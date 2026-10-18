# Word 文档格式的静态检查

只读、不改文件、可进 CI。实现：[docxaudit](https://github.com/Asgoresmilelong/docxaudit)（Python，MIT）。

输入一个 `.docx`，输出 `(规则, 级别, 段落号, 说明)`。约束四条：不修改原文件、结果确定、定位到段、有退出码。

## 一、属性要按级联求值

Word 的格式不是读出来的一个值，而是四级叠加：**段落直接格式 → 段落样式链（`basedOn`）→ 表格样式 → `docDefaults`**。`python-docx` 只回答"挂了哪个样式"，回答不了"实际是什么"。

```python
# 伪代码
def resolve(node, prop):
    if node.direct[prop] is not None:
        return node.direct[prop]
    for style in style_chain(node.style):      # 沿 basedOn 回溯，需防环
        if style[prop] is not None:
            return style[prop]
    return doc_defaults[prop]
```

三个容易写错的细节：`w:sz` 是半磅，`w:line` 在 `exact` 下是 1/20 磅、在 `auto` 下是 1/240 行；`w:rFonts` 可能是 `theme:minorEastAsia` 这类主题引用，静态解析不出真字体——本实现直接返回 `None` 回退，宁可缺失，不要错值。

## 二、编号是域，不是文本

题注和交叉引用在 Word 里是 `SEQ` / `REF` 域。复杂域的结构是 `begin → instrText → separate → 结果 → end`，`separate` 之后只是**上次求值的缓存**。文档流转中被改过又没按 F9，缓存就是过期的。

所以本工具不判断"编号对不对"，只判断"序列自不自洽"：

```python
for kind, num, p in captions:                  # captions = [(图|表, "3-2", 段落)]
    parts = num.split("-")                     # 支持「图 3-2」章节式编号
    groups.setdefault((kind, "-".join(parts[:-1])), []).append((parts[-1], p))

expected = None
for raw, p in groups[key]:
    if expected is not None and int(raw) != expected:
        report(f"{kind} 编号不连续：出现 {raw}，应为 {expected}")
    expected = int(raw) + 1
```

`w:instrText` 已解析进 `Paragraph.field_codes`，后续可重算域再比对，判断"编号是否过期"——v0.1 未实现。

## 三、顺序不等于 XML 顺序

段落会藏在 `w:tbl`、`w:sdt`、`w:customXml` 里。漏掉任一层，不是少算几段，而是所有"第 N 段"整体错位——而这正是诊断唯一可操作的部分。另外 `w:del` 里被删的文字不能计入正文。

```python
if tag == qn("p"):
    yield el, in_table
elif tag == qn("tbl"):
    for row in el.findall(qn("tr")):
        for cell in row.findall(qn("tc")):
            yield from _iter_paragraphs(cell, True)
elif tag == qn("sdt"):
    yield from _iter_paragraphs(el.find(qn("sdtContent")), in_table)
```

复杂度 O(N)，N 为元素总数。

## 四、统计规则先定义母体

字体/行距一致性是统计判断，不是逻辑判断。第一版把参考文献条目也算进正文：

| 字体 | 票数 | 来源 |
|---|---|---|
| 宋体 | 5 | 正文 |
| 未设置 | 3 | 文献条目 |
| 黑体 | 1 | 真正的异常 |

5/9 = 0.56 < 0.6，规则直接退出——**真问题被合法的少数派稀释**。修法是把参考文献区整段排除。阈值定为：样本 n ≥ 3 且多数派占比 ≥ 0.6，否则静默（证据不足不报）。

投票以**段落**为单位而非 run：run 数与文本长度强相关，按 run 投票会让长段落主导结果。

## 五、一个跨平台坑

26 个测试全过、本地正常，CI 上 Windows 三个任务全红：

```
UnicodeEncodeError: 'charmap' codec can't encode characters in position 36-37
  File "...\encodings\cp1252.py", line 19, in encode
```

Windows runner 控制台是 cp1252，报告里的中文直接 `print` 就炸；本地是 UTF-8 控制台，所以测试没测到。修法在入口统一改写输出流：

```python
def ensure_utf8_console() -> None:
    for stream in (sys.stdout, sys.stderr):
        if stream is None:
            continue
        enc = (getattr(stream, "encoding", None) or "").lower().replace("-", "").replace("_", "")
        if enc == "utf8":
            continue
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:          # 被重定向到 StringIO 时没有该方法
            continue
        try:
            reconfigure(encoding="utf-8", errors="replace")
        except (ValueError, OSError):
            pass
```

回归测试用 `TextIOWrapper(encoding="cp1252", errors="strict")` 接管 stdout，不依赖宿主编码。

> 凡是会输出非 ASCII 的 CLI，CI 里必须有 Windows。

## 实测

| 指标 | 值 |
|---|---|
| 真实文档 | 947 段 / 1629 run / 22 表；`document.xml` 306 KiB，**`styles.xml` 342 KiB**（比正文还大） |
| 耗时 | 解析 86.0 ms，规则 6.6 ms（7 次中位数），解析占 93% |
| 真实语料 | 报 1 条提示（21 条文献未被引用），人工核对属实 |
| 自检样本 | 埋 6 处错误，六条规则全部命中，无额外误报 |
| 测试 | 27 个用例，夹具现场生成，无外部样本 |

**局限**：主题字体解析不出、域未重算、只支持 `.docx`、只认 `[n]` 顺序编码制。

```bash
pip install git+https://github.com/Asgoresmilelong/docxaudit
docxaudit check report.docx --json
```
