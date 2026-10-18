# Word 格式检查为什么必须直接读 OOXML：一次踩坑记录

> 起因很朴素：一百多页的课程报告，导师批注只有四个字——「格式再改改」。
> 我盯着导航窗格翻了半小时，找出三处图号跳号、两处「如图 5 所示」其实没有图 5、一段莫名变成黑体的正文。
> 于是有了 [docxaudit](https://github.com/Asgoresmilelong/docxaudit)：一个只读、不改文件、能进 CI 的 Word 格式体检工具。

## 0. 先说结论

- `python-docx` 很适合**写**文档，不适合**判**文档：它给你的是「XML 里有什么」，而格式检查要的是「Word 最终认什么」。
- 有三件事必须下钻到 OOXML：**样式继承链**、**域代码（SEQ / REF）**、**文档顺序（表格 / 内容控件 / 修订）**。
- 还有一个和 Word 无关、但一定会踩的坑：**你的 CLI 在 Ubuntu 上绿了，不代表在 Windows 上活着**（后面有真实事故）。

下面每条都附真实代码，来自仓库 `docxaudit/loader.py` 与 `rules.py`。

---

## 1. 坑一：样式不是属性，是一条继承链

新手最容易写的代码是这样的：

```python
from docx import Document

doc = Document("报告.docx")
for p in doc.paragraphs:
    print(p.style.name, p.text[:20])
```

问题在于，`p.style.name` 只回答「这个段落挂了哪个样式」，回答不了「这个段落**实际**是几级标题、**实际**是什么字体」。

真实的 OOXML 里，同一个属性有三个来源，优先级从高到低：

| 层级 | 位置 | 例子 |
|---|---|---|
| 段落直接格式 | `w:p/w:pPr` | 你手动把某段字号改成 14 磅 |
| 样式 | `w:styles/w:style` | 「标题 3」定义了大纲级别 |
| 样式基类 | `w:style/w:basedOn` 指向的祖先 | 「标题 3」继承自「标题」 |
| 文档默认 | `w:docDefaults/w:rPrDefault` | 全局宋体五号 |

而 Word 判断「这是几级标题」，看的是 `w:outlineLvl`——它可能写在段落上，也可能写在样式上，还可能只写在**基类样式**上。

所以 `docxaudit` 自己写了一层解析，先把样式表读成 `styleId → Style`，再沿着 `basedOn` 往回走：

```python
def style_chain(self, style_id: str | None, limit: int = 12) -> list[Style]:
    """按 basedOn 向上回溯的样式链（不含 None）。"""
    out: list[Style] = []
    seen: set[str] = set()
    cur = style_id
    while cur and cur in self.styles and cur not in seen and len(out) < limit:
        seen.add(cur)
        st = self.styles[cur]
        out.append(st)
        cur = st.based_on
    return out
```

字体同理，取「第一个定义了的值」：

```python
for st in style_chain:
    if east is None and st.east_asia:
        east = st.east_asia
    if latin is None and st.latin:
        latin = st.latin
    if size is None and st.size_pt is not None:
        size = st.size_pt

east = east or doc.default_east_asia
latin = latin or doc.default_latin
```

**这里还有个坑中坑**：`w:rFonts` 里存的可能是 `w:eastAsiaTheme="minorEastAsia"` 这种主题字体引用，而不是字体名。主题定义在 `theme1.xml` 里，要把 `minorEastAsia` 映射成「宋体」得再解析一层主题。当前版本选择退回样式或 `docDefaults` 的值——宁可不知道，也不要猜错。

---

## 2. 坑二：题注编号是「域」，不是文本

这是最容易被忽略的一层。

你在 Word 里看到的「图 3」，在 XML 里可能长这样：

```xml
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> SEQ 图 \* ARABIC </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>3</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

也就是说，`3` 只是**上次更新域时的缓存结果**。文档传给别人、对方没按 F9 更新，显示的就是过期编号——你看到的「图 3」可能早就该是「图 4」。

这带来两个现实后果：

1. **只读文本，永远不知道编号是不是自动生成的**。一个被手动改成「图 3」的题注，和自动编号的「图 3」，在纯文本视角下完全一样。
2. **做连续性校验时，不能相信单个编号**，只能相信「顺序」。所以 `docxaudit` 的 R2 规则是按出现顺序检查递增：

```python
groups: dict[tuple[str, str], list[tuple[str, Paragraph]]] = {}
for kind, num, p in caps:          # caps: [(图/表, "3-2", 段落), ...]
    parts = num.split("-")
    prefix = "-".join(parts[:-1])  # 支持「图 3-2」这种章节式编号
    groups.setdefault((kind, prefix), []).append((parts[-1], p))
```

同一章内，期望值 = 上一个 + 1，对不上就报「按顺序应为 图 2」。重复编号则直接报错。

域指令本身没有浪费——`loader.py` 把它们也读了出来：

```python
def _paragraph_fields(p) -> list[str]:
    return [el.text.strip() for el in p.iter(qn("instrText")) if el.text and el.text.strip()]
```

现在还没有规则消费它，但这是个明确的扩展点：将来可以比对「域算出来的编号」和「缓存文本里的编号」，不一致就提示「编号已过期，请按 Ctrl+A、F9 更新域」。这比现在只查连续性更准。

---

## 3. 坑三：文档的「顺序」不等于 XML 的「顺序」

如果你的报告只有正文段落，`for p in doc.paragraphs` 够用。但真实的 Word 文档里，正文段落会藏在三个地方：

1. **表格**：`w:tbl → w:tr → w:tc → w:p`，单元格里的段落也是段落；
2. **内容控件**：`w:sdt → w:sdtContent → w:p`，从网页粘贴、插入目录、模板填空都会产生；
3. **修订记录**：`w:del` 里是被删掉的文字，`w:ins` 里是新增的。

漏掉任何一层，「第 N 段」的定位就全错——而「第 N 段」是这个工具唯一有用的东西。

所以遍历是手写的，不是 `doc.paragraphs`：

```python
def _iter_paragraphs(container, in_table: bool = False):
    """按文档顺序产出 (w:p, in_table)，会下钻表格与内容控件（sdt）。"""
    for el in container:
        tag = el.tag
        if tag == qn("p"):
            yield el, in_table
        elif tag == qn("tbl"):
            for row in el.findall(qn("tr")):
                for cell in row.findall(qn("tc")):
                    yield from _iter_paragraphs(cell, True)
        elif tag == qn("sdt"):
            content = el.find(qn("sdtContent"))
            if content is not None:
                yield from _iter_paragraphs(content, in_table)
```

提取文字时也要绕开被删除的内容：

```python
if node.tag == qn("t"):
    anc = node.getparent()
    while anc is not None:
        if anc.tag == qn("del"):
            deleted = True
            break
        anc = anc.getparent()
```

不这么做，一段改过三次的正文，字数统计会多出两遍。

---

## 4. 彩蛋坑：Ubuntu 绿了，不代表 Windows 活着

代码写完，26 个单元测试全过，本地 PowerShell 跑得好好的。推到 GitHub，配了个 2×3 的 CI 矩阵（Ubuntu / Windows × Python 3.9 / 3.11 / 3.12）。

第一次结果：

```
ubuntu-latest, 3.9    success
ubuntu-latest, 3.11   success
ubuntu-latest, 3.12   success
windows-latest, 3.9   failure
windows-latest, 3.11  failure
windows-latest, 3.12  failure
```

日志里这一行是全部原因：

```
UnicodeEncodeError: 'charmap' codec can't encode characters in position 36-37:
character maps to <undefined>
  File "C:\hostedtoolcache\...\lib\encodings\cp1252.py", line 19, in encode
```

Windows runner 的控制台编码是 **cp1252**，而我的报告输出里全是中文和 `·` 这种符号，`print` 直接炸。本地之所以没事，是因为我的 PowerShell 是 UTF-8 控制台——**测试环境比生产环境"更友好"，于是测试没测到真问题**。

修法不是在 workflow 里加一行 `PYTHONIOENCODING` 就完事（虽然也加了，作为双保险），而是让库自己扛住：

```python
def ensure_utf8_console() -> None:
    """让 stdout/stderr 能输出中文。

    Windows 控制台和 CI 管道默认是 cp1252 / GBK，直接 print 中文会抛
    UnicodeEncodeError；这里能改就改成 UTF-8，改不了（比如被重定向到
    StringIO）就原样保留。
    """
    for stream in (sys.stdout, sys.stderr):
        if stream is None:
            continue
        encoding = (getattr(stream, "encoding", None) or "").lower().replace("-", "").replace("_", "")
        if encoding == "utf8":
            continue
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:
            continue
        try:
            reconfigure(encoding="utf-8", errors="replace")
        except (ValueError, OSError):
            pass
```

然后补了一条**模拟 cp1252 控制台**的回归测试，让这个问题永远回不来：

```python
def test_non_utf8_console_does_not_crash(self):
    """Windows / CI 管道默认 cp1252，之前这里会抛 UnicodeEncodeError。"""
    raw = io.BytesIO()
    stream = io.TextIOWrapper(raw, encoding="cp1252", errors="strict")
    with contextlib.redirect_stdout(stream):
        code = cli_main(["check", str(path), "--no-fail"])
    stream.flush()

    self.assertEqual(code, 0)
    self.assertIn("docxaudit", raw.getvalue().decode("utf-8"))
```

经验只有一句：**凡是会 `print` 中文的 CLI，CI 里必须有 Windows。** Ubuntu 的 UTF-8 环境会替你掩盖掉一整类问题。

---

## 5. 判据设计：先定义「母体」，再谈「多数派」

「正文行距应该统一」这句话，机器执行起来有个隐藏前提：**哪些段落算正文？**

第一版我把「非标题、非题注、长度够」的段落都算正文，然后在示例文档上跑，R5（字体一致性）和 R6（行距一致性）一条都没报。

原因很蠢：参考文献那 21 条，条目格式本来就特殊（字号更小、无首行缩进），它们把「未设置」的票数拉到了 3 票。最终投票结果：

```
宋体   5 票   ← 真正的正文
未设置 3 票   ← 参考文献条目
黑体   1 票   ← 那个真问题
```

5 / 9 = 0.55，没过 0.6 的阈值，规则直接静默退出——**真问题被"合法的少数派"稀释掉了**。

修法是给规则划定母体：参考文献区（从「参考文献」标题到下一个同级标题）整段排除。

```python
def body_paragraphs(doc: Document, min_chars: int = 20) -> Iterable[Paragraph]:
    """正文段落：排除表格、标题、题注、目录，以及参考文献区。

    参考文献区有自己的排版规范（悬挂缩进、字号更小），把它算进“正文多数派”
    会把真正的正文偏差稀释掉。
    """
    excluded = _reference_para_nos(doc)
    for p in doc.paragraphs:
        if p.no in excluded:
            continue
        if p.in_table or p.heading_level is not None or is_caption(p) or is_toc_like(p):
            continue
        if len(p.stripped) < min_chars:
            continue
        yield p
```

另外两个保守设计，都是为了压误报：

- 票数少于 3 票时不判（文档太短，多数派没有意义）；
- 最多数占比低于 60% 时不判（说明这份文档本来就有多种合法排版）。

统计类规则的通病是「看起来很聪明，实际全是误报」。宁可漏报，不要乱报——用户被误报骗两次，就再也不会打开了。

---

## 6. 真实效果

用 `examples/make_sample.py` 生成一份埋了 6 处典型错误的文档，跑出来是这样：

```
docxaudit · sample-broken.docx
[R1 错误] 第 4 段：标题层级跳跃：从 1 级直接跳到 3 级
[R3 错误] 第 5 段：正文引用了「图 6」，但全文没有这个编号的题注
[R2 警告] 第 7 段：图 编号不连续：出现 图 3，按顺序应为 图 2
[R5 警告] 第 8 段：正文中文字体应为「宋体」，此段为「黑体」
[R6 警告] 第 9 段：段落间距为「固定行距 30.0 磅，无首行缩进」，
          多数正文为「固定行距 22.0 磅，首行缩进 24.0 磅」
[R4 提示] 第 13 段：有 2 条参考文献在正文中从未被引用：[2]、[3]

共 6 条：错误 2 · 警告 3 · 提示 1
```

六条规则，六个问题，没有一条是我手动指定的。

换到一份真实的 947 段报告上（有 50 个标题、6 个题注、21 条参考文献），只报了 1 条提示：

```
[R4 提示] 第 927 段：有 21 条参考文献在正文中从未被引用：[1]、[2]、…
共 1 条：错误 0 · 警告 0 · 提示 1
```

这条是**真的**——那份报告的文献表确实没在正文里引用过。

---

## 7. 局限（写在前面比藏起来好）

- **只支持 `.docx`**。`.doc` 是二进制格式，请先另存。
- **主题字体无法解析**，遇到 `theme:minorEastAsia` 会退回样式值，可能判错。
- **只认顺序编码制**（`[1]`、`[2]`），著者-出版年制（APA）暂不支持。
- **域编号未参与校验**，目前只查连续性，不判断「编号是否过期」。
- 统计类规则本质是启发式，不保证零误报。

---

## 8. 最后

工具本身不复杂：六个规则、一千多行代码、27 个测试。真正花时间的是**把 Word 的隐式语义还原成显式结构**——继承链、域、文档顺序，这三件事都不在高层 API 的视野里。

```bash
pip install git+https://github.com/Asgoresmilelong/docxaudit
docxaudit check 你的报告.docx
```

仓库：https://github.com/Asgoresmilelong/docxaudit

如果它帮你省下了一次「格式再改改」，那就值了。也欢迎直接提 Issue 骂误报——误报率是这类工具唯一的护城河。
