# Word 文档格式的静态检查：OOXML 结构解析、规则设计与工程约束

> **摘要**：本文讨论对 `.docx` 做**只读格式检查**的可行实现路径。核心论点是：Word 的"格式"不是一组可以直接读取的属性，而是一个需要按继承规则求值的**级联系统**，其中还混入了域（field）这种**派生值**。忽略任一层次，静态检查的结论都会失真。文中给出解析模型、六条规则的判定算法与复杂度、误报控制策略，以及一次 Windows CI 上的编码事故复盘。实现见 [docxaudit](https://github.com/Asgoresmilelong/docxaudit)（Python，MIT）。

## 1. 问题定义

给定文档 D（一个 `.docx` 包），输出有限的诊断集合

F = {(rule, severity, para_no, message)}

其中 rule 为规则编号，severity ∈ {error, warn, info}，para_no 为段落定位，message 为自然语言说明。

约束条件：

1. **只读**：不得修改 D 的任何字节，检查过程不产生副作用；
2. **确定性**：同一输入必得同一输出（无网络、无时间戳、结果按 `(severity, para_no, rule)` 全序排序）；
3. **可解释**：每条诊断必须能定位到具体段落并说明成因，不接受"格式不规范"这类不可操作结论；
4. **可集成**：`--json` 提供结构化输出，存在 error 时返回非零退出码。

**非目标**：不做视觉渲染比对（那是另一条技术路线，见 §9），不做自动修复，不做内容质量评价。

## 2. 包结构与解析入口

`.docx` 是 ZIP 容器，格式检查真正需要的 part 只有三个：

| Part | 作用 | 规模（实测，一份 947 段报告） |
|---|---|---|
| `word/document.xml` | 正文、表格、域、修订 | 306.2 KiB |
| `word/styles.xml` | 样式定义与文档默认值 | **341.6 KiB** |
| `word/numbering.xml` | 列表/多级编号定义 | 可选 |

有一个反直觉的数据点：**样式表比正文还大**。原因是模板沿用了 `python-docx` 默认模板，内含近百个未使用的样式定义。这意味着两件事：全量解析样式表的成本不可忽略；但样式表可整体缓存，与段落数无关。

`python-docx` 的高层 API 在此不合适——它暴露的是"XML 里有什么"，而检查需要"Word 最终认什么"。因此解析层直接用 `lxml`：

```python
with zipfile.ZipFile(path) as zf:
    document_xml = zf.read("word/document.xml")
    styles_xml = zf.read("word/styles.xml") if "word/styles.xml" in zf.namelist() else None
```

解析结果收敛为不可变的 `Document / Paragraph / Run` 结构，规则层只依赖该结构，不接触 XML。这一分层使得规则可以被独立测试（全部夹具用 `python-docx` 现场生成，无需外部样本）。

实测开销（Windows，Python 3.12，取 7 次中位数）：

| 阶段 | 耗时 |
|---|---|
| `load_docx`（解压 + 两次 XML 解析 + 947 段属性求值） | 86.0 ms |
| `audit`（6 条规则全跑） | 6.6 ms |

解析占 93%，规则占 7%。因此优化方向明确在解析：若需支持大文档，可对 `styles.xml` 做惰性解析，只解析被引用的样式。

## 3. 有效属性求值：级联模型

### 3.1 级联顺序

WordprocessingML 中，一个段落属性的值可能来自四个层级，优先级由高到低：

| 优先级 | 来源 | XML 位置 |
|---|---|---|
| 1 | 段落直接格式 | `w:p/w:pPr` |
| 2 | 段落样式链（含 `basedOn` 祖先） | `w:styles/w:style/w:pPr` |
| 3 | 表格样式（段落位于表格内时） | `w:tblPr/w:tblStyle` 指向的样式 |
| 4 | 文档默认 | `w:docDefaults/w:pPrDefault` |

字符属性（字体、字号、加粗）同理，但多一层字符样式 `w:rStyle`，位于 run 直接格式与段落样式之间。求值伪代码：

```
resolve(node, prop):
    if node.direct[prop] is not None:      return node.direct[prop]
    for style in style_chain(node.style):  # 沿 basedOn 回溯
        if style[prop] is not None:        return style[prop]
    if node.in_table and table_style[prop] is not None:
        return table_style[prop]
    return doc_defaults[prop]
```

### 3.2 `basedOn` 链

样式链是求值的基础，必须处理**环形引用**与**超长链**：

```python
def style_chain(self, style_id: str | None, limit: int = 12) -> list[Style]:
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

`seen` 集合保证在畸形文档上终止；`limit` 是防御性上限（实测正常文档链深 ≤ 4，Word 自身限制为 11 层）。

### 3.3 开关属性的语义陷阱

`w:b`（加粗）、`w:i`、`w:caps` 这类属性是**开关（toggle）类型**：元素出现即表示真，`w:val="0"` 表示假。但它们在多层定义时的组合语义并不等价于"取第一个非空值"——`ECMA-376` 对 toggle 属性规定了切换语义，且不同实现之间存在差异。

本工具采用**首个显式定义者获胜**的近似：

```python
def _toggle(el) -> bool | None:
    if el is None:
        return None
    val = _attr(el, "val")
    if val is None:
        return True
    return val.lower() not in {"0", "false", "off"}
```

对"样式里加粗、run 上再加粗一次"这类情形，近似值与真实渲染可能相反。这是有意的取舍：精确实现切换语义需要完整复刻 Word 的格式合并算法，收益远小于成本；该偏差已写入局限说明。

### 3.4 单位换算

三个单位最容易写错：

| 属性 | 单位 | 换算 |
|---|---|---|
| `w:sz` | 半磅 | `pt = val / 2` |
| `w:spacing/@w:line`（`lineRule="exact"`） | 二十分之一磅 | `pt = val / 20` |
| `w:spacing/@w:line`（`lineRule="auto"`） | 240 分之一行 | `倍数 = val / 240` |
| `w:ind/@w:firstLineChars` | 1/100 字符 | `字符 = val / 100` |

报告里的「固定行距 22.0 磅」「首行缩进 24.0 磅」就是由这些换算得到的。

### 3.5 主题字体：静态解析的边界

`w:rFonts` 中可能不是字体名，而是主题引用：

```xml
<w:rFonts w:eastAsiaTheme="minorEastAsia" w:asciiTheme="minorHAnsi"/>
```

真实字体定义在 `word/theme/theme1.xml` 的 `a:fontScheme` 中，且区分 `majorFont`（标题）与 `minorFont`（正文），按语言脚本（`<a:font script="Hans" typeface="宋体"/>`）进一步细化。要正确解析，需要读第四个 part 并实现脚本匹配。

当前实现的策略是**显式放弃**：遇到 `theme:` 前缀返回 `None`，回退到样式或 `docDefaults` 的值。

```python
def _theme_font(raw: str | None) -> str | None:
    return None if raw and raw.startswith("theme:") else raw
```

宁可信息缺失，也不要给出一个看似确定的错误结论——这与 §6.7 的误报控制是同一原则。

### 3.6 标题级别的识别

Word 判断"这是几级标题"的依据是 `w:outlineLvl`（0 基，0 即一级），它同样遵循 §3.1 的级联。但现实文档里存在两种退化情形：

1. 样式名是 `Heading 3` 或本地化的 `标题 3`，但 `outlineLvl` 未显式声明（继承自基类）；
2. 段落直接套用了 `outlineLvl`，样式名却是 `Normal`。

因此识别逻辑是"`outlineLvl` 优先，样式名兜底"：

```python
@property
def heading_level(self) -> int | None:
    if self.outline_level is not None and 0 <= self.outline_level <= 8:
        return self.outline_level
    name = (self.style_name or "").strip()
    m = re.match(r"^(?:heading|标题)\s*([1-9])", name, re.I)
    return int(m.group(1)) - 1 if m else None
```

## 4. 域：派生值与观测值

### 4.1 域的两种编码

题注编号（`SEQ`）与交叉引用（`REF`、`PAGEREF`）在 Word 中是**域**，有两种编码：

**简单域**：

```xml
<w:fldSimple w:instr=" SEQ 图 \* ARABIC ">
  <w:r><w:t>3</w:t></w:r>
</w:fldSimple>
```

**复杂域**（由三个 `w:fldChar` 分隔）：

```xml
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> SEQ 图 \* ARABIC </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>3</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

`separate` 与 `end` 之间的内容是**上次求值时的缓存结果**。这带来一个重要的认识论问题：

> 文档里显示的编号是**派生值的历史快照**，不是当前应有的值。

若文档在流转中被编辑而未更新域，缓存结果可能与正确编号不一致。`w:fldSimple/@w:dirty="true"` 与 `settings.xml` 中的 `w:updateFields` 是仅有的两个"需要更新"提示，且都不是强制。

### 4.2 对本工具的影响

由于要重新计算域需要实现 `SEQ`/`REF` 的求值语义（含书签表、章节编号、编号格式），本工具采用**观测值一致性**策略：不判断编号是否正确，只判断**观测序列是否自洽**。

R2 的判定：按文档顺序分组统计，同组内期望值严格 +1。

```python
for kind, num, p in caps:                 # caps = [(图|表, "3-2", 段落)]
    parts = num.split("-")
    prefix = "-".join(parts[:-1])         # 章节式编号的组键
    groups.setdefault((kind, prefix), []).append((parts[-1], p))

expected = None
for raw, p in groups[(kind, prefix)]:
    if expected is not None and int(raw) != expected:
        report(f"{kind} 编号不连续：出现 {kind} {raw}，按顺序应为 {kind} {expected}")
    expected = int(raw) + 1
```

复杂度 O(n)，且对缓存是否过期免疫——只要编号是递增的，就认为自洽。

域指令本身被解析并保存在 `Paragraph.field_codes` 中：

```python
def _paragraph_fields(p) -> list[str]:
    return [el.text.strip() for el in p.iter(qn("instrText")) if el.text and el.text.strip()]
```

这为后续能力留了接口：**重算域并与缓存比对**，即可判断"编号已过期"。实现路径有两条：实现 `SEQ` 的子集语义，或调用 LibreOffice headless 更新域后做 diff。前者精确但工作量大，后者需要外部依赖，均未在 v0.1 实现。

### 4.3 交叉引用

`REF` 域引用的是**书签**（`w:bookmarkStart/@w:name`），文章里"如图 3 所示"在专业排版中应是 `REF _Ref12345 \h`。但大量实际文档是**手打的静态文本**——这正是 R3 存在的原因：手打文本没有域来保证一致性。

## 5. 文档顺序与内容提取

### 5.1 三类容器

`document.xml` 的 `w:body` 并不是段落数组，而是一个混合序列，段落可能出现在三个容器中：

| 容器 | 路径 | 说明 |
|---|---|---|
| 表格 | `w:tbl → w:tr → w:tc → w:p` | 单元格内段落必须计入序号 |
| 内容控件 | `w:sdt → w:sdtContent → ...` | 目录、模板占位、网页粘贴产物 |
| 自定义 XML | `w:customXml → ...` | 少见，但存在于旧版本文档 |

漏掉任一层的后果不是"少统计几个段落"，而是**所有"第 N 段"的定位整体错位**——而这正是诊断唯一可操作的部分。

```python
def _iter_paragraphs(container, in_table: bool = False):
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

递归深度受嵌套表格层数约束（实际 ≤ 3），复杂度 O(N)，N 为元素总数。

### 5.2 修订标记

`w:ins` 是新增内容，`w:del` 是删除内容；区别在于被删文字位于 `w:delText` 而非 `w:t`，但**并非所有生成器都严格遵循**（LibreOffice 与部分在线编辑器存在把 `w:t` 置于 `w:del` 内的情况）。因此提取时不能只按标签名过滤，还要检查祖先：

```python
if node.tag == qn("t"):
    anc = node.getparent()
    while anc is not None:
        if anc.tag == qn("del"):
            deleted = True
            break
        anc = anc.getparent()
```

同类标签还有 `w:moveFrom`（移动的源）与 `w:moveTo`（移动的目标），在"比较文档并接受修订"的产物中常见。

### 5.3 文本提取规则

| XML 元素 | 提取为 |
|---|---|
| `w:t`（且不在 `w:del` / `w:moveFrom` 内） | 原文 |
| `w:tab` | `\t` |
| `w:br`（含 `w:type="page"`）/ `w:cr` | `\n` |
| 其他 | 忽略 |

段落序号定义为**文档序上的 1 基整数**，包含表格内段落。该定位在同一次解析中稳定、可复现；但它随内容编辑而漂移，因此不适合作为跨版本的持久标识（持久标识应使用 `w:paraId`，Word 2013+ 生成，但并非所有生成器都写）。

## 6. 规则设计

### 6.1 接口

每条规则是纯函数 `Document → list[Finding]`，不读文件、不打日志、不共享可变状态：

```python
@dataclass
class Rule:
    id: str
    name: str
    description: str
    run: Callable[[Document], list[Finding]]
```

这让每条规则都能用最小夹具独立验证，也让 `--only/--ignore` 的过滤成为简单集合运算。

### 6.2 R1 标题层级

对标题序列 h₁, h₂, …, h_k（h_i ∈ {0,…,8}）：

- 若 h₁ ≠ 0，报 warn（文档通常以一级标题开头）；
- 若 h_i > h_{i−1} + 1，报 error（跳级）。

复杂度 O(k)。跳级之所以定为 error，是因为它会破坏导航窗格与自动目录的结构，是**结构性**缺陷而非风格偏好。

### 6.3 R2 题注编号

见 §4.2。补充两点：

- **重复检测**：同一 `(kind, normalized_number)` 出现两次即 error，用于捕捉"复制粘贴后未更新域"；
- **归一化**：`3-1`、`3.1`、`3–1`（en dash）统一为 `3-1`，避免因连字符字符集不同而误判。

### 6.4 R3 交叉引用

集合运算：令 C 为题注编号集合，M 为正文中的提及集合（排除题注段落与目录段落）。对每个 m ∈ M：

1. 若 m ∈ C，通过；
2. 否则若 m 的**尾段**（最后一个 `-` 之后的部分）出现在 C 的尾段集合中，通过；
3. 否则报 error。

规则 2 用于容忍"图 3-1"被正文简称为"图 1"的写法，代价是可能放过"图 1"实为"图 4-1"的笔误。这是**精确率与召回率的显式权衡**，选择偏向减少误报。

### 6.5 R4 参考文献

两条独立判定：

1. **悬空引用**：[n] 出现在正文但不在文献表 → error；
2. **未被引用**：文献表条目从未在正文出现 → info（聚合为一条，避免刷屏）。

引用区间需要在解析阶段展开，且必须防御恶意/畸形输入：

```python
m = re.match(r"^(\d+)\s*[-–—]\s*(\d+)$", part)
if m:
    a, b = int(m.group(1)), int(m.group(2))
    if a <= b and b - a <= 200:          # 上限防御：避免 [1-999999] 造成内存放大
        nums.extend(range(a, b + 1))
```

上限 200 是工程折衷：正常文献表的引用区间远小于此值，而恶意输入可构造 10⁶ 量级的展开。

### 6.6 R5 / R6 一致性判定

这两条不是逻辑判断，而是**统计判断**，因此必须明确定义三件事：**母体、投票粒度、阈值**。

**母体（population）**。第一版定义为"非标题、非题注、长度 ≥ 20 字符的段落"，在示例文档上 R5/R6 全部静默。原因是文献表的 21 条条目被计入：

| 字体 | 票数 | 来源 |
|---|---|---|
| 宋体 | 5 | 正文 |
| 未设置 | 3 | 参考文献条目 |
| 黑体 | 1 | 真正的异常段落 |

5/9 = 0.56 < 0.6，规则直接退出。**真实的少数派被合法的多数派稀释**——这是统计类规则的典型失效模式。

修正后的母体显式排除参考文献区（从「参考文献」标题到下一个同级标题）：

```python
excluded = _reference_para_nos(doc)
for p in doc.paragraphs:
    if p.no in excluded or p.in_table or p.heading_level is not None:
        continue
    if is_caption(p) or is_toc_like(p) or len(p.stripped) < 20:
        continue
    yield p
```

**投票粒度**。以**段落**为单位投票，而非以 run 为单位。原因是 run 的数量与文本长度强相关：一段 800 字的正文可能被拆成 40 个 run，而一段 30 字的异常段落只有 2 个 run，按 run 投票会让长段落主导结果。

**阈值**。样本数 n ≥ 3 且多数派占比 ≥ 0.6 才判定，否则静默。两个阈值的含义不同：前者排除"文档太短，多数派无意义"，后者排除"文档本身存在多种合法排版"。

### 6.7 误报的代价

这类工具的唯一护城河是**误报率**。用户被误报欺骗两次，就会永久卸载。因此所有规则遵循"证据不足即不报"：主题字体解析不出→跳过；样本不足→跳过；多数派不显著→跳过。

代价是召回率下降，这是显式设计选择，而非缺陷。

## 7. 工程问题：编码与跨平台

### 7.1 事故

26 个单元测试通过、本地 PowerShell 输出正常，推送到 CI 后 2×3 矩阵中**Windows 三个任务全部失败**：

```
UnicodeEncodeError: 'charmap' codec can't encode characters in position 36-37:
character maps to <undefined>
  File "C:\hostedtoolcache\...\lib\encodings\cp1252.py", line 19, in encode
```

根因不在 Word，也不在 `lxml`，而在输出层：

- Windows runner 的控制台/管道编码为 **cp1252**；
- 报告文本包含中文与 `·` 等非 ASCII 字符；
- `print()` 在写入时按 `sys.stdout.encoding` 编码，遇到不可映射字符即抛异常。

本地之所以未暴露，是因为开发机控制台为 UTF-8——**测试环境比生产环境更宽容，测试因此失效**。这是一个可推广的结论：任何非 ASCII 输出的 CLI，其测试矩阵必须包含默认非 UTF-8 的宿主。

### 7.2 处置

在 CLI 入口统一改写输出流的编码，而非在调用点零散处理：

```python
def ensure_utf8_console() -> None:
    for stream in (sys.stdout, sys.stderr):
        if stream is None:
            continue
        enc = (getattr(stream, "encoding", None) or "").lower()
        if enc.replace("-", "").replace("_", "") == "utf8":
            continue
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:
            continue
        try:
            reconfigure(encoding="utf-8", errors="replace")
        except (ValueError, OSError):
            pass
```

三个细节：`reconfigure` 仅存在于 `TextIOWrapper`（被重定向为 `StringIO` 时不存在，静默跳过）；`errors="replace"` 保证即使 `reconfigure` 失败也不会崩溃；判断 UTF-8 时同时兼容 `utf-8` / `utf_8` 两种拼写。

### 7.3 回归测试

修复必须在测试层固化，否则同类问题会以另一种形式回归。做法是构造一个 **cp1252 的 `TextIOWrapper`** 接管 stdout：

```python
raw = io.BytesIO()
stream = io.TextIOWrapper(raw, encoding="cp1252", errors="strict")
with contextlib.redirect_stdout(stream):
    code = cli_main(["check", str(path), "--no-fail"])
stream.flush()

assert code == 0
assert "docxaudit" in raw.getvalue().decode("utf-8")
```

该测试在 Linux/macOS 上同样有效——它不依赖宿主编码，而是显式构造了故障条件。

## 8. 验证

### 8.1 单元测试

27 个用例，覆盖每条规则的命中与不命中、过滤器、CLI 退出码、JSON 输出、非 UTF-8 控制台。夹具由 `python-docx` 现场生成，因此测试不依赖仓库内的二进制文件，也不存在夹具与实现同时腐化的风险。

### 8.2 真实语料

一份 947 段、50 个标题、6 个题注、21 条参考文献的真实报告：

- 解析：947 段、1629 个 run、22 个表格；
- 输出：1 条 info（21 条文献从未被引用），0 error、0 warn；
- 人工核对：该结论属实。

对照组：`examples/make_sample.py` 生成的 14 段文档，埋入 6 处典型错误，六条规则**全部命中且无额外误报**。

### 8.3 未覆盖

| 项 | 原因 | 可能路径 |
|---|---|---|
| 主题字体解析 | 需读 `theme1.xml` 并实现脚本匹配 | 增加 theme 解析层 |
| 域重算 | 需实现 `SEQ`/`REF` 子集语义 | 或调用 LibreOffice headless |
| `.doc` | 二进制格式，需 CFB 解析 | 明确不支持 |
| 著者-出版年制引用 | 规则基于 `[n]` 正则 | 增加引用风格插件 |
| 渲染差异 | 静态检查不涉及布局引擎 | 属另一技术路线 |

## 9. 与相关工作对比

| 方案 | 范式 | 局限 |
|---|---|---|
| `python-docx` | 对象模型 API | 不提供级联求值，缺域支持 |
| WordFormat 类工具 | **修改**文档 | 不产生诊断，无法进 CI |
| Word「比较文档」 | 渲染/结构差异 | 需两份文档，非规则检查 |
| 本文方案 | 静态规则检查 | 无自动修复，依赖启发式 |

差异在于目标函数：前几者优化"改得对"，本方案优化"**说得清**"。

## 10. 结论与后续

对 `.docx` 做格式检查的难点不在规则本身，而在**把 Word 的隐式语义还原为显式结构**：级联求值、域的派生性、文档序的容器嵌套。三者之中任何一个处理不当，诊断的定位或结论就会失真。

后续方向按收益排序：

1. **域重算比对**：检测"编号已过期"，这是当前最大的能力缺口；
2. **主题字体解析**：解决 `theme:` 前缀导致的静默跳过；
3. **增量检查**：基于 `w:paraId` 做跨版本定位，支持 CI 中只报新增问题。

```bash
pip install git+https://github.com/Asgoresmilelong/docxaudit
docxaudit check report.docx --json
```

仓库：https://github.com/Asgoresmilelong/docxaudit
