# docxaudit

[![ci](https://github.com/Asgoresmilelong/docxaudit/actions/workflows/ci.yml/badge.svg)](https://github.com/Asgoresmilelong/docxaudit/actions/workflows/ci.yml)
[![license: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![python](https://img.shields.io/badge/python-3.9%2B-blue.svg)](pyproject.toml)

**Word 文档格式体检**：把 .docx 里那些「肉眼要翻一百页才能发现」的格式问题，一条命令查出来。

```bash
docxaudit check 毕业论文.docx
```

```
docxaudit · sample-broken.docx
[R1 错误] 第 4 段：标题层级跳跃：从 1 级直接跳到 3 级
    … 1.1.1 研究背景
    → 中间缺一级标题会导致导航窗格和自动目录结构错乱。
[R3 错误] 第 5 段：正文引用了「图 6」，但全文没有这个编号的题注
    … 如图 6 所示，杆顶位移随高度增加而迅速增大，这一结论与现场实测一致。
    → 图号写错，或题注被删除后正文未同步修改。
[R2 警告] 第 7 段：图 编号不连续：出现 图 3，按顺序应为 图 2
    … 图 3 弯矩与应力沿杆长分布
    → 手动改过编号，或删图后没有更新域。
[R5 警告] 第 8 段：正文中文字体应为「宋体」，此段为「黑体」
    → 全文正文应统一字体；标题和题注可以不同。
[R6 警告] 第 9 段：段落间距为「固定行距 30.0 磅，无首行缩进」，多数正文为「固定行距 22.0 磅，首行缩进 24.0 磅」
    → 行距/缩进不一致在打印稿上非常明显，且会影响页数统计。
[R4 提示] 第 13 段：有 2 条参考文献在正文中从未被引用：[2]、[3]

共 6 条：错误 2 · 警告 3 · 提示 1
```

这段输出不是编的：`python examples/make_sample.py` 会生成一份埋了 6 处典型错误的文档，
再跑 `docxaudit check` 就能得到上面的结果（CI 里也是这么自检的）。

## 为什么不是"再写一个 Word 排版工具"

市面上多数工具做的是**改**格式（套模板、刷字体）；`docxaudit` 做的是**查**：

1. **不修改你的文件**，只读，跑完文档一个字节都没变；
2. **给出原因和位置**，不是“格式不规范”这种废话，而是“第 7 段，图 3，应该是图 2”；
3. **能进 CI**：`--json` 输出结构化结果，错误数超阈值就返回非零退出码。

## 安装

```bash
pip install git+https://github.com/Asgoresmilelong/docxaudit
```

或从源码：

```bash
git clone https://github.com/Asgoresmilelong/docxaudit
cd docxaudit
pip install -e ".[dev]"
```

只依赖 `lxml`；测试用 `python-docx` 现场生成夹具，不需要任何样本文件。（PyPI 发布后可简化为 `pip install docxaudit`。）

## 用法

```bash
docxaudit check 报告.docx                  # 人看的报告
docxaudit check 报告.docx --json           # 机器读的 JSON
docxaudit check 报告.docx --ignore R5,R6   # 跳过字体和行距（往往是模板决定的）
docxaudit check 报告.docx --no-fail        # 只出报告，不让 CI 失败
docxaudit rules                            # 列出规则
```

退出码：`0` 正常，`1` 有错误（可用 `--max-errors N` 调阈值），`2` 文件读不了。

Python API：

```python
from docxaudit import check

for f in check("报告.docx"):
    print(f.rule, f.severity, f.para_no, f.message)
```

## 规则

| ID | 名称 | 查什么 | 严重级 |
|---|---|---|---|
| R1 | 标题层级 | 首个标题不是 1 级；层级跳跃（1 级直接跳到 3 级） | 错误 / 警告 |
| R2 | 题注编号 | 图/表编号重复、跳号；支持 `图 3-1` 章节式编号 | 错误 / 警告 |
| R3 | 交叉引用 | 正文写「如图 6 所示」但全文没有「图 6」 | 错误 |
| R4 | 参考文献 | 正文引用了 `[7]` 但文献表没有；文献表条目从未被引用 | 错误 / 提示 |
| R5 | 中文字体 | 正文中文字体与多数段落不一致；同一段内混用多种中文字体 | 警告 |
| R6 | 行距缩进 | 正文行距 / 首行缩进与多数段落不一致 | 警告 |

规则编号稳定，`--ignore R5` 这种写法可以长期写在 CI 配置里。

## 技术说明：为什么必须直接读 OOXML

`python-docx` 的高层 API 够不到这几件事，而它们恰好是格式检查的核心：

1. **属性继承**：段落的标题级别可能写在段落 `pPr`、样式、乃至于样式基类上；字体同理，最后兜底在 `docDefaults/rPrDefault`。要"有效值"，必须自己沿 `basedOn` 链回溯。
2. **域代码**：题注编号和交叉引用在 Word 里是 `SEQ` / `REF` 域，值存在 `w:instrText` 和缓存结果里。直接读缓存文本只能拿到"显示出来的数字"，判断不了它是不是自动生成、更新过没有。
3. **结构顺序**：`w:tbl` 里的段落也要算进段落序号，且要跟正文段落保持文档顺序；`w:sdt`（内容控件）会包住正文，直接遍历 `body/*` 会漏。
4. **修订痕迹**：`w:del` 里被删掉的文字不能计入正文，否则字数、引用统计全错。

`docxaudit` 因此自己写了一层解析（`loader.py`），用 `lxml` 读 `word/document.xml` 和 `word/styles.xml`，把结果收敛成 `Paragraph` / `Run` 结构，规则层只面对这些结构，不碰 XML。

## 局限（说在前面）

- 只支持 `.docx`；`.doc` 请先用 Word / LibreOffice 另存。
- 无法解析主题字体（`theme:minorEastAsia` 这类引用），此时会退回样式或 docDefaults 的值。
- R5 / R6 用"多数派"判断，正文段落不足 5 段时不报告，避免把特殊排版误判成错误。
- 参考文献只认 `[n]` 形式的顺序编码制；著者-出版年制暂不支持。

## 开发

```bash
python -m unittest discover -s tests -t .   # 27 个用例
python examples/make_sample.py              # 生成埋了 6 处错误的示例文档
```

## License

MIT
