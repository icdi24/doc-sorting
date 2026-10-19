# 端到端流程

目标：把一份**已有图片与代码**的中文实验报告 / 课程报告，补全空缺文字与表格，
加图注表注，并规范化排版，且**不损坏原有内容**。

## 阶段 0 — 勘察（只读）

```bash
PY=python
S=~/.workbuddy/skills/doc-sorting/scripts

# 段落模型：文本 + 字号 + 加粗 + 对齐 + 缩进 + 是否含图
$PY $S/docx_scan.py "D:/report.docx" > scan.tsv

# 只看图片段与标签段
$PY $S/docx_scan.py "D:/report.docx" --images
$PY $S/docx_scan.py "D:/report.docx" --only "文字说明"
```

先把结构摸清：大纲、表格数量、图片数量、空缺位置（`程序截图：` / `运行结果：` /
`文字说明：` 后面是否为空）。

**同时把截图内容看一遍**——图注必须按截图实际内容命名。从 `word/media/` 抽出图片
拼成 contact sheet 目视，比逐张看好得多：

```python
# 从 docx 里取图并拼图（Pillow）
import zipfile, math
from PIL import Image, ImageDraw, ImageFont
z = zipfile.ZipFile("D:/report.docx")
names = [n for n in z.namelist() if n.startswith("word/media/")]
# ...按网格 paste，并在每格左上角标注文件名，便于对号入座
```

同时读源码目录（如果报告配套有 `src/`）——补写的文字说明必须与真实代码一致，
不能凭空编造类名与方法名。

## 阶段 1 — 补内容（编辑器 SDK，插入）

用 `docx_insert.py`。**先 dry run 看顺序**，确认每个 `idx` 落在预期位置再执行。

```bash
# 1) 拿锚点坐标
$PY $S/docx_insert.py anchors "D:/report.docx" "文字说明：" "运行结果："
# 2) 拿图片坐标
$PY $S/docx_insert.py images  "D:/report.docx"
# 3) 写计划并干跑
$PY $S/docx_insert.py plan plan.json
# 4) 执行（脚本自动按 idx 降序）
$PY $S/docx_insert.py run  plan.json
```

计划里的两类操作：

* 图注：`{"after_image": <图片index>, "text": "图 3-14 xxx", ...}`
  —— 脚本会插到 `index + 1`，即图片正下方
* 补文字 / 补标签：`{"idx": <doc_find 的 begin>, "text": "...", ...}`

**补换行**（把粘在说明段末尾的「运行结果…」拆成独立段落）：
在粘连位置插一个空段落即可——`{"idx": <粘连位置>, "text": ""}`，
断点前的文本继承原段落属性，断点后成为新段落。

**补表格**：用 `doc_set_table_cells`，`cells: [{row, col, text}]`，
一次性写完整张表。表注另用插入操作放在表格上方。

## 阶段 2 — 规范排版（XML 直改，坐标无关）

复制一份起步规则文件再改：

```bash
cp ~/.workbuddy/skills/doc-sorting/assets/course-report.rules.json ./rules.json
# 改 match 里的 style_id / size_pt 以匹配本文档（先用 docx_scan.py 看真实值）
$PY $S/docx_format.py "D:/report.docx" --spec rules.json --dry-run
$PY $S/docx_format.py "D:/report.docx" --spec rules.json --out clean.docx
```

规则按**内容特征**匹配（正则 / 精确文本 / 样式 id / 当前字号 / 含图 / 空段），
不依赖坐标，因此可重复执行、幂等、不会误伤邻居。

**中文文档常见的"内容段"特征**：`pStyle` 为某个正文样式 id，且首个 run 的
`sz=21`（10.5pt）。用 `docx_scan.py` 先确认本文档的实际值，再写进 `match`。

## 阶段 3 — 校验（两层）

```bash
# XML 层断言
$PY $S/docx_verify.py clean.docx --spec rules.json --strict-stray
# 目视层：Word 导出 PDF → 逐页 PNG
powershell -ExecutionPolicy Bypass -File $S/docx2pdf.ps1 -Src clean.docx -Pdf clean.pdf
$PY $S/pdf2png.py clean.pdf pages/
```

XML 断言查：每条规则的段落数量与格式、`mc:Ignorable` 前缀完整性、
残留 `&#13;`、空 `w:rPr`。目视查：分页、图注位置、表格列宽、题干有没有被误改。

**必须看渲染结果。** 表格第一列变窄、题干三个字变大，XML 断言都发现不了。

## 阶段 4 — 落盘与复盘

* 写回前先备份（`doc_format.py --inplace` 会自动生成 `.bak`）
* 重新 `open_file` 刷新编辑器预览
* 记录：本次用到的样式 id、匹配正则、误伤与修复

## 交付前自查

- [ ] 每张图下面都有图注，编号连贯，内容描述准确
- [ ] 表注在表格上方，编号与正文引用一致
- [ ] `程序截图：`/`运行结果：`/`文字说明：` 三个标签格式统一
- [ ] 正文段落首行缩进 2 字符，题干**未**被改动
- [ ] 表格单元格没有被加缩进（第一列不换行）
- [ ] 补写的文字与真实代码/运行结果一致，没有杜撰的类名、输出值
- [ ] Word 能正常打开，PDF 分页正常
- [ ] 题目要求与代码不符之处（如题目说 3+9i，代码写 3+2i）**已向用户指出**，
      没有擅自替换用户的截图

## 一条重要纪律

如果发现**代码/截图与题目要求不符**，不要替用户改截图或改代码结果去"凑答案"。
把差异明确指出来（题目要什么、代码实际是什么、正确结果应该是多少），
由用户决定。这是唯一一处必须停下来问的地方。
