# doc-sorting

> 给中文课程报告 / 实验报告的 `.docx` 补内容、加图注表注、统一排版 —— 并且**不破坏**文档里已有的截图、表格和题干。

`doc-sorting` 是一个面向中文教学文档的排版整理工具集。它解决的不是「怎么改 Word」，而是**怎么改才不会把文档改坏**。

```text
报告.docx  ──▶  扫描 → 补空 → 规范化 → 校验 ──▶  clean.docx
                 只读    编辑器SDK   XML直改    断言+目视
```

---

## 目录

- [为什么需要它](#为什么需要它)
- [核心设计原则](#核心设计原则)
- [特性](#特性)
- [环境要求](#环境要求)
- [安装](#安装)
- [快速开始](#快速开始)
- [工作流](#工作流)
- [脚本速查](#脚本速查)
- [规则文件](#规则文件)
- [目录结构](#目录结构)
- [五个致命坑](#五个致命坑)
- [边界与限制](#边界与限制)

---

## 为什么需要它

整理一份已有的实验报告，常见的诉求是这样的：

- 「`文字说明：` 后面空着，帮我补上」
- 「每张程序截图下面都要有图注」
- 「图注和表注字号调小一点、居中」
- 「正文每个自然段首行缩进两字符」
- 「`运行结果…` 这句话跟上一段粘在一起了，该换行」

用 `python-docx` 盲写、或者对 Word 调 COM 全篇刷格式，结果往往是：截图跑位、表格列宽被改、题干被一起放大、段落标记的字体属性丢失。

真正致命的是**编辑器 SDK 的坐标模型**：它按字符流计数，每插入一个字，后面所有偏移量都会失效。一次批量改 140 个段落，很容易误伤 149 个 —— 包括图注、标签、表格单元格和空段。

这个 skill 把整套踩过的坑固化成脚本和规则。

## 核心设计原则

> **插入内容走编辑器 SDK，套用格式走直接改 XML。**
>
> 绝不用编辑器坐标做批量格式化，也绝不把编辑器的 `save_file` 当作权威写回。

| 任务 | 走哪条路 | 为什么 |
|---|---|---|
| 读结构、定位锚点 | `doc_find` / `docx_scan.py` | 坐标由工具给出，不手算 |
| 插入段落、图注、标签 | 编辑器 SDK（`docx_insert.py`） | 需要文档的实时视图与段落属性继承 |
| 批量改字号 / 加粗 / 对齐 / 缩进 | 直接改 `document.xml`（`docx_format.py`） | 按**内容签名**匹配，完全不用坐标，天然幂等 |
| 证明改对了 | `docx_verify.py` + PDF 目视 | 断言抓结构，渲染抓版式 |

## 特性

- **坐标无关的批量格式化** —— 规则按正则 / 精确文本 / 样式 ID / 当前字号 / 是否含图来选中段落，写入顺序不影响结果，可以反复跑。
- **锚点式安全插入** —— 计划里的操作按 `idx` **降序**执行，前面的插入不会让后面的坐标失效。
- **图注自动落位** —— `after_image: <index>` 直接把图注插到图片下一段。
- **修复合段** —— 用空段落切断「`运行结果…` 粘在上一段末尾」的粘连。
- **幂等 + 自动备份** —— `--inplace` 自动生成 `.docx.bak`，`--dry-run` 先看再写。
- **双层校验** —— XML 层断言（每规则命中数、格式值、命名空间完整性、残留 `&#13;`）＋ Word 导出 PDF 逐页目视。
- **可恢复** —— `--clear-stray-indent` 专门清理被编辑器写坏后残留的「假首行缩进」。

## 环境要求

| 项 | 要求 | 说明 |
|---|---|---|
| 操作系统 | Windows | 依赖 Word COM 导出 PDF，以及本地 Office 编辑器 SDK |
| Python | 3.9+ | 建议用独立虚拟环境 |
| Python 包 | `lxml` `pypdfium2` `pillow` | 前两个必需；Pillow 用于图片拼图核对 |
| Office | Microsoft Word（`docx2pdf.ps1` 用） | 或用 WPS 手动另存为 PDF |
| 编辑器 SDK | 本地 Office 编辑能力（edsdk） | `edsdk_client.py --where` 可探测；缺失时仅「插入」阶段不可用，格式化与校验仍可跑 |

```bash
pip install lxml pypdfium2 pillow
```

## 安装

作为 WorkBuddy skill 安装（推荐，路径即约定）：

```bash
git clone git@github.com:icdi24/doc-sorting.git ~/.workbuddy/skills/doc-sorting
```

或者把整个目录拷进 `~/.workbuddy/skills/doc-sorting`。克隆后重启会话，`SKILL.md` 里的 `description` 会命中「补全实验报告 / 加图注 / 首行缩进 / 排版美化」等说法，自动触发。

## 快速开始

```bash
PY=python
S=~/.workbuddy/skills/doc-sorting/scripts

# 1) 先看清楚文档长什么样（只读，不改文件）
$PY $S/docx_scan.py 报告.docx > scan.tsv
$PY $S/docx_scan.py 报告.docx --images          # 只看图片段落
$PY $S/docx_scan.py 报告.docx --only "文字说明"  # 找空着的标签

# 2) 拿锚点，写插入计划，先干跑
$PY $S/docx_insert.py anchors 报告.docx "文字说明：" "运行结果："
$PY $S/docx_insert.py plan plan.json            # dry run
$PY $S/docx_insert.py run  plan.json            # 执行并保存

# 3) 规则化排版：改好 style_id / size_pt 后先 dry-run
cp $S/../assets/course-report.rules.json ./rules.json
$PY $S/docx_format.py 报告.docx --spec rules.json --dry-run
$PY $S/docx_format.py 报告.docx --spec rules.json --out clean.docx

# 4) 两层校验，断言必须过
$PY $S/docx_verify.py clean.docx --spec rules.json --strict-stray

# 5) 看真实分页
powershell -ExecutionPolicy Bypass -File $S/docx2pdf.ps1 -Src clean.docx -Pdf clean.pdf
$PY $S/pdf2png.py clean.pdf pages/
```

> `docx_scan.py` 的输出里，**真实 `style_id` 和字号**是第 3 步必须用来改规则的值 —— 不同模板的正文样式 id 不同（示例模板是 `00000e`）。

## 工作流

### 阶段 0 — 勘察（只读）

跑 `docx_scan.py` 建立段落模型：文本 + 字号 + 加粗 + 对齐 + 缩进 + 是否含图 + 近似偏移。

同时做两件人工判断：

1. 把 `word/media/*` 抽出来拼成带文件名的联系表（Pillow 网格），**逐张看截图**。图注必须描述截图里真实的内容。
2. 若有配套源码目录，读一遍。补写的文字必须和真实代码一致 —— 不要编造类名、方法名、输出值。

### 阶段 1 — 补内容（编辑器 SDK，插入）

```json
{ "file_id": "D:/报告.docx", "save": true,
  "ops": [
    {"after_image": 5429, "text": "图 2-5 完善后的 CSSE 类代码",
     "para": {"jc": "center"}, "font": {"size_pt": 9}},
    {"idx": 14702, "text": "运行结果输出了 5 个行政办公室。"}
  ] }
```

- `after_image` → 插到 `image.index + 1`，图注紧跟图片。
- `idx` → 文本落点；用 `doc_find` 拿现成坐标，不要手算。
- 合段修复：在粘连处插一个空段落 `{"idx": <位置>, "text": ""}`，前文继承原段落属性。
- 表格填内容用编辑器的 `doc_set_table_cells`（一次调用传全部 `{row, col, text}`），表注用普通插入放在表格上方。

### 阶段 2 — 规范化排版（XML 直改，坐标无关）

规则按**内容签名**选段落，不按位置。跑多少遍结果都一样。

```bash
$PY $S/docx_format.py 报告.docx --spec rules.json --out clean.docx
```

### 阶段 3 — 校验（两层）

| 层 | 手段 | 抓什么 |
|---|---|---|
| 结构层 | `docx_verify.py --strict-stray` | 每规则命中数、格式值、`mc:Ignorable` 前缀丢失、残留 `&#13;`、空 `w:rPr`、假缩进。**退出码为 0 才允许落盘** |
| 视觉层 | `docx2pdf.ps1` → `pdf2png.py` | 真实分页、图注位置、表格列宽、题干是否被误放大 |

XML 断言看不出分页 —— **永远要看渲染出来的页面**。

### 阶段 4 — 落盘

`--inplace` 写回（自动 `.bak`），或把 `clean.docx` 覆盖到目标路径，然后重新在编辑器里打开刷新预览。

## 脚本速查

| 脚本 | 作用 | 关键参数 |
|---|---|---|
| `scripts/docx_scan.py` | 只读段落模型，不需要编辑器 | `--images` 只看图片段 · `--only <文本>` 过滤 · `--find <文本>` 定位 · `--classify` 分类统计 · `--base N` 校准偏移 · `--json` |
| `scripts/docx_insert.py` | 锚点/图片定位 + 有序安全插入 | 子命令 `anchors` `images` `plan` `run` `captions`；`--image-offset`、`--captions-file`、`--dry-run` |
| `scripts/docx_format.py` | 规则驱动的 lxml 格式化 | `--spec` 规则文件 · `--out` · `--inplace` · `--dry-run` · `--clear-stray-indent` |
| `scripts/docx_verify.py` | XML 层断言，退出码即门禁 | `--spec`（必需） · `--strict-stray` · `--max-report N` |
| `scripts/docx2pdf.ps1` | Word COM 导出 PDF（`wdFormatPDF = 17`） | `-Src` `-Pdf` |
| `scripts/pdf2png.py` | PDF 栅格化（pypdfium2） | `--scale` · `--pages` |
| `scripts/edsdk_client.py` | 驱动本地编辑器 SDK | `--where` 探测 · `--list` · `--schema <tool>` · `--pool` · 任意工具直调 |

## 规则文件

`assets/course-report.rules.json` 是起步模板。**用前必改**：`body` 规则的 `style_id` / `size_pt` 要换成你文档的真实值。

```json
{
  "defaults": { "base_pt": 10.5 },
  "rules": [
    {
      "name": "figure-caption",
      "match": { "text_regex": "^图\\s*\\d+-\\d+" },
      "format": { "size_pt": 9, "bold": false, "align": "center", "clear_indent": true }
    },
    {
      "name": "body",
      "match": { "style_id": "00000e", "size_pt": 10.5 },
      "format": { "first_line_indent_chars": 2 }
    }
  ]
}
```

**`match` 可用键**

| 键 | 含义 |
|---|---|
| `text_regex` | 段落文本正则匹配 |
| `text_in` | 文本等于列表中的某一项 |
| `style_id` | 段落样式 ID（`docx_scan.py` 里能看到真实值） |
| `size_pt` | 段落当前字号 |
| `has_image` | 段落是否含图片 |

**`format` 可用键**

| 键 | 含义 |
|---|---|
| `size_pt` | 字号（pt） |
| `bold` | 加粗 |
| `align` | 对齐：`left` / `center` / `right` / `both` |
| `clear_indent` | 清除缩进 |
| `first_line_indent_chars` | 首行缩进字符数（按 `base_pt` 换算 twips） |

**规则按顺序匹配，第一个命中的生效** —— 越具体的规则越要放前面。

## 目录结构

```text
doc-sorting/
├── SKILL.md                          # skill 定义与硬性规则
├── README.md
├── assets/
│   └── course-report.rules.json      # 起步排版规则模板
├── references/
│   ├── workflow.md                   # 端到端流程 + 交付前自查清单
│   ├── edsdk.md                      # 编辑器工具清单、idx 语义、偏移模型
│   ├── pitfalls.md                   # 五个文档损坏坑与恢复次序
│   └── layout-spec.md                # 中文字号对照、缩进 twips、图注命名、补写语气
└── scripts/
    ├── docx_scan.py
    ├── docx_insert.py
    ├── docx_format.py
    ├── docx_verify.py
    ├── edsdk_client.py
    ├── docx2pdf.ps1
    └── pdf2png.py
```

## 五个致命坑

以下每一条都在真实文档上造成过损坏，详见 [`references/pitfalls.md`](references/pitfalls.md)。

| # | 坑 | 后果 | 对策 |
|---|---|---|---|
| 0 | 编辑器的 `save_file` 会整篇重新序列化 | 594 → 509 段，段落标记 `rPr` 全丢，表格单元格被重套缩进 | 权威写回一律走 XML；编辑器只读和插 |
| 1 | 用编辑器坐标批量改格式 | 140 个目标误伤成 149 个 + 33 处题干被放大 | 批量格式走 `docx_format.py`；插入按 `idx` 降序 |
| 2 | 用 `ElementTree` 重写 `document.xml` | Word 报「文件可能已损坏」（`ns0:` 前缀与 `mc:Ignorable` 对不上） | 一律用 `lxml`，只替换单个 zip 条目 |
| 3 | 手算偏移与编辑器坐标漂移 | 写入位置错位 | 先 `doc_find` 拿坐标；`--base N` 反推基线 |
| 4 | 远端 `doc_to_image` 不可用 | 无法目视校验 | 本机 Word 导出 PDF 再栅格化 |

## 边界与限制

- **只面向 `.docx`**，不处理 `.doc` / `.wps`（需先另存）。
- **Windows 专用**，`docx2pdf.ps1` 依赖 Word COM。
- 插入阶段**依赖本地编辑器 SDK**；没有该能力时，只能做扫描 / 格式化 / 校验。
- 规则是**正则与样式匹配**的启发式，不是语义理解 —— 换一份新模板，必须重新勘察并调整规则。
- 不会替你改题干或结论。**当截图内容与题目要求矛盾时**（例如题目写 `3+9i`，而 `main` 里是 `3+2i`），脚本会报告差异，由人决定怎么处理。

---

## 致谢

这套规则与坑位清单来自真实的中文实验报告整理过程，每一条都以「文档被改坏」为代价换来的。
