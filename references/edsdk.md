# 本地 Office 编辑器 SDK（edsdk）速查

`tencent-local-office-edit` 内置技能提供了一个本地 JSON-RPC 服务，
`scripts/edsdk_client.py` 会自动定位并加载它的 `edsdk.py`。

```bash
PY=python   # 用隔离环境里的解释器
S=~/.workbuddy/skills/doc-sorting/scripts

$PY $S/edsdk_client.py --where                    # edsdk.py 所在目录
$PY $S/edsdk_client.py --pool                     # 打开的文件与 file_id
$PY $S/edsdk_client.py --list doc_                # 工具清单
$PY $S/edsdk_client.py --schema doc_insert_paragraph_with_text
$PY $S/edsdk_client.py open_file --json '{"file_path":"D:/x.docx","wait":true}'
```

`file_id` 即文档绝对路径（正斜杠可用）。写操作后**必须** `save_file` 才会落盘。

## 关键工具

### 读

| 工具 | 用途 | 返回 |
|---|---|---|
| `doc_find` | 按文本定位 | `locations[]`：`begin` `end` `paragraph_id` `row` `col` `table_id` `related_text` |
| `doc_get_images` | 列图片 | `images[]`：`index`（脚注式 DOC 偏移）`image_url` `source` |
| `doc_list_tables` | 列表格 | `table_id` `table_start/end` `row_count` `col_count` |
| `doc_get_table_info` | 单元格文本 | `block.table.cells[]`：`row` `col` `text` |
| `doc_get_paragraph_property` | 段落属性 | `alignment` `indent` `style` … |
| `doc_resolve_document_structure` | 结构树 | 段落/表格嵌套关系 |
| `get_pool_status` | 已打开文件 | `file_id` / `file_path` |

**`doc_find` 是最可靠的锚点来源** —— 拿到 `begin` 就能直接用于插入，
不必自己维护坐标模型。

### 写

| 工具 | 用途 | 要点 |
|---|---|---|
| `doc_insert_paragraph_with_text` | **一步插入段落 + 格式** | `idx` 为插入点；`text` / `runs`（段内混合格式）/ `paragraph_property` / `text_format`；返回 `next_index`、`range`、`paragraph_id` |
| `doc_insert_text` | 只插纯文本 | 不带格式 |
| `doc_modify_paragraph` | 改段落属性 | `ranges` + `jc` / `first_line_indent_chars` / 行距 |
| `doc_update_text_property` | 改字符格式 | `ranges` + `bold` / `font_size`(pt) / `font_family` / `color` |
| `doc_set_table_cells` | 批量写单元格 | `table_id` + `cells[{row,col,text}]` |
| `save_file` / `close_file` | 落盘 / 关闭 | `close_file` 加 `force:true` 丢弃未保存改动 |

### 渲染

`doc_to_image`（远端）—— 经常不可用，优先用本机 Word：
`scripts/docx2pdf.ps1` → `scripts/pdf2png.py`。

## 语义陷阱

* **`doc_insert_paragraph_with_text` 的 `idx` 是"插入点"，不是"某段之后"**
  - 文档开头：`idx = -1`（等价 0）
  - **某段之后**：`该段 range.end + 1`
  - 图片正下方：`image.index + 1`（图片占 1 个坐标单位）
  - 链式追加：用上一次返回的 `next_index`（已自动跳过 ¶）
  - ⚠️ 误用 `range.end` 会让下一段文本落在 ¶ 之前并与当前段合并
    （`"Base"` + `"Second"` → `"BaseSecond"`）
* **自动换段保护**：若 `idx` 前一个字符是普通文本，服务端会先插一个 ¶ 切断前文。
  这是"补换行"的正规做法——在粘连处传 `idx = 粘连位置` 并给空 `text`。
* **不传 `level`** 时按正文（Normal）插入，不继承上一段的标题字号。
* `first_line_indent_chars` 写的 `w:ind` **只有 `firstLineChars`**，
  没有 `firstLine`。若需精确排版，改用 `docx_format.py` 写全两个属性。
* 编辑器的 `bold` / `font_size` 等作用于 **run** 层；段落标记上的 `w:rPr`
  要用 XML 工具设置才能改变整段默认字号。

## 偏移校准（万不得已才用）

编辑器坐标 = 从 `w:body` 起点起算：

```
每个 w:t           → len(text)
w:drawing / w:pict / w:tab / w:br / w:sym / w:noBreakHyphen / w:object → 1
每个段落           → +1
每个单元格、每一行 → +1
```

用 `docx_scan.py --base N` 施加基线。取 `doc_find` 的 `begin` 减去模型里
同一段落的 `begin` 即得 `N`。**仅用于交叉验证，不要用来替代 `doc_find`。**
