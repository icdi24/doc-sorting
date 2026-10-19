# 血泪教训：本类任务的五个致命坑

以下每一条都在真实文档上**造成过损坏**，不是理论风险。

## 0. 编辑器的 `save_file` 会整篇重新序列化（最隐蔽，也最严重）

**现象**：为了改一个段落的字号而走编辑器通道（`doc_update_text_property` +
`save_file`），保存后整篇文档被编辑器按自己的模型重写：
段落总数 594 → 509（约 85 个空段被剪掉）、所有段落的段落标记 `w:rPr` 丢失
（图注的段落标记字号从 18 变成 None）、2.4 表格的 `NO` / `可以` 单元格被重新
套上首行缩进、正文段的 `pStyle` 丢失导致按样式匹配的规则全部失配。

**根因**：编辑器保存时把内存中的文档模型序列化回 OOXML，这个 round-trip
**不保真**——它会规范化空段、丢弃它不建模的属性、并把它自己写入的
`firstLineChars`（无 `firstLine`）散布到表格单元格里。

**对策**：

* **权威写回一律走 XML**（`docx_format.py`），编辑器只用于读取和插入
* 若确实通过编辑器改了内容，保存后**必须**重跑 `docx_verify.py`——
  断言会立刻暴露 `pStyle` 丢失和表格单元格缩进
* 每次动手前先 `docx_format.py --inplace`（自动 `.bak`）或手工备份，
  这样才有可回滚的基线
* 编辑器的 `close_file --force` 会释放文件锁；WPS 预览窗口占用时会
  `PermissionError`，关掉编辑器会话后重试即可

## 1. 批量改格式用编辑器坐标 → 误伤邻居

**现象**：对 140 个段落批量调用 `doc_update_text_property` /
`doc_modify_paragraph`，结果 149 个段落被误加缩进（图注、标签、表格单元格、
空段全中），另有 33 处题干文字被误加粗放大到 12pt。

**根因**：编辑器的 DOC 坐标是**字符流偏移**。每写一次，后面所有偏移都会变。
一次批量是从同一份快照算出的坐标，写第一项之后，剩下的坐标全部失效。

**对策**：

* 批量**格式**一律走 `docx_format.py`（XML 直改，按内容匹配，完全不用坐标）
* 只在**插入**时用编辑器坐标，且必须先按 `idx` **降序**执行
* 千万不要为了"顺带设个格式"把 `range.end + 1` 传进去——如果这个 `end+1`
  跨过了段落标记，编辑器会在 `w:t` 里塞进**字面量 `\r`（`&#13;`）**，
  使该字符后的坐标整体漂移一位，后续全部错位

## 2. ElementTree 重序列化 → Word 判定"文件已损坏"

**现象**：用 `xml.etree.ElementTree` 解析 `document.xml`、改完再
`ET.tostring()` 写回，Word 打开时报「文件可能已损坏」。

**根因**：ET 把所有元素统一改写成 `ns0:`、`ns1:` 前缀，而根元素上的
`mc:Ignorable="w14 wp14 ..."` 仍然引用**原来的前缀名**。前缀未声明 →
OOXML 校验失败。

**对策**：一律用 **`lxml`**。lxml 保留元素自身的 nsmap 与原始前缀，
`etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)`
可以直接写回。写回时只替换 `word/document.xml` 这一个 zip 条目，
其余条目**原样复制**，不要整包重建。

`docx_verify.py` 会检查 `mc:Ignorable` 引用的前缀是否都已声明，作为回归防线。

## 3. 编辑器坐标与磁盘 XML 模型会漂移

**现象**：自己按 XML 算出的偏移，在编辑器多次写操作后与 `doc_find` 结果对不上。

**根因**：编辑器的 DOC 坐标模型（每个 `w:t` 计 `len(text)`，每个图片/制表符/
换行计 1，每段 +1，每个单元格 +1，每行 +1）本身是对的，但**基线**取决于文档
起点，且写了之后会变。

**对策**：

* 任何要用坐标的操作，**先 `doc_find` 拿现成坐标**，不要手算
* 需要理解结构时用 `docx_scan.py` 报告的近似偏移，仅作参考
* 必须校准时：`docx_scan.py --base N`，用已知锚点反推 `N`
  （`doc_find` 给出的 `begin` 减去模型里同一段的 `begin`）

## 4. 远端 `doc_to_image` 不可用

**现象**：`doc_to_image` 持续返回 50xxx / 437xxx 错误。

**对策**：改用本机 Word 导出 PDF 再栅格化——
`docx2pdf.ps1`（Word COM，17 = wdFormatPDF）→ `pdf2png.py`（pypdfium2）。
这也是唯一能看到**真实分页**的办法，XML 断言看不出分页和表格列宽。

## 附：中文路径与参数编码

* 编辑器 `file_id` 就是文档绝对路径，正斜杠可用
* shell 里传中文参数容易被破坏：改用 **UTF-8 的 `--json-file`**
* `edsdk._rpc('tools/call', ...)` **直接返回结果对象**，不要再取 `obj['result']`
* 无法渲染时先在纯英文路径放一份副本（如 `C:/tmp/src.docx`）再调工具

## 附：修复受损文档的次序

若已经踩坑并写坏了文档：

1. **不要**在编辑器里继续"撤销式"批量操作 —— 坐标已经不可信
2. 关闭编辑器（`close_file --force`），拿磁盘上的文件副本
3. 用 `docx_format.py --clear-stray-indent` 清掉 SDK 特征缩进
   （`w:ind` 有 `firstLineChars` 但没有 `firstLine`）
4. 用 lxml 清掉所有 `w:t` 里的 `\r`
5. 用 `docx_verify.py --strict-stray` 断言 + Word 导出 PDF 逐页目视
