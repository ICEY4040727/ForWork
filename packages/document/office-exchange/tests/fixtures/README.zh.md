# Office Exchange 测试夹具

[English](README.md) | 中文

供 `@forwork/office-exchange` 的往返、映射、保留与不支持特性检测测试使用的确定性二进制夹具。语料库包含 34 个夹具，覆盖 `specs/002-in-place-ooxml-save` 中 FR-004 的每个构造类别。

| 文件 | 用途 |
| --- | --- |
| `roundtrip.xlsx` | 可编辑工作簿：包含带粗体表头的 `Sales` 工作表、三行数据、`=SUM(B2:B4)`、一个合并单元格和 `$#,##0.00` 货币格式；`Notes` 工作表包含文本单元格。检测结果干净。 |
| `roundtrip.docx` | 最小 OOXML 文档：一个 `Heading1` 段落、一个带粗体 run 的普通段落，以及一个 2x2 表格。不含宏、批注、修订、页眉、绘图或嵌入对象。检测结果干净。 |
| `preserve-image.docx` | 带内嵌 JPEG 的内联图片（`r:embed`）、一个 `word/media/` 部件、对应图片关系以及 JPEG 内容类型。检测结果必须保持干净，图片、媒体字节与关系必须经受保存。 |
| `preserve-external-image.docx` | 链接到外部目标（`r:link`、`TargetMode="External"`）且没有媒体部件的内联图片。检测结果必须保持干净，图片与外部关系必须经受保存。 |
| `preserve-notes.docx` | 正文含 `w:footnoteReference` 与 `w:endnoteReference` 标记，并带有 `word/footnotes.xml`、`word/endnotes.xml` 及其关系与内容类型。检测结果必须保持干净，注释部件与引用 run 必须经受保存。 |
| `preserve-header-image.docx` | 含内联 PNG 的页眉部件、页眉关系以及页眉到媒体的关系。检测结果必须保持干净，页眉图片必须经受保存。 |
| `preserve-footer-image.docx` | 含内联 PNG 的页脚部件、页脚关系以及页脚到媒体的关系。检测结果必须保持干净，页脚图片必须经受保存。 |
| `preserve-note-image.docx` | 脚注内含内联 PNG，并带脚注关系与媒体部件；另有一个文本尾注。检测结果必须保持干净，注释图片与部件必须经受保存。 |
| `preserve-multi-media.docx` | 同一段落内含 PNG、JPEG 与 EMF 内联图片，三条图片关系与三个媒体部件。检测结果必须保持干净，每个媒体条目必须经受保存。 |
| `preserve-table-cell-image.docx` | 表格单元格内的内联 PNG。检测结果必须保持干净，单元格图片、其关系与媒体部件必须经受保存。 |
| `offset-shift.docx` | 使用 CRLF 的文档，在两个段落后放置含内联 PNG 的段落，使保留 run 从非零字节偏移开始。检测结果必须保持干净，且在别处编辑后图片必须保留。 |
| `preserve-id-collision.docx` | 保留图片使用 `rId2`、页眉使用 `rId3`，与会重新生成时通常分配的 id 冲突。检测结果必须保持干净，且每条关系保留原始 id。 |
| `preserve-comment.docx` | 正文含 `w:commentRangeStart`、`w:commentRangeEnd` 与 `w:commentReference`，并带 `word/comments.xml`、`word/commentsExtended.xml`、`word/commentsIds.xml` 与 `word/people.xml` 及其关系与内容类型。检测必须报告 `comment`，且线程化批注部件必须经受保存。 |
| `preserve-chart.docx` | 正文含图表绘图与 `word/charts/chart1.xml`、其关系与内容类型。图表部件必须经受保存。 |
| `preserve-embedded-object.docx` | 正文含 VML `w:object` 预览与 `word/embeddings/oleObject1.bin`、其 OLE 关系、图片关系与内容类型。检测必须报告 `embedded-object`。 |
| `preserve-shape.docx` | 正文含 VML `w:pict` 形状与 `word/drawings/vmlDrawing1.vml`、其关系与内容类型。检测必须报告 `shape`。 |
| `preserve-external-link.docx` | 正文中指向外部目标（`TargetMode="External"`）的超链接。检测结果必须保持干净，链接目标与关系必须经受保存。 |
| `preserve-field.docx` | 正文含 `w:fldSimple` 域以及由 `w:fldChar` 与 `w:instrText` 组成的复杂域。域 run 必须经受保存。 |
| `preserve-bookmark.docx` | 正文含 `w:bookmarkStart`、`w:bookmarkEnd` 与使用 `w:anchor` 的 `w:hyperlink`。书签标记必须经受保存。 |
| `preserve-content-control.docx` | 正文含包裹段落的 `w:sdt` 内容控件。内容控件必须经受保存。 |
| `preserve-unmodeled-parts.docx` | 文档加 `word/theme/theme1.xml`、`word/fontTable.xml`、`word/webSettings.xml`、`docProps/core.xml`、`docProps/app.xml`、`customXml/item1.xml` 与 `customXml/itemProps1.xml`，含其关系与内容类型。每个未建模部件必须经受保存。 |
| `preserve-pivot-table.xlsx` | 含数据透视表、透视缓存定义与记录、其关系与内容类型的工作簿。所有透视条目必须经受保存。 |
| `preserve-excel-table.xlsx` | 含 Excel 表格部件、其关系与内容类型的工作表。表格部件必须经受保存。 |
| `preserve-defined-names.xlsx` | 含常量定义名称、区域定义名称与打印区域定义名称的工作簿。定义名称必须经受保存。 |
| `preserve-calc-chain.xlsx` | 含公式单元格与 `xl/calcChain.xml`、其关系与内容类型的工作簿。计算链必须经受保存。 |
| `macro.docm` | 宏启用文档，含宏启用的主内容类型、`word/vbaProject.bin`、其内容类型覆盖与关系。检测必须报告 `macro`，且保存不得改变格式。 |
| `macro.xlsm` | 宏启用工作簿，含宏启用的主内容类型、`xl/vbaProject.bin`、其内容类型覆盖与关系。检测必须报告 `macro`，且保存不得改变格式。 |
| `container-stored.docx` | `roundtrip.docx` 包，其中 `word/styles.xml` 以存储方式写入。无操作保存必须保留存储方法与本地记录。 |
| `container-comment.docx` | `roundtrip.docx` 包，带非空归档注释。无操作保存必须保留注释字节。 |
| `container-descriptor.docx` | `roundtrip.docx` 包，其条目使用非 ZIP64 数据描述符（通用标志位 3）。无操作保存必须保留每个描述符。 |
| `unsupported-macro.docx` | `roundtrip.docx` 加上一个 `word/vbaProject.bin` 条目、其内容类型覆盖与关系。检测必须报告 `macro`。 |
| `unsupported-tracked.docx` | `roundtrip.docx` 在 `word/document.xml` 中带有 `<w:ins>` 与 `<w:del>` 修订。检测必须报告 `tracked-changes`。 |
| `unsupported-header.docx` | `roundtrip.docx` 加上带 `Confidential header` 文本的 `word/header1.xml` 及其 content-type 覆盖与关系。检测不得报告任何内容，且保存后页眉文本必须保留。 |
| `legacy.xls` | 由 SheetJS 写入的 BIFF8 工作簿，包含一个 `Legacy` 工作表。用于演练旧版转换。 |
| `generate.mjs` | 重新生成上述所有夹具。 |

## 重新生成

```sh
node packages/document/office-exchange/tests/fixtures/generate.mjs
```

`generate.mjs` 从 `packages/client/ui-sidebar-documentpreview`（它们所属的 workspace 包）解析 `exceljs` 与 `xlsx`，从 `packages/document/office-exchange` 解析 `fflate`，并回退到仓库根目录。

## 确定性

```sh
node packages/document/office-exchange/tests/fixtures/generate.mjs --check
```

`--check` 在内存中重建每个夹具并与已检入的二进制比较，不写入文件。任何不一致的夹具会输出 `missing:` 或 `drift:` 并以非零状态退出；干净的工作树会输出 `fixtures are deterministic (34 checked)`。编辑 `generate.mjs` 或任意夹具断言后请运行它。

输出可逐字节复现：归档条目经过排序，并使用固定的 `FIXED_DATE` 时间戳写入，工作簿属性固定为同一日期，且不使用随机 id 或当前时间值。在干净的工作树上重新运行脚本会得到相同的哈希。

不要手工编辑生成的二进制文件；请编辑 `generate.mjs` 并重新生成。`generate.mjs` 中的 `FIXED_DATE` 全局变量是确定性的唯一来源，只有在有意为所有夹具重建基线时才可更改。

如果安装的 SheetJS 构建无法写入 `bookType: 'xls'`，脚本会跳过 `legacy.xls`，删除任何过期副本，并打印原因。
