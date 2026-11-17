# Office Exchange 测试夹具

[English](README.md) | 中文

供 `@forwork/office-exchange` 的往返、映射与不支持特性检测测试使用的确定性二进制夹具。

| 文件 | 用途 |
| --- | --- |
| `roundtrip.xlsx` | 可编辑工作簿：包含带粗体表头的 `Sales` 工作表、三行数据、`=SUM(B2:B4)`、一个合并单元格和 `$#,##0.00` 货币格式；`Notes` 工作表包含文本单元格。检测结果干净。 |
| `roundtrip.docx` | 最小 OOXML 文档：一个 `Heading1` 段落、一个带粗体 run 的普通段落，以及一个 2x2 表格。不含宏、批注、修订、页眉、绘图或嵌入对象。检测结果干净。 |
| `unsupported-macro.docx` | `roundtrip.docx` 加上一个 `word/vbaProject.bin` 条目。检测必须报告 `macro`。 |
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

输出可逐字节复现：归档条目经过排序，并使用固定的 `FIXED_DATE` 时间戳写入，工作簿属性固定为同一日期，且不使用随机 id 或当前时间值。在干净的工作树上重新运行脚本会得到相同的哈希。

不要手工编辑生成的二进制文件；请编辑 `generate.mjs` 并重新生成。`generate.mjs` 中的 `FIXED_DATE` 全局变量是确定性的唯一来源，只有在有意为所有夹具重建基线时才可更改。

如果安装的 SheetJS 构建无法写入 `bookType: 'xls'`，脚本会跳过 `legacy.xls`，删除任何过期副本，并打印原因。
