
## ⚙️ 核心实现原理

### 1. Office OpenXML 文件（Excel / PPT / Word）
- 使用 `JSZip` 解压文件，读取 `docProps/core.xml` 和 `docProps/app.xml`。
- 通过 DOM 解析和修改相应的 XML 节点（例如 `<dc:creator>`、`<cp:keywords>`）。
- 重新压缩并生成 Blob，保持原始扩展名和 MIME 类型。

### 2. CSV 文件
- 利用 `SheetJS` 将 CSV 解析为工作簿对象。
- 手动为 `workbook.Props` 设置作者、标题等属性。
- 输出为 XLSX 格式的 Blob。

### 3. PDF 文件
- 使用 `pdf-lib` 加载原始 PDF。
- 调用 `setAuthor`、`setTitle`、`setSubject`、`setKeywords`（关键词需传入数组）。
- 保存并返回新的 PDF Blob。

### 4. 批量打包
- 所有处理结果存入 `JSZip` 实例。
- 最终调用 `generateAsync` 生成 ZIP 文件，使用 `FileSaver` 触发下载。

## ⚠️ 注意事项

- **CSV 转换**：CSV 文件本身不包含文档属性，因此会**转换为 XLSX 格式**输出，文件名后缀变为 `.xlsx`。
- **宏文件支持**：`.xlsm`、`.pptm`、`.docm` 中的宏代码不会被修改或损坏，属性修改仅影响元数据区域。
- **PDF 关键词格式**：`pdf-lib` 要求关键词为字符串数组，工具会自动将输入的关键词（按中英文逗号或空格分割）转为数组。
- **浏览器兼容性**：需要支持 `File API`、`Blob`、`Promise`、`ArrayBuffer` 的现代浏览器。不支持 IE。
- **文件大小限制**：受浏览器内存限制，建议单个文件不超过 100 MB，总文件数不超过 50 个（视内存而定）。

## 📄 开源许可

本项目采用 **MIT 许可证**。您可以自由使用、修改和分发，但需保留原始版权声明。

