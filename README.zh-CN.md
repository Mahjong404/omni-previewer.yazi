# omni-previewer.yazi

[English](./README.md) | **简体中文**

面向 [Yazi](https://github.com/sxyazi/yazi) 的一站式文档预览插件，重点覆盖 Windows + Office 格式。

## 功能

### Word 文档（`docx`、`doc`、`docm`、`dotx`、`dotm`、`rtf`）

- **渐进式预览**：先秒出提取文本，转换完成后自动替换为渲染页面图。
- 经 Microsoft Word → PDF → `pdftoppm` 渲染，保留排版、图片、表格。
- `J`/`K` 翻页，`F6` 在页面图/纯文本模式间切换。
- 隐藏的可复用 Word 服务让文档热转换保持在约 0.5 秒，空闲 10 分钟自动退出。
- 安全：只读打开、宏强制禁用（`AutomationSecurity=3`）、拒绝 ActiveX/`vbaProject`/可达外链，绝不触碰你正在使用的 Word 会话。
- 有界缓存（`%LOCALAPPDATA%\yazi\docx-pages`）：总量 64 MiB，每文档最多 3 张页图，3 天未用自动清理。

### PDF

- `pdftoppm` 按预览区尺寸直接渲染——超大页面（如巨型思维导图导出）不再触发图像尺寸超限。
- `J`/`K` 翻页；状态栏显示 `p.N/M` 页码。
- 只缓存页图，不复制源 PDF。

### Excel 工作簿（`xlsx`）

- 带边框的对齐表格预览，中英文宽度计算正确。
- 真实单元格样式转 ANSI：填充色、字体色、加粗。
- 合并单元格跨列渲染，居中保留。
- 按文件指纹缓存输出，滚动零延迟（`%LOCALAPPDATA%\yazi\xlsx-cache`）。

## 依赖

- Windows + 已安装 Microsoft Word（Word 管线）
- Python 3，含 `openpyxl`、`pywin32`、`psutil`（`render.py`/`xlsx.py` 通过 `python.exe` 运行）
- [Poppler](https://github.com/oschwartz10612/poppler-windows)（`pdftoppm.exe` / `pdfinfo.exe`）
- [Pandoc](https://pandoc.org/)（文本回退，依赖 `docx-preview.yazi` 插件）

## 安装

```sh
ya pkg add Mahjong404/omni-previewer.yazi
```

`yazi.toml`：

```toml
[[plugin.prepend_previewers]]
url = "*.{docx,DOCX,doc,DOC,docm,DOCM,dotx,DOTX,dotm,DOTM,rtf,RTF}"
run = "omni-previewer"

[[plugin.prepend_previewers]]
url = "*.{xlsx,XLSX}"
run = "omni-previewer"

[[plugin.prepend_previewers]]
url = "*.{pdf,PDF}"
run = "omni-previewer"

[[plugin.prepend_preloaders]]
url = "*.{docx,DOCX,doc,DOC,docm,DOCM,dotx,DOTX,dotm,DOTM,rtf,RTF,pdf,PDF}"
run = "omni-previewer"
```

`keymap.toml`（可选，图文/纯文本切换）：

```toml
[[mgr.prepend_keymap]]
on = "<F6>"
run = "plugin omni-previewer"
desc = "切换 Word 图文/纯文本预览"
```

将 `main.lua` 顶部的 `PYTHON` 常量改成你的解释器路径。

## 路线图

- 更多格式：`pptx`、`ppt`、`xls`、`csv`、`md` 快速预览
- `.ipynb` 渲染；`.git` 仓库信息（log/status 摘要）支持
- 归档（`zip`/`rar`）预览增强：提速 + 更丰富的列表
- macOS/Linux 支持：Word 管线目前依赖 Windows COM，跨平台需改用 LibreOffice 无界面转换（`soffice --convert-to pdf`），xlsx 管线本身跨平台可用
- `docx-preview` 回退依赖解耦（内置文本提取）

## License

MIT
