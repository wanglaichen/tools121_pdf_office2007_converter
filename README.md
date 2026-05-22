# PDF Office 2007 Converter

把 PDF 每一页渲染成图片，再写入 `.docx`，用于生成 Word 2007 更容易打开的图片版 Word 文件。

## 安装

```powershell
cd F:\简历\26-28\pdf_office2007_converter
python -m venv .venv
.\.venv\Scripts\activate
python -m pip install -r requirements.txt
```

## 运行

```powershell
python app.py
```

浏览器打开：

```text
http://127.0.0.1:7633
```

端口在 `.env` 里配置：

```text
HOST=127.0.0.1
PORT=7633
```

也可以直接双击或运行：

```powershell
.\start_server.bat
```

默认清晰度是 `220 DPI`，对应这次可用的 `*_office2007.docx` 转换方式。

转换方式：

- `可编辑文字版`：默认模式，抽取 PDF 文字并写入 Word，尽量保留字号、粗体、颜色和段落位置。
- `图片兼容版`：把每页 PDF 转成图片放进 Word，Office 2007 打开最稳，但文字不能编辑。
