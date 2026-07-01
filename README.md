# DOCX 划词批注

前端：`frontend`，Vue 3 + Vite + mammoth.js + web-highlighter。

后端：

- `backend-python`：Flask + python-docx，默认端口 `5000`。
- `backend-python-aspose`：Flask + aspose-words，默认端口 `5001`。
- `backend-java-aspose`：Spring Boot + Aspose.Words for Java，默认端口 `8080`。

## 运行

```bash
cd backend-python
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Aspose 后端：

```bash
cd backend-python-aspose
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Java Aspose 后端：

```bash
cd backend-java-aspose
mvn spring-boot:run
```

```bash
cd frontend
npm install
npm run dev
```

默认地址：

- 前端：http://localhost:5173
- python-docx 后端：http://localhost:5000
- Aspose 后端：http://localhost:5001
- Java Aspose 后端：http://localhost:8080

前端已配置两条 Vite 代理：

- `/api/python-docx/*` -> `http://localhost:5000/api/*`
- `/api/aspose/*` -> `http://localhost:5001/api/*`
- `/api/java-aspose/*` -> `http://localhost:8080/api/*`

## 功能

- 前端提供三个上传入口和三个预览 Tab：mammoth.js 在浏览器端解析预览，Python Aspose/Java Aspose 后端分别解析成 HTML 后返回预览；三个 Tab 的预览区域、批注列表和高亮状态彼此独立。
- 在预览区域划词后，web-highlighter 自动高亮选区。
- 保存批注后，可分别点击“生成批注文件（python-docx）”或“生成批注文件（Aspose）”。Aspose 按钮会复用后端上传时保存的原始 DOCX。
- `backend-python` 用 python-docx 的 `Document.add_comment()` 在匹配文本上生成 Word 批注。
- `backend-python-aspose` 用 Aspose.Words 的 `CommentRangeStart`、`CommentRangeEnd`、`Comment` 生成 Word 批注。
- `backend-java-aspose` 用 Aspose.Words for Java 的 `CommentRangeStart`、`CommentRangeEnd`、`Comment` 生成 Word 批注。

## 说明

当前版本按 `选中文本 + 第几次出现` 定位原文。建议批注单段落内的文本；跨段落选择、复杂表格/页眉页脚内容可能需要进一步增强定位逻辑。
