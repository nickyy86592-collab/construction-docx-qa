# DOCX PDF 页数预检（GitHub Actions）

此工具在 GitHub Actions 的 Ubuntu 运行器安装 LibreOffice、Poppler 和中文字体，将 DOCX 转 PDF，统计正文页数，并输出 PDF、`report.json`、`report.txt`。**不是 Microsoft Word 实测，不代替原 Word 硬审脚本的 PASS。**

## 使用

1. 将本包的 `pdf_page_gate.py`、`smoke_test.py`、`.github/workflows/docx-pdf-page-check.yml` 放到**私有** GitHub 仓库相同路径。代码推送后自动用一次性示例运行自测。
2. 将要检测的 DOCX 放在该私有仓库，例如 `input/draft.docx`。投标文件不应放入公开仓库。
3. 在仓库的 Actions → `DOCX PDF page preflight` → Run workflow，输入路径、正文目标页数、容差、前置页数。长丰义井镇的已知要求：正文目标 50、容差 1、前置 1。
4. 运行完成后下载 `docx-pdf-page-preflight` 工件，查看 PDF 和报告。若本地 Microsoft Word 的分页与 PDF 不同，以 Word 实际分页为准。

本工具只检查 PDF 页数并给出标题/强制分页/横版线索；不会假称完成原版式硬审，也不会把 37 页稿自动补成 50 页。每次修改文稿后须重新运行。若文档使用特殊字体，仍需核对所用字体和字体替换情况。
