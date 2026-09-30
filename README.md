# Construction DOCX QA

本仓库用于施工组织设计 DOCX 的版式与页数校验。当前长丰义井镇项目同时保留三条通道：

1. **阶段 Microsoft Word 页数实测**：用于正文编制约 40% / 75% 时校准后续篇幅，不执行最终 50±1 页硬审。
2. **最终 Microsoft Word 正式交付硬审**：运行项目页数硬审 + 通用正式交付硬审，并导出 Word 实际版式 PDF。
3. **LibreOffice PDF 页数预检**：仅作云端预检，不代替 Microsoft Word 的最终结果。

## 长丰义井镇：阶段 Word 实测

阶段页数的目标值、容差和封面/目录排除规则**唯一读取**：

`audit/changfeng/page_gate.py`

因此阶段实测不会另维护一份“50页”参数。

阶段稿固定使用：

- 约 40%：`input/changfeng/stage-40.docx`
- 约 75%：`input/changfeng/stage-75.docx`
- 最终稿：`input/changfeng/design.docx`

将阶段稿更新到对应路径后，会自动触发：

`Changfeng Word stage page measurement`

Windows 自托管 runner 会调用本机 Microsoft Word 重新分页，产出阶段 JSON/TXT 报告。报告中的 `page_source` 必须为 `microsoft_word`。

阶段报告仅用于校准后续章节容量；粗略线性投影不是硬性页数判定，不能机械凑页。

## 长丰义井镇：最终正式交付硬审

最终稿仍放在：

`input/changfeng/design.docx`

现有工作流 `Changfeng Word formal delivery audit` 保持不变，继续运行项目页数硬审和通用正式交付硬审。只有最终稿才执行正式页数 PASS/FAIL。

## Windows 自托管 runner

阶段实测和最终硬审共用现有 Windows runner 标签：

`self-hosted, Windows, X64, word-audit`

runner 窗口显示 `Listening for Jobs` 时才在线。

本地原先用于配置 runner 的 ZIP 安装包，在已经成功解压、配置并能稳定接单后可以删除 ZIP 本身；**不要删除当前正在使用的已解压 runner 文件夹，也不要删除日常启动所用的 `start-word-runner.cmd`。**

## LibreOffice PDF 预检

仓库仍保留旧的 `DOCX PDF page preflight`。它只作辅助检查，不代替 Microsoft Word 阶段校准或最终硬审。
