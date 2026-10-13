# 通用施工组织设计 Word 审核

通用编制、格式和图文核验不绑定工程专业、名称、章节、字体或页数。项目事实与格式要求由AI从当次招标及技术资料提取；正文页数目标/范围、容差和封面目录扣除页数只来自当前唯一项目页数硬审脚本。

- 通用代码：`audit/common/`。
- 当前项目输入：`input/project/`，仅保留一个`*页数硬审*.py`，名字可随项目变化。
- 运行期正式稿：`input/project/design.docx`及AI生成的`format_rules.json`。
- 代表性样章：`input/project/pilot.docx`，Word测量会导出实际PDF和非表格文字量。
- 阶段快照：`input/project/stage-40.docx`、`stage-75.docx`。
- 历史项目资料：`archive/previous-projects/`，不参加当前项目审核。

## 换项目

更换项目页数脚本（删除同目录旧页数文件，避免两份并存），正常提供新项目招标/图纸/知识库。AI重新生成项目状态、格式规则、照片清单和正文，将当前稿同步至统一输入路径。用户不需改通用技能、通用脚本或工作流。

项目页数脚本保留三个字面量整数常量`EXPECTED_BODY_PAGES`、`PAGE_TOLERANCE`、`FRONT_MATTER_PAGES`；最终审核CLI支持`document --report 路径`及0/2/3退出码。没有强制固定的项目文件名。阶段实测直接读取同一脚本常量并调用Word，不要求项目脚本提供阶段模式。

## 执行入口

`Word representative sample measurement`测量编制前代表性样章并导出Word PDF，供AI逐页复核和预算校准。`Word stage page measurement`实测40%/75%快照，仅产MEASURED，不把阶段稿按成稿范围判FAIL。`Word formal delivery audit`执行静态图文、项目页数、Word版式检查并导出Word PDF。这些工作流复用`self-hosted, Windows, X64, word-audit`标签的运行器和缓存依赖。只有实际Word返回的分页报告才是Word证据。

三张技术图必须彩色，真实照片或编制阶段带框占位必须存在。图件、照片主题、章节及字体来自运行期规则，通用仓库不保存项目默认章节/照片主题/字号。

当前项目尚无DOCX及格式规则时，工作流明确记录INPUT_READY=false并跳过Word任务，不声称通过Word审查。完整交付仍由Skill核对启动、阶段状态、评分、内容溯源和WARN；Word组件报告不替代完整delivery门禁。

`DOCX PDF page preflight`只作LibreOffice估值，同样读取唯一项目页数脚本，不在工作流填写另一份目标。`Generic audit synchronization checks`用临时合成文件检查通用脚本及换项目接口，不运行Windows Word。

## 内容规划与长文门禁

Skill先将评分子条件和施工范围落到章—节—实质内容模块，再按文字/表格/图片与机动量预算，预编制总量严格在当前页数脚本范围内。目标达到200页的长文必须先经Word代表性样章试排及密度校准，draft-ready通过才批量编制。模块保存与恢复不限次数，用户交付最多三批。

模块正文源、实质复核与阶段快照绑定哈希。新模块顺序登记，到40%/75%节点先校准才能继续后续模块；阶段实测前缀加剩余预算须落在最终允许范围。最终合并稿检查正文探针与模块主体覆盖，避免漏合并正文。规则及专业判断由AI根据项目资料维护，结构门禁不等于专业质量证明。合成测试不代表Windows Word实测。
