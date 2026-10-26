# 格式检测接入

“格式检测”页面把模板论文和待检测论文提交到独立的后台串行队列，复用外部 `thesis-format-checker` 项目。检测不使用 MinerU 或模型 API。页面布局与样式保持原样，原有状态文字和按钮展示任务进度、受限项数量及下载结果。

## 环境

需要 Windows、已配置好的桌面版 Microsoft Word，以及安装了检测项目依赖的 Python 3.11+。默认发现文衡项目旁的 `thesis-format-checker/thesis-format-checker` 目录，使用该目录的 `.venv-run/Scripts/python.exe`，不会调用旧的 `.venv` 或启动脚本。Linux Docker 镜像不能执行当前完整检测器。

可在项目根目录 `.env` 中覆盖以下配置，修改后重启 `npm start`：

```dotenv
FORMAT_CHECKER_ROOT=F:/NJUSTwork/thesis-format-checker/thesis-format-checker
FORMAT_CHECKER_PYTHON=F:/NJUSTwork/thesis-format-checker/thesis-format-checker/.venv-run/Scripts/python.exe
FORMAT_CHECKER_TIMEOUT_SECONDS=1800
```

启动时自动执行 `003_format_tasks.sql`。检测任务与原 PDF 评审记录分别存储，共用已有登录账号和归属检查。

## 使用和恢复

每个区域选择一份不超过 30 MB 的 `.docx`，点击“开始检测”。后台验证 Word 文档结构，调用 Python 检测器，保存原生批注文档和 JSON 报告。页面每三秒查询状态，刷新页面会恢复当前账号最近的一次检测。任务进行时锁定文档选择；完成后可下载、替换任意一份文档或直接重新检测。刷新后仍可复用服务端已保存的两份输入。

提交响应中断时，页面查询后台已接受的任务，避免直接重复提交。首次连接失败后自动重试，恢复连接时清除旧提示。重新选择两份文件时创建独立任务；只替换其中一份时复用另一份已上传的文档。

队列通过数据库命名锁串行执行；同一账号同一时间只能有一个排队或运行中的任务。服务重启后继续排队任务，将失去执行进程的任务标记为中断，允许重新检测。默认超时 30 分钟。Python 包装器通过 Windows Job Object 管理它创建的 Word 实例，超时或终止 Python 时一起释放这些实例，不按进程名称结束用户的其他 Word 窗口。

Python 正常退出不代表所有格式合格。平台读取 JSON 中的问题数、无法判定、仅报告及跳过项，并在原结果区域提示。检测器的部分 Word 取证失败时可能仍生成文件，这类结果会显示“部分项目需核查”。

Word 启动或进程隔离失败会使任务失败，即使外部检测器捕获异常并生成部分结果，也不会被标记为完成。内部异常堆栈保存在任务的 `worker.log` 中。

## 文件和接口

文件保存在 `REVIEW_DATA_DIR/format/<任务ID>/`（默认 `data/format/`），含 `template.docx`、`document.docx`、`annotated.docx`、`report.json`、`worker.log`。缓存按账号存放在 `format/cache/`，避免不同账号共享解析缓存。备份脚本包含任务输入、批注文件和 JSON 报告。

所有接口需要登录：

| 接口 | 用途 |
| --- | --- |
| `GET /api/format/config` | 环境配置状态及文件大小限制 |
| `POST /api/format/tasks` | multipart 提交 `template`、`document`；可用 `source_task_id` 复用本人已有输入 |
| `GET /api/format/tasks/latest` | 恢复本人最近一次任务 |
| `GET /api/format/tasks/:id` | 查询本人任务状态和摘要 |
| `GET /api/format/tasks/:id/download` | 下载本人已完成任务的批注文档 |

初次配置建议用小型、非保密的真实 DOCX 验证 Word 可正常启动及批注输出；接口配置状态只检查运行系统与程序路径，完整可用性以实际任务结果为准。

## 验证

格式检测的服务与页面测试：

```powershell
node --test tests/format-checker.test.mjs tests/frontend.test.mjs
```

测试使用独立的 `wenheng_test_*` 数据库，`MYSQL_*` 需指向允许创建、删除测试库的测试实例，不要为测试扩大业务账号权限。

Windows 上可显式运行真实 Word 与 Chromium 验证，默认测试不会启动 Word：

```powershell
$env:RUN_WORD_INTEGRATION = '1'
node --test tests/format-http.test.mjs
Remove-Item Env:RUN_WORD_INTEGRATION
```

该测试自动生成虚构模板与论文，验证浏览器上传、真实检测、刷新恢复、下载、账号隔离、损坏文档拒绝，以及输入文件和输出正文内容保持不变。截图、批注文档和验证结果保存在 `test-results/format-integration/`。

2026-10-19 本机验证通过：样例生成 49 条 Word 原生批注（41 个错误、8 个警告）；正文文本、样式与媒体内容保持不变，输入文件校验值未变化，浏览器无脚本错误。3 项无法判定、2 项仅报告、37 项跳过如实显示为“部分项目需核查”。本轮功能接入未修改页面 HTML 与 CSS，二者 SHA-256 与接入前一致。
