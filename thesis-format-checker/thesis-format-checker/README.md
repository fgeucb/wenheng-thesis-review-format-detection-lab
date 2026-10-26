# 论文格式对照检测系统

在 Windows 上调用桌面版 Microsoft Word，以模板 A 为基准，按**同一层级 + 同一位置**
逐项对照检查待检测文档 B，并把发现的问题写成 **Word 原生批注**。输入支持
`.doc` / `.docx`，输出统一为 `.docx`，**输入原件永不被修改**。

设计目标：**一个大类一个类、一条规则一个方法**。要检测的 165 条规则全部登记在
规则目录里，改哪条规则就去对应类里改那一个方法。

## 目录结构

```text
thesis-format-checker/
├─ pyproject.toml              项目与依赖配置（可 pip install -e .）
├─ thesis-checker.cmd          免安装启动器
├─ docs/
│  ├─ 架构说明.md              分层设计、扩展方式、实现状态
│  ├─ 模板编辑规范.md          给高校的模板写法要求 + 交付前自检清单
│  └─ 规则清单.md              165 个规则族 / 356 条细粒度规则明细（自动生成）
├─ tools/
│  └─ check_template.py        模板体检：按规范逐项检查一份模板
├─ examples/
│  ├─ mapping.example.yaml     可选映射配置示例
│  └─ create_demo_documents.py 生成演示用模板/论文（需要 Word）
├─ src/thesis_format_checker/
│  ├─ cli.py                   命令行：compare / rules
│  ├─ core/                    基础设施（不判断规则）
│  │  ├─ models.py             数据模型：Finding、RunSummary、规格对象
│  │  ├─ docx_context.py       只读的 docx 段落/格式上下文
│  │  ├─ effective_format.py   有效文字格式与中英文脚本切分
│  │  ├─ word_session.py       独立 Word 进程、只读打开、异常清理
│  │  ├─ word_layout.py        排版快照：分页、每页页眉页脚变体、图形/表格坐标
│  │  ├─ point_comments.py     字符级批注（按页锚定）与无损拆分 run
│  │  ├─ review_settings.py    可选：输出副本打开“显示批注”开关
│  │  ├─ cache.py              结果缓存：复用 Word 段落属性与排版快照
│  │  ├─ scope.py              正文范围识别
│  │  └─ comments.py           原生批注锚定与写入
│  ├─ checks/                  规则层：大类 → 类，规则 → 方法
│  │  ├─ catalog.py            165 条规则的唯一权威清单（标题/状态/来源/备注）
│  │  ├─ base.py               CheckResult + CategoryChecker 基类
│  │  ├─ c01_structure.py … c15_appendix.py   15 个大类检测器
│  │  └─ registry.py           大类 key → 检测器类
│  ├─ detectors/               取证实现层：按大类读取模板与文档对象
│  │  ├─ structure.py … appendices.py         与 15 个大类一一对应
│  │  ├─ fields.py / text_runs.py             共用读取工具
│  │  └─ word_effective/      旧版 Word 有效格式检测（页边距/字体/行距/标点等）
│  └─ engine/
│     ├─ pipeline.py           调度：模板基准 → 逐类取证 → 批注落盘
│     └─ reporting.py          报告：命令行摘要、规则清单、JSON 报告
└─ tests/                      单元测试与真实 Word 集成测试
```

## 安装

```powershell
Set-Location -LiteralPath 'D:\Lenovo\Documents\南理\thesis-format-checker'
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -e '.[test]'
```

需要 Windows 10/11、Python 3.11+、可正常启动的桌面版 Microsoft Word。

## 使用

```powershell
# 检测并写入原生批注
.\thesis-checker.cmd compare `
  --template 'D:\path\template_a.docx' `
  --document 'D:\path\document_b.doc' `
  --output   'D:\path\document_b_格式检测批注.docx' `
  --report-json 'D:\path\检测报告.json'

# 查看规则清单与实现状态（15 大类、165 条）
.\thesis-checker.cmd rules
.\thesis-checker.cmd rules --category c07     # 只看“字体与段落检测”
.\thesis-checker.cmd rules --pending          # 只看尚未实现的规则

# 模板体检：交付模板前先自查（详见 docs/模板编辑规范.md）
.\.venv\Scripts\python.exe tools\check_template.py 'D:\path\template_a.docx'
```

安装后也可以直接使用命令名 `thesis-checker`，或用
`python -m thesis_format_checker`。

参数与退出码：默认不覆盖已有输出（需要时加 `--overwrite`）；输出路径必须以
`.docx` 结尾且不能与输入相同。发现格式问题仍返回 `0`，参数错误返回 `2`，
运行/文档/配置错误返回 `3`。

### 速度与缓存

最慢的一段是 Word COM（有效格式 + 分页）。同模板/论文的**第二次检测**会复用
上次解析好的 Word 段落属性、模板剖面与排版快照，通常压到 1 分钟以内
（67 页样例论文实测：首次约 7 分钟，第二次约 1 分钟）：

```powershell
.\thesis-checker.cmd compare --template ... --document ... --output ...   # 第一次：完整检测，写缓存
.\thesis-checker.cmd compare --template ... --document ... --output ...   # 第二次：命中缓存，不再开 Word

--no-cache                 # 完全关闭缓存（改模板/改代码后想强制全量时用）
--cache-dir D:\cache\路径   # 指定缓存目录，默认 %LOCALAPPDATA%\thesis-checker\cache
```

缓存键包含源码指纹，因此**改了代码或模板/论文内容，缓存会自动失效**，
不需要手工清理，也不会读到旧结论。

## 规则、等级与批注格式

规则分成两层：

* **规则族**：沿用 001–165 的编号，用于归类与对旧文档，也是今天检测代码的输出单位；
* **细粒度规则**：**一条只讲一件事**，编号形如 `013.3`（`族编号.序号`），
  共 **356 条**。

每条细粒度规则都带等级与状态：

| 等级 | 含义 | 处理要求 |
| --- | --- | --- |
| 错误 | 违反学校模板的硬性要求 | 必须修改 |
| 警告 | 影响规范性或与模板不一致 | 建议修改 |
| 提醒 | 排版细节、风格偏好 | 可改可不改 |

当前 356 条中：307 条已独立输出结论、15 条已检测但暂与同族条目合并输出、34 条待确认、
0 条可确定性待实现；另有 21 个规则族因为**必须依赖大模型语义理解**已整体删除
（编号与原因保留在清单里）。逐条明细见 [docs/规则清单.md](docs/规则清单.md)，
设计说明见 [docs/架构说明.md](docs/架构说明.md)。

批注文案统一为「等级 + 问题标题 + 应该 + 实际」，批注本身直接锚在出问题的文字上
（不写规则编号、不写“位置：第 N 段”）：

```
【错误】封面题目格式不符合模板
应该：字号 19 磅、加粗
实际：字号 16 磅、未加粗
```

需要 Word 渲染才能判定的规则（如空白页、页码、题注位置）在跑完整流程时给出确定
结论；所有“无法判定”条目（包括 Word 路径和低层写入入口）**只进 JSON 报告，不写批注**。

Word 不允许在页眉页脚中插入批注。页眉页脚问题使用受影响页正文首末边界的
**零长度锚点**，不圈选封面题目、校徽或正文，不增加空格或段落。缺乏可靠边界时只进报告，
不回退到文首。显示公式段落也仅报告（实测 Word 原生批注会改变其排版）；显式分页符前禁止写入页级锚点。
空白页、复杂对象边界及同页跨节可能无法定位；短页的底部指该页现有正文末端，
不能在不改排版的前提下凭空创建纸张下边距锚点。
首页/奇偶页按 Word 实际变体及打印页码确定；共享配置仅写一条批注，
所有受影响页序及实际锚点保存在 JSON 的 `affected_pages` / `page_comments` 中。
可见的页眉页脚内容、距离差异在每个可确认的受影响页分别标注；同页同一问题跨规则合并。
字符类问题按确切字符范围无损拆分 run，只圈选出错词或符号；填写项缺失只标注对应填写区域。

`thesis-checker rules` 可以打印规则族与细粒度规则及其等级、状态、定位方式；
`thesis-checker rules --pending` 只列尚未实现的条目。

## 扩展方式

1. **新增/细化一条规则**：在 `checks/rule_specs.py` 的 `FINE_RULE_ROWS` 登记
   （编号、标题、等级、状态、定位方式），重新生成文档
   （`python tools/generate_rule_catalog.py`）。
2. **实现某条细规则**：在所属大类的检测器里产出对应结论；若算法在 `detectors/`
   里，就把族编号加进 `catalog.PROVIDERS` 对应条目的 `rules`，并把它在清单里的
   状态从“已实现（聚合输出）”改成“已实现”。
3. **新增大类**：在 `catalog.CATEGORIES` 登记，新建 `checks/cXX_*.py` 与
   `detectors/xxx.py`，再到 `checks/registry.py` 注册。

## 测试

```powershell
# 架构自检 + 全部单元测试（不需要 Word）
.\.venv\Scripts\python.exe -m pytest -m "not integration"

# 真实 Word 端到端测试（需要本机可用的 Word）
$env:RUN_WORD_INTEGRATION = '1'
.\.venv\Scripts\python.exe -m pytest -m integration -s
```

## 能力边界

- 只检测主文档正文，以及显式读取的页眉页脚、标题、目录、图表题注、公式、
  脚注尾注与附录；表格内部文字、文本框内容不参与正文格式比较。
- 依赖语义判断的规则（摘要是否完整、图表是否清晰、语言是否规范等）本阶段
  不实现，代码中保留了方法骨架与原因说明。
- 程序只添加批注，不自动修改论文内容或格式；正文允许为字符边界批注无损拆分同格式 run，
  去除批注、规范化 run 后内容树必须一致；样式、图片、页眉页脚等非批注部件逐字节一致。
