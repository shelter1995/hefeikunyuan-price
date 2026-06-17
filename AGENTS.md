# AGENTS.md — 项目报价更新系统

> 本文件面向 AI 编程助手。如果你正在阅读此文件，说明你被期望修改、调试或扩展本项目。请仔细阅读后再动手。

---

## 1. 项目概述

本项目是一个**钢材价格自动化更新系统**，服务于建筑钢材报价业务。核心功能包括：

- **网价更新**：通过 Playwright 自动登录 mysteel.com 抓取最新钢材网价，写入项目 Excel 的 `G1/G3/G4` 单元格（日期/盘螺/螺纹）。
- **图片/文档价更新**：使用 MiniMax 多模态视觉模型识别线下报价图片（`.jpg`/`.png`）中的价格和库存信息，配合文本解析（`.txt`），写入项目 Excel 的 `H1/H3/H4` 单元格。
- **库存颜色标注**：根据库存状态（充足/告警/缺货）在"报价表" Sheet 中为对应单元格标注蓝色/黄色/红色。
- **安全回写机制**：强制先 `--dry-run` 预览，确认无异常后再 `--confirm-write` 写入；价格偏离网价超过 ±1000元/吨 或 ±20% 时自动阻断。

项目使用 `pyproject.toml` 配置 pytest 与 ruff（无 `setup.py`），以普通 Python 包形式直接运行。

---

## 2. 技术栈

| 层级 | 技术 |
|------|------|
| 语言 | Python 3.12+（`pyproject.toml` 中 ruff `target-version = py312`） |
| Excel 处理 | `openpyxl` |
| 浏览器自动化 | `playwright`（Chromium） |
| 视觉识别 | MiniMax VLM API（HTTP/JSON） |
| 传统 OCR（备用/诊断） | PaddleOCR layout-parsing API |
| HTTP 请求 | `requests` |
| 测试框架 | `pytest` |

---

## 3. 项目结构

```
d:\GitHub_WorkSpace\hefeikunyuan
├── ocr_price/                  # 核心业务代码包（18 个业务模块 + __init__.py）
│   ├── pipeline.py             # 主流程编排（single / batch CLI）
│   ├── web_price.py            # 网价抓取、登录、回写（G列）
│   ├── writeback_image_doc.py  # 图片/文档价回写（H列）
│   ├── minimax_vision.py       # MiniMax VLM 视觉识别客户端
│   ├── semantic_adjustment.py  # 补充报价 LLM 语义理解（MiniMax 文本模型）
│   ├── parser.py               # 文本/表格/价格解析引擎
│   ├── rules.py                # 业务规则（规格提取等）
│   ├── inventory_color.py      # 库存颜色标注（现代实现）
│   ├── inventory.py            # 库存解析（传统实现）
│   ├── offline_validation.py   # 线下价格安全校验
│   ├── reporting.py            # Markdown/JSON 报告生成
│   ├── audit.py                # 回写后审计
│   ├── manifest.py             # dry-run manifest 生成与校验
│   ├── events.py               # 结构化事件记录
│   ├── env_loader.py           # .env 文件加载器
│   ├── paddle_api.py           # PaddleOCR 备用客户端
│   ├── xlsx_utils.py           # openpyxl 安全加载器（XML 修复）
│   ├── cli.py                  # 独立单文件 OCR 提取 CLI
│   └── __init__.py             # 包入口，暴露 parser 函数
├── tests/                      # pytest 测试套件
│   ├── conftest.py             # 注入项目根目录到 sys.path
│   └── test_*.py               # 19 个测试模块
├── skills/quote-update/        # Agent Skill（执行脚本、规则、流程文档）
│   ├── scripts/
│   │   ├── run_single.py       # 单文件运行入口（推荐）
│   │   ├── run_batch.py        # 批量运行入口（推荐）
│   │   ├── apply_confirmations.py    # 用户确认后更新待确认对照表
│   │   ├── make_agent_prompt.py      # 生成给其他 Agent 的执行指令
│   │   └── check_pending_mapping.py  # 待确认映射检查
│   ├── references/
│   │   ├── commands.md         # 命令模板
│   │   ├── rules.md            # 业务规则
│   │   └── workflow.md         # 完整工作流程
│   └── SKILL.md                # Skill 定义（触发条件、执行流程、规则）
├── archive/one_off_scripts/    # 已归档的一次性调试脚本（不参与正式流程）
├── doc/                        # 中文架构文档（库存颜色、标准流程、流程图）
├── 项目报价/                   # 待更新的项目 Excel 文件（gitignore，运行时创建）
│   └── 备份/                   # 自动备份目录
├── 线下报价/                   # 线下报价源文件（.jpg/.png/.txt，gitignore）
├── 运行产物/                   # 报告、对照表、OCR JSON 中间产物、备份（gitignore）
├── pyproject.toml              # pytest 与 ruff 配置
├── requirements-ocr.txt        # Python 依赖清单
├── .env / .env.example         # 环境变量（API Key 等）
└── README.md                   # 给其他 Agent 的迁移指南（极详细）
```

> 说明：`项目报价/`、`线下报价/`、`运行产物/`、`.env`、`网站账号密码.txt` 等均已列入 `.gitignore`，属于运行时数据，仓库中不一定存在。

### 3.1 关键文件说明

| 文件 | 说明 |
|------|------|
| `ocr_price/pipeline.py` | **唯一主入口**。提供 `single` 和 `batch` 子命令，协调 web 流程与 image_doc 流程，支持 `--dry-run` / `--confirm-write` / `--manifest` / `--refresh-*-artifacts`。 |
| `skills/quote-update/scripts/run_single.py` | 推荐使用的单文件包装脚本，对 `python -m ocr_price.pipeline single` 做了一层 CLI 参数透传。 |
| `skills/quote-update/scripts/run_batch.py` | 推荐使用的批量包装脚本，对 `python -m ocr_price.pipeline batch` 做了一层 CLI 参数透传。 |
| `ocr_price/web_price.py` | 网价全链路：Playwright 登录 → 价格抓取 → 生成清单 Excel → 厂家对照映射 → 回写 G1/G3/G4。 |
| `ocr_price/writeback_image_doc.py` | 图片/文档价全链路：加载 OCR JSON → 厂家对照映射 → 价格偏差校验 → 回写 H1/H3/H4 → 库存颜色标注。 |
| `ocr_price/minimax_vision.py` | MiniMax 视觉 API 客户端。支持 3 次重试 + 结果合并，将图片（`.jpg`/`.png`）转为结构化价格+库存 JSON。 |
| `ocr_price/offline_validation.py` | 线下价动态校验：对比 H3/H4 与 G3/G4，偏差超过 ±1000 或 ±20% 则阻断写入。 |
| `ocr_price/xlsx_utils.py` | `load_workbook_safe()` — 自动修复 openpyxl 样式 XML 损坏。 |

---

## 4. 构建与运行命令

### 4.0 小白用户一句话入口

本项目面向的小白用户不需要提供命令行参数，也不需要复制长模板。只要用户在 agent 对话框里说一句自然语言，Agent 必须自己完成文件匹配、模式判断、dry-run 执行和结果汇报。

常见一句话示例：

| 用户说法 | Agent 应理解为 |
|----------|----------------|
| `更新报价` | 如果 `项目报价/` 下只有一个 `.xlsx`，对该文件执行 `mode=both` 的 dry-run；如果有多个文件，列出文件让用户选择 |
| `更新景明苑报价` | 在 `项目报价/` 中按关键词 `景明苑` 查找实际文件名，唯一匹配时执行单文件 `mode=both` dry-run |
| `只更新景明苑网价` | 查找 `景明苑` 项目文件，执行单文件 `mode=web` dry-run |
| `只更新景明苑线下报价` | 查找 `景明苑` 项目文件，执行单文件 `mode=image_doc` dry-run |
| `批量更新报价` | 对 `项目报价/*.xlsx` 执行批量 `mode=both` dry-run |
| `确认写入景明苑报价` | 仅在已经有同一次 dry-run 的 Manifest 且报告无待确认/异常时，使用该 Manifest 执行 confirm-write |

一句话入口的强制流程：

1. 先读取 `项目报价/` 实际文件名，禁止让用户自己拼路径。
2. 根据用户话术识别模式：默认 `both`；出现“只更新网价”用 `web`；出现“只更新线下/图片/文档”用 `image_doc`；出现“批量”用 batch。
3. 如果匹配到 0 个项目文件，说明未找到并列出可选文件；如果匹配到多个项目文件，只展示候选项并让用户选择，禁止猜测。
4. 默认先 dry-run，报告 Manifest、Report、MarkdownReport、Events、待确认项和异常项。
5. dry-run 通过后，必须等待用户明确说“确认写入/可以写入/执行写入”，才能用同一次 Manifest 执行 confirm-write。
6. 禁止要求用户自己拼命令，除非用户明确要求查看命令。
7. 用户确认厂家映射后，必须用 `skills/quote-update/scripts/apply_confirmations.py` 更新网价和图片/文档待确认表；pending 为 0 后直接复用同一次 Manifest 执行 confirm-write。确认映射后禁止重新 dry-run，避免重复登录和重复抓网价。
8. 图片/文档厂家也必须完整列入确认清单；例如来源文件为“徐刚”、项目 sheet 为“徐钢”时，应作为同一厂家进入待确认匹配，确认后才能写入 H1/H3/H4 和库存颜色。

换句话说，小白用户只需要说：`更新景明苑报价`。Agent 负责执行安全流程：先 dry-run，报告 Manifest，等待用户确认，再 confirm-write。

### 4.1 环境准备

在项目根目录执行：

```powershell
pip install -r requirements-ocr.txt
playwright install chromium
```

然后复制环境变量模板并填写：

```powershell
copy .env.example .env
# 编辑 .env，填入 MINIMAX_API_KEY 和 PaddleOCR 配置（可选）
```

确认 `网站账号密码.txt` 存在且包含 mysteel.com 的登录凭据（格式为 `用户名,密码`）。

### 4.2 推荐运行方式（通过包装脚本）

**单文件 dry-run（必须先执行）：**
```powershell
python skills/quote-update/scripts/run_single.py `
  --project "项目报价/<项目Excel路径>" `
  --mode both `
  --dry-run `
  --headless
```

**单文件确认写入：**
```powershell
python skills/quote-update/scripts/run_single.py `
  --project "项目报价/<项目Excel路径>" `
  --mode both `
  --confirm-write `
  --headless `
  --manifest "运行产物/<项目名>/dry_run_manifest.json"
```

**批量 dry-run：**
```powershell
python skills/quote-update/scripts/run_batch.py `
  --project-dir "项目报价" `
  --glob "*.xlsx" `
  --mode both `
  --dry-run `
  --headless
```

### 4.3 直接模块运行方式（等效）

```powershell
# 单文件
python -m ocr_price.pipeline single --project "..." --mode both --dry-run

# 批量
python -m ocr_price.pipeline batch --project-dir "项目报价" --glob "*.xlsx" --mode both --dry-run

# 独立 OCR 提取（不写入 Excel，仅生成 JSON）
python -m ocr_price.cli --input "线下报价/徐钢报价.jpg" --location "蚌埠"
```

### 4.4 常用 CLI 参数

| 参数 | 说明 |
|------|------|
| `--mode {web,image_doc,both}` | 更新模式 |
| `--dry-run` | 预演模式，不修改任何 Excel 文件 |
| `--confirm-write` | 确认写入模式；应配合同一次 dry-run 生成的 `--manifest` 使用 |
| `--manifest PATH` | dry-run 生成的 manifest 路径；单文件 confirm-write 默认复用它 |
| `--refresh-web-artifacts` | 强制重新抓取网价 |
| `--refresh-image-artifacts` | 强制重新跑 OCR |
| `--headless` | 无头浏览器模式（登录态持久化在 `.chrome_user_data/`） |
| `--manual-login-timeout N` | 手动登录等待超时（默认 180 秒） |

---

## 5. 测试说明

### 5.1 运行测试

```powershell
pytest tests/
```

测试发现规则由 `pyproject.toml` 的 `[tool.pytest.ini_options]` 配置。

### 5.2 测试目录结构

| 测试文件 | 测试内容 |
|----------|----------|
| `test_pipeline_safety.py` | dry-run 安全：确认 workbook 不会被修改 |
| `test_pipeline_artifact_apply.py` | confirm-write 复用产物逻辑、待确认阻断 |
| `test_offline_price_deviation.py` | 线下价偏离网价的阻断阈值（±1000 / ±20%） |
| `test_offline_validation.py` | 线下价 payload 校验（价格范围、地点、库存枚举） |
| `test_inventory_writeback.py` | 库存颜色回写（蓝/黄/红） |
| `test_web_price_interactive_login.py` | 手动登录 fallback 逻辑（Mock Playwright） |
| `test_location_parse.py` | 从文件名解析地点（如 `安徽合肥-安徽蚌埠`） |
| `test_minimax_conversion.py` | MiniMax 响应格式转换 |
| `test_reporting.py` | Markdown 报告格式 |
| `test_audit.py` | 回写后审计（值匹配校验） |

### 5.3 测试约定

- `conftest.py` 仅做 `sys.path` 注入，使 `from ocr_price import ...` 能正确解析。
- 大量使用 `tmp_path` 创建隔离的 Excel/JSON 夹具。
- 大量使用 `monkeypatch` Mock 外部依赖（MiniMax API、Playwright、文件系统）。
- 测试风格混合**纯单元测试**（数学计算、字符串解析）和**集成测试**（openpyxl 完整读写回环）。

### 5.4 归档的一次性调试脚本（非 pytest）

历史根级 `test_*.py`、`reprocess_images*.py`、`retry_failed.py` 已归档到 `archive/one_off_scripts/`。这些是**一次性手工调试脚本**，不是正式测试用例。修改核心代码后无需保证它们通过；如需参考，先确认当前 `ocr_price/` API 是否已变化。

---

## 6. 代码风格与开发约定

### 6.1 语言与注释

- 项目文档、README、doc/ 目录、代码内关键注释均使用**中文**。
- 模块级 docstring 和函数签名使用英文（如 `ocr_price/__init__.py`）。
- 变量命名混合拼音缩写（如 `xlsx`）与英文业务术语（如 `writeback`、`pipeline`）。

### 6.2 模块组织原则

- `ocr_price/` 是唯一的业务代码包，无子包扁平化组织。
- `pipeline.py` 是**唯一协调者**，不直接操作 Excel，只调用 `web_price.py` 和 `writeback_image_doc.py`。
- `web_price.py` 和 `writeback_image_doc.py` 各自独立，可单独作为 CLI 子模块运行（`python -m ocr_price.web_price ...`）。
- `parser.py`、`minimax_vision.py`、`paddle_api.py` 是纯数据提取层，无 Excel 依赖。

### 6.3 安全与状态管理约定

1. **强制两步写入**：任何写入操作必须先 `--dry-run`，再 `--confirm-write`。`pipeline.py` 在 CLI 层面强制二者必须选其一；单文件 `confirm-write` 应复用同一次 dry-run 输出的 `--manifest`。
2. **自动备份**：`--confirm-write` 时会在 `运行产物/<项目>/` 下生成带时间戳的 `.xlsx` 备份。
3. **待确认阻断**：新厂家出现时生成 `"待确认"` 状态记录，必须经用户在对话中确认后才能改为 `"已确认匹配"`。禁止代码自动确认。
4. **价格偏差阻断**：`writeback_image_doc.py` 在回写 H3/H4 前读取同 sheet 的 G3/G4 作为参考，偏差过大则跳过并报告。
5. **页签颜色隔离**：网价链路允许"先清空后标红未匹配"；图片/文档链路**禁止**修改页签（sheet tab）颜色。
6. **Manifest 约束**：如果 dry-run 没有输出 `Manifest:` 路径，不得执行 `confirm-write`。

### 6.4 Excel 单元格约定

| 列 | 含义 | 更新来源 |
|----|------|----------|
| G1 | 网价日期 | web_price.py |
| G3 | 网价盘螺(Φ8-10) | web_price.py |
| G4 | 网价螺纹(Φ18) | web_price.py |
| H1 | 图片/文档报价日期 | writeback_image_doc.py |
| H3 | 图片/文档盘螺价 | writeback_image_doc.py |
| H4 | 图片/文档螺纹价 | writeback_image_doc.py |

### 6.5 库存颜色码

| 状态 | 颜色 | RGB（openpyxl 填充） |
|------|------|----------------------|
| 充足 | 蓝色 | `FF0070C0` |
| 告警 | 黄色 | `FFFFC000` |
| 缺货 | 红色 | `FFFF0000` |

---

## 7. 安全注意事项

- **敏感文件**：`.env` 包含 `MINIMAX_API_KEY`；`网站账号密码.txt` 包含站点凭据。二者均已列入 `.gitignore`，**严禁提交到 Git**。
- **浏览器持久化数据**：`.chrome_user_data/` 保存登录态 Cookie，也已列入 `.gitignore`。
- **API Key 泄露风险**：`minimax_vision.py` 和 `paddle_api.py` 通过 `os.environ` 读取 Key，不会在日志中打印。修改时请勿添加 `print(api_key)` 之类的调试语句。
- **Excel 写入安全**：`xlsx_utils.py` 的 `load_workbook_safe()` 会原地修复样式 XML。如果进一步修改修复逻辑，务必先在副本上验证，避免破坏原始文件。
- **价格偏差校验是硬规则**：`offline_validation.py` 中的 ±1000 和 ±20% 阈值是业务安全底线，**不可随意放宽**。

---

## 8. 给 Agent 的快速参考

### 8.1 修改代码前必读

1. 阅读 `skills/quote-update/SKILL.md` 了解触发条件和执行流程。
2. 阅读 `skills/quote-update/references/rules.md` 了解业务规则。
3. 阅读 `skills/quote-update/references/workflow.md` 了解完整工作流。
4. 阅读本项目的 `README.md`（它是给其它 Agent 的迁移指南，包含结果返回格式规范）。

### 8.2 常见修改场景

| 场景 | 应修改的文件 |
|------|-------------|
| 调整网价抓取逻辑 | `ocr_price/web_price.py` |
| 调整 MiniMax 视觉识别 prompt/解析 | `ocr_price/minimax_vision.py`、`ocr_price/parser.py` |
| 调整库存颜色标注规则 | `ocr_price/inventory_color.py`、`ocr_price/inventory.py` |
| 调整价格偏差安全阈值 | `ocr_price/offline_validation.py` |
| 调整报告格式 | `ocr_price/reporting.py` |
| 调整主流程编排 | `ocr_price/pipeline.py` |
| 调整 Excel 加载/保存行为 | `ocr_price/xlsx_utils.py` |
| 新增测试 | `tests/test_*.py`（遵循现有命名和夹具风格） |

### 8.3 验证修改的正确方式

```powershell
# 1. 运行相关测试
pytest tests/test_<对应模块>.py -v

# 2. 对单个项目做 dry-run 验证
python skills/quote-update/scripts/run_single.py `
  --project "项目报价/<某个项目>.xlsx" `
  --mode both `
  --dry-run `
  --headless

# 3. 检查 运行产物/<项目>/ 下生成的 JSON/Excel 是否符合预期
```

---

## 9. 外部依赖与配置

### 9.1 依赖

运行依赖（`requirements-ocr.txt` / `pyproject.toml` 的 `[project.dependencies]`）：

```text
requests>=2.31.0
openpyxl>=3.1.5
playwright>=1.54.0
```

开发/测试依赖（`pyproject.toml` 的 `[project.optional-dependencies].dev`）：

```text
pytest>=8.3.0
ruff>=0.6.0
```

安装运行依赖用 `pip install -r requirements-ocr.txt`；安装开发依赖用 `pip install -e .[dev]`。

### 9.2 .env 变量

| 变量 | 用途 |
|------|------|
| `MINIMAX_API_KEY` | MiniMax VLM API 认证（主链路必需） |
| `PADDLEOCR_BASE_URL` | PaddleOCR layout-parsing 服务端点（备用） |
| `PADDLEOCR_API_KEY` | PaddleOCR API Key（备用） |
| `PADDLEOCR_AUTH_SCHEME` | PaddleOCR 认证方式（如 `token`） |

---

## 10. 注意事项（必读）

- **禁止直接修改 Excel 单元格**：任何 Agent 都不得绕过 `pipeline.py` 或 `run_single.py` / `run_batch.py` 直接操作 `.xlsx` 文件。
- **禁止伪造数据**：所有价格必须从实际文件或 API 响应中提取，禁止编造。
- **禁止自动确认映射**：新厂家的 `"待确认"` 状态必须由人类用户在对话中确认，Agent 不得私自改为 `"已确认匹配"`。
- **结果返回格式**：执行完成后必须按 `README.md` 第 5.0 节的标准格式返回结果（包含时间、模式、已更新/未更新明细、汇总统计表、报告文件路径、库存颜色统计）。
