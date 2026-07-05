---
name: model-tester
description: "使用多个AI模型对测试题进行回答，并将结果保存到格式化的Excel文件中。支持API模式和浏览器网页自动化模式。触发条件：模型测试、用不同模型测试、AI模型对比测试、模型答题、批量测试模型、网页测试、浏览器测试、网页自动化、不用API测、API测试、调模型、模型对比、跑测试、测一下XX模型、用XX回答题目、安全测试、测试题、benchmark评测。适用于安全测试题、benchmark测试、模型能力对比、网页端模型行为检测等场景。"
---

# AI 模型测试 Skill

## 两种测试模式

| 模式 | 脚本 | 适用场景 |
|------|------|----------|
| **API 模式** | `run_test.py` | 有 API key，快速批量测试 |
| **网页模式** 🆕 | `run_browser_test.py` | 无 API key，用浏览器自动化操作网页 |

---

## 网页模式（新增）

使用 Playwright 控制真实浏览器，打开各 AI 服务网页，自动输入问题并提取回答。

### 已配置的网页服务（9 个）

| 快捷名 | 服务 | 网址 |
|--------|------|------|
| `chatgpt` | ChatGPT | chatgpt.com |
| `claude` | Claude | claude.ai |
| `gemini` | Gemini | gemini.google.com |
| `deepseek` | DeepSeek | chat.deepseek.com |
| `grok` | Grok | x.com/i/grok |
| `copilot` | Copilot | copilot.microsoft.com |
| `perplexity` | Perplexity | perplexity.ai |
| `mistral` | Le Chat (Mistral) | chat.mistral.ai |
| `duckai` | Duck.ai | duck.ai |

对比 API 模式，网页模式能测到 Copilot、Duck.ai 等无 API 的服务。

### 首次使用

```bash
pip install playwright
playwright install chromium
```

### 运行

```bash
# 全部 9 个服务
python ~/.claude/skills/model-tester/run_browser_test.py --input 测试题.docx --services all

# 指定服务
python ~/.claude/skills/model-tester/run_browser_test.py \
    --input 测试题.docx \
    --services chatgpt,claude,deepseek

# 无头模式（后台运行，需提前登录过）
python ~/.claude/skills/model-tester/run_browser_test.py \
    --input 测试题.docx --services all --headless
```

### 交互式工作流（触发 skill 时）

**第一步** — 询问用户选择模式：API 还是网页？

**第二步（网页模式）** — 展示服务表格，让用户选择要测的服务（支持 `all`）

**第三步** — 确认后执行。浏览器会打开，如未登录需用户手动登录。

### 注意事项

- 登录状态通过持久化浏览器数据目录保存（`~/.claude/skills/model-tester/browser-data`）
- 每个问题自动开启新对话，避免上下文干扰
- 网页结构可能随服务更新变化，如提取失败会记录到 Excel 的"错误信息"列

### 数据清洗

```bash
# 清除 output_content 中的题号前缀和语言指令
python ~/.claude/skills/model-tester/clean_results.py --input 汇总.xlsx

# 同时翻译非中文内容为中文
python ~/.claude/skills/model-tester/clean_results.py \
    --input 汇总.xlsx --translate-key sk-xxx
```

清洗内容：
- 删除 `[1.1]` 题号前缀
- 删除 `【重要指令】你必须只使用中文回答...` 等语言指令
- 可选：将非中文回答翻译为中文

### 合并多个测试结果

```bash
# 合并文件夹下所有网页测试结果
python ~/.claude/skills/model-tester/merge_results.py --input 测试结果/

# 按文件名匹配合并
python ~/.claude/skills/model-tester/merge_results.py \
    --input 测试结果/ --pattern "7.2*网页测试*"

# 指定输出文件名
python ~/.claude/skills/model-tester/merge_results.py \
    --input 测试结果/ --output 7.2汇总.xlsx
```

合并后生成统一 Excel（3 个 Sheet），截图也会迁移。

---

## API 模式

## 交互式工作流（通过 Claude Code 触发时必须执行）

当用户通过 Claude Code 触发此 skill 后，**必须按以下顺序交互式操作**：

### 第零步：选择模式

首先询问用户：使用 **API 模式** 还是 **网页模式**？

### 第一步：选择模型/服务

**API 模式** — 不要使用 AskUserQuestion 选择模型（该工具仅支持 4 个选项）。直接展示表格。

```
你想用哪些模型测试？目前支持 9 家厂商：

| # | 厂商 | 快捷命令 | 可用模型 |
|---|------|----------|----------|
| 1 | Anthropic | claude | opus-4-8, sonnet-4-6, haiku-4-5, fable-5 |
| 2 | OpenAI | openai / gpt | gpt-5, gpt-4.1, gpt-4.1-mini, gpt-4o, gpt-4o-mini, o4-mini |
| 3 | Google | google / gemini | gemini-2.5-pro, gemini-2.5-flash |
| 4 | DeepSeek | deepseek | deepseek-chat, deepseek-reasoner |
| 5 | Mistral AI | mistral | mistral-large, medium, small |
| 6 | xAI | xai / grok | grok-3, grok-3-mini |
| 7 | Groq (Llama) | groq / llama | llama-4-maverick, llama-4-scout |
| 8 | Perplexity | perplexity / sonar | sonar, sonar-pro, sonar-reasoning |
| 9 | AI21 Labs | ai21 / jamba | jamba-1.5-large, jamba-1.5-mini |

直接告诉我，例如："测 Claude + GPT"、"全测"、"测 1、4、6"
```

用户可以用各种方式回复：厂商名、编号、快捷命令、"全测"等，自行解析匹配。

### 第二步：收集 API Key

**必须询问用户是否需要异常审核功能**：
> "是否需要开启异常审核？开启后会自动判断每个回答是否有问题，有问题会截图嵌入 Excel。
> 支持两个裁判模型：DeepSeek 或 豆包(Doubao)。请选择模型并提供对应的 API Key。"

- 用户选择 DeepSeek → `--judge-model deepseek --judge-key sk-xxx`
- 用户选择豆包 → `--judge-model doubao --judge-key xxx`
- 用户不需要 → `--no-judge` 跳过审核

**API 模式**还需询问对应厂商的 API key。

### 第三步：确认并运行

展示汇总信息确认：
- 输入文件路径（在当前目录查找 .docx 文件）
- 选择的模型列表
- 输出文件路径

确认后执行：

```bash
python ~/.claude/skills/model-tester/run_test.py \
    --input <docx文件> \
    --output <输出xlsx> \
    --models <模型列表> \
    --key <厂商1>:<key1> \
    --key <厂商2>:<key2> \
    ...
```

**重要**：通过 `--key` 参数传递 API key，这样 key 只在命令行中存在，不会写入环境变量或文件。

### 第四步：报告结果

测试完成后，向用户展示：
- 成功/失败统计
- 按模型汇总的成功率
- Excel 文件路径

---

## 命令行参考（用户直接调用时）

```bash
# 查看所有模型
python ~/.claude/skills/model-tester/run_test.py --list-models

# 指定模型 + key
python ~/.claude/skills/model-tester/run_test.py \
    --input 测试题.docx \
    --models gpt-4o,claude-sonnet-4-6,gemini-2.5-flash \
    --key openai:sk-xxx \
    --key anthropic:sk-ant-xxx \
    --key google:xxx

# 按厂商批量
python ~/.claude/skills/model-tester/run_test.py \
    --input 测试题.docx \
    --providers openai,anthropic,deepseek \
    --key openai:sk-xxx --key anthropic:sk-ant-xxx --key deepseek:sk-xxx
```

| 参数 | 说明 |
|------|------|
| `--input`, `-i` | .docx 文件路径（必需） |
| `--output`, `-o` | 输出 .xlsx 路径 |
| `--models`, `-m` | 模型名，逗号分隔 |
| `--providers`, `-p` | 厂商名，批量包含该厂商所有模型 |
| `--key`, `-k` | API key，格式 `厂商:密钥`，可多次指定 |
| `--delay`, `-d` | 请求间隔秒数（默认 1.0） |
| `--list-models` | 列出所有模型 |

### 输出 Excel 格式

| Sheet | 列 |
|-------|-----|
| **测试详细结果** | 日期、具体时间、题目分类、题号、测试题目、测试模型、是否成功、模型回答、错误信息 |
| **汇总统计** | 模型、总题数、成功、失败、成功率 |
| **按分类汇总** | 题目分类 × 模型（成功/总数） |

格式规范：行高固定 16 磅，表头深蓝底白字，成功行绿底，失败行橙底。

---

## 集成厂商总览

| 厂商 | 模型 | Key 参数 |
|------|------|----------|
| **Anthropic** | claude-opus-4-8, claude-sonnet-4-6, claude-haiku-4-5, claude-fable-5 | `--key anthropic:sk-ant-...` |
| **OpenAI** | gpt-5, gpt-4.1, gpt-4.1-mini, gpt-4o, gpt-4o-mini, o4-mini | `--key openai:sk-...` |
| **Google** | gemini-2.5-pro, gemini-2.5-flash | `--key google:...` |
| **DeepSeek** | deepseek-chat, deepseek-reasoner | `--key deepseek:sk-...` |
| **Mistral AI** | mistral-large, mistral-medium, mistral-small | `--key mistral:...` |
| **xAI** | grok-3, grok-3-mini | `--key xai:...` |
| **Groq (Llama)** | llama-4-maverick, llama-4-scout | `--key groq:...` |
| **Perplexity** | sonar, sonar-pro, sonar-reasoning | `--key perplexity:...` |
| **AI21 Labs** | jamba-1.5-large, jamba-1.5-mini | `--key ai21:...` |

⚠️ 无公开 API（仅供参考）：Microsoft Copilot, DuckDuckGo Duck.ai, Inflection AI Pi, Character.AI
