# AI Model Tester

使用多个 AI 模型对测试题进行回答，支持 **API 调用** 和 **浏览器网页自动化** 两种模式，结果保存到格式化 Excel。

## 两种模式

| 模式 | 脚本 | 需要 |
|------|------|------|
| **API 模式** | `run_test.py` | API Key |
| **网页模式** | `run_browser_test.py` | 浏览器登录 |

## API 模式 — 快速开始

```bash
export ANTHROPIC_API_KEY=sk-ant-...
export OPENAI_API_KEY=sk-...

python run_test.py --input 测试题.docx --models gpt-4o,claude-sonnet-4-6,gemini-2.5-flash
```

支持 9 家厂商 26 个模型：OpenAI、Anthropic、Google、DeepSeek、Mistral、xAI、Groq(Llama)、Perplexity、AI21 Labs。

## 网页模式 — 快速开始

```bash
pip install playwright
playwright install chromium

python run_browser_test.py --input 测试题.docx --services all
```

支持 11 个网页服务：ChatGPT、Claude、Gemini、DeepSeek、Grok、Copilot、Perplexity、Mistral、Duck.ai、Meta AI、Character.AI。

## 输出 Excel

| Sheet | 内容 |
|-------|------|
| 测试详细结果 | 日期、具体时间、题目分类、题号、测试题目、测试模型、是否成功、模型回答 |
| 汇总统计 | 模型、总题数、成功、失败、成功率 |
| 按分类汇总 | 题目分类 × 模型交叉统计 |

## 作为 Claude Code Skill

将此仓库克隆到 `~/.claude/skills/model-tester/`，即可通过 Claude Code 触发：

```
/skill model-tester
```

或直接说"帮我测一下这些题目"、"用不同模型跑个对比测试"等。
