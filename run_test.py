#!/usr/bin/env python3
"""
AI模型测试脚本 — 多厂商多模型版
读取 .docx 测试题文件，使用多个 AI 模型依次回答，结果保存到格式化 Excel 文件。

用法:
    python run_test.py --input <docx文件> [--output <xlsx输出>] [--models 模型列表] [--providers 厂商列表]

示例:
    # 用默认模型测试
    python run_test.py --input 新一轮测试题.docx

    # 指定模型
    python run_test.py --input 测试题.docx --models gpt-4o,claude-sonnet-4-6,gemini-2.5-flash

    # 按厂商批量测试（测试该厂商所有模型）
    python run_test.py --input 测试题.docx --providers openai,anthropic

    # 列出所有可用模型
    python run_test.py --list-models

必需环境变量（按需设置）:
    ANTHROPIC_API_KEY   — Anthropic Claude
    OPENAI_API_KEY      — OpenAI GPT
    GEMINI_API_KEY      — Google Gemini
    DEEPSEEK_API_KEY    — DeepSeek
    MISTRAL_API_KEY     — Mistral AI
    XAI_API_KEY         — xAI Grok
    GROQ_API_KEY        — Groq (托管 Llama 等开源模型)
    PERPLEXITY_API_KEY  — Perplexity AI
    AI21_API_KEY        — AI21 Labs Jamba
"""

import os
import sys
import argparse
import json
import time
import re
import urllib.request
import urllib.error
from datetime import datetime
from zipfile import ZipFile
from xml.etree import ElementTree
from pathlib import Path


# ============================================================================
# 提供商 & 模型配置（单一数据源）
# ============================================================================

# provider → { api_type, base_url, env_var, models: { short_name → model_id } }
# api_type: "anthropic" | "gemini" | "openai_compatible"

PROVIDERS = {
    "anthropic": {
        "name": "Anthropic",
        "api_type": "anthropic",
        "base_url": "https://api.anthropic.com/v1/messages",
        "env_var": "ANTHROPIC_API_KEY",
        "models": {
            "claude-opus-4-8":        "claude-opus-4-8-20250805",
            "claude-sonnet-4-6":      "claude-sonnet-4-6-20250805",
            "claude-haiku-4-5":       "claude-haiku-4-5-20251001",
            "claude-fable-5":         "claude-fable-5-20251001",
        },
    },
    "openai": {
        "name": "OpenAI",
        "api_type": "openai_compatible",
        "base_url": "https://api.openai.com/v1/chat/completions",
        "env_var": "OPENAI_API_KEY",
        "models": {
            "gpt-5":                  "gpt-5",
            "gpt-4.1":                "gpt-4.1",
            "gpt-4.1-mini":           "gpt-4.1-mini",
            "gpt-4o":                 "gpt-4o",
            "gpt-4o-mini":            "gpt-4o-mini",
            "o4-mini":                "o4-mini",
        },
    },
    "google": {
        "name": "Google",
        "api_type": "gemini",
        "base_url": "https://generativelanguage.googleapis.com/v1beta/models",
        "env_var": "GEMINI_API_KEY",
        "models": {
            "gemini-2.5-pro":         "gemini-2.5-pro",
            "gemini-2.5-flash":       "gemini-2.5-flash",
        },
    },
    "deepseek": {
        "name": "DeepSeek",
        "api_type": "openai_compatible",
        "base_url": "https://api.deepseek.com/v1/chat/completions",
        "env_var": "DEEPSEEK_API_KEY",
        "models": {
            "deepseek-chat":          "deepseek-chat",
            "deepseek-reasoner":      "deepseek-reasoner",
        },
    },
    "mistral": {
        "name": "Mistral AI",
        "api_type": "openai_compatible",
        "base_url": "https://api.mistral.ai/v1/chat/completions",
        "env_var": "MISTRAL_API_KEY",
        "models": {
            "mistral-large":          "mistral-large-latest",
            "mistral-medium":         "mistral-medium-latest",
            "mistral-small":          "mistral-small-latest",
        },
    },
    "xai": {
        "name": "xAI",
        "api_type": "openai_compatible",
        "base_url": "https://api.x.ai/v1/chat/completions",
        "env_var": "XAI_API_KEY",
        "models": {
            "grok-3":                 "grok-3-latest",
            "grok-3-mini":            "grok-3-mini-latest",
        },
    },
    "groq": {
        "name": "Groq (Llama 等)",
        "api_type": "openai_compatible",
        "base_url": "https://api.groq.com/openai/v1/chat/completions",
        "env_var": "GROQ_API_KEY",
        "description": "Meta Llama 模型通过 Groq 云服务调用",
        "models": {
            "llama-4-maverick":       "meta-llama/llama-4-maverick-17b-128e-instruct",
            "llama-4-scout":          "meta-llama/llama-4-scout-17b-16e-instruct",
        },
    },
    "perplexity": {
        "name": "Perplexity AI",
        "api_type": "openai_compatible",
        "base_url": "https://api.perplexity.ai/chat/completions",
        "env_var": "PERPLEXITY_API_KEY",
        "models": {
            "sonar":                  "sonar",
            "sonar-pro":              "sonar-pro",
            "sonar-reasoning":        "sonar-reasoning",
        },
    },
    "ai21": {
        "name": "AI21 Labs",
        "api_type": "openai_compatible",
        "base_url": "https://api.ai21.com/studio/v1/chat/completions",
        "env_var": "AI21_API_KEY",
        "models": {
            "jamba-1.5-large":        "jamba-1.5-large",
            "jamba-1.5-mini":         "jamba-1.5-mini",
        },
    },
}

# ---------------------------------------------------------------------------
# 以下厂商无公开 API，不可编程调用，仅列在 --list-models 中供参考
# ---------------------------------------------------------------------------
NO_API_PROVIDERS = {
    "microsoft": {
        "name": "Microsoft",
        "models": {"copilot": "Copilot (基于 GPT-4，无公开 API)"},
        "note": "无公开 API，仅供手动使用",
    },
    "meta": {
        "name": "Meta",
        "models": {"llama-open": "Llama 开源系列"},
        "note": "开源模型，可通过 Groq / Together AI / Replicate 调用。本工具已内置 Groq 渠道 (groq provider)。",
    },
    "duckduckgo": {
        "name": "DuckDuckGo",
        "models": {"duck-ai": "Duck.ai (聚合 GPT/Claude/Llama/Mistral，无公开 API)"},
        "note": "无公开 API，仅供手动使用",
    },
    "inflection": {
        "name": "Inflection AI",
        "models": {"pi": "Pi (无公开 API)"},
        "note": "无公开 API，仅供手动使用",
    },
    "characterai": {
        "name": "Character.AI",
        "models": {"cai": "自研角色模型 (无公开 API)"},
        "note": "无公开 API，仅供手动使用",
    },
}

# ============================================================================
# 从 PROVIDERS 派生的查找表
# ============================================================================

def build_model_lookup():
    """构建 model_short_name → (provider_key, model_id) 的查找表。"""
    lookup = {}
    for pkey, pinfo in PROVIDERS.items():
        for short, full in pinfo["models"].items():
            if short in lookup:
                print(f"⚠️ 模型名冲突: '{short}' 同时存在于 '{lookup[short][0]}' 和 '{pkey}'")
            lookup[short] = (pkey, full)
    return lookup

MODEL_LOOKUP = build_model_lookup()
ALL_MODEL_NAMES = sorted(MODEL_LOOKUP.keys())

# 默认测试模型（不依赖特定 API key 也能报出清晰的缺失提示）
DEFAULT_MODELS = [
    "gpt-4o",
    "claude-sonnet-4-6",
    "gemini-2.5-flash",
]

# 回答语言配置
LANG_PROMPTS = {
    "zh": "请用中文回答以下问题：\n\n",
    "en": "Please answer the following question in English:\n\n",
    "auto": "",  # 不附加指令，模型自行判断
}

DEFAULT_ANSWER_LANG = "zh"


# ============================================================================
# DOCX 解析（与原版一致）
# ============================================================================

def parse_docx(filepath: str) -> list[dict]:
    """解析 .docx 文件，提取题目及分类层级。
    支持两种格式：
    - 紧凑格式：题号+题目在同一段（如 "1.1 问题文本"）
    - 拆分格式：题号单独一段，题目正文在下一段（如藏语 docx）
    """
    with ZipFile(filepath, 'r') as z:
        with z.open('word/document.xml') as f:
            tree = ElementTree.parse(f)
            root = tree.getroot()

    paragraphs = []
    for p in root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'):
        texts = []
        for t in p.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
            if t.text:
                texts.append(t.text)
        line = ''.join(texts).strip()
        if line:
            paragraphs.append(line)

    questions = []
    current_category = ""
    pending_question = None  # 拆分格式：上一行是独立题号

    for line in paragraphs:
        # 独立的题号（如 "1.1" 独占一行），后面段落是题目正文
        if re.match(r'^\d+\.\d+\s*$', line):
            if pending_question:
                questions.append(pending_question)
            pending_question = {
                "category": current_category,
                "question_num": line.strip(),
                "question_text": "",
            }
            continue

        # 题号+题目在同一行（如 "1.1 问题文本"）
        m = re.match(r'^(\d+\.\d+)\s+(.+)', line)
        if m:
            if pending_question:
                questions.append(pending_question)
                pending_question = None
            questions.append({
                "category": current_category,
                "question_num": m.group(1),
                "question_text": m.group(2).strip(),
            })
            continue

        # 大类标题（如 "1. 标题文本"）
        sec = re.match(r'^(\d+)\.\s+(.+)', line)
        if sec:
            if pending_question:
                questions.append(pending_question)
                pending_question = None
            current_category = line
            continue

        # 独立大类编号（如 "1." 独占一行）
        if re.match(r'^\d+\.\s*$', line):
            if pending_question:
                questions.append(pending_question)
                pending_question = None
            continue

        # 有 pending question，当前行就是题目正文
        if pending_question:
            pending_question["question_text"] = line
            questions.append(pending_question)
            pending_question = None
            continue

    if pending_question:
        questions.append(pending_question)

    return questions


# ============================================================================
# API 调用层 — 三种 API 类型
# ============================================================================

def _call_anthropic_api(model_id: str, question: str, api_key: str,
                        max_tokens: int = 4096, timeout: int = 120) -> dict:
    """Anthropic Messages API。"""
    url = "https://api.anthropic.com/v1/messages"
    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
    }
    body = {
        "model": model_id,
        "max_tokens": max_tokens,
        "messages": [{"role": "user", "content": question}],
    }

    try:
        req = urllib.request.Request(url, data=json.dumps(body).encode(), headers=headers, method="POST")
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            result = json.loads(resp.read().decode())
        text = ""
        for block in result.get("content", []):
            if block.get("type") == "text":
                text += block.get("text", "")
        usage = result.get("usage", {})
        return {"success": True, "response": text.strip(), "error": None,
                "usage": {"input_tokens": usage.get("input_tokens", 0),
                          "output_tokens": usage.get("output_tokens", 0)}}
    except urllib.error.HTTPError as e:
        body = e.read().decode() if e.fp else str(e)
        return {"success": False, "response": "", "error": f"HTTP {e.code}: {body[:500]}", "usage": {}}
    except Exception as e:
        return {"success": False, "response": "", "error": str(e)[:500], "usage": {}}


def _call_gemini_api(model_id: str, question: str, api_key: str,
                     max_tokens: int = 4096, timeout: int = 120) -> dict:
    """Google Gemini API (generateContent)。"""
    url = (f"https://generativelanguage.googleapis.com/v1beta/models/{model_id}:generateContent"
           f"?key={api_key}")
    headers = {"Content-Type": "application/json"}
    body = {
        "contents": [{"parts": [{"text": question}]}],
        "generationConfig": {"maxOutputTokens": max_tokens},
    }

    try:
        req = urllib.request.Request(url, data=json.dumps(body).encode(), headers=headers, method="POST")
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            result = json.loads(resp.read().decode())
        # Gemini 响应格式
        candidates = result.get("candidates", [])
        text = ""
        for c in candidates:
            for part in c.get("content", {}).get("parts", []):
                text += part.get("text", "")
        usage = result.get("usageMetadata", {})
        return {"success": True, "response": text.strip(), "error": None,
                "usage": {"input_tokens": usage.get("promptTokenCount", 0),
                          "output_tokens": usage.get("candidatesTokenCount", 0)}}
    except urllib.error.HTTPError as e:
        body = e.read().decode() if e.fp else str(e)
        return {"success": False, "response": "", "error": f"HTTP {e.code}: {body[:500]}", "usage": {}}
    except Exception as e:
        return {"success": False, "response": "", "error": str(e)[:500], "usage": {}}


def _call_openai_compatible_api(base_url: str, model_id: str, question: str, api_key: str,
                                max_tokens: int = 4096, timeout: int = 120) -> dict:
    """OpenAI 兼容 Chat Completions API（OpenAI / DeepSeek / Mistral / xAI / Groq / Perplexity / AI21）。"""
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    body = {
        "model": model_id,
        "max_tokens": max_tokens,
        "messages": [{"role": "user", "content": question}],
    }

    try:
        req = urllib.request.Request(base_url, data=json.dumps(body).encode(), headers=headers, method="POST")
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            result = json.loads(resp.read().decode())
        text = result["choices"][0]["message"]["content"]
        usage = result.get("usage", {})
        return {"success": True, "response": text.strip(), "error": None,
                "usage": {"input_tokens": usage.get("prompt_tokens", 0),
                          "output_tokens": usage.get("completion_tokens", 0)}}
    except urllib.error.HTTPError as e:
        body = e.read().decode() if e.fp else str(e)
        return {"success": False, "response": "", "error": f"HTTP {e.code}: {body[:500]}", "usage": {}}
    except Exception as e:
        return {"success": False, "response": "", "error": str(e)[:500], "usage": {}}


def call_model(provider_key: str, model_id: str, question: str,
               api_keys: dict, max_tokens: int = 4096, timeout: int = 120) -> dict:
    """根据 provider 类型分发到对应的 API 调用函数。"""
    pinfo = PROVIDERS[provider_key]
    api_type = pinfo["api_type"]
    api_key = api_keys.get(provider_key, "")

    if api_type == "anthropic":
        return _call_anthropic_api(model_id, question, api_key, max_tokens, timeout)
    elif api_type == "gemini":
        return _call_gemini_api(model_id, question, api_key, max_tokens, timeout)
    else:  # openai_compatible
        return _call_openai_compatible_api(pinfo["base_url"], model_id, question, api_key, max_tokens, timeout)


# ============================================================================
# Excel 输出（三 Sheet：详细 / 汇总 / 分类）
# ============================================================================

def save_to_excel(results: list[dict], output_path: str):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("正在安装 openpyxl...")
        os.system(f"{sys.executable} -m pip install openpyxl -q")
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        from openpyxl.utils import get_column_letter

    wb = Workbook()

    # 样式
    hdr_font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
    hdr_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c_font = Font(name="微软雅黑", size=10)
    c_align = Alignment(vertical="top", wrap_text=True)
    c_center = Alignment(horizontal="center", vertical="top", wrap_text=True)
    ok_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    bad_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    border = Border(left=Side("thin", "D9D9D9"), right=Side("thin", "D9D9D9"),
                    top=Side("thin", "D9D9D9"), bottom=Side("thin", "D9D9D9"))

    def write_header(ws, headers):
        for c, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.font = hdr_font; cell.fill = hdr_fill; cell.alignment = hdr_align; cell.border = border

    # ======== Sheet 1: 详细结果 ========
    ws = wb.active
    ws.title = "测试详细结果"
    headers = ["日期", "具体时间", "题目分类", "题号", "测试题目", "测试模型",
               "是否成功", "模型回答", "错误信息"]
    write_header(ws, headers)

    for i, r in enumerate(results, 2):
        vals = [r.get("test_date", ""), r.get("test_time", ""),
                r.get("category", ""), r.get("question_num", ""),
                r.get("question_text", ""), r.get("model_name", ""),
                "✓ 成功" if r.get("success") else "✗ 失败",
                r.get("model_response", ""), r.get("error") or ""]
        row_fill = ok_fill if r.get("success") else bad_fill
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=i, column=c, value=v)
            cell.font = c_font; cell.border = border
            cell.alignment = c_center if c in (1, 2, 4, 6, 7) else c_align
            if c == 7:
                cell.fill = row_fill

    # 行高 16 磅
    for row in ws.iter_rows(min_row=1, max_row=len(results) + 1):
        ws.row_dimensions[row[0].row].height = 16

    for col, w in {1: 14, 2: 12, 3: 28, 4: 8, 5: 48, 6: 22, 7: 10, 8: 58, 9: 28}.items():
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(results) + 1}"

    # ======== Sheet 2: 汇总统计 ========
    ws2 = wb.create_sheet("汇总统计")
    write_header(ws2, ["模型", "总题数", "成功", "失败", "成功率"])

    model_stats = {}
    for r in results:
        m = r.get("model_name", "unknown")
        if m not in model_stats:
            model_stats[m] = {"total": 0, "success": 0, "fail": 0}
        model_stats[m]["total"] += 1
        model_stats[m]["success" if r.get("success") else "fail"] += 1

    for i, (m, s) in enumerate(sorted(model_stats.items()), 2):
        rate = f"{s['success']/s['total']*100:.1f}%" if s["total"] else "N/A"
        for c, v in enumerate([m, s["total"], s["success"], s["fail"], rate], 1):
            cell = ws2.cell(row=i, column=c, value=v)
            cell.font = c_font; cell.alignment = c_center; cell.border = border
    for row in ws2.iter_rows(min_row=1, max_row=len(model_stats) + 1):
        ws2.row_dimensions[row[0].row].height = 16
    for col, w in {1: 24, 2: 10, 3: 8, 4: 8, 5: 10}.items():
        ws2.column_dimensions[get_column_letter(col)].width = w
    ws2.freeze_panes = "A2"

    # ======== Sheet 3: 按分类汇总 ========
    ws3 = wb.create_sheet("按分类汇总")
    all_models = sorted(set(r.get("model_name", "") for r in results))
    write_header(ws3, ["题目分类"] + [f"{m}\n(成功/总数)" for m in all_models])

    cat_stats = {}
    for r in results:
        cat = r.get("category", "未分类")
        cat_stats.setdefault(cat, {}).setdefault(r["model_name"], {"total": 0, "success": 0})
        cat_stats[cat][r["model_name"]]["total"] += 1
        if r.get("success"):
            cat_stats[cat][r["model_name"]]["success"] += 1

    for i, (cat, md) in enumerate(sorted(cat_stats.items()), 2):
        ws3.cell(row=i, column=1, value=cat).font = c_font
        ws3.cell(row=i, column=1).alignment = c_align
        ws3.cell(row=i, column=1).border = border
        for j, mn in enumerate(all_models, 2):
            d = md.get(mn, {"total": 0, "success": 0})
            cell = ws3.cell(row=i, column=j, value=f"{d['success']}/{d['total']}")
            cell.font = c_font; cell.alignment = c_center; cell.border = border
    ws3.column_dimensions["A"].width = 35
    for j in range(2, len(all_models) + 2):
        ws3.column_dimensions[get_column_letter(j)].width = 18
    for row in ws3.iter_rows(min_row=1, max_row=len(cat_stats) + 1):
        ws3.row_dimensions[row[0].row].height = 16
    ws3.freeze_panes = "B2"

    wb.save(output_path)
    print(f"\n✅ Excel 文件已保存到: {output_path}")


# ============================================================================
# 主流程
# ============================================================================

def print_model_list():
    """打印所有模型（含不可用厂商的说明）。"""
    print("\n📋 可编程调用模型（需对应 API key）:\n")
    max_name_len = max(len(k) for k in ALL_MODEL_NAMES) if ALL_MODEL_NAMES else 10
    for pkey, pinfo in PROVIDERS.items():
        print(f"  [{pinfo['name']}]  (env: {pinfo['env_var']})")
        for short, full in pinfo["models"].items():
            print(f"    {short:{max_name_len+2}s} → {full}")
        print()

    print("-" * 70)
    print("\n⚠️  以下厂商无公开 API，不可编程调用，仅供参考:\n")
    for pkey, pinfo in NO_API_PROVIDERS.items():
        print(f"  [{pinfo['name']}]  {pinfo['note']}")
        for short, full in pinfo["models"].items():
            print(f"    {short:{max_name_len+2}s} → {full}")
        print()


def main():
    parser = argparse.ArgumentParser(
        description="AI模型测试工具 — 多厂商多模型版",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=f"""
示例:
  python run_test.py --input 测试题.docx
  python run_test.py --input 测试题.docx --models gpt-4o,claude-sonnet-4-6,gemini-2.5-flash
  python run_test.py --input 测试题.docx --providers openai,anthropic
  python run_test.py --list-models

已集成厂商 (有API): {', '.join(PROVIDERS.keys())}
无API(仅参考): {', '.join(NO_API_PROVIDERS.keys())}
        """,
    )
    parser.add_argument("--input", "-i", help="输入的 .docx 测试题文件路径")
    parser.add_argument("--output", "-o", default=None, help="输出 .xlsx 路径（默认自动生成）")
    parser.add_argument("--models", "-m", default=None,
                        help=f"要测试的模型（逗号分隔）。默认: {','.join(DEFAULT_MODELS)}")
    parser.add_argument("--providers", "-p", default=None,
                        help="按厂商批量测试（逗号分隔），测试该厂商的全部模型")
    parser.add_argument("--delay", "-d", type=float, default=1.0, help="请求间隔秒数（默认 1.0）")
    parser.add_argument("--list-models", action="store_true", help="列出所有可用模型并退出")
    parser.add_argument("--max-tokens", type=int, default=4096, help="最大输出 token 数（默认 4096）")
    parser.add_argument("--timeout", type=int, default=120, help="单次请求超时秒数（默认 120）")
    parser.add_argument("--answer-lang", "-l", default=DEFAULT_ANSWER_LANG,
                        help=f"要求模型回答的语言（默认: {DEFAULT_ANSWER_LANG}）。可选: zh/en/auto")
    parser.add_argument("--key", "-k", action="append", default=None,
                        help="API key，格式: 厂商:密钥。可多次指定。\n"
                             "例如: --key openai:sk-xxx --key anthropic:sk-ant-xxx\n"
                             "优先级: 命令行 --key > 环境变量")

    args = parser.parse_args()

    if args.list_models:
        print_model_list()
        return

    # 必须有输入文件
    if not args.input:
        print("❌ 请指定输入文件: --input <文件路径>")
        print("   使用 --list-models 查看可用模型")
        sys.exit(1)

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"❌ 文件不存在: {args.input}")
        sys.exit(1)

    # --- 确定模型列表 ---
    model_names = []

    if args.providers:
        # 按厂商批量选择
        for p in args.providers.split(","):
            p = p.strip()
            if p not in PROVIDERS:
                print(f"❌ 未知厂商: {p}")
                print(f"   可用厂商: {', '.join(PROVIDERS.keys())}")
                sys.exit(1)
            model_names.extend(PROVIDERS[p]["models"].keys())

    if args.models:
        model_names.extend(m.strip() for m in args.models.split(",") if m.strip())

    if not model_names:
        model_names = DEFAULT_MODELS

    # 去重并校验
    model_names = list(dict.fromkeys(model_names))  # 保序去重
    invalid = [m for m in model_names if m not in MODEL_LOOKUP]
    if invalid:
        print(f"❌ 未知模型: {', '.join(invalid)}")
        print(f"   使用 --list-models 查看完整列表")
        sys.exit(1)

    # 校验回答语言
    if args.answer_lang not in LANG_PROMPTS:
        print(f"❌ 未知语言: {args.answer_lang}")
        print(f"   可选: {', '.join(LANG_PROMPTS.keys())}")
        sys.exit(1)

    # --- 收集所需 API keys ---
    needed_providers = set()
    for mn in model_names:
        needed_providers.add(MODEL_LOOKUP[mn][0])

    # 先解析命令行 --key 参数（优先级最高）
    cli_keys = {}
    if args.key:
        for item in args.key:
            if ":" not in item:
                print(f"❌ --key 格式错误: '{item}'，应为 厂商:密钥")
                print(f"   例如: --key openai:sk-xxx --key anthropic:sk-ant-xxx")
                sys.exit(1)
            p, k = item.split(":", 1)
            p, k = p.strip(), k.strip()
            if p not in PROVIDERS:
                print(f"❌ 未知厂商: '{p}'（--key 参数）")
                print(f"   可用厂商: {', '.join(PROVIDERS.keys())}")
                sys.exit(1)
            cli_keys[p] = k

    # 合并：命令行 > 环境变量
    api_keys = {}
    missing = []
    for pkey in needed_providers:
        pinfo = PROVIDERS[pkey]
        key = cli_keys.get(pkey) or os.environ.get(pinfo["env_var"], "")
        if key:
            api_keys[pkey] = key
        else:
            missing.append((pkey, pinfo))

    if missing:
        print("❌ 缺少以下 API 密钥:\n")
        for pkey, pinfo in missing:
            print(f"   {pinfo['name']} ({pkey})")
            print(f"   方式一: --key {pkey}:<your-key>")
            print(f"   方式二: export {pinfo['env_var']}=<your-key>")
            print(f"   模型: {', '.join(PROVIDERS[pkey]['models'].keys())}\n")
        print(f"   补齐后重新运行即可。")
        sys.exit(1)

    # --- 输出路径 ---
    if args.output is None:
        args.output = f"测试结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    # --- 开始 ---
    print("=" * 70)
    print("  AI 模型测试工具 — 多厂商多模型版")
    print("=" * 70)
    print(f"  输入文件:     {args.input}")
    print(f"  输出文件:     {args.output}")
    print(f"  测试模型:     {', '.join(model_names)}")
    print(f"  请求间隔:     {args.delay} 秒")
    print("-" * 70)

    # 1. 解析
    print("\n📖 正在解析测试题...")
    questions = parse_docx(str(input_path))
    print(f"   共解析出 {len(questions)} 道题目")
    for q in questions:
        cat = (q['category'] or '无分类')[:40]
        print(f"     [{q['question_num']}] {cat}")

    if not questions:
        print("❌ 未解析到任何题目"); sys.exit(1)

    # 回答语言指令
    lang_prompt = LANG_PROMPTS.get(args.answer_lang, LANG_PROMPTS[DEFAULT_ANSWER_LANG])

    # 2. 测试
    total_tasks = len(questions) * len(model_names)
    print(f"\n🚀 开始测试，共 {total_tasks} 个任务（{len(questions)} 题 × {len(model_names)} 个模型）")
    if lang_prompt:
        print(f"   回答语言要求: {args.answer_lang}（将在每道题目前附加语言指令）")
    print()

    all_results = []
    task_count = 0

    for q in questions:
        # 构建完整提示：语言指令 + 原问题
        full_question = lang_prompt + q["question_text"] if lang_prompt else q["question_text"]

        for mn in model_names:
            task_count += 1
            pkey, model_id = MODEL_LOOKUP[mn]
            now = datetime.now()
            test_date = f"{now.year}/{now.month}/{now.day}"
            test_time_detail = now.strftime("%H:%M")

            q_preview = q['question_text'][:60]
            print(f"[{task_count}/{total_tasks}] 📝 [{mn}] {q['question_num']} {q_preview}...",
                  end=" ", flush=True)

            result = call_model(pkey, model_id, full_question, api_keys,
                               max_tokens=args.max_tokens, timeout=args.timeout)

            status = "✅" if result["success"] else "❌"
            u = result.get("usage", {})
            print(f"{status} (in:{u.get('input_tokens',0)} out:{u.get('output_tokens',0)})")

            all_results.append({
                "test_date": test_date,
                "test_time": test_time_detail,
                "category": q["category"],
                "question_num": q["question_num"],
                "question_text": q["question_text"],
                "model_name": mn,
                "model_response": result["response"],
                "success": result["success"],
                "error": result.get("error"),
            })

            if task_count < total_tasks:
                time.sleep(args.delay)

    # 3. 保存
    print(f"\n📊 正在生成 Excel 报告...")
    save_to_excel(all_results, args.output)

    # 4. 总结
    ok = sum(1 for r in all_results if r["success"])
    fail = len(all_results) - ok

    print(f"\n{'=' * 70}")
    print(f"  测试完成!")
    print(f"  总任务数:     {len(all_results)}")
    print(f"  成功:         {ok}")
    print(f"  失败:         {fail}")
    print(f"  结果文件:     {args.output}")
    print(f"{'=' * 70}")


if __name__ == "__main__":
    main()
