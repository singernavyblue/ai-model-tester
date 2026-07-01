#!/usr/bin/env python3
"""
AI模型网页自动化测试脚本
使用 Playwright 控制浏览器，打开各 AI 模型网页，自动输入问题并提取回答。

用法:
    python run_browser_test.py --input <docx文件> [--services 服务列表] [--output <xlsx>]

前置条件:
    1. 安装: pip install playwright && playwright install chromium
    2. 首次运行需在打开的浏览器中手动登录各 AI 服务（登录状态会持久化）

服务快捷名:
    all       — 全部已配置的服务
    chatgpt   — ChatGPT (chat.openai.com)
    claude    — Claude (claude.ai)
    gemini    — Gemini (gemini.google.com)
    deepseek  — DeepSeek (chat.deepseek.com)
    grok      — Grok (x.com/i/grok)
    copilot   — Microsoft Copilot (copilot.microsoft.com)
    perplexity— Perplexity (perplexity.ai)
    mistral   — Mistral (chat.mistral.ai)
    duckai    — Duck.ai (duck.ai)
"""

import os
import sys
import argparse
import json
import time
import re
from datetime import datetime
from pathlib import Path

# ============================================================================
# 服务处理器配置（每个 AI 服务网页的选择器）
# ============================================================================

SERVICE_HANDLERS = {
    "chatgpt": {
        "name": "ChatGPT",
        "provider": "OpenAI",
        "url": "https://chatgpt.com",
        "new_chat_url": "https://chatgpt.com/?model=gpt-4o",
        "input_selector": "#prompt-textarea",
        "input_fallback": 'div[contenteditable="true"].ProseMirror',
        "submit_strategy": "click",
        "submit_selector": 'button[data-testid="send-button"]',
        "submit_enter": True,
        "response_selector": 'div[data-message-author-role="assistant"]',
        "wait_done_strategy": "stop_button_gone",
        "stop_button_selector": 'button[data-testid="stop-button"]',
        "wait_timeout": 120000,
        "extract": "last",
        "post_wait": 2000,
    },
    "claude": {
        "name": "Claude",
        "provider": "Anthropic",
        "url": "https://claude.ai",
        "new_chat_url": "https://claude.ai/new",
        "input_selector": 'div[contenteditable="true"].ProseMirror',
        "submit_strategy": "click",
        "submit_selector": 'button[aria-label="Send Message"]',
        "submit_enter": True,
        "response_selector": 'div.font-claude-message',
        "wait_done_strategy": "stop_button_gone",
        "stop_button_selector": 'button[aria-label="Stop"]',
        "wait_timeout": 120000,
        "extract": "last",
        "post_wait": 2000,
    },
    "gemini": {
        "name": "Gemini",
        "provider": "Google",
        "url": "https://gemini.google.com/app",
        "new_chat_url": "https://gemini.google.com/app",
        "input_selector": 'div[contenteditable="true"]',
        "input_fallback": 'textarea, div[contenteditable="true"]',
        "submit_strategy": "click",
        "submit_selector": 'button[aria-label="Send message"]',
        "submit_enter": True,
        "response_selector": 'model-response, [data-model-response]',
        "response_fallback": 'message-content, .response-content, div.message:last-child',
        "wait_done_strategy": "stable_dom",
        "wait_timeout": 120000,
        "stable_wait": 5000,
        "extract": "last",
        "post_wait": 2000,
    },
    "deepseek": {
        "name": "DeepSeek",
        "provider": "DeepSeek",
        "url": "https://chat.deepseek.com",
        "new_chat_url": "https://chat.deepseek.com",
        "input_selector": 'textarea#chat-input',
        "input_fallback": 'textarea',
        "submit_strategy": "click",
        "submit_selector": 'button[data-action="send"], button:has(svg)',
        "submit_enter": True,
        "response_selector": 'div.ds-markdown',
        "response_fallback": '.message-content, div[class*="message"]',
        "wait_done_strategy": "stable_dom",
        "wait_timeout": 120000,
        "stable_wait": 5000,
        "extract": "all_after_last_prompt",
        "post_wait": 2000,
    },
    "grok": {
        "name": "Grok",
        "provider": "xAI",
        "url": "https://x.com/i/grok",
        "new_chat_url": "https://x.com/i/grok",
        "input_selector": 'textarea, div[contenteditable="true"]',
        "submit_strategy": "enter",
        "submit_enter": True,
        "response_selector": 'div[class*="message"]',
        "wait_done_strategy": "stable_dom",
        "wait_timeout": 120000,
        "stable_wait": 5000,
        "extract": "last",
        "post_wait": 2000,
    },
    "copilot": {
        "name": "Copilot",
        "provider": "Microsoft",
        "url": "https://copilot.microsoft.com",
        "new_chat_url": "https://copilot.microsoft.com",
        "input_selector": 'textarea, div[contenteditable="true"]',
        "submit_strategy": "click",
        "submit_selector": 'button[aria-label="Send message"], button[type="submit"]',
        "submit_enter": True,
        "response_selector": 'div[class*="response"], div[class*="message"], div[class*="answer"]',
        "wait_done_strategy": "stable_dom",
        "wait_timeout": 120000,
        "stable_wait": 5000,
        "extract": "last",
        "post_wait": 2000,
    },
    "perplexity": {
        "name": "Perplexity",
        "provider": "Perplexity AI",
        "url": "https://www.perplexity.ai",
        "new_chat_url": "https://www.perplexity.ai",
        "input_selector": 'textarea',
        "input_fallback": 'div[contenteditable="true"]',
        "submit_strategy": "click",
        "submit_selector": 'button[aria-label="Submit"], button:has(svg[class*="arrow"])',
        "submit_enter": True,
        "response_selector": 'div[class*="prose"], div[class*="response"], div[class*="answer"]',
        "wait_done_strategy": "stable_dom",
        "wait_timeout": 120000,
        "stable_wait": 5000,
        "extract": "last",
        "post_wait": 2000,
    },
    "mistral": {
        "name": "Le Chat",
        "provider": "Mistral AI",
        "url": "https://chat.mistral.ai/chat",
        "new_chat_url": "https://chat.mistral.ai/chat",
        "input_selector": 'textarea',
        "input_fallback": 'div[contenteditable="true"]',
        "submit_strategy": "enter",
        "submit_enter": True,
        "response_selector": 'div[class*="prose"], div[class*="message"]',
        "wait_done_strategy": "stable_dom",
        "wait_timeout": 120000,
        "stable_wait": 5000,
        "extract": "last",
        "post_wait": 2000,
    },
    "duckai": {
        "name": "Duck.ai",
        "provider": "DuckDuckGo",
        "url": "https://duck.ai",
        "new_chat_url": "https://duck.ai",
        "input_selector": 'textarea, div[contenteditable="true"]',
        "submit_strategy": "enter",
        "submit_enter": True,
        "response_selector": 'div[class*="response"], div[class*="message"], div[class*="answer"]',
        "wait_done_strategy": "stable_dom",
        "wait_timeout": 120000,
        "stable_wait": 5000,
        "extract": "last",
        "post_wait": 2000,
    },
    "characterai": {
        "name": "Character.AI",
        "provider": "Character.AI",
        "url": "https://character.ai",
        "new_chat_url": "https://character.ai",
        "input_selector": 'textarea, div[contenteditable="true"]',
        "input_fallback": '[placeholder*="message"], [placeholder*="Message"]',
        "submit_strategy": "enter",
        "submit_enter": True,
        "response_selector": 'div[class*="message"], div[class*="Message"], div[class*="msg"]',
        "response_fallback": 'p, div[class*="response"]',
        "wait_done_strategy": "stable_dom",
        "wait_timeout": 120000,
        "stable_wait": 6000,
        "extract": "last",
        "post_wait": 3000,
    },
    "meta": {
        "name": "Meta AI",
        "provider": "Meta",
        "url": "https://www.meta.ai",
        "new_chat_url": "https://www.meta.ai",
        "input_selector": 'textarea, div[contenteditable="true"]',
        "input_fallback": '[role="textbox"], input[type="text"]',
        "submit_strategy": "enter",
        "submit_enter": True,
        "response_selector": 'div[class*="message"], div[class*="response"], div[class*="prose"]',
        "wait_done_strategy": "stable_dom",
        "wait_timeout": 180000,
        "stable_wait": 12000,
        "extract": "full_page",
        "post_wait": 5000,
        "scroll_to_bottom": True,
    },
}


# 回答语言配置
LANG_PROMPTS = {
    "zh": "【重要指令】你必须只使用中文回答下面的问题。不要使用任何其他语言，无论问题是什么语言，你的回答都必须是中文。\n\n",
    "en": "【IMPORTANT】You must answer the following question in English only. Do not use any other language.\n\n",
    "auto": "",
}

# ============================================================================
# 浏览器自动化核心
# ============================================================================

def ensure_playwright():
    """确保 Playwright 已安装。"""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("正在安装 playwright...")
        os.system(f"{sys.executable} -m pip install playwright -q")
        print("正在安装 Chromium 浏览器...")
        os.system(f"{sys.executable} -m playwright install chromium")
        from playwright.sync_api import sync_playwright
    return sync_playwright


def find_element(page, selectors, timeout=5000):
    """尝试多个选择器，返回第一个找到的元素。"""
    if isinstance(selectors, str):
        selectors = [selectors]
    for sel in selectors:
        try:
            el = page.wait_for_selector(sel, timeout=timeout, state="visible")
            if el:
                return el
        except Exception:
            continue
    return None


def wait_for_response_completion(page, handler, timeout=120000):
    """等待 AI 回答完成。"""
    strategy = handler.get("wait_done_strategy", "stable_dom")
    start = time.time()

    if strategy == "stop_button_gone":
        # 等待"停止生成"按钮消失（表示回答完成）
        stop_sel = handler.get("stop_button_selector", "")
        if stop_sel:
            try:
                # 先等待停止按钮出现（说明正在生成）
                page.wait_for_selector(stop_sel, timeout=15000, state="visible")
            except Exception:
                pass
            # 再等待它消失
            try:
                page.wait_for_selector(stop_sel, timeout=timeout, state="detached")
                time.sleep(handler.get("post_wait", 1000) / 1000)
                return True
            except Exception:
                # 超时也继续，可能回答已经完成但选择器变了
                pass

    elif strategy == "stable_dom":
        # 等待 DOM 稳定（无新内容出现）
        stable_wait = handler.get("stable_wait", 5000) / 1000
        time.sleep(stable_wait)

        # 再额外等一会儿确保流式输出完成
        try:
            # 尝试找"复制"按钮或类似已完成标记
            page.wait_for_selector(
                'button[aria-label="Copy"], button[title="Copy"], [data-testid="copy-button"]',
                timeout=min(timeout - (time.time() - start) * 1000, 30000)
            )
            time.sleep(1)
        except Exception:
            pass
        return True

    # 默认：等待足够时间
    remaining = timeout / 1000 - (time.time() - start)
    if remaining > 0:
        time.sleep(min(remaining, handler.get("stable_wait", 5000) / 1000))
    return True


def extract_response(page, handler):
    """提取当前页面中的 AI 回答。"""
    extract_mode = handler.get("extract", "last")
    selectors = handler.get("response_selector", "")
    fallback = handler.get("response_fallback", "")
    all_selectors = [selectors] if isinstance(selectors, str) else selectors
    if fallback:
        all_selectors.extend([fallback] if isinstance(fallback, str) else fallback)

    # 滚动到底部，确保懒加载内容全部出现
    if handler.get("scroll_to_bottom"):
        try:
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            time.sleep(1)
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            time.sleep(0.5)
        except Exception:
            pass

    texts = []
    for sel in all_selectors:
        try:
            els = page.query_selector_all(sel)
            for el in els:
                # 用 text_content 获取全部文本（比 inner_text 更完整）
                t = el.text_content().strip()
                if t and len(t) > 5:
                    texts.append(t)
            if texts:
                break
        except Exception:
            continue

    if not texts:
        # 兜底：获取页面主体文本
        try:
            body = page.query_selector("body")
            if body:
                return body.inner_text().strip()[-5000:]
        except Exception:
            pass
        return "[未能提取回答]"

    if extract_mode == "last":
        return texts[-1]
    elif extract_mode == "full_page":
        return "\n\n".join(texts)
    else:
        return "\n\n".join(texts) if len(texts) > 1 else texts[-1]


def send_question(page, handler, question_text, timeout=120000):
    """在指定页面输入问题、提交、等待回答完成、提取回答。"""
    # 1. 找到输入框
    input_selectors = [handler["input_selector"]]
    if handler.get("input_fallback"):
        if isinstance(handler["input_fallback"], list):
            input_selectors.extend(handler["input_fallback"])
        else:
            input_selectors.append(handler["input_fallback"])

    input_el = find_element(page, input_selectors, timeout=10000)
    if not input_el:
        # 兜底：尝试点击页面再找
        page.click("body", position={"x": 500, "y": 500})
        time.sleep(2)
        input_el = find_element(page, input_selectors, timeout=10000)
    if not input_el:
        return {"success": False, "response": "", "error": "找不到输入框，页面可能已改版"}

    # 2. 清空并输入问题
    try:
        input_el.click()
        time.sleep(0.5)
        # 尝试清空
        input_el.fill("")
        # 输入问题（可能很长，用 type 或 fill）
        try:
            input_el.fill(question_text)
        except Exception:
            # fallback: 逐字输入
            input_el.type(question_text, delay=10)
        time.sleep(0.5)
    except Exception as e:
        return {"success": False, "response": "", "error": f"输入失败: {e}"}

    # 3. 提交
    submit_strategy = handler.get("submit_strategy", "enter")
    if submit_strategy == "click":
        submit_sel = handler.get("submit_selector", "")
        if submit_sel:
            try:
                btn = find_element(page, submit_sel, timeout=5000)
                if btn:
                    btn.click()
                else:
                    input_el.press("Enter")
            except Exception:
                input_el.press("Enter")
        else:
            input_el.press("Enter")
    else:
        input_el.press("Enter")

    time.sleep(2)

    # 4. 等待回答完成
    try:
        wait_for_response_completion(page, handler, timeout)
    except Exception as e:
        pass  # 即使超时也尝试提取

    # 5. 提取回答
    response_text = extract_response(page, handler)

    if response_text and response_text != "[未能提取回答]":
        return {"success": True, "response": response_text, "error": None}
    else:
        return {"success": False, "response": response_text,
                "error": "未能提取有效回答，需手动检查页面"}


def run_browser_test(services: list[str], questions: list[dict],
                     output_path: str, headless: bool = False,
                     user_data_dir: str = None, question_delay: float = 3.0,
                     answer_lang: str = "zh", doc_lang: str = ""):
    """主测试流程：对每个服务、每个问题执行网页自动化测试。"""
    sync_playwright = ensure_playwright()

    if user_data_dir is None:
        user_data_dir = str(Path.home() / ".claude" / "skills" / "model-tester" / "browser-data")

    lang_prompt = LANG_PROMPTS.get(answer_lang, "")

    all_results = []

    with sync_playwright() as p:
        # 使用持久化上下文，保留登录状态
        context = p.chromium.launch_persistent_context(
            user_data_dir=user_data_dir,
            headless=headless,
            args=["--no-sandbox", "--disable-blink-features=AutomationControlled"],
            viewport={"width": 1280, "height": 900},
            slow_mo=100,
        )

        page = context.new_page()

        for svc_key in services:
            handler = SERVICE_HANDLERS[svc_key]
            print(f"\n{'=' * 60}")
            print(f"  🖥️  [{handler['name']}] ({handler['provider']})")
            print(f"     打开: {handler['new_chat_url']}")
            print(f"{'=' * 60}")

            # 打开新聊天页面
            try:
                page.goto(handler["new_chat_url"], wait_until="domcontentloaded", timeout=30000)
                time.sleep(3)  # 等页面完全加载
            except Exception as e:
                print(f"  ⚠️ 页面加载失败: {e}")
                # 尝试备用 URL
                try:
                    page.goto(handler["url"], wait_until="domcontentloaded", timeout=30000)
                    time.sleep(3)
                except Exception as e2:
                    print(f"  ❌ 无法访问该服务，跳过")
                    for q in questions:
                        all_results.append(_fail_result(q, svc_key, handler,
                                                        f"无法访问: {e2}", doc_lang))
                    continue

            # 检查是否需要登录
            # 简单启发式：检测输入框是否存在
            input_exists = find_element(page, handler["input_selector"], timeout=5000)
            if not input_exists and handler.get("input_fallback"):
                input_exists = find_element(page, handler["input_fallback"], timeout=5000)

            if not input_exists:
                print(f"  ⚠️ 未找到输入框，可能未登录或页面已改版")
                print(f"     请在浏览器中手动登录 {handler['name']}，然后按 Enter 继续...")
                if not headless:
                    input("  >>> 按 Enter 继续 <<<")
                    page.goto(handler["new_chat_url"], wait_until="domcontentloaded", timeout=30000)
                    time.sleep(3)

            # 逐题测试
            for qi, q in enumerate(questions):
                q_num = q["question_num"]
                q_text = lang_prompt + q["question_text"] if lang_prompt else q["question_text"]
                now = datetime.now()
                test_date = f"{now.year}/{now.month}/{now.day}"
                test_time = now.strftime("%H:%M")

                print(f"  [{qi+1}/{len(questions)}] [{q_num}] {q_text[:50]}...",
                      end=" ", flush=True)

                result = send_question(page, handler, q_text,
                                       timeout=handler.get("wait_timeout", 120000))

                status = "✅" if result["success"] else "❌"
                print(status)

                all_results.append({
                    "test_date": test_date,
                    "test_time": test_time,
                    "category": q["category"],
                    "question_num": q_num,
                    "doc_lang": doc_lang,
                    "question_text": q_text,
                    "model_name": f"{handler['name']} (网页)",
                    "model_response": result["response"],
                    "success": result["success"],
                    "error": result.get("error"),
                })

                # 题间延迟
                time.sleep(question_delay)

                # 每道题后刷新页面（开新对话），避免上下文干扰
                if qi < len(questions) - 1:
                    try:
                        page.goto(handler["new_chat_url"],
                                  wait_until="domcontentloaded", timeout=20000)
                        time.sleep(2)
                    except Exception:
                        pass

        context.close()

    # 保存 Excel
    save_to_excel(all_results, output_path)
    return all_results


def _fail_result(question: dict, svc_key: str, handler: dict, error: str, doc_lang: str = "") -> dict:
    """生成失败记录。"""
    now = datetime.now()
    return {
        "test_date": f"{now.year}/{now.month}/{now.day}",
        "test_time": now.strftime("%H:%M"),
        "category": question.get("category", ""),
        "question_num": question.get("question_num", ""),
        "doc_lang": doc_lang,
        "question_text": question.get("question_text", ""),
        "model_name": f"{handler['name']} (网页)",
        "model_response": "",
        "success": False,
        "error": error,
    }


# ============================================================================
# Excel 输出（复用 API 模式的格式）
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
    headers = ["日期", "具体时间", "题目分类", "题号", "语言", "测试题目", "测试模型",
               "是否成功", "模型回答", "错误信息"]
    write_header(ws, headers)

    for i, r in enumerate(results, 2):
        vals = [r.get("test_date", ""), r.get("test_time", ""),
                r.get("category", ""), r.get("question_num", ""),
                r.get("doc_lang", ""),
                r.get("question_text", ""), r.get("model_name", ""),
                "✓ 成功" if r.get("success") else "✗ 失败",
                r.get("model_response", ""), r.get("error") or ""]
        row_fill = ok_fill if r.get("success") else bad_fill
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=i, column=c, value=v)
            cell.font = c_font; cell.border = border
            cell.alignment = c_center if c in (1, 2, 4, 5, 7, 8) else c_align
            if c == 8:
                cell.fill = row_fill

    for row in ws.iter_rows(min_row=1, max_row=len(results) + 1):
        ws.row_dimensions[row[0].row].height = 16

    for col, w in {1: 14, 2: 12, 3: 28, 4: 8, 5: 10, 6: 48, 7: 22, 8: 10, 9: 58, 10: 28}.items():
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(results) + 1}"

    # ======== Sheet 2: 汇总统计 ========
    ws2 = wb.create_sheet("汇总统计")
    write_header(ws2, ["模型", "总题数", "成功", "失败", "成功率"])
    model_stats = {}
    for r in results:
        m = r.get("model_name", "unknown")
        model_stats.setdefault(m, {"total": 0, "success": 0, "fail": 0})
        model_stats[m]["total"] += 1
        model_stats[m]["success" if r.get("success") else "fail"] += 1
    for i, (m, s) in enumerate(sorted(model_stats.items()), 2):
        rate = f"{s['success']/s['total']*100:.1f}%" if s["total"] else "N/A"
        for c, v in enumerate([m, s["total"], s["success"], s["fail"], rate], 1):
            cell = ws2.cell(row=i, column=c, value=v)
            cell.font = c_font; cell.alignment = c_center; cell.border = border
    for row in ws2.iter_rows(min_row=1, max_row=len(model_stats) + 1):
        ws2.row_dimensions[row[0].row].height = 16
    for col, w in {1: 28, 2: 10, 3: 8, 4: 8, 5: 10}.items():
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
# DOCX 解析
# ============================================================================

def parse_docx(filepath: str) -> list[dict]:
    """解析 .docx 文件，提取题目及分类层级。
    支持两种格式：
    - 紧凑格式：题号+题目在同一段（如 "1.1 问题文本"）
    - 拆分格式：题号单独一段，题目正文在下一段（如藏语 docx）
    """
    from zipfile import ZipFile
    from xml.etree import ElementTree

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
        if re.match(r'^\d+\.\d+\.?\s*$', line):
            if pending_question:
                questions.append(pending_question)
            pending_question = {
                "category": current_category,
                "question_num": line.strip(),
                "question_text": "",
            }
            continue

        # 题号+题目在同一行（如 "1.1 问题文本" 或 "1.1问题文本" 无空格）
        m = re.match(r'^(\d+\.\d+)\.?\s*(.+)', line)
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

        # 大类标题（如 "1. 标题文本" 或 "1.标题文本" 无空格）
        sec = re.match(r'^(\d+)\.\s*(.+)', line)
        if sec:
            if pending_question:
                questions.append(pending_question)
                pending_question = None
            current_category = line
            continue

        # 独立大类编号（如 "1." 独占一行），后面是标题
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
# 主入口
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="AI模型网页自动化测试工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 用全部服务测试
  python run_browser_test.py --input 测试题.docx --services all

  # 指定服务
  python run_browser_test.py --input 测试题.docx --services chatgpt,claude,deepseek

  # 后台运行（不显示浏览器窗口）
  python run_browser_test.py --input 测试题.docx --services all --headless

可用服务:
  all, chatgpt, claude, gemini, deepseek, grok, copilot, perplexity, mistral, duckai
        """,
    )
    parser.add_argument("--input", "-i", help="输入的 .docx 测试题文件路径")
    parser.add_argument("--output", "-o", default=None, help="输出 .xlsx 路径（默认自动生成）")
    parser.add_argument("--services", "-s", default="all",
                        help="要测试的服务，逗号分隔（默认 all）")
    parser.add_argument("--delay", "-d", type=float, default=3.0,
                        help="题间延迟秒数（默认 3.0）")
    parser.add_argument("--headless", action="store_true",
                        help="无头模式（不显示浏览器窗口，需已登录）")
    parser.add_argument("--answer-lang", "-l", default="zh",
                        help="要求模型回答的语言（默认: zh）。可选: zh/en/auto")
    parser.add_argument("--doc-lang", default="",
                        help="测试题文档的语言标识（如 中文/英文/蒙古语/藏语/哈萨克语/维吾尔语）")
    parser.add_argument("--user-data-dir",
                        help="浏览器用户数据目录（默认 ~/.claude/skills/model-tester/browser-data）")
    parser.add_argument("--list-services", action="store_true",
                        help="列出所有已配置的服务并退出")

    args = parser.parse_args()

    if args.list_services:
        print("\n📋 已配置的网页服务:\n")
        for key, h in SERVICE_HANDLERS.items():
            print(f"  {key:12s} → {h['name']} ({h['provider']})")
            print(f"             {h['new_chat_url']}")
        print()
        return

    # 解析服务列表
    if args.services == "all":
        services = list(SERVICE_HANDLERS.keys())
    else:
        services = [s.strip() for s in args.services.split(",")]
        invalid = [s for s in services if s not in SERVICE_HANDLERS]
        if invalid:
            print(f"❌ 未知服务: {', '.join(invalid)}")
            print(f"   可用: {', '.join(SERVICE_HANDLERS.keys())}")
            sys.exit(1)

    # 校验回答语言
    if args.answer_lang not in LANG_PROMPTS:
        print(f"❌ 未知语言: {args.answer_lang}")
        print(f"   可选: {', '.join(LANG_PROMPTS.keys())}")
        sys.exit(1)

    # 输入文件
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"❌ 文件不存在: {args.input}")
        sys.exit(1)

    # 输出路径
    if args.output is None:
        # 输出到与"测试题"文件夹同级的"测试结果"目录
        p = input_path.parent
        while p.name != "测试题" and p.parent != p:
            p = p.parent
        out_dir = (p.parent if p.name == "测试题" else input_path.parent) / "测试结果"
        out_dir.mkdir(parents=True, exist_ok=True)
        args.output = str(out_dir / f"网页测试结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

    # 解析题目
    print("📖 正在解析测试题...")
    questions = parse_docx(str(input_path))
    print(f"   共 {len(questions)} 道题目")
    for q in questions:
        print(f"     [{q['question_num']}] {(q['category'] or '无分类')[:40]}")

    # 检查 Playwright
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("\n⚠️ 需要安装 Playwright:")
        print("   pip install playwright")
        print("   playwright install chromium")
        sys.exit(1)

    # 提示
    total = len(services) * len(questions)
    print(f"\n🚀 将用 {len(services)} 个服务测试 {len(questions)} 道题，共 {total} 个任务")
    if not args.headless:
        print(f"⚠️  将打开浏览器窗口，请确保已在各服务中登录")
        print(f"   浏览器数据目录: {args.user_data_dir or '~/.claude/skills/model-tester/browser-data'}")
    print()

    # 执行
    results = run_browser_test(
        services=services,
        questions=questions,
        output_path=args.output,
        headless=args.headless,
        user_data_dir=args.user_data_dir,
        question_delay=args.delay,
        answer_lang=args.answer_lang,
        doc_lang=args.doc_lang,
    )

    # 总结
    ok = sum(1 for r in results if r["success"])
    fail = len(results) - ok
    print(f"\n{'=' * 60}")
    print(f"  网页测试完成!")
    print(f"  总任务数: {len(results)}")
    print(f"  成功: {ok}")
    print(f"  失败: {fail}")
    print(f"  结果文件: {args.output}")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
