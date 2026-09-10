"""
AI 小红书内容增长工作台 (Streamlit)
============================================
基于「小红书数据采集工具」二次开发：保留原有真实能力
  · 小红书登录（采集时自动处理，cookies.json 持久化）
  · 批量采集竞品帖子（Playwright）
  · Excel 上传 / 数据处理
  · 高清无水印原图下载
并重塑为 7 页产品结构：
  · 首页：流程入口 + 当前任务 + 最近任务
  · 小红书采集：登录状态 + 实时采集进度
  · 帖子数据：当前真实帖子列表 + 下载 Excel + 生成爆款分析表
  · 爆款分析：爆款卡片 + 详情面板 + 基于此爆款生成内容
  · 内容生成：产品/用户/场景 → 小红书内容提示词
  · 历史记录：任务卡片 + 已生成提示词
  · 设置：Excel 导入 + 数据管理

核心链路：
  小红书采集 → 帖子数据 → 爆款分析表 → 内容生成 → 小红书内容提示词

运行：streamlit run app.py  （http://localhost:8501）
"""
import io
import json
import os
import random
import re
import subprocess
import sys
import threading
import time
import uuid
from datetime import datetime

import pandas as pd
import streamlit as st

from xhs_scraper import XHSScraper
from analyzer import analyze_note, compute_heat, summarize_rules
from generate_content import gen_notes, build_content_model, load_brand, gen_one
from generate_prompt import build_prompt, save_prompt, load_prompts

# ---- 云端兼容：Playwright Chromium 缺失时自动补装（Streamlit Cloud 等环境）----
# 本机已安装时此函数为纯检查（毫秒级），不影响正常启动；
# 仅当浏览器缺失时才会触发一次下载安装，保证采集功能在云端可用。
def _ensure_chromium():
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            if os.path.exists(p.chromium.executable_path):
                return
    except Exception:
        pass
    try:
        subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            timeout=900, check=True,
        )
        print("[bootstrap] playwright chromium installed")
    except Exception as e:
        print(f"[bootstrap] playwright chromium install failed: {e}")


_ensure_chromium()

# ---- 云端环境检测：Streamlit Community Cloud 将仓库挂载到 /mount/src，且为无显示器 Linux 容器。
# Chromium 在无显示环境只能以 headless 模式运行（有头模式会因无法创建显示器而失败），
# 因此云端采集/图片下载一律强制无头；本机有桌面环境时保持有头（扫码弹窗）。
def _is_cloud_env():
    return os.path.exists("/mount/src") or os.environ.get("STREAMLIT_CLOUD") == "1"


HERE = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(HERE, "xhs_notes.json")
MASTER_FILE = os.path.join(HERE, "xhs_master.json")
DOWNLOAD_DIR = os.path.join(HERE, "downloads")
COOKIE_FILE = os.path.join(HERE, "cookies.json")
# 登录验证文件：只有真正扫码登录成功后才会写入。
# 仅依赖 cookies.json 会把"曾经登录过"误判为"当前已登录"——
# web_session 在 Playwright 持久化时可能 expires=-1，无法靠 expires 判断是否还有效。
# 加入 LOGIN_FLAG + 30 天有效期，既准确，又不会让用户每次启动都要重新扫码。
LOGIN_FLAG = os.path.join(HERE, "xhs_login.json")
LOGIN_VALID_SECONDS = 30 * 86400  # 30 天
TASKS_FILE = os.path.join(HERE, "xhs_tasks.json")
TASK_DIR = os.path.join(HERE, "tasks")
SCRAPE_PROGRESS = os.path.join(HERE, "_scrape_progress.txt")
SCRAPE_STATUS = os.path.join(HERE, "_scrape_status.json")
# 云端无头模式下，采集等待扫码期间的登录页截图（含二维码），由进度页 st.image 展示
SCRAPE_QR = os.path.join(HERE, "_scrape_qr.png")
IMG_PROGRESS = os.path.join(HERE, "_img_progress.txt")
IMG_STATUS = os.path.join(HERE, "_img_status.json")
IMG_RESULT = os.path.join(HERE, "_img_result.json")
REPORT_FILE = os.path.join(HERE, "xhs_report.xlsx")
CONTENT_FILE = os.path.join(HERE, "xhs_content.xlsx")
PROMPT_FILE = os.path.join(HERE, "xhs_prompts.json")
# 当前工作上下文快照：只记最近一次有效任务的 task_id。
# 采集/导入/加载数据落盘时同步更新；应用刷新/重启后据此自动恢复，避免"刷新数据就没了"。
CUR_CTX_FILE = os.path.join(HERE, "_cur_ctx.json")

# ---- 访客模式（每个访问者用自己的小红书账号，登录态相互隔离）----
SESSION_DIR = os.path.join(HERE, "sessions")


def session_cookie_path(vid):
    return os.path.join(SESSION_DIR, f"{vid}.json")


def session_qr_path(vid):
    return os.path.join(SESSION_DIR, f"{vid}_qr.png")


def visitor_status_file(vid):
    return os.path.join(HERE, f"_qr_{vid}.json")


def visitor_progress_file(vid):
    return os.path.join(HERE, f"_qr_{vid}.txt")


def read_visitor_cookie(vid):
    """读取某访客独立登录态文件，返回 cookies 列表（无效/缺失返回 None）。"""
    p = session_cookie_path(vid)
    if not os.path.exists(p):
        return None
    try:
        cookies = json.load(open(p, encoding="utf-8"))
    except Exception:
        return None
    has = any(isinstance(c, dict) and c.get("name") == "web_session" and c.get("value")
              for c in cookies)
    return cookies if has else None


def is_local_browser():
    """判断当前浏览器会话是否从本机（localhost / 127.0.0.1 / [::1]）访问。

    用于默认采集身份选择：本机默认「主人账号」，经公网链接访问默认「访客扫码」。
    st.context（Streamlit ≥1.36）提供当前页面 url 与请求头；拿不到时保守返回 False。
    """
    try:
        ctx = st.context
        url = str(getattr(ctx, "url", "") or "")
        if url.startswith(("http://localhost", "http://127.0.0.1", "http://[::1]")):
            return True
        try:
            h = str((ctx.headers or {}).get("Host") or "")
        except Exception:
            h = ""
        if any(k in h for k in ("localhost", "127.0.0.1", "[::1]")):
            return True
    except Exception:
        pass
    return False


def account_identity():
    """汇总当前会话的采集账号身份，供顶栏 / 首页 / 采集页 / 门禁统一判断。

    单身份模式（2026-09-09 按用户要求回退）：不再区分主人/访客，
    一律使用本机保存的 cookies.json；未登录时点「开始采集」自动走扫码登录。
    保留原有返回结构，旧调用点无需改动。
    """
    vid = st.session_state.get("visitor_id") or ""
    return {
        "vid": vid,
        "choice": "main",
        "main_logged": check_login(),
        "visitor_bound": False,
    }


def cleanup_sessions(days=3):
    """清理过期的访客会话文件（sessions/*.json 保留 3 天，_qr_* / 二维码图 1 天）。

    vid 已沉淀到 URL(?vid=…)，正常刷新不产生孤儿文件；这里兜底清理长期不回的访客，
    防止 sessions/ 无限膨胀。每次进程启动最多执行一次（按日期标记文件）。
    """
    try:
        mark = os.path.join(HERE, "_sessions_cleanup.day")
        today = time.strftime("%Y-%m-%d")
        if os.path.exists(mark) and open(mark, encoding="utf-8").read().strip() == today:
            return
        now = time.time()
        for d in (SESSION_DIR,):
            if os.path.isdir(d):
                for fn in os.listdir(d):
                    p = os.path.join(d, fn)
                    try:
                        if os.path.getmtime(p) < now - days * 86400:
                            os.remove(p)
                    except Exception:
                        pass
        for fn in os.listdir(HERE):
            if fn.startswith("_qr_") and fn.endswith((".json", ".txt", ".png")):
                p = os.path.join(HERE, fn)
                try:
                    if os.path.getmtime(p) < now - 86400:
                        os.remove(p)
                except Exception:
                    pass
        with open(mark, "w", encoding="utf-8") as f:
            f.write(today)
    except Exception:
        pass

NAV = ["首页", "小红书采集", "帖子数据", "爆款分析", "内容生成", "历史记录", "设置"]
GOALS = ["竞品内容分析", "爆款内容分析", "用户需求分析", "内容机会分析"]

# 页面标题（顶栏「当前页面」使用，与 NAV 一一对应）
PAGE_TITLES = {
    "首页": "首页 · 产品入口",
    "小红书采集": "小红书采集 · 登录与配置",
    "帖子数据": "帖子数据 · 当前任务",
    "爆款分析": "爆款分析 · 内容规律",
    "内容生成": "内容生成 · 小红书内容提示词",
    "历史记录": "历史记录 · 全部任务",
    "设置": "设置 · 数据导入",
}

# 小红书红 / 浅灰 / 白 的企业级 SaaS 主题
st.set_page_config(page_title="AI 小红书内容增长工作台", layout="wide", initial_sidebar_state="expanded")
st.markdown("""
<style>
:root{ --xr:#FF2442; --xr2:#C00000; --ink:#1f1f1f; --sub:#666; --line:#ececec; --bg:#fafafa; }
*{font-family:-apple-system,"PingFang SC","Microsoft YaHei",Segoe UI,sans-serif;}
.stApp{background:var(--bg);}
header[data-testid="stHeader"]{background:transparent;}
.kpi{background:#fff;border:1px solid var(--line);border-radius:12px;padding:14px 16px;}
.kpi .v{font-size:26px;font-weight:700;color:var(--xr2);}
.kpi .l{font-size:12px;color:var(--sub);margin-top:2px;}
.sec-title{font-size:18px;font-weight:700;color:var(--ink);margin:0 0 10px;}
.tag{display:inline-block;background:#fff0f2;color:var(--xr2);border:1px solid #ffd6dc;
      border-radius:20px;padding:2px 10px;font-size:12px;margin:2px 4px 2px 0;}
.muted{color:var(--sub);font-size:13px;}
.navbtn button{font-weight:600;}
/* 首页快速入口等高卡片：固定 caption 容器高度，让底部按钮对齐到同一行 */
.qc-card{background:#fff;border:1px solid var(--line);border-radius:12px;padding:14px 16px;height:100%;box-sizing:border-box;}
.qc-card h3{margin:0 0 6px;font-size:16px;font-weight:700;color:var(--ink);display:flex;align-items:center;gap:6px;}
.qc-cap{font-size:13px;color:var(--sub);line-height:1.55;min-height:64px;margin-bottom:8px;}
/* 节点卡片：略高，加色边线让"已完成"更醒目 */
.flow-node{background:#fff;border:1px solid var(--line);border-radius:12px;padding:14px 12px;height:100%;box-sizing:border-box;}
.flow-node.done{border-color:#ffd6dc;background:#fff8f9;}
.flow-icon{font-size:22px;}
.flow-title{font-weight:700;margin-top:4px;color:var(--ink);}
.flow-sub{font-size:12px;color:var(--sub);margin:4px 0 6px;line-height:1.5;min-height:42px;}
.flow-tag{font-size:12px;color:var(--sub);}
.flow-tag.done{color:var(--xr2);}
</style>
""", unsafe_allow_html=True)


# ======================================================================
# 会话状态初始化
# ======================================================================
def init_state():
    defaults = {
        "page": "首页",
        "notes": [],
        "analysis_summary": None,
        "current_task": None,
        "tasks": load_tasks(),
        "scraping": False,
        "scrape_cfg": None,
        "downloading_images": False,
        "selected_note_id": None,
        "baokuan_detail_id": None,
        "drafts": [],
        "draft_model": None,
        "draft_themes": None,
        "draft_context": {},
        "auto_gen": False,
        "content_source": None,
        # 内容生成（提示词）
        "prompt_product": "",
        "prompt_audience": "",
        "prompt_scenario": "",
        "prompt_selling": "",
        "prompt_direction": "真实体验分享",
        "prompt_extra": "",
        "prompt_result": None,           # build_prompt 输出
        "prompt_ref_title": "",          # 来自爆款的参考标题
        # 内容生成输入控件（与页面稳定对应）
        "content_product": "",
        "content_audience": "",
        "content_scenario": "",
        "content_selling": "",
        "content_direction": "真实体验分享",
        "content_extra": "",
        "content_ref_title": "",
        "last_prompt": None,             # build_prompt 输出
        "last_product": {},              # 生成时的产品信息
        "prompts": [],                   # 已保存的提示词记录
        # 内容创作工作室（PRD §20-§30：标题工厂/成稿/评分/风险/封面/选题）
        "studio": None,                  # studio_pipeline 完整结果
        "studio_product": {},            # 生成时的产品信息
        "studio_title": "",              # 成稿标题（text_input key，用户可编辑）
        "studio_body": "",               # 成稿正文（text_area key，用户可编辑）
        "studio_last_action": None,      # (kind, msg) 操作反馈（普通 key，非控件）
        # 采集设置
        "keyword": "",
        "scope": "近30天",
        "count": 100,
        "max_scrolls": 30,
        "with_details": True,
        "detail_limit": 0,
        "download_images": False,
        "analysis_goals": {g: True for g in GOALS},
        # 首页「快速开始采集」卡（独立 widget key，避免与侧栏 keyword/count Session State 冲突）
        "home_kw": "",
        "home_count": 100,
        # 访客模式：每个访问者用自己的小红书账号（公开访问时互不干扰主账号）
        "visitor_id": uuid.uuid4().hex[:10],
        "visitor_mode": False,
        "visitor_bound": False,
        "qr_running": False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v
    # 单身份模式：不再向 URL 写 ?vid= 访客标记，也不读。
    # （visitor_id 等 key 仅为兼容旧 helper 函数保留，业务上不再使用。）
    # 采集目标勾选键：单独初始化，控件用 key 直接读写，避免 value= 冲突
    for g in GOALS:
        if ("goal_" + g) not in st.session_state:
            st.session_state["goal_" + g] = True
    # AI 创作条件输入框默认值（控件只用 key= 直接读写，避免 value= 冲突）
    # 注意：这些键不能用 value= 传入控件，否则会与 Session State 冲突报错
    ctx_defaults = {
        "ctx_competitor": "",
        "ctx_keyword": "",
        "ctx_user": "年轻女性 / 关注品质的用户",
        "ctx_goal": "新品推广 / 种草",
        "ctx_style": "真实测评",
        "ctx_ref": "",
        "ctx_focus_in": "",
        "ctx_user_focus": "",
        "ctx_opp_in": "",
        "ctx_opp": "",
        "ctx_opp_kw": "",
        "ctx_brand": "",
        "ctx_product": "",
    }
    for k, v in ctx_defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v
    # 新会话（刷新/重启）首次运行：自动恢复上次工作上下文，保证采集数据不丢
    if "ctx_booted" not in st.session_state:
        st.session_state["ctx_booted"] = True
        restore_current_context()


# ======================================================================
# 工具函数
# ======================================================================
def load_tasks():
    if os.path.exists(TASKS_FILE):
        try:
            return json.load(open(TASKS_FILE, encoding="utf-8"))
        except Exception:
            return []
    return []


def save_tasks():
    json.dump(st.session_state.tasks, open(TASKS_FILE, "w", encoding="utf-8"),
              ensure_ascii=False, indent=2)


def add_recent_task(task):
    st.session_state.tasks = [t for t in st.session_state.tasks if t.get("id") != task["id"]]
    st.session_state.tasks.insert(0, task)
    save_tasks()


def check_login():
    """判定当前是否处于「已登录」状态。

    判定依据（任一为真即视为未登录）：
      - cookies.json 缺失或里面没有 web_session；
      - xhs_login.json 缺失（从未通过界面扫码登录过）；
      - xhs_login.json 中的 at 时间戳早于 30 天前（视为过期，需重新扫码）。

    注意：仅看 cookies.json 的 web_session 是不够的，因为 Playwright 持久化时
    一些 cookie（特别是 session cookie）的 expires=-1，无法靠 expires 判断是否还有
    实际登录态；会出现"显示已登录、实际打开页面才发现已掉线"的问题。
    """
    if not os.path.exists(COOKIE_FILE):
        return False
    try:
        cookies = json.load(open(COOKIE_FILE, encoding="utf-8"))
    except Exception:
        return False
    has_ws = any(isinstance(c, dict) and c.get("name") == "web_session" and c.get("value")
                 for c in cookies)
    if not has_ws:
        return False
    if not os.path.exists(LOGIN_FLAG):
        return False
    try:
        flag = json.load(open(LOGIN_FLAG, encoding="utf-8"))
    except Exception:
        return False
    if not flag.get("logged_in"):
        return False
    at = float(flag.get("at") or 0)
    if not at or (time.time() - at) > LOGIN_VALID_SECONDS:
        return False
    return True


def mark_login_success():
    """在扫码登录成功时调用，写入持久化登录标志。"""
    try:
        json.dump({"logged_in": True, "at": time.time(),
                   "ua": "xhs_scraper.login_only"},
                  open(LOGIN_FLAG, "w", encoding="utf-8"),
                  ensure_ascii=False, indent=2)
    except Exception as e:
        print("[mark_login_success] failed:", e)


def clear_login():
    """退出登录：清掉 cookies.json 和 xhs_login.json（下次启动就是未登录）。"""
    for p in (COOKIE_FILE, LOGIN_FLAG):
        if os.path.exists(p):
            try:
                os.remove(p)
            except Exception:
                pass


def normalize_notes(notes):
    recs = []
    for n in notes:
        recs.append({
            "id": n.get("id") or n.get("url") or "",
            "title": n.get("title", "") or "",
            "author": n.get("author", "") or "",
            "url": n.get("url", "") or "",
            "cover": n.get("cover", "") or "",
            "content": n.get("content", "") or "",
            "publish_time": n.get("publish_time", "") or "",
            "likes": int(n.get("likes") or 0),
            "collects": int(n.get("collects") or 0),
            "comments": int(n.get("comments") or 0),
            "shares": int(n.get("shares") or 0),
            "fans": int(n.get("fans") or 0),
            "images": n.get("images") or [],
            "local_images": n.get("local_images") or [],
        })
    return recs


def analyze_notes(notes):
    recs = normalize_notes(notes)
    if not recs:
        return recs, None
    heats = [compute_heat(r) for r in recs]
    for r, h in zip(recs, heats):
        r["analysis"] = analyze_note(r, heats)
    summary = summarize_rules(recs, [r["analysis"] for r in recs])
    return recs, summary


def ensure_analysis():
    if st.session_state.notes and not st.session_state.analysis_summary:
        _, summary = analyze_notes(st.session_state.notes)
        st.session_state.analysis_summary = summary


def make_task(cfg, notes, status):
    tid = "T" + datetime.now().strftime("%Y%m%d%H%M%S")
    return {
        "id": tid,
        "name": f"{cfg['keyword']} 竞品分析",
        "keyword": cfg["keyword"],
        "scope": cfg.get("scope", "近30天"),
        "count": cfg.get("count", len(notes)),
        "goals": cfg.get("goals", []),
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "status": status,
        "note_count": len(notes),
    }


def save_task_data(task, notes):
    os.makedirs(TASK_DIR, exist_ok=True)
    json.dump(notes, open(os.path.join(TASK_DIR, task["id"] + ".json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=2)


def load_task_data(task_id):
    p = os.path.join(TASK_DIR, task_id + ".json")
    if os.path.exists(p):
        try:
            return json.load(open(p, encoding="utf-8"))
        except Exception:
            return None
    return None


def persist_context(task=None, recs=None, register=True):
    """把「当前工作上下文」持久化（任务数据入库 + 快照），让刷新/重启后能无缝续上。

    - 任务数据 → tasks/{task_id}.json：历史记录页「继续分析 / 查看爆款 / 导出」从此读取；
    - 快照 _cur_ctx.json：只记录 task_id，应用启动时 restore_current_context() 自动恢复；
    - register=True 时同步把任务登记/置顶到任务列表（xhs_tasks.json）。
    未指定参数时自动读取 session 里最新的 notes / current_task。
    """
    task = task if task is not None else st.session_state.get("current_task")
    recs = recs if recs is not None else st.session_state.get("notes")
    if task is None or not recs or not task.get("id"):
        return False
    try:
        save_task_data(task, recs)
        if register:
            add_recent_task(task)
        json.dump({"task_id": task["id"],
                   "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
                  open(CUR_CTX_FILE, "w", encoding="utf-8"),
                  ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print("[persist_context] failed:", e)
        return False


def restore_current_context():
    """新会话（浏览器刷新 / 应用重启）首次运行时调用：自动恢复上次工作上下文。

    优先读快照指向的任务数据文件；若任务数据缺失则回退到最近一次采集落盘的
    xhs_notes.json（DATA_FILE）。恢复成功后 session 内 notes / analysis_summary /
    current_task 即就绪，「帖子数据 / 爆款分析 / 内容生成」无需重新采集即可继续。
    """
    try:
        notes, task = None, None
        if os.path.exists(CUR_CTX_FILE):
            snap = json.load(open(CUR_CTX_FILE, encoding="utf-8"))
            tid = snap.get("task_id")
            if tid:
                for t in st.session_state.tasks:
                    if t.get("id") == tid:
                        task = t
                        break
                data = load_task_data(tid)
                if data:
                    notes = data
        if notes is None and os.path.exists(DATA_FILE):
            # 快照任务数据缺失/从未落盘 → 回退最近一次采集的原始数据
            try:
                raw = json.load(open(DATA_FILE, encoding="utf-8"))
                notes = raw if raw else None
            except Exception:
                notes = None
        if not notes:
            return False
        recs, summary = analyze_notes(notes)
        if task is None:
            task = make_task({"keyword": "最近一次采集", "goals": []}, recs, "已分析")
        st.session_state.notes = recs
        st.session_state.analysis_summary = summary
        st.session_state.current_task = task
        st.session_state["ctx_banner"] = (
            f"已自动恢复上次工作上下文：{task['name']} · {len(recs)} 篇（无需重新采集）")
        return True
    except Exception as e:
        print("[restore_current_context] failed:", e)
        return False


def append_progress(path, msg):
    with open(path, "a", encoding="utf-8") as f:
        f.write(msg + "\n")


# ======================================================================
# 真实采集（后台线程 + 文件进度，供界面实时展示 AI Agent 工作流）
# ======================================================================
def start_scrape(cfg):
    for p in (SCRAPE_PROGRESS, SCRAPE_STATUS):
        if os.path.exists(p):
            os.remove(p)
    # 访客模式：使用该访客独立的登录态，浏览器无头运行（不在本机反复弹窗）
    vid = st.session_state.get("visitor_id")
    cookie_file = None
    # 云端无显示器强制无头；本机主账号保持有头（扫码弹窗）；访客模式二维码显示在页面
    headless = _is_cloud_env() or bool(st.session_state.get("visitor_mode"))
    is_visitor_run = bool(st.session_state.get("visitor_mode")
                          and read_visitor_cookie(vid))
    if is_visitor_run and vid:
        cookie_file = session_cookie_path(vid)
        headless = True
    st.session_state.scraping = True
    st.session_state.scrape_cfg = cfg

    def worker():
        try:
            # 云端无头模式：浏览器在服务器后台运行，不弹窗口；
            # 登录二维码会截图显示在页面下方，用手机小红书 App 扫码即可。
            if headless:
                append_progress(SCRAPE_PROGRESS,
                                "云端无头模式：浏览器在服务器后台运行，如需要登录，"
                                "二维码会显示在页面下方，请用手机小红书 App 扫码")
            scraper = XHSScraper(
                headless=headless,
                progress=lambda m: append_progress(SCRAPE_PROGRESS, m),
                cookie_file=cookie_file,
                qr_capture_path=SCRAPE_QR if headless else None,
            )
            notes = scraper.run(
                keyword=cfg["keyword"],
                target_count=cfg["count"],
                max_scrolls=cfg["max_scrolls"],
                with_details=cfg["with_details"],
                detail_limit=cfg["detail_limit"] if cfg["detail_limit"] else None,
                download_images=cfg["download_images"],
            )
            with open(DATA_FILE, "w", encoding="utf-8") as f:
                json.dump(notes, f, ensure_ascii=False, indent=2)
            # 仅主账号（非访客）采集成功才翻转本机全局登录态；
            # 访客用自己的账号，不应污染主账号的登录标记。
            if not is_visitor_run:
                mark_login_success()
            json.dump({"running": False, "done": True, "error": None, "count": len(notes)},
                      open(SCRAPE_STATUS, "w", encoding="utf-8"))
        except Exception as e:
            json.dump({"running": False, "done": True, "error": str(e), "count": 0},
                      open(SCRAPE_STATUS, "w", encoding="utf-8"))

    threading.Thread(target=worker, daemon=True).start()


def show_scrape_progress():
    status = {"running": True, "done": False}
    if os.path.exists(SCRAPE_STATUS):
        try:
            status = json.load(open(SCRAPE_STATUS, encoding="utf-8"))
        except Exception:
            pass
    if status.get("done"):
        err = status.get("error")
        st.session_state.scraping = False
        if err:
            st.error("采集失败：" + str(err))
        else:
            notes = json.load(open(DATA_FILE, encoding="utf-8"))
            recs, summary = analyze_notes(notes)
            st.session_state.notes = recs
            st.session_state.analysis_summary = summary
            cfg = st.session_state.scrape_cfg or {"keyword": "竞品", "goals": []}
            task = make_task(cfg, recs, "待分析")
            st.session_state.current_task = task
            st.session_state.selected_note_id = None
            st.session_state.baokuan_detail_id = None
            # 采集结果落地：任务数据文件 + 任务列表 + 快照（刷新后可自动续上）
            persist_context(task, recs)
            if os.path.exists(SCRAPE_STATUS):
                os.remove(SCRAPE_STATUS)
            if os.path.exists(SCRAPE_QR):
                try:
                    os.remove(SCRAPE_QR)
                except Exception:
                    pass
            st.success(f"采集完成，共 {len(recs)} 条竞品帖子，已自动完成爆款分析")
            # 自动进入下一步：帖子数据
            st.session_state.page = "帖子数据"
            st.rerun()
        return
    lines = []
    if os.path.exists(SCRAPE_PROGRESS):
        lines = open(SCRAPE_PROGRESS, encoding="utf-8").read().splitlines()[-18:]
    st.markdown("### 🤖 AI 采集 Agent 执行中")
    with st.status("正在采集竞品内容…", state="running", expanded=True):
        for ln in lines:
            st.write("· " + ln)
    # 云端无头模式下把登录二维码展示在页面上，用户手机扫码完成登录
    if os.path.exists(SCRAPE_QR):
        try:
            st.markdown("#### 📱 扫码登录小红书（云端无头模式，浏览器不弹窗）")
            st.image(SCRAPE_QR, width=280,
                     caption="用手机小红书 App 扫一扫（二维码约 180 秒有效，过期可点「重新发起」；"
                             "登录成功后本页会自动继续采集）")
        except Exception:
            pass
    time.sleep(0.6)
    st.rerun()


# ======================================================================
# 访客扫码登录（公开访问：每个访问者用自己的小红书账号，登录态相互隔离）
# ======================================================================
def start_visitor_login(vid, relogin=False):
    """发起访客扫码登录：后台线程把小红书登录二维码截图给界面展示，登录态写入 sessions/{vid}.json。

    - 浏览器以 headless 运行（二维码直接显示在工作台页面，本机不弹窗打扰主人）；
    - 与主人 cookies.json 完全隔离；同一会话同时只允许一个扫码流程（qr_running 守卫）；
    - relogin=True 表示「切换账号」：先清除旧登录态文件，强制重新扫码。
    """
    if not vid or st.session_state.get("qr_running") or st.session_state.get("scraping"):
        return
    # 注意：acct_choice 是采集页 radio 控件 key，脚本主体内一律不赋值；
    # 调用方必须先让身份处于 visitor（radio 选择或访客默认），此处只置非控件标志。
    st.session_state["visitor_mode"] = True
    st.session_state["visitor_bound"] = False
    if relogin:
        cp = session_cookie_path(vid)
        if os.path.exists(cp):
            try:
                os.remove(cp)
            except Exception:
                pass
    for f in (visitor_status_file(vid), visitor_progress_file(vid), session_qr_path(vid)):
        if os.path.exists(f):
            try:
                os.remove(f)
            except Exception:
                pass
    st.session_state.qr_running = True

    def worker():
        try:
            scraper = XHSScraper(
                headless=True,
                login_timeout=180,
                progress=lambda m: append_progress(visitor_progress_file(vid), m),
                cookie_file=session_cookie_path(vid),
            )
            scraper.login_with_qr_capture(session_qr_path(vid))
            json.dump({"running": False, "done": True, "error": None},
                      open(visitor_status_file(vid), "w", encoding="utf-8"))
        except Exception as e:
            json.dump({"running": False, "done": True, "error": str(e)},
                      open(visitor_status_file(vid), "w", encoding="utf-8"))

    threading.Thread(target=worker, daemon=True).start()


def show_visitor_login_progress(vid):
    """访客扫码进行中：展示二维码图片 + 进度行，完成后刷新状态。"""
    status = {"running": True}
    sp = visitor_status_file(vid)
    if os.path.exists(sp):
        try:
            status = json.load(open(sp, encoding="utf-8"))
        except Exception:
            pass
    if status.get("done"):
        st.session_state.qr_running = False
        err = status.get("error")
        for f in (sp, visitor_progress_file(vid)):
            if os.path.exists(f):
                try:
                    os.remove(f)
                except Exception:
                    pass
        if err:
            st.warning("⚠️ 访客登录未完成：" + str(err) + "（可点上方「🔑 扫码登录」重新发起）")
        else:
            st.session_state.visitor_bound = bool(read_visitor_cookie(vid))
            st.success("✅ 访客账号登录成功！登录态已独立保存，点「🚀 开始批量采集」即使用你自己的账号。")
        return
    lines = []
    pp = visitor_progress_file(vid)
    if os.path.exists(pp):
        lines = open(pp, encoding="utf-8").read().splitlines()[-12:]
    st.markdown("### 🔑 访客扫码登录进行中")
    with st.status("请在下方二维码上用手机小红书 App 扫码…", state="running", expanded=True):
        for ln in lines:
            st.write("· " + ln)
    qp = session_qr_path(vid)
    if os.path.exists(qp):
        try:
            st.image(qp, caption="用手机小红书 App 扫一扫（二维码约 180 秒有效，过期可重新发起）",
                     width=280)
        except Exception:
            pass
    time.sleep(1.2)
    st.rerun()


# ======================================================================
# 批量下载高清原图（后台线程 + 文件进度，复用已保存的登录态）
# ======================================================================
def start_image_download(ids):
    """对给定笔记 id 列表批量下载高清无水印原图。"""
    for p in (IMG_PROGRESS, IMG_STATUS):
        if os.path.exists(p):
            os.remove(p)
    st.session_state.downloading_images = True
    payload = [dict(r) for r in st.session_state.notes if r["id"] in set(ids)]

    def worker():
        try:
            scraper = XHSScraper(headless=_is_cloud_env(),
                                 progress=lambda m: append_progress(IMG_PROGRESS, m))
            notes, n_img = scraper.download_images_for(payload)
            # 结果写文件，由界面线程合并回 session（线程间不共享 st.session_state）
            json.dump(notes, open(IMG_RESULT, "w", encoding="utf-8"), ensure_ascii=False)
            json.dump({"running": False, "done": True, "error": None,
                       "count": len(notes), "images": n_img},
                      open(IMG_STATUS, "w", encoding="utf-8"))
        except Exception as e:
            json.dump({"running": False, "done": True, "error": str(e),
                       "count": 0, "images": 0},
                      open(IMG_STATUS, "w", encoding="utf-8"))

    threading.Thread(target=worker, daemon=True).start()


def show_image_progress():
    status = {"running": True, "done": False}
    if os.path.exists(IMG_STATUS):
        try:
            status = json.load(open(IMG_STATUS, encoding="utf-8"))
        except Exception:
            pass
    if status.get("done"):
        err = status.get("error")
        st.session_state.downloading_images = False
        if err:
            st.error("图片下载失败：" + str(err))
        else:
            n_img = status.get("images", 0)
            # 合并本地图片路径回当前数据，并保存到数据文件
            if os.path.exists(IMG_RESULT):
                try:
                    updated = json.load(open(IMG_RESULT, encoding="utf-8"))
                    by_id = {n["id"]: n for n in updated}
                    for r in st.session_state.notes:
                        u = by_id.get(r["id"])
                        if u:
                            r["images"] = u.get("images") or r.get("images") or []
                            r["local_images"] = u.get("local_images") or r.get("local_images") or []
                    if os.path.exists(DATA_FILE):
                        try:
                            all_notes = json.load(open(DATA_FILE, encoding="utf-8"))
                            fb = {n["id"]: n for n in all_notes}
                            for u in updated:
                                if u["id"] in fb:
                                    fb[u["id"]]["local_images"] = u.get("local_images") or []
                                    fb[u["id"]]["images"] = u.get("images") or []
                            json.dump(all_notes, open(DATA_FILE, "w", encoding="utf-8"),
                                      ensure_ascii=False, indent=2)
                        except Exception:
                            pass
                    # 本地图片路径已合并回当前数据 → 同步刷新任务数据与快照
                    persist_context(register=False)
                except Exception:
                    pass
                os.remove(IMG_RESULT)
            if os.path.exists(IMG_STATUS):
                os.remove(IMG_STATUS)
            st.success(f"图片下载完成，共新增 {n_img} 张高清原图（保存在 downloads/ 目录，按笔记 id 分文件夹）")
            st.caption("回到「采集数据」点击任意笔记即可在详情面板查看本地图片。")
        return
    lines = []
    if os.path.exists(IMG_PROGRESS):
        lines = open(IMG_PROGRESS, encoding="utf-8").read().splitlines()[-15:]
    st.markdown("### 🖼️ 批量下载原图执行中")
    with st.status("正在通过登录态批量下载高清无水印原图…", state="running", expanded=True):
        for ln in lines:
            st.write("· " + ln)
    time.sleep(0.6)
    st.rerun()


# ======================================================================
# 采集参数：侧栏为唯一设置入口，首页/各处复用同一份配置
# ======================================================================
def current_cfg():
    """读取当前采集参数；首页「快速开始采集」输入（home_kw / home_count）优先于侧栏设置。
    注意：不要在脚本主体或按钮分支里给 keyword / count 这些 widget key 赋值
    （Streamlit 会在 widget 实例化后写入时报 Session State 冲突）；统一在这里做优先级合并。
    """
    return {
        "keyword": (st.session_state.get("home_kw") or "").strip()
                  or (st.session_state.get("keyword") or "").strip(),
        "scope": st.session_state.scope,
        "count": st.session_state.get("home_count")
                  or st.session_state.get("count") or 100,
        "max_scrolls": st.session_state.max_scrolls,
        "with_details": st.session_state.with_details,
        "detail_limit": st.session_state.detail_limit,
        "download_images": st.session_state.download_images,
        "goals": [g for g in GOALS if st.session_state.get("goal_" + g, True)],
    }


def start_scrape_from_settings():
    """按当前采集配置启动真实采集。

    返回值：
      - True                已启动采集线程；
      - False               关键词为空（由调用方提示）；
      - "need_visitor_login" 当前是访客身份但尚未扫码绑定（调用方应引导访客扫码）。
    关键词/数量由 current_cfg() 内部合并首页输入与侧栏设置。
    """
    if not current_cfg()["keyword"]:
        return False
    ident = account_identity()
    if ident["choice"] == "visitor":
        # 访客身份：登录态必须来自该访客自己的隔离文件；未绑定则不启动（防误用主人 cookies）
        st.session_state["visitor_mode"] = True
        if not ident["visitor_bound"]:
            return "need_visitor_login"
    else:
        st.session_state["visitor_mode"] = False
    start_scrape(current_cfg())
    return True


def chips_html(items):
    return "".join(f'<span class="tag">{i}</span>' for i in items if i)


# ======================================================================
# 登录已并入采集流程：采集时 XHSScraper 会打开小红书页面，
# 若需要登录会停留在扫码页等待（浏览器持续可见），扫码后继续采集。
# 采集 worker 成功后会调用 mark_login_success() 写登录标记，顶栏状态随之翻转。
# 不再提供独立的「打开浏览器扫码登录」按钮（Windows 下 Playwright 子进程不稳定）。
# ======================================================================


# ======================================================================
# 洞察 / 机会（全部基于真实采集数据计算）
# ======================================================================
def get_insights(recs, summary):
    if not recs or not summary:
        return {}
    top = [r for r in recs if r["analysis"]["baokuan"] in ("爆款", "潜力爆款")]
    # 竞品主要内容方向
    directions = []
    for it in summary["hook_stats"][:4]:
        directions.append(f"{it[0]}类内容（{it[1]}篇，平均赞{it[2]}）")
    for it in summary["cover_stats"][:3]:
        directions.append(f"{it[0]}封面（{it[1]}篇）")
    # 高频关键词
    freq_kw = [{"kw": k, "cnt": c} for k, c in summary["keyword_top"][:12]]
    # 高互动关键词：按平均(赞+藏)排序
    eng = {}
    for r in recs:
        try:
            kws = []
            for v in r["analysis"]["keywords"].values():
                kws += v
        except Exception:
            kws = []
        e = (r["likes"] + r["collects"]) or 1
        for k in set(kws):
            eng.setdefault(k, []).append(e)
    high = sorted(({"kw": k, "avg": int(sum(v) / len(v))} for k, v in eng.items()),
                  key=lambda x: -x["avg"])[:10]
    # 用户主要关注点：卖点词频（高表现作品）
    focus_counter = {}
    for r in top:
        for sp in r["analysis"]["selling_points"]:
            focus_counter[sp] = focus_counter.get(sp, 0) + 1
    focus = sorted(focus_counter.items(), key=lambda x: -x[1])[:8]
    # 热门内容类型
    hot_types = [{"t": it[0], "cnt": it[1]} for it in summary["cover_stats"][:6]]
    # 爆款共性
    common = summary["features"]
    # 竞品卖点
    selling = [{"p": k, "cnt": c} for k, c in summary["selling_top"][:10]]
    return {
        "directions": directions,
        "freq_kw": freq_kw,
        "high_kw": high,
        "focus": focus,
        "hot_types": hot_types,
        "common": common,
        "selling": selling,
    }


def get_opportunities(recs, summary):
    if not recs or not summary:
        return []
    top = [r for r in recs if r["analysis"]["baokuan"] in ("爆款", "潜力爆款")]
    total = max(len(recs), 1)
    opps = []
    # 真实测评：人物使用场景 / 亲测 钩子占比
    real_cnt = sum(1 for r in top if r["analysis"]["cover"] in ("人物使用场景", "生活方式场景")
                   or r["analysis"]["hook_type"] in ("经验分享型", "种草型"))
    if real_cnt:
        opps.append({"title": "真实测评内容", "kw": "测评",
                     "desc": "高表现笔记多为真人亲测 / 使用场景，用户更信真实体验。",
                     "evidence": f"爆款中约 {real_cnt}/{max(len(top),1)} 篇为真实测评 / 场景类"})
    # 场景化内容
    scene_kw = ["客厅", "卧室", "出租屋", "小户型", "婚房", "书房", "阳台"]
    scene_cnt = sum(1 for r in recs if any(k in (r["title"] + r["content"]) for k in scene_kw))
    if scene_cnt:
        opps.append({"title": "场景化种草", "kw": "客厅",
                     "desc": "围绕具体生活场景（客厅/卧室/出租屋）展开，收藏表现更好。",
                     "evidence": f"含场景词的笔记 {scene_cnt} 篇"})
    # 高性价比
    cheap_cnt = sum(1 for r in recs if any(k in (r["title"] + r["content"])
                      for k in ["平价", "性价比", "便宜", "平替", "划算", "学生党"]))
    if cheap_cnt:
        opps.append({"title": "高性价比对比", "kw": "性价比",
                     "desc": "用户对价格敏感，做「平替 / 性价比」对比内容有流量。",
                     "evidence": f"提及性价比/平替的笔记 {cheap_cnt} 篇"})
    # 品牌对比
    vs_cnt = sum(1 for r in recs if "vs" in (r["title"] + r["content"]).lower()
                 or "对比" in (r["title"] + r["content"]))
    if vs_cnt:
        opps.append({"title": "品牌横向对比", "kw": "对比",
                     "desc": "竞品间横向对比内容互动高，适合做选购决策类。",
                     "evidence": f"含对比/VS 的笔记 {vs_cnt} 篇"})
    # 内容缺口：高互动但低频的关键词方向
    if summary["keyword_top"]:
        kw, c = summary["keyword_top"][0]
        opps.append({"title": f"聚焦「{kw}」方向", "kw": kw,
                     "desc": f"「{kw}」是高频且高表现关键词，可围绕它做系列内容。",
                     "evidence": f"「{kw}」出现 {c} 次，集中于高表现笔记"})
    if not opps:
        opps.append({"title": "真实体验种草", "kw": "真实",
                     "desc": "用真实使用体验切入，比纯产品展示更易种草。", "evidence": "默认机会"})
    return opps[:5]


# ======================================================================
# AI 创作：生成 / 变换 / 质量评分
# ======================================================================
PROMO = ["点击下方", "左下角", "链接在这", "快冲", "赶紧买", "限时", "促销", "全网最低",
         "厂家直销", "私信", "加微信", "一件也是批发价", "购买链接", "优惠", "错过",
         "再不买", "点链接", "链接见", "领券", "秒杀"]


def build_context():
    recs = st.session_state.notes
    kw = st.session_state.get("ctx_keyword", st.session_state.keyword)
    focus = []
    if recs:
        heats = [compute_heat(r) for r in recs]
        top = [r for r, h in zip(recs, heats) if r["analysis"]["baokuan"] in ("爆款", "潜力爆款")] or recs
        for r in top[:8]:
            focus += r["analysis"]["selling_points"]
    # 手动填写的「用户关注点」与数据归纳的关注点合并
    manual = [s.strip() for s in re.split(r"[、,，/]",
              st.session_state.get("ctx_focus_in", "") or "") if s.strip()]
    focus += manual
    focus = list(dict.fromkeys(focus))[:6]
    return {
        "competitor": st.session_state.get("ctx_competitor", st.session_state.keyword),
        "keyword": kw,
        "goal_user": st.session_state.get("ctx_user", "年轻女性 / 关注品质的用户"),
        "goal_content": st.session_state.get("ctx_goal", "新品推广 / 种草"),
        "style": st.session_state.get("ctx_style", "真实测评"),
        "ref_baokuan": st.session_state.get("ctx_ref", ""),
        "user_focus": "、".join(focus),
        "content_opportunity": st.session_state.get("ctx_opp", ""),
        "opportunity_kw": st.session_state.get("ctx_opp_kw", ""),
        "focus_kws": focus,
        "brand": st.session_state.get("ctx_brand", ""),
        "product": st.session_state.get("ctx_product", ""),
    }


def standalone_model(context):
    """无采集数据时的独立创作模型：由关键词/关注点直接构建。"""
    p = (context.get("product") or "").strip() or (context.get("keyword") or "").strip() or "好物"
    style = (context.get("style") or "").strip()
    selling = [s for s in (context.get("focus_kws") or []) if s] or ["颜值", "实用", "质感"]
    return {
        "cover_top": ["场景展示型", "人物场景型", "细节特写型"],
        "hook_top": ["种草型", "痛点共鸣型", "干货攻略型"],
        "topics": [f"{p}怎么选", f"{p}真实体验分享", f"{p}避坑指南", f"{style or '心动'}{p}推荐"],
        "title_formulas": [
            "终于找到适合【场景】的【产品】了",
            "被问疯了的【风格】【产品】，谁懂啊",
            "别再乱买了！【产品】看这篇就够了",
        ],
        "title_examples": [],
        "prod": [p],
        "style": [style] if style else [],
        "scene": [],
        "need": [],
        "tags": [p] + ([style] if style else []),
        "selling_top": selling,
        "features": [],
    }


def standalone_themes(context, n=3):
    """无采集数据时的合成选题（gen_one 只读这些字段）。"""
    p = (context.get("product") or "").strip() or (context.get("keyword") or "").strip() or "好物"
    style = (context.get("style") or "").strip()
    selling = [s for s in (context.get("focus_kws") or []) if s] or ["颜值", "舒适", "性价比"]
    hooks = ["种草型", "痛点共鸣型", "干货攻略型"]
    covers = ["场景展示型", "人物场景型", "细节特写型"]
    scenes = ["日常", "通勤", "户外"]
    themes = []
    for i in range(n):
        themes.append({
            "analysis": {
                "keywords": {"产品关键词": [p], "风格关键词": [style] if style else [],
                              "场景关键词": [scenes[i % len(scenes)]],
                              "用户需求关键词": []},
                "selling_points": selling,
                "hook_type": hooks[i % len(hooks)],
                "cover": covers[i % len(covers)],
                "heat": 100,
                "baokuan": "潜力爆款",
            }
        })
    return themes


def generate_drafts(context):
    recs = st.session_state.notes
    bp = load_brand(context["brand"]) if context.get("brand") else None
    if not recs:
        # 独立创作模式：无需采集数据，用关键词/风格/关注点直接生成
        model = standalone_model(context)
        themes = standalone_themes(context, 3)
        notes = [gen_one(t, model, bp, context.get("brand", ""),
                         (context.get("product") or "").strip() or None, idx=i)
                 for i, t in enumerate(themes)]
        st.session_state.content_source = "independent"
        return notes, model, themes
    st.session_state.content_source = "data"
    pool = recs
    kw = context.get("opportunity_kw")
    if kw:
        matched = [r for r in recs if kw in (r["title"] + r["content"])]
        if matched:
            pool = matched
    notes = gen_notes(pool, n=3, bp=bp, brand_name=context.get("brand", ""),
                      product=context.get("product"))
    model = build_content_model(pool)
    return notes, model, pool


def reduce_ad(text):
    t = text
    for p in PROMO:
        t = t.replace(p, "")
    t = re.sub(r"！+", "！", t)
    t = re.sub(r"!+", "!", t)
    t = re.sub(r"链接在左下角.*", "", t)
    t = re.sub(r"左下角有链接.*", "", t)
    return t.strip()


def enhance_real(text):
    insert = "说点真实体验：我自己用了一段时间，不是广子，纯分享真实感受。"
    parts = text.split("\n\n", 1)
    if len(parts) > 1:
        return parts[0] + "\n\n" + insert + "\n\n" + parts[1]
    return text + "\n\n" + insert


def optimize_title(title):
    cands = [
        ("终于找到！" + title) if not title.startswith("终于") else title,
        title + "（亲测）",
        ("被问爆的" + title) if len(title) < 14 else title,
        title + "｜真实体验分享",
    ]
    for c in cands:
        if c != title:
            return c
    return title + "｜推荐"


def compute_quality(title, body, ctx):
    text = (title or "") + "\n" + (body or "")
    kws = [ctx.get("keyword")] + (ctx.get("focus_kws") or [])
    kws = [k for k in kws if k]
    cov = (sum(1 for k in kws if k in text) / len(kws)) if kws else 0.85
    keyword_cov = int(min(1, cov) * 100)
    need_words = ["我", "真实", "亲测", "体验", "用了", "感受", "日常", "场景", "适合", "自己"]
    user_rel = int(min(1, sum(1 for w in need_words if w in body) / 6) * 100)
    ad_hits = sum(1 for p in PROMO if p in body)
    excl = body.count("！") + body.count("!")
    ad_score = max(0, 100 - ad_hits * 14 - max(0, excl - 3) * 4)
    length = len(body)
    if 300 <= length <= 1000:
        plat = 90
    elif length < 300:
        plat = int(50 + length / 8)
    else:
        plat = max(60, 100 - (length - 1000) // 60)
    if "#" in body:
        plat = min(100, plat + 6)
    completeness = 55
    if length > 200:
        completeness += 15
    if "。" in body or "\n" in body:
        completeness += 10
    if any(w in body for w in ["总结", "所以", "最后", "建议", "入手", "推荐", "种草"]):
        completeness += 15
    completeness = min(100, completeness)
    return {"用户相关性": user_rel, "关键词覆盖": keyword_cov, "平台适配度": plat,
            "广告感": ad_score, "内容完整度": completeness}


def export_drafts_xlsx(drafts_values, out):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("种草笔记")
    headers = ["篇号", "标题", "正文", "标签", "参考钩子", "核心卖点"]
    ws.append(headers)
    fill = PatternFill("solid", fgColor="C00000")
    hf = Font(bold=True, color="FFFFFF")
    for i, d in enumerate(drafts_values, 1):
        ws.append([i, d["title"], d["body"], d["tags"], d.get("hook", ""), d.get("selling", "")])
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = fill
        cell.font = hf
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for r in range(2, len(drafts_values) + 2):
        ws.cell(row=r, column=3).alignment = Alignment(wrap_text=True, vertical="top")
    for i, w in enumerate([5, 30, 70, 30, 12, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    wb.save(out)
    return out


# ======================================================================
# 顶栏
# ======================================================================
def render_topbar():
    page_name = st.session_state.page
    page_title = PAGE_TITLES.get(page_name, page_name)

    # 第 1 行：产品名 / 副标题（去掉"新建任务"按钮——按用户要求：流程极简化，去除冗余入口）
    st.markdown("### 📕 AI 小红书内容增长工作台")
    st.caption(f"从竞品采集 → 爆款分析 → 小红书内容提示词 · 当前页面：**{page_title}**")

    # 第 2 行：AI 状态 + 小红书登录状态（单账号：本机 cookies.json）
    ident = account_identity()
    s1, s2 = st.columns([1.2, 5])
    with s1:
        st.markdown('<span class="tag">● AI 分析可用</span>', unsafe_allow_html=True)
    with s2:
        if ident["main_logged"]:
            st.markdown('<span class="tag" style="background:#fff5f8;color:#c00;">● 小红书已登录</span>',
                        unsafe_allow_html=True)
            st.caption("&nbsp;&nbsp;登录态随采集自动更新，无需重复扫码", unsafe_allow_html=False)
        else:
            st.markdown('<span class="tag" style="background:#fff0f0;color:#c00;border-color:#ffd6d6;">'
                        '● 小红书未登录（开始采集时会自动弹出扫码）</span>',
                        unsafe_allow_html=True)
            st.caption("&nbsp;&nbsp;尚未登录过；采集时会自动打开小红书，扫码一次即可",
                       unsafe_allow_html=False)

    st.divider()


# ======================================================================
# 侧栏：导航 + 采集设置（精简版——只保留最常用项，详细配置到采集页）
# ======================================================================
def render_sidebar():
    with st.sidebar:
        st.markdown("#### 🧭 导航")
        for name in NAV:
            active = st.session_state.page == name
            if st.button(("▸ " if not active else "● ") + name,
                         key="nav_" + name, use_container_width=True,
                         type="primary" if active else "secondary"):
                st.session_state.page = name
                st.rerun()
        st.divider()

        st.markdown("#### ⚙️ 常用设置")
        # 关键词（首页/采集页都共用）
        st.text_input("搜索关键词", key="keyword",
                      placeholder="如：小户型沙发")
        quick = ["喜茶", "奈雪", "防晒霜", "女装", "护肤品", "小户型沙发"]

        def set_quick_kw(q):
            st.session_state.keyword = q
        # 一行快捷词按钮
        qcols = st.columns(3)
        for i, q in enumerate(quick):
            qcols[i % 3].button(q, key="quick_side_" + q,
                                on_click=set_quick_kw, args=(q,),
                                use_container_width=True)
        st.caption("快捷词会回填到左侧「搜索关键词」框")

        st.selectbox("采集数量", [30, 50, 100, 200, 500], key="count")
        st.caption("更多配置（采集详情 / 下载原图 / 分析目标）请到「小红书采集」页。")


# ======================================================================
# 页面：首页（产品入口 + 4 节点工作流 + 最近任务）
# ======================================================================
def page_home():
    ident = account_identity()
    logged = ident["main_logged"]
    has_data = bool(st.session_state.notes)
    note_count = len(st.session_state.notes) if has_data else 0
    has_analysis = bool(st.session_state.analysis_summary)
    has_prompts = bool(st.session_state.get("prompts"))

    # 刷新/重启后自动恢复的提示（仅展示一次）
    banner = st.session_state.pop("ctx_banner", None)
    if banner:
        st.success(banner, icon="💾")

    # ===== Hero =====
    st.markdown("""
<div class="hero">
  <div class="hero-title">📕 AI 小红书内容增长工作台</div>
  <div class="hero-sub">从竞品采集 → 爆款分析 → 小红书内容提示词 · 把流量变成可复用结构</div>
</div>
""", unsafe_allow_html=True)

    # ===== 快速开始采集卡（极简一步：关键词+数量+开始；登录自动并入采集过程）=====
    # 注意：home_kw / home_count 是独立 widget key，**不与侧栏 keyword / count 双向同步**
    # （侧栏 input 已先于本函数实例化，对它赋值会触发 StreamlitAPIException）。
    # 由 current_cfg() 内部按 home_kw → keyword / home_count → count 合并优先级。
    with st.container(border=True):
        st.markdown("##### 🚀 快速开始采集")
        st.caption("填好下面两项，点「开始采集」即可。第一次采集时如需登录，会自动停留到扫码完成。")
        f1, f2, f3 = st.columns([3, 1.2, 1.6])
        with f1:
            st.text_input("搜索关键词", key="home_kw",
                          placeholder="如：小户型沙发 / 防晒霜",
                          label_visibility="collapsed")
        with f2:
            st.selectbox("采集数量", [30, 50, 100, 200, 500], key="home_count",
                         label_visibility="collapsed")
        with f3:
            started = st.button("🚀 开始采集竞品内容", type="primary",
                                use_container_width=True, key="home_start")
        # 快捷词：on_click 里同步到侧栏 keyword 是安全的（回调在下一次脚本重跑前运行）
        st.caption("快捷词：")
        quick = ["喜茶", "奈雪", "防晒霜", "女装", "护肤品", "小户型沙发"]

        def set_home_kw(q):
            st.session_state.keyword = q
            st.session_state.home_kw = q
        bcols = st.columns(6)
        for i, q in enumerate(quick):
            bcols[i].button(q, key="home_quick_" + q,
                            on_click=set_home_kw, args=(q,),
                            use_container_width=True)

        if started:
            kw = (st.session_state.get("home_kw") or "").strip()
            if not kw:
                st.warning("请先填写搜索关键词")
            else:
                # 不在按钮分支里给 sidebar widget key 赋值（避免 Session State 冲突）；
                # start_scrape_from_settings() 内部由 current_cfg() 合 home_kw 优先级读取。
                start_scrape_from_settings()
                st.rerun()

    st.divider()

    # ===== 4 节点工作流（全部可点跳转）=====
    st.markdown('<p class="sec-title">小红书内容增长流程</p>', unsafe_allow_html=True)
    node_labels = [
        ("01 小红书采集", "打开浏览器 · 批量采集 · 自动处理登录", "📕", "小红书采集"),
        ("02 帖子数据", "查看采集结果 · 筛选/排序 · 下载 Excel", "📊", "帖子数据"),
        ("03 爆款分析", "生成爆款分析表 · 提取爆款规律 / 痛点 / 卖点", "🔥", "爆款分析"),
        ("04 内容提示词", "输入产品 → 生成小红书创作提示词", "✨", "内容生成"),
    ]
    cols = st.columns(4, gap="medium")
    for i, (title_text, sub, icon, target) in enumerate(node_labels):
        with cols[i]:
            done = (i == 0 and logged) or (i == 1 and has_data) or (i == 2 and has_analysis) or (i == 3 and has_prompts)
            tag = "✅ 已完成" if done else "○ 待进行"
            done_cls = " done" if done else ""
            st.markdown(f"""
<div class="flow-node{done_cls}">
  <div class="flow-icon">{icon}</div>
  <div class="flow-title">{title_text}</div>
  <div class="flow-sub">{sub}</div>
  <div class="flow-tag{done_cls}">{tag}</div>
</div>
""", unsafe_allow_html=True)
            if st.button("进入该节点", key=f"home_node_{i}", use_container_width=True):
                st.session_state.page = target
                st.rerun()

    st.divider()

    # ===== 三个快速入口 =====
    st.markdown('<p class="sec-title">快速入口</p>', unsafe_allow_html=True)
    qc = st.columns(3, gap="medium")
    # 顺序：开始采集 / 爆款分析 / 生成提示词
    with qc[0]:
        st.markdown("""
<div class="qc-card">
  <h3>📕 小红书采集</h3>
  <div class="qc-cap">完整流程：自动登录 / 高级配置 / 实时进度 / 下载原图。第一次用必看。</div>
</div>
""", unsafe_allow_html=True)
        if st.button("去采集页", key="home_qc_collect", use_container_width=True):
            st.session_state.page = "小红书采集"
            st.rerun()
    with qc[1]:
        st.markdown("""
<div class="qc-card">
  <h3>🔥 爆款分析</h3>
  <div class="qc-cap">对已采集帖子生成爆款分析表，提取标题/内容/痛点/卖点规律。</div>
</div>
""", unsafe_allow_html=True)
        if st.button("去爆款分析", key="home_qc_bao", use_container_width=True,
                     disabled=not has_data):
            ensure_analysis()
            st.session_state.page = "爆款分析"
            st.rerun()
    with qc[2]:
        st.markdown("""
<div class="qc-card">
  <h3>✨ 生成提示词</h3>
  <div class="qc-cap">输入产品信息，一键生成小红书内容提示词（含角色 / 结构 / 风格 / 输出格式）。</div>
</div>
""", unsafe_allow_html=True)
        if st.button("去生成提示词", key="home_qc_prompt", use_container_width=True):
            st.session_state.page = "内容生成"
            st.rerun()

    st.divider()

    # ===== 最近任务（真实任务优先，无则给示例）=====
    st.markdown('<p class="sec-title">最近任务</p>', unsafe_allow_html=True)
    real_tasks = st.session_state.tasks[:5] if st.session_state.tasks else []
    if real_tasks:
        for t in real_tasks:
            with st.container(border=True):
                r1, r2, r3 = st.columns([2.5, 1.4, 1])
                r1.write(f"**{t['name']}**")
                r2.caption(f"关键词：{t['keyword']} · {t.get('created_at','')[:10]}")
                r3.caption(f"状态：{t.get('status','—')}")
                b1, b2 = st.columns([1, 5])
                if b1.button("继续分析", key="home_rt_go_" + t["id"]):
                    data = load_task_data(t["id"])
                    if data is None:
                        st.warning("任务数据缺失，请重新采集")
                    else:
                        recs, summary = analyze_notes(data)
                        st.session_state.notes = recs
                        st.session_state.analysis_summary = summary
                        st.session_state.current_task = t
                        # 续上历史任务时同步刷新任务数据与快照（下次刷新仍可恢复）
                        persist_context(t, recs, register=False)
                        st.session_state.page = "帖子数据"
                        st.rerun()
    else:
        # 少量示例帮助理解（不会写入实际任务数据）
        examples = [
            ("小户型沙发竞品分析", "沙发 / 小户型", "爆款分析"),
            ("护肤品竞品研究", "护肤品 / 美白", "帖子采集"),
            ("通勤穿搭灵感", "通勤穿搭", "提示词已生成"),
        ]
        for i, (name, kw, stage) in enumerate(examples):
            with st.container(border=True):
                r1, r2, r3 = st.columns([2.5, 1.4, 1])
                r1.write(f"**{name}**")
                r2.caption(f"关键词：{kw}")
                r3.caption(f"当前阶段：{stage}")
                st.caption("（示例）开始真实采集后，这里会自动出现你的真实任务。")

    # ===== 当前进度总览（一行就够了，不做数据大屏）=====
    st.divider()
    st.markdown('<p class="sec-title">当前进度</p>', unsafe_allow_html=True)
    pc = st.columns(4)
    pc[0].metric("小红书", "已登录" if logged else "未登录")
    pc[1].metric("已采集", f"{note_count} 篇")
    pc[2].metric("爆款分析", "已完成" if has_analysis else "未开始")
    pc[3].metric("提示词", f"{len(st.session_state.get('prompts', []))} 个")

    # 未登录提示
    if not logged:
        st.info("💡 采集需要登录小红书账号 → 在上方「快速开始采集」填好关键词，点"
                "「🚀 开始采集竞品内容」，采集会打开小红书页面，扫码一次后登录态自动保存。", icon="🔐")
    elif not has_data:
        st.info("下一步：在上方「快速开始采集」填关键词 → 点「🚀 开始采集竞品内容」→ 自动跳到「帖子数据」。"
                "也可跳过采集，到「设置」上传已有 Excel。")


# ======================================================================
# 页面：小红书采集（登录状态 + 扫码 + 配置 + 实时进度）
# ======================================================================
def page_xhs_collect():
    st.markdown('<p class="sec-title">📕 小红书采集</p>', unsafe_allow_html=True)
    st.caption("① 确认登录状态 → ② 填关键词 → 点「开始批量采集」（未登录会自动弹扫码）。采集过程实时显示在下方进度区。")

    ident = account_identity()          # 单身份模式：恒为 main（cookies.json）
    logged = ident["main_logged"]

    # ----- 1. 登录状态（单账号：本机 cookies.json；未登录时点开始采集自动扫码）-----
    st.markdown("**① 采集账号登录状态**")
    with st.container(border=True):
        if logged:
            st.success("● 小红书账号已登录", icon="✅")
            st.caption("登录态保存在本机 cookies.json。采集自动复用，无需重复扫码。")
        else:
            st.warning("● 小红书尚未登录", icon="⚠️")
            st.caption("点下方「开始批量采集」，浏览器会自动打开小红书登录页，扫码一次后登录态即保存。")

    st.divider()

    # ----- 2. 采集配置（直接复用侧栏的全局配置，并允许这里再次检查）-----
    with st.container(border=True):
        st.markdown("**② 采集配置**")
        cfg = current_cfg()
        kw = cfg["keyword"]
        if not kw:
            st.warning("⚠️ 还没有设置关键词 → 在左侧「常用设置」填写，或回到首页「快速开始采集」卡里填。",
                       icon="⚠️")
        else:
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("关键词", kw)
            c2.metric("范围", cfg["scope"])
            c3.metric("数量", cfg["count"])
            c4.metric("详情", "含详情+原图" if cfg["with_details"] and cfg["download_images"]
                      else ("含详情" if cfg["with_details"] else "仅列表"))
            st.caption("高级设置里可调整：最大采集页数 / 详情上限 / 下载高清原图。")

        busy = bool(st.session_state.scraping)
        c5, c6 = st.columns([3, 1])
        with c5:
            if st.button("🚀 开始批量采集", type="primary",
                         use_container_width=True, key="page_collect_start",
                         disabled=busy):
                if not kw:
                    st.warning("请先在左侧「常用设置」填写关键词")
                else:
                    start_scrape_from_settings()
                    st.rerun()
        with c6:
            if st.session_state.scraping:
                st.info("采集进行中…", icon="⏳")
            elif st.session_state.notes:
                st.success(f"已采集 {len(st.session_state.notes)} 篇", icon="✅")

    # ----- 3. 采集过程（实时进度）-----
    st.divider()
    st.markdown('<p class="sec-title">3. 采集过程</p>', unsafe_allow_html=True)

    # 阶段勾选（与真实流程一一对应，基于真实采集进度文件判断）
    p1, p2, p3, p4, p5, p6 = st.columns(6)
    prog_file = SCRAPE_PROGRESS
    prog = open(prog_file, encoding="utf-8").read().strip() if os.path.exists(prog_file) else ""

    def mark(col, label, done):
        col.markdown("✅ " + label if done else "○ " + label)

    is_running = bool(st.session_state.scraping)
    login_ok = logged
    mark(p1, "打开浏览器", bool(prog) or is_running)
    mark(p2, "进入小红书", bool(prog) or is_running)
    mark(p3, "检查登录", login_ok or "登录" in prog or is_running)
    mark(p4, "搜索关键词", "滑页" in prog or "采集" in prog or "搜索" in prog)
    mark(p5, "采集帖子", "滑页" in prog or "详情" in prog or "采集" in prog)
    mark(p6, "保存数据", "采集完成" in prog or "已下载" in prog)

    if prog:
        st.code(prog, language="text")

    if is_running:
        st.info("正在采集…结果会自动同步到「帖子数据」页。", icon="⏳")

    st.divider()

    # ----- 4. 采集完成快捷入口（只在采集后显示）-----
    if st.session_state.notes:
        st.markdown('<p class="sec-title">4. 下一步</p>', unsafe_allow_html=True)
        b1, b2, b3, b4 = st.columns(4)
        with b1:
            if st.button("📊 查看帖子数据", key="after_posts", use_container_width=True):
                st.session_state.page = "帖子数据"
                st.rerun()
        with b2:
            if st.button("🔥 生成爆款分析表", key="after_bao", use_container_width=True):
                ensure_analysis()
                st.session_state.page = "爆款分析"
                st.rerun()
        with b3:
            if st.button("✨ 生成内容提示词", key="after_prompt", use_container_width=True):
                st.session_state.page = "内容生成"
                st.rerun()
        with b4:
            buf = io.BytesIO()
            ed = pd.DataFrame(st.session_state.notes)[
                ["title", "content", "likes", "collects", "comments",
                 "author", "publish_time", "url"]].copy()
            ed.insert(0, "序号", range(1, len(ed) + 1))
            ed = ed.rename(columns={"title": "标题", "content": "正文内容",
                                    "likes": "点赞数", "collects": "收藏数",
                                    "comments": "评论数", "author": "作者",
                                    "publish_time": "发布时间", "url": "链接"})
            ed.to_excel(buf, index=False, engine="openpyxl")
            buf.seek(0)
            st.download_button("⬇️ 下载帖子 Excel",
                               buf.getvalue(),
                               file_name=f"xhs_{kw or 'data'}_采集数据.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)


# ======================================================================
# 页面：帖子数据（当前任务的帖子列表 / 筛选 / 排序 / 详情 / 下载）
# ======================================================================
def page_posts():
    if not st.session_state.notes:
        st.info("当前暂无帖子数据，请先完成竞品采集。可：① 在左侧采集竞品；"
                "② 从「设置」上传 Excel；③ 加载本地已采集数据。")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("📕 去采集", key="empty_to_collect", use_container_width=True):
                st.session_state.page = "小红书采集"
                st.rerun()
        with c2:
            if st.button("📂 加载本地已采集数据", key="load_local"):
                if os.path.exists(DATA_FILE):
                    notes = json.load(open(DATA_FILE, encoding="utf-8"))
                    recs, summary = analyze_notes(notes)
                    st.session_state.notes = recs
                    st.session_state.analysis_summary = summary
                    task = make_task({"keyword": "本地数据", "goals": []}, recs, "已分析")
                    st.session_state.current_task = task
                    # 入库 + 登记 + 快照：刷新后仍可从此处续上
                    persist_context(task, recs)
                    st.rerun()
                else:
                    st.warning("暂无本地数据文件")
        return

    ensure_analysis()
    recs = st.session_state.notes
    task = st.session_state.current_task
    kw = task["keyword"] if task else "—"

    st.markdown('<p class="sec-title">竞品采集数据</p>', unsafe_allow_html=True)
    h1, h2 = st.columns([3, 1])
    with h1:
        st.caption(f"当前关键词：**{kw}**")
    with h2:
        st.caption(f"采集数量：**{len(recs)}** 篇")

    f1, f2, f3 = st.columns([2, 2, 2])
    with f1:
        search = st.text_input("🔍 搜索（标题/正文）", key="collect_search",
                                placeholder="输入关键词过滤")
    with f2:
        filt = st.radio("筛选", ["全部", "高赞", "高收藏", "高评论"], horizontal=True, key="collect_filt")
    with f3:
        sort_by = st.selectbox("排序", ["互动量 ↓", "点赞 ↓", "收藏 ↓", "评论 ↓", "最新"], key="collect_sort")

    df = pd.DataFrame(recs)
    mask = pd.Series(True, index=df.index)
    if search.strip():
        s = search.strip().lower()
        mask = df["title"].str.lower().str.contains(s, na=False) | \
               df["content"].str.lower().str.contains(s, na=False)
    if filt == "高赞":
        thr = df["likes"].quantile(0.7) if len(df) else 0
        mask &= df["likes"] >= thr
    elif filt == "高收藏":
        thr = df["collects"].quantile(0.7) if len(df) else 0
        mask &= df["collects"] >= thr
    elif filt == "高评论":
        thr = df["comments"].quantile(0.7) if len(df) else 0
        mask &= df["comments"] >= thr
    view = df[mask].copy()
    sv = {"互动量 ↓": "_heat", "点赞 ↓": "likes", "收藏 ↓": "collects",
          "评论 ↓": "comments", "最新": "publish_time"}
    if sort_by == "互动量 ↓":
        view["_heat"] = view.apply(lambda r: compute_heat(r), axis=1)
    view = view.sort_values(sv[sort_by], ascending=(sort_by == "最新"))

    # 主从布局：左列表（一行一个「查看」），右详情（上下文操作集中在这里）
    left, right = st.columns([1.6, 1])
    with left:
        for _, r in view.iterrows():
            rid = r["id"]
            a = r.get("analysis", {})
            row = st.container(border=True)
            cA, cB = row.columns([4, 1])
            with cA:
                st.markdown(f"**{r['title'][:40] or '(无标题)' }**")
                st.caption(f"🔥 互动量 {compute_heat(r)}　👍 {r['likes']}　❤️ {r['collects']}　💬 {r['comments']}　"
                           f"🕒 {r['publish_time'] or '—'}　🏷️ {a.get('cover','—')}")
            with cB:
                if st.button("查看详情", key="v_" + rid, use_container_width=True):
                    st.session_state.selected_note_id = rid
        if len(view) == 0:
            st.caption("无匹配结果")

    with right:
        sel = next((r for r in recs if r["id"] == st.session_state.selected_note_id), None)
        if sel:
            st.markdown('<p class="sec-title">帖子详情</p>', unsafe_allow_html=True)
            with st.container(border=True):
                st.markdown(f"### {sel['title'] or '(无标题)'}")
                st.write(sel["content"] or "（无正文，可能为纯图文笔记）")
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("点赞", sel["likes"])
                m2.metric("收藏", sel["collects"])
                m3.metric("评论", sel["comments"])
                m4.metric("互动量", compute_heat(sel))
                st.caption(f"发布时间：{sel['publish_time'] or '—'}")
                st.caption(f"关键词：{kw}")
                if sel.get("local_images"):
                    st.image(sel["local_images"][:4], width=120)
                elif sel.get("cover"):
                    try:
                        st.image(sel["cover"], width=200)
                    except Exception:
                        st.caption("（封面预览被防盗链拦截，可用下方「批量下载筛选图片」下载后查看）")
                # 上下文操作集中在详情面板，避免列表行按钮过多
                a1, a2 = st.columns(2)
                with a1:
                    if st.button("🔥 分析此帖", key="sel_analyze", use_container_width=True):
                        st.session_state.baokuan_detail_id = sel["id"]
                        st.session_state.page = "爆款分析"
                        st.rerun()
                with a2:
                    if st.button("✍️ 以此生成内容", key="sel_gen", use_container_width=True):
                        st.session_state.ctx_ref = sel["title"]
                        st.session_state.page = "内容生成"
                        st.session_state.auto_gen = True
                        st.rerun()
                if sel["url"]:
                    st.markdown(f"[🔗 打开原帖]({sel['url']})")
        else:
            st.info("点击左侧「查看详情」查看帖子详情，并可在此分析此帖 / 生成内容")

    # 底部操作区：导出 / 批量下载 / 下一步
    st.divider()
    n1, n2 = st.columns(2)
    with n1:
        if st.button("🔥 下一步：爆款分析与内容机会", key="next_baokuan", use_container_width=True):
            st.session_state.page = "爆款分析"
            st.rerun()
    with n2:
        if st.button("✍️ 下一步：生成小红书内容提示词", key="next_create", use_container_width=True):
            st.session_state.page = "内容生成"
            st.rerun()
    e1, e2 = st.columns(2)
    with e1:
        if st.button("⬇️ 导出当前数据 Excel", key="export_collect"):
            ed = pd.DataFrame(recs)[["title", "content", "likes", "collects", "comments",
                                     "author", "publish_time", "url"]].copy()
            ed.insert(0, "序号", range(1, len(ed) + 1))
            ed = ed.rename(columns={"title": "标题", "content": "正文内容", "likes": "点赞数",
                                     "collects": "收藏数", "comments": "评论数", "author": "作者",
                                     "publish_time": "发布时间", "url": "链接"})
            buf = io.BytesIO()
            ed.to_excel(buf, index=False, engine="openpyxl")
            buf.seek(0)
            st.download_button("下载 Excel", buf.getvalue(),
                               file_name=f"xhs_{kw}_采集数据.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with e2:
        # 批量下载当前筛选结果的图片（跳过已有本地图的笔记）
        by_id = {r["id"]: r for r in recs}
        view_ids = [rid for rid in view["id"]]
        need_ids = [rid for rid in view_ids
                    if not (by_id.get(rid) or {}).get("local_images")]
        have = len(view_ids) - len(need_ids)
        if st.button(f"🖼️ 批量下载筛选图片（高清无水印）· 待下载 {len(need_ids)} 篇"
                     + (f" · 已有 {have} 篇" if have else ""),
                     key="dl_images", disabled=not need_ids):
            start_image_download(need_ids)
            st.rerun()


# ======================================================================
# 页面：爆款分析 + 竞品洞察 + 内容机会
# ======================================================================
def page_baokuan():
    if not st.session_state.notes:
        st.info("当前暂无帖子数据，请先生成爆款分析表。", icon="ℹ️")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("📕 去采集", key="bao_empty_collect", use_container_width=True):
                st.session_state.page = "小红书采集"
                st.rerun()
        with c2:
            if st.button("⬆️ 去上传 Excel", key="bao_empty_upload", use_container_width=True):
                st.session_state.page = "设置"
                st.rerun()
        return

    ensure_analysis()
    recs = st.session_state.notes
    summary = st.session_state.analysis_summary
    task = st.session_state.current_task
    kw = task["keyword"] if task else "—"

    st.markdown("# 🔥 爆款分析")
    st.caption(f"从已采集的竞品内容中，找出值得学习的内容规律。　"
               f"当前任务：{kw}　|　样本 {len(recs)} 篇")

    # 上方摘要三栏：爆款内容 / 爆款分析结果 / 重点内容
    baokuan_recs = [r for r in recs if r["analysis"].get("baokuan")]
    sum1, sum2, sum3 = st.columns(3)
    sum1.metric("爆款内容", f"{len(baokuan_recs)} 篇")
    top_heat = max((r["analysis"]["heat"] for r in recs), default=0)
    sum2.metric("最高互动量", f"{top_heat}")
    sum3.metric("重点内容", f"{min(len(baokuan_recs), 5)} 篇（按热度 Top）")

    st.divider()

    # 左侧卡片列表 + 右侧详情面板
    ranked = sorted(recs, key=lambda r: r["analysis"]["heat"], reverse=True)
    top_cards = ranked[:30]

    # 选中爆款：若还没选，默认选第一条
    if "bao_selected_id" not in st.session_state:
        st.session_state.bao_selected_id = top_cards[0]["id"] if top_cards else None
    # 如果任务切换导致选中不在当前集合里，归位到第一条
    if st.session_state.bao_selected_id not in {r["id"] for r in top_cards}:
        st.session_state.bao_selected_id = top_cards[0]["id"] if top_cards else None

    left, right = st.columns([1.4, 1.6], gap="large")
    with left:
        st.markdown('<p class="sec-title">爆款卡片</p>', unsafe_allow_html=True)
        for r in top_cards:
            a = r["analysis"]
            selected = (r["id"] == st.session_state.bao_selected_id)
            border = True
            with st.container(border=border):
                if selected:
                    st.markdown('<span class="tag" style="background:#ffe7e7;color:#c00;border-color:#ffd6d6;">● 当前选中</span>',
                                unsafe_allow_html=True)
                st.markdown(f"**{r['title'][:48] or '(无标题)'}**")
                st.caption(f"🔥 {a['heat']}　👍 {r['likes']}　❤️ {r['collects']}　💬 {r['comments']}　"
                           f"{a['baokuan']}")
                b1, b2 = st.columns([1, 4])
                if b1.button("查看", key="bao_card_" + r["id"], use_container_width=True):
                    st.session_state.bao_selected_id = r["id"]
                    st.rerun()
                b2.caption("👇 右侧查看完整爆款拆解")

    with right:
        st.markdown('<p class="sec-title">爆款详情</p>', unsafe_allow_html=True)
        sel = next((r for r in recs if r["id"] == st.session_state.bao_selected_id), None)
        if not sel:
            st.info("请选择一条爆款内容进行分析。")
        else:
            a = sel["analysis"]
            # ---- 原始帖子 ----
            with st.container(border=True):
                st.markdown("**原始帖子**")
                st.markdown(f"### {sel['title'] or '(无标题)'}")
                st.write(sel["content"] or "（无正文，可能为纯图文笔记）")
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("点赞", sel["likes"])
                m2.metric("收藏", sel["collects"])
                m3.metric("评论", sel["comments"])
                m4.metric("综合互动", a["heat"])
                st.caption(f"发布时间：{sel.get('publish_time','—')}　|　关键词：{kw}")
                if sel.get("url"):
                    st.markdown(f"[🔗 打开原帖]({sel['url']})")

            # ---- 为什么它容易爆？ ----
            with st.container(border=True):
                st.markdown("**为什么它容易爆？**")
                st.markdown(f"· 爆款原因：{a['baokuan']}（热度位于全部样本前段）")
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("**标题结构**")
                    st.write(f"· 钩子类型：{a['hook_type']}")
                    st.write(f"· 钩子句：{a['hook_sentence']}")
                    st.markdown("**核心卖点**")
                    for p in (a["selling_points"] or ["待补充"])[:5]:
                        st.write("· " + p)
                with c2:
                    st.markdown("**用户痛点**")
                    st.write("· " + (a["hook_sentence"] or "（从正文未提取到明确痛点）"))
                    st.markdown("**内容情绪**")
                    st.write("· 真实 / 实用 / 经验分享（基于正文与钩子识别）")

                st.markdown("**内容结构**")
                steps = ["痛点引入", "产品体验", "使用场景", "具体评价", "购买建议"]
                present = []
                txt = (sel["title"] or "") + (sel["content"] or "")
                if any(k in txt for k in ["踩雷", "避雷", "别再", "为什么", "后悔", "劝退", "坑"]):
                    present.append("痛点引入")
                if any(k in txt for k in ["体验", "用了", "感觉", "上手"]):
                    present.append("产品体验")
                if any(k in txt for k in ["场景", "客厅", "卧室", "日常", "时候"]):
                    present.append("使用场景")
                if any(k in txt for k in ["整体", "来说", "总结", "评价"]):
                    present.append("具体评价")
                if any(k in txt for k in ["推荐", "入手", "买", "建议", "链接"]):
                    present.append("购买建议")
                flow = " ↓ ".join(present) if present else " → ".join(steps)
                st.info(flow)

            # ---- 可复用爆款结构 ----
            with st.container(border=True):
                st.markdown("**可复用爆款结构**")
                st.markdown("**标题公式**：" + a["hook_type"])
                st.markdown("**正文公式**：问题 → 场景 → 产品 → 使用体验 → 总结")
                st.markdown("**结尾钩子**：总结 + 建议 + 互动（提问 / 收藏引导）")
                # 一键带入内容生成
                if st.button("✨ 基于此爆款生成内容", type="primary",
                             use_container_width=True, key="bao_to_content"):
                    # 自动把爆款标题/钩子/结构/卖点带入 AI 创作页的 ctx_*
                    st.session_state.ctx_competitor = kw
                    st.session_state.ctx_keyword = kw
                    st.session_state.ctx_ref = sel["title"]
                    st.session_state.ctx_focus_in = a["hook_sentence"] or ""
                    st.session_state.ctx_opp_in = a["baokuan"]
                    st.session_state.ctx_style = a["hook_type"]
                    st.session_state.page = "内容生成"
                    st.session_state.auto_gen = True
                    st.rerun()

    st.divider()

    # AI 竞品洞察
    st.markdown('<p class="sec-title">AI 竞品洞察</p>', unsafe_allow_html=True)
    ins = get_insights(recs, summary)
    if ins:
        i1, i2 = st.columns(2)
        with i1:
            with st.container(border=True):
                st.markdown("**竞品主要内容方向**")
                for d in ins["directions"]:
                    st.write("· " + d)
            with st.container(border=True):
                st.markdown("**高频关键词**")
                st.write("、".join(k["kw"] for k in ins["freq_kw"][:12]) or "—")
            with st.container(border=True):
                st.markdown("**热门内容类型**")
                for t in ins["hot_types"]:
                    st.write(f"· {t['t']}（{t['cnt']}篇）")
        with i2:
            with st.container(border=True):
                st.markdown("**高互动关键词（按平均赞+藏）**")
                for k in ins["high_kw"][:10]:
                    st.write(f"· {k['kw']}　均 {k['avg']}")
            with st.container(border=True):
                st.markdown("**用户主要关注点**")
                st.write("、".join(p for p, _ in ins["focus"][:8]) or "—")
            with st.container(border=True):
                st.markdown("**竞品卖点**")
                st.write("、".join(s["p"] for s in ins["selling"][:10]) or "—")
        with st.container(border=True):
            st.markdown("**爆款共性**")
            for f in ins["common"]:
                st.write("· " + f)

    st.divider()

    # 内容机会
    st.markdown('<p class="sec-title">内容机会</p>', unsafe_allow_html=True)
    opps = get_opportunities(recs, summary)
    ocols = st.columns(len(opps))
    for i, op in enumerate(opps):
        with ocols[i]:
            with st.container(border=True):
                st.markdown(f"**{op['title']}**")
                st.caption(op["evidence"])
                st.write(op["desc"])
                if st.button("使用这个机会创作 →", key="opp_" + str(i)):
                    st.session_state.ctx_opp = op["title"]
                    st.session_state.ctx_opp_kw = op["kw"]
                    st.session_state.ctx_keyword = kw
                    st.session_state.ctx_competitor = kw
                    st.session_state.page = "内容生成"
                    st.session_state.auto_gen = True
                    st.rerun()

    st.divider()
    b1, b2 = st.columns(2)
    with b1:
        if st.button("✍️ 下一步：生成小红书内容提示词", key="next_create_from_baokuan",
                     use_container_width=True):
            st.session_state.page = "内容生成"
            st.rerun()
    with b2:
        if st.button("⬇️ 导出爆款分析表 Excel", key="export_report", use_container_width=True):
            tmp = os.path.join(HERE, "_cur_notes.json")
            json.dump([{**r, "analysis": r["analysis"]} for r in recs],
                      open(tmp, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
            try:
                subprocess.run([sys.executable, "build_report.py", "--input", tmp,
                                "--output", REPORT_FILE, "--nomaster", "--force"],
                               cwd=HERE, check=True)
                with open(REPORT_FILE, "rb") as f:
                    data = f.read()
                st.download_button("下载分析表", data, file_name="xhs_report.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.success("已生成爆款分析表")
            except Exception as e:
                st.error("导出失败：" + str(e))


# ======================================================================
# 页面：AI 创作
# ======================================================================
# AI 创作：草稿编辑框相关回调
# 用 on_click 回调在「下一次脚本执行、控件实例化之前」写入 session_state，
# 避免在控件实例化后再赋值触发 StreamlitAPIException。
# ======================================================================
def clear_draft_edits():
    """清空草稿编辑框缓存 key（title_i / body_i）。必须在这些控件被实例化之前调用。"""
    for k in list(st.session_state.keys()):
        if k.startswith("title_") or k.startswith("body_"):
            del st.session_state[k]


def adopt_title_cb(i, t):
    st.session_state[f"title_{i}"] = t
    st.session_state["last_action"] = ("success", f"已采用标题：{t}")


def optimize_title_cb(i, fallback):
    cur = st.session_state.get(f"title_{i}", fallback)
    new = optimize_title(cur)
    if new == cur:
        st.session_state["last_action"] = ("info", "这个标题已经比较紧凑，暂无更优改写方案")
        return
    st.session_state[f"title_{i}"] = new
    st.session_state["last_action"] = ("success", f"标题已优化：{new}")


def reduce_ad_cb(i, fallback):
    cur = st.session_state.get(f"body_{i}", fallback)
    new = reduce_ad(cur)
    if new == cur:
        st.session_state["last_action"] = ("info", "正文未检测到营销话术或堆叠感叹号，广告感已较低，无需弱化")
        return
    st.session_state[f"body_{i}"] = new
    st.session_state["last_action"] = ("success", "已降低广告感：移除营销话术、收敛感叹号")


def enhance_real_cb(i, fallback):
    cur = st.session_state.get(f"body_{i}", fallback)
    new = enhance_real(cur)
    if new == cur:
        st.session_state["last_action"] = ("info", "正文已包含真实体验描述，无需重复添加")
        return
    st.session_state[f"body_{i}"] = new
    st.session_state["last_action"] = ("success", "已增强真实体验：补充第一人称使用感受")


def regen_draft_cb(i):
    ctx = st.session_state.draft_context or {}
    drafts = st.session_state.drafts
    themes = st.session_state.draft_themes or st.session_state.notes
    model = st.session_state.draft_model or build_content_model(st.session_state.notes)
    bp = load_brand(ctx.get("brand")) if ctx.get("brand") else None
    # 换一个选题角度重新生成（避免同一主题导致标题一成不变）
    n = max(len(themes), 1)
    j = (i + random.randint(1, max(n - 1, 1))) % n
    new = gen_one(themes[j], model, bp,
                  ctx.get("brand", ""), ctx.get("product"), idx=random.randint(0, 9))
    drafts[i] = new
    st.session_state.drafts = drafts
    st.session_state[f"title_{i}"] = new["titles"][0]
    st.session_state[f"body_{i}"] = new["body"]
    st.session_state["last_action"] = ("success", f"已换角度重新生成笔记 {i+1}")


# ======================================================================
def page_create():
    st.markdown('<p class="sec-title">AI 种草内容创作</p>', unsafe_allow_html=True)

    # 数据来源模式提示：有数据走竞品驱动，无数据走独立创作（填关键词即可生成）
    if st.session_state.notes:
        st.caption(f"✅ 竞品数据驱动模式：已载入 {len(st.session_state.notes)} 条真实采集数据，"
                   "生成文案会自动带入爆款规律与高互动卖点")
    else:
        st.info("✍️ 独立创作模式：无需先采集数据。填写下方「关键词」等条件即可直接生成种草笔记；"
                "之后采集/导入竞品数据，文案会自动升级为竞品爆款驱动。")

    # 预填创作条件：若已采集且输入框为空，自动带入竞品/关键词（仅首次，保留用户后续编辑）
    if not st.session_state.get("ctx_competitor"):
        st.session_state.ctx_competitor = st.session_state.keyword
    if not st.session_state.get("ctx_keyword"):
        st.session_state.ctx_keyword = st.session_state.keyword

    # 来自「内容机会」的自动生成
    if st.session_state.auto_gen:
        st.session_state.auto_gen = False
        ctx = build_context()
        drafts, model, themes = generate_drafts(ctx)
        if drafts:
            clear_draft_edits()  # 清掉上一批草稿的编辑缓存，避免残留旧内容
            st.session_state.drafts = drafts
            st.session_state.draft_model = model
            st.session_state.draft_themes = themes
            st.session_state.draft_context = ctx
            st.success(f"已基于「{ctx.get('content_opportunity') or '竞品数据'}」生成 {len(drafts)} 篇种草笔记")

    left, right = st.columns([1, 1.4])
    with left:
        st.markdown("**创作条件**")
        st.text_input("当前竞品", key="ctx_competitor")
        st.text_input("关键词", key="ctx_keyword")
        st.text_input("目标用户", key="ctx_user")
        st.text_input("内容目标", key="ctx_goal")
        st.text_input("内容风格", key="ctx_style")
        st.text_input("参考爆款", key="ctx_ref")
        st.text_input("用户关注点", key="ctx_focus_in")
        st.text_input("内容机会", key="ctx_opp_in")
        st.text_input("品牌名（可选）", key="ctx_brand")
        st.text_input("产品名（可选）", key="ctx_product")
        if st.button("🪄 生成种草笔记", type="primary", use_container_width=True, key="gen_drafts"):
            # 同步两个非 widget key 的输入到 context；其余 ctx_* 已由控件 key 自动写入 session_state
            # 注意：禁止在此再赋值 ctx_competitor/ctx_keyword 等 widget key，会触发 StreamlitAPIException
            st.session_state["ctx_user_focus"] = st.session_state.get("ctx_focus_in", "")
            st.session_state["ctx_opp"] = st.session_state.get("ctx_opp_in", "")
            ctx = build_context()
            drafts, model, themes = generate_drafts(ctx)
            if drafts:
                clear_draft_edits()  # 清掉上一批草稿的编辑缓存，避免残留旧内容
                st.session_state.drafts = drafts
                st.session_state.draft_model = model
                st.session_state.draft_themes = themes
                st.session_state.draft_context = ctx
                st.rerun()

    with right:
        st.markdown("**生成结果**")
        # 显示上一次操作（采用/优化/降广告感/增强体验/重新生成）的真实反馈
        last_action = st.session_state.pop("last_action", None)
        if last_action:
            kind, msg = last_action
            getattr(st, kind, st.info)(msg)
        drafts = st.session_state.drafts
        if not drafts:
            st.info("点击左侧「🪄 生成种草笔记」开始创作："
                    + ("系统会基于当前竞品与爆款数据生成多标题、正文、标签。"
                       if st.session_state.notes else
                       "填写关键词即可直接生成；若填了品牌名/产品名，会套用品牌调性。"))
            return
        ctx = st.session_state.draft_context
        # 首次渲染时把生成的标题/正文写入 session_state，供可编辑控件显示
        # 必须在 text_input/text_area 实例化之前完成，且不能再用 value= 传默认值
        for i, d in enumerate(drafts):
            if f"title_{i}" not in st.session_state:
                st.session_state[f"title_{i}"] = d["titles"][0]
            if f"body_{i}" not in st.session_state:
                st.session_state[f"body_{i}"] = d["body"]
        for i, d in enumerate(drafts):
            with st.container(border=True):
                st.markdown(f"**▍笔记 {i+1}**")
                # 标题候选
                st.markdown("标题推荐：")
                for j, t in enumerate(d["titles"]):
                    tc1, tc2 = st.columns([5, 1])
                    tc1.write(f"{j+1}. {t}")
                    tc2.button("采用", key=f"adopt_{i}_{j}",
                               on_click=adopt_title_cb, args=(i, t))
                st.text_input("当前标题", key=f"title_{i}")
                st.text_area("正文（可直接修改）", key=f"body_{i}", height=240)
                st.caption("标签：" + d.get("tags", ""))
                bc1, bc2, bc3, bc4 = st.columns(4)
                bc1.button("优化标题", key=f"opt_{i}",
                           on_click=optimize_title_cb, args=(i, d["titles"][0]))
                bc2.button("降低广告感", key=f"ad_{i}",
                           on_click=reduce_ad_cb, args=(i, d["body"]))
                bc3.button("增强真实体验", key=f"real_{i}",
                           on_click=enhance_real_cb, args=(i, d["body"]))
                bc4.button("重新生成", key=f"re_{i}",
                           on_click=regen_draft_cb, args=(i,))
                # 内容质量
                q = compute_quality(st.session_state.get(f"title_{i}", d["titles"][0]),
                                    st.session_state.get(f"body_{i}", d["body"]), ctx)
                st.markdown("**内容质量**")
                for k, v in q.items():
                    col = st.columns([1, 3])
                    col[0].caption(k)
                    col[1].progress(v / 100.0)
                    col[1].caption(str(v))

        if st.session_state.drafts:
            if st.button("⬇️ 导出种草笔记 Excel", key="export_content"):
                vals = []
                for i, d in enumerate(st.session_state.drafts):
                    vals.append({
                        "title": st.session_state.get(f"title_{i}", d["titles"][0]),
                        "body": st.session_state.get(f"body_{i}", d["body"]),
                        "tags": d.get("tags", ""),
                        "hook": d.get("hook", ""),
                        "selling": d.get("selling", ""),
                    })
                out = export_drafts_xlsx(vals, CONTENT_FILE)
                with open(out, "rb") as f:
                    st.download_button("下载 xhs_content.xlsx", f.read(),
                                       file_name="xhs_content.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ======================================================================
# 页面：内容生成（输出小红书内容提示词，规格第 18-21 节）
# 数据来源 = 用户产品信息 + 爆款分析（上一阶段结果）
# ======================================================================
def page_content():
    from generate_prompt import build_prompt, save_prompt, load_prompts

    PROMPTS_FILE = os.path.join(HERE, "xhs_prompts.json")
    prompts = load_prompts(PROMPTS_FILE)
    st.session_state.prompts = prompts

    st.markdown("# ✨ AI 小红书内容生成")
    st.caption("把竞品帖子 → 爆款规律 → 与你自己的产品结合 → 生成一份可直接给任意 LLM 使用的创作提示词。"
               "这不是聊天框，提示词是「分析结果」的产物。")

    # 数据来源模式提示
    has_data = bool(st.session_state.notes)
    has_analysis = bool(st.session_state.analysis_summary)
    if has_data:
        st.markdown('<span class="tag" style="background:#fff5f5;color:#c00;border-color:#ffd6d6;">● 爆款规律驱动模式</span>',
                    unsafe_allow_html=True)
        st.caption(f"已载入 {len(st.session_state.notes)} 条真实帖子 + 爆款分析结果，"
                   "提示词的「参考爆款结构」段落会自动从这些数据里抽取。")
    else:
        st.markdown('<span class="tag">● 独立创作模式</span>', unsafe_allow_html=True)
        st.caption("暂无已采集数据。提示词仍可生成，但「参考爆款结构」将使用通用模板。"
                   "想要更精准的提示词，先到「小红书采集」采集竞品。")

    st.divider()

    # ===== 上方：参考爆款规律（来自爆款分析结果） =====
    with st.expander("📊 参考爆款规律（来自爆款分析）", expanded=False):
        if not has_data:
            st.caption("暂无爆款分析结果。可在「爆款分析」页生成。")
        else:
            kw = (st.session_state.current_task or {}).get("keyword", "")
            ranked = sorted(st.session_state.notes,
                            key=lambda r: r["analysis"]["heat"], reverse=True)
            top = ranked[:3]
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**标题结构（前 3 条爆款）**")
                for r in top:
                    st.write("· " + (r["title"][:40] or "（无标题）"))
            with c2:
                st.markdown("**用户痛点 / 核心卖点**")
                pains, sellings = [], []
                for r in top:
                    a = r.get("analysis", {})
                    pains.append(a.get("hook_sentence", ""))
                    sellings.extend(a.get("selling_points", []) or [])
                st.write("· " + "、".join(p[:30] for p in pains if p) or "—")
                st.write("· " + "、".join(sellings[:20]) or "—")
            if kw:
                st.caption(f"参考关键词：{kw}")

    # ===== 自动带入（来自爆款「基于此爆款生成内容」按钮） =====
    if st.session_state.auto_gen and not st.session_state.get("content_product"):
        # 从 ctx_* 自动填入到 content_*
        mapping = {
            "content_product": st.session_state.ctx_keyword or st.session_state.keyword,
            "content_audience": st.session_state.get("ctx_user", ""),
            "content_scenario": "",
            "content_selling": "",
            "content_direction": st.session_state.get("ctx_style", "") or "真实体验分享",
            "content_extra": st.session_state.get("ctx_focus_in", ""),
            "content_ref_title": st.session_state.get("ctx_ref", ""),
        }
        for k, v in mapping.items():
            if v:
                st.session_state[k] = v
        st.session_state.auto_gen = False

    st.divider()

    # ===== 输入区 =====
    st.markdown('<p class="sec-title">我的产品信息</p>', unsafe_allow_html=True)
    with st.container(border=True):
        c1, c2 = st.columns(2)
        with c1:
            st.text_input("我的产品（必填）", key="content_product",
                          placeholder="如：小户型布艺沙发")
            st.text_input("目标用户", key="content_audience",
                          placeholder="如：25-35岁租房女性")
            st.text_input("使用场景", key="content_scenario",
                          placeholder="如：10㎡左右小客厅")
        with c2:
            st.text_input("核心卖点（顿号分隔）", key="content_selling",
                          placeholder="如：省空间 / 高颜值 / 好打理")
            st.text_input("内容方向", key="content_direction",
                          placeholder="如：干货测评 / 真实体验 / 故事分享")
            st.text_input("附加要求（可选）", key="content_extra",
                          placeholder="如：增加具体使用场景、避免营销话术")

        if st.session_state.get("content_ref_title"):
            st.info(f"参考爆款标题：《{st.session_state.content_ref_title}》")

    st.divider()

    # ===== 主按钮：完整种草内容 / 结构化提示词 =====
    bc1, bc2 = st.columns([1.15, 1])
    with bc1:
        # 生成逻辑在 run_studio_cb（on_click 回调）里执行，避免按钮分支写 widget key
        st.button("📝 生成完整种草内容", type="primary",
                  use_container_width=True, key="studio_gen",
                  on_click=run_studio_cb)
    with bc2:
        if st.button("🧠 仅生成提示词（给 LLM 用）", use_container_width=True,
                     key="content_gen"):
            if not (st.session_state.content_product or "").strip():
                st.warning("请先填写「我的产品」")
            else:
                product = {
                    "product": st.session_state.content_product,
                    "audience": st.session_state.content_audience,
                    "scenario": st.session_state.content_scenario,
                    "selling_points": st.session_state.content_selling,
                    "content_direction": st.session_state.content_direction,
                    "extra": st.session_state.content_extra,
                    "ref_baokuan_title": st.session_state.content_ref_title,
                }
                recs = st.session_state.notes or []
                analyses = [r.get("analysis", {}) for r in recs]
                obj = build_prompt(product, recs, analyses)
                st.session_state.last_prompt = obj
                st.session_state.last_product = product
                st.success("已生成小红书内容提示词", icon="✨")
                st.rerun()

    st.divider()

    # 创作工作室操作反馈（成功/采用/优化/保存/空产品告警）——在结果区之前消费，
    # 保证即使尚未生成（studio 为空）也能显示「请先填写产品」类提示
    _last_action = st.session_state.pop("studio_last_action", None)
    if _last_action:
        _kind, _msg = _last_action
        getattr(st, _kind, st.info)(_msg)

    # ===== 输出区 ①：创作工作室（完整种草内容，PRD §20-§30） =====
    if st.session_state.get("studio"):
        render_studio_result()
        st.divider()

    # ===== 输出区 ②：结构化提示词（原能力保留） =====
    obj = st.session_state.get("last_prompt")
    if not obj:
        if not st.session_state.get("studio"):
            st.info("👆 填写「我的产品信息」后，点「📝 生成完整种草内容」直接成稿；"
                    "或点「🧠 仅生成提示词」让任意 LLM 代写。")
        return
    if st.session_state.get("studio"):
        st.markdown('<p class="sec-title">🧠 结构化提示词（可选：交给任意 LLM 深度创作）</p>',
                    unsafe_allow_html=True)

    st.markdown('<p class="sec-title">小红书内容生成提示词</p>', unsafe_allow_html=True)
    with st.container(border=True):
        st.markdown(f"**产品**：{obj['refs'].get('product','—')}　"
                    f"**参考爆款**：{obj['refs'].get('ref_baokuan_title') or '—'}")
        # 用 text_area 让用户可复制（自带复制按钮）
        # 注意：不带 key=，直接用 value= 展示提示词；避免 key 已存在时报 StreamlitAPIException
        st.text_area("提示词（点击右上角图标复制）", value=obj["prompt"], height=420)
        # 文件下载
        st.download_button("⬇️ 下载提示词为 .md",
                           obj["prompt"].encode("utf-8"),
                           file_name=f"xhs_prompt_{obj['refs'].get('product','data')}.md",
                           mime="text/markdown",
                           use_container_width=True)
        bc1, bc2, bc3 = st.columns(3)
        with bc1:
            # 用 Streamlit 自带的 toast / info 提示「复制」（剪贴板受浏览器限制，统一用下载兜底）
            if st.button("📋 一键复制到剪贴板", key="content_copy",
                         use_container_width=True):
                # 用 streamlit 的剪贴板支持：v1.27+ 提供 clipboard，但兼容性差；这里同时给下载 + 提示
                st.info("已把提示词选中（请按 Ctrl+C / ⌘+C 复制下方选区）；"
                        "如浏览器禁用了剪贴板，请使用「下载提示词」按钮。",
                icon="📋")
        with bc2:
            if st.button("🔄 重新生成", key="content_regen",
                         use_container_width=True):
                # 用当前输入重新 build_prompt（不清空，覆盖旧结果）
                product = {
                    "product": st.session_state.content_product,
                    "audience": st.session_state.content_audience,
                    "scenario": st.session_state.content_scenario,
                    "selling_points": st.session_state.content_selling,
                    "content_direction": st.session_state.content_direction,
                    "extra": st.session_state.content_extra,
                    "ref_baokuan_title": st.session_state.content_ref_title,
                }
                recs = st.session_state.notes or []
                analyses = [r.get("analysis", {}) for r in recs]
                st.session_state.last_prompt = build_prompt(product, recs, analyses)
                st.session_state.last_product = product
                st.success("已重新生成小红书内容提示词", icon="✨")
                st.rerun()
        with bc3:
            if st.button("💾 保存到历史记录", key="content_save",
                         type="primary", use_container_width=True):
                task_id = (st.session_state.current_task or {}).get("id", "")
                product = st.session_state.last_product or {}
                record = save_prompt(PROMPTS_FILE, task_id, product, obj)
                st.session_state.prompts = load_prompts(PROMPTS_FILE)
                st.success(f"已保存提示词 #{record['id']}，到「历史记录」查看")
                st.rerun()


# ======================================================================
# 创作工作室（PRD §20-§30）：标题工厂 / 成稿 / 评分 / 风险 / 封面 / 选题
# ======================================================================
def _collect_studio_product():
    """从 content_* 输入控件读当前产品信息。"""
    return {
        "product": st.session_state.get("content_product", ""),
        "audience": st.session_state.get("content_audience", ""),
        "scenario": st.session_state.get("content_scenario", ""),
        "selling_points": st.session_state.get("content_selling", ""),
        "content_direction": st.session_state.get("content_direction", ""),
        "extra": st.session_state.get("content_extra", ""),
        "ref_baokuan_title": st.session_state.get("content_ref_title", ""),
    }


def _studio_sources():
    recs = st.session_state.notes or []
    analyses = [r.get("analysis", {}) for r in recs]
    return recs, analyses


def run_studio_cb():
    """生成完整种草内容（回调内写 widget key，规则 2 安全时序）。"""
    from content_studio import studio_pipeline
    product = _collect_studio_product()
    if not (product.get("product") or "").strip():
        st.session_state["studio_last_action"] = ("warning", "请先填写「我的产品」")
        return
    recs, analyses = _studio_sources()
    try:
        res = studio_pipeline(product, recs, analyses)
    except Exception as e:  # 引擎异常不阻塞页面
        st.session_state["studio_last_action"] = ("error", f"内容生成失败：{e}")
        return
    st.session_state.studio = res
    st.session_state.studio_product = product
    st.session_state.studio_title = res["note"]["title"]
    st.session_state.studio_body = res["note"]["body"]
    src = res["source"]
    st.session_state["studio_last_action"] = (
        "success", f"已生成完整种草内容（{src}）：10 个标题 + 正文 + 7 维评分 + 风险检测 + 封面建议")


def regen_studio_cb():
    from content_studio import studio_pipeline
    product = _collect_studio_product()
    if not (product.get("product") or "").strip():
        st.session_state["studio_last_action"] = ("warning", "请先填写「我的产品」")
        return
    recs, analyses = _studio_sources()
    try:
        res = studio_pipeline(product, recs, analyses)
    except Exception as e:
        st.session_state["studio_last_action"] = ("error", f"重新生成失败：{e}")
        return
    st.session_state.studio = res
    st.session_state.studio_product = product
    st.session_state.studio_title = res["note"]["title"]
    st.session_state.studio_body = res["note"]["body"]
    st.session_state["studio_last_action"] = ("success", "已基于当前输入重新生成 ✨")


def adopt_studio_title_cb(i):
    res = st.session_state.get("studio") or {}
    titles = res.get("titles") or []
    if 0 <= i < len(titles):
        t = titles[i]["title"]
        st.session_state.studio_title = t
        st.session_state["studio_last_action"] = ("success", f"已采用标题：{t}")


def deai_studio_cb():
    """✨ 优化表达：去 AI 感。"""
    from content_studio import deai_text
    cur = st.session_state.get("studio_body", "") or ""
    if not cur:
        st.session_state["studio_last_action"] = ("info", "正文为空，无可优化内容")
        return
    new, changed = deai_text(cur)
    if new == cur or not changed:
        st.session_state["studio_last_action"] = (
            "info", "正文已经很自然：无套话、含真实口吻与互动问句，无需改动")
        return
    st.session_state.studio_body = new
    st.session_state["studio_last_action"] = (
        "success", f"✨ 已优化表达，调整 {len(changed)} 处：" + "；".join(changed[:3])
        + ("…" if len(changed) > 3 else ""))


def save_studio_cb():
    """把当前成稿（含用户编辑）保存到内容资产库。"""
    from content_studio import save_asset
    res = st.session_state.get("studio")
    if not res:
        return
    title = st.session_state.get("studio_title", "") or ""
    body = st.session_state.get("studio_body", "") or ""
    note = res.get("note", {})
    asset = {
        "created_at": res.get("created_at", ""),
        "product": res.get("product", ""),
        "audience": res.get("audience", ""),
        "direction": res.get("direction", ""),
        "source": res.get("source", ""),
        "title": title,
        "body": body,
        "tags": note.get("tags", ""),
        "score": res.get("score", {}),
        "cover": res.get("cover", {}),
        "risks": res.get("risks", []),
    }
    try:
        rec = save_asset(asset)
        st.session_state["studio_last_action"] = (
            "success", f"已保存到内容资产库 #{rec['id']}，可在「历史记录」查看")
    except Exception as e:
        st.session_state["studio_last_action"] = ("error", f"保存失败：{e}")


def render_studio_result():
    """内容生成页：创作工作室结果区（PRD §20-§30）。"""
    res = st.session_state.get("studio") or {}

    note = res.get("note", {})
    score = res.get("score", {})
    total = score.get("total", 0)
    risks = res.get("risks", [])
    titles = res.get("titles", [])
    product = res.get("product", "")
    source = res.get("source", "")

    st.markdown('<p class="sec-title">📝 完整种草内容</p>', unsafe_allow_html=True)
    st.caption(f"产品：{product}　|　内容方向:{res.get('direction','—')}　|　"
               f"来源：{source}　|　生成时间：{res.get('created_at','')}")

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("综合评分", f"{total}/100")
    m2.metric("风险提示", f"{len(risks)} 条" if risks else "0 条")
    m3.metric("标题候选", f"{len(titles)} 个")
    m4.metric("正文字数", f"{len(note.get('body',''))}")
    m5.metric("选题", f"{len(res.get('topics',[]))} 个")

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
        ["📝 笔记成稿", "🎯 10 个标题", "🩺 质量评分", "🛡️ 风险检测", "🎨 封面建议", "🔥 选题库"])

    # ---- Tab1 成稿 ----
    with tab1:
        st.text_input("标题（可编辑）", key="studio_title")
        st.text_area("正文（可直接修改）", key="studio_body", height=340)
        st.caption("标签：" + (note.get("tags", "") or "—"))
        c1, c2, c3, c4 = st.columns([1, 1, 1, 1])
        with c1:
            st.button("✨ 优化表达（去 AI 感）", key="studio_deai",
                      on_click=deai_studio_cb, use_container_width=True)
        with c2:
            st.button("🔄 重新生成", key="studio_regen",
                      on_click=regen_studio_cb, use_container_width=True)
        with c3:
            st.button("💾 保存到内容库", key="studio_save", type="primary",
                      on_click=save_studio_cb, use_container_width=True)
        with c4:
            _body = st.session_state.get("studio_body", note.get("body", ""))
            st.download_button("⬇️ 下载 .md", _body.encode("utf-8"),
                               file_name=f"xhs_note_{product or 'draft'}.md",
                               mime="text/markdown", use_container_width=True)
        st.caption("小贴士：正文修改后可直接点「✨ 优化表达」去 AI 感；保存时以当前编辑内容为准。")

    # ---- Tab2 标题工厂 ----
    with tab2:
        if not titles:
            st.caption("暂无标题，先在上方生成。")
        for i, c in enumerate(titles):
            s = c.get("scores", {})
            tag_colors = {"高点击型": "#ffe58f", "痛点型": "#ffccc7", "清单型": "#d3f0d3",
                          "测评型": "#d6e4ff", "避坑型": "#ffd6e7"}
            bg = tag_colors.get(c.get("category", ""), "#f0f0f0")
            score_html = " · ".join(
                f"{k} <b>{v}</b>" for k, v in s.items())
            cL, cR = st.columns([5.4, 1])
            with cL:
                st.markdown(
                    f"""<div style="border:1px solid #eee;border-radius:10px;padding:8px 12px;margin:2px 0;background:#fff;">
<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
<span style="background:{bg};border-radius:12px;padding:1px 8px;font-size:12px;color:#333;">{c.get('category','')}</span>
<span style="font-weight:600;">{c.get('title','')}</span>
</div>
<div style="font-size:12px;color:#888;margin-top:4px;">{score_html}　<span style="color:#c00;">广告感越低越好</span></div>
<div style="font-size:12px;color:#666;margin-top:2px;">为什么：{c.get('reason','')}</div>
</div>""", unsafe_allow_html=True)
            with cR:
                st.button("✔ 采用", key=f"studio_adopt_{i}",
                          on_click=adopt_studio_title_cb, args=(i,),
                          use_container_width=True)
        st.caption("点击「采用」后，标题会写回上方「笔记成稿」的标题输入框。")

    # ---- Tab3 质量评分 ----
    with tab3:
        dims = score.get("dims", [])
        if not dims:
            st.caption("暂无评分。")
        for name, v in dims:
            col = st.columns([1.2, 3, 1.2])
            col[0].caption(name)
            col[1].progress(v / 100.0)
            col[2].caption(str(v))
        st.markdown(f"**综合评分：{total}/100**")
        tips = score.get("tips", {})
        if tips:
            st.caption("提升建议：")
            for k, v in tips.items():
                st.caption(f"· {k}：{v}")

    # ---- Tab4 风险检测 ----
    with tab4:
        if not risks:
            st.success("✅ 未检测到绝对化表达 / 夸大宣传 / 过度营销 / 可疑数据，可放心使用。")
        else:
            for r in risks:
                st.warning(f"⚠️ **{r['type']}**（命中：「{r.get('word','')}」）　→　{r.get('suggestion','')}")
            st.caption("提示：可用上方「✨ 优化表达」或手动修改后重新评分。")

    # ---- Tab5 封面建议 ----
    with tab5:
        cover = res.get("cover", {})
        st.markdown(f"**封面标题**：{cover.get('cover_title', '—')}")
        for k, txt in [("画面建议", cover.get("visual", "")),
                       ("排版建议", cover.get("layout", "")),
                       ("配色建议", cover.get("color", ""))]:
            if txt:
                st.write(f"**{k}**：{txt}")

    # ---- Tab6 选题库 ----
    with tab6:
        for i, tp in enumerate(res.get("topics", [])):
            c = st.columns([5, 1, 1.5])
            c[0].write(f"{i+1}. {tp['topic']}")
            c[1].markdown(f'<span class="tag">热度 {tp["heat"]}</span>', unsafe_allow_html=True)
            c[2].button("用作主题", key=f"studio_topic_{i}",
                        on_click=_adopt_topic_cb, args=(tp["topic"],))
        st.caption("「用作主题」会把选题追加到正文开头，方便按选题方向微调成稿。")

def _adopt_topic_cb(topic):
    body = st.session_state.get("studio_body", "") or ""
    if topic and not body.startswith(f"「{topic}」"):
        st.session_state.studio_body = f"「{topic}」\n\n" + body
        st.session_state["studio_last_action"] = ("success", f"已应用选题：{topic}")


# ======================================================================
# 页面：历史记录（任务卡片 + Prompt 记录 + 节点状态）
# ======================================================================
def page_history():
    from generate_prompt import load_prompts
    PROMPTS_FILE = os.path.join(HERE, "xhs_prompts.json")
    prompts = load_prompts(PROMPTS_FILE)

    st.markdown("# 🗂️ 历史记录")
    st.caption("所有真实任务、提示词与内容成稿都在这里保存。")

    # ---- 内容资产库（PRD §30）----
    from content_studio import load_assets, delete_asset
    assets = load_assets()
    st.markdown(f'<p class="sec-title">📚 我的内容资产库（{len(assets)}）</p>',
                unsafe_allow_html=True)
    if not assets:
        st.caption("暂无内容。在「内容生成」页点「📝 生成完整种草内容」，点「💾 保存到内容库」后出现在这里。")
    else:
        for a in assets[:20]:
            with st.container(border=True):
                r1, r2, r3, r4 = st.columns([2.4, 1.2, 1.2, 1])
                r1.write(f"**{a.get('title','（无标题）')[:42]}**")
                r2.caption(f"{a.get('product','')[:12]} · {a.get('source','')}")
                r3.caption(f"综合 {a.get('score',{}).get('total','—')} 分")
                r4.caption(a.get("created_at", "")[:16])
                with st.expander("查看全文 / 标签 / 封面建议"):
                    st.text_area("", value=a.get("body", ""), height=260,
                                 key=f"asset_body_{a.get('id')}")
                    st.caption("标签：" + (a.get("tags", "") or "—"))
                    cover = a.get("cover", {}) or {}
                    if cover.get("visual"):
                        st.write(f"**封面标题**：{cover.get('cover_title','—')}")
                        st.caption(cover.get("visual", ""))
                    st.caption("质量分：" + "，".join(
                        f"{k} {v}" for k, v in (a.get("score", {}) or {}).get("dims", [])))
                    if st.button("🗑 删除此内容", key="asset_del_" + a.get("id", "")):
                        delete_asset(a.get("id", ""))
                        st.rerun()
                    st.download_button("⬇️ 下载 .md", a.get("body", "").encode("utf-8"),
                                       file_name=f"xhs_note_{a.get('product','draft')}.md",
                                       mime="text/markdown",
                                       key="asset_dl_" + a.get("id", ""))
    st.divider()

    # ---- 提示词历史（最新生成的可独立取用）----
    st.markdown('<p class="sec-title">小红书内容提示词历史</p>', unsafe_allow_html=True)
    if not prompts:
        st.info("暂无提示词记录。在「内容生成」页生成提示词并点「保存到历史记录」即可出现。")
    else:
        for p in prompts[:10]:
            with st.container(border=True):
                r1, r2, r3, r4 = st.columns([2, 1.4, 1.4, 1])
                r1.write(f"**{p.get('product') or '（未命名产品）'}**")
                r2.caption(f"目标用户：{p.get('audience','—')[:24]}")
                r3.caption(f"时间：{p.get('created_at','')[:16]}")
                if r4.button("查看提示词", key="hp_view_" + p["id"]):
                    st.session_state["history_view_prompt"] = p
                    st.rerun()
                with st.expander("查看完整提示词", expanded=bool(st.session_state.get("history_view_prompt", {}).get("id") == p["id"])):
                    # 不带 key=，直接用 value= 展示；避免 key 已存在时报 StreamlitAPIException
                    st.text_area("", value=p.get("prompt", ""), height=300)
                    st.download_button("⬇️ 下载",
                                       p.get("prompt", "").encode("utf-8"),
                                       file_name=f"prompt_{p['id']}.md",
                                       mime="text/markdown",
                                       key="hd_" + p["id"])

    st.divider()

    # ---- 任务历史 ----
    st.markdown('<p class="sec-title">采集 / 分析任务</p>', unsafe_allow_html=True)
    if not st.session_state.tasks:
        st.info("暂无任务。在首页创建竞品分析任务后，会自动保存到此处。")
        return
    for t in st.session_state.tasks:
        with st.container(border=True):
            r1, r2, r3, r4 = st.columns([2, 1.2, 1.2, 3])
            r1.write(f"**{t['name']}**")
            r2.caption("关键词：" + t["keyword"])
            r3.caption("时间：" + t["created_at"])
            r4.caption(f"状态：{t.get('status', '—')}　|　数量：{t.get('note_count', 0)} 篇")

            b1, b2, b3, b4, b5 = st.columns(5)
            if b1.button("继续分析", key="htk_v_" + t["id"], use_container_width=True):
                data = load_task_data(t["id"])
                if data is None:
                    st.warning("任务数据缺失，请到「小红书采集」重采")
                else:
                    recs, summary = analyze_notes(data)
                    st.session_state.notes = recs
                    st.session_state.analysis_summary = summary
                    st.session_state.current_task = t
                    st.session_state.selected_note_id = None
                    persist_context(t, recs, register=False)
                    st.session_state.page = "帖子数据"
                    st.rerun()
            if b2.button("查看爆款", key="htk_b_" + t["id"], use_container_width=True):
                data = load_task_data(t["id"])
                if data is None:
                    st.warning("任务数据缺失")
                else:
                    recs, summary = analyze_notes(data)
                    st.session_state.notes = recs
                    st.session_state.analysis_summary = summary
                    st.session_state.current_task = t
                    persist_context(t, recs, register=False)
                    st.session_state.page = "爆款分析"
                    st.rerun()
            if b3.button("去生成提示词", key="htk_p_" + t["id"], use_container_width=True):
                data = load_task_data(t["id"])
                if data is not None:
                    recs, _ = analyze_notes(data)
                    st.session_state.notes = recs
                    st.session_state.current_task = t
                    persist_context(t, recs, register=False)
                st.session_state.page = "内容生成"
                st.rerun()
            if b4.button("导出 Excel", key="htk_x_" + t["id"], use_container_width=True):
                data = load_task_data(t["id"])
                if not data:
                    st.warning("任务数据缺失")
                else:
                    buf = io.BytesIO()
                    ed = pd.DataFrame(data)[["title", "content", "likes", "collects",
                                              "comments", "author", "publish_time", "url"]].copy()
                    ed.insert(0, "序号", range(1, len(ed) + 1))
                    ed = ed.rename(columns={"title": "标题", "content": "正文内容",
                                            "likes": "点赞数", "collects": "收藏数",
                                            "comments": "评论数", "author": "作者",
                                            "publish_time": "发布时间", "url": "链接"})
                    ed.to_excel(buf, index=False, engine="openpyxl")
                    buf.seek(0)
                    st.download_button("下载", buf.getvalue(),
                                       file_name=f"{t['name']}_{t['id']}.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       key="hex_" + t["id"])
            if b5.button("删除", key="htk_d_" + t["id"], use_container_width=True):
                # 删除的是当前工作上下文 → 同步清空会话与快照，避免恢复到已删除数据
                cur = st.session_state.get("current_task")
                if cur and cur.get("id") == t["id"]:
                    st.session_state.notes = []
                    st.session_state.analysis_summary = None
                    st.session_state.current_task = None
                    if os.path.exists(CUR_CTX_FILE):
                        try:
                            os.remove(CUR_CTX_FILE)
                        except Exception:
                            pass
                st.session_state.tasks = [x for x in st.session_state.tasks if x["id"] != t["id"]]
                p = os.path.join(TASK_DIR, t["id"] + ".json")
                if os.path.exists(p):
                    os.remove(p)
                save_tasks()
                st.rerun()


# ======================================================================
# 页面：设置（Excel 导入 + 基础设置）
# ======================================================================
def page_settings():
    st.markdown("# ⚙️ 设置")
    st.caption("数据导入、登录态管理、采集参数默认值等。")

    # ---- 数据导入 ----
    st.markdown('<p class="sec-title">数据导入</p>', unsafe_allow_html=True)
    st.caption("无网络或暂不想登录小红书时，可上传已有的小红书采集表（含 标题/正文/点赞/收藏/评论/链接），"
               "系统将自动分析并可进入帖子数据 / 爆款分析 / 内容生成完整流程。")
    uploaded = st.file_uploader("选择 .xlsx 文件", type=["xlsx"], key="upload_xlsx")
    if uploaded and st.button("🔍 生成分析表", key="upload_btn", use_container_width=True):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(uploaded, data_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                st.error("文件为空")
                return
            header = [str(h or "").strip() for h in rows[0]]

            def col(*names):
                for nm in names:
                    for i, h in enumerate(header):
                        if nm in h:
                            return i
                return None

            i_title = col("标题")
            i_content = col("正文", "内容")
            i_likes = col("点赞")
            i_collects = col("收藏")
            i_comments = col("评论")
            i_url = col("链接", "笔记链接")
            i_author = col("作者")
            i_date = col("发布", "时间")
            if i_title is None or i_likes is None:
                st.error("未找到「标题」或「点赞」列，请检查表头。")
                return

            def to_int(v):
                try:
                    return int(float(str(v).replace(",", "").replace("万", "0000").replace("+", "")))
                except Exception:
                    return 0

            records = []
            for r in rows[1:]:
                if not r or not r[i_title]:
                    continue
                url = str(r[i_url]) if i_url is not None and r[i_url] else ""
                records.append({
                    "id": url or str(r[i_title]),
                    "title": str(r[i_title]),
                    "author": str(r[i_author]) if i_author is not None and r[i_author] else "",
                    "url": url, "cover": "",
                    "content": str(r[i_content]) if i_content is not None and r[i_content] else "",
                    "publish_time": str(r[i_date]) if i_date is not None and r[i_date] else "",
                    "likes": to_int(r[i_likes]),
                    "collects": to_int(r[i_collects]) if i_collects is not None else 0,
                    "comments": to_int(r[i_comments]) if i_comments is not None else 0,
                    "shares": 0, "fans": 0,
                })
            if not records:
                st.error("未解析到任何数据行")
                return
            recs, summary = analyze_notes(records)
            st.session_state.notes = recs
            st.session_state.analysis_summary = summary
            task = make_task({"keyword": "Excel导入", "goals": []}, recs, "已分析")
            st.session_state.current_task = task
            # 导入数据也入库 + 登记 + 快照：刷新后「历史记录」可继续分析
            persist_context(task, recs)
            st.success(f"已分析 {len(recs)} 条，可进入「帖子数据 / 爆款分析 / 内容生成」")
        except Exception as e:
            st.error("上传分析失败：" + str(e))


# ======================================================================
# 主流程
# ======================================================================
def main():
    init_state()
    cleanup_sessions()          # 每天最多执行一次：清理过期访客会话文件
    render_topbar()
    render_sidebar()

    # 采集 / 图片下载 实时进度
    if st.session_state.scraping:
        show_scrape_progress()
        return
    if st.session_state.downloading_images:
        show_image_progress()
        return

    page = st.session_state.page
    if page == "首页":
        page_home()
    elif page == "小红书采集":
        page_xhs_collect()
    elif page == "帖子数据":
        page_posts()
    elif page == "爆款分析":
        page_baokuan()
    elif page == "内容生成":
        page_content()
    elif page == "历史记录":
        page_history()
    elif page == "设置":
        page_settings()
    # 兼容旧名称（防止旧链接 / session_state 残留导致崩溃）
    elif page == "工作台":
        page_home()
    elif page == "采集数据":
        page_posts()
    elif page == "AI创作":
        page_content()
    elif page == "任务中心":
        page_history()
    elif page == "数据导入":
        page_settings()


if __name__ == "__main__":
    main()
