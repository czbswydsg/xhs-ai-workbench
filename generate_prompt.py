# -*- coding: utf-8 -*-
"""
小红书内容提示词生成器
============================================
输入：用户产品信息（产品/用户/场景/卖点/内容方向）+ 爆款分析规律（来自 analyzer）
输出：一份完整、结构化的小红书创作提示词，可直接复制给任何 LLM 使用。

设计原则：
- 不做 AI 聊天，不直接生成最终文案，专注生成高质量提示词。
- 提示词结构参考二十节示例：角色 / 产品 / 用户 / 场景 / 卖点 / 爆款结构 / 内容结构 / 表达要求 / 输出格式。
- 数据完全基于真实采集数据 + 用户输入，不编造。
"""
from datetime import datetime
from typing import Dict, List, Optional

try:
    from analyzer import extract_hook, extract_keywords, extract_selling_points
except Exception:  # analyzer 不可用时也能降级
    extract_hook = extract_keywords = extract_selling_points = None


def _safe_get(rec, key, default=""):
    if not rec:
        return default
    v = rec.get(key) if isinstance(rec, dict) else None
    return v if v is not None else default


def _baokuan_rules(notes, analyses, top_n=3):
    """从已分析的真实数据里抽取爆款规律，用于提示词的「参考爆款结构」段落。"""
    rules = {"title_patterns": [], "content_patterns": [], "pains": [], "sellings": []}
    if not notes or not analyses:
        return rules
    # 取互动量最高的若干条
    pairs = list(zip(notes, analyses))
    pairs.sort(key=lambda x: _safe_get(x[0], "likes", 0) + _safe_get(x[0], "collects", 0) * 1.5, reverse=True)
    top = pairs[:max(1, top_n)]
    for n, _ in top:
        title = _safe_get(n, "title", "")
        if title:
            # 标题结构：用空格/破折号/换行切出短片段
            segs = [s.strip() for s in re_split(title) if s.strip()]
            rules["title_patterns"].append({"title": title[:60], "segments": segs[:6]})
        # 内容结构：从 analyzer 的 hook 字段提取
        if extract_hook:
            try:
                _hook = extract_hook(n)
                # extract_hook 返回 (hook_type, hook_sentence) 元组；兼容字符串/None
                if isinstance(_hook, tuple):
                    h = _hook[1] if len(_hook) > 1 else _hook[0]
                else:
                    h = _hook
                if h:
                    rules["content_patterns"].append({"hook": str(h).strip()})
            except Exception:
                pass
        # 用户痛点 / 卖点
        if extract_selling_points:
            try:
                sps = extract_selling_points(n) or []
                for sp in sps[:3]:
                    s = sp.get("text") if isinstance(sp, dict) else str(sp)
                    if s:
                        rules["sellings"].append(s)
            except Exception:
                pass
        kw_list = _safe_get(n, "keywords", []) or []
        rules["pains"].extend(kw_list[:3])
    # 去重保序
    rules["pains"] = _dedup(rules["pains"])[:8]
    rules["sellings"] = _dedup(rules["sellings"])[:8]
    return rules


def re_split(s):
    import re as _re
    return _re.split(r"[\s|｜·。!！？?\-—]+", s or "")


def _dedup(seq):
    seen = set()
    out = []
    for x in seq:
        k = str(x).strip()
        if not k or k in seen:
            continue
        seen.add(k)
        out.append(k)
    return out


def build_prompt(product: Dict, baokuan_notes: Optional[list] = None,
                 baokuan_analyses: Optional[list] = None) -> Dict:
    """
    根据用户产品信息 + 爆款规律生成一份完整提示词。

    product keys:
      - product: 我的产品（必填）
      - audience: 目标用户
      - scenario: 使用场景
      - selling_points: 核心卖点（可用顿号/逗号分隔）
      - content_direction: 内容方向（如：干货测评 / 真实体验 / 故事分享）
      - ref_baokuan_title: 来自上一阶段选中的爆款标题（可选）
      - extra: 任意附加要求
    """
    product_name = (product.get("product") or "").strip()
    if not product_name:
        product_name = "（请填写你的产品）"
    audience = (product.get("audience") or "").strip() or "（未指定）"
    scenario = (product.get("scenario") or "").strip() or "（未指定）"
    selling = (product.get("selling_points") or "").strip() or "（未指定）"
    direction = (product.get("content_direction") or "").strip() or "真实体验分享"
    extra = (product.get("extra") or "").strip()
    ref_title = (product.get("ref_baokuan_title") or "").strip()

    rules = _baokuan_rules(baokuan_notes or [], baokuan_analyses or [])

    # 组合标题/正文结构示例（基于真实爆款提取 + 默认模板）
    title_pattern_hint = "痛点 + 场景 + 结果" if not rules["title_patterns"] else \
        "参考真实爆款标题分片：" + " | ".join(
            "/".join(p["segments"]) for p in rules["title_patterns"][:3]
        )
    content_pattern_hint = "问题 → 场景 → 解决方案 → 真实体验 → 总结建议"
    if rules["content_patterns"]:
        def _hook_str(p):
            v = p.get("hook")
            if isinstance(v, tuple):
                v = v[1] if len(v) > 1 else v[0]
            return str(v).strip() if v else ""
        content_pattern_hint += "\n参考爆款开头：" + " / ".join(
            s for s in (_hook_str(p) for p in rules["content_patterns"][:2]) if s
        )
    pains_hint = "、".join(rules["pains"]) if rules["pains"] else "（可补充）"
    sellings_hint = "、".join(rules["sellings"]) if rules["sellings"] else "（可补充）"

    lines = [
        "# 小红书内容生成提示词",
        "",
        f"> 由「AI 小红书内容增长工作台」基于真实爆款数据自动生成 · {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "",
        "## 角色",
        "你是一名擅长小红书内容创作的内容运营，熟悉小红书的标题套路、互动心理与平台规则。",
        "",
        "## 任务",
        f"请围绕「{product_name}」创作一条小红书笔记。",
        "",
        "## 目标用户",
        audience,
        "",
        "## 使用场景",
        scenario,
        "",
        "## 核心卖点",
        selling,
        "",
        "## 内容方向",
        direction,
    ]
    if extra:
        lines += ["", "## 附加要求", extra]
    if ref_title:
        lines += ["", "## 参考爆款标题", f"《{ref_title}》"]
    lines += [
        "",
        "## 参考爆款结构（来自真实采集数据）",
        f"- 标题结构：{title_pattern_hint}",
        f"- 内容结构：{content_pattern_hint}",
        f"- 用户痛点：{pains_hint}",
        f"- 核心卖点：{sellings_hint}",
        "",
        "## 表达要求",
        "- 真实、生活化，避免 AI 模板化与夸张营销。",
        "- 用具体使用场景代替空话。",
        "- 不要使用绝对化用词（最、第一、必买）。",
        "- 加入 1-2 个小细节或个人体验。",
        "",
        "## 输出格式",
        "1. 标题（1 条，可附 2 条备选）",
        "2. 正文（300-600 字）",
        "3. 标签（5-8 个 #话题）",
    ]
    prompt_text = "\n".join(lines)
    return {
        "prompt": prompt_text,
        "rules": rules,
        "refs": {
            "product": product_name,
            "audience": audience,
            "scenario": scenario,
            "selling_points": selling,
            "content_direction": direction,
            "extra": extra,
            "ref_baokuan_title": ref_title,
        },
    }


def save_prompt(prompts_file: str, task_id: str, product: Dict, prompt_obj: Dict) -> Dict:
    """把生成的提示词落盘，返回完整记录。"""
    import json, os
    data = []
    if os.path.exists(prompts_file):
        try:
            data = json.load(open(prompts_file, encoding="utf-8"))
        except Exception:
            data = []
    record = {
        "id": "P" + datetime.now().strftime("%Y%m%d%H%M%S"),
        "task_id": task_id or "",
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "product": product.get("product", ""),
        "audience": product.get("audience", ""),
        "scenario": product.get("scenario", ""),
        "selling_points": product.get("selling_points", ""),
        "content_direction": product.get("content_direction", ""),
        "ref_baokuan_title": product.get("ref_baokuan_title", ""),
        "prompt": prompt_obj["prompt"],
    }
    data.insert(0, record)
    json.dump(data, open(prompts_file, "w", encoding="utf-8"),
              ensure_ascii=False, indent=2)
    return record


def load_prompts(prompts_file: str):
    import json, os
    if not os.path.exists(prompts_file):
        return []
    try:
        return json.load(open(prompts_file, encoding="utf-8"))
    except Exception:
        return []
