# -*- coding: utf-8 -*-
"""
小红书内容创作引擎（离线启发式版，不依赖外部 LLM API）
========================================================
对应 PRD §20-§30：创作输入 → 标题工厂(10个×5类×评分) → 正文成稿 →
话题标签 → 质量评分(7维+综合) → 风险检测 → ✨去AI感优化 → 封面建议 → 选题库。

纯函数模块，不依赖 Streamlit；供 app.py 内容生成页调用。
真实爆款数据存在时自动带入 rules（标题分片/开头句/痛点/卖点）。

后续如接入 LLM，只需替换 gen_titles / gen_full_note 内部实现，函数签名保持不变。
"""
import os
import random
import re
from datetime import datetime

try:
    from generate_prompt import _baokuan_rules
except Exception:  # 独立运行/依赖缺失时可降级
    _baokuan_rules = None

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS_FILE = os.path.join(HERE, "xhs_assets.json")

# ----------------------------------------------------------------------
# 词库
# ----------------------------------------------------------------------
ABSOLUTE_WORDS = ["最好用", "最强", "最值得", "最划算", "最便宜", "全网第一",
                  "绝对", "100%", "百分百", "百分之一百", "必定", "保证",
                  "永不", "永久", "根治", "断根", "全部都能", "所有都",
                  "无敌", "零风险", "必买", "闭眼入不出错", "绝对有效"]
EXAGGERATE_WORDS = ["秒杀", "暴瘦", "逆袭", "神效", "奇迹", "一夜", "立刻见效",
                    "马上见效", "急速", "狂甩", "暴涨", "躺着瘦", "不反弹"]
MARKETING_WORDS = ["限量", "仅限", "最后一天", "亏本", "骨折价", "清仓",
                   "点击下方", "左下角", "私我", "加微信", "+v", "加V",
                   "扫码购买", "立即下单", "秒回"]
STOCK_AI_PHRASES = [
    ("总的来说", ""), ("总而言之", ""), ("综上所述", ""), ("值得注意的是", ""),
    ("需要注意的是", ""), ("首先", ""), ("其次", ""), ("最后", ""),
    ("该产品", "它"), ("此款产品", "它"), ("这款产品", "它"),
    ("具有良好的", "挺"), ("非常值得推荐", "我自己是挺喜欢的"),
    ("有着很好的", "很"), ("在...方面表现优异", "表现不错"),
]
REAL_WORDS = ["我", "自己", "真实", "用了", "感受", "体验", "日常", "闺蜜",
              "我妈", "室友", "通勤", "周末", "上手", "质感", "手感"]

ASSET_DIMS = ["标题吸引力", "用户匹配度", "内容完整度", "可读性",
              "种草能力", "原创度", "自然度"]


# ----------------------------------------------------------------------
# 通用
# ----------------------------------------------------------------------
def _dedup(seq):
    seen, out = set(), []
    for x in seq:
        k = str(x).strip()
        if k and k not in seen:
            seen.add(k)
            out.append(k)
    return out


def _seg(v):
    """顿号/逗号/分号分隔 → 词条列表"""
    if not v:
        return []
    return [s.strip() for s in re.split(r"[、，,；;|/]", str(v)) if s.strip()]


def _seeded(product, audience):
    seed = sum(map(ord, (product or "x") + (audience or "y"))) % 9973
    return random.Random(seed)


# ----------------------------------------------------------------------
# 标题工厂：10 个标题 = 5 类 × 2，每标题 4 维评分 + 原因
# ----------------------------------------------------------------------
def gen_titles(product, audience, scenario, selling_items, rules=None):
    rules = rules or {}
    pains = rules.get("pains") or []
    pain = pains[0] if pains else "踩雷/交智商税"
    if len(pain) > 12:
        pain = pain[:12]
    sells = selling_items or []
    s1 = sells[0] if sells else "适合自己"
    s2 = sells[1] if len(sells) > 1 else s1
    aud = (audience or "").strip() or "姐妹们"
    if len(aud) > 14:
        aud = aud[:14]
    prod = (product or "").strip() or "这个单品"
    scn = (scenario or "").strip() or "日常"
    if len(scn) > 10:
        scn = scn[:10]

    # 真实爆款标题分片（若有）→ 作为"测评型"灵感
    real_slice = ""
    tps = rules.get("title_patterns") or []
    if tps:
        segs = tps[0].get("segments") or []
        real_slice = "｜".join(segs[:3]) if segs else ""

    templates = [
        # 高点击型（好奇/悬念）
        ("高点击型",
         f"{aud}注意！关于{prod}，没人告诉你的几个真相",
         "人群 + 悬念句式，制造「信息差」点击欲"),
        ("高点击型",
         f"别再乱买{prod}了，{pain}的人先把这篇看完",
         "否定指令 + 人群限定 + 痛点，反差感强"),
        # 痛点型
        ("痛点型",
         f"{aud}最怕的{pain}，选{prod}前先看这几点",
         "直接命中{pain}，目标用户代入感强"),
        ("痛点型",
         f"如果{scn}要选{prod}，这 3 个坑千万别踩",
         "场景化痛点 + 数字清单，风险提示明显"),
        # 清单型
        ("清单型",
         f"{aud}闭眼抄作业：{prod}挑选清单，照着买不踩雷",
         "清单承诺 + 低决策成本，收藏率高"),
        ("清单型",
         f"{prod}选购 5 个关键点：从{pain}到{s1}一次说清",
         "数字结构化 + 覆盖完整决策链路"),
        # 测评型
        ("测评型",
         f"亲测一个月｜{prod}到底值不值？{aud}说实话",
         "真实测评时间锚点 + 直接提问，可信度高"),
        ("测评型",
         f"{prod}测评日记：{s1}是真的吗？用真实体验说话",
         "带着质疑去验证卖点，弱广告感"),
        # 避坑型
        ("避坑型",
         f"{aud}避坑！{prod}最容易踩雷的地方，下单前看完",
         "避坑心智 + 行动指令，负面预警点击高"),
        ("避坑型",
         f"劝退帖｜{prod}并不适合所有人，{scn}党下单前先想清楚",
         "反向劝退 + 限定人群，制造「是否适合我」的好奇"),
    ]
    # 有真实爆款分片时，给第 8 条追加一行参考真实结构
    if real_slice:
        templates[7] = ("测评型",
                        f"{s1}是真是假？{prod}真实测评（结构参考：{real_slice}）",
                        "复用真实爆款标题分片结构 + 测评立场")

    titles = []
    for i, (cat, t, why) in enumerate(templates):
        # ---- 4 维评分（启发式）----
        click = 62 + min(14, len(t)) // 2
        if re.search(r"[0-9０-９]|注意|真相|别|坑|避|到底|值不值", t):
            click += 8
        if re.search(r"[？！?!]", t):
            click += 4
        # 用户相关性：人群词/你/我 + 痛点命中
        rel = 60
        rel += 18 if aud in t or "你" in t or "我" in t else 0
        rel += 12 if any(p[:4] in t for p in pains) else 0
        rel += 10 if sells and any(s[:4] in t for s in sells) else 0
        # 差异化：含具体产品/具体限定场景越具体越高
        diff = 58
        diff += 20 if prod != "这个单品" else 4
        diff += 12 if len(aud) > 2 and aud != "姐妹们" else 0
        diff += 8 if scn and scn != "日常" else 0
        # 广告感（越低越好）：绝对化/营销词命中即高
        ad = 8 + random.Random(hash(t) % 100).randint(0, 8)
        if any(w in t for w in ABSOLUTE_WORDS + MARKETING_WORDS):
            ad += 70
        elif "闭眼抄" in t:
            ad += 26
        click = min(98, click)
        rel = min(97, rel)
        diff = min(96, diff)
        ad = min(95, ad)
        titles.append({
            "category": cat,
            "title": t,
            "scores": {"点击潜力": click, "用户相关性": rel,
                       "差异化": diff, "广告感": ad},
            "reason": why,
        })
    return titles


# ----------------------------------------------------------------------
# 正文成稿：Hook → 痛点 → 场景 → 方案 → 体验 → 细节 → 总结 → 互动
# ----------------------------------------------------------------------
def _fmt_body(parts):
    return "\n\n".join(x.strip() for x in parts if x and x.strip())


def gen_full_note(product, audience, scenario, selling_items, direction,
                  rules=None, topic=""):
    rules = rules or {}
    pains = rules.get("pains") or []
    sells = selling_items or []
    aud = (audience or "").strip() or "姐妹们"
    scn = (scenario or "").strip() or "平时"
    prod = (product or "").strip() or "这个"
    s1 = sells[0] if sells else "整体挺满意"
    s2 = sells[1] if len(sells) > 1 else "性价比在线"

    # 真实爆款开头（优先）→ 否则生成
    hooks = [p for p in (rules.get("content_patterns") or []) if p.get("hook")]
    hook = hooks[0]["hook"].strip() if hooks else (
        f"说真的，作为{aud}，我在{prod}这件事上踩过的坑比谁都多。")
    if len(hook) > 70:
        hook = hook[:70] + "…"

    pain_txt = "、".join(pains[:3]) if pains else "盲目跟风、买完后悔"
    if not topic:
        topic = f"{prod}真实体验分享"

    p_hook = hook
    p_pain = (f"先说下我的情况：{pain_txt}，这些问题我基本都经历过，"
              f"所以在挑{prod}的时候格外谨慎，做了很多功课才下手。")
    p_scene = (f"平时{scn}用到它的频率很高，入手到现在也用了有一阵子了。"
               f"不是广告，就是自己花钱买的，想给同样纠结的{aud}一点参考。")
    p_solve = (f"我最后选它的理由其实很简单：{s1}，{s2}，"
               f"刚好每一项都戳中我的需求，没有为了某个亮点硬买。")
    p_exp = (f"实际用下来的感受是——{s1}这点是实实在在能感觉到的，"
             f"不是那种宣传页上说说而已。")
    p_detail = (f"一个小细节：我自己用了一两周才敢来写这篇，"
                f"刚到手的时候也怕踩雷，结果比预期稳。"
                f"身边{aud}也有人来问我到底怎么样，我都是同一句：按自己需求挑，别盲目跟风。")
    p_sum = (f"总之，如果{aud}的预算和需求刚好和它匹配，{prod}值得放进备选；"
             f"如果只是想跟风买，那再想想，钱花在刀刃上比较好。")
    p_ask = "你们在挑的时候最在意什么？评论区聊聊，我看到都会回～"
    tag_words = _dedup(
        [re.sub(r"[\s#]", "", prod), topic[:12],
         pain_txt[:8]] +
        [t[:8] for t in sells[:2]] +
        [direction[:8] if direction else ""] + ["真实分享", "好物推荐", "理性种草"])
    tags = " ".join("#" + w for w in tag_words[:6])

    body = _fmt_body([p_hook, p_pain, p_scene, p_solve, p_exp, p_detail, p_sum, p_ask, tags])
    # 内部先给一个标题兜底；流水线里会用标题工厂第 1 张卡覆盖
    title = f"{aud}亲测｜{prod}这样选不踩雷（{s1}是认真的）"
    return {"title": title, "body": body, "tags": tags,
            "topic": topic, "hook": hook}


# ----------------------------------------------------------------------
# 质量评分：PRD §25 七维 + 综合
# ----------------------------------------------------------------------
def score_content(title, body, audience, scenario, selling_items, pains=None):
    pains = pains or []
    sells = selling_items or []
    text = (title or "") + "\n" + (body or "")
    b = body or ""

    # 标题吸引力
    ta = 60
    if title and 10 <= len(title) <= 26:
        ta += 18
    if re.search(r"[0-9０-９]", title or ""):
        ta += 8
    if re.search(r"[？！?!]", title or ""):
        ta += 6
    if any(w in (title or "") for w in ["避", "别", "真相", "亲测", "测评", "清单", "值不值"]):
        ta += 8
    ta = min(98, ta)

    # 用户匹配度
    um = 45
    if audience and audience[:6] in text:
        um += 20
    um += sum(2 for w in (pains[:3]) if w[:4] and w[:4] in b) * 4
    um += sum(1 for w in ["你", "我", "姐妹", "自己"] if w in b) * 4
    um = min(98, um)

    # 内容完整度：结构段是否齐全
    marks = 0
    for kw in ["踩", "坑", "纠结", "问题", "功课"]:
        if kw in b:
            marks += 1
            break
    for kw in sells[:2] + (["体验", "用"] if True else []):
        if kw[:4] and kw[:4] in b:
            marks += 1
    if re.search(r"总之|所以|建议|值得|结论", b):
        marks += 1
    if "?" in b or "？" in b or "聊" in b:
        marks += 1
    if "#" in b:
        marks += 1
    complete = min(98, 38 + marks * 12)

    # 可读性
    rd = 70
    if 250 <= len(b) <= 800:
        rd += 12
    elif len(b) < 200:
        rd -= 12
    paragraphs = b.count("\n\n")
    if 3 <= paragraphs <= 9:
        rd += 6
    ex = b.count("！") + b.count("!")
    rd -= max(0, ex - 4) * 3
    rd -= sum(2 for w in ["综上所述", "总而言之", "由此可见", "值得注意的是"] if w in b)
    rd = max(30, min(97, rd))

    # 种草能力
    gz = 40
    gz += sum(8 for s in sells[:3] if s[:4] and s[:4] in b)
    gz += sum(3 for w in ["质感", "手感", "上手", "用了一", "一周", "两周", "亲测",
                          "真实", "细节", "不广告"] if w in b)
    gz += 6 if "推荐" in b or "值得" in b else 0
    gz = min(98, gz)

    # 原创度：命中真实爆款整句越少越好
    og = 88
    org_bad = 0
    for p in (pains or [])[:2]:
        if p[:8] and p[:8] in b:
            org_bad += 1
    og -= org_bad * 10
    og = max(50, og)

    # 自然度（AI 感低 = 分高）
    na = 80
    na -= sum(5 for w in ["综上所述", "总而言之", "值得注意的是", "需要注意的是",
                          "该产品", "这款产品", "首先", "其次", "最后"] if w in b)
    na -= max(0, ex - 5) * 4
    na += sum(2 for w in REAL_WORDS if w in b)  # 上限补偿
    na = max(25, min(98, na))

    dims = [
        ("标题吸引力", ta), ("用户匹配度", um), ("内容完整度", complete),
        ("可读性", rd), ("种草能力", gz), ("原创度", og), ("自然度", na),
    ]
    total = round(sum(v for _, v in dims) / len(dims))
    tips = {
        "标题吸引力": ("标题建议" if ta < 80 else "标题不错") + "：加入数字/悬念词或控制 12~22 字",
        "用户匹配度": ("点明人群" if um < 80 else "人群清晰") + "：开头直接带出目标用户与场景",
        "内容完整度": ("补结构" if complete < 80 else "结构完整") + "：痛点→体验→总结→互动问句齐全",
        "可读性": ("精简长句" if rd < 80 else "读起来顺") + "：控制感叹号密度，多用短句与分段",
        "种草能力": ("加细节" if gz < 80 else "有真实感") + "：写具体使用时长、手感等可感知细节",
        "原创度": ("规避撞句" if og < 80 else "表达够新") + "：避免直接照搬参考爆款的完整句子",
        "自然度": ("再口语化" if na < 80 else "比较像真人") + "：删掉书面套话，换成生活化说法",
    }
    return {"dims": dims, "total": total, "tips": tips}


# ----------------------------------------------------------------------
# 风险检测：PRD §26
# ----------------------------------------------------------------------
def detect_risks(title, body):
    text = (title or "") + "\n" + (body or "")
    risks = []
    for w in ABSOLUTE_WORDS:
        if w in text:
            risks.append({
                "type": "绝对化表达",
                "word": w,
                "suggestion": f"把「{w}」改为留有余地的说法，如「我自己用下来没怎么/基本…」",
            })
            break
    for w in EXAGGERATE_WORDS:
        if w in text:
            risks.append({
                "type": "夸大宣传",
                "word": w,
                "suggestion": f"「{w}」属于功效承诺，建议换成真实体验描述。",
            })
            break
    for w in MARKETING_WORDS:
        if w in text:
            risks.append({
                "type": "过度营销",
                "word": w,
                "suggestion": f"「{w}」营销感过强，小红书社区对导流/促销词敏感，建议删除。",
            })
            break
    m = re.search(r"[0-9０-９]{2,}\s*(万|w|%|倍|斤|kg|天|小时|次)", text)
    if m:
        risks.append({
            "type": "数据可信度",
            "word": m.group(0),
            "suggestion": "如非真实可溯源数据，建议删掉数字或改成「大概」这类模糊表述。",
        })
    return risks


# ----------------------------------------------------------------------
# ✨ 去 AI 感优化（PRD §24）
# ----------------------------------------------------------------------
def deai_text(text):
    changed = []
    t = text
    for a, b in STOCK_AI_PHRASES:
        if a in t:
            t = t.replace(a, b)
            changed.append(f"删除/替换套话「{a}」")
    t = re.sub(r"！+", "！", t)
    t = re.sub(r"!+", "!", t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    # 纯书面第三人称开头（无"我"）时插入真实体验口吻
    first_para_end = t.find("\n\n")
    head = t[:first_para_end] if first_para_end > 0 else t
    if "我" not in t[:max(60, len(t) // 4)]:
        insert = "先声明：纯自用分享，不是广告，大家按需参考。"
        t = head + "\n\n" + insert + (t[first_para_end:] if first_para_end > 0 else "")
        changed.append("开头补充「自用声明」，降低推广感")
    # 句尾问句若缺失给互动引导
    if "?" not in t and "？" not in t:
        t = t.rstrip() + "\n\n你们觉得呢？有同款的评论区举个手～"
        changed.append("末尾补互动问句")
    return t.strip(), changed


# ----------------------------------------------------------------------
# 封面建议（PRD §28）
# ----------------------------------------------------------------------
def cover_suggestion(product, scenario, selling_items, titles):
    sells = selling_items or []
    cover_title = ""
    for cand in reversed(titles):
        if cand["category"] in ("避坑型", "清单型") and len(cand["title"]) <= 14:
            cover_title = cand["title"]
            break
    if not cover_title and titles:
        t = titles[0]["title"]
        cover_title = t if len(t) <= 14 else t[:12] + "…"
    kw = (sells[:2] + [scenario]) if scenario else sells[:2]
    return {
        "cover_title": cover_title,
        "visual": (f"画面：{scenario or '日常'}场景 + {product}产品主体 + "
                   f"{'、'.join(str(k)[:6] for k in kw[:3])} 等关键词贴纸"),
        "layout": "排版：封面主标题居中放大（白字/深色描边），2~3 个卖点放底部小字，"
                  "左上角可加「真实测评」角标提升点击",
        "color": "配色：浅色背景为主（米白/奶白），产品本身作为视觉焦点，避免花哨滤镜",
    }


# ----------------------------------------------------------------------
# 选题库（PRD §29）
# ----------------------------------------------------------------------
def suggest_topics(product, audience, scenario, selling_items, rules=None, n=5):
    rules = rules or {}
    pains = rules.get("pains") or []
    sells = selling_items or []
    aud = (audience or "").strip() or "用户"
    scn = (scenario or "").strip() or ""
    topics = []
    seen = set()

    def add(t, heat):
        if t and t not in seen:
            seen.add(t)
            topics.append({"topic": t, "heat": max(30, min(98, heat))})

    add(f"{aud}实测：{product}怎么选才不踩雷", 88 + len(pains))
    if scn:
        add(f"{scn}场景下的{product}挑选指南", 84 + min(10, len(sells) * 2))
    if sells:
        add(f"{product}的{'、'.join(sells[:2])}到底是不是真的？", 86)
    if pains:
        add(f"{pains[0][:14]}？这些坑我替你踩过了", 90)
    add(f"{product}平价/进阶怎么选（真实对比）", 80)
    add(f"一年用下来，{product}最打动我的 3 个点", 82)
    # 真实数据里出现过的标题方向（若有）
    for p in (rules.get("title_patterns") or [])[:2]:
        tt = (p.get("title") or "")[:16]
        if tt:
            add(tt, 93)
    return topics[:n]


# ----------------------------------------------------------------------
# 一键流水线
# ----------------------------------------------------------------------
def studio_pipeline(product, baokuan_notes=None, baokuan_analyses=None):
    """page_content 主按钮调用：一次产出完整创作结果。"""
    prod = (product.get("product") or "").strip()
    audience = (product.get("audience") or "").strip()
    scenario = (product.get("scenario") or "").strip()
    selling_items = _seg(product.get("selling_points"))
    direction = (product.get("content_direction") or "").strip() or "真实体验分享"

    rules = {}
    if _baokuan_rules is not None and (baokuan_notes or baokuan_analyses):
        try:
            rules = _baokuan_rules(baokuan_notes or [], baokuan_analyses or [])
        except Exception:
            rules = {}

    topics = suggest_topics(prod, audience, scenario, selling_items, rules)
    titles = gen_titles(prod, audience, scenario, selling_items, rules)
    note = gen_full_note(prod, audience, scenario, selling_items, direction,
                         rules, topic=topics[0]["topic"] if topics else "")
    note["title"] = titles[0]["title"]  # 笔记标题与标题工厂第 1 张卡保持一致
    cover = cover_suggestion(prod, scenario, selling_items, titles)
    score = score_content(note["title"], note["body"], audience, scenario,
                          selling_items, rules.get("pains"))
    risks = detect_risks(note["title"], note["body"])
    return {
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "product": prod or "（未填写）",
        "audience": audience or "（未指定）",
        "scenario": scenario or "（未指定）",
        "direction": direction,
        "selling_items": selling_items,
        "rules": rules,
        "topics": topics,
        "titles": titles,
        "note": note,
        "cover": cover,
        "score": score,
        "risks": risks,
        "source": "数据驱动" if (baokuan_notes or baokuan_analyses) else "独立创作",
    }


# ----------------------------------------------------------------------
# 内容资产库（PRD §30）
# ----------------------------------------------------------------------
def load_assets():
    import json
    if not os.path.exists(ASSETS_FILE):
        return []
    try:
        with open(ASSETS_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def save_asset(asset):
    """把一条成稿存进内容资产库，返回记录 dict。"""
    import json
    data = load_assets()
    record = {
        "id": "A" + datetime.now().strftime("%Y%m%d%H%M%S"),
        "created_at": asset.get("created_at", datetime.now().strftime("%Y-%m-%d %H:%M")),
        "product": asset.get("product", ""),
        "audience": asset.get("audience", ""),
        "direction": asset.get("direction", ""),
        "source": asset.get("source", ""),
        "title": asset.get("title", ""),
        "body": asset.get("body", ""),
        "tags": asset.get("tags", ""),
        "score": asset.get("score", {}),
        "cover": asset.get("cover", {}),
        "risks_count": len(asset.get("risks", [])),
    }
    data.insert(0, record)
    import json as _json
    with open(ASSETS_FILE, "w", encoding="utf-8") as f:
        _json.dump(data, f, ensure_ascii=False, indent=2)
    return record


def delete_asset(aid):
    import json as _json
    data = [a for a in load_assets() if a.get("id") != aid]
    with open(ASSETS_FILE, "w", encoding="utf-8") as f:
        _json.dump(data, f, ensure_ascii=False, indent=2)
    return True
