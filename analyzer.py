"""
小红书爆款内容分析引擎（离线规则版，不依赖联网 / 不调用外部 API）

输入：一条笔记的字典（title / content / likes / collects / comments / shares ...）
输出：综合热度、爆款标记、封面类型、文案钩子+钩子句、核心卖点、热搜关键词(分5类)

设计原则：
- 封面/钩子/卖点/关键词 全部基于“原帖文本”判断，绝不凭空编造；
- 无法判断的项标记「待人工确认」，不强行猜测；
- 综合热度与爆款标记由“数值分布”自动判定，不写死固定阈值。
"""

# ======================================================================
# 1) 综合热度权重
#    小红书爆款的核心信号：收藏(强意向) > 评论(深互动) > 分享(破圈) > 点赞(基数大、水分多)
#    综合热度 = 点赞×1 + 收藏×1.5 + 评论×2 + 分享×3
# ======================================================================
WEIGHTS = {"likes": 1.0, "collects": 1.5, "comments": 2.0, "shares": 3.0}

# 爆款标记 → 星级（用于 Excel 直观展示）
STAR_MAP = {"爆款": "★★★★★", "潜力爆款": "★★★★", "普通": "★★", "低表现": "★"}


def stars_of(baokuan):
    return STAR_MAP.get(baokuan, "—")


def compute_heat(note):
    """综合热度（加权互动值）。"""
    l = note.get("likes") or 0
    c = note.get("collects") or 0
    m = note.get("comments") or 0
    s = note.get("shares") or 0
    return int(l * WEIGHTS["likes"] + c * WEIGHTS["collects"]
               + m * WEIGHTS["comments"] + s * WEIGHTS["shares"])


# ======================================================================
# 2) 爆款标记：按“综合热度在全部作品中的百分排位”自动判定
#    爆款      : 前 10%（热度百分排位 ≥ 90）
#    潜力爆款  : 10%～30%（≥ 70）
#    普通      : 30%～70%（≥ 30）
#    低表现    : 后 30%（< 30）
#    保险：当数据集整体热度都很低（最高热度 < 200）时，不轻易标“爆款”，
#          最高只给“潜力爆款”，避免空数据集里硬造爆款。
# ======================================================================
def classify_baokuan(heat, all_heats):
    total = len(all_heats)
    if total == 0:
        return "普通"
    # 百分排位：比当前值小的样本数 / 总数 * 100
    lower = sum(1 for h in all_heats if h < heat)
    pct = lower / total * 100.0
    max_heat = max(all_heats) if all_heats else 0
    if pct >= 90 and max_heat >= 200:
        return "爆款"
    if pct >= 70:
        return "潜力爆款"
    if pct >= 30:
        return "普通"
    return "低表现"


# ======================================================================
# 3) 封面类型：基于标题+正文关键词启发式判断（无图像识别）
#    优先级从高到低，命中即返回；都未命中返回「待人工确认」
# ======================================================================
COVER_RULES = [
    ("对比型", ["对比", "vs", "before", "前后", "测评对比", "横评"]),
    ("信息图型", ["攻略", "清单", "合集", "干货", "步骤", "怎么做", "教程", "一览", "汇总", "避雷指南"]),
    ("情绪氛围型", ["氛围", "ins", "治愈", "质感", "情绪", "松弛", "氛围感", "高级感", "生活态度"]),
    ("人物使用场景", ["真人", "上身", "实拍", "博主", "模特", "我坐在", "我躺在", "使用感受", "亲测"]),
    ("客厅整体场景", ["客厅", "卧室", "房间", "全家", "整屋", "全屋", "户型", "样板间"]),
    ("生活方式场景", ["日常", "vlog", "居家", "生活", "一日", "记录", "探店", "逛"]),
    ("产品+文字", ["文案", "标题党", "大字", "标语", "金句", "一句话"]),
    ("产品局部细节", ["细节", "特写", "面料", "材质", "做工", "纹理", "近看", "局部"]),
    ("产品正面展示", ["开箱", "展示", "种草", "安利", "测评", "实拍", "好物"]),
]


def classify_cover(note):
    text = (note.get("title", "") + " " + note.get("content", "")).lower()
    for ctype, kws in COVER_RULES:
        for kw in kws:
            if kw.lower() in text:
                return ctype
    return "待人工确认"


# ======================================================================
# 4) 文案钩子：分析标题+正文开头，给出钩子类型 + 最关键钩子句
# ======================================================================
HOOK_RULES = [
    ("痛点型", ["踩雷", "避雷", "别再", "为什么", "后悔", "劝退", "坑", "翻车", "智商税", "千万别"]),
    ("好奇型", ["竟然", "居然", "没想到", "揭秘", "真相", "原来", "被惊艳", "万万", "不敢相信", "藏在"]),
    ("情绪型", ["绝了", "太香", "哭死", "爱了", "封神", "上头", "谁懂", "破防", "心动", "一眼"]),
    ("反差型", ["平价", "平替", "学生党", "穷", "大牌", "便宜", "白菜", "源头", "不到一百", "几十块"]),
    ("清单型", ["3个", "5个", "8个", "10个", "几款", "合集", "清单", "必买", "全收录", "盘一盘", "汇总"]),
    ("场景型", ["客厅", "卧室", "出租屋", "小户型", "婚房", "书房", "阳台", "宿舍", "办公室", "厨房"]),
    ("种草型", ["安利", "种草", "推荐", "宝藏", "私藏", "墙裂", "入股", "闭眼入", "冲"]),
    ("经验分享型", ["攻略", "经验", "心得", "干货", "亲测", "整理", "总结", "避坑", "教程"]),
]


def _split_sentences(text):
    """按换行/常见标点切句，保留非空短句。"""
    parts = []
    for seg in text.replace("\r", "\n").split("\n"):
        seg = seg.strip()
        if not seg:
            continue
        # 进一步按中文句号/感叹/问号切，但保留原片段
        for s in seg.split("。"):
            s = s.strip("！？!?。\t ")
            if s:
                parts.append(s)
    return parts


def extract_hook(note):
    title = (note.get("title", "") or "").strip()
    content = (note.get("content", "") or "").strip()
    # 钩子句候选池：标题优先，其次正文前若干句
    candidates = [title] if title else []
    candidates += _split_sentences(content)[:8]

    hook_type = "待人工确认"
    hook_sentence = title or (candidates[1] if len(candidates) > 1 else "")

    # 提问型：标题/首句带问号
    first_text = (title + " " + content[:60])
    if "？" in first_text or "?" in first_text:
        hook_type = "提问型"
        # 取含问号的句子
        for s in candidates:
            if "？" in s or "?" in s:
                hook_sentence = s
                break

    # 其它类型按优先级匹配，取命中关键词所在的最相关句子
    for htype, kws in HOOK_RULES:
        for kw in kws:
            if kw.lower() in (title + " " + content).lower():
                hook_type = htype
                # 找包含该关键词的句子作为钩子句
                for s in candidates:
                    if kw.lower() in s.lower():
                        hook_sentence = s
                        break
                break
        if hook_type != "待人工确认":
            break

    hook_sentence = (hook_sentence or "").strip()
    if len(hook_sentence) > 60:
        hook_sentence = hook_sentence[:60] + "…"
    return hook_type, hook_sentence


# ======================================================================
# 5) 核心卖点：仅从原帖文本出现的维度中挑选，绝不创造
# ======================================================================
SELLING_POINTS = {
    "舒适度": ["舒适", "软", "躺", "坐感", "包裹", "回弹", "贴合", "承托", "久坐"],
    "面料": ["面料", "布艺", "棉", "麻", "绒", "科技布", "雪尼尔", " linen", "羊毛"],
    "材质": ["实木", "板材", "岩板", "金属", "皮革", "乳胶", "核桃木", "橡木", "碳钢", "大理石"],
    "颜色": ["颜色", "色", "奶油", "莫兰迪", "原木色", "奶白", "米白", "杏色", "雾霾蓝"],
    "造型": ["造型", "设计", "颜值", "外观", "款式", "线条", "曲线", "圆润", "极简"],
    "尺寸": ["尺寸", "大小", "长宽", "占地", "厘米", "cm", "小巧", "大容量"],
    "收纳": ["收纳", "储物", "隐藏", "抽屉", "置物", "可放", "装下"],
    "功能": ["功能", "可拆", "可洗", "多功能", "变形", "模块化", "可调节", "两用", "折叠"],
    "性价比": ["性价比", "平价", "便宜", "划算", "学生党", "源头", "平替", "闭眼入", "不到"],
    "品牌": ["品牌", "大牌", "源头工厂", "旗舰店", "官方", "正品"],
    "空间搭配": ["搭配", "户型", "小户型", "出租屋", "客厅", "卧室", "整体", "协调", "风格统一"],
    "耐用": ["耐用", "结实", "质量", "做工", "承重", "稳固", "不掉", "不开裂"],
    "安装": ["安装", "组装", "送货", "上门", "师傅", "免打孔", "自己装"],
    "颜值": ["颜值", "好看", "高级", "ins风", "出片", "上镜", "精致", "氛围感"],
}


def extract_selling_points(note, limit=6):
    text = (note.get("title", "") + " " + note.get("content", "")).lower()
    found = []
    for point, kws in SELLING_POINTS.items():
        if any(kw.lower() in text for kw in kws):
            found.append(point)
    return found[:limit]


# ======================================================================
# 6) 热搜关键词：分 5 类提取（产品 / 风格 / 场景 / 用户需求 / 热门话题）
# ======================================================================
KW_PRODUCT = ["沙发", "床垫", "茶几", "餐桌", "书桌", "椅子", "餐椅", "床", "衣柜", "橱柜",
              "电视柜", "窗帘", "地毯", "抱枕", "灯具", "吊灯", "台灯", "落地灯", "收纳柜",
              "置物架", "镜子", "挂画", "装饰画", "花瓶", "香薰", "绿植", "餐具", "锅",
              "行李箱", "包", "鞋", "护肤品", "化妆品", "香水", "耳机", "手机", "电脑",
              "相机", "手表", "摆件", "收纳盒", "四件套", "毛巾", "浴巾"]
KW_STYLE = ["奶油风", "极简", "北欧", "原木", "复古", "中古", "法式", "现代", "ins风", "侘寂",
            "新中式", "日式", "轻奢", "工业风", "美式", "韩系", "盐系", "多巴胺", "美拉德", "复古风"]
KW_SCENE = ["客厅", "卧室", "出租屋", "小户型", "婚房", "书房", "阳台", "儿童房", "玄关",
            "办公室", "宿舍", "卫生间", "厨房", "榻榻米", "飘窗", " Loft"]
KW_NEED = ["性价比", "平价", "便宜", "划算", "学生党", "租房", "懒人", "刚需", "母婴", "宠物",
           "小户型", "收纳", "二手", "源头工厂", "平替", "大牌平替", "打工人", "独居"]
KW_TOPIC = ["家居好物", "好物分享", "装修灵感", "种草", "平价好物", "生活好物", "居家", "断舍离",
            "收纳整理", "家居", "家装", "软装", "布置", "好物", "必入"]


def _extract_hashtags(text):
    """提取 #话题# 形式的标签。"""
    import re
    return re.findall(r"#([^#\s]+)#?", text)


def extract_keywords(note):
    title = note.get("title", "") or ""
    content = note.get("content", "") or ""
    text = title + " " + content
    text_l = text.lower()

    result = {
        "产品关键词": [k for k in KW_PRODUCT if k.lower() in text_l],
        "风格关键词": [k for k in KW_STYLE if k.lower() in text_l],
        "场景关键词": [k for k in KW_SCENE if k.lower() in text_l],
        "用户需求关键词": [k for k in KW_NEED if k.lower() in text_l],
    }
    # 热门话题：优先取 #标签，其次命中常见话题词
    tags = _extract_hashtags(content + " " + title)
    topic_words = [k for k in KW_TOPIC if k.lower() in text_l]
    merged = []
    seen = set()
    for t in tags + topic_words:
        if t not in seen:
            seen.add(t)
            merged.append(t)
    result["热门话题关键词"] = merged
    return result


# ======================================================================
# 7) 长尾关键词：基于已提取的 产品/风格/场景/需求 词，组合成用户搜索习惯短语
#    （离线启发式，贴合“小红书搜索词”形态）
# ======================================================================
def generate_long_tail(note, keywords):
    prod = keywords.get("产品关键词", [])
    style = keywords.get("风格关键词", [])
    scene = keywords.get("场景关键词", [])
    need = keywords.get("用户需求关键词", [])
    phrases = []
    seen = set()

    def add(p):
        p = (p or "").strip()
        if p and p not in seen:
            seen.add(p)
            phrases.append(p)

    main_prod = prod[0] if prod else "家具"
    main_style = style[0] if style else ""
    main_scene = scene[0] if scene else ""

    if main_scene:
        add(f"{main_scene}适合什么{main_prod}")
        add(f"{main_scene}{main_prod}怎么选")
    if main_style and main_scene:
        add(f"{main_style}{main_scene}{main_prod}推荐")
    elif main_style:
        add(f"{main_style}{main_prod}推荐")
    if main_style:
        add(f"{main_style}{main_prod}搭配")
    if need:
        add(f"{need[0]}{main_prod}推荐")
        add(f"{need[0]}党{main_prod}")
    add(f"高级感{main_prod}")
    add(f"{main_prod}选购攻略")
    add(f"{main_prod}避坑指南")
    return phrases[:8]


# ======================================================================
# 8) 单条分析入口
# ======================================================================
def analyze_note(note, all_heats):
    heat = compute_heat(note)
    baokuan = classify_baokuan(heat, all_heats)
    cover = classify_cover(note)
    hook_type, hook_sentence = extract_hook(note)
    selling = extract_selling_points(note)
    keywords = extract_keywords(note)
    long_tail = generate_long_tail(note, keywords)
    # 把关键词压成可写入单元格的字符串（分号分隔）
    kw_flat = "；".join(
        [f"【{cat}】" + "、".join(vals) for cat, vals in keywords.items() if vals]
    )
    return {
        "heat": heat,
        "baokuan": baokuan,
        "stars": stars_of(baokuan),
        "cover": cover,
        "hook_type": hook_type,
        "hook_sentence": hook_sentence,
        "selling_points": selling,
        "keywords": keywords,
        "keywords_flat": kw_flat,
        "long_tail": long_tail,
        "long_tail_flat": "；".join(long_tail),
    }


# ======================================================================
# 8) 关键词聚合（供 Sheet3 使用）
# ======================================================================
def aggregate_keywords(notes, analyses):
    """
    notes: 列表（已含 likes/collects/baokuan）
    analyses: 与 notes 同序的分析结果列表
    返回：[(关键词, 类型, 出现次数, 对应作品数, 平均点赞, 平均收藏, 爆款出现次数), ...]
    按出现次数降序。
    """
    agg = {}  # (kw, type) -> [count, sum_likes, sum_collects, baokuan_count]
    for note, a in zip(notes, analyses):
        likes = note.get("likes") or 0
        collects = note.get("collects") or 0
        is_bao = a["baokuan"] in ("爆款", "潜力爆款")
        for cat, vals in a["keywords"].items():
            for kw in vals:
                key = (kw, cat)
                if key not in agg:
                    agg[key] = [0, 0, 0, 0]
                agg[key][0] += 1
                agg[key][1] += likes
                agg[key][2] += collects
                if is_bao:
                    agg[key][3] += 1
    rows = []
    for (kw, cat), v in agg.items():
        cnt = v[0]
        rows.append((
            kw, cat, cnt, cnt,
            round(v[1] / cnt), round(v[2] / cnt), v[3],
        ))
    rows.sort(key=lambda r: (-r[2], -r[5], r[0]))
    return rows


# ======================================================================
# 9) 爆款规律总结（供 Sheet4 使用）
# ======================================================================
def summarize_rules(notes, analyses):
    # 爆款 / 潜力爆款 视为“高表现”
    top = [(n, a) for n, a in zip(notes, analyses) if a["baokuan"] in ("爆款", "潜力爆款")]
    all_pairs = list(zip(notes, analyses))

    def group_stats(field_fn):
        """按某维度分组，返回 维度->(数量, 平均点赞, 平均收藏)。"""
        g = {}
        for n, a in all_pairs:
            key = field_fn(n, a)
            if not key or key == "待人工确认":
                continue
            g.setdefault(key, [0, 0, 0])
            g[key][0] += 1
            g[key][1] += (n.get("likes") or 0)
            g[key][2] += (n.get("collects") or 0)
        out = []
        for k, v in g.items():
            out.append((k, v[0], round(v[1] / v[0]), round(v[2] / v[0])))
        out.sort(key=lambda x: -x[1])
        return out

    cover_stats = group_stats(lambda n, a: a["cover"])
    hook_stats = group_stats(lambda n, a: a["hook_type"])

    # 高频核心卖点（高表现作品中）
    sp_counter = {}
    for n, a in top:
        for sp in a["selling_points"]:
            sp_counter[sp] = sp_counter.get(sp, 0) + 1
    sp_top = sorted(sp_counter.items(), key=lambda x: -x[1])[:15]

    # 高频关键词（高表现作品中）
    kw_counter = {}
    for n, a in top:
        for cat, vals in a["keywords"].items():
            for kw in vals:
                kw_counter[kw] = kw_counter.get(kw, 0) + 1
    kw_top = sorted(kw_counter.items(), key=lambda x: -x[1])[:20]

    # 爆款共同特征（几条统计性结论）
    features = []
    if top:
        avg_c_l_ratio = sum((n.get("collects") or 0) / max(n.get("likes") or 1, 1) for n, a in top) / len(top)
        cao_ratio = sum(1 for n, a in top if (n.get("collects") or 0) >= (n.get("likes") or 0)) / len(top)
        scene_ratio = sum(1 for n, a in top if a["cover"] in ("客厅整体场景", "生活方式场景", "人物使用场景")) / len(top)
        features.append(f"高表现作品共 {len(top)} 篇。")
        features.append(f"其中收藏/点赞 平均比值 ≈ {avg_c_l_ratio:.2f}（>1 表示更偏“收藏型”内容）。")
        features.append(f"收藏≥点赞 的作品占比 {cao_ratio*100:.0f}%。")
        features.append(f"使用“场景/生活/人物”类封面的占比 {scene_ratio*100:.0f}%。")
        if cover_stats:
            features.append(f"最常见封面类型：{cover_stats[0][0]}（{cover_stats[0][1]} 篇）。")
        if hook_stats:
            features.append(f"最常见文案钩子：{hook_stats[0][0]}（{hook_stats[0][1]} 篇）。")
    else:
        features.append("当前数据中暂无高表现（爆款/潜力爆款）作品，可扩大采集量后重新分析。")

    return {
        "top_count": len(top),
        "cover_stats": cover_stats,
        "hook_stats": hook_stats,
        "selling_top": sp_top,
        "keyword_top": kw_top,
        "features": features,
    }


if __name__ == "__main__":
    import json
    data = json.load(open("xhs_notes.json", encoding="utf-8"))
    heats = [compute_heat(n) for n in data]
    for n in data[:3]:
        a = analyze_note(n, heats)
        print("标题:", n.get("title"))
        print("  热度:", a["heat"], "| 爆款:", a["baokuan"], "| 封面:", a["cover"])
        print("  钩子:", a["hook_type"], "| 句:", a["hook_sentence"])
        print("  卖点:", a["selling_points"])
        print("  关键词:", a["keywords_flat"])
        print()
