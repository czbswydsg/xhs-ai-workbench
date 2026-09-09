"""
小红书种草内容生成模块（离线启发式版，不依赖外部 API）

输入：分析后的主库 xhs_master.json（含每篇作品的 analysis）
输出：xhs_content.xlsx，含两张 Sheet：
      1) 爆款笔记模型：封面规律 / 选题方向 / 标题公式 / 正文结构 / 高频标签
      2) 种草笔记：自动生成 3~5 篇小红书笔记（多标题 / 正文 / 封面方案 / 配图建议 / 标签）

「品牌模式」：
  python generate_content.py --brand 法兰丝
  python generate_content.py --brand 芝华仕 --product 奶油风布艺沙发
  读取 brand_profile.json 中该品牌的 voice（语气/口头禅/句式/形容词/金句/视觉/规避词），
  据此重写标题、开篇、正文、封面方案、标签，让文案真正带出品牌调性，并规避低价促销口吻。

升级为更自然文案：在 gen_one() 中接入 LLM API（见文件末尾说明），
或直接在对话里让 WorkBuddy 帮你写。

用法：
  python generate_content.py                       # 用默认数据生成 3~5 篇（通用）
  python generate_content.py --n 5                # 指定篇数
  python generate_content.py --brand 芝华仕        # 品牌模式
  python generate_content.py --brand 法兰丝 --product 奶油风布艺沙发  # 品牌×产品
  python generate_content.py --output 我的内容.xlsx             # 指定输出文件
"""
import argparse
import json
import os
import random

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from analyzer import analyze_note, compute_heat, summarize_rules

HERE = os.path.dirname(os.path.abspath(__file__))
MASTER_FILE = os.path.join(HERE, "xhs_master.json")
NOTES_FILE = os.path.join(HERE, "xhs_notes.json")
BRAND_FILE = os.path.join(HERE, "brand_profile.json")
CONTENT_FILE = os.path.join(HERE, "xhs_content.xlsx")

HEADER_FILL = PatternFill("solid", fgColor="C00000")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=10)
TITLE_FONT = Font(bold=True, size=14, color="C00000")
WRAP_TOP = Alignment(wrap_text=True, vertical="top")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


# ----------------------------------------------------------------------
# 数据加载：--input 优先（上传/筛选后的分析表），缺失回退主库
# ----------------------------------------------------------------------
def load_records(input_path=None):
    if input_path and os.path.exists(input_path):
        src = input_path
    else:
        src = MASTER_FILE if os.path.exists(MASTER_FILE) else NOTES_FILE
    if not os.path.exists(src):
        raise SystemExit("未找到分析数据，请先采集/上传生成分析表，或运行 build_report.py")
    data = json.load(open(src, encoding="utf-8"))
    # 主库以 {笔记ID: 记录} 形式存储，转成列表
    records = list(data.values()) if isinstance(data, dict) else data
    # 补全 analysis
    heats = [compute_heat(r) for r in records]
    for r, h in zip(records, heats):
        r.setdefault("analysis", None)
        if not r.get("analysis"):
            r["analysis"] = analyze_note(r, heats)
    return records


def load_brand(name):
    if not name:
        return None
    prof = {}
    if os.path.exists(BRAND_FILE):
        try:
            prof = json.load(open(BRAND_FILE, encoding="utf-8"))
        except Exception:
            prof = {}
    return prof.get(name)


# ----------------------------------------------------------------------
# 爆款笔记模型：从分析结果归纳
# ----------------------------------------------------------------------
def build_content_model(records):
    analyses = [r["analysis"] for r in records if r.get("analysis")]
    s = summarize_rules(records, analyses)

    # 高频词（产品/风格/场景/需求）
    prod, style, scene, need = [], [], [], []
    for r in records:
        kw = (r.get("analysis") or {}).get("keywords", {})
        prod += kw.get("产品关键词", [])
        style += kw.get("风格关键词", [])
        scene += kw.get("场景关键词", [])
        need += kw.get("用户需求关键词", [])
    # 取高频（去重保序）
    def top(lst, n=4):
        seen, out = set(), []
        for x in lst:
            if x not in seen:
                seen.add(x); out.append(x)
        return out[:n]
    prod, style, scene, need = top(prod), top(style), top(scene), top(need)

    cover_top = [it[0] for it in s["cover_stats"][:3]]
    hook_top = [it[0] for it in s["hook_stats"][:3]]

    # 选题方向：由高频关键词组合而成
    topics = []
    if scene and prod:
        topics.append(f"{scene[0]}优化")
    if style:
        topics.append(f"{style[0]}空间打造")
    topics += ["家装改造", "氛围感营造", "高级感软装", "小户型显大"]

    # 标题公式（带占位，便于套用）
    title_formulas = [
        "终于找到适合【场景】的【产品】了",
        "被问疯了的【风格】【产品】，谁懂啊",
        "入住一年后，我最庆幸买的【产品】",
        "抄作业！【风格】【场景】【产品】这么搭真的绝",
        "别再乱买了！【产品】看这篇就够了",
        "一眼心动的【风格】【产品】，氛围感拉满",
    ]
    # 填几个示例
    p = prod[0] if prod else "家具"
    st = style[0] if style else ""
    sc = scene[0] if scene else ""
    title_examples = []
    for f in title_formulas[:4]:
        ex = f.replace("【产品】", p).replace("【风格】", st).replace("【场景】", sc)
        title_examples.append(ex)

    # 高频标签
    tags = []
    for r in records:
        kw = (r.get("analysis") or {}).get("keywords", {})
        for cat in ("产品关键词", "风格关键词", "热门话题关键词"):
            tags += kw.get(cat, [])
    tags = top(tags, 12)

    return {
        "cover_top": cover_top,
        "hook_top": hook_top,
        "topics": topics[:5],
        "title_formulas": title_formulas,
        "title_examples": title_examples,
        "prod": prod, "style": style, "scene": scene, "need": need,
        "tags": tags,
        "selling_top": [k for k, _ in s["selling_top"][:8]],
        "features": s["features"],
    }


# ----------------------------------------------------------------------
# 占位符填充工具
# ----------------------------------------------------------------------
def fill_ph(text, p, st, sc, adj=""):
    return (text.replace("{p}", p)
                .replace("{st}", st)
                .replace("{sc}", sc)
                .replace("{adj}", adj))


def clean_forbidden(text, avoid):
    """把品牌规避词做温和替换，避免出现违和口吻。"""
    repl = {
        "低价": "高性价比", "便宜": "实在", "促销": "入手", "冲量": "种草",
        "绝绝子": "", "yyds": "", "家人们": "", "爆款": "好物",
        "廉价": "实在", "凑合": "将就", "将就": "将就", "奢侈": "讲究",
        "高冷": "克制", "轻奢": "质感", "土味": "", "花哨": "",
    }
    for w in avoid or []:
        if w in repl and repl[w]:
            text = text.replace(w, repl[w])
        elif w in text:
            # 无替代词则直接删掉该词（如 绝绝子 / 家人们冲）
            text = text.replace(w, "")
    return text


# ----------------------------------------------------------------------
# 生成单篇笔记
# ----------------------------------------------------------------------
# 通用（非品牌）模板
OPENINGS = [
    "装修前我真的做了好多功课，踩过的坑不想你们再踩。",
    "说真的，这件东西改变了我对【scene】的认知。",
    "之前一直没敢下手，直到我真正用了一个月……",
    "被朋友圈问爆了，干脆出一篇笔记统一回答。",
    "如果你也在纠结【product】，这篇可能是你最需要的。",
]
EXPERIENCE_TPL = (
    "先说结论：它最打动我的是【sp1】和【sp2】。"
    "【prod】整体【style】感很强，放在【scene】里特别协调，"
    "不管是自己用还是朋友来家里都很有面子。"
)
ENDINGS = [
    "家不是样板间，但有一点自己的审美和舒服，真的会让人更愿意回家。",
    "生活的质感，往往就藏在这些小事里。",
    "希望你们也能找到那个让自己心动的角落。",
    "慢慢来，把家过成自己喜欢的样子。",
]
COVER_TEXTS = [
    "终于找到我的理想【product】",
    "被问爆的【style】【scene】",
    "入住一年后最庆幸的购入",
    "抄作业！这样搭真的绝",
]
GEN_TITLE_FORMULAS = [
    "终于找到适合{sc}的{p}了",
    "被问疯了的{st}{p}，谁懂啊",
    "入住一年后，我最庆幸买的{p}",
    "抄作业！{st}{sc}{p}这么搭真的绝",
    "别再乱买了！{p}看这篇就够了",
]


def gen_one(theme, m, bp=None, brand_name="", product=None, idx=0):
    """theme: 一条高表现笔记的 record（取关键词/钩子/卖点）。
    bp:        品牌资料 dict（含 voice）；brand_name: 品牌名字符串。
    product:   具体产品名（选项2）。idx: 篇序号，用于轮换开篇/金句。
    """
    a = theme.get("analysis", {})
    # 产品词：选项2 用指定的具体产品覆盖，其余沿用分析表归纳出的词
    prod = [product] if product else (a.get("keywords", {}).get("产品关键词") or m["prod"] or ["家具"])
    style = (a.get("keywords", {}).get("风格关键词") or m["style"] or [""])
    scene = (a.get("keywords", {}).get("场景关键词") or m["scene"] or ["家里"])
    sp = a.get("selling_points") or m["selling_top"] or ["颜值", "舒适"]
    hook = a.get("hook_type", "种草型")

    p = prod[0]; st = style[0]; sc = scene[0]
    sp1 = sp[0] if sp else "颜值"
    sp2 = sp[1] if len(sp) > 1 else "舒适"

    cover_type = a.get("cover", "场景展示型")

    # ============ 品牌模式：用 voice 真正驱动文案 ============
    if bp:
        v = bp.get("voice", {})
        adj_bank = v.get("adjectives", [])
        expr = bp.get("expression", [])
        openings = v.get("openings", [])
        patterns = v.get("title_patterns", [])
        phrases = v.get("phrases", [])
        visual = bp.get("visual", "")
        avoid = bp.get("avoid", [])
        tone_emoji = v.get("emoji", False)
        adj = random.choice(adj_bank) if adj_bank else ""

        # 标题：优先用品牌专属句式；不足时用通用句式兜底并加上品牌名
        titles = []
        for f in patterns:
            t = fill_ph(f, p, st, sc, adj)
            if t and t not in titles:
                titles.append(t)
        # 兜底 + 品牌×产品专属收尾句
        for f in GEN_TITLE_FORMULAS:
            t = fill_ph(f, p, st, sc)
            if t and t not in titles:
                titles.append(t)
            if len(titles) >= 3:
                break
        if product:
            titles.append(f"{brand_name}的{product}，氛围感拉满")
        elif brand_name:
            titles.append(f"{brand_name}这套{st}{p}，真的耐看")
        titles = titles[:4]

        # 开篇：用品牌专属开场（按篇序号轮换）
        if openings:
            opening = openings[idx % len(openings)]
            # 把产品/场景自然嵌进去
            opening = opening.replace("它", p).replace("沙发", p if "沙发" in p else "沙发")
        else:
            opening = fill_ph(random.choice(OPENINGS), p, st, sc).replace("【product】", p).replace("【scene】", sc)

        # 正文：用品牌形容词 + 表达词重写体验段
        adj1 = adj_bank[0] if adj_bank else "舒服"
        expr_word = expr[0] if expr else ""
        experience = (
            f"说点实在的：它最打动我的是{sp1}，{expr_word}这一点是真的加分。"
            f"整体{st}里带着{adj1}的质感，放进{sc}不抢戏却很出彩，"
            f"朋友来家里总忍不住多坐一会儿。"
        )

        # 结尾：品牌金句
        ending = phrases[idx % len(phrases)] if phrases else random.choice(ENDINGS)

        # 封面方案：用品牌视觉调性
        cover_text = fill_ph(random.choice(patterns) if patterns else "{p}真的耐看", p, st, sc, adj)
        if not cover_text or len(cover_text) > 16:
            cover_text = f"我的理想{st}{p}"
        cover_plan = (
            f"图片类型：{cover_type}\n"
            f"封面文字：「{cover_text}」\n"
            f"构图：产品占主体约 70%，留白 30%\n"
            f"色调：{visual or (st + '低饱和、自然光、有质感')}"
        )

        # 标签：品牌名 + 表达词 + 产品/风格/通用
        tags = [f"#{brand_name}"]
        for t in (expr[:2] + prod[:2] + style[:2] + m["tags"][:3]):
            if t and len(t) >= 2:
                tags.append(f"#{t}")
        tags += ["#家居好物", "#装修灵感", "#种草"]
        tags = list(dict.fromkeys(tags))[:10]

        body = f"{opening}\n\n{experience}\n\n{ending}"
        body = clean_forbidden(body, avoid)
        cover_plan = clean_forbidden(cover_plan, avoid)

        return {
            "titles": titles,
            "body": body,
            "cover_plan": cover_plan,
            "images": [
                f"第1张：{sc}整体氛围图（建立场景感）",
                f"第2张：{sp1}细节特写（材质/做工）",
                f"第3张：真实使用场景（人在其中）",
                f"第4张：{st}搭配参考 / 尺寸示意",
                "第5张：一句金句收尾图（引导收藏）",
            ],
            "tags": " ".join(tags),
            "hook": hook,
            "cover_type": cover_type,
            "selling": "、".join(sp[:4]),
        }

    # ============ 通用模式（无品牌） ============
    titles = [fill_ph(f, p, st, sc) for f in GEN_TITLE_FORMULAS[:3]]
    opening = fill_ph(random.choice(OPENINGS), p, st, sc).replace("【product】", p).replace("【scene】", sc)
    experience = EXPERIENCE_TPL.replace("【sp1】", sp1).replace("【sp2】", sp2) \
        .replace("【prod】", p).replace("【style】", st).replace("【scene】", sc)
    ending = random.choice(ENDINGS)
    body = f"{opening}\n\n{experience}\n\n{ending}"

    cover_text = fill_ph(random.choice(COVER_TEXTS), p, st, sc)
    cover_plan = (
        f"图片类型：{cover_type}\n"
        f"封面文字：「{cover_text}」\n"
        f"构图：产品占主体约 70%，留白 30%\n"
        f"色调：{st + '低饱和' if st else '低饱和'}、自然光、有质感"
    )

    tags = [f"#{t}" for t in (prod[:2] + style[:2] + m["tags"][:3]) if t and len(t) >= 2]
    tags += ["#家居好物", "#装修灵感", "#种草"]
    tags = list(dict.fromkeys(tags))[:10]

    return {
        "titles": titles,
        "body": body,
        "cover_plan": cover_plan,
        "images": [
            f"第1张：{sc}整体氛围图（建立场景感）",
            f"第2张：{sp1}细节特写（材质/做工）",
            f"第3张：真实使用场景（人在其中）",
            f"第4张：{st}搭配参考 / 尺寸示意",
            "第5张：一句金句收尾图（引导收藏）",
        ],
        "tags": " ".join(tags),
        "hook": hook,
        "cover_type": cover_type,
        "selling": "、".join(sp[:4]),
    }


def gen_notes(records, n=5, bp=None, brand_name="", product=None):
    analyses = [(r, r["analysis"]) for r in records if r.get("analysis")]
    # 优先用高表现作品作为选题素材（封面/钩子/卖点规律来自分析表）
    top = [r for r, a in analyses if a.get("baokuan") in ("爆款", "潜力爆款")]
    pool = top if top else [r for r, _ in analyses]
    # 按热度取前 n 篇作为主题
    pool = sorted(pool, key=lambda r: r["analysis"]["heat"], reverse=True)[:n]
    if not pool:
        pool = records[:n]
    model = build_content_model(records)
    out = []
    for i, r in enumerate(pool):
        out.append(gen_one(r, model, bp, brand_name, product, idx=i))
    return out


# ----------------------------------------------------------------------
# 写出 Excel
# ----------------------------------------------------------------------
def write_content_xlsx(model, notes, bp=None, brand_name="", mode="通用", out_file=CONTENT_FILE):
    wb = Workbook()
    wb.remove(wb.active)  # 删除默认 Sheet

    # ---- Sheet 1: 爆款笔记模型 ----
    ws = wb.create_sheet("爆款笔记模型")
    ws.append([f"📌 爆款笔记模型（由数据分析表自动归纳 · 生成模式：{mode}）"])
    ws["A1"].font = TITLE_FONT
    r = 3

    def block(title, lines):
        nonlocal r
        ws.cell(row=r, column=1, value=title).font = Font(bold=True, size=11, color="C00000")
        r += 1
        for ln in lines:
            ws.cell(row=r, column=1, value="• " + ln).alignment = WRAP_TOP
            r += 1
        r += 1

    block("一、爆款封面规律", [
        f"高频封面类型：{('、'.join(model['cover_top']) or '待积累')}",
        "建议产品占画面主体约 70%，留白 30%，自然光 + 低饱和色调更显高级",
        "场景图 / 人物使用场景 的收藏表现通常优于单纯产品白底图",
    ])
    block("二、爆款选题方向", [f"· {t}" for t in model["topics"]])
    block("三、标题公式（可直接套用）", model["title_formulas"] + ["", "示例："] + model["title_examples"])
    block("四、正文结构", [
        "开头：情绪 / 痛点切入，先共情",
        "中间：产品真实体验，落到 1~2 个核心卖点",
        "结尾：生活方式表达，引导收藏 / 关注",
    ])
    block("五、高频标签关键词", ["、".join(model["tags"]) or "（暂无）"] + [
        f"产品词：{('、'.join(model['prod']) or '—')}",
        f"风格词：{('、'.join(model['style']) or '—')}",
        f"场景词：{('、'.join(model['scene']) or '—')}",
    ])

    # 品牌调性说明（品牌模式时显示，帮助判断文案是否对味）
    if bp:
        v = bp.get("voice", {})
        block("六、本次品牌调性（文案据此生成）", [
            f"品牌：{brand_name}",
            f"定位：{bp.get('positioning', '—')}",
            f"目标人群：{bp.get('audience', '—')}",
            f"说话语气：{v.get('tone', '—')}",
            f"视觉调性：{bp.get('visual', '—')}",
            f"常用表达：{('、'.join(bp.get('expression', [])) or '—')}",
            f"规避口吻：{('、'.join(bp.get('avoid', [])) or '—')}",
        ])

    ws.column_dimensions["A"].width = 90
    ws.sheet_view.showGridLines = False

    # ---- Sheet 2: 种草笔记 ----
    ws2 = wb.create_sheet("种草笔记")
    headers = ["篇号", "标题方案(多选)", "正文", "封面设计方案", "配图建议", "标签", "参考钩子", "封面类型", "核心卖点"]
    ws2.append(headers)
    for i, nt in enumerate(notes, 1):
        ws2.append([
            i,
            "\n".join(f"{j}. {t}" for j, t in enumerate(nt["titles"], 1)),
            nt["body"],
            nt["cover_plan"],
            "\n".join(nt["images"]),
            nt["tags"],
            nt["hook"],
            nt["cover_type"],
            nt["selling"],
        ])
    # 样式
    for c in range(1, len(headers) + 1):
        cell = ws2.cell(row=1, column=c)
        cell.fill = HEADER_FILL; cell.font = HEADER_FONT; cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in range(2, len(notes) + 2):
        for col in (2, 3, 4, 5, 6):
            ws2.cell(row=row, column=col).alignment = WRAP_TOP
        for col in (1, 7, 8, 9):
            ws2.cell(row=row, column=col).alignment = Alignment(horizontal="center", vertical="top")
    ws2.freeze_panes = "A2"
    # 列宽
    for i, w in enumerate([5, 34, 60, 40, 34, 30, 12, 14, 18], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    wb.save(out_file)
    return out_file


# ----------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--n", type=int, default=5, help="生成篇数(3~5)")
    ap.add_argument("--brand", default="", help="品牌模式：品牌名（需 brand_profile.json 含该品牌）")
    ap.add_argument("--product", default="", help="具体产品名（如：奶油风布艺沙发）；与 --brand 配合生成品牌×产品内容")
    ap.add_argument("--input", default="", help="指定已分析数据 JSON（上传/筛选后的分析表），覆盖主库")
    ap.add_argument("--output", default="", help="指定输出 xlsx 路径（默认覆盖 xhs_content.xlsx）")
    args = ap.parse_args()

    records = load_records(args.input or None)
    brand_name = args.brand.strip()
    bp = load_brand(brand_name) if brand_name else None
    if brand_name and not bp:
        print(f"提示：未在 brand_profile.json 找到品牌「{brand_name}」，将用通用调性生成。"
              f"你可编辑 brand_profile.json 补充该品牌的 voice（语气/口头禅/句式/表达）。")
    notes = gen_notes(records, n=max(3, min(5, args.n)),
                      bp=bp, brand_name=brand_name, product=args.product.strip() or None)
    mode = "通用" if not brand_name else (f"品牌×产品（{brand_name}·{args.product}）" if args.product.strip() else f"品牌（{brand_name}）")
    out = args.output.strip() or CONTENT_FILE
    write_content_xlsx(records and build_content_model(records), notes, bp, brand_name, mode, out)
    print(f"已生成 {len(notes)} 篇种草笔记 [{mode}]：{out}")


if __name__ == "__main__":
    main()

# ======================================================================
# 【升级为更自然文案】说明
#   当前为离线模板生成，文案偏结构化。要更接近真人写作，可：
#   1) 在 gen_one() 中调用 LLM API（如 OpenAI / 通义 / 文心），把 theme + model
#      作为 prompt 输入，让模型返回 标题/正文/封面方案/标签；
#   2) 或直接把「爆款笔记模型 + 你想写的品牌/产品」发给我（WorkBuddy），
#      我现场帮你写出可直接发布的小红书笔记。
# ======================================================================
