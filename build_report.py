"""
小红书内容数据采集 -> Excel 爆款分析表 生成器

工作流：
  1. 读取采集器产出的 xhs_notes.json（每次采集覆盖）
  2. 与历史主库 xhs_master.json 合并去重（按 笔记ID）；
     - 旧作品：更新点赞/收藏等指标，保留你手动修正过的分析字段
     - 新作品：自动分析（封面/钩子/卖点/关键词）
  3. 重算 综合热度 与 爆款标记（随指标动态变化）
  4. 生成 xhs_report.xlsx，含 5 个 Sheet：
       原始数据 / 爆款分析 / 关键词分析 / 爆款规律总结 / 使用说明

用法：
  python build_report.py                # 常规刷新（保留手动修正）
  python build_report.py --force        # 强制重新分析全部作品
  python build_report.py --input xxx.json --output yyy.xlsx
"""
import argparse
import json
import os
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from analyzer import (
    analyze_note, classify_baokuan, compute_heat,
    aggregate_keywords, summarize_rules, WEIGHTS,
    stars_of, generate_long_tail,
)

HERE = os.path.dirname(os.path.abspath(__file__))
NOTES_FILE = os.path.join(HERE, "xhs_notes.json")
MASTER_FILE = os.path.join(HERE, "xhs_master.json")
REPORT_FILE = os.path.join(HERE, "xhs_report.xlsx")

# ---------- 样式 ----------
HEADER_FILL = PatternFill("solid", fgColor="C00000")   # 小红书红
HEADER_FONT = Font(bold=True, color="FFFFFF", size=10)
TITLE_FONT = Font(bold=True, size=14, color="C00000")
SUB_FONT = Font(bold=True, size=11, color="333333")
WRAP_TOP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
SEC_FILL = PatternFill("solid", fgColor="F2F2F2")


# ======================================================================
# 合并 / 去重 / 分析
# ======================================================================
def load_master():
    if os.path.exists(MASTER_FILE):
        try:
            return json.load(open(MASTER_FILE, encoding="utf-8"))
        except Exception:
            return {}
    return {}


def normalize_note(n):
    return {
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
    }


def merge_and_analyze(force=False, use_master=True):
    master = load_master() if use_master else {}
    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    new_notes = []
    if os.path.exists(NOTES_FILE):
        try:
            new_notes = json.load(open(NOTES_FILE, encoding="utf-8"))
        except Exception as e:
            print("读取 xhs_notes.json 失败：", e)
            new_notes = []

    merged = {}  # id -> record (完整 schema)
    for n in new_notes:
        rec = normalize_note(n)
        rid = rec["id"]
        if not rid:
            continue
        if rid in master:
            old = master[rid]
            rec["collect_time"] = old.get("collect_time", now)
            # 保留手写修正过的分析字段；仅当 force 时重新分析
            if force:
                rec["analysis"] = None
            else:
                rec["analysis"] = dict(old.get("analysis", {}))
        else:
            rec["collect_time"] = now
            rec["analysis"] = None
        merged[rid] = rec

    # 保留主库里、本次未重新采集到的旧作品（累计多关键词采集）
    for rid, old in master.items():
        if rid not in merged:
            merged[rid] = old

    records = list(merged.values())
    all_heats = [compute_heat(r) for r in records]

    for r, h in zip(records, all_heats):
        a = r.get("analysis")
        if not a:  # 新作品 或 force
            a = analyze_note(r, all_heats)
            r["analysis"] = a
        else:
            # 旧作品：动态指标更新，文本分析保留
            a["heat"] = h
            a["baokuan"] = classify_baokuan(h, all_heats)
        # 补齐新版本引入的派生字段（兼容旧主库里缺字段的分析）
        a["stars"] = stars_of(a.get("baokuan", ""))
        if "keywords" in a:
            a["long_tail"] = generate_long_tail(r, a["keywords"])
            a["long_tail_flat"] = "；".join(a["long_tail"])
        else:
            a["long_tail"] = []
            a["long_tail_flat"] = ""

    # 回写主库
    if use_master:
        json.dump(merged, open(MASTER_FILE, "w", encoding="utf-8"),
                  ensure_ascii=False, indent=2)
    return records


# ======================================================================
# 写入工具
# ======================================================================
def style_header(ws, ncols, row=1):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = BORDER


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def autofit_columns(ws, max_w=50, min_w=6):
    """按单元格实际内容自适应列宽（CJK 按 2 字符宽估算），避免无效空白。"""
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            col = cell.column_letter
            s = str(cell.value)
            if s.startswith("="):  # 公式列按数字宽度处理，避免公式文本撑宽
                w = 9
            else:
                w = 0
                for line in s.split("\n"):
                    lw = sum(2 if ord(ch) > 0x2E80 else 1 for ch in line)
                    w = max(w, lw)
            widths[col] = max(widths.get(col, 0), w)
    for col, w in widths.items():
        ws.column_dimensions[col].width = max(min_w, min(max_w, w + 2))


def add_table(ws, name, ncols, nrows):
    if nrows < 1:
        return
    ref = f"A1:{get_column_letter(ncols)}{nrows + 1}"
    tbl = Table(displayName=name, ref=ref)
    tbl.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False)
    ws.add_table(tbl)


# ======================================================================
# Sheet 1：原始数据
# ======================================================================
RAW_HEADERS = ["序号", "笔记ID", "笔记链接", "笔记标题", "作者", "发布时间",
               "点赞量", "收藏量", "评论量", "分享量", "粉丝量", "采集时间",
               "综合热度", "爆款标记", "封面类型", "文案钩子", "核心卖点",
               "热搜关键词", "长尾关键词", "封面链接", "正文内容", "钩子句"]


def build_raw_sheet(wb, records):
    ws = wb.create_sheet("原始数据")
    ws.append(RAW_HEADERS)
    for i, r in enumerate(records, 1):
        a = r["analysis"]
        ws.append([
            i, r["id"], r["url"], r["title"], r["author"], r["publish_time"],
            r["likes"], r["collects"], r["comments"], r["shares"], r["fans"],
            r["collect_time"],
            f"=[@点赞量]*{WEIGHTS['likes']}+[@收藏量]*{WEIGHTS['collects']}"
            f"+[@评论量]*{WEIGHTS['comments']}+[@分享量]*{WEIGHTS['shares']}",
            f"{a['baokuan']} {a['stars']}", a["cover"], a["hook_type"],
            "、".join(a["selling_points"]), a["keywords_flat"], a["long_tail_flat"],
            r["cover"], r["content"], a["hook_sentence"],
        ])
    nrows = len(records)
    style_header(ws, len(RAW_HEADERS))
    autofit_columns(ws)
    # 综合热度为公式列，整列设为常规
    for row in range(2, nrows + 2):
        for col in (3, 4, 18, 19, 20, 21, 22):  # 长文本列
            ws.cell(row=row, column=col).alignment = WRAP_TOP
        for col in (7, 8, 9, 10, 11, 13):  # 数值列
            ws.cell(row=row, column=col).alignment = CENTER
    ws.freeze_panes = "A2"
    if nrows:
        add_table(ws, "原始数据", len(RAW_HEADERS), nrows)
    return ws


# ======================================================================
# Sheet 2：爆款分析（整行排序，杜绝错位）
# ======================================================================
ANL_HEADERS = ["排名", "标题", "链接", "作者", "发布时间", "点赞", "收藏", "评论",
               "分享", "综合热度", "点赞排名", "收藏排名", "爆款标记", "封面类型",
               "文案钩子", "核心卖点", "热搜关键词", "长尾关键词"]


def build_analysis_sheet(wb, records):
    ws = wb.create_sheet("爆款分析")
    ws.append(ANL_HEADERS)

    # 综合热度排序（整行一起移动，绝不会错位）
    sorted_recs = sorted(records, key=lambda r: r["analysis"]["heat"], reverse=True)
    # 计算维度排名
    by_likes = sorted(records, key=lambda r: r["likes"], reverse=True)
    by_collects = sorted(records, key=lambda r: r["collects"], reverse=True)
    like_rank = {id(r): i + 1 for i, r in enumerate(by_likes)}
    collect_rank = {id(r): i + 1 for i, r in enumerate(by_collects)}

    for rank, r in enumerate(sorted_recs, 1):
        a = r["analysis"]
        ws.append([
            rank, r["title"], r["url"], r["author"], r["publish_time"],
            r["likes"], r["collects"], r["comments"], r["shares"], a["heat"],
            like_rank[id(r)], collect_rank[id(r)], f"{a['baokuan']} {a['stars']}",
            a["cover"], a["hook_type"], "、".join(a["selling_points"]),
            a["keywords_flat"], a["long_tail_flat"],
        ])
        # 链接做成可点击超链接
        link_cell = ws.cell(row=rank + 1, column=3)
        if r["url"]:
            link_cell.hyperlink = r["url"]
            link_cell.font = Font(color="0563C1", underline="single", size=9)

    nrows = len(sorted_recs)
    style_header(ws, len(ANL_HEADERS))
    autofit_columns(ws)
    for row in range(2, nrows + 2):
        for col in (2, 17, 18):
            ws.cell(row=row, column=col).alignment = WRAP_TOP
        for col in (1, 6, 7, 8, 9, 10, 11, 12):
            ws.cell(row=row, column=col).alignment = CENTER
    ws.freeze_panes = "A2"
    if nrows:
        add_table(ws, "爆款分析", len(ANL_HEADERS), nrows)
    return ws


# ======================================================================
# Sheet 3：关键词分析
# ======================================================================
def build_keyword_sheet(wb, records):
    ws = wb.create_sheet("关键词分析")
    headers = ["关键词", "类型", "出现次数", "对应作品数", "平均点赞", "平均收藏", "爆款出现次数"]
    ws.append(headers)
    analyses = [r["analysis"] for r in records]
    rows = aggregate_keywords(records, analyses)
    for row in rows:
        ws.append(list(row))
    nrows = len(rows)
    style_header(ws, len(headers))
    autofit_columns(ws)
    for r in range(2, nrows + 2):
        for c in (3, 4, 5, 6, 7):
            ws.cell(row=r, column=c).alignment = CENTER
    ws.freeze_panes = "A2"
    if nrows:
        add_table(ws, "关键词分析", len(headers), nrows)
    ws.cell(row=nrows + 3, column=1,
            value="说明：本表由工具自动统计，重新运行 build_report.py 后自动刷新。").font = Font(italic=True, color="808080", size=9)
    return ws


# ======================================================================
# Sheet 4：爆款规律总结
# ======================================================================
def build_summary_sheet(wb, records):
    ws = wb.create_sheet("爆款规律总结")
    analyses = [r["analysis"] for r in records]
    s = summarize_rules(records, analyses)

    r = 1
    ws.cell(row=r, column=1, value="📊 爆款规律总结").font = TITLE_FONT
    r += 2

    # 公式与规则说明
    ws.cell(row=r, column=1, value="一、综合热度公式与爆款判定规则").font = SUB_FONT
    r += 1
    rules = [
        f"综合热度 = 点赞×{WEIGHTS['likes']} + 收藏×{WEIGHTS['collects']} + 评论×{WEIGHTS['comments']} + 分享×{WEIGHTS['shares']}",
        "（收藏=强意向、评论=深互动、分享=破圈，权重依次提高，更贴近“爆款”本质）",
        "爆款标记按“综合热度在全部作品中的百分排位”自动判定：",
        "   爆款：前 10%（且最高热度≥200）   潜力爆款：10%~30%",
        "   普通：30%~70%   低表现：后 30%",
        "→ 全部随数据采集动态变化，不写死固定阈值。",
    ]
    for line in rules:
        ws.cell(row=r, column=1, value=line).alignment = WRAP_TOP
        r += 1
    r += 1

    def write_block(title, header, data_rows, widths):
        nonlocal r
        ws.cell(row=r, column=1, value=title).font = SUB_FONT
        r += 1
        for c, h in enumerate(header, 1):
            cell = ws.cell(row=r, column=c, value=h)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = CENTER
        r += 1
        for drow in data_rows:
            for c, val in enumerate(drow, 1):
                cell = ws.cell(row=r, column=c, value=val)
                cell.border = BORDER
                if c > 1 and isinstance(val, (int, float)):
                    cell.alignment = CENTER
            r += 1
        r += 1

    # 爆款常见封面类型
    write_block("二、爆款常见封面类型（高表现作品）",
                ["封面类型", "篇数", "平均点赞", "平均收藏"],
                [[it[0], it[1], it[2], it[3]] for it in s["cover_stats"][:10]],
                [16, 8, 10, 10])
    # 爆款常见文案钩子
    write_block("三、爆款常见文案钩子（高表现作品）",
                ["文案钩子", "篇数", "平均点赞", "平均收藏"],
                [[it[0], it[1], it[2], it[3]] for it in s["hook_stats"][:10]],
                [16, 8, 10, 10])
    # 高频核心卖点
    write_block("四、高频核心卖点（高表现作品 TOP）",
                ["核心卖点", "出现篇数"],
                [[k, v] for k, v in s["selling_top"][:15]],
                [20, 10])
    # 高频关键词
    write_block("五、高频关键词（高表现作品 TOP）",
                ["关键词", "出现篇数"],
                [[k, v] for k, v in s["keyword_top"][:20]],
                [24, 10])
    # 不同封面类型平均表现（全部作品）
    write_block("六、不同封面类型的平均表现（全部作品）",
                ["封面类型", "篇数", "平均点赞", "平均收藏"],
                [[it[0], it[1], it[2], it[3]] for it in s["cover_stats"]],
                [16, 8, 10, 10])
    # 不同文案钩子平均表现
    write_block("七、不同文案钩子的平均表现（全部作品）",
                ["文案钩子", "篇数", "平均点赞", "平均收藏"],
                [[it[0], it[1], it[2], it[3]] for it in s["hook_stats"]],
                [16, 8, 10, 10])
    # 爆款共同特征
    ws.cell(row=r, column=1, value="八、爆款内容共同特征").font = SUB_FONT
    r += 1
    for feat in s["features"]:
        ws.cell(row=r, column=1, value="• " + feat).alignment = WRAP_TOP
        r += 1

    set_widths(ws, [26, 14, 12, 12, 12, 12])
    ws.sheet_view.showGridLines = False
    return ws


# ======================================================================
# Sheet 5：使用说明
# ======================================================================
def build_help_sheet(wb):
    ws = wb.create_sheet("使用说明")
    lines = [
        ("📕 小红书内容采集 + 爆款分析表 · 使用说明", "title"),
        ("", ""),
        ("【整体工作流】", "sub"),
        ("采集数据 → 导入/合并 → 自动更新指标 → 自动排序 → 自动判定爆款 → 自动分析封面/钩子/卖点/关键词 → 总结爆款规律", ""),
        ("", ""),
        ("【如何刷新这份表】", "sub"),
        ("方式一（推荐）：双击项目目录里的 run_analysis.bat", ""),
        ("方式二：在命令行执行  python build_report.py", ""),
        ("如需重新分析全部作品（忽略手动修正）：python build_report.py --force", ""),
        ("", ""),
        ("【数据从哪来】", "sub"),
        ("1) 先用采集工具（streamlit run app.py）采集，结果存为 xhs_notes.json", ""),
        ("2) build_report.py 会把每次采集合并进 xhs_master.json（按笔记ID去重）", ""),
        ("3) 同一篇笔记被重复采集 → 自动更新点赞/收藏，不会产生重复行", ""),
        ("", ""),
        ("【各 Sheet 用途】", "sub"),
        ("原始数据：每条笔记的完整记录 + 分析字段；综合热度是 Excel 公式，修改点赞/收藏会自动重算", ""),
        ("爆款分析：已按综合热度整行排序，可直接看排名；点列头可按其它维度排序（整表一起动，不会错位）", ""),
        ("关键词分析：自动统计高频关键词及平均点赞/收藏/爆款出现次数", ""),
        ("爆款规律总结：自动归纳爆款封面/钩子/卖点/关键词与共同特征", ""),
        ("", ""),
        ("【如何排序而不错位】", "sub"),
        ("「原始数据」「爆款分析」「关键词分析」都是 Excel 表格（点表头三角可排序/筛选）", ""),
        ("排序时是整行一起移动，标题/点赞/收藏/爆款/封面/钩子等全部跟着走，不会错位", ""),
        ("", ""),
        ("【种草内容生成】", "sub"),
        ("运行 python generate_content.py 可基于分析结果，自动生成 3~5 篇小红书种草笔记", ""),
        ("（含多标题、正文、封面方案、配图建议、标签）；输出 xhs_content.xlsx", ""),
        ("如需更自然的文案：在 generate_content.py 中接入 LLM API，或直接在对话里让我帮你写", ""),
        ("", ""),
        ("【字段含义】", "sub"),
        ("综合热度：点赞×1+收藏×1.5+评论×2+分享×3 的加权互动值", ""),
        ("爆款标记：爆款(★★★★★)/潜力爆款(★★★★)/普通(★★)/低表现(★)，按热度百分排位自动判定", ""),
        ("封面类型：由标题/正文启发式判断；无法判断时标「待人工确认」", ""),
        ("文案钩子：痛点/好奇/情绪/反差/清单/场景/种草/经验/提问 等类型 + 关键钩子句", ""),
        ("核心卖点：仅从原帖文本提取，不编造", ""),
        ("热搜关键词：分 产品/风格/场景/用户需求/热门话题 五类", ""),
        ("长尾关键词：由上述关键词组合成的搜索短语（如「小户型适合什么沙发」），用于 SEO/投流", ""),
        ("", ""),
        ("【你可以手动改什么】", "sub"),
        ("可在「原始数据」表里直接修改 封面类型 / 文案钩子 / 核心卖点 / 热搜关键词", ""),
        ("→ 下次刷新时会保留你的手动修正（除非用 --force）", ""),
        ("", ""),
        ("【注意事项】", "sub"),
        ("· 小红书前端会改版，详情页选择器失效会导致 收藏/评论/分享/粉丝 采不到（显示0），届时告诉我更新选择器", ""),
        ("· 封面类型基于文本判断，非图像识别；如要更准确可人工复核「待人工确认」项", ""),
        ("· 分享量/粉丝量需详情页采集，采集时请保持「采集详情」勾选", ""),
    ]
    r = 1
    for text, kind in lines:
        cell = ws.cell(row=r, column=1, value=text)
        if kind == "title":
            cell.font = TITLE_FONT
        elif kind == "sub":
            cell.font = SUB_FONT
        else:
            cell.alignment = WRAP_TOP
            cell.font = Font(size=10)
        r += 1
    ws.column_dimensions["A"].width = 110
    ws.sheet_view.showGridLines = False
    return ws


# ======================================================================
# 主流程
# ======================================================================
def main():
    global NOTES_FILE
    ap = argparse.ArgumentParser()
    ap.add_argument("--force", action="store_true", help="强制重新分析全部作品")
    ap.add_argument("--input", default=NOTES_FILE, help="采集结果 JSON")
    ap.add_argument("--output", default=REPORT_FILE, help="输出 Excel")
    ap.add_argument("--nomaster", action="store_true",
                    help="不读取/写入主库，仅分析本次输入的 JSON（用于上传文件隔离分析）")
    ap.add_argument("--emit-json", default="",
                    help="把分析后的记录（含 analysis）写出到此 JSON，供内容生成模块复用")
    args = ap.parse_args()

    NOTES_FILE = args.input

    records = merge_and_analyze(force=args.force, use_master=not args.nomaster)
    print(f"合并后共 {len(records)} 篇作品")

    wb = Workbook()
    wb.remove(wb.active)  # 删掉默认 sheet
    build_raw_sheet(wb, records)
    build_analysis_sheet(wb, records)
    build_keyword_sheet(wb, records)
    build_summary_sheet(wb, records)
    build_help_sheet(wb)

    wb.save(args.output)
    print(f"已生成报表：{args.output}")

    if args.emit_json:
        json.dump({r["id"]: r for r in records}, open(args.emit_json, "w", encoding="utf-8"),
                  ensure_ascii=False, indent=2)
        print(f"已导出分析数据：{args.emit_json}")


if __name__ == "__main__":
    main()
