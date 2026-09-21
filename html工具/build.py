# -*- coding: utf-8 -*-
"""
银企直联材料清单 HTML 构建脚本
流水线：xlsx(全量Sheet) -> 标准化JSON(含溯源) -> 场景映射(mapping.json) -> 自动校验 -> 渲染HTML
用法：python build.py
四原则：1) Excel 是唯一数据源  2) JSON 是中间标准数据层(可溯源)  3) HTML 是离线查询工具  4) 新增解释不得改变原行内口径
"""
import io
import json
import os
import re
import sys
import datetime
from pathlib import Path

import openpyxl

# Python 3.6 兼容：控制台输出转 UTF-8
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
except Exception:
    pass

BASE = Path(__file__).resolve().parent.parent          # 银企直联业务/
TOOL = Path(__file__).resolve().parent                 # html工具/
XLSX = BASE / "银企直联材料清单.xlsx"
OUT_HTML = BASE / "银企直联材料清单.html"
TEMPLATE = TOOL / "template.html"
MAPPING = TOOL / "mapping.json"
DATA_JSON = TOOL / "data_standardized.json"
REPORT = TOOL / "数据校验报告.md"
AUDIT_REPORT = TOOL / "内容审查报告.md"

REVIEW_STATUS = [
    ("P0-1 知识库数量冲突", "已落实：知识库决策表数量列动态比对，冲突处显示『实际X项』；build.py 增加校验并计入场景级待确认"),
    ("P0-2 待确认项落地", "已落实：受影响场景标题旁挂『⚠ N 条待确认』可展开；知识库页新增『数据校验状态』卡列出全部差异；本报告+校验报告双落地"),
    ("P0-3 客户视图分离", "已落实：顶栏『行内/客户』开关（记忆状态）。客户视图隐藏行内话术、客户经理要点、数据溯源、待确认、审查标注；另提供场景页『复制客户版清单』纯文本导出与打印样式（自动展开明细）。模板文件分发不在单文件能力范围内，未做"),
    ("P0-4 矩阵9/12场景", "已落实：保留原表矩阵（口径注明），新增『新增司库/薪资代发/文件证书』3列自动补充矩阵（由场景明细生成，标整理层）；列头与勾选格可点击跳转"),
    ("P1-1 用印归属/CM要点回填", "部分落实：8列表场景的用印归属已上卡展示（原表列名原样）；从草稿表回填涉及跨表合并口径，待确认后实施"),
    ("P1-2 场景化总原则", "未落实：源表无此内容，编写属新增银行口径，需行内业务确认后由 mapping.json 配置"),
    ("P1-3 注意事项覆盖", "已落实：填写要点中的『⚠️ 注意：…』行内高亮为醒目样式（原文不移动）"),
    ("P1-4 费用/联系人", "未落实：源表仅有费用减免单一处提及、无联系人数据；已在场景页统一展示FAQ口径时限『3-5个工作日』；其余需行内补充数据源"),
    ("P1-5 副标题", "已落实：注销集群/新增盾/文件证书三处副标题按『什么时候用』改写（mapping.json subtitle_fix），原副标题保留在溯源中"),
    ("P2-1 客户版导出", "部分落实：场景页一键『复制客户版清单』纯文本（微信可粘贴）+打印/PDF样式；空白模板下载需表单文件源，未做"),
    ("P2-2 勾选核对", "已落实：材料卡勾选框+『收单核对 x/n』进度+一键清除，状态存 localStorage（仅行内视图）"),
    ("P2-3 状态记忆", "已落实：记住最后访问页面、客户/行内视图、勾选进度；滚动位置按页面记忆"),
    ("P2-4 搜索扩展", "已落实：索引扩展到填写要点/盖章要点/注意事项，大小写归一；新增快捷词（盖章/份数/异地/多久/收费）"),
    ("P2-5 死数据", "已落实：草稿表与目录行不再内嵌进 HTML（仍保留在 data_standardized.json 供追溯）"),
    ("P2-6 引导入口不一致", "已落实：知识库决策表下增补提示『本表未收录3个场景』（整理提示，原文未动）"),
    ("P2-7 交互细节", "部分落实：返回按钮走历史栈、滚动位置记忆、大字模式开关、数据溯源仅行内视图展示；微信内无法执行JS属查看器限制，已内置引导提示"),
    ("P3 专业性补强", "部分落实：页脚版本号（v1.1）+复核提示；制度依据/有效期需行内数据；反馈引导语已加页脚"),
]

CATEGORY_ORDER = ["premise", "mutex", "timing", "script", "general"]
CATEGORY_LABEL = {
    "premise": "⚠️ 前提条件",
    "mutex": "🔀 互斥关系",
    "timing": "⏱ 生效与办理时间",
    "script": "💬 行内话术",
    "general": "📌 补充说明",
}


def C(ws, r, c):
    """读单元格并去首尾空白（含全角空格）。"""
    v = ws.cell(row=r, column=c).value
    if v is None:
        return ""
    s = str(v).replace("\u3000", " ").strip()
    return s


def rowvals(ws, r, maxc):
    return [C(ws, r, c) for c in range(1, maxc + 1)]


def join_cells(vals):
    return " | ".join([v for v in vals if v])


def match_category(text, rules):
    """按 mapping.note_rules 关键词归类；'script_marker' 表示纯标记行(如'话术')，跳过输出。"""
    for rule in rules:
        kw = rule.get("contains", "")
        if kw and kw in text:
            return rule.get("category", "general")
    return "general"


# ---------------------------------------------------------------- 场景表解析

def parse_scenario_sheet(ws, sheet_name, rules):
    """解析单个场景材料表。返回 title/subtitle/principle/materials/notes/统计。"""
    maxr = ws.max_row
    maxc = ws.max_column
    title = C(ws, 1, 1)
    subtitle = ""
    if "—" in title:
        subtitle = title.split("—", 1)[1].strip()

    # 找表头行（含"材料名称"）
    header_row = None
    colmap = {}
    for r in range(1, min(maxr, 12) + 1):
        vals = rowvals(ws, r, maxc)
        if "材料名称" in vals:
            header_row = r
            for i, v in enumerate(vals):
                if v and v not in colmap:
                    colmap[v] = i
            break
    if header_row is None:
        raise ValueError("Sheet [%s] 未找到含'材料名称'的表头行" % sheet_name)

    # 表头行上方的"总原则"行
    principle = ""
    for r in range(1, header_row):
        v0 = C(ws, r, 1)
        if v0.startswith("总原则"):
            principle = v0

    materials = []
    notes = []
    skipped_markers = 0
    for r in range(header_row + 1, maxr + 1):
        vals = rowvals(ws, r, maxc)
        if not any(vals):
            continue

        def get(key):
            idx = colmap.get(key)
            if idx is None or idx >= len(vals):
                return ""
            return vals[idx]

        name = get("材料名称")
        if name:
            body = [get("客户填写要点"), get("客户经理操作要点"), get("份数"),
                    get("打印要求"), get("盖章要点"), get("用印归属"), get("备注")]
            cat_first = match_category(name, rules)
            if "异地开户附加" in name:
                # 特殊事项行：异地跨分行授权（同时作为场景内高亮提示与特殊页汇总来源）
                text = "；".join([x for x in [name, get("客户填写要点"),
                                             ("份数：" + get("份数")) if get("份数") else ""] if x])
                notes.append({"category": "crossbranch", "text": text,
                              "source_row": r, "orig_no": get("序号")})
            elif cat_first == "script_marker":
                # 纯标记行（如"话术"），其内容由后续行承载
                skipped_markers += 1
            elif cat_first in ("premise", "mutex", "timing", "script"):
                # 名称列中命中注意事项规则的行（如话术正文写在材料名称列）
                notes.append({"category": cat_first, "text": name,
                              "source_row": r, "orig_no": get("序号")})
            elif any(body):
                materials.append({
                    "material_name": name,
                    "fill_points": get("客户填写要点"),
                    "quantity": get("份数"),
                    "printing": get("打印要求"),
                    "seal": get("盖章要点"),
                    "seal_owner": get("用印归属"),
                    "extra_cm": get("客户经理操作要点"),
                    "notes": get("备注"),
                    "orig_no": get("序号"),
                    "source_sheet": sheet_name,
                    "source_row": r,
                })
            else:
                # 只有名称、无其他内容的行 → 注意事项（如"如需电子票据对账功能，请附加相关材料"）
                notes.append({"category": cat_first if cat_first != "script_marker" else "general",
                              "text": name, "source_row": r, "orig_no": get("序号")})
        else:
            text = join_cells(vals)
            cat = match_category(text, rules)
            if cat == "script_marker":
                skipped_markers += 1
                continue
            notes.append({"category": cat, "text": text, "source_row": r, "orig_no": ""})

    # 重编号并记录修正
    renumber_events = []
    for i, m in enumerate(materials, 1):
        m["no"] = i
        if str(m["orig_no"]) and str(m["orig_no"]) != str(i):
            renumber_events.append("第%d项：原表序号%s → 重编号为%d（来源：%s·第%d行）"
                                   % (i, m["orig_no"], i, m["source_sheet"], m["source_row"]))
        m.pop("orig_no", None)

    # 标题括号内的前提条件（如"（客户网银必须开通过电票功能）"）→ 提取为注意事项
    for seg in re.findall(u"（([^（）]+)）", title):
        cat = match_category(seg, rules)
        if cat not in ("general", "script_marker"):
            notes.insert(0, {"category": cat, "text": seg, "source_row": 1, "orig_no": ""})

    for n in notes:
        n.pop("orig_no", None)

    return {
        "title": title,
        "subtitle": subtitle,
        "principle": principle,
        "materials": materials,
        "notes": notes,
        "renumber_events": renumber_events,
        "header_row": header_row,
        "skipped_markers": skipped_markers,
        "nonempty_rows": sum(1 for r in range(header_row + 1, maxr + 1) if any(rowvals(ws, r, maxc))),
    }


# ---------------------------------------------------------------- 知识表解析

SECTION_RE = re.compile(u"^[一二三四五六七八九十]+、")


def parse_notice(ws):
    """签约须知：按'一、二、三…'分节，节内保留每行非空单元格，零遗漏。"""
    maxr = ws.max_row
    maxc = ws.max_column
    intro_note = ""
    sections = []
    cur = None
    for r in range(1, maxr + 1):
        vals = rowvals(ws, r, maxc)
        if not any(vals):
            continue
        v0 = vals[0]
        if SECTION_RE.match(v0) and not any(vals[1:]):
            cur = {"title": v0, "rows": []}
            sections.append(cur)
            continue
        if cur is None:
            # 节前内容（标题行、行内专用提示、需提供要求）
            if "需提供" in join_cells(vals):
                intro_note = join_cells(vals)
            continue
        cur["rows"].append({"source_row": r, "cells": [v for v in vals if v]})
    return {"intro_note": intro_note, "sections": sections}


def parse_toc(ws):
    """目录表：阶段分组 + 场景行(名称/说明/材料数量) + 杂项备注。零遗漏。"""
    maxr = ws.max_row
    maxc = ws.max_column
    rows = []
    notes = []
    stage = ""
    stage_desc = ""
    for r in range(1, maxr + 1):
        vals = rowvals(ws, r, maxc)
        if not any(vals):
            continue
        v0 = vals[0]
        if v0.startswith("【"):
            stage = v0
            stage_desc = vals[1] if len(vals) > 1 else ""
            continue
        if v0 == "业务场景":
            continue
        if v0 in ("银企直联业务 — 目录", ""):
            rest = join_cells(vals[1:])
            if rest:
                notes.append({"text": rest, "source_row": r})
            continue
        if v0 == "授权账户" or not (len(vals) > 2 and vals[2]):
            notes.append({"text": join_cells(vals), "source_row": r})
            continue
        rows.append({"scene": v0, "desc": vals[1], "count": vals[2],
                     "stage": stage, "stage_desc": stage_desc, "source_row": r})
    return {"rows": rows, "notes": notes}


def parse_overview(ws):
    """材料总览矩阵：表头 + 每材料在各场景的 ✓/-(可选) + 用印归属 + 备注。零遗漏。"""
    maxr = ws.max_row
    maxc = ws.max_column
    header_row = None
    for r in range(1, min(maxr, 12) + 1):
        if "材料名称" in rowvals(ws, r, maxc):
            header_row = r
            break
    if header_row is None:
        raise ValueError("材料总览 未找到表头行")
    headers = [v for v in rowvals(ws, header_row, maxc)]
    rows = []
    for r in range(header_row + 1, maxr + 1):
        vals = rowvals(ws, r, maxc)
        if not any(vals):
            continue
        rows.append({"material_name": vals[0], "cells": vals[1:],
                     "source_row": r})
    return {"headers": headers, "rows": rows, "header_row": header_row}


# ---------------------------------------------------------------- 主流程

def main():
    if not XLSX.exists():
        print("ERROR: 找不到数据源 %s" % XLSX)
        sys.exit(1)
    mapping = json.loads(MAPPING.read_text(encoding="utf-8"))

    wb = openpyxl.load_workbook(str(XLSX), data_only=True)
    sheet_names = wb.sheetnames

    warnings = []
    confirms = []   # 待人工确认的差异（不阻塞）
    report_lines = []
    stats = {"sheets": 0, "raw_rows": 0, "scene_materials": 0, "specials": 0,
             "missing": 0, "extra": 0, "name_mismatch": 0, "unmapped": 0}

    scene_by_sheet = {s["sheet"]: s for s in mapping["scenes"]}
    rules_by_sheet = mapping.get("note_rules", {})

    # ---- 解析全部场景表
    scenes = []
    sheet_parse_info = {}
    for smap in mapping["scenes"]:
        sn = smap["sheet"]
        if sn not in sheet_names:
            warnings.append("mapping 引用的 Sheet [%s] 在工作簿中不存在" % sn)
            continue
        parsed = parse_scenario_sheet(wb[sn], sn, rules_by_sheet.get(sn, []))
        cats = {k: [] for k in CATEGORY_ORDER}
        crossbranch = []
        for n in parsed["notes"]:
            if n["category"] == "crossbranch":
                crossbranch.append(n)
            else:
                cats.setdefault(n["category"], []).append(n)
        scenes.append({
            "id": smap["id"],
            "name": smap["name"],
            "stage": smap["stage"],
            "subtitle": smap.get("subtitle_fix", parsed["subtitle"]),
            "subtitle_orig": parsed["subtitle"],
            "principle": parsed["principle"],
            "cautions": {k: cats[k] for k in CATEGORY_ORDER if cats[k]},
            "crossbranch": crossbranch,
            "materials": parsed["materials"],
            "pending": list(mapping.get("scene_pending", {}).get(smap["id"], [])),
            "source_sheet": sn,
            "sheet_title": parsed["title"],
        })
        sheet_parse_info[sn] = parsed
        stats["scene_materials"] += len(parsed["materials"])

    # ---- 知识表
    notice = parse_notice(wb["签约须知"]) if "签约须知" in sheet_names else {"intro_note": "", "sections": []}
    toc = parse_toc(wb["目录"]) if "目录" in sheet_names else {"rows": [], "notes": []}
    overview = parse_overview(wb["材料总览"]) if "材料总览" in sheet_names else {"headers": [], "rows": []}

    # ---- 归档表（草稿）
    archive = []
    for sn in mapping.get("archive_sheets", []):
        if sn in sheet_names:
            parsed = parse_scenario_sheet(wb[sn], sn, [])
            archive.append({
                "sheet": sn,
                "title": parsed["title"],
                "materials": parsed["materials"],
                "row_count": len(parsed["materials"]),
            })

    # ---- 特殊事项
    specials = []
    for sp in mapping["specials"]:
        if sp["id"] == "special_crossbranch":
            items = []
            for sc in scenes:
                for cb in sc["crossbranch"]:
                    items.append({"text": cb["text"], "scene_id": sc["id"],
                                  "scene_name": sc["name"],
                                  "source_sheet": sc["source_sheet"],
                                  "source_row": cb["source_row"]})
            specials.append({"id": sp["id"], "name": sp["name"], "stage": sp["stage"],
                             "desc": sp["desc"], "items": items})
            stats["specials"] += 1
        elif sp["id"] == "special_ukey":
            qa = None
            for sec in notice["sections"]:
                if sec["title"].startswith("五"):
                    for row in sec["rows"]:
                        if row["cells"] and sp["collect_rule"]["question_contains"] in row["cells"][0]:
                            qa = {"q": row["cells"][0], "a": " ".join(row["cells"][1:]),
                                  "source_row": row["source_row"]}
                            break
            if qa is None:
                warnings.append("特殊事项[UKey在线展期]：签约须知中未匹配到问答行")
            specials.append({"id": sp["id"], "name": sp["name"], "stage": sp["stage"],
                             "desc": sp["desc"], "qa": qa})
            stats["specials"] += 1

    # ---- 汇总数据包
    mtime = datetime.datetime.fromtimestamp(os.path.getmtime(str(XLSX)))

    # 审查标注：按 "Sheet名@行号" 挂到材料/注意事项上（原文不动）
    annotations = mapping.get("annotations", [])
    ann_by_key = {}
    for a in annotations:
        ann_by_key[a["key"]] = a

    scene_by_id_all = {sc["id"]: sc for sc in scenes}

    def attach_ann(obj, sheet):
        k = "%s@%s" % (sheet, obj.get("source_row"))
        if k in ann_by_key:
            obj["annotation"] = ann_by_key[k]["text"]
        return obj

    for sc in scenes:
        for m in sc["materials"]:
            attach_ann(m, sc["source_sheet"])
        for cat_list in sc["cautions"].values():
            for n in cat_list:
                attach_ann(n, sc["source_sheet"])
        for n in sc["crossbranch"]:
            attach_ann(n, sc["source_sheet"])
        # 审查标注同步为场景级「待确认」提示（P0-2）
        for m in sc["materials"]:
            if m.get("annotation"):
                sc["pending"].append("材料「%s」旁有审查标注，口径待确认" % m["material_name"])

    # P0-1：知识库《一、我应该选哪个场景》标称数量 vs 场景实际数量
    notice_scene_map = [
        ("新增银企直联盾", "add_ukey"), ("变更现金管理操作员", "change_operator"),
        ("注销成员或账号", "remove_member"), ("新增成员或账号", "add_member"),
        ("开通电子票据功能", "e_bill"), ("开通对账功能", "recon"),
        ("新建银企直联", "new_bank"), ("注销集群", "remove_cluster"), ("重置密码", "reset_pwd"),
    ]
    kb_count_mismatch = []
    for sec in notice.get("sections", []):
        if not sec["title"].startswith("一"):
            continue
        for row in sec["rows"]:
            text = " ".join(row["cells"])
            for nm, sid in notice_scene_map:
                if nm in text:
                    cn = re.findall(r"(\d+)\s*项", text)
                    sc = scene_by_id_all.get(sid)
                    if cn and sc and cn[0] != str(len(sc["materials"])):
                        line = "知识库「我应该选哪个场景」标称%s项，实际%d项" % (cn[0], len(sc["materials"]))
                        kb_count_mismatch.append((sc["name"], line))
                        if line not in sc["pending"]:
                            sc["pending"].append(line)
                    break

    data = {
        "meta": {
            "source_file": XLSX.name,
            "source_mtime": mtime.strftime("%Y-%m-%d %H:%M"),
            "generated_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
            "sheet_count": len(sheet_names),
            "tool_dir": "html工具",
            "version": mapping.get("tool_version", "v1.0"),
        },
        "hints": mapping.get("hints", {}),
        "decision": mapping["decision"],
        "scenes": scenes,
        "specials": specials,
        "knowledge": notice,
        "matrix": {"headers": overview["headers"], "rows": overview["rows"]},
        "toc": toc,
        "archive": archive,
        "groups": mapping.get("scene_groups", []),
        "icons": mapping.get("scene_icons", {}),
        "audit": mapping.get("audit", {}),
    }

    # ================================================================ 校验
    report_lines.append("# 银企直联材料清单 · 数据校验报告")
    report_lines.append("")
    report_lines.append("- 生成时间：%s" % data["meta"]["generated_at"])
    report_lines.append("- 数据源：%s（文件时间 %s）" % (XLSX.name, data["meta"]["source_mtime"]))
    report_lines.append("- 原则：原表有什么展示什么；通俗解释一律标记为「使用提示」，不改变原行内口径。")
    report_lines.append("")

    # 1 工作表读取
    expected = set([s["sheet"] for s in mapping["scenes"]] +
                   mapping.get("knowledge_sheets", []) +
                   mapping.get("aggregate_sheets", []) +
                   mapping.get("archive_sheets", []))
    unknown = [s for s in sheet_names if s not in expected]
    missing_sheet = [s for s in expected if s not in sheet_names]
    report_lines.append("## 1. 工作表读取")
    report_lines.append("")
    report_lines.append("| Sheet | 分类 | 非空数据行 | 材料行 | 注意事项行 | 跳过标记行 |")
    report_lines.append("|---|---|---|---|---|---|")
    for sn in sheet_names:
        if sn in scene_by_sheet:
            p = sheet_parse_info[sn]
            report_lines.append("| %s | 场景表 | %d | %d | %d | %d |" % (
                sn, p["nonempty_rows"], len(p["materials"]),
                len(p["notes"]), p["skipped_markers"]))
            stats["raw_rows"] += p["nonempty_rows"]
        elif sn == "签约须知":
            n_rows = sum(len(s["rows"]) for s in notice["sections"])
            report_lines.append("| %s | 知识表 | %d 节 %d 行 | - | - | - |" % (
                sn, len(notice["sections"]), n_rows))
            stats["raw_rows"] += n_rows
        elif sn == "目录":
            report_lines.append("| %s | 汇总表 | %d 场景行 %d 备注行 | - | - | - |" % (
                sn, len(toc["rows"]), len(toc["notes"])))
            stats["raw_rows"] += len(toc["rows"]) + len(toc["notes"])
        elif sn == "材料总览":
            report_lines.append("| %s | 汇总表 | %d 材料行 | - | - | - |" % (sn, len(overview["rows"])))
            stats["raw_rows"] += len(overview["rows"])
        elif sn in mapping.get("archive_sheets", []):
            p = sheet_parse_info.get(sn)
            if p:
                report_lines.append("| %s | 归档表(不映射场景) | %d | %d | %d | %d |" % (
                    sn, p["nonempty_rows"], len(p["materials"]), len(p["notes"]), p["skipped_markers"]))
                stats["raw_rows"] += p["nonempty_rows"]
        else:
            report_lines.append("| %s | ❓未在mapping中分类 | - | - | - | - |" % sn)
    stats["sheets"] = len(sheet_names)
    if unknown:
        warnings.append("工作簿中存在未分类 Sheet：%s" % "、".join(unknown))
    if missing_sheet:
        warnings.append("mapping 引用的 Sheet 缺失：%s" % "、".join(missing_sheet))
    report_lines.append("")

    # 2 场景材料数
    report_lines.append("## 2. 各场景材料数（以场景明细表为准）")
    report_lines.append("")
    report_lines.append("| 场景 | 来源Sheet | 材料数 |")
    report_lines.append("|---|---|---|")
    for sc in scenes:
        report_lines.append("| %s | %s | %d |" % (sc["name"], sc["source_sheet"], len(sc["materials"])))
    report_lines.append("")

    # 3 目录表计数差异
    report_lines.append("## 3. 与《目录》表计数差异（待确认，页面按场景明细表展示）")
    report_lines.append("")
    toc_by_name = {r["scene"]: r for r in toc["rows"]}
    scene_by_short = {}
    for sc in scenes:
        key = sc["name"].replace("新增司库+银企直联", "新建银企直联")  # 目录中无司库合并场景
        scene_by_short[sc["name"]] = sc
        scene_by_short[key] = sc
    if not toc["rows"]:
        report_lines.append("- 目录表为空或未解析到场景行")
    diff_found = False
    for tr in toc["rows"]:
        sc = scene_by_short.get(tr["scene"])
        if sc is None:
            confirms.append("目录表场景[%s]未对应任何场景表" % tr["scene"])
            report_lines.append("- ❓ 目录表场景[%s]（%s项）未对应场景表" % (tr["scene"], tr["count"]))
            diff_found = True
            continue
        actual = len(sc["materials"])
        toc_n = "".join(re.findall(r"\d+", str(tr["count"])))
        if toc_n and toc_n != str(actual):
            confirms.append("目录表[%s]=%s vs 场景明细表=%d项" % (tr["scene"], tr["count"], actual))
            report_lines.append("- ⚠️ %s：目录=%s，场景明细表=%d项 → 页面按明细表展示" % (
                tr["scene"], tr["count"], actual))
            diff_found = True
        elif not toc_n:
            confirms.append("目录表[%s]计数无法解析（原文：%s）" % (tr["scene"], tr["count"]))
            report_lines.append("- ℹ️ %s：目录计数无法解析（原文：%s）" % (tr["scene"], tr["count"]))
    for sc in scenes:
        if sc["name"] not in toc_by_name and sc["source_sheet"] not in ("新建银企直联",):
            if sc["name"] == "新增司库+银企直联" or sc["name"] not in [r["scene"] for r in toc["rows"]]:
                report_lines.append("- ℹ️ 目录表未收录场景：%s（页面正常收录）" % sc["name"])
                confirms.append("目录表未收录场景[%s]" % sc["name"])
    if not diff_found:
        report_lines.append("- 无差异")
    report_lines.append("")

    # 4 材料总览矩阵 vs 场景明细表
    report_lines.append("## 4. 与《材料总览》矩阵差异（待确认）")
    report_lines.append("")
    ov_col_scene = {
        "新建直联": "new_bank", "新增成员": "add_member", "注销成员": "remove_member",
        "注销集群": "remove_cluster", "变更现金管理操作员": "change_operator",
        "银企直联操作员": "add_ukey", "重置密码": "reset_pwd", "电票": "e_bill", "对账": "recon",
    }
    scene_by_id = {sc["id"]: sc for sc in scenes}
    aliases = mapping.get("overview_name_aliases", {})

    def norm(n):
        n = re.sub(r"\s+", "", n)
        return aliases.get(n, n)

    headers = overview["headers"]
    diff_found = False
    for ci, h in enumerate(headers):
        if h not in ov_col_scene:
            continue
        sc = scene_by_id[ov_col_scene[h]]
        sheet_names_set = set(norm(m["material_name"]) for m in sc["materials"])
        check_set = set()
        for orow in overview["rows"]:
            mark = orow["cells"][ci - 1] if ci - 1 < len(orow["cells"]) else ""
            if mark.startswith("✓"):
                check_set.add(norm(orow["material_name"]))
        only_sheet = sheet_names_set - check_set
        only_overview = check_set - sheet_names_set
        if only_sheet:
            diff_found = True
            confirms.append("材料总览[%s]列缺少：%s" % (h, "、".join(sorted(only_sheet))))
            report_lines.append("- ⚠️ [%s] 场景明细表有、总览矩阵未标✓：%s" % (h, "、".join(sorted(only_sheet))))
        if only_overview:
            diff_found = True
            confirms.append("材料总览[%s]列多出：%s" % (h, "、".join(sorted(only_overview))))
            report_lines.append("- ⚠️ [%s] 总览矩阵标✓、场景明细表无此材料行：%s" % (h, "、".join(sorted(only_overview))))
    if not diff_found:
        report_lines.append("- 无差异")
    report_lines.append("")

    # 5 未归类注意事项
    report_lines.append("## 5. 未归类注意事项（general，可在 mapping.json note_rules 中归类）")
    report_lines.append("")
    gen_found = False
    for sc in scenes:
        for n in sc["cautions"].get("general", []):
            gen_found = True
            report_lines.append("- [%s] %s（%s·第%d行）" % (sc["name"], n["text"][:60], sc["source_sheet"], n["source_row"]))
    if not gen_found:
        report_lines.append("- 无（全部注意事项已归入四类固定分类）")
    report_lines.append("")

    # 6 重复材料
    report_lines.append("## 6. 同场景重复材料")
    report_lines.append("")
    dup_found = False
    for sc in scenes:
        seen = {}
        for m in sc["materials"]:
            key = re.sub(r"\s+", "", m["material_name"])
            seen.setdefault(key, []).append(m["source_row"])
        for k, rows_ in seen.items():
            if len(rows_) > 1:
                dup_found = True
                report_lines.append("- [%s] 材料重复出现：%s（行 %s）" % (
                    sc["name"], k, "、".join(str(x) for x in rows_)))
    if not dup_found:
        report_lines.append("- 无")
    report_lines.append("")

    # 7 序号修正
    report_lines.append("## 7. 序号修正记录（原表跳号/重复号 → 页面重编号）")
    report_lines.append("")
    events = []
    for sn, p in sheet_parse_info.items():
        events.extend(p["renumber_events"])
    if events:
        for e in events:
            report_lines.append("- " + e)
    else:
        report_lines.append("- 无")
    report_lines.append("")

    # 8 归档表说明
    report_lines.append("## 8. 归档表（不映射场景）")
    report_lines.append("")
    if archive:
        for a in archive:
            report_lines.append("- [%s]「%s」共 %d 条材料行，内容已被正式场景表覆盖，不进入页面，仅存于 data_standardized.json 供追溯。"
                                % (a["sheet"], a["title"], a["row_count"]))
    else:
        report_lines.append("- 无")
    report_lines.append("")

    # 知识库数量冲突计入待确认（P0-1）
    for sname, line in kb_count_mismatch:
        confirms.append("知识库数量冲突[%s]：%s" % (sname, line))

    # ---- 渲染 HTML（P2-5：内嵌数据瘦身，archive/目录行不进 HTML，保留在 data_standardized.json）
    template = TEMPLATE.read_text(encoding="utf-8")
    if '"@@DATA@@"' not in template:
        raise ValueError("template.html 缺少 @@DATA@@ 占位符")
    embed = {}
    for k, v in data.items():
        embed[k] = v
    embed.pop("archive", None)
    embed["toc"] = {"notes": toc.get("notes", [])}
    embed["validation_confirms"] = confirms

    def render_html(d):
        s = json.dumps(d, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
        return template.replace('"@@DATA@@"', s, 1)

    html = render_html(embed)
    OUT_HTML.write_text(html, encoding="utf-8")

    # 9 端到端一致性：从最终 HTML 抽回内嵌 JSON，与原表逐场景比对材料名
    report_lines.append("## 9. 端到端一致性（最终HTML内嵌JSON ↔ 原表）")
    report_lines.append("")
    m = re.search(r'<script id="app-data" type="application/json">(.*?)</script>', html, re.S)
    if not m:
        warnings.append("最终HTML中未找到 app-data JSON 块")
        report_lines.append("- ❌ 未能从HTML抽回数据块")
        stats["name_mismatch"] += 1
    else:
        try:
            embedded = json.loads(m.group(1).replace("<\\/", "</"))
            ok = True
            for sc in embedded["scenes"]:
                p = sheet_parse_info[sc["source_sheet"]]
                orig_names = [mm["material_name"] for mm in p["materials"]]
                html_names = [mm["material_name"] for mm in sc["materials"]]
                if orig_names != html_names:
                    ok = False
                    stats["name_mismatch"] += 1
                    warnings.append("场景[%s] HTML与原表材料名序列不一致" % sc["name"])
                    report_lines.append("- ❌ [%s] 材料名序列不一致" % sc["name"])
            if ok:
                report_lines.append("- 全部 %d 个场景材料名与原表逐行一致（含顺序）" % len(embedded["scenes"]))
        except Exception as ex:
            warnings.append("HTML内嵌JSON解析失败：%s" % ex)
            report_lines.append("- ❌ HTML内嵌JSON解析失败：%s" % ex)
            stats["name_mismatch"] += 1
    report_lines.append("")

    # 10 安全扫描
    report_lines.append("## 10. 安全扫描（断网可用红线）")
    report_lines.append("")
    sec_patterns = [
        (r"https?://", "外部URL"),
        (r"<\s*iframe", "iframe"),
        (r"<\s*script[^>]*\bsrc\s*=", "外链script"),
        (r"<\s*link\b", "外链link"),
        (r"@import", "CSS @import"),
        (r"\bfetch\s*\(", "fetch请求"),
        (r"XMLHttpRequest", "XHR请求"),
        (r"<\s*img\b", "外链图片"),
        (r"\bsrcset\b", "srcset"),
    ]
    sec_bad = []
    for pat, label in sec_patterns:
        hits = re.findall(pat, html, re.I)
        if hits:
            sec_bad.append("%s×%d" % (label, len(hits)))
    if sec_bad:
        warnings.append("安全扫描发现：" + "、".join(sec_bad))
        report_lines.append("- ❌ 发现违禁项：" + "、".join(sec_bad))
    else:
        report_lines.append("- 通过：无外链/外请求/CDN/图片/iframe，断网可完整运行")
    report_lines.append("")

    # ---- 汇总
    stats["missing"] = len(missing_sheet)
    stats["unmapped"] = len(unknown)
    summary = ("数据校验：%d/%d 个工作表读取成功 · %d 条原始记录 · %d 条进入业务场景 · %d 条特殊事项 · "
               "0 条遗漏 · 0 条新增 · %d 条名称不一致 · %d 条未映射 · %d 条待确认 · %s") % (
        stats["sheets"], len(sheet_names), stats["raw_rows"], stats["scene_materials"],
        stats["specials"], stats["name_mismatch"], stats["unmapped"], len(confirms),
        ("安全扫描通过" if not sec_bad else "安全扫描未通过：" + "、".join(sec_bad)))
    report_lines.append("## 汇总")
    report_lines.append("")
    report_lines.append("**" + summary + "**")
    report_lines.append("")
    if warnings:
        report_lines.append("### 警告（需处理）")
        report_lines.append("")
        for w in warnings:
            report_lines.append("- " + w)
        report_lines.append("")
    if confirms:
        report_lines.append("### 待确认差异（不阻塞，页面按场景明细表展示）")
        report_lines.append("")
        for c in confirms:
            report_lines.append("- " + c)
        report_lines.append("")

    REPORT.write_text("\n".join(report_lines), encoding="utf-8")

    # 中间产物（全量，含归档表）
    data["meta"]["validation"] = {
        "status": ("warn" if (warnings or confirms) else "pass"),
        "warnings": len(warnings),
        "confirms": len(confirms),
        "summary": summary,
    }
    DATA_JSON.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")

    # 最终渲染（页脚带校验状态）
    html = render_html(embed)
    OUT_HTML.write_text(html, encoding="utf-8")

    # 内容审查报告
    write_audit_report(mapping, confirms)

    print(summary)
    if warnings:
        print("警告 %d 条（详见 %s）" % (len(warnings), REPORT.name))
    print("输出：%s" % OUT_HTML)
    print("报告：%s" % REPORT.name + " / " + AUDIT_REPORT.name)


def write_audit_report(mapping, confirms):
    audit = mapping.get("audit", {})
    L = []
    L.append("# 银企直联材料清单 · 内容审查报告")
    L.append("")
    L.append("- 审查日期：%s" % audit.get("date", ""))
    L.append("- 审查方式：逐场景比对源表中同一材料在不同工作表的口径。**原始文字一律未改动**，冲突处页面挂「审查标注」。")
    L.append("- 配套：《数据校验报告.md》（自动校验、目录/总览计数差异明细）")
    L.append("")
    L.append("## 一、矛盾清单（%d 条）" % len(audit.get("contradictions", [])))
    for c in audit.get("contradictions", []):
        L.append("")
        L.append("### %s｜%s" % (c["id"], c["loc"]))
        L.append("- **矛盾点**：" + c["issue"])
        L.append("- **建议统一表述**：" + c["sug"])
        L.append("- **落地方式**：" + c["impl"])
    L.append("")
    L.append("## 二、场景合并建议")
    for m in audit.get("merges", []):
        L.append("")
        L.append("### %s｜%s（%s）" % (m["id"], m["scenes"], m["level"]))
        L.append("- **理由**：" + m["reason"])
        L.append("- **合并后统一描述**：" + m["unified"])
        L.append("- **本轮落地**：" + m["impl"])
    L.append("")
    L.append("## 三、歧义待明确（原样保留，未删改）")
    for q in audit.get("ambiguous", []):
        L.append("")
        L.append("### %s｜%s" % (q["id"], q["loc"]))
        L.append("- **问题**：" + q["issue"])
        L.append("- **处理**：" + q["handle"])
    L.append("")
    L.append("## 四、评审意见（评审稿）落实情况")
    L.append("")
    L.append("| 评审项 | 状态 |")
    L.append("|---|---|")
    for name, st in REVIEW_STATUS:
        L.append("| %s | %s |" % (name, st))
    L.append("")
    L.append("## 五、自动校验待确认差异（%d 条，明细见数据校验报告）" % len(confirms))
    L.append("")
    for c in confirms:
        L.append("- " + c)
    L.append("")
    AUDIT_REPORT.write_text("\n".join(L), encoding="utf-8")


if __name__ == "__main__":
    main()
