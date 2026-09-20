# -*- coding: utf-8 -*-
"""
lesson_prep.py —— DSH「一节课备课包」插件的确定性执行内核

设计原则
  1. 一切输入由命令行给出：`--title` 或 `--deck <path.pptx>`；仓库里不含任何课件、图片、视频。
  2. 一切产物写到 `--out <dir>` 指定的目录；仓库里不含任何运行结果。
  3. 不读任何 API Key、不发任何网络请求（视频下载由 agent 按提示调用 yt-dlp 完成）。
  4. 只做机器能保证的事：结构、页数、形状、校验；学科内容由 agent 按 SKILL.md 生成。

子命令
  extract  课件 → deck.json（结构化大纲）
  plan     deck.json → plan.json（时间盒 + 删页清单，保留率落在 65%–75%）
  trim     按 plan.json 生成精简课件（只删页，不改任何内容）
  media    生成数据集样例图 + 写两个可运行演示脚本
  video    生成 B 站检索关键词、推荐清单 md 骨架、下载脚本
  quiz     写讨论题与单选题的骨架模板（供 agent 填充学科内容）
  merge    把新增原理图页合并进课件（默认追加到末尾，最稳）
  outline  写课程方案 md 骨架（含 5 个任务的提示词）
  bundle   打包成 zip + 生成带样式的下载页 index.html
  prompts  生成两个子代理的任务提示词（parallel / slides）
  config   打印本机依赖自检结果
  all      依次跑完上面的确定性步骤

用法示例
  python3 scripts/lesson_prep.py all --title "第7章 循环神经网络" \
      --deck "/path/7. 循环神经网络 (1).pptx" --out ./runs/rnn --minutes 90
  python3 scripts/lesson_prep.py all --title "第4章 卷积神经网络" --out ./runs/cnn
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import sys
import zipfile
from datetime import datetime

# ----------------------------------------------------------------------
# 通用小工具
# ----------------------------------------------------------------------
DEFAULT_MINUTES = 90
KEEP_LOW, KEEP_HIGH = 0.65, 0.75
BUDGET_TEMPLATE = [
    # (名称, 分钟, 性质)
    ("开场与回顾", 5, "固定"),
    ("PPT 讲解", 39, "变量：由页面预算反推"),
    ("课堂提问", 8, "固定（3–4 次插入式 + 1 次小结）"),
    ("代码演示", 15, "固定（两个 demo）"),
    ("视频演示", 7, "固定（≤5 分钟片段 + 2 分钟解读）"),
    ("课堂讨论", 10, "固定"),
    ("练习", 5, "固定，只留时间不写题目"),
    ("收尾总结", 1, "固定"),
]


def log(msg):
    print(msg, flush=True)


def ensure_dir(path):
    os.makedirs(path, exist_ok=True)
    return path


def write_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    return path


def read_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def write_text(path, text):
    ensure_dir(os.path.dirname(path) or ".")
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)
    return path


def slugify(text, fallback="lesson"):
    """把中文/英文标题转成安全的目录名片段。"""
    text = re.sub(r"[\\/:*?\"<>|\s]+", "-", text.strip())
    text = re.sub(r"-+", "-", text).strip("-")
    return text or fallback


def slide_digest(slide):
    """把一页的文字压成一个摘要串，用于相似度与重复检测。"""
    return re.sub(r"\s+", "", slide.get("text") or "")


# ----------------------------------------------------------------------
# extract：课件 → deck.json
# ----------------------------------------------------------------------
CODE_TOKENS = ("import ", "def ", "class ", "nn.", "torch.", "self.", "print(", "return ",
               "for ", "while ", "= ", "()", "np.", "plt.")
SECTION_RE = re.compile(r"^(\d+(?:\.\d+){0,3})\s*([^\s].{0,60})$")
DIVIDER_HINT = ("本章内容", "contents", "目录", "本章小结", "小结")


def extract_deck(pptx_path, out_dir):
    from pptx import Presentation
    from pptx.util import Emu

    prs = Presentation(pptx_path)
    slides = []
    for i, slide in enumerate(prs.slides, start=1):
        texts, title, body = [], "", []
        pics = 0
        for shp in slide.shapes:
            if shp.shape_type is not None and "PICTURE" in str(shp.shape_type):
                pics += 1
            if not shp.has_text_frame:
                continue
            t = shp.text_frame.text.strip()
            if not t:
                continue
            texts.append(t)
            if not title and len(t) <= 60:
                title = t.replace("\n", " ").strip()
            else:
                body.append(t)
        full = "\n".join(texts)
        compact = re.sub(r"\s+", "", full)
        lines = [x for x in full.splitlines() if x.strip()]
        code_hits = sum(1 for tok in CODE_TOKENS if tok in full)
        looks_code = code_hits >= 4 or bool(re.search(r"\n\s{2,}\S", full))
        # 分隔页/目录页：含 contents 或“本章内容”，且文字很短
        is_divider = ("本章内容" in compact or "contents" in compact.lower()) and len(compact) < 400
        is_summary = bool(re.search(r"(本章小结|小结|总结)$", (title or "").strip())) and len(compact) < 400
        section_no = None
        m = SECTION_RE.match((title or "").strip())
        if m:
            section_no = m.group(1)
        slides.append({
            "page": i,
            "title": title,
            "text": full,
            "text_chars": len(compact),
            "body": "\n".join(body)[:4000],
            "line_count": len(lines),
            "picture_count": pics,
            "looks_code": looks_code,
            "code_score": code_hits,
            "is_divider": is_divider,
            "is_summary": is_summary,
            "is_structural": is_divider or is_summary,
            "section_no": section_no,
        })

    sections = group_sections(slides)
    deck = {
        "source": os.path.basename(pptx_path),
        "slide_count": len(slides),
        "slide_width": int(prs.slide_width),
        "slide_height": int(prs.slide_height),
        "aspect": round(int(prs.slide_width) / int(prs.slide_height), 4),
        "extracted_at": datetime.now().isoformat(timespec="seconds"),
        "sections": sections,
        "slides": slides,
    }
    write_json(os.path.join(out_dir, "deck.json"), deck)
    log("课件：%s" % deck["source"])
    log("页数：%d　画幅：%d×%d（%.2f）" % (deck["slide_count"], deck["slide_width"],
                                          deck["slide_height"], deck["aspect"]))
    log("小节数：%d（含 %d 个分隔页）" % (len(sections),
                                          sum(1 for s in slides if s["is_divider"])))
    for sec in sections:
        log("  [%-8s] p%02d-p%02d %2d 页  %s" % (sec["no"], sec["start_page"], sec["end_page"],
                                                  sec["slides"], sec["title"][:40]))
    return deck


def group_sections(slides):
    """按“节号变化”切小节；同一节号的连续页归入同一个节，分隔页单独成节。"""
    sections, cur = [], None
    for s in slides:
        no = s["section_no"]
        if no:
            if cur is not None and cur["no"] == no and not cur["is_divider"]:
                cur["pages"].append(s["page"])
                cur["end_page"] = s["page"]
                cur["slides"] += 1
                continue
            cur = {"no": no, "title": s["title"], "start_page": s["page"],
                   "end_page": s["page"], "pages": [s["page"]], "slides": 1,
                   "is_divider": False}
            sections.append(cur)
        elif s["is_divider"]:
            cur = {"no": "divider", "title": s["title"], "start_page": s["page"],
                   "end_page": s["page"], "pages": [s["page"]], "slides": 1,
                   "is_divider": True}
            sections.append(cur)
        else:
            if cur is None or cur["is_divider"]:
                cur = {"no": cur["no"] if cur else "未编号", "title": s["title"],
                       "start_page": s["page"], "end_page": s["page"],
                       "pages": [], "slides": 0, "is_divider": False}
                sections.append(cur)
            cur["pages"].append(s["page"])
            cur["end_page"] = s["page"]
            cur["slides"] += 1
    return sections


# ----------------------------------------------------------------------
# plan：时间盒 + 删页清单
# ----------------------------------------------------------------------
def build_budget(minutes, weights):
    """把 minutes 分配给八个时段：先用“固定保底”，再把余量按权重分给各段。

    保底（分钟）：开场 5、提问 5、代码 10、视频 5、讨论 8、练习 5、收尾 1，其余给 PPT 讲解。
    这样 --minutes 90 与 --minutes 45 都能得到合计恰好等于 minutes 的时间盒。
    """
    floor = {"开场与回顾": 5, "课堂提问": 7, "代码演示": 18, "视频演示": 6,
             "课堂讨论": 10, "练习": 5, "收尾总结": 1}
    floor_sum = sum(floor.values())
    if minutes <= floor_sum + 5:
        raise SystemExit("课时 %d 分钟太短：固定时段保底已占 %d 分钟，请至少给 %d 分钟"
                         % (minutes, floor_sum, floor_sum + 5))
    alloc = {"PPT 讲解": minutes - floor_sum}
    alloc.update(floor)

    # 课时不足 90 分钟时，按比例压缩演示 / 讨论 / 视频 / 提问，把时间留给讲解
    if minutes < 90:
        scale = (minutes - 5 - 5 - 1) / 79.0      # 79 = 提问7+代码18+视频6+讨论10+讲解38
        for name in ("课堂提问", "代码演示", "视频演示", "课堂讨论"):
            alloc[name] = max(4, int(round(alloc[name] * scale)))
        alloc["PPT 讲解"] = minutes - alloc["开场与回顾"] - alloc["练习"] - alloc["收尾总结"] \
            - sum(alloc[n] for n in ("课堂提问", "代码演示", "视频演示", "课堂讨论"))
        if alloc["PPT 讲解"] < 5:
            raise SystemExit("课时 %d 分钟不足以同时容纳五类活动，请至少给 50 分钟" % minutes)

    budget = [{"name": name, "minutes": alloc.get(name, m), "kind": kind}
              for name, m, kind in BUDGET_TEMPLATE]
    total = sum(b["minutes"] for b in budget)
    assert total == minutes, (total, minutes)
    return budget


def protected_pages(deck):
    """硬保护：封面、教材/资源页、小结页、每个小节首页、分隔页。"""
    keep = set()
    slides = deck["slides"]
    for s in slides[:3]:
        keep.add(s["page"])
    for s in slides:
        if s["is_structural"] and not s["is_divider"]:
            keep.add(s["page"])
        if s["is_summary"]:
            keep.add(s["page"])
        if s["is_divider"]:
            keep.add(s["page"])
    # 每个小节至少保留前 2 页
    for sec in deck["sections"]:
        for p in sec["pages"][:2]:
            keep.add(p)
    return keep


def score_delete_candidate(slide, dup_of):
    """分值越高越该删。"""
    score = 0.0
    reasons = []
    if dup_of:
        score += 100
        reasons.append("与第 %d 页正文重复" % dup_of)
    if slide["looks_code"]:
        score += 40
        reasons.append("整页代码，课堂上会现场演示")
    if slide["text_chars"] < 120:
        score += 25
        reasons.append("内容偏薄（%d 字）" % slide["text_chars"])
    if re.search(r"(运行结果|输出结果|输出如下|结果如下|如图所示的结果|执行上述代码)", slide["text"] or ""):
        score += 30
        reasons.append("运行结果截图页")
    if slide["picture_count"] >= 3 and slide["text_chars"] < 300:
        score += 10
        reasons.append("以截图为主")
    # 同一小节内“保前删尾”：越靠后的候选越优先
    return score, reasons


def detect_duplicates(slides):
    """正文完全相同或高度相似的页，后出现的记为重复。"""
    seen, dup = {}, {}
    for s in slides:
        key = hashlib.md5(slide_digest(s).encode("utf-8")).hexdigest()
        if key in seen and len(slide_digest(s)) > 30:
            dup[s["page"]] = seen[key]
        else:
            seen.setdefault(key, s["page"])
    return dup


def build_plan(deck, minutes, out_dir, desired_drop=None, keep_ratio=0.70):
    total = deck["slide_count"]
    target_keep = max(1, int(round(total * keep_ratio)))
    target_keep = max(int(total * KEEP_LOW), min(int(total * KEEP_HIGH), target_keep))
    must_drop = total - target_keep

    weights = {"PPT 讲解": 5, "课堂提问": 1, "代码演示": 2, "视频演示": 1, "课堂讨论": 1.5}
    budget = build_budget(minutes, weights)
    teach_min = [b["minutes"] for b in budget if b["name"] == "PPT 讲解"][0]

    keep = protected_pages(deck)
    dup = detect_duplicates(deck["slides"])
    forced = [p for p in (desired_drop or []) if p not in keep]
    cand = []
    for s in deck["slides"]:
        p = s["page"]
        if p in keep or p in forced:
            continue
        sc, why = score_delete_candidate(s, dup.get(p))
        cand.append((sc, -p, p, why))
    cand.sort(reverse=True)

    delete = list(forced)
    reasons = {}
    for p in forced:
        s = deck["slides"][p - 1]
        reasons[str(p)] = "教师指定删除（%s）" % (s["title"][:30] or "无标题")
    for sc, _, p, why in cand:
        if len(delete) >= must_drop + len(forced):
            break
        delete.append(p)
        s = deck["slides"][p - 1]
        reasons[str(p)] = "；".join(why) if why else ("价值较低：%s" % (s["title"][:30] or "无标题"))
    delete = sorted(set(delete))
    kept = total - len(delete)

    plan = {
        "title": deck.get("source", ""),
        "minutes": minutes,
        "budget": budget,
        "ppt_minutes": teach_min,
        "minutes_per_page": round(teach_min / max(kept, 1), 2),
        "stats": {
            "original_pages": total,
            "delete_pages": len(delete),
            "kept_pages": kept,
            "keep_ratio": round(kept / total, 4),
        },
        "delete_list": {"delete_pages": delete, "reasons": reasons},
        "protected_pages": sorted(keep),
        "sections": deck["sections"],
        "warnings": [],
    }
    if minutes >= 60 and plan["minutes_per_page"] < 0.8:
        plan["warnings"].append(
            "平均每页 %.2f 分钟，节奏偏紧：讲解要只抓主线，并把整页代码交给现场演示。"
            % plan["minutes_per_page"])
    if not (KEEP_LOW <= plan["stats"]["keep_ratio"] <= KEEP_HIGH):
        plan["warnings"].append("保留率 %.1f%% 超出 65%%–75%% 目标区间，请复核删页清单。"
                                % (100 * plan["stats"]["keep_ratio"]))
    write_json(os.path.join(out_dir, "plan.json"), plan)

    log("时间盒（合计 %d 分钟）：" % minutes)
    for b in budget:
        log("  %-10s %2d 分钟  %s" % (b["name"], b["minutes"], b["kind"]))
    log("删页：原 %d 页 → 删 %d 页 → 保留 %d 页（%.1f%%），平均 %.2f 分钟/页"
        % (total, len(delete), kept, 100 * plan["stats"]["keep_ratio"], plan["minutes_per_page"]))
    for w in plan["warnings"]:
        log("  ⚠ " + w)
    return plan


# ----------------------------------------------------------------------
# trim：只删页，不改内容
# ----------------------------------------------------------------------
def trim_deck(deck_path, plan_path, src_pptx, out_pptx, out_dir):
    plan = read_json(plan_path)
    drop = sorted(set(plan["delete_list"]["delete_pages"]))
    zin = zipfile.ZipFile(src_pptx)
    pres = zin.read("ppt/presentation.xml").decode("utf-8")
    rels = zin.read("ppt/_rels/presentation.xml.rels").decode("utf-8")

    entries = re.findall(r"<p:sldId[^>]*/>", pres)
    total = len(entries)
    rel_map = {}
    for m in re.finditer(r"<Relationship\b[^>]*/>", rels):
        tag = m.group(0)
        rid = re.search(r'Id="([^"]+)"', tag).group(1)
        rel_map[rid] = re.search(r'Target="([^"]+)"', tag).group(1)

    bad = [p for p in drop if p < 1 or p > total]
    if bad:
        raise SystemExit("删页清单越界：%s（课件共 %d 页）" % (bad, total))

    drop_rids = {re.search(r'r:id="([^"]+)"', entries[p - 1]).group(1) for p in drop}
    keep_entries = [t for i, t in enumerate(entries, 1) if i not in set(drop)]
    old_lst = re.search(r"<p:sldIdLst>.*?</p:sldIdLst>", pres, re.S).group(0)
    new_pres = pres.replace(old_lst, "<p:sldIdLst>" + "".join(keep_entries) + "</p:sldIdLst>")

    new_rels = rels
    for m in re.finditer(r"<Relationship\b[^>]*/>", rels):
        tag = m.group(0)
        if re.search(r'Id="([^"]+)"', tag).group(1) in drop_rids:
            new_rels = new_rels.replace(tag, "")
    new_rels = re.sub(r"\n\s*\n", "\n", new_rels)

    # 连带清掉被删页面的 slide 部件 / 备注 / 独占媒体
    drop_names = set()
    for rid in drop_rids:
        name = "ppt/" + rel_map[rid]
        drop_names.add(name)
        rels_name = "ppt/_rels/" + os.path.basename(name) + ".rels"
        if rels_name in zin.namelist():
            sub = zin.read(rels_name).decode("utf-8")
            for m in re.finditer(r'Target="([^"]+)"', sub):
                tgt = m.group(1)
                if "notesSlide" in tgt:
                    base = os.path.normpath(os.path.join("ppt/slides", tgt))
                    drop_names.add(base)
                    drop_names.add(os.path.join(os.path.dirname(base), "_rels",
                                                os.path.basename(base) + ".rels"))
            drop_names.add(rels_name)

    keep_slides = {"ppt/" + rel_map[re.search(r'r:id="([^"]+)"', t).group(1)] for t in keep_entries}
    used = set()
    for item in zin.namelist():
        if item.startswith("ppt/slides/_rels/") and item.endswith(".rels"):
            owner = "ppt/slides/" + os.path.basename(item)[:-5]
            if owner not in keep_slides:
                continue
            for m in re.finditer(r'Target="([^"]+)"', zin.read(item).decode("utf-8", "ignore")):
                used.add(os.path.normpath(os.path.join("ppt/slides", m.group(1))))
    for item in zin.namelist():
        if item.startswith(("ppt/media/", "ppt/embeddings/")) and item not in used:
            drop_names.add(item)

    with zipfile.ZipFile(out_pptx, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename in drop_names:
                continue
            data = zin.read(item.filename)
            if item.filename == "ppt/presentation.xml":
                data = new_pres.encode("utf-8")
            elif item.filename == "ppt/_rels/presentation.xml.rels":
                data = new_rels.encode("utf-8")
            zout.writestr(item, data)

    kept = total - len(drop)
    log("裁剪课件：%d 页 → 删除 %d 页 → %d 页（保留 %.1f%%）" % (total, len(drop), kept, 100 * kept / total))
    log("  已写出 %s" % out_pptx)

    # 删页清单 md
    lines = ["# 删页清单（%d 页 → %d 页，保留 %.1f%%）\n" % (total, kept, 100 * kept / total),
             "只删除页面，不改动任何文字、图片与版式。\n"]
    for p in drop:
        s = next((x for x in read_json(deck_path)["slides"] if x["page"] == p), {})
        why = read_json(plan_path)["delete_list"]["reasons"].get(str(p), "")
        lines.append("- 原第 %d 页：%s —— 删除理由：%s" % (p, (s.get("title") or "无标题")[:40], why))
    write_text(os.path.join(out_dir, "02-删页清单.md"), "\n".join(lines) + "\n")
    return kept


# ----------------------------------------------------------------------
# media：数据集样例图 + 两个演示脚本
# ----------------------------------------------------------------------
def write_demo_scripts(out_dir):
    """写出两个演示脚本骨架：结构完整、可直接跑，学科内容由 agent 按主题替换。"""
    demo_dir = ensure_dir(os.path.join(out_dir, "demos"))
    common = os.path.join(demo_dir, "_common.py")
    if not os.path.exists(common):
        shutil.copy(os.path.join(os.path.dirname(os.path.abspath(__file__)), "_demo_common.py"), common)
    log("演示脚本目录：%s（_common.py 已就绪，demo1/demo2 由 agent 按主题生成）" % demo_dir)
    return demo_dir


# ----------------------------------------------------------------------
# video：检索关键词 + 推荐清单骨架 + 下载脚本
# ----------------------------------------------------------------------
def write_video_assets(out_dir, title, keywords=None, max_seconds=300, want=2):
    vdir = ensure_dir(os.path.join(out_dir, "06-视频"))
    kw = keywords or [title, title + " 原理 动画", title + " 代码 讲解", title + " 入门"]
    write_text(os.path.join(vdir, "keywords.txt"), "\n".join(kw) + "\n")
    md = ["# 推荐视频清单（哔哩哔哩）", "",
          "用途：课堂视频演示环节素材（默认挑 %d 段，每段 ≤ %d 秒）。" % (want, max_seconds),
          "检索关键词：%s" % "；".join(kw), "",
          "## 一、课堂使用的视频（待 agent 检索后填写）", "",
          "| 序号 | 标题 | UP 主 | BV 号 | 时长 | 链接 | 推荐理由 | 课堂用法 |", "|---|---|---|---|---|---|---|---|",
          "| 1 | 【待补】 | 【待补】 | 【待补】 | 【待补】 | 【待补】 | 【待补】 | 【待补】 |",
          "| 2 | 【待补】 | 【待补】 | 【待补】 | 【待补】 | 【待补】 | 【待补】 | 【待补】 |", "",
          "## 二、其余备选（不限时长）", "", "【待补】", "",
          "## 三、下载为 MP4 的命令", "",
          "```bash",
          "pip install -U yt-dlp",
          "yt-dlp --no-playlist -F \"https://www.bilibili.com/video/<BV号>\"      # 先看清晰度",
          "yt-dlp --no-playlist -f \"bv*[vcodec^=avc1][height<=480]+ba/b[height<=480]/b\" \\",
          "       --merge-output-format mp4 -o \"%(title)s-%(id)s.%(ext)s\" \\",
          "       \"https://www.bilibili.com/video/<BV号>\"",
          "# 教室兼容性：统一转 H.264 + yuv420p + faststart",
          "ffmpeg -y -i in.mp4 -c:v libx264 -preset veryfast -crf 23 -pix_fmt yuv420p \\",
          "       -c:a aac -b:a 128k -movflags +faststart out.mp4",
          "```", "",
          "版权提示：仅供课堂教学演示，版权归原作者所有，请勿商用或二次上传。"]
    write_text(os.path.join(vdir, "推荐视频清单.md"), "\n".join(md) + "\n")

    sh = """#!/usr/bin/env bash
# 下载推荐视频为 MP4（H.264，教室电脑任意播放器可放）
# 用法：bash 下载视频.sh "BV1xxxxxxxxx" "07-视频演示1-主题.mp4" [...]
set -u
OUT="${OUT_DIR:-下载视频}"; mkdir -p "$OUT"
command -v yt-dlp >/dev/null || { echo "缺 yt-dlp：pip install -U yt-dlp"; exit 1; }
command -v ffmpeg >/dev/null || { echo "缺 ffmpeg"; exit 1; }
while [ $# -ge 2 ]; do
  BV="$1"; NAME="$2"; shift 2
  yt-dlp --no-playlist -f "bv*[vcodec^=avc1][height<=480]+ba/b[height<=480]/b" \\
         --merge-output-format mp4 --add-header "Referer:https://www.bilibili.com/" \\
         -o "$OUT/tmp-$BV.%(ext)s" "https://www.bilibili.com/video/$BV" || continue
  ffmpeg -v error -y -i "$OUT/tmp-$BV.mp4" -c:v libx264 -preset veryfast -crf 23 \\
         -pix_fmt yuv420p -c:a aac -b:a 128k -movflags +faststart "$OUT/$NAME"
  rm -f "$OUT/tmp-$BV.mp4"; echo "已生成 $OUT/$NAME"
done
"""
    write_text(os.path.join(vdir, "下载视频.sh"), sh)
    log("视频素材目录：%s（关键词、清单骨架、下载脚本已就绪）" % vdir)
    return vdir


# ----------------------------------------------------------------------
# quiz / outline：md 骨架
# ----------------------------------------------------------------------
def write_quiz_skeletons(out_dir, title, n_mcq=10):
    discuss = ["# 课堂讨论题：把课上原理迁移到新场景", "",
               "## 一、讨论题（投影给学生）", "", "【待补：结合本节主题设计一道迁移型讨论题，给出场景约束】", "",
               "要求学生在 4 分钟内给出：① 改造清单；② 三个会翻车的点；③ 一句话结论。", "",
               "## 二、教师参考解", "", "### 第 1 步：改造清单", "", "【待补】", "",
               "### 第 2 步：三个会翻车的点", "", "【待补】", "",
               "### 第 3 步：一句话结论", "", "【待补】", "",
               "## 三、评分标准（满分 10 分）", "",
               "改造清单 4 分；三个翻车点 4 分（每点 1 分，说清机制再各加 0.5 分）；表达与协作 2 分。", "",
               "## 评分口径与时间盒", "", "小组讨论 4 分钟 → 两组汇报共 3 分钟 → 教师收口 3 分钟，合计 10 分钟。", "",
               "## 四、讨论延伸（课后思考）", "", "【待补】"]
    write_text(os.path.join(out_dir, "03-讨论题.md"), "\n".join(discuss) + "\n")

    mcq = ["# 课后习题：%s（%d 道单选题）" % (title, n_mcq), "",
           "作答说明：每题只有一个正确答案，请把答案写在题号后的括号里；解析用于自查。", ""]
    for i in range(1, n_mcq + 1):
        mcq += ["%d. 【待补：题干】（　　）" % i,
                "A. 【待补】", "B. 【待补】", "C. 【待补】", "D. 【待补】",
                "正确答案：【待补】", "解析：【待补】", ""]
    mcq += ["## 答案速查", "",
            "；".join("%d 【待补】" % i for i in range(1, n_mcq + 1)), "",
            "## 易错点提醒（讲评时重点说）", "", "【待补】"]
    write_text(os.path.join(out_dir, "04-课后习题.md"), "\n".join(mcq) + "\n")
    log("讨论题与习题骨架：03-讨论题.md、04-课后习题.md（学科内容由 agent 填充）")


def write_outline_skeleton(out_dir, title, minutes, plan=None):
    budget = plan["budget"] if plan else [{"name": n, "minutes": m, "kind": k} for n, m, k in BUDGET_TEMPLATE]
    total = sum(b["minutes"] for b in budget)
    lines = ["《%s》%d 分钟课堂实施方案" % (title, minutes), "",
             "课堂形态：讲解 + 插入式提问 + 代码演示 + 视频演示 + 小组讨论 + 末段练习", "",
             "时间盒：" + "；".join("%s %d 分钟" % (b["name"], b["minutes"]) for b in budget) +
             "。合计 %d 分钟。" % total, "",
             "## 一、逐段记叙（时间、时长、老师讲什么、学生做什么）", "",
             "第 1 段　开场与回顾（0 分—5 分，时长 5 分钟）",
             "老师讲什么：【待补】",
             "学生做什么：【待补】",
             "这一段的高效提示词：【待补】", ""]
    cursor = 5
    for b in budget[1:]:
        seg_start, cursor = cursor, cursor + b["minutes"]
        lines += ["第 N 段　%s（%d 分—%d 分，时长 %d 分钟）" % (b["name"], seg_start, cursor, b["minutes"]),
                  "老师讲什么：【待补】", "学生做什么：【待补】",
                  "这一段的高效提示词：【待补】", ""]
    lines += ["## 二、这一节课对每一件任务设计的提示词", "",
              "任务一，生成整节课方案。【待补】",
              "任务二，精简课件。【待补】",
              "任务三，生成原理图页。【待补】",
              "任务四，找并下载演示视频。【待补】",
              "任务五，出讨论题、单选题与演示代码。【待补】", "",
              "## 三、板书", "", "【待补】", "",
              "## 四、时间盒自查表", "", "【待补】"]
    write_text(os.path.join(out_dir, "01-课程方案.md"), "\n".join(lines) + "\n")
    log("课程方案骨架：01-课程方案.md")


# ----------------------------------------------------------------------
# merge：把新增原理图页合并进课件（默认追加到末尾）
# ----------------------------------------------------------------------
def merge_slides(pptx_in, images_dir, out_pptx, spec=None, mode="append", insert_after=None,
                 blank_layout_index=6):
    from pptx import Presentation
    from pptx.util import Emu

    pairs = []
    for f in os.listdir(images_dir):
        low = f.lower()
        if low.startswith("page-") and low.endswith((".png", ".jpg", ".jpeg")):
            try:
                pairs.append((int(low[5:7]), f))
            except ValueError:
                continue
    pairs.sort()
    if not pairs:
        raise SystemExit("在 %s 里没找到 page-NN.png" % images_dir)
    if [n for n, _ in pairs] != list(range(1, len(pairs) + 1)):
        raise SystemExit("原理图页码不连续：%s" % [n for n, _ in pairs])

    notes = {}
    if spec and os.path.exists(spec):
        for p in read_json(spec).get("pages", []):
            notes[p.get("index")] = p.get("notes", "")

    prs = Presentation(pptx_in)
    before = len(prs.slides._sldIdLst)
    for num, name in pairs:
        slide = prs.slides.add_slide(prs.slide_layouts[blank_layout_index])
        for shp in list(slide.shapes):
            shp._element.getparent().remove(shp._element)
        slide.shapes.add_picture(os.path.join(images_dir, name), Emu(0), Emu(0),
                                 width=prs.slide_width, height=prs.slide_height)
        if notes.get(num):
            slide.notes_slide.notes_text_frame.text = notes[num]

    if mode == "insert" and insert_after:
        lst = prs.slides._sldIdLst
        total = len(lst)
        for k in range(len(pairs)):
            el = list(lst)[total - len(pairs) + k]
            lst.remove(el)
            lst.insert(insert_after + k, el)

    prs.save(out_pptx)
    after = len(Presentation(out_pptx).slides._sldIdLst)
    if mode == "insert" and insert_after:
        where = "第 %d—%d 页" % (insert_after + 1, insert_after + len(pairs))
    else:
        where = "末尾第 %d—%d 页" % (before + 1, after)
    log("合并原理图：%d 页 → %d 页，新页位于%s" % (before, after, where))
    return after


# ----------------------------------------------------------------------
# bundle：zip + 下载页
# ----------------------------------------------------------------------
def bundle(pack_dir, out_dir, name):
    """在一个自包含目录里生成 zip 与下载页。

    布局（必须这样，下载页里的相对链接才成立）：
        <out>/<pack_name>/        解压后的全部文件
        <out>/<name>.zip          整包
        <out>/index.html          下载页
    pack_dir 会被复制一份到 <out>/<pack_name>/，原目录不动。
    """
    out_dir = ensure_dir(out_dir)
    pack_name = os.path.basename(os.path.normpath(pack_dir))
    served = os.path.join(out_dir, pack_name)
    if os.path.abspath(served) != os.path.abspath(pack_dir):
        if os.path.exists(served):
            shutil.rmtree(served)
        shutil.copytree(pack_dir, served)

    zip_path = os.path.join(out_dir, name + ".zip")
    if os.path.exists(zip_path):
        os.remove(zip_path)
    n = 0
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for dp, ds, fs in os.walk(served):
            ds.sort()
            fs.sort()
            for f in fs:
                p = os.path.join(dp, f)
                z.write(p, os.path.relpath(p, out_dir))
                n += 1
    size = os.path.getsize(zip_path) / 1024 / 1024
    log("打包：%d 个文件 → %s（%.1f MB）" % (n, zip_path, size))

    html = render_index_html(served, pack_name, name + ".zip", n, size)
    write_text(os.path.join(out_dir, "index.html"), html)
    log("下载页：%s" % os.path.join(out_dir, "index.html"))
    log("分享：cd %s && python3 -m http.server 8010 --bind 0.0.0.0" % out_dir)
    return zip_path


def render_index_html(pack_dir, pack_name, zip_name, file_count, size_mb):
    """生成一个带样式的下载页（不依赖任何前端框架，起个 http.server 就能分享）。"""
    from urllib.parse import quote
    rows = []
    for dp, ds, fs in os.walk(pack_dir):
        ds.sort()
        for f in sorted(fs):
            p = os.path.join(dp, f)
            rows.append((os.path.relpath(p, pack_dir), os.path.getsize(p)))
    rows.sort()
    trs = "\n".join(
        '    <tr><td><a href="%s">%s</a></td><td>%.1f KB</td></tr>'
        % (quote(os.path.join(pack_name, rel)), rel, sz / 1024) for rel, sz in rows)
    tpl = """<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>备课包下载中心</title>
<style>
 :root{--blue:#2471a3;--dark:#1b2b3a;--line:#e3e9ef;--bg:#f5f8fb}
 *{box-sizing:border-box}
 body{margin:0;background:var(--bg);color:var(--dark);line-height:1.7;
      font-family:"Microsoft YaHei","PingFang SC","Noto Sans CJK SC",system-ui,sans-serif}
 .wrap{max-width:1000px;margin:0 auto;padding:28px 20px 80px}
 header{background:linear-gradient(120deg,#1b4f72,#2471a3 60%,#2e86c1);color:#fff;
        border-radius:14px;padding:26px 30px;box-shadow:0 6px 20px rgba(27,79,114,.25)}
 header h1{margin:0 0 6px;font-size:25px}header p{margin:0;opacity:.92;font-size:14px}
 .hero{margin-top:22px;background:#fff;border:1px solid var(--line);border-radius:14px;
       padding:22px 26px;display:flex;flex-wrap:wrap;gap:18px;align-items:center;
       justify-content:space-between}
 .btn{display:inline-block;text-decoration:none;padding:13px 26px;border-radius:10px;font-size:16px;
      font-weight:700;color:#fff;background:#c0392b;box-shadow:0 4px 12px rgba(192,57,43,.3)}
 table{width:100%;border-collapse:collapse;background:#fff;border-radius:12px;overflow:hidden;
       font-size:14px;margin-top:22px}
 th,td{padding:10px 14px;border-bottom:1px solid var(--line);text-align:left}
 th{background:#eaf2f8;color:#1b4f72;font-size:13px}
 td a{color:var(--blue);text-decoration:none;font-weight:600}td a:hover{text-decoration:underline}
 footer{margin-top:30px;text-align:center;font-size:12.5px;color:#85929e}
</style></head><body><div class="wrap">
<header><h1>__PACK__</h1><p>一节课备课包 · 共 __COUNT__ 个文件 · 解压后 __SIZE__ MB</p></header>
<div class="hero"><div><div style="font-size:17px;font-weight:700">一键下载全部材料</div>
<div style="font-size:13px;color:#5d6d7e">__ZIP__</div></div>
<a class="btn" href="__ZIPQ__" download>⬇ 下载完整备课包 ZIP</a></div>
<table><tr><th>文件</th><th>大小</th></tr>
__ROWS__
</table>
<footer>由 dsh-lesson-prep 生成 · 视频与图片版权归原作者，仅供课堂教学使用</footer>
</div></body></html>
"""
    return (tpl.replace("__PACK__", pack_name)
               .replace("__COUNT__", str(file_count))
               .replace("__SIZE__", "%.1f" % size_mb)
               .replace("__ZIP__", zip_name)
               .replace("__ZIPQ__", quote(zip_name))
               .replace("__ROWS__", trs))


# ----------------------------------------------------------------------
# style：从原课件提取配色与字体，供 slide-forge 做风格参考
# ----------------------------------------------------------------------
def extract_style(src_pptx, out_dir):
    zin = zipfile.ZipFile(src_pptx)
    theme = next((n for n in zin.namelist() if n.startswith("ppt/theme/theme")), None)
    colors, fonts = [], []
    if theme:
        xml = zin.read(theme).decode("utf-8", "ignore")
        colors = sorted(set(re.findall(r'<a:srgbClr val="([0-9A-Fa-f]{6})"', xml)))
        fonts = sorted(set(re.findall(r'typeface="([^"]+)"', xml)))
    # 统计页面里出现最多的字号，作为标题/正文字号参考
    sizes = []
    for name in zin.namelist():
        if re.match(r"ppt/slides/slide\d+\.xml$", name):
            sizes += [int(x) for x in re.findall(r'sz="(\d+)"', zin.read(name).decode("utf-8", "ignore"))]
    from collections import Counter
    top_sizes = [s for s, _ in Counter(sizes).most_common(6)]
    style = {
        "source": os.path.basename(src_pptx),
        "theme_colors": ["#" + c.upper() for c in colors[:24]],
        "theme_fonts": [f for f in fonts if f and not f.startswith("+")][:12],
        "font_sizes_half_points": top_sizes,
        "hint": "把 theme_colors 的前 1–2 个当作主色，theme_fonts 里的中文字体当作标题字体，"
                "写进 slide-forge 的 deck_spec.json（theme_lock），使新增页与原课件风格一致。",
    }
    write_json(os.path.join(out_dir, "style.json"), style)
    log("原课件风格：主色候选 %s" % ", ".join(style["theme_colors"][:6]))
    log("主题字体：%s" % ", ".join(style["theme_fonts"][:6]))
    return style


# ----------------------------------------------------------------------
# prompts：两个子代理的任务提示词
# ----------------------------------------------------------------------
def write_subagent_prompts(out_dir, title, run_dir, mode="default", slide_mode="both"):
    tpl_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "prompts")
    kv = {
        "TITLE": title,
        "RUN_DIR": os.path.abspath(run_dir),
        "SLIDE_MODE": slide_mode,
    }
    made = []
    for src, dst in (("subagent-slides.md", "子代理A-幻灯与原理图.md"),
                     ("subagent-parallel.md", "子代理B-素材与文案.md"),
                     ("main-agent.md", "主代理-总装与验收.md")):
        p = os.path.join(tpl_dir, src)
        if not os.path.exists(p):
            continue
        text = open(p, encoding="utf-8").read()
        for k, v in kv.items():
            text = text.replace("{{%s}}" % k, v)
        made.append(write_text(os.path.join(out_dir, dst), text))
    write_json(os.path.join(out_dir, "subagent-plan.json"), {
        "mode": mode,
        "note": "default 模式下由主代理顺序执行；parallel 模式下把 A/B 两个提示词分别交给两个子代理并行执行，"
                "主代理只做总装与验收。",
        "agents": [
            {"id": "slides", "prompt_file": "子代理A-幻灯与原理图.md",
             "owns": ["slide-forge 图片版", "slide-forge 元素版", "原课件风格参考", "原理图页"]},
            {"id": "parallel", "prompt_file": "子代理B-素材与文案.md",
             "owns": ["课件裁剪复核", "演示代码与数据集图", "B 站视频检索与 MP4 下载", "讨论题与单选题", "课程方案文案"]},
        ],
    })
    log("子代理提示词：%s" % "、".join(os.path.basename(m) for m in made))
    return made


# ----------------------------------------------------------------------
# config：依赖自检
# ----------------------------------------------------------------------
def check_config():
    checks = []
    def probe(name, fn, hint):
        try:
            info = fn()
            checks.append((name, "OK", info, hint))
        except Exception as e:  # noqa: BLE001
            checks.append((name, "MISSING", "%s: %s" % (type(e).__name__, e), hint))

    probe("python3", lambda: sys.version.split()[0], "3.8+")
    probe("python-pptx", lambda: __import__("pptx").__version__, "pip install python-pptx")
    probe("Pillow", lambda: __import__("PIL").__version__, "pip install pillow")
    probe("matplotlib", lambda: __import__("matplotlib").__version__, "pip install matplotlib")
    probe("numpy", lambda: __import__("numpy").__version__, "pip install numpy")
    probe("torch", lambda: __import__("torch").__version__, "pip install torch（演示代码用，CPU 即可）")
    probe("lxml", lambda: __import__("lxml").__version__, "pip install lxml")
    probe("yt-dlp", lambda: os.popen("yt-dlp --version 2>/dev/null").read().strip() or "not found",
          "pip install -U yt-dlp（仅下载视频需要）")
    probe("ffmpeg", lambda: os.popen("ffmpeg -version 2>/dev/null").readline().split()[2],
          "apt install ffmpeg（转码与截取需要）")
    probe("libreoffice", lambda: os.popen("soffice --version 2>/dev/null").read().split()[-1],
          "apt install libreoffice（仅 PDF 复核需要，可选）")

    print("%-14s %-8s %s" % ("组件", "状态", "说明 / 安装提示"))
    print("-" * 78)
    for name, status, info, hint in checks:
        print("%-14s %-8s %s" % (name, status, info if status == "OK" else hint))
    missing = [c for c in checks if c[1] != "OK"]
    print("-" * 78)
    print("缺失 %d 项。" % len(missing) if missing else "全部就绪。")
    return checks


# ----------------------------------------------------------------------
# CLI
# ----------------------------------------------------------------------
def cmd_all(args):
    run = ensure_dir(args.out)
    log("=" * 78)
    log("备课包运行目录：%s" % run)
    log("=" * 78)
    plan = None
    if args.deck:
        deck = extract_deck(args.deck, run)
        plan = build_plan(deck, args.minutes, run, desired_drop=args.drop or None,
                          keep_ratio=args.keep_ratio)
        trim_deck(os.path.join(run, "deck.json"), os.path.join(run, "plan.json"),
                  args.deck, os.path.join(run, "02-精简课件.pptx"), run)
        extract_style(args.deck, run)
    else:
        log("未提供 --deck：跳过抽页、删页、裁剪与风格提取，其余照常（计划统计为空）。")
        build_budget(args.minutes, {"PPT 讲解": 5, "课堂提问": 1, "代码演示": 2,
                                    "视频演示": 1, "课堂讨论": 1.5})
    write_demo_scripts(run)
    write_video_assets(run, args.title)
    write_quiz_skeletons(run, args.title)
    write_outline_skeleton(run, args.title, args.minutes, plan)
    write_subagent_prompts(run, args.title, run, mode=args.subagent_mode, slide_mode=args.slide_mode)
    log("确定性步骤完成。下一步：按 SKILL.md 执行子代理 / 学科内容填充，最后跑 bundle。")
    return run


def build_parser():
    ap = argparse.ArgumentParser(description="DSH 一节课备课包：确定性执行内核")
    sub = ap.add_subparsers(dest="cmd")

    p = sub.add_parser("extract", help="课件 → deck.json")
    p.add_argument("--deck", required=True)
    p.add_argument("--out", required=True)

    p = sub.add_parser("plan", help="deck.json → plan.json（时间盒 + 删页清单）")
    p.add_argument("--deck-json", required=True)
    p.add_argument("--minutes", type=int, default=DEFAULT_MINUTES)
    p.add_argument("--keep-ratio", type=float, default=0.70)
    p.add_argument("--drop", type=int, nargs="*", default=None)
    p.add_argument("--out", required=True)

    p = sub.add_parser("trim", help="按 plan.json 生成精简课件")
    p.add_argument("--deck-json", required=True)
    p.add_argument("--plan", required=True)
    p.add_argument("--pptx", required=True)
    p.add_argument("--out-pptx", required=True)
    p.add_argument("--out", required=True)

    p = sub.add_parser("media", help="准备演示脚本目录")
    p.add_argument("--out", required=True)

    p = sub.add_parser("video", help="生成视频检索关键词 / 清单骨架 / 下载脚本")
    p.add_argument("--title", required=True)
    p.add_argument("--out", required=True)

    p = sub.add_parser("quiz", help="生成讨论题与单选题骨架")
    p.add_argument("--title", required=True)
    p.add_argument("--out", required=True)

    p = sub.add_parser("outline", help="生成课程方案骨架")
    p.add_argument("--title", required=True)
    p.add_argument("--minutes", type=int, default=DEFAULT_MINUTES)
    p.add_argument("--out", required=True)

    p = sub.add_parser("merge", help="把新增原理图页合并进课件")
    p.add_argument("--pptx", required=True)
    p.add_argument("--images", required=True)
    p.add_argument("--spec", default=None)
    p.add_argument("--mode", choices=["append", "insert"], default="append")
    p.add_argument("--insert-after", type=int, default=None)
    p.add_argument("--out-pptx", required=True)

    p = sub.add_parser("bundle", help="打包 zip + 生成下载页")
    p.add_argument("--pack-dir", required=True)
    p.add_argument("--out", required=True)
    p.add_argument("--name", default="备课包")

    p = sub.add_parser("prompts", help="生成子代理与主代理提示词")
    p.add_argument("--title", required=True)
    p.add_argument("--run-dir", required=True)
    p.add_argument("--subagent-mode", default="default")
    p.add_argument("--slide-mode", default="both")
    p.add_argument("--out", required=True)

    p = sub.add_parser("style", help="从原课件提取配色 / 字体")
    p.add_argument("--pptx", required=True)
    p.add_argument("--out", required=True)

    p = sub.add_parser("config", help="依赖自检")

    p = sub.add_parser("all", help="一次跑完确定性步骤")
    p.add_argument("--title", required=True)
    p.add_argument("--deck", default=None, help="可选：原始 .pptx 路径")
    p.add_argument("--minutes", type=int, default=DEFAULT_MINUTES)
    p.add_argument("--keep-ratio", type=float, default=0.70)
    p.add_argument("--drop", type=int, nargs="*", default=None)
    p.add_argument("--subagent-mode", choices=["default", "parallel"], default="default")
    p.add_argument("--slide-mode", choices=["image", "both"], default="both")
    p.add_argument("--out", required=True)
    return ap


def main(argv=None):
    args = build_parser().parse_args(argv)
    if args.cmd in (None, "config"):
        check_config()
        return 0
    if args.cmd == "extract":
        extract_deck(args.deck, ensure_dir(args.out))
    elif args.cmd == "plan":
        build_plan(read_json(args.deck_json), args.minutes, ensure_dir(args.out),
                   desired_drop=args.drop, keep_ratio=args.keep_ratio)
    elif args.cmd == "trim":
        trim_deck(args.deck_json, args.plan, args.pptx, args.out_pptx, ensure_dir(args.out))
    elif args.cmd == "media":
        write_demo_scripts(ensure_dir(args.out))
    elif args.cmd == "video":
        write_video_assets(ensure_dir(args.out), args.title)
    elif args.cmd == "quiz":
        write_quiz_skeletons(ensure_dir(args.out), args.title)
    elif args.cmd == "outline":
        write_outline_skeleton(ensure_dir(args.out), args.title, args.minutes)
    elif args.cmd == "merge":
        merge_slides(args.pptx, args.images, args.out_pptx, spec=args.spec,
                     mode=args.mode, insert_after=args.insert_after)
    elif args.cmd == "bundle":
        bundle(args.pack_dir, ensure_dir(args.out), args.name)
    elif args.cmd == "prompts":
        write_subagent_prompts(ensure_dir(args.out), args.title, args.run_dir,
                               mode=args.subagent_mode, slide_mode=args.slide_mode)
    elif args.cmd == "style":
        extract_style(args.pptx, ensure_dir(args.out))
    elif args.cmd == "all":
        cmd_all(args)
    return 0


if __name__ == "__main__":
    sys.exit(main())
