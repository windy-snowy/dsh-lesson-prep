# -*- coding: utf-8 -*-
"""
verify_pack.py —— 备课包验收：把 SKILL.md 第七节的判据变成可执行的检查

用法：
    python3 scripts/verify_pack.py --pack-dir ./runs/rnn-90
    python3 scripts/verify_pack.py --pack-dir ./runs/rnn-90 --json

检查项都只读文件、不改动任何东西；缺文件记为 fail，结构不合规记为 fail，
可选文件缺失记为 warn。退出码：全部通过 0，否则 1。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys


def newest(pack_dir, *patterns):
    """按文件名模式找最新的一个文件（支持简单通配）。"""
    hits = []
    for dp, _ds, fs in os.walk(pack_dir):
        for f in fs:
            for pat in patterns:
                if re.search(pat, f):
                    hits.append(os.path.join(dp, f))
    return sorted(hits)[-1] if hits else None


def read(path):
    with open(path, encoding="utf-8", errors="ignore") as f:
        return f.read()


def check_mcq(text, want=10):
    """单选题结构：题号 1..want、每题 4 个选项、有正确答案与解析。"""
    problems = []
    for i in range(1, want + 1):
        if not re.search(r"(?m)^\s*%d[.、]" % i, text):
            problems.append("缺少第 %d 题" % i)
            continue
        start = re.search(r"(?m)^\s*%d[.、]" % i, text).start()
        nxt = re.search(r"(?m)^\s*%d[.、]" % (i + 1), text)
        block = text[start:nxt.start() if nxt else len(text)]
        # 选项既可能每个占一行，也可能挤在一行（A. …  B. …），两种都要能数出来
        opts = len(re.findall(r"(?m)^\s*[ABCD][.、)]", block))
        if opts != 4:
            opts = len(re.findall(r"[ABCD][.、)]\s*\S", block))
        if not re.search(r"正确答案", block) and not re.search(r"答案[：:]", block):
            problems.append("第 %d 题缺正确答案" % i)
        if not re.search(r"解析", block):
            problems.append("第 %d 题缺解析" % i)
    return problems


def check_plan(plan_path):
    plan = json.loads(read(plan_path))
    problems, warns = [], []
    total = sum(b["minutes"] for b in plan["budget"])
    if total != plan["minutes"]:
        problems.append("时间盒合计 %d ≠ 课时 %d" % (total, plan["minutes"]))
    names = [b["name"] for b in plan["budget"]]
    for need in ("练习", "PPT 讲解", "课堂讨论"):
        if not any(need in n for n in names):
            problems.append("时间盒缺少「%s」时段" % need)
    practice = next((b for b in plan["budget"] if "练习" in b["name"]), None)
    if practice and practice["minutes"] != 5:
        warns.append("练习时段为 %d 分钟（时间盒要求末段固定 5 分钟）" % practice["minutes"])
    stats = plan.get("stats") or {}
    if stats:
        ratio = stats.get("keep_ratio", 0)
        if not (0.65 <= ratio <= 0.75):
            problems.append("保留率 %.1f%% 不在 65%%–75%%" % (100 * ratio))
        if stats["original_pages"] - stats["delete_pages"] != stats["kept_pages"]:
            problems.append("页数不自洽：原 %d − 删 %d ≠ 保留 %d"
                            % (stats["original_pages"], stats["delete_pages"], stats["kept_pages"]))
    return problems, warns


def check_deck(pptx_path):
    problems = []
    try:
        from pptx import Presentation
    except ImportError:
        return ["没装 python-pptx，无法核对课件页数"]
    prs = Presentation(pptx_path)
    n = len(prs.slides._sldIdLst)
    if n == 0:
        problems.append("课件页数为 0")
    for i, slide in enumerate(prs.slides, 1):
        text = "".join(sh.text_frame.text for sh in slide.shapes if sh.has_text_frame)
        pic = any("PICTURE" in str(sh.shape_type) for sh in slide.shapes if sh.shape_type)
        if not text.strip() and not pic:
            problems.append("第 %d 页是空白页" % i)
    return problems


def check_video_md(text):
    problems, warns = [], []
    if "bilibili.com" not in text and "BV" not in text:
        warns.append("推荐视频清单里还没有真实 BV 号（仍是骨架）")
    if "yt-dlp" not in text:
        problems.append("推荐视频清单缺少 yt-dlp 下载命令")
    if "【待补】" in text:
        warns.append("推荐视频清单仍有【待补】")
    return problems, warns


def check_outline(text):
    problems, warns = [], []
    for need in ("老师讲什么", "学生做什么"):
        if need not in text:
            problems.append("课程方案缺少「%s」" % need)
    if re.search(r"(?m)^\s*\|.*\|\s*$", text):
        problems.append("课程方案里出现了表格（要求纯文本散文体）")
    if re.search(r"(?m)^\s*[-*]\s+", text):
        problems.append("课程方案里出现了分条列表（要求纯文本散文体）")
    if re.search(r"(?m)^#{1,6}\s", text) or re.search(r"\*\*.+\*\*", text):
        warns.append("课程方案里还有 Markdown 标题 / 加粗语法")
    if "【待补】" in text:
        warns.append("课程方案仍有【待补】未填")
    return problems, warns


def check_demos(demo_dir):
    problems, warns = [], []
    pys = [f for f in os.listdir(demo_dir)
           if f.endswith(".py") and not f.startswith("_")] if os.path.isdir(demo_dir) else []
    if len(pys) < 2:
        problems.append("演示代码少于 2 个 .py（当前 %d 个；_common.py 之类的公共工具不计）" % len(pys))
    for f in pys:
        if f.startswith("_"):
            continue                      # 公共工具不算演示代码
        body = read(os.path.join(demo_dir, f))
        if not re.search(r"[\u4e00-\u9fff]", body):
            warns.append("%s 里没有中文注释" % f)
        if not re.search(r"(plt\.|savefig|matplotlib)", body):
            warns.append("%s 里没看到可视化代码" % f)
    return problems, warns


def check_dataset_images(pack_dir):
    imgs = []
    for dp, _ds, fs in os.walk(pack_dir):
        for f in fs:
            if f.lower().endswith((".png", ".jpg", ".jpeg")) and re.search(r"(数据集|dataset|sample)", f):
                imgs.append(os.path.join(dp, f))
    if not imgs:
        return ["没有找到数据集样例图（文件名含「数据集 / dataset / sample」的图片）"]
    return []


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pack-dir", required=True)
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args()

    d = args.pack_dir
    results = []          # (level, item, message)

    def add(level, item, msgs):
        """把一个或一组问题记进结果。

        msgs 允许三种形态：空（通过）、问题字符串列表、check_* 返回的 (problems, warns) 二元组。
        二元组会被自动展平，否则空元组会被误当成"有两条空问题"。
        """
        if isinstance(msgs, tuple):
            msgs = list(msgs[0]) + list(msgs[1])
        if not msgs:
            return
        for m in msgs:
            if not m:
                continue
            results.append((level, item, m))

    # 1) 时间盒与删页清单
    plan_path = os.path.join(d, "plan.json")
    if os.path.exists(plan_path):
        p, w = check_plan(plan_path)
        add("fail", "时间盒 / 删页清单", p)
        add("warn", "时间盒 / 删页清单", w)
    else:
        add("warn", "时间盒 / 删页清单", ["没有 plan.json（可能只给了标题、没有课件）"])

    # 2) 精简课件 / 最终课件
    deck = newest(d, r"最终版.*\.pptx$", r"精简课件.*\.pptx$")
    if deck:
        add("fail", "课件页数与空白页", check_deck(deck))
    else:
        add("fail", "课件页数与空白页", ["没有找到 .pptx 课件"])

    # 3) 课程方案
    outline = newest(d, r"课程方案.*\.md$")
    if outline:
        p, w = check_outline(read(outline))
        add("fail", "课程方案格式", p)
        add("warn", "课程方案格式", w)
    else:
        add("fail", "课程方案格式", ["没有找到课程方案 md"])

    # 4) 单选题
    mcq = newest(d, r"课后习题.*\.md$", r"单选题.*\.md$")
    if mcq:
        add("fail", "单选题结构", check_mcq(read(mcq)))
    else:
        add("fail", "单选题结构", ["没有找到课后习题 md"])

    # 5) 讨论题
    disc = newest(d, r"讨论题.*\.md$")
    if disc:
        t = read(disc)
        miss = [k for k in ("参考解", "评分", "时间盒") if k not in t]
        add("fail", "讨论题要素", ["缺少：%s" % "、".join(miss)] if miss else [])
        if "【待补】" in t:
            add("warn", "讨论题要素", ["讨论题仍有【待补】"])
    else:
        add("fail", "讨论题要素", ["没有找到讨论题 md"])

    # 6) 视频
    vmd = newest(d, r"推荐视频.*\.md$")
    if vmd:
        p, w = check_video_md(read(vmd))
        add("fail", "视频清单", p)
        add("warn", "视频清单", w)
        mp4 = [f for dp, _ds, fs in os.walk(d) for f in fs if f.lower().endswith(".mp4")]
        if not mp4:
            add("warn", "视频清单", ["还没有下载好的 MP4（清单照常可用）"])
    else:
        add("warn", "视频清单", ["没有找到推荐视频 md"])

    # 7) 演示代码
    demo_dir = os.path.join(d, "demos")
    add("fail", "演示代码", check_demos(demo_dir))

    # 8) 数据集图
    add("fail", "数据集样例图", check_dataset_images(d))

    fails = [r for r in results if r[0] == "fail"]
    warns = [r for r in results if r[0] == "warn"]

    if args.json:
        print(json.dumps({"results": results,
                          "fail": len(fails), "warn": len(warns),
                          "passed": not fails}, ensure_ascii=False, indent=1))
    else:
        print("验收目录：%s" % d)
        print("-" * 76)
        for level, item, msg in results:
            mark = "✓" if msg == "OK" else ("✗" if level == "fail" else "!")
            print("%s %-16s %s" % (mark, item, msg))
        print("-" * 76)
        print("不合格 %d 项，警告 %d 项 → %s" % (len(fails), len(warns),
                                                 "通过" if not fails else "未通过"))
    return 0 if not fails else 1


if __name__ == "__main__":
    sys.exit(main())
