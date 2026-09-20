# -*- coding: utf-8 -*-
"""
check_download_page.py —— 校验 bundle 产出的下载页：每条链接都要真的能打开

用法：
    python3 scripts/check_download_page.py --dist-dir /path/to/dist
    python3 scripts/check_download_page.py --dist-dir /path/to/dist --json

背景：下载页里的 href 是 URL 编码过的（中文文件名必须编码），所以校验时必须先
`urllib.parse.unquote` 再判断文件存在，否则会把「中文名编码后与磁盘文件名不一致」
误判成断链（这是 CI 第一次跑红的原因）。

退出码：全部链接可用 0，否则 1。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import urllib.parse


def check(dist_dir):
    index = os.path.join(dist_dir, "index.html")
    problems, links = [], []
    if not os.path.exists(index):
        return (["没有 index.html：" + dist_dir], [])
    html = open(index, encoding="utf-8").read()
    for href in re.findall(r'(?:href|src)="([^"]+)"', html):
        if href.startswith(("http://", "https://", "#", "mailto:")):
            continue
        target = urllib.parse.unquote(href.split("?")[0].split("#")[0])
        path = os.path.join(dist_dir, target)
        exists = os.path.isfile(path) or os.path.isdir(path)
        links.append((href, target, exists))
        if not exists:
            problems.append("链接打不开：%s（解码后 %s）" % (href, target))
    if not links:
        problems.append("下载页里没有任何文件链接")
    # 整包 zip 必须存在且非空
    zips = [f for f in os.listdir(dist_dir) if f.lower().endswith(".zip")]
    if not zips:
        problems.append("目录里没有 zip 整包")
    elif all(os.path.getsize(os.path.join(dist_dir, z)) == 0 for z in zips):
        problems.append("zip 整包是空文件")
    return problems, links


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dist-dir", required=True)
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args()

    problems, links = check(args.dist_dir)
    if args.json:
        print(json.dumps({"dist_dir": args.dist_dir, "links": len(links),
                          "problems": problems, "passed": not problems},
                         ensure_ascii=False, indent=1))
    else:
        print("下载页自检：%s" % args.dist_dir)
        print("  链接 %d 条，其中失效 %d 条" % (len(links), len(problems)))
        for p in problems:
            print("  ✗ " + p)
        print("  → %s" % ("通过" if not problems else "未通过"))
    return 0 if not problems else 1


if __name__ == "__main__":
    sys.exit(main())
