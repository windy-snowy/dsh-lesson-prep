#!/usr/bin/env bash
# 最小试跑：用当前目录里的任意 pptx 做一次完整备课包（默认 90 分钟）
# 用法：bash examples/quickstart.sh /path/to/deck.pptx "第7章 循环神经网络" [out_dir]
set -euo pipefail
DECK="${1:?用法: bash examples/quickstart.sh <deck.pptx> <标题> [out_dir]}"
TITLE="${2:?请给标题}"
OUT="${3:-./runs/$(basename "$DECK" .pptx)}"

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

echo "== 1/4 环境自检"
python3 "$ROOT/scripts/lesson_prep.py" config

echo "== 2/4 确定性步骤（抽页 / 时间盒 / 裁剪 / 风格 / 骨架）"
python3 "$ROOT/scripts/lesson_prep.py" all --title "$TITLE" --deck "$DECK" --out "$OUT" --minutes 90

echo "== 3/4 验收（此时原理图与学科内容还没填，会有 warn）"
python3 "$ROOT/scripts/verify_pack.py" --pack-dir "$OUT" || true

echo "== 4/4 打包 + 下载页"
python3 "$ROOT/scripts/lesson_prep.py" bundle --pack-dir "$OUT" --out "$OUT/../dist" --name "$(basename "$TITLE")-90分钟备课包"
echo
echo "完成后按 SKILL.md 执行子代理，或直接对 agent 说：用子代理模式继续"
