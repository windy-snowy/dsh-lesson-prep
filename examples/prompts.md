# 输入 prompt 范例库

把范例里的话题和路径换掉即可直接使用。所有范例都**不要求**你准备任何额外文件：
给课件就多出"精简课件 + 风格一致的原理图"，只给标题也能出整套材料。

---

## 1. 最小可用：只给标题

```
给「第4章 卷积神经网络」出一节课的备课包。
```

产出：90 分钟课程方案、讨论题、10 道单选题、两个演示代码、数据集图、打包下载页。
（没有抽页 / 裁剪 / 原理图，因为没给课件。）

## 2. 标准用法：课件 + 90 分钟

```
帮我把这份课件做成 90 分钟的备课包：
/path/to/7. 循环神经网络 (1).pptx
要求课上要有提问、讨论、代码演示、视频演示、PPT 讲解，最后留 5 分钟练习。
所有文件打成一个可以下载的包，给出链接。
```

## 3. 并行加速 + 原理图页 + 风格对齐

```
用子代理模式跑这节课的备课包，主题「第7章 循环神经网络」，
课件 /path/to/7. 循环神经网络 (1).pptx，90 分钟。
原课件第一节的原理部分只有静态图和公式，请加 3 页原理图，用流程图讲清原理，
风格跟原课件一致（配色和字体照 style.json 来）。
原理图要图片版和元素版都要；最后打包并给出下载链接。
```

## 4. 换课时（60 分钟）

```
给「第5章 迁移学习」做 60 分钟的备课包，课件在 /path/to/ch5.pptx。
时间要紧一点，代码演示保留两个但各缩短；最后仍然留 5 分钟练习。
```

## 5. 指定删页与保留要求

```
把 /path/to/ch3.pptx 精简到 90 分钟能讲完，保留原页数的 70% 左右。
第 12、13、20 页我确定用不到，请一定删掉；
每个小节至少留两页，分隔页和本章小结不要删。
```

## 6. 只要原理图页（不打包）

```
参考 /path/to/ch7.pptx 的风格，给「LSTM 的三个门」生成 3 页原理图，
图片版 + 元素版都要，元素版要能单独改文字。
```

## 7. 只要题目

```
根据 /path/to/ch7.pptx 出 10 道单选题（含解析）和 1 道课堂讨论题（含参考解、评分标准）。
格式要纯文本段落，不要表格。
```

## 8. 英文 prompt

```
Build a 90-minute lesson prep pack from /path/to/chapter7.pptx.
Include: lecture plan with timings, a trimmed deck (keep ~70% of pages),
3 principle-diagram slides styled like the original deck, two runnable demo scripts,
a discussion question with a reference answer, 10 single-choice questions with explanations,
two short demo videos (<=5 min), dataset sample figures, and a downloadable zip with an index page.
Use subagent mode for speed.
```

---

## 命令行等价写法

```bash
# 对应范例 2
python3 scripts/lesson_prep.py all --title "第7章 循环神经网络" \
  --deck "/path/to/7. 循环神经网络 (1).pptx" --out ./runs/rnn-90 --minutes 90

# 对应范例 3（并行 + 原理图图片版与元素版）
python3 scripts/lesson_prep.py all --title "第7章 循环神经网络" \
  --deck "/path/to/7. 循环神经网络 (1).pptx" --out ./runs/rnn-90 \
  --minutes 90 --subagent-mode parallel --slide-mode both

# 对应范例 5（指定删页）
python3 scripts/lesson_prep.py all --title "第3章" --deck "/path/to/ch3.pptx" \
  --out ./runs/ch3 --keep-ratio 0.70 --drop 12 13 20
```

## 参数速查

| 参数 | 默认 | 说明 |
| --- | --- | --- |
| `--title` | 必填 | 章节 / 课时标题 |
| `--deck` | 无 | 原始 `.pptx`；不给就跳过抽页与裁剪 |
| `--minutes` | 90 | 课时；时间盒合计恰好等于它 |
| `--keep-ratio` | 0.70 | 课件保留率（钳到 65%–75%） |
| `--drop` | 无 | 必删页码，空格分隔 |
| `--subagent-mode` | `default` | `parallel` 启用双子代理 |
| `--slide-mode` | `both` | `image` 只要图片版 |
| `--out` | 必填 | 运行目录 |
