# 流水线：八个阶段与各自的责任边界

本文件说明 `lesson_prep.py` 每个子命令做什么、输入输出是什么、失败怎么降级。
判据见 [quality-gates.md](quality-gates.md)，子代理分工见 [subagent-design.md](subagent-design.md)。

---

## 阶段 0 · 环境自检（`config`）

探测 Python、python-pptx、Pillow、matplotlib、numpy、torch、lxml、yt-dlp、ffmpeg、LibreOffice。
缺必需项就明确提示安装命令；缺可选项只降级对应能力（例如缺 LibreOffice 就不做 PDF 页数复核）。

## 阶段 1 · 抽页（`extract`）

输入 `.pptx`，输出 `deck.json`。对每一页记录：

- `title` / `text` / `text_chars`：标题与全文；
- `looks_code` / `code_score`：是否是整页代码（按 `def`、`class`、`nn.`、`self.` 等词频判断）；
- `is_divider` / `is_summary` / `is_structural`：分隔页（含「本章内容」「contents」且文字很短）、
  小结页（标题以「小结/总结」结尾）、结构性页面；
- `picture_count`：图片数量，用于识别「以截图为主」的页；
- `section_no`：从标题里抽出的节号（如 `7.1.1`）。

`group_sections` 按**节号变化**切小节，同一节号的连续页归入一节，分隔页单独成节。
不依赖任何模板或命名约定，换教材也能用。

**降级**：课件打不开 / 加密 / 非 pptx → 明确报错并中止，不产出半成品。

## 阶段 2 · 时间盒与删页清单（`plan`）

### 时间盒

八段：开场与回顾、PPT 讲解、课堂提问、代码演示、视频演示、课堂讨论、练习、收尾总结。
90 分钟时的默认分配是 **5 / 38 / 7 / 18 / 6 / 10 / 5 / 1**；其他课时按比例缩放，
并保证：练习固定 5 分钟、总时长恰好等于 `--minutes`、无负值。

### 删页

1. **硬保护**（永不删除）：前 3 页（封面、教材资源、章标题）、小结页、所有分隔页、
   每个小节的**前 2 页**（防腰斩）。
2. **候选打分**（越高越该删）：与前面的页正文重复 +100；整页代码 +40；
   出现「运行结果 / 输出如下 / 执行上述代码」+30；正文 <120 字 +25；以截图为主 +10。
3. **目标保留率**：默认 70%，钳制在 65%–75% 区间内换算成删除页数。
4. 教师用 `--drop` 指定的页号优先删除，不受打分影响（但仍不能删硬保护页）。

输出 `plan.json`：`budget`、`ppt_minutes`、`minutes_per_page`、`delete_list.delete_pages`、
`delete_list.reasons`（逐页理由）、`protected_pages`、`warnings`。

**降级**：`minutes_per_page < 0.8` 且课时 ≥60 时给 warning（讲解只抓主线）；
保留率越界也给 warning 而不是静默通过。

## 阶段 3 · 裁剪课件（`trim`）

只做一件事：**从 page 序列里摘掉指定页**，并连带清理被删页面的 `slide` 部件、备注页与
仅被它们引用的媒体文件。改写的是 `ppt/presentation.xml` 的 `sldIdLst` 与
`ppt/_rels/presentation.xml.rels`，**不触碰任何 slide 的 XML 内容**，所以文字、图片、
版式一定不变。

同时输出 `02-删页清单.md`（每条删除的页号 + 原文标题 + 理由），便于教师复核；
想保留某页就从 `plan.json` 里去掉该页号重跑。

## 阶段 4 · 风格提取（`style`）

从 `ppt/theme/theme*.xml` 抽 `srgbClr` 与 `typeface`，从所有 `slide*.xml` 统计出现最多的字号，
写进 `style.json`。它的唯一用途是给原理图页做风格参考：把主色写进 `deck_spec.json` 的
`theme_lock.accents`，把中文字体写进 `typography.cjk_font`。

**降级**：课件没有 theme 部件时输出空列表并提示「由 slide-forge 预设决定配色」。

## 阶段 5 · 原理图页（`slide-forge` + `merge`）

由子代理 A 执行（详见 [subagent-design.md](subagent-design.md)）：

1. 写 `slides/deck_spec.json`（3 页，每页一个原理，`preset` 走浅色学术风，配色取自 `style.json`）；
2. `slide-forge` 出图片版（整页图 PPTX/PDF），`--slide-mode both` 时再出元素版（对象级可编辑）；
3. `merge` 把 `page-NN.png` 合成新页写进课件：
   - `--mode append`（默认）：追加到末尾，**原页序完全不变**，最稳；
   - `--mode insert --insert-after N`：插到第 N 页之后，用于「原理图放进第一节」；
   - 若 `deck_spec.json` 里有 `pages[].notes`，同步写进讲者备注。

**降级**：缺生图密钥 → 停在 `deck_spec.json` 与逐页提示词；元素版校验失败 → 如实报告并逐页交付。

## 阶段 6 · 素材与文案（子代理 B）

六个子步骤：复核裁剪结果（只读）、两个演示脚本、数据集样例图、B 站视频检索与 MP4 下载、
讨论题与 10 道单选题、课程方案正文。全部要求见 `prompts/subagent-parallel.md`。

**降级**：B 站搜索 412 / 断网 → 保留关键词与清单骨架；找不到 ≤5 分钟视频 → 提示换关键词；
缺 yt-dlp → 只影响下载；数据集不可下载 → 演示代码回退内置 / 合成数据。

## 阶段 7 · 打包（`bundle`）

生成自包含目录：`<out>/<pack_name>/`（全部文件）+ `<out>/<name>.zip` + `<out>/index.html`。
下载页是纯静态 HTML + CSS，带整包下载按钮与逐文件列表，相对链接指向同目录下的
`<pack_name>/`，因此**必须**保持这个布局，起 `python3 -m http.server` 就能分享。

zip 用 Python 的 `zipfile` 写入 UTF-8 文件名，Windows 解压不乱码。

## 阶段 8 · 验收（`verify_pack.py`）

把 [quality-gates.md](quality-gates.md) 的判据变成可执行检查：时间盒合计、保留率、
课件空白页、课程方案是否含表格/分条、单选题结构与题数、讨论题要素、视频清单、
演示代码数量与可视化、数据集图。输出 fail / warn 两级，`fail` 非零退出。

**fail 必须修**；**warn 必须如实告诉用户**，不能吞掉。
