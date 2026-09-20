---
name: lesson-prep-90
description: 把「教学章节标题」或「一份 .pptx 课件」变成一整套可直接上课的备课包：90 分钟课堂方案（纯文本散文体）、按课时精简后的 PPT、原课件风格的新增原理图页（slide-forge 图片版 + 元素版）、B 站演示视频（含 MP4 下载）、课堂讨论题（含参考解）、10 道单选题（含解析）、两个可运行带可视化的演示代码、数据集样例图，最后打包成 zip 与带样式的下载页。适用于高校教师备课、试讲、企业内训；不适用于只想要一份全新 PPT（用 slide-forge）或把图片版 PPT 转可编辑（用 image-to-editable-ppt）。
when-to-use: 用户说「这节课怎么上」「帮我把这个 PPT 精简到 90 分钟」「出一节课的备课包 / 教案 / 课件+习题+视频+代码」，或给出一个教学章节标题或 .pptx 课件，希望得到整套课堂材料。
---

# lesson-prep-90（一节课备课包）

把一份教学输入变成**可直接上课**的一套材料。四条铁律：

1. **时间先行**：先算出课时怎么分配（讲解 / 提问 / 讨论 / 代码 / 视频 / 练习各占多少），再倒推 PPT 该留几页；
2. **只删页、不改内容**：裁剪课件只做「删除多余页面」，不动任何文字、图片与版式；
3. **不编事实**：课件和用户材料里没有的数据、数字、结论一律写 `【待补】`，绝不编造；
4. **机器管结构、模型管学科**：页数、形状、保留率、题目结构由脚本校验；讲解词、题目、解析、代码由你写。

仓库里**不含**任何课件、图片、视频、运行产物和 API Key；所有输入由命令行给出，所有产物写到 `--out` 指定的目录。

---

## 一、什么时候用

- 用户给了一个章节标题（如「第 7 章 循环神经网络」）或一份 `.pptx`，要求出**一节课的教学方案**；
- 用户说「把这节课的 PPT 精简到 90 分钟」「找个短视频当课堂演示」「出 10 道单选题」「给两个能跑的演示代码」「把材料打成一个能下载的包」。

**不要用**的场景：只想生成一份全新 PPT → 用 `slide-forge`；已有图片版 PPT 想转可编辑 → 用 `image-to-editable-ppt`。

---

## 二、环境与依赖

先自检，缺什么按提示装：

```bash
python3 scripts/lesson_prep.py config
```

| 组件 | 必需性 | 用途 | 安装 |
| --- | --- | --- | --- |
| Python 3.8+ | 必需 | 跑全部脚本 | 系统自带或 conda |
| python-pptx | 必需 | 读写 PPTX、合并原理图页 | `pip install python-pptx` |
| Pillow / numpy / matplotlib | 必需 | 数据集样例图与演示可视化 | `pip install pillow numpy matplotlib` |
| lxml | 必需 | 解析 PPTX 主题配色 | `pip install lxml` |
| torch | 演示代码用 | 训练小模型（CPU 即可） | `pip install torch`（CPU 版足够） |
| yt-dlp | 下载视频用 | 从 B 站拉 MP4 | `pip install -U yt-dlp` |
| ffmpeg | 转码用 | 统一转 H.264、截取片段 | `apt install ffmpeg` |
| LibreOffice | 可选 | 把 PPTX 转 PDF 做页数复核 | `apt install libreoffice` |
| slide-forge 技能 | 要用原理图时必需 | 生成新增原理图页（图片版 + 元素版） | 见第六节 |

---

## 三、输入（只认这两样）

| 输入 | 必需性 | 说明 |
| --- | --- | --- |
| `--title` | 必需 | 章节 / 课时标题，例如「第 7 章 循环神经网络」。标题也会用于文件命名与视频检索关键词 |
| `--deck <path.pptx>` | 可选 | 原始课件。给了就抽页、删页、裁剪、提风格；不给则跳过这些步骤，其余照常 |
| `--minutes` | 可选，默认 90 | 课时；脚本按目标分钟等比缩放时间盒，合计必须恰好等于它 |
| `--keep-ratio` | 可选，默认 0.70 | 课件保留率目标，脚本会钳到 65%–75% |
| `--drop` | 可选 | 教师指定必须删除的页码，优先于算法判断 |
| `--subagent-mode` | 可选，默认 `default` | `default` 顺序执行；`parallel` 启用双子代理并行（见第五节） |
| `--slide-mode` | 可选，默认 `both` | 原理图页产出形态：`image` 只要图片版，`both` 图片版 + 元素版 |
| `--out` | 必需 | 运行目录，全部产物写在这里 |

---

## 四、一条命令跑完确定性步骤

```bash
python3 scripts/lesson_prep.py all \
  --title "第7章 循环神经网络" \
  --deck "/path/to/7. 循环神经网络 (1).pptx" \
  --out ./runs/rnn-90 --minutes 90 --subagent-mode parallel
```

`all` 依次完成：抽页 → 时间盒与删页清单 → 裁剪课件 → 提取原课件配色字体 → 准备演示脚本目录 → 生成视频检索关键词与下载脚本 → 生成讨论题与习题骨架 → 生成课程方案骨架 → 生成子代理提示词。

产物（都在 `--out` 下）：

| 文件 | 说明 |
| --- | --- |
| `deck.json` | 课件结构化大纲（每页标题、正文、是否代码页 / 分隔页 / 小结页） |
| `plan.json` | 时间盒（八段合计 = 课时）、删页清单与逐页理由、保留率、警告 |
| `02-精简课件.pptx` | 只删页、未改内容的精简课件 |
| `02-删页清单.md` | 每条删除的页号 + 原文标题 + 删除理由 |
| `style.json` | 从原课件提取的主题色与字体（供原理图页做风格参考） |
| `01-课程方案.md` | 课程方案骨架（分段、时间、待填内容） |
| `03-讨论题.md` / `04-课后习题.md` | 题目骨架（结构齐备、学科内容待填） |
| `06-视频/keywords.txt`、`推荐视频清单.md`、`下载视频.sh` | 视频检索与下载素材 |
| `demos/_common.py` | 演示代码公共工具（中文字体、数据集加载） |
| `子代理A-幻灯与原理图.md`、`子代理B-素材与文案.md`、`主代理-总装与验收.md` | 子代理与主代理提示词 |
| `subagent-plan.json` | 本轮采用哪种模式、两个子代理各负责什么 |

**分步执行**（要精细控制时）：

```bash
python3 scripts/lesson_prep.py extract --deck "x.pptx" --out ./runs/x
python3 scripts/lesson_prep.py plan    --deck-json ./runs/x/deck.json --minutes 90 --out ./runs/x
python3 scripts/lesson_prep.py trim    --deck-json ./runs/x/deck.json --plan ./runs/x/plan.json \
       --pptx "x.pptx" --out-pptx "./runs/x/02-精简课件.pptx" --out ./runs/x
python3 scripts/lesson_prep.py style   --pptx "x.pptx" --out ./runs/x
python3 scripts/lesson_prep.py merge   --pptx "./runs/x/02-精简课件.pptx" --images ./runs/x/slides \
       --spec ./runs/x/slides/deck_spec.json --mode append --out-pptx "./runs/x/07-循环神经网络-最终版-2.pptx"
python3 scripts/lesson_prep.py bundle  --pack-dir ./runs/x --out ./dist --name "循环神经网络-90分钟备课包"
```

---

## 五、子代理方案（默认不带子代理）

**默认 `--subagent-mode default`：不召唤任何子代理。** 由主代理顺序执行全部步骤，适合小课时、单页原理图、或想省上下文预算的场景。开工前只说一句提示、不追问：

> 默认不使用子代理：所有步骤由我顺序执行；若想并行加速，说「用子代理模式」。

### 开启方式（三种，任选其一）

1. 命令行：`--subagent-mode parallel`（`all` 子命令支持）；
2. 自然语言：用户说「**用子代理模式**」「并行跑」「两个子代理分工」；
3. 单独取提示词：
   ```bash
   python3 scripts/lesson_prep.py prompts --title "第7章 循环神经网络" \
       --run-dir ./runs/rnn-90 --subagent-mode parallel --slide-mode both --out ./runs/rnn-90
   ```
   然后把 `子代理A-幻灯与原理图.md` 与 `子代理B-素材与文案.md` 分别交给两个子代理。

### 两个子代理怎么分（互不写同一个文件）

| 子代理 | 负责 | 独占产物 |
| --- | --- | --- |
| **A · 幻灯与原理图** | 调 `slide-forge` 生成原理图页：先读 `style.json` 把原课件配色字体写进 `deck_spec.json` 的 `theme_lock`；出**图片版**（整页图 PPTX/PDF）并继续跑**元素版**（对象级可编辑 PPTX）；再用 `merge` 合进课件 | `slides/`（图片、逐页提示词、图片版 PPTX/PDF、元素版 PPTX）、`07-*-最终版-2.pptx` |
| **B · 素材与文案** | 演示代码（两个 `.py` + 数据集样例图）、B 站视频检索与 MP4 下载、讨论题与参考解、10 道单选题与解析、课程方案正文（纯文本散文体）、验收自检 | `demos/`、`06-视频/`、`01-课程方案.md`、`03-讨论题.md`、`04-课后习题.md`、`08-数据集图片/` |

主代理只做三件事：跑 `all` 的确定性步骤、把两份提示词交给两个子代理、最后跑 `merge` + `bundle` 并验收。

### 为什么这样切不影响质量

- **不共享写入目标**：A 只写 `slides/` 与最终课件，B 只写文档 / 代码 / 视频 / 图片，双方都不改对方的产物，合并时不会互相覆盖；
- **有先后依赖的步骤留在主代理**：裁剪课件先于原理图（A 依赖裁剪结果），合并在两个子代理都完成之后（主代理串行做）；
- **并行的是"无依赖且耗时"的部分**：B 的代码运行、视频下载、题目写作与 A 的逐页生图 / 元素版重建完全无交集，这是加速的全部来源；
- **质量闸门不并行**：`qa` 门与人工确认点仍然串行、由主代理把关，不因为并行而跳过。

### 召唤子代理时用的提示词范例

```
用子代理模式跑这一节课的备课包，主代理负责总装：
- 子代理 A 负责调用 slide-forge 生成 3 页原理图（图片版 + 元素版），风格参考原课件 style.json；
- 子代理 B 负责演示代码、数据集图、B 站视频下载、讨论题与 10 道单选题、课程方案正文；
- 两边都完成后由主代理合并课件并打包。
```

更完整的角色提示词见 `prompts/main-agent.md`、`prompts/subagent-slides.md`、`prompts/subagent-parallel.md`。

---

## 六、原理图页与 slide-forge 的分工

原理图页解决的是「原课件只有静态结构图和公式、学生第一次看不懂信息怎么流动」这个缺口。流程：

1. 读 `style.json`（由 `style` 子命令从原课件提取的主题色与字体）；
2. 写 `slides/deck_spec.json`：3 页左右，`preset` 用 `clean-corporate` 或 `academic-paper`，并把 `style.json` 的前 1–2 个主色写进 `theme_lock.accents`、中文字体写进 `theme_lock.typography`，使新增页与原课件一致；
3. 图片版：`python3 <slide-forge>/scripts/deck.py gen|qa|pptx --spec slides/deck_spec.json --out-dir slides`；
4. 元素版：`python3 <slide-forge>/scripts/editable_run.py prepare|...`（详见 slide-forge 技能；元素版产出的 PPTX 文字 / 形状可单独选中修改）；
5. 合并：`lesson_prep.py merge`。默认 `--mode append` 把新页追加到末尾（最稳、绝不破坏原页序）；想插到某一节之前用 `--mode insert --insert-after <页码>`。

---

## 七、验收判据（收尾必跑）

| 产出 | 判据 |
| --- | --- |
| 课程方案 | 含讲解 / 提问 / 讨论 / 代码 / 视频五类活动；时段合计 = 课时；末段留 5 分钟练习且**不写题目**；正文为连续散文，**无表格、无分条** |
| 精简课件 | 页数 = 原页数 − 删除数；保留率 65%–75%；封面 / 资源页 / 分隔页 / 小结页 / 每个小节首页未被删；无空白页 |
| 原理图页 | 每页一个知识点、图为流程图或原理示意图；主色与原课件一致；文字无错字乱码；图片版与元素版都能打开 |
| 视频 | 至少 1 条 ≤5 分钟；推荐清单含标题 / UP 主 / BV 号 / 时长 / 链接 / 推荐理由；下载脚本可执行；MP4 为 H.264 |
| 讨论题 | 背景 / 分组 / 时间盒 / 参考解 / 评分标准齐备 |
| 单选题 | 恰好 10 题、每题 4 选项、唯一正确答案、解析非空 |
| 演示代码 | 两个 `.py` 能在 CPU 上数十秒跑完并出图；有中文注释与可视化 |
| 数据集图 | 至少 1 张带类别标注的真实样本图；标注来源与许可 |
| 打包 | zip 可解压、文件名不乱码；`index.html` 里每条相对链接都能打开 |

自检命令：

```bash
python3 scripts/verify_pack.py --pack-dir ./runs/rnn-90
```

---

## 八、失败与降级

| 情况 | 行为 |
| --- | --- |
| 只给标题、没有课件 | 跳过抽页 / 删页 / 裁剪 / 风格提取，其余照常，`plan.stats` 为空 |
| 课件打不开 / 加密 / 非 pptx | 明确报错并中止，不产出半成品 |
| B 站搜索 412 或断网 | 保留关键词与清单骨架，提示换关键词重搜；其余产物不受影响 |
| 找不到 ≤5 分钟视频 | 提示换关键词；推荐清单与下载脚本照常产出 |
| 缺 yt-dlp | 只影响下载，不影响清单 |
| 无 LibreOffice | 跳过 PDF 页数复核，PPTX 照常交付 |
| 无 GPU | 演示代码走 CPU 小样本，数十秒跑完 |
| 无生图密钥 / 未配 slide-forge | 原理图页停在 `deck_spec.json` 与提示词，交给用户补密钥后重跑；其余产物照常 |
| 元素版合并校验失败 | 如实报告并降级为逐页交付 `page.pptx`，不伪造校验结果 |

---

## 九、红线（不要做）

- 不把用户课件、图片、视频、运行产物、API Key 写进本仓库或提交到 git；
- 不改写课件原有页面的文字、图片、版式（裁剪只删页）；
- 不编造数据、引用、结论；不确定的写 `【待补】`；
- 不跨页复用同一张原理图；不伪造 QA 或校验结果；
- 用户没要求时**不召唤子代理**，也不要为「是否用子代理」反复追问。
