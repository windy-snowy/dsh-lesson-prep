# dsh-lesson-prep ｜ DSH 插件：一节课备课包生成器

把「一个教学章节标题」或「一份 `.pptx` 课件」变成一整套**可直接上课**的备课包：
课堂方案、按课时精简的课件、原课件风格的新增原理图页、演示视频（含 MP4 下载）、
讨论题与参考解、10 道单选题与解析、两个可运行带可视化的演示代码、数据集样例图，
最后打包成 zip 与带样式的下载页。

这是一个 **DSH 技能（skill）插件**：`SKILL.md` 是本体，`scripts/` 是可执行的确定性内核，
`prompts/` 是两个子代理与主代理的角色提示词，`references/` 是流程与判据说明。

- 默认 **90 分钟**课时（`--minutes` 可改，时间盒会等比缩放且合计恰好等于课时）
- 默认 **不召唤子代理**；说「用子代理模式 / 并行跑」即启用**双子代理并行**（见下）
- 仓库内**不含**任何课件、图片、视频、运行产物与 API Key

---

## 一、安装（3 种方式，任选其一）

### 方式 1：DSH 用户技能目录（推荐）

DSH 会扫描 `<DSH_HOME>/skills`（`DSH_HOME` 默认 `~/.dsh`），把技能目录放进去即可被识别：

```bash
git clone https://github.com/<you>/dsh-lesson-prep.git ~/.dsh/skills/lesson-prep-90
# 或指定 DSH_HOME
DSH_HOME=/path/to/dsh-home git clone https://github.com/<you>/dsh-lesson-prep.git "$DSH_HOME/skills/lesson-prep-90"
```

DSH 带目录监听，新技能无需重启就会被扫到（若你的配置关了 watch，重开会话即可）。

### 方式 2：当前项目的技能目录

在任意 git 项目里：

```bash
mkdir -p .dsh/skills
git clone https://github.com/<you>/dsh-lesson-prep.git .dsh/skills/lesson-prep-90
```

### 方式 3：自定义技能根目录

如果你用 `@deepseek-ai/dsh-skill-filesystem` 的 `customSkillDirs` 指定了别的目录，把本仓库克隆进去即可。

> 目录名可以随意（技能名取自 `SKILL.md` frontmatter 的 `name: lesson-prep-90`）。
> 技能正文按需加载，不会常驻上下文。

## 二、依赖

```bash
python3 scripts/lesson_prep.py config      # 一键自检，缺什么会告诉你装什么
```

| 组件 | 必需性 | 用途 | 安装 |
| --- | --- | --- | --- |
| Python 3.8+ | 必需 | 跑全部脚本 | 系统自带 / conda |
| python-pptx | 必需 | 读写 PPTX、合并原理图页 | `pip install python-pptx` |
| Pillow、numpy、matplotlib | 必需 | 数据集图与演示可视化 | `pip install pillow numpy matplotlib` |
| lxml | 必需 | 解析 PPTX 主题配色 | `pip install lxml` |
| torch | 演示代码用 | 训练小模型（CPU 即可） | `pip install torch` |
| yt-dlp | 下载视频用 | 从 B 站拉 MP4 | `pip install -U yt-dlp` |
| ffmpeg | 转码用 | 统一 H.264、截片段 | `apt install ffmpeg` |
| LibreOffice | 可选 | PPTX → PDF 页数复核 | `apt install libreoffice` |
| slide-forge 技能 | 生成原理图时必须 | 图片版 + 元素版原理图页 | 见下 |

原理图页由 `slide-forge` 生成。它需要一份生图后端配置（OpenAI 兼容 API 或 OpenRouter），
密钥只放在环境变量或 `~/.slideforge/config.json`（chmod 600），**绝不要写进本仓库**：

```bash
export SLIDEFORGE_IMAGE_API_KEY="<your-key>"
export SLIDEFORGE_IMAGE_BASE_URL="https://<your-endpoint>/v1"
export SLIDEFORGE_IMAGE_MODEL="gpt-image-2"
```

## 三、输入：只认 prompt + 两个参数

| 输入 | 必需性 | 说明 |
| --- | --- | --- |
| `--title` | 必需 | 章节 / 课时标题，例如「第 7 章 循环神经网络」 |
| `--deck <path.pptx>` | 可选 | 原始课件；不给就跳过抽页 / 删页 / 裁剪 / 风格提取，其余照常 |
| `--minutes` | 可选，默认 90 | 课时；时间盒八段合计恰好等于它 |
| `--keep-ratio` | 可选，默认 0.70 | 课件保留率，脚本钳到 65%–75% |
| `--drop` | 可选 | 教师指定必删页码 |
| `--subagent-mode` | 可选，默认 `default` | `parallel` 启用双子代理 |
| `--slide-mode` | 可选，默认 `both` | 原理图产出：`image` 或 `image + element` |
| `--out` | 必需 | 运行目录 |

## 四、输入 prompt 范例（直接抄）

**范例 1 · 有课件、90 分钟、默认不用子代理**

```
帮我把这份课件做成 90 分钟的备课包：
/path/to/7. 循环神经网络 (1).pptx
要求：课上要有提问、讨论、代码演示、视频演示、PPT 讲解，最后留 5 分钟练习。
```

**范例 2 · 有课件、要并行加速、要原理图页**

```
用子代理模式跑这节课的备课包，主题「第7章 循环神经网络」，
课件在 /path/to/7. 循环神经网络 (1).pptx，90 分钟。
在原 PPT 第一节的原理部分加 3 页原理图，用流程图讲原理，风格跟原课件一致。
最后打成一个能下载的包，给出下载链接。
```

**范例 3 · 只有标题、没有课件**

```
给「第4章 卷积神经网络」出一节课的备课包，60 分钟，要 10 道单选题和两个能跑的演示代码。
```

**范例 4 · 直接告诉它用并行 + 元素版**

```
生成备课包：主题「第5章 迁移学习」，课件 /path/to/ch5.pptx。
启用子代理并行模式；原理图要图片版和元素版都要；视频找 2 段 5 分钟以内的。
```

命令行等价写法：

```bash
python3 scripts/lesson_prep.py all \
  --title "第7章 循环神经网络" \
  --deck "/path/to/7. 循环神经网络 (1).pptx" \
  --out ./runs/rnn-90 --minutes 90 --subagent-mode parallel --slide-mode both
```

## 五、子代理方案

**默认 `--subagent-mode default`：不使用子代理**，主代理顺序执行，适合小课时或想省上下文预算的场景。

**启用方式**：命令行 `--subagent-mode parallel`，或直接对 agent 说「用子代理模式 / 并行跑」。
启用后主代理会把两个子代理**同时**派出去，二者写入目标互不重叠：

| 子代理 | 负责 | 独占产物 |
| --- | --- | --- |
| **A · 幻灯与原理图** | `slide-forge` 生成原理图：读 `style.json` 复用原课件配色字体 → 图片版 → 元素版（对象级可编辑） | `slides/`（PNG、逐页提示词、图片版 PPTX/PDF、元素版 PPTX） |
| **B · 素材与文案** | 演示代码 + 数据集图、B 站视频检索与 MP4 下载、讨论题与参考解、10 道单选题与解析、课程方案正文 | `demos/`、`06-视频/`、`08-数据集图片/`、`01-课程方案.md`、`03-讨论题.md`、`04-课后习题.md` |

主代理只做：跑确定性步骤 → 派两个子代理 → 合并课件（`merge`）→ 打包（`bundle`）→ 验收（`verify_pack`）。
**有依赖的步骤仍串行**（裁剪先于原理图、合并晚于两者），所以并行只加速无依赖的部分，不牺牲质量。
细节见 [references/subagent-design.md](references/subagent-design.md)。

## 六、产出结构

```
runs/rnn-90/
├── deck.json                     课件结构化大纲
├── plan.json                     时间盒 + 删页清单 + 保留率
├── style.json                    原课件主题色与字体
├── 01-课程方案.md                 散文体课堂方案（无表格、无分条）
├── 02-精简课件.pptx               只删页、不改内容
├── 02-精简课件-2.pptx             合并原理图后的最终课件
├── 02-删页清单.md                 每条删除的理由
├── 03-讨论题.md / 04-课后习题.md
├── 06-视频/                       关键词、推荐清单、下载脚本、MP4
├── 07-代码或 demos/               两个演示脚本 + 公共工具 + 输出图
├── 08-数据集图片/                 数据集样例图
├── slides/                       原理图页（图片版 / 元素版 / deck_spec.json）
└── 子代理A-*.md、子代理B-*.md、主代理-*.md
```

打包后（`bundle` 产出的自包含目录）：

```
dist-<标题>/
├── <标题>-90分钟备课包.zip
├── index.html                    带样式的下载页
└── runs/…                       解压后的同一份文件（供下载页逐文件下载）
```

## 七、自检与验收

```bash
python3 scripts/lesson_prep.py config                        # 环境依赖
python3 scripts/verify_pack.py --pack-dir ./runs/rnn-90      # 交付物判据
python3 scripts/check_download_page.py --dist-dir ./dist     # 下载页链接逐条可开
```

判据清单见 [references/quality-gates.md](references/quality-gates.md)。

## 八、仓库里没有什么

- ❌ 任何课件、图片、视频、数据集等输入文件
- ❌ 任何运行产物（`runs/`、`dist-*/`、`*.zip`、`output/`）
- ❌ 任何 API Key、token、账号信息（`.gitignore` 已排除 `config.json`、`*.key`、`.env`）

## 九、许可与致谢

- 代码以 MIT 许可发布（见 [LICENSE](LICENSE)）。
- 原理图页由 `slide-forge` 技能生成；课件裁剪、时间盒与验收逻辑为本仓库自研。
- 课堂上使用的 B 站视频与公开数据集版权归各自原作者，仅供教学演示，请勿商用或二次上传。

## 十、常见问题

**Q：一定要有课件吗？**
A：不一定。只给标题也能产出课程方案、题目、代码、数据集图与打包；只是没有精简课件与原理图（原理图仍可生成，只是没有原课件风格可参考）。

**Q：时间盒会被我改坏吗？**
A：不会。脚本按目标分钟等比缩放，并在写入 `plan.json` 前断言八段合计 == `--minutes`；不满足就报错而不是静默出错。

**Q：为什么默认不用子代理？**
A：默认顺序执行更省上下文、更可预期；并行只对「无依赖且耗时」的环节有收益，所以做成按需开启。

**Q：原理图页会不会破坏原课件？**
A：默认 `--mode append` 追加到末尾，原页序完全不变；想插到某节之前用 `--mode insert --insert-after N`。
