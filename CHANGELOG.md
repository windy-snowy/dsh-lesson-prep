# 变更记录

遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/) 与语义化版本。

## [0.1.0] - 2026-09-20

### 新增
- `SKILL.md`：DSH 技能本体，含适用场景、依赖清单、输入参数、子代理方案、原理图与 slide-forge 的分工、验收判据、失败降级与红线。
- `scripts/lesson_prep.py`：确定性内核，12 个子命令
  `extract / plan / trim / media / video / quiz / outline / merge / bundle / prompts / style / config / all`。
  - `plan`：八段时间盒（合计断言等于 `--minutes`）+ 删页清单；硬保护封面、资源页、分隔页、小结页与每小节前两页；候选按「重复页 → 整页代码 → 运行结果截图 → 内容偏薄页」打分；保留率钳制在 65%–75%。
  - `trim`：只改 `presentation.xml` 的 `sldIdLst` 与关系文件，并连带清理被删页的部件与独占媒体，绝不触碰 slide 内容。
  - `style`：从原课件主题提取配色与字体，供 `slide-forge` 做风格参考。
  - `merge`：把原理图页合进课件，`append`（默认，最稳）或 `insert --insert-after N`。
  - `bundle`：自包含目录 + zip + 带样式的下载页（相对链接全部可用，UTF-8 文件名不乱码）。
  - `prompts`：生成两个子代理与主代理的角色提示词。
- `scripts/verify_pack.py`：把交付判据变成可执行的 fail/warn 检查。
- `prompts/`：`subagent-slides.md`（slide-forge：图片版 + 元素版 + 原课件风格）、`subagent-parallel.md`（代码/视频/题目/方案）、`main-agent.md`（总装与验收）。
- `references/`：`pipeline.md`（八阶段责任与降级）、`subagent-design.md`（默认单代理、按需双代理、零重叠写入）、`quality-gates.md`（判据与下载页规格）。
- `examples/`：8 个输入 prompt 范例 + 命令行等价写法 + `quickstart.sh` 最小试跑脚本。

### 设计取舍
- 默认**不召唤子代理**；只有 `--subagent-mode parallel` 或用户明确要求时才并行，且并行只覆盖无依赖且耗时的环节，质量闸门保持串行。
- 仓库不含任何课件、图片、视频、运行产物与 API Key；`.gitignore` 排除 `runs/`、`dist-*/`、`*.pptx`、`*.mp4`、`config.json`、`.env` 等。
