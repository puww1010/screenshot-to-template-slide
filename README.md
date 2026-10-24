# screenshot-to-template-slide

把 **幻灯片截图（原图）+ 讲稿 + 可选参考 slide**，在你指定的 **PPTX 模版**上重建成一页**可编辑、图文并茂**的幻灯片。生成的页面全部用原生形状（不是图片），配色用模版的主题色，模版里的任何文件都不改。

这是一个 [Claude](https://claude.ai) Agent Skill。规则和工具来自一次真实的多轮协作：用二十多页 explainer 迭代出来，用户纠正过的每一处都写进了硬规则。

## 特点
- **规则先行**：原图文字逐字保留，讲稿不上页面（可以写进演讲者备注），版式结构照原图，模版保持原样，每页只出一个方向，同级元素样式统一。
- **slidekit**：原生 OOXML 形状库，包括 26 种图标、20 多个插画部件，以及标签条、进度条、阴影卡片、荧光高亮等页面辅助函数。
- **一条命令完成构建和 QA**：注入形状 → 统一字体（自动检测模版标题字体）→ 用 validate 校验 → 和模版逐文件比对 → 字号放大 6% 做压力渲染 → 导出页面文字供核对。
- **整套 deck 合并**：按版式名称对应，合并多个单页文件。
- **中文模式**：`--lang zh`，中文字体单独设置，术语保留英文。

## 安装
- **Claude 桌面版 / Cowork**：把本仓库目录打包成 `screenshot-to-template-slide.skill`（zip，SKILL.md 放在 `screenshot-to-template-slide/` 目录下），在对话里打开后点 **Save skill**。
- **Claude Code**：把仓库复制到 `~/.claude/skills/screenshot-to-template-slide/`。

依赖：Python 3、`zip`、LibreOffice 与 `pdftoppm`（用于 QA 渲染）、`markitdown`；校验和渲染脚本复用官方 `pptx` skill（`scripts/office/validate.py`、`soffice.py`）。

## 快速开始
```bash
python3 scripts/list_layouts.py MyTemplate.pptx --grep Section
python3 scripts/build_slide.py --template MyTemplate.pptx --page examples/how_it_works_tabs.py \
        --out out.pptx --qa-dir qa/ [--layout "Section Divider"] [--notes script.txt] [--lang zh]
python3 scripts/merge_deck.py --out deck.pptx cover.pptx agenda.pptx ch1.pptx ...
```
`page.py` 只需定义 `build()`，返回 slidekit 形状列表，写法参考 `examples/`。

## 目录
```
SKILL.md                      规则与工作流（给 Claude 读）
scripts/slidekit.py           原生形状与图标库
scripts/build_slide.py        单页构建与 QA
scripts/merge_deck.py         合并整套 deck
scripts/list_layouts.py       列出模版版式
examples/                     章节分隔页 / How-it-works 标签页 / 插画流程页
references/                   页型结构、QA 清单、常用说法对照、模版预设示例
```
