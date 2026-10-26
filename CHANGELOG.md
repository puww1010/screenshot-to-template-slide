# Changelog

## v1.2.1 — 2026-10-26
按"7 步构建 Skill"的标准做了一次体检，修了一个静默失败的 bug，并补齐了测试和样例。
- **修复**：`slidekit.inject()` 原来只认 `</p:grpSpPr>`。模版页写成自闭合的 `<p:grpSpPr/>` 时（python-pptx 生成的模版就是这样），形状一个都不会注入，校验却照样显示 PASSED。现在两种写法都支持，都匹配不到时直接报错。
- **断言**：`build_slide.py` 发现幻灯片部件和模版完全一样时，打印 `[FAIL]` 并以非零码退出，不再默默交付一张空页。
- **description**：写明前提（必须同时有原图截图和用户模版）和边界（只想把截图转成可编辑 PPT 的用 image-to-editable-ppt；从零写 deck 的不用本 skill），避免和其他 PPT skill 抢触发。
- **SKILL.md**：头信息新增 `metadata.version: "1.2.1"`；开头加了 5 条红线；每个 reference 都写明了什么时候读；`--qa-dir` 改为要求绝对路径，并且放在看图工具能读到的目录里；补充了 `[FAIL]` 和渲染被跳过时该怎么处理。
- **新增** `examples/illustrated_flow_io.md`：一组"输入 → 输出"对照（文字清单、示意讲稿、判断、页面文字核对），演示"讲稿不上页面"。
- **新增** `references/testing.md`：冒烟测试（demo 模版和自闭合空白模版各跑一遍 4 个 example）和触发测试。
- `references/prompt-patterns.md`：和 SKILL.md 重复的处理方式改成"见 §x"，只保留说法到章节的映射和几条补充做法。

## v1.1.1 — 2026-10-24
- README 新增"效果示例"一节，附 4 张示例图（docs/images/）。
- 新增示例 `examples/stat_timeline.py`：大数字环形图、三步链路、THEN/NOW 时间线和结论横幅。
- 新增虚构品牌 ACME 的演示模版 `examples/demo-template.pptx` 及其生成脚本 `make_demo_template.js`，没有自己的模版也能试跑。
- 示例 `how_it_works_tabs.py` 补全了 QA gate 的第二个测试项和 Operational impact 说明。

## v1.1.0 — 2026-10-24
提升在其他 AI Agent 中的可用性。
- `build_slide.py` 不再强制依赖官方 `pptx` skill：
  - 内置删除多余幻灯片（`prune_unlisted_slides`），模版里有多页时也能干净地只留一页；
  - 内置基础包结构检查（`basic_validate`）：XML 是否合法、rels 目标是否存在、每个部件是否有 content type、slide 内 shape id 是否重复；
  - 内置 LibreOffice 直接渲染，使用独立 profile，超时 240 秒；
  - 没有 `markitdown` 时自己提取页面文字；
  - 新增 `--pptx-skill none`，可强制走内置流程。
- `SKILL.md` 和参考文档里的工具名改成通用说法，并注明在 Claude、Cowork 里对应的工具。
- README 新增"在其他 AI Agent 中使用"和英文简介。

## v1.0.0 — 2026-10-24
首个公开版本：规则与工作流、slidekit、build_slide、merge_deck、list_layouts、3 个示例页、4 份参考文档。MIT License。
