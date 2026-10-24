# Changelog

## Unreleased
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
