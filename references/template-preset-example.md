# 模版预设示例（新模版第一次使用时填写一份）

每接一个新模版，先花 2 分钟给它写一份预设，存成 `references/<模版名>-preset.md`，之后整个系列都按这份预设来做。

## 怎么采集信息
```bash
python3 scripts/list_layouts.py Template.pptx            # 列出全部版式和占位符位置
python3 scripts/list_layouts.py Template.pptx --grep Section
unzip -p Template.pptx ppt/theme/theme1.xml | grep -o '<a:clrScheme.*</a:clrScheme>'   # 主题色
```
- 标题字体不用手动填，build_slide.py 的 `--font auto` 会自动检测。
- 先拿模版里的示例页渲染一张图看看，留意页脚的位置（通常在 y≈5.3 附近）和 logo 的位置，内容不要压到它们。

## 填写模板
- **画布**：16:9（10" × 5.625"）或 4:3
- **主题色**：dk1 / lt1 / dk2 / lt2 / accent1（主强调色）/ accent2（高亮色）/ accent3–6；accent 里哪个代表"成功 / 通过"（通常是绿色）
- **字体**：标题 ___，正文 ___，系列统一用 ___
- **常用版式**：内容页 ___、章节分隔页 ___、封面 ___
- **安全区**：内容不要超过 y ___（页脚从这里开始）
- **系列惯例**：页面语言、交付目录、文件名前缀、章节页格式（比如白卡插图 + "0N / 0M" 进度条）、过程页标签组

## 示例（虚构品牌 ACME）
- 主题色：accent1 #FA582D 橙红（主强调色），accent2 #FFCB06 黄（高亮用），accent4 #00CC66 绿（通过）
- 标题字体：Helvetica Neue
- 内容页：「Title Only - light」；章节页：「Section Divider」，渐变背景，标题为白色 36pt 居中
- 页脚从 y 5.33 开始
- 文件名用 `ACME_<主题>.pptx`，章节页用白卡格式，加 "0N / 05" 进度条
