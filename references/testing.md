# 测试：改了 skill 之后先跑这几项

## 1. 冒烟测试（脚本）

修改 `scripts/` 或 `examples/` 之后跑。会用两个模版各跑一遍 4 个 example，因为两个模版里 `<p:grpSpPr>` 的写法不同，两种情况都要覆盖：

- `examples/demo-template.pptx`：写法是 `<p:grpSpPr>…</p:grpSpPr>`。
- python-pptx 现场生成的 16:9 空白模版：写法是自闭合的 `<p:grpSpPr/>`。v1.2.1 之前，这种写法下形状会静默注入失败。生成它需要 `pip install python-pptx`，只在测试时用到。

```bash
SK=<本 skill 目录>; OUT=<输出目录>/smoke; mkdir -p "$OUT"
python3 -c "
from pptx import Presentation; from pptx.util import Inches
p=Presentation(); p.slide_width=Inches(10); p.slide_height=Inches(5.625)
s=p.slides.add_slide(p.slide_layouts[5]); s.shapes.title.text='x'; p.save('$OUT/blank.pptx')"
for T in "$SK/examples/demo-template.pptx" "$OUT/blank.pptx"; do
  t=$(basename "$T" .pptx)
  for e in stat_timeline how_it_works_tabs illustrated_flow section_divider; do
    L=(); [ "$t" = demo-template ] && [ $e = section_divider ] && L=(--layout "ACME Section Divider")
    python3 "$SK/scripts/build_slide.py" --template "$T" --page "$SK/examples/$e.py" \
        --out "$OUT/${t}_$e.pptx" --qa-dir "$OUT/qa" "${L[@]}"; echo "$t / $e exit=$?"
  done
done
```

在空白模版上，section_divider 没有对应的版式。它的标题、副标题和进度条是白字，在白底上看不见，这是正常现象，只要插图卡片画出来就算通过。

**通过标准**（8 次构建都要满足）：
- `exit=0`，输出里没有 `[FAIL]`。
- `[validate]`（或 `[validate:basic]`）显示 PASSED。
- `[diff vs template]` 里有 `ppt/slides/slide1.xml`。
- `[text on slide]` 打印的是 example 里的文字，不是模版原来的文字。
- 打开 `$OUT/qa/*_stress-1.jpg`，确认形状都画出来了。空白模版的标题是 44pt，长标题折行属于正常。

只想测内置兜底流程时，在命令后面加 `--pptx-skill none`。

## 2. 回归测试（脚本）

```bash
python3 tests/test_regressions.py
```
覆盖已经踩过的坑：行距单位（lnpct=120 → 120000）、lint 拦截 PowerPoint 会照原值排版的异常值、`--lang zh` 给中文文字标 zh-CN、非 10 英寸画布的提示、非系统字体的提示。需要 python-pptx，用来临时生成模版。最后一行显示 `all passed` 才算通过。修 bug 时先在这里加一个能复现问题的测试，再改代码。

## 3. 触发测试（description）

修改 description 之后，在一个装了本 skill 的新对话里逐条发下面几句话，看触发的是哪个 skill：

| 用户说法 | 应该触发 |
|---|---|
| "以 MyTemplate.pptx 为模版，把这张截图重做一页，讲稿如下……" | 本 skill |
| "参考这个 slide 的风格，按原图在我的模版上重建，不要改动模版的 logo 和颜色" | 本 skill |
| "把这张截图转成可编辑的 PPT"（没有给模版） | image-to-editable-ppt（环境里装了的话） |
| "这个扫描版 PDF 帮我变成可编辑的 pptx" | image-to-editable-ppt（环境里装了的话） |
| "帮我做一份关于零信任的 10 页 PPT"（没有截图） | 通用的 PPT skill，不是本 skill |

如果前两条没有触发本 skill，或者后三条触发了本 skill，就回去改 description 里的触发词和边界说明。
