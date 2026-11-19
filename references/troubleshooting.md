# 排错手册：脚本输出 → 原因 → 怎么办

`build_slide.py` 和 `doctor.py` 的每一行提示都以 `[FAIL]`、`[WARN]`、`[NOTE]` 开头，或者以 `[qa render]`、`[notes]` 这类步骤名开头。先在下面找到对应的那一行，按"怎么办"处理。处理完重新构建，直到没有 `[FAIL]`，`[WARN]` 也都确认过。

- **`[FAIL]`**：构建停止，不会交付文件。
- **`[WARN]`**：文件已经生成，但很可能有问题，交付前必须处理或向用户说明。
- **`[NOTE]`**：提醒用户注意，不影响构建。

## 环境和安装（先跑 `python3 scripts/doctor.py --selftest`）

| 提示 | 原因 | 怎么办 |
|---|---|---|
| `[FAIL] package file missing: …` | skill 包不完整，常见于只复制了 SKILL.md | 从 GitHub Releases 下载完整的 `.skill` 包重新安装，不要只拷 SKILL.md |
| `[FAIL] Python … is too old` | Python 版本低于 3.8 | 安装 Python 3.8 或更新版本 |
| `LibreOffice  missing` | 没装 LibreOffice，压力渲染会被跳过 | 按 doctor 打印的命令安装；装不了时走下面的"降级路径" |
| `pdftoppm  missing` | 没装 poppler，渲染结果只有 PDF，没有 JPG | 按 doctor 打印的命令安装，或者直接打开 PDF 检查 |
| `markitdown` / `Anthropic pptx skill`  missing | 可选组件 | 不用处理，脚本会改用内置流程 |
| `[FAIL] selftest build failed` | 环境里有更深层的问题 | 看后面打印的 8 行日志，再到本页对照具体的 `[FAIL]` |

**降级路径**：没有 LibreOffice，或者环境不能执行命令时，按下面的顺序退一步，并在交付时告诉用户少做了哪一步。
1. 有 Python、没有 LibreOffice：照常构建。lint、模版比对、文字核对都会执行，只少了渲染图。交付时请用户在 PowerPoint 里逐页检查版面。
2. 完全不能执行命令：这个 skill 的自动构建跑不起来。告诉用户需要一个能执行 shell 和 Python 的环境，比如 Claude 桌面版、Cowork 或 Claude Code。不要只凭 SKILL.md 手写 XML 交付。

## 构建时的 `[FAIL]`

| 提示 | 原因 | 怎么办 |
|---|---|---|
| `--template file not found` / `--page file not found` | 路径写错，或者用了相对路径 | 改用绝对路径 |
| `--template is not a .pptx file` | 传进来的是 .key、.pdf、.potx，或者文件已损坏 | 先在 PowerPoint 里另存为 .pptx |
| `--template has no slides` | 模版里一页幻灯片都没有 | 在 PowerPoint 里加一页空白页再保存 |
| `layout "…" not found … Available layouts:` | `--layout` 的名字写错了 | 从打印出来的列表里复制准确的名字，或者跑 `list_layouts.py` |
| `--page … could not be loaded` | page.py 有语法错误或 import 错误 | 按提示修正；开头那行 `sys.path.insert` 照抄 `examples/page_template.py` |
| `must define build()` | page.py 里没有 `build()` 函数 | 从 `examples/page_template.py` 复制骨架 |
| `build() … raised …` | `build()` 运行时出错，提示里有行号 | 修改对应的那一行 |
| `must return a list of strings` | `build()` 返回的不是形状列表 | `return S`，S 是由 slidekit 函数返回值组成的列表 |
| `inject: no <p:grpSpPr>…</p:spTree>` | 模版页的 XML 结构不规范 | 在 PowerPoint 里打开模版重新保存 |
| `… is identical to the template` | 形状没有注入进去 | 检查 `build()` 是否返回了空列表 |
| `lint: line spacing spcPct val="120"` | 行距单位写错：120 被当成 0.12% | 用 `para(..., lnpct=120)`，不要手写 `spcPct` |
| `lint: font size sz="…"` | 字号超出 1–400 pt | 检查是不是多乘了 100，或者 K 用错了 |
| `lint: spcBef/spcAft …` | 段前、段后间距超出 0–200 pt | 检查 `para(..., before=…)` 的值，单位是 pt |

## `[WARN]`（文件已生成，交付前要处理）

| 提示 | 原因 | 怎么办 |
|---|---|---|
| `[WARN] canvas 20 x 11.25 in … x2.0` | 模版画布不是 10 英寸 | 在 page.py 里设 `K = 2`，所有坐标和字号乘以 K（见 `examples/page_template.py`）；`progress_meter` 传 `scale=K` |
| `[WARN] … contain CJK text but are not tagged zh-CN` | 页面有中文，但没加 `--lang zh` | 重新构建时加 `--lang zh` |
| `[WARN] section divider …` | 章节页上画了插图 | 章节页只放标题、副标题和进度条（SKILL.md §2.4）；如果是用户明确要求的插图，交付时说明 |
| `[validate:basic] ISSUES: …` | 包结构有问题 | 按列出的问题修正；通常是 page.py 手写的 XML 有错 |

## `[NOTE]` 和步骤提示

| 提示 | 含义 | 怎么办 |
|---|---|---|
| `[NOTE] font "…" is not a standard Windows/macOS font` | 字体不是两个系统都自带的 | 交付时提醒用户在放映电脑上安装字体，或在 PowerPoint 里嵌入字体 |
| `[NOTE] font "PingFang SC" ships with macOS only` | 苹方只有 Mac 才有 | Windows 上放映时用 `--ea-font "Microsoft YaHei"` |
| `[qa render] LibreOffice not found — skipped` / `timed out` | 渲染被跳过 | 见上面的"降级路径" |
| `[qa render] PDF at …` | 没装 pdftoppm | 直接打开 PDF 检查 |
| `[notes] slide has no notesSlide part — skipped` | 模版页没有备注页 | 讲稿写不进备注；在 PowerPoint 里给这页加一次备注后保存，再重新构建 |

## 渲染图看着没问题，但 PowerPoint 里不对

LibreOffice 和 PowerPoint 有已知差异，见 `references/lessons.md` 的"LibreOffice 和 PowerPoint 的差异"。遇到新的差异时：
1. 在 `tests/test_regressions.py` 里加一个能复现问题的测试；
2. 修改脚本，或者在 lint 里加一条规则；
3. 在 lessons.md 里补一条记录。
