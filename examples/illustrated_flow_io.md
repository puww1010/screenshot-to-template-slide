# 输入 → 输出对照：插画型流程页（illustrated_flow.py）

> **说明**：这是根据 `illustrated_flow.py` 反推出来的文字版样例，没有附原图截图。原图文字清单是按成品页面反推的；**讲稿是为了演示规则写的示意文本，不是真实讲稿**。要演示的是三件事：哪些字能上页面，哪些不能，最后怎么核对。

## 输入

**原图**：以插画为主，没有标题。4 张卡片横排，卡片之间用箭头连接，每张卡片的上方和下方各有插图。

**第 1 步：抄文字清单**（按原图从左到右、从上到下）

| # | 位置 | 原图文字 |
|---|---|---|
| 1 | 卡片 1 标题 | Plan |
| 2 | 卡片 1 说明 | Pick scope & strategy |
| 3 | 卡片 2 标题 | Build |
| 4 | 卡片 2 说明 | Parallel work & candidates |
| 5 | 卡片 3 标题 | Verify |
| 6 | 卡片 3 说明 | Tested in a sandbox |
| 7 | 卡片 4 标题 | Release |
| 8 | 卡片 4 说明 | Score, approve, publish |
| 9 | 卡片 3 上方插图里的小标签 | sand-box |

原图里的插图元素（重画时要一一对应）：

| 卡片 | 上方 | 下方 |
|---|---|---|
| 卡片 1 | 放大镜、指南针 | 服务器 |
| 卡片 2 | 二进制数据流 | 齿轮 |
| 卡片 3 | 沙盒 | 半边盾牌 |
| 卡片 4 | 大门 | 对勾、火箭 |

**讲稿（示意）**

> Let's look inside the delivery pipeline. We start by planning: picking the scope and the strategy for this round. The build stage then runs several streams of work in parallel and collects candidates. Every candidate is tested in a sandbox before anyone sees it, so nothing ships on trust alone. Finally the release step scores each result, gets approval and publishes it — and the whole loop runs every day without anyone pushing it along.

## 判断

- 原图**没有标题**，属于 §2.1 说的"原图几乎没有文字"。这时从讲稿里取最少量的文字补一个短标题，措辞贴近讲稿原话：**Inside the delivery pipeline**。交付时要说明这个标题是补的、可以删。
- 讲稿里其他句子**一律不上页面**，比如 "so nothing ships on trust alone"、"runs every day without anyone pushing it along"。这些句子只用来理解逻辑，比如帮你决定第 3 张卡配沙盒、第 4 张卡配大门和火箭。需要的话，可以用 `--notes` 把讲稿写进演讲者备注。
- 结构照原图来（§2.2）：4 张卡横排，箭头连接，插图放在卡片上下。不要改成竖排，也不要另加编号徽章。

## 输出

**page.py**：见 `illustrated_flow.py`。构建命令：

```bash
python3 scripts/build_slide.py --template examples/demo-template.pptx --page examples/illustrated_flow.py \
    --out OUT.pptx --qa-dir <输出目录>/QA [--notes script.txt]
```

**第 6 步：核对文字**（把 `[text on slide]` 的输出和文字清单逐条对照）

| 页面文字 | 来源 | 结论 |
|---|---|---|
| Inside the delivery pipeline | 讲稿（补的标题） | ✅ 允许，交付时说明 |
| Plan / Pick scope & strategy | 清单 #1–2 | ✅ |
| Build / Parallel work & candidates | 清单 #3–4 | ✅ |
| Verify / Tested in a sandbox | 清单 #5–6 | ✅ |
| Release / Score, approve, publish | 清单 #7–8 | ✅ |
| sand-box | 清单 #9 | ✅ |
| 0110101011 等 4 行 | `binary()` 插图部件自带的装饰字符 | ✅ 属于插图，不算正文 |

清单里的 9 条全都出现了，也没有多出讲稿里的句子，核对通过。

**交付汇报里要提到**：
- 标题是从讲稿补的，可以删。
- 4 张卡片和插图元素都按原图的对应关系重画了。
- 校验通过，模版只改了 `slide1.xml`。
