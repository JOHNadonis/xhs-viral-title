# xhs-viral-title

小红书爆款标题生成器 —— WorkBuddy / Claude Skill。

用 **116 个经过验证的高赞标题句式结构**，给任意产品批量生成标题文案。输出可直接发布的 Excel 表，每条默认控制在 20 字内。

## 核心思路

**结构是固定的，只有槽位里的产品信息在变。**

这套结构库来自一份爆款标题排行榜（116 个句式，每条都带 3 个真实高赞示例、相似标题数、点赞中位数）。做标题不是每次从零想创意，而是**把产品信息套进已经跑通的骨架里**。

每个结构都标注了三件事：
- **槽位**（`slots`）：该填什么，比如「对象」「痛点场景」「价格」
- **品类倾向**（`fit`）：`F` 服饰 / `U` 美妆个护 / `H` 家居百货 / `G` 食品饮品 / `D` 数码 / `E` 母婴 / `K` 知识课程 / `*` 通用
- **语气**（`tone`）：`price` 价格冲击 / `hype` 夸张赞叹 / `story` 故事自述 / `pain` 痛点共情 / `soft` 温柔治愈 / `social` 社交背书

## 安装

### 方式一：手动复制

```bash
# Windows
xcopy /E /I xhs-viral-title "%USERPROFILE%\.workbuddy\skills\xhs-viral-title"

# macOS / Linux
cp -r xhs-viral-title ~/.workbuddy/skills/
```

### 方式二：git clone

```bash
git clone https://github.com/JOHNadonis/xhs-viral-title.git ~/.workbuddy/skills/xhs-viral-title
```

### 依赖

Python 标准库 + `openpyxl`：

```bash
pip install openpyxl
```

## 快速开始

```bash
cd ~/.workbuddy/skills/xhs-viral-title

# 1. 看全部 116 个结构
python scripts/xvt.py list

# 2. 看单个结构详情（槽位/示例/语气/品类）
python scripts/xvt.py show --rank 5

# 3. 导出待填模板（默认全量 116 个，不筛选）
python scripts/xvt.py pick --per-row 3 --out template.json

# 4. 填完文案后校验（字数/重复/空值）
python scripts/xvt.py check --titles filled.json --strict

# 5. 装配 Excel（主表 + 产品信息卡）
python scripts/xvt.py assemble --filled filled.json --out 产品名_标题文案.xlsx
```

## 标准流程

```
提取产品卡  →  导出结构模板  →  填文案（每结构 3 条）  →  校验  →  装配 Excel
```

### 第 1 步：提取产品卡

从图片 / 文档 / 口述里抽结构化信息，存成 JSON：

```json
{
  "产品名": "AYT 云感安睡护腰枕",
  "到手价": "39.9（原价 79.9）",
  "规格": "蓝白色；大号/小号可拆洗",
  "核心材质": "羽丝棉 / 透气网眼 / 静音魔术贴",
  "卖点": ["填满腰窝", "软而不塌不硌骨", "静音魔术贴不滑位"],
  "场景": ["平躺睡", "侧卧睡", "久坐办公", "出差旅行"],
  "痛点": ["起床腰酸", "侧睡腰悬空", "硬靠垫硌骨"],
  "买家原话": ["侧卧时候最舒服，能给到支撑", "魔术贴静音粘的很紧"]
}
```

**「买家原话」是最值钱的一栏。** 从评价区截图里抄真实的口语，比自己造卖点效果好一个档。

### 第 2 步：导出结构模板

```bash
python scripts/xvt.py pick --per-row 3 --out template.json
```

默认出全部 116 个。只在需要时才加筛选：

| 参数 | 作用 |
|---|---|
| `--category H` | 只出家居百货适用结构（`*` 通用结构始终包含） |
| `--tone price,hype` | 只要价格冲击 + 夸张赞叹 |
| `--min-likes 600` | 只要点赞中位数 600 以上 |
| `--rank 4,5,69` | 指定某几条 |
| `--limit 20` | 只取前 N 条 |

### 第 3 步：填文案

读 template.json，对每个 structure 的 slots 填产品信息，产出：

```json
{"rows": [
  {"rank": 1, "titles": ["标题一", "标题二", "标题三"]},
  {"rank": 4, "titles": ["...", "...", "..."]}
]}
```

### 第 4 步：校验

```bash
python scripts/xvt.py check --titles filled.json --strict
```

只查三件事：**超 20 字、重复、空值**。

这是格式校验，不扫词、不做内容审查。能不能发由你自己判断。

### 第 5 步：装配 Excel

```bash
python scripts/xvt.py assemble --filled filled.json \
  --out 产品名_标题文案_N条.xlsx \
  --product-card card.json
```

生成两个 sheet：主表（序号 / 排名 / 结构 / 相似数 / 点赞中位数 / 文案1-N）+ 产品信息卡。

## 目录结构

```
xhs-viral-title/
├── SKILL.md                          # 主文档（技能入口）
├── assets/
│   ├── structures.json               # 116 个结构主数据
│   └── 爆款标题排行榜1026终版.xlsx     # 原始排行榜
├── references/
│   ├── writing-rules.md              # 怎么写才像人话
│   └── risk-guide.md                 # 发布前自查参考（按需查阅）
└── scripts/
    ├── xvt.py                        # 主工具（5 个子命令）
    └── build_structures.py           # 从原始 xlsx 重建 structures.json
```

## 三条写文案的硬规则

### 1. 结构不重写，只换槽位

句式骨架、重复节奏、括号补充、省略号叹气、emoji 位置全部保留。看到「39.9！39.9！39.9！说三遍！」就照写三遍——**啰嗦就是它的爆点**。

### 2. 同结构出 3 条，必须切不同角度

不要三条都是「39.9 好便宜」。三种切法：

- **按卖点切**：价格 → 材质 → 场景
- **按人群切**：侧睡党 → 久坐族 → 出差党
- **按情绪切**：强冲击 → 中对比 → 弱陈述

### 3. 20 字是硬上限

小红书标题超 20 字会被截断。写完必须跑 `check`。

超字了先砍什么？**别先砍语气词。**「啊」「诶」「。。。」「❗️」是这类标题的呼吸感，要砍就砍修饰、砍重复的形容词。

## 语感校准（`references/writing-rules.md`）

爆款标题的「爆」不在词藻，在**说话的方式**。同一句话换个说法就不像小红书了。

拆开 116 个结构，真正让人想点开的动作只有五种：

| 动作 | 机制 | 例子 |
|---|---|---|
| **反差** | 预期被打碎 | 「以为要一百多。。。结果39.9」 |
| **具体** | 抽象变可量 | 「穿一次被夸三天」而不是「很显气质」 |
| **共情** | 说出对方没说出口的感受 | 「侧睡的时候腰是悬空的」 |
| **稀缺** | 制造错失感 | 「唯一的缺点：只有小号还在」 |
| **身份** | 让人找到自己 | 「请大数据推给所有侧睡党」 |

填槽位前先问：这条在用哪个动作？五个都没有，就是白写。

## 更新结构库

榜单更新了？把新 xlsx 覆盖到 `assets/`，然后：

```bash
python scripts/build_structures.py
```

会重建 `structures.json`。新增结构的槽位标注在脚本里的 `META` 字典，需要补一下（每个结构要标：槽位名列表、品类码、语气）。

## 实测数据

拿护腰枕跑过一轮：116 结构 × 3 = **348 条**，全部 ≤20 字、0 重复、0 空值。

## License

MIT
