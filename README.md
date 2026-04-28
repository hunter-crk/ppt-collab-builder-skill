# 数据咨询类型生成PPT生成 Skill

一个专门用于生成**中度或重度依赖数据分析的PPT**的 Claude Code skill（v2.2）。

这个skill不是通用的“做一份好看的PPT”工具，而是面向经营分析、业务复盘、品类分析、销售汇报、ROI评估、目标拆解、选品分析等数据驱动场景。它的核心目标是：在保证数据口径、图表定义、洞察结论可追溯的前提下，生成可继续精确调整的 `.pptx` 文件。

## 为什么需要这个skill

Codex、通用Agent、在线PPT工具或图片生成工具可以快速生成视觉上很好看的页面，但在中重度数据分析型PPT里，经常会遇到这些问题：

| 常见问题 | 对数据型PPT的影响 |
|----------|------------------|
| 数据幻觉 | 图表、数字、同比环比、排序和结论可能与真实Excel不一致 |
| 图表不可追溯 | 很难确认图表来自哪份数据、哪个字段、哪种计算口径 |
| 调整精度低 | 图片式PPT好看但难以逐个修改表格、文字、图表位置和数据 |
| 业务结构不稳定 | 缺少需求确认、板块规划、图表方案确认，容易生成“像PPT但不像汇报”的内容 |
| 风格与数据割裂 | 页面样式漂亮，但图表标签溢出、表格不可读、洞察位置不服务分析逻辑 |

这个skill通过**6阶段协作流程 + Excel数据模板 + 图表方案确认 + 数据洞察 + JS/ECharts/pptxgenjs生成**来降低这些问题。

## 适用场景

适合：
- 经营分析、销售分析、品类分析、渠道分析、用户分析
- 目标达成、增长机会、ROI、预算投入产出分析
- SKU/商品/门店/区域等多维指标汇报
- 需要用户填Excel、复核数据口径、确认图表方案的PPT
- 需要最终产出可编辑、可复用、可二次精修的 `.pptx`

不适合：
- 纯品牌视觉稿、概念海报式PPT、极少数据的创意提案
- 只需要一张漂亮图片或一次性视觉草图的场景
- 不需要数据核对、不关心图表口径的轻量PPT

## 快速开始

当用户说“帮我做一个经营分析PPT”“把这个Excel做成PPT”“生成数据汇报/销售分析/品类分析PPT”等类似需求时，这个skill就会激活。

## 核心亮点

| 亮点 | 说明 |
|------|------|
| 数据先行 | 先确认业务目标、受众、板块和图表方案，再生成数据模板，避免凭空编数字 |
| 图表口径可确认 | 每个图表都明确类型、维度、系列、数据来源和编号 |
| 纯表格/ECharts分流 | 精确数值用pptxgenjs表格，复杂趋势/占比/分布图用浏览器+ECharts生成高清图 |
| 洞察可追溯 | 数据填好后先做数据分析和洞察更新，再进入最终PPT生成 |
| 可编辑输出 | 最终输出为`.pptx`，表格、文本、图片和页面元素可继续调整 |
| 自适应排版 | 根据项目风格画像选择版式变体，避免所有页面套同一种模板 |
| 适合Agent接入 | 把信息收集、图表规划、资源准备、最终生成拆成明确阶段，便于Agent稳定执行 |

## Skill 结构

```
ppt-collab-builder/
├── SKILL.md              # 主skill定义
├── README.md             # 这个文件
├── references/           # 参考文档
│   ├── workflow.md       # 6阶段工作流详细说明
│   ├── templates.md      # 页面模板与布局规范
│   └── adaptive-layout.md # 项目风格画像与自适应排版规则
├── scripts/              # 辅助脚本
│   └── ppt_builder.js    # PPT构建基础参考工具（v2.1）
└── examples/             # 示例
```

## 重要说明：scripts/ppt_builder.js 的定位

**`scripts/ppt_builder.js` 是一个基础参考工具，不是万能的通用脚本。**

### 为什么每个项目需要单独的生成脚本？

1. **页面结构差异大**：不同项目的页面类型、数量、布局组合都不同
2. **视觉规范定制化**：配色、字体、尺寸等每个项目都可能有特殊要求
3. **数据格式不统一**：Excel模板结构、数据来源每个项目都不一样
4. **表格图表组合独特**：每个项目的表格和图表的组合方式都是独一无二的

### ppt_builder.js 提供的价值

| 组件 | 用途 |
|------|------|
| 辅助函数 | `cell()` / `headerCell()` / `highlightCell()` / `createBusinessTable()` |
| 最佳实践 | pptxgenjs 表格API正确用法、ECharts配置生成逻辑 |
| 基础方法 | `createProject()` / `generateFullOutlineMarkdown()` / `pregenerateECharts()` |
| 配色规范 | `TABLE_COLORS` 常量定义 |

### 使用方式

1. **参考** `ppt_builder.js` 中的代码结构和辅助函数
2. **根据项目需求** 编写专门的 `{project_name}_builder.js`
3. **复用** 辅助函数和最佳实践，但整体流程需要定制

## 工作原理（v2.2 更新）

这个skill通过多轮对话指导用户完成6个阶段，重点解决“数据从哪里来、图表怎么定义、结论如何生成、最终PPT如何可编辑”：

1. **需求分析** - 了解项目目的、受众、**必需板块**和数据分析目标，AI生成2-3个PPT结构方案供选择
2. **大纲完善** - 选择主题（可上传参考截图），AI生成项目风格画像并建议每页版式，最终大纲包含【风格画像、视觉规范、结论呈现规范、每页目的/版式变体/布局/图表编号】
3. **图表规划** - AI先判断：纯表格还是ECharts？然后给出建议方案，再确认
4. **资源准备** - 为确认后的图表/表格生成Excel数据模板，指导用户填数和放置图片
5. **数据洞察** - 读取已填数据，分析趋势、峰值、结构、增长率等，生成结论并更新大纲
6. **最终生成** - 按顺序执行：①数据洞察更新大纲→②检查视觉风格与版式适配→③仅预生成ECharts复杂图表（纯表格跳过）→④用JS+ECharts+pptxgenjs按项目风格合成最终PPT

## 技术栈要求（严格执行）

**必须使用以下技术栈，不得变更**：

| 层级 | 技术 | 用途 |
|------|------|------|
| 运行时 | Node.js (JavaScript) | 脚本运行环境 |
| 图表渲染 | 浏览器(Puppeteer/Playwright) + ECharts库 | 复杂图表渲染，**不使用pptxgenjs内置图表**，用echarts+浏览器生成2倍分辨率高清PNG |
| PPT生成 | pptxgenjs | 直接生成.pptx文件，纯表格(直接用API)、文本、插入ECharts图片 |
| 数据处理 | 原生JS (fs/path) | 文件读写、数据处理 |

**禁止使用**：Python或其他技术栈方案。

**ECharts图表生成方式说明**：
- ECharts复杂图表**不使用**pptxgenjs内置的图表功能
- 使用**浏览器自动化工具(Puppeteer/Playwright) + ECharts库**在浏览器中渲染
- 生成2倍分辨率(@2x)高清PNG图片
- 然后用pptxgenjs将生成的图片插入PPT

## 环境准备（重要！）

在阶段6预生成ECharts图表之前，必须确保Playwright和浏览器内核已正确安装。

### Playwright 安装

```bash
# 安装Playwright依赖
npm install playwright

# 或者使用yarn
yarn add playwright
```

### 浏览器内核配置（Mac优先使用本地Chrome）

**Mac 用户**：

1. **优先使用本地Chrome**（推荐，无需下载）：
   ```bash
   # 检查Chrome是否安装
   ls "/Applications/Google Chrome.app"
   ```
   
   代码中配置：
   ```javascript
   const browser = await playwright.chromium.launch({
     executablePath: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
   });
   ```

2. **如果没有Chrome，安装Chromium**：
   ```bash
   npx playwright install chromium
   ```

**Windows/Linux 用户**：

```bash
npx playwright install chromium
```

### 验证安装

```bash
node -e "const playwright = require('playwright'); console.log('Playwright OK');"
```

### 常见问题

| 问题 | 解决方案 |
|------|----------|
| `Could not find Chromium` | 运行 `npx playwright install chromium` |
| Mac Chrome路径错误 | 确认路径：`/Applications/Google Chrome.app/Contents/MacOS/Google Chrome` |
| 权限错误 | 使用Playwright安装的Chromium，或检查本地Chrome权限 |

## 纯表格 vs ECharts 决策标准

在阶段3图表规划时，AI会先判断使用哪种呈现方式：

### 纯表格（直接用pptxgenjs生成，无需预生成图片）
- 适合场景：数据清单、详细指标对比、SKU列表、多维度小数据量表格
- 数据特征：行数≤15行，列数≤8列，以展示精确数值为主
- 页面类型：super_3d（超单3D选品）、essential（必备品全景数据部分）、structure右侧数据表格

### ECharts复杂图表（需要预生成高清图片）
- 适合场景：趋势分析、占比分析、对比分析、分布分析、象限分析
- 数据特征：需要可视化趋势、对比、分布关系，数据量较大
- 页面类型：overview（指标卡片除外）、efficiency、structure左侧饼图、super_performance、growth、investment、targets

## 使用示例

### 示例1: 从头创建一个经营分析PPT

```
用户: 帮我做一个差旅用品的经营分析PPT
```

Skill会按数据分析型PPT流程引导：
1. 询问项目名称、目的、受众、**必需板块**
2. AI生成2-3个PPT结构方案供选择
3. 确认主题风格（可上传截图），AI生成项目风格画像并建议每页版式
4. AI建议图表/表格方案，逐页确认数据维度和系列
5. 生成Excel模板
6. 用户填好数据后，先做数据洞察，再按4步骤执行最终生成

### 示例2: 基于已有数据创建PPT

```
用户: 我有一个Excel文件，帮我做成PPT
```

Skill会：
1. 先读取Excel了解数据结构
2. 判断哪些内容适合做图表、哪些适合做表格、哪些需要补充数据
3. 建议合适的PPT结构和图表方案
4. 按6阶段流程完成

### 示例3: 数据填完了，生成最终PPT

```
用户: 数据已经填好了，生成PPT吧
```

Skill会按顺序执行：
1. 数据洞察 → 更新完善内容大纲
2. 检查视觉风格和版式适配是否明确
3. 仅预生成ECharts复杂图表（确保排版和文字不溢出）
4. 用pptxgenjs合成可编辑PPT

## 输出文件

生成的项目包含：
- `data_templates/` - Excel数据模板（含TABLE_和CHART_两种模板）
- `images/` - 图片资源文件夹
- `tmp_charts/` - 临时ECharts目录（仅存放ECharts复杂图表，纯表格不生成图片）
- `output/` - 最终PPT输出目录
- `project_state.json` - 项目状态文件
- `{project_name}_大纲.md` - 完整内容大纲（Markdown，含项目风格画像、视觉规范、结论规范、每页详情）

## 最终大纲标准

最终生成的内容大纲必须包含：

### 1. 项目风格画像与版式策略

| 项目 | 策略 |
|------|------|
| 使用场景 | ... |
| 目标受众 | ... |
| 视觉气质 | ... |
| 信息密度 | ... |
| 版式节奏 | ... |

### 2. 整体视觉规范

| 项目 | 规范 |
|------|------|
| 设计风格 | ... |
| 页面尺寸 | ... |
| 边距 | ... |
| 标题字体 | ... |
| 正文字体 | ... |
| 主色（标题栏） | ... |
| ... | ... |

### 3. 每页结论呈现规范

| 位置 | 样式 | 用途 |
|------|------|------|
| 页面标题栏 | ... | ... |
| 内容区域卡片 | ... | ... |
| 底部"核心洞察"栏 | ... | ... |
| ... | ... | ... |

### 4. 每页详情

```
### PX：页面标题

**页面目的**：...

**版式变体**：...

**页面布局**：
- 顶部标题：...
- 左侧（占X%）主图表：...
  - X轴：...
  - Y轴：...
  - 系列1：...
- 右侧（占X%）区域：...

**图表编号**：CHART_XX_XXX
```

## 自定义主题

支持3套预设主题：
- `business` - 清爽商务风（蓝色）
- `simple` - 简约专业风（灰色）
- `tech` - 科技活力风（绿色）

**提示**：阶段2时会询问用户是否有参考的风格截图可以上传。

## 页面类型

支持9种页面模板：
- `overview` - 生意概览
- `efficiency` - 单店效能
- `structure` - 品类结构
- `essential` - 必备品全景
- `super_3d` - 超单3D选品
- `super_performance` - 超单品表现
- `targets` - 目标设定
- `growth` - 增长机会
- `investment` - 投资回报

## 脚本使用

```bash
# 创建新项目
node scripts/ppt_builder.js create-project "项目名称"

# 数据洞察
node scripts/ppt_builder.js generate-insights ./data_templates

# 预生成ECharts
node scripts/ppt_builder.js pregenerate-echarts ./tmp_charts

# 生成PPT
node scripts/ppt_builder.js generate-ppt ./output/presentation.pptx

# 完整最终生成流程
node scripts/ppt_builder.js final-generation ./data_templates ./output
```
