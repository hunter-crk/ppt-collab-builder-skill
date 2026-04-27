---
name: ppt-collab-builder
description: This skill should be used when the user asks to create a data-driven PowerPoint presentation, build a PPT from data, generate charts and insights for a presentation, or discusses collaborative PPT building workflows. Use this skill whenever the user mentions "PPT", "PowerPoint", "presentation", "charts", "data visualization", "business analysis", or wants to create any kind of data-driven presentation, even if they don't explicitly ask for a "PPT builder".
version: 2.2.0
---

# PPT 协作构建器

协作式数据型PPT生成工具 - 通过6阶段工作流帮助用户构建专业的数据驱动型PPT。

## 工作流概述

这个skill指导用户完成一个结构化的6阶段PPT构建过程：

1. **需求分析** - 收集项目目的、受众、必需板块，AI生成2-3个PPT结构方案供选择
2. **大纲完善** - 确定项目风格画像（可上传参考截图）、视觉规范和逐页版式策略，最终大纲包含【风格画像、视觉规范、结论呈现规范、每页目的/版式变体/布局/图表编号】
3. **图表规划** - AI先给出图表建议方案（类型、维度、系列），与用户确认
4. **资源准备** - 生成Excel数据模板，指导用户准备图片
5. **数据洞察** - 分析已填数据，自动生成关键结论，更新大纲
6. **最终生成** - 按顺序执行：①数据洞察更新大纲；②检查视觉风格与版式适配；③预生成所有ECharts（确保排版和文字不溢出）；④按项目风格合成最终PPT

## 何时使用此skill

- 用户需要创建包含数据图表的PPT
- 用户有Excel数据需要可视化
- 用户需要商业分析/经营分析类PPT
- 用户希望通过多轮对话协作式构建PPT
- 用户提到"生意分析"、"经营数据"、"数据汇报"等

## 核心概念

### 页面类型

支持9种预设页面类型：
- `overview` - 生意概览（关键指标卡片）
- `efficiency` - 单店效能（趋势分析）
- `structure` - 品类结构（饼图/占比）
- `essential` - 必备品全景（商品网格）
- `super_3d` - 超单3D选品（数据表格）
- `super_performance` - 超单品表现（对比图表）
- `targets` - 目标设定（进度展示）
- `growth` - 增长机会（散点/象限）
- `investment` - 投资回报（ROI分析）

### 图表类型

- `line` - 折线图（趋势）
- `bar` - 柱状图（对比）
- `pie` - 饼图（占比）
- `area` - 面积图（累积）
- `scatter` - 散点图（分布）

### 内容呈现方式决策标准

在阶段3图表规划时，AI必须先判断使用纯表格还是ECharts复杂图表：

**纯表格（直接用pptxgenjs生成，无需预生成图片）**：
- 适合场景：数据清单、详细指标对比、SKU列表、多维度小数据量表格
- 数据特征：行数≤15行，列数≤8列，以展示精确数值为主
- 页面类型：super_3d（超单3D选品）、essential（必备品全景数据部分）、structure右侧数据表格

**ECharts复杂图表（需要预生成高清图片）**：
- 适合场景：趋势分析、占比分析、对比分析、分布分析、象限分析
- 数据特征：需要可视化趋势、对比、分布关系，数据量较大
- 页面类型：overview（指标卡片除外）、efficiency、structure左侧饼图、super_performance、growth、investment、targets

**决策原则**：AI先建议呈现方式，说明理由，再与用户确认。

### 主题风格

- `business` - 清爽商务风（蓝色系）
- `simple` - 简约专业风（灰色系）
- `tech` - 科技活力风（绿色系）

这些主题只是起点，不是固定模板。Agent必须根据项目目的、受众、行业、参考截图、内容密度生成**项目风格画像**，并据此选择每页版式变体。详细规则见 `references/adaptive-layout.md`。

### 自适应排版原则

- 保持6阶段协作流程、图表规划流程、资源准备流程、最终生成技术栈不变
- 不把所有页面都套成同一种标题栏+白卡片+左右分栏
- 每页先判断信息任务：结论页、数据密集页、对比页、趋势页、商品/图片页、机会判断页
- 根据项目风格画像决定留白、信息密度、标题形态、图表占比、洞察位置和页面节奏
- 最终PPT必须体现统一视觉系统，同时允许页面之间有明确的版式变化

## 使用指南

### 第一步：理解用户需求

先通过对话了解：
1. 项目名称（如：差旅用品经营分析）
2. PPT的主要目的（向客户展示、内部汇报、竞品分析等）
3. 目标受众（客户、管理层、销售团队等）
4. 您认为这份PPT中**必须包含**的板块有哪些？

然后AI根据必需板块，生成2-3个PPT结构方案供用户选择确认。

### 第二步：创建项目结构

根据用户确认的PPT结构，在当前目录下创建项目结构：

```
{project_name}/
├── data_templates/    # Excel数据模板
├── images/            # 图片资源文件夹
└── output/            # 最终输出
```

### 第三步：大纲完善（项目风格+每页配置）

1. 询问主题风格，并提示："您是否有参考的风格截图可以上传？"
2. 生成项目风格画像：场景、受众、视觉气质、信息密度、版式节奏、适合/避免的排版方式
3. 逐页细化时，AI先给出建议（页面类型、页面任务、版式变体、目的、布局、是否需要图表/表格/图片、排版比例），再确认
4. 最终生成的大纲必须包含：
   - **项目风格画像与版式策略**
   - **整体视觉规范**（详细表格）
   - **每页结论呈现规范**（位置、样式、用途）
   - **每页详情**：页面目的、版式变体、页面布局、图表编号

### 第四步：图表规划（AI先建议）

对每个需要数据展示的页面，按以下步骤进行：

1. **呈现方式决策**：AI先判断并建议使用**纯表格**还是**ECharts复杂图表**，说明推荐理由
2. **纯表格页面**：确定表格结构（列定义、样式规范）
3. **ECharts图表页面**：AI给出图表建议方案（图表类型、X轴/Y轴维度、数据系列），再与用户确认

**决策依据**：
- 纯表格：数据清单、≤15行≤8列、以精确数值展示为主
- ECharts：趋势/对比/分布分析、需要可视化展示、数据量较大

### 第五步：生成数据模板

为每个需要图表的页面生成Excel模板，包含：
- 说明页（图表信息、使用指南）
- 数据页（带示例数据的表格）

模板命名格式：`CHART_{pageId}_{chartIndex}.xlsx`

### 第六步：最终生成（必须按顺序执行）

当用户告知数据填完了，必须按以下顺序执行：

1. **数据洞察并更新大纲**：读取数据，生成洞察文案，更新到内容大纲
2. **检查视觉风格与版式适配**：确认项目风格画像、视觉规范、每页版式变体是否已明确；检查页面之间是否过于单一
3. **预生成ECharts复杂图表（仅复杂图表）**：
   - 只对标记为"ECharts复杂图表"的内容进行预生成
   - 确保符合排版比例，检查文字不溢出，生成2倍分辨率高清图
   - 纯表格内容跳过此步骤，直接在步骤4用pptxgenjs生成
4. **合成最终PPT**：使用pptxgenjs生成PPT
   - 纯表格：直接用pptxgenjs表格API生成
   - ECharts图表：插入预生成的高清图片
   - 插入其他图片、洞察，应用视觉规范和每页版式策略

**技术栈要求**：此步骤必须使用 **JS + ECharts(浏览器渲染) + pptxgenjs**，禁止使用Python或其他技术栈。

**ECharts生成方式说明**：
- ECharts复杂图表**不使用**pptxgenjs内置图表功能
- 使用**浏览器（Puppeteer/Playwright）+ ECharts库**渲染生成2倍分辨率高清PNG图片
- 然后用pptxgenjs将生成的图片插入PPT

## 最终大纲标准格式

最终大纲必须包含以下四个部分：

### 1. 项目风格画像与版式策略

| 项目 | 策略 |
|------|------|
| 使用场景 | ... |
| 目标受众 | ... |
| 视觉气质 | ... |
| 信息密度 | ... |
| 版式节奏 | ... |
| 适合的版式 | ... |
| 避免的版式 | ... |

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
  - ...

**图表编号**：CHART_XX_XXX
```

## 参考资源

- `references/workflow.md` - 详细的6阶段工作流说明
- `references/templates.md` - 页面模板和布局规范
- `references/adaptive-layout.md` - 项目风格画像、版式变体和最终版式审查规则
- `scripts/ppt_builder.js` - PPT构建基础参考工具（重要：每个项目需要单独编写生成脚本）

## 重要说明

### scripts/ppt_builder.js 的定位

**`scripts/ppt_builder.js` 是一个基础参考工具，不是万能的通用脚本。**

每个具体项目都需要：
1. 基于 `ppt_builder.js` 中的辅助函数和最佳实践
2. 根据项目的具体页面结构、数据格式、视觉规范
3. **单独编写项目专用的PPT生成脚本**

### 为什么需要单独脚本？

- 不同项目的页面类型、布局结构、数据来源差异很大
- 视觉规范（配色、字体、尺寸）每个项目都可能不同
- 数据处理逻辑需要针对具体Excel模板定制
- 表格和图表的组合方式每个项目都独一无二

### ppt_builder.js 提供的价值

✅ **可复用的辅助函数**：
- `PPTBuilder.cell()` / `headerCell()` / `highlightCell()` - 表格单元格生成
- `PPTBuilder.createBusinessTable()` - 标准业务表格配置
- `TABLE_COLORS` - 配色规范参考

✅ **可参考的最佳实践**：
- pptxgenjs 表格API的正确用法
- ECharts 配置生成逻辑
- 6阶段工作流的代码结构

✅ **可复用的基础方法**：
- `createProject()` - 项目结构创建
- `generateFullOutlineMarkdown()` - 大纲生成
- `pregenerateECharts()` - ECharts预生成流程

## 示例对话

**示例1：简单经营分析PPT**
```
用户：帮我做一个差旅用品的经营分析PPT
你的回应：好的！让我们通过6阶段协作来完成这个PPT。首先了解一下需求...
（按新流程：先问必需板块，再AI出方案）
```

**示例2：带数据的PPT**
```
用户：我有一个Excel表，帮我做成PPT
你的回应：好的！先让我看看你的数据，然后我们规划一下PPT结构...
```

**示例3：数据填完了**
```
用户：数据已经填好了，帮我生成PPT吧
你的回应：好的！我会按以下步骤执行：
1. 先进行数据洞察，更新完善内容大纲
2. 检查项目风格画像、视觉规范和每页版式适配是否已明确
3. 预生成所有ECharts图表（确保排版和文字不溢出）
4. 最后按项目风格合成最终PPT
让我们开始...
```

## 关键原则

1. **协作优先** - 多轮对话确认，不要假设
2. **AI先建议** - 页面配置、图表方案、呈现方式都由AI先给出建议
3. **数据驱动** - 所有图表和洞察基于实际数据
4. **专业标准** - 保持商务PPT的专业质感，最终大纲必须包含视觉规范、结论规范、每页详情
5. **自适应排版** - 根据项目风格画像选择版式变体，避免整份PPT排版单一
6. **分步生成** - 最终生成必须按顺序执行：数据洞察→检查风格与版式适配→预生成ECharts（仅复杂图表）→合成PPT
7. **迭代优化** - 随时可以返回修改之前的决策
8. **技术栈锁定** - 最终生成PPT必须使用 **JS + ECharts(浏览器渲染) + pptxgenjs** 技术栈，禁止使用Python或其他技术栈方案

## 技术栈要求（严格执行）

**必须使用以下技术栈，不得变更**：

| 层级 | 技术 | 用途 | 说明 |
|------|------|------|------|
| 运行时 | Node.js (JavaScript) | 脚本运行环境 | 所有构建脚本必须用JS编写 |
| 图表渲染 | 浏览器(Puppeteer/Playwright) + ECharts库 | 复杂图表渲染 | **不使用pptxgenjs内置图表**，用浏览器渲染生成2倍分辨率高清PNG |
| PPT生成 | pptxgenjs | 直接生成.pptx文件 | 纯表格（直接用API）、文本、插入ECharts生成的图片 |
| 数据处理 | 原生JS (fs/path) | 文件读写、数据处理 | 不使用Python pandas等第三方库 |

**纯表格生成方式**：
- 对于纯表格内容，直接使用pptxgenjs的表格API生成，不经过ECharts，无需预生成图片。

**ECharts复杂图表生成方式（重要）**：
- **不使用**pptxgenjs内置的图表功能
- 使用**浏览器（Puppeteer/Playwright）+ ECharts库**在浏览器中渲染
- 生成2倍分辨率（@2x）高清PNG图片
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

### 浏览器内核配置

**Mac 用户（推荐优先使用本地Chrome）**：

1. **检查是否已安装Chrome**：
   ```bash
   # 检查Chrome是否存在
   ls "/Applications/Google Chrome.app"
   ```

2. **如果有本地Chrome**：
   - 直接使用本地Chrome，无需额外下载
   - 在代码中配置executablePath指向本地Chrome：
     ```javascript
     const browser = await playwright.chromium.launch({
       executablePath: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
     });
     ```

3. **如果没有Chrome**：
   ```bash
   # 使用Playwright安装Chromium
   npx playwright install chromium
   ```

**Windows/Linux 用户**：

```bash
# 安装Playwright默认浏览器
npx playwright install chromium
```

### 验证安装

```bash
# 简单测试Playwright是否可用
node -e "const playwright = require('playwright'); console.log('Playwright OK');"
```

### 常见问题

| 问题 | 解决方案 |
|------|----------|
| `Could not find Chromium` | 运行 `npx playwright install chromium` |
| Mac本地Chrome无法启动 | 检查Chrome路径是否正确：`/Applications/Google Chrome.app/Contents/MacOS/Google Chrome` |
| 权限错误 | 确保Chrome有执行权限，或使用Playwright安装的Chromium |
| ECharts渲染超时 | 增加超时时间，或检查网络连接（如需加载外部资源） |

**Skill执行提醒**：在进入阶段6之前，先检查Playwright环境，如未安装则指导用户完成安装后再继续。

## 快速开始

如果用户说"开始吧"或类似的话，直接启动6阶段流程：

1. 从需求分析开始：项目名称、目的、受众、必需板块
2. AI生成2-3个PPT结构方案供选择
3. 帮助用户完善大纲（可上传风格截图，AI生成项目风格画像并建议每页版式）
4. AI建议图表方案，逐页确认
5. 生成资源模板
6. 用户填好数据后，按4步骤执行最终生成
