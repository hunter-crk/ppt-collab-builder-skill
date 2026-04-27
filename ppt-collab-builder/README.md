# PPT Collab Builder

协作式数据型PPT构建器 - 通过多轮对话协作完成数据型PPT的自动生成。

## 工作流

### 阶段1: 需求分析
- 多轮对话询问用户需求
- 确认每个部分的作用和内容
- 产出: 初步内容大纲

### 阶段2: 大纲完善
- 视觉规范定义(颜色、字体、尺寸)
- 风格定义
- 章节内容细化
- 内容排版比例
- 产出: 详细内容大纲

### 阶段3: 图表规划
- 识别需要图片/图表的页面
- 逐图确认: 图表类型、维度、数据定义
- 产出: 图表规划清单

### 阶段4: 资源准备
- 生成Excel数据模板
- 告知用户图片存放位置
- 用户填写数据、放置图片

### 阶段5: 数据洞察
- 读取填好的数据
- 自动分析生成结论/洞察
- 更新内容大纲(添加数据驱动的文案)

### 阶段6: 最终生成
- 复杂图表: ECharts渲染
- 纯表格: 直接在PPT中生成
- 整合图片、内容、洞察
- 产出: 最终PPT

## 安装

```bash
npm install
npx playwright install chromium
```

## 使用

```bash
npm start
```

## 项目结构

```
ppt-collab-builder/
├── package.json
├── README.md
├── src/
│   ├── index.js              # 主入口
│   ├── stages/               # 各个阶段处理
│   │   ├── 1_requirements.js
│   │   ├── 2_outline.js
│   │   ├── 3_charts.js
│   │   ├── 4_templates.js
│   │   ├── 5_insights.js
│   │   └── 6_generate.js
│   ├── core/                 # 核心模块
│   │   ├── chart_renderer.js
│   │   ├── ppt_generator.js
│   │   ├── data_analyzer.js
│   │   └── template_builder.js
│   └── config/               # 配置
│       ├── themes.js
│       └── defaults.js
├── data/
│   ├── templates/             # Excel数据模板
│   └── images/               # 图片资源
└── output/                   # 输出目录
```
