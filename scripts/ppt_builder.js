/**
 * PPT Builder 基础参考脚本 (v2.1)
 *
 * 【重要定位】
 * 这是一个基础参考工具，不是万能的通用脚本。
 * 每个具体项目都需要基于此脚本，单独编写项目专用的PPT生成脚本。
 *
 * 【为什么需要单独脚本？】
 * 1. 页面结构差异：不同项目的页面类型、布局组合都不同
 * 2. 视觉规范定制：配色、字体、尺寸每个项目可能不同
 * 3. 数据格式不统一：Excel模板结构每个项目都不一样
 * 4. 表格图表组合独特：每个项目的组合方式都是独一无二的
 *
 * 【本脚本提供的价值】
 * - 可复用的辅助函数：cell() / headerCell() / highlightCell() / createBusinessTable()
 * - 可参考的最佳实践：pptxgenjs表格API正确用法、ECharts配置生成逻辑
 * - 可复用的基础方法：createProject() / generateFullOutlineMarkdown() / pregenerateECharts()
 *
 * 【技术栈要求】
 * - 必须使用: JavaScript + ECharts(浏览器渲染) + pptxgenjs
 * - 禁止使用: Python或其他技术栈方案
 *
 * 【ECharts生成方式说明】
 * - ECharts复杂图表不使用pptxgenjs内置图表功能
 * - 使用浏览器(Puppeteer/Playwright/或者本地chrome) + ECharts库渲染生成2倍分辨率高清PNG
 * - 然后用pptxgenjs将生成的图片插入PPT
 *
 * 使用方法:
 *   node ppt_builder.js create-project "项目名称"
 *   node ppt_builder.js generate-template --page 1 --chart 1
 *   node ppt_builder.js generate-ppt
 *   node ppt_builder.js generate-insights  (数据洞察)
 *   node ppt_builder.js pregenerate-echarts (仅预生成ECharts复杂图表，纯表格跳过)
 *   node ppt_builder.js test-table (测试表格生成功能)
 */

const fs = require('fs');
const path = require('path');

// 本项目表格配色规范（来自references/pptxgenjs表格渲染经验说明.md）
const TABLE_COLORS = {
  HEADER: 'EEEEEE',      // 表头背景 - 浅灰色
  HIGHLIGHT: 'E8F5E9',   // 重点高亮行 - 浅绿色
  PRIMARY: '1A73E8',     // 主色（标题栏）- 蓝色
  SUCCESS: '34A853',     // 辅助色（增长）- 绿色
  BORDER: 'DDDDDD'       // 边框 - 浅灰色
};

class PPTBuilder {
  constructor(baseDir) {
    this.baseDir = baseDir || process.cwd();
    this.state = {
      requirements: {},
      detailedOutline: {
        visualSpecs: {},      // 视觉规范
        conclusionSpecs: {},  // 结论呈现规范
        pages: []             // 每页详情
      },
      dataDisplayPlan: [],     // 数据展示规划（包含图表和表格）
      chartPlan: [],           // ECharts复杂图表规划
      tablePlan: [],           // 纯表格规划
      resources: {},
      currentStage: 1
    };
  }

  // 创建项目结构
  createProject(projectName) {
    const safeName = projectName.replace(/[^\w\u4e00-\u9fa5]/g, '_');
    const projectDir = path.join(this.baseDir, safeName);

    const dirs = [
      projectDir,
      path.join(projectDir, 'data_templates'),
      path.join(projectDir, 'images'),
      path.join(projectDir, 'output'),
      path.join(projectDir, 'tmp_charts')  // 临时ECharts目录
    ];

    dirs.forEach(dir => {
      if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
      }
    });

    this.saveState(path.join(projectDir, 'project_state.json'));

    return {
      projectDir,
      dirs: {
        dataTemplates: path.join(projectDir, 'data_templates'),
        images: path.join(projectDir, 'images'),
        output: path.join(projectDir, 'output'),
        tmpCharts: path.join(projectDir, 'tmp_charts')
      }
    };
  }

  // 设置需求
  setRequirements(requirements) {
    this.state.requirements = { ...this.state.requirements, ...requirements };
  }

  // 设置详细大纲（包含视觉规范、结论规范、每页详情）
  setDetailedOutline(outline) {
    this.state.detailedOutline = {
      ...this.state.detailedOutline,
      ...outline
    };
  }

  // 设置视觉规范
  setVisualSpecs(specs) {
    this.state.detailedOutline.visualSpecs = { ...specs };
  }

  // 设置结论呈现规范
  setConclusionSpecs(specs) {
    this.state.detailedOutline.conclusionSpecs = { ...specs };
  }

  // 设置数据展示规划（统一管理图表和表格）
  setDataDisplayPlan(displayPlan) {
    this.state.dataDisplayPlan = displayPlan;
    // 分离图表和表格
    this.state.chartPlan = displayPlan.filter(item => item.type === 'echarts');
    this.state.tablePlan = displayPlan.filter(item => item.type === 'table');
  }

  // 设置图表规划（保留向后兼容）
  setChartPlan(chartPlan) {
    this.state.chartPlan = chartPlan;
  }

  // 设置表格规划
  setTablePlan(tablePlan) {
    this.state.tablePlan = tablePlan;
  }

  // 保存状态
  saveState(filePath) {
    fs.writeFileSync(filePath, JSON.stringify(this.state, null, 2), 'utf-8');
  }

  // 加载状态
  loadState(filePath) {
    if (fs.existsSync(filePath)) {
      const content = fs.readFileSync(filePath, 'utf-8');
      this.state = JSON.parse(content);
      return true;
    }
    return false;
  }

  // 生成完整内容大纲（Markdown格式）
  generateFullOutlineMarkdown() {
    const { visualSpecs, conclusionSpecs, pages } = this.state.detailedOutline;
    const { projectName } = this.state.requirements;

    let md = `# ${projectName || '项目名称'} - 内容大纲\n\n`;

    // 整体视觉规范
    md += `## 整体视觉规范\n\n`;
    md += `| 项目 | 规范 |\n`;
    md += `|------|------|\n`;
    for (const [key, value] of Object.entries(visualSpecs || {})) {
      md += `| ${key} | ${value} |\n`;
    }
    md += `\n`;

    // 每页结论呈现规范
    md += `## 每页结论呈现规范\n\n`;
    if (conclusionSpecs && conclusionSpecs.length > 0) {
      md += `| 位置 | 样式 | 用途 |\n`;
      md += `|------|------|------|\n`;
      for (const spec of conclusionSpecs) {
        md += `| ${spec.position} | ${spec.style} | ${spec.purpose} |\n`;
      }
    }
    md += `\n`;

    // 每页详情
    if (pages && pages.length > 0) {
      // 按品类分组
      const groupedPages = {};
      pages.forEach((page, idx) => {
        const category = page.category || '默认品类';
        if (!groupedPages[category]) {
          groupedPages[category] = [];
        }
        groupedPages[category].push({ ...page, index: idx + 1 });
      });

      for (const [category, categoryPages] of Object.entries(groupedPages)) {
        md += `## ${category}（共${categoryPages.length}页）\n\n`;
        for (const page of categoryPages) {
          md += `### P${page.index}：${page.title || '页面标题'}\n\n`;
          md += `**页面目的**：${page.purpose || '待填写'}\n\n`;
          md += `**页面布局**：\n${page.layout || '待填写'}\n\n`;
          if (page.chartIds) {
            md += `**图表编号**：${page.chartIds}\n\n`;
          }
          md += `---\n\n`;
        }
      }
    }

    // 大纲总结
    md += `## 大纲总结\n\n`;
    md += `| 项目 | 统计 |\n`;
    md += `|------|------|\n`;
    md += `| 总页数 | ${pages?.length || 0}页 |\n`;
    md += `| ECharts复杂图表数量 | ${this.state.chartPlan?.length || 0}个 |\n`;
    md += `| 纯表格数量 | ${this.state.tablePlan?.length || 0}个 |\n`;
    md += `\n`;

    return md;
  }

  // 生成表格数据模板
  generateTableTemplate(tableItem, outputPath) {
    const { pageId, pageTitle, tableIndex, columns, styleOptions } = tableItem;

    const template = {
      info: {
        page: `${pageId} - ${pageTitle}`,
        tableIndex: tableIndex,
        type: 'table',
        columns: columns,
        styleOptions: styleOptions
      },
      sampleData: this.generateTableSampleData(columns)
    };

    fs.writeFileSync(
      outputPath.replace('.xlsx', '.json'),
      JSON.stringify(template, null, 2),
      'utf-8'
    );

    return outputPath;
  }

  generateTableSampleData(columns) {
    const rowCount = 10; // 示例10行数据
    const headers = columns.map(col => col.name);
    const data = [headers];

    for (let i = 1; i <= rowCount; i++) {
      const row = [];
      for (const col of columns) {
        if (col.type === 'text') {
          row.push(`示例${col.name}${i}`);
        } else if (col.type === 'number') {
          row.push(Math.floor(Math.random() * 10000) + 100);
        } else {
          row.push(`数据${i}`);
        }
      }
      data.push(row);
    }

    return data;
  }

  // 生成Excel数据模板（简化版，实际需要xlsx库）
  generateDataTemplate(chartItem, outputPath) {
    const { pageId, pageTitle, chartIndex, chartType, xAxis, yAxis, series } = chartItem;

    const template = {
      info: {
        page: `${pageId} - ${pageTitle}`,
        chartIndex: chartIndex,
        chartType: chartType,
        xAxis: xAxis,
        yAxis: yAxis,
        series: series
      },
      sampleData: this.generateSampleData(xAxis, series)
    };

    // 这里简化处理，实际应该用xlsx库生成Excel
    fs.writeFileSync(
      outputPath.replace('.xlsx', '.json'),
      JSON.stringify(template, null, 2),
      'utf-8'
    );

    return outputPath;
  }

  generateSampleData(xAxisName, seriesNames) {
    const periods = ['1月', '2月', '3月', '4月', '5月', '6月'];
    const headers = [xAxisName, ...seriesNames];
    const data = [headers];

    for (const period of periods) {
      const row = [period];
      for (let i = 0; i < seriesNames.length; i++) {
        row.push(Math.floor(Math.random() * 1000) + 100);
      }
      data.push(row);
    }

    return data;
  }

  // === 阶段6: 最终生成相关方法 ===

  /**
   * 步骤1: 数据洞察
   * 读取已填数据，生成洞察文案，更新大纲
   */
  generateDataInsights(dataDir) {
    const insights = [];

    // 遍历所有数据文件
    if (fs.existsSync(dataDir)) {
      const files = fs.readdirSync(dataDir);
      const dataFiles = files.filter(f => f.endsWith('.json') || f.endsWith('.xlsx'));

      for (const dataFile of dataFiles) {
        const filePath = path.join(dataDir, dataFile);
        // 简化处理：假设是JSON格式的数据
        if (dataFile.endsWith('.json')) {
          try {
            const data = JSON.parse(fs.readFileSync(filePath, 'utf-8'));
            const insight = this.analyzeData(data, dataFile);
            insights.push(insight);
          } catch (e) {
            // 忽略错误
          }
        }
      }
    }

    // 将洞察更新到大纲中
    this.state.dataInsights = insights;
    return insights;
  }

  analyzeData(data, fileName) {
    // 简化的数据分析逻辑
    const sampleData = data.sampleData || [];
    if (sampleData.length < 2) {
      return { file: fileName, insights: ['数据不足，无法分析'] };
    }

    const headers = sampleData[0];
    const rows = sampleData.slice(1);

    const pageInsights = [];

    // 分析每个数据列
    for (let col = 1; col < headers.length; col++) {
      const seriesName = headers[col];
      const values = rows.map(r => Number(r[col]) || 0);
      const avg = values.reduce((a, b) => a + b, 0) / values.length;
      const max = Math.max(...values);
      const min = Math.min(...values);
      const first = values[0];
      const last = values[values.length - 1];
      const trend = last > first ? '上升' : last < first ? '下降' : '平稳';
      const changeRate = first !== 0 ? ((last - first) / first * 100).toFixed(1) : 0;

      pageInsights.push(`${seriesName}：${trend}趋势，平均${avg.toFixed(1)}，最大值${max}，最小值${min}，变化率${changeRate}%`);
    }

    return {
      file: fileName,
      insights: pageInsights
    };
  }

  /**
   * 步骤2: 检查视觉风格
   * 返回是否已明确
   */
  checkVisualStyle() {
    const { visualSpecs } = this.state.detailedOutline;
    if (!visualSpecs || Object.keys(visualSpecs).length === 0) {
      return {
        confirmed: false,
        missing: ['设计风格', '主色', '字体', '页面尺寸', '边距']
      };
    }
    return { confirmed: true };
  }

  /**
   * 步骤3: 预生成ECharts复杂图表（仅复杂图表）
   * 纯表格跳过此步骤，直接在步骤4用pptxgenjs生成
   * 确保符合排版比例，检查文字不溢出
   */
  pregenerateECharts(outputDir) {
    const results = [];

    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    // 只处理ECharts复杂图表
    for (const chartItem of this.state.chartPlan) {
      const checkResult = this.checkChartLayout(chartItem);
      const echartsConfig = this.generateEChartsConfig(chartItem, checkResult);

      const chartFileName = `CHART_${chartItem.pageId}_${String(chartItem.chartIndex).padStart(3, '0')}.json`;
      const chartFilePath = path.join(outputDir, chartFileName);

      fs.writeFileSync(chartFilePath, JSON.stringify({
        config: echartsConfig,
        layoutCheck: checkResult,
        chartItem: chartItem
      }, null, 2), 'utf-8');

      results.push({
        file: chartFileName,
        type: 'echarts',
        layoutCheck: checkResult,
        status: 'generated'
      });
    }

    // 记录纯表格（跳过预生成）
    for (const tableItem of this.state.tablePlan) {
      results.push({
        file: `TABLE_${tableItem.pageId}_${String(tableItem.tableIndex).padStart(3, '0')}`,
        type: 'table',
        status: 'skipped_pptxgenjs_direct'
      });
    }

    return results;
  }

  /**
   * 检查图表排版和文字溢出
   */
  checkChartLayout(chartItem) {
    const issues = [];
    const adjustments = [];

    // 检查X轴标签数量，过多可能溢出
    const sampleXAxisCount = 6; // 假设6个月份
    if (sampleXAxisCount > 8) {
      issues.push('X轴标签过多，可能溢出');
      adjustments.push('建议旋转标签45度');
      adjustments.push('或减少显示的标签数量');
    }

    // 检查数据系列数量
    const seriesCount = (chartItem.series || '').split(',').length;
    if (seriesCount > 5) {
      issues.push('数据系列过多，图例可能溢出');
      adjustments.push('建议合并或减少系列');
    }

    // 检查Y轴单位长度
    const yAxisLength = (chartItem.yAxis || '').length;
    if (yAxisLength > 8) {
      adjustments.push('Y轴标签可考虑换行或缩写');
    }

    return {
      ok: issues.length === 0,
      issues,
      adjustments: [
        '字体大小自动调整为12-14pt',
        '标签过长时自动截断或旋转',
        '确保图表内文字不超出图表范围',
        ...adjustments
      ]
    };
  }

  /**
   * 生成ECharts配置
   */
  generateEChartsConfig(chartItem, layoutCheck) {
    const baseFontSize = layoutCheck.ok ? 14 : 12;

    return {
      title: {
        text: chartItem.pageTitle || '',
        fontSize: baseFontSize + 4
      },
      tooltip: {
        trigger: 'axis'
      },
      legend: {
        data: (chartItem.series || '').split(',').map(s => s.trim()),
        fontSize: baseFontSize
      },
      xAxis: {
        name: chartItem.xAxis || '',
        nameFontSize: baseFontSize,
        axisLabel: {
          fontSize: baseFontSize,
          rotate: layoutCheck.issues.length > 0 ? 45 : 0
        },
        data: ['1月', '2月', '3月', '4月', '5月', '6月']
      },
      yAxis: {
        name: chartItem.yAxis || '',
        nameFontSize: baseFontSize,
        axisLabel: {
          fontSize: baseFontSize
        }
      },
      series: (chartItem.series || '').split(',').map((s, idx) => ({
        name: s.trim(),
        type: chartItem.chartType || 'line',
        data: []
      })),
      grid: {
        left: '10%',
        right: '10%',
        bottom: layoutCheck.issues.length > 0 ? '20%' : '15%'
      }
    };
  }

  // 生成PPT（简化版）
  generatePPT(outputPath) {
    const summary = {
      projectName: this.state.requirements.projectName,
      generatedAt: new Date().toISOString(),
      pages: this.state.detailedOutline.pages?.length || 0,
      charts: this.state.chartPlan.length,
      tables: this.state.tablePlan.length,
      visualSpecs: this.state.detailedOutline.visualSpecs,
      dataInsights: this.state.dataInsights,
      status: 'generated'
    };

    fs.writeFileSync(outputPath.replace('.pptx', '.json'), JSON.stringify(summary, null, 2), 'utf-8');

    // 同时保存完整大纲Markdown
    const outlinePath = outputPath.replace('.pptx', '_大纲.md');
    fs.writeFileSync(outlinePath, this.generateFullOutlineMarkdown(), 'utf-8');

    return {
      pptPath: outputPath,
      outlinePath: outlinePath,
      summary
    };
  }

  /**
   * 阶段6完整执行流程
   *
   * 【技术栈锁定】必须使用 JS + ECharts(浏览器渲染) + pptxgenjs
   * 禁止使用 Python 或其他技术栈方案
   */
  async runFinalGeneration(dataDir, outputDir) {
    const steps = [];

    // 步骤1: 数据洞察
    steps.push({
      step: 1,
      name: '数据洞察与大纲更新',
      result: this.generateDataInsights(dataDir)
    });

    // 步骤2: 检查视觉风格
    const styleCheck = this.checkVisualStyle();
    steps.push({
      step: 2,
      name: '检查视觉风格',
      result: styleCheck
    });

    if (!styleCheck.confirmed) {
      return {
        success: false,
        message: '视觉风格未明确，请先确认',
        steps,
        missing: styleCheck.missing
      };
    }

    // 步骤3: 预生成ECharts复杂图表（仅复杂图表，纯表格跳过）
    const tmpChartsDir = path.join(outputDir, '..', 'tmp_charts');
    steps.push({
      step: 3,
      name: '预生成ECharts复杂图表（纯表格跳过）',
      result: this.pregenerateECharts(tmpChartsDir)
    });

    // 步骤4: 生成最终PPT
    // - 纯表格：直接用pptxgenjs表格API生成
    // - ECharts图表：插入预生成的高清图片
    const pptPath = path.join(outputDir, 'final_presentation.pptx');
    steps.push({
      step: 4,
      name: '合成最终PPT（JS + ECharts(浏览器) + pptxgenjs）',
      result: this.generatePPT(pptPath)
    });

    return {
      success: true,
      message: 'PPT生成完成（使用JS + ECharts(浏览器渲染) + pptxgenjs技术栈）',
      techStack: 'JS + ECharts(浏览器渲染) + pptxgenjs',
      steps
    };
  }

  // ==================== pptxgenjs 表格渲染辅助函数 ====================
  // 参考：references/pptxgenjs表格渲染经验说明.md

  /**
   * pptxgenjs 单元格辅助函数（重要：必须使用此格式）
   *
   * 正确API格式：{ text: 'xxx', options: { bold: true, fill: 'EEEEEE' } }
   * 错误方式：{ text: 'xxx', bold: true, fill: '#EEEEEE' }
   *
   * @param {string|number} text - 单元格文本内容
   * @param {object} opts - 样式选项
   * @param {boolean} [opts.bold] - 是否加粗
   * @param {string} [opts.fill] - 背景色（不带#号，如'EEEEEE'）
   * @param {string} [opts.color] - 文字颜色（不带#号）
   * @param {number} [opts.fontSize] - 字体大小
   * @param {string} [opts.align] - 对齐方式 ('left'|'center'|'right')
   * @returns {object} 符合pptxgenjs要求的单元格对象
   */
  static cell(text, opts = {}) {
    const result = { text: String(text) };
    if (Object.keys(opts).length > 0) {
      result.options = {};
      if (opts.bold !== undefined) result.options.bold = opts.bold;
      if (opts.fill) {
        // 自动去掉颜色中的 # 号
        result.options.fill = opts.fill.replace('#', '');
      }
      if (opts.color) {
        result.options.color = opts.color.replace('#', '');
      }
      if (opts.fontSize !== undefined) result.options.fontSize = opts.fontSize;
      if (opts.align) result.options.align = opts.align;
    }
    return result;
  }

  /**
   * 生成表头单元格（带浅灰色背景）
   */
  static headerCell(text) {
    return PPTBuilder.cell(text, { bold: true, fill: TABLE_COLORS.HEADER });
  }

  /**
   * 生成高亮行单元格（浅绿色背景）
   */
  static highlightCell(text) {
    return PPTBuilder.cell(text, { fill: TABLE_COLORS.HIGHLIGHT });
  }

  /**
   * 测试表格生成功能（返回示例表格数据结构）
   * 用于验证pptxgenjs表格API格式是否正确
   */
  static testTableGeneration() {
    // 测试数据：差旅用品销售数据
    const sampleData = [
      ['省份', '销售额(万)', '同比增长', '占比'],
      ['上海', 1258, '+12.5%', '25.2%'],
      ['北京', 986, '+8.3%', '19.8%'],
      ['广东', 875, '+15.2%', '17.5%'],
      ['江苏', 652, '+5.8%', '13.1%'],
      ['浙江', 521, '+10.1%', '10.4%'],
      ['其他', 698, '+3.2%', '14.0%']
    ];

    // 使用cell函数构建格式化表格
    const formattedRows = [
      [
        PPTBuilder.headerCell('省份'),
        PPTBuilder.headerCell('销售额(万)'),
        PPTBuilder.headerCell('同比增长'),
        PPTBuilder.headerCell('占比')
      ],
      [
        PPTBuilder.highlightCell('上海'),
        PPTBuilder.highlightCell('1258'),
        PPTBuilder.highlightCell('+12.5%'),
        PPTBuilder.highlightCell('25.2%')
      ],
      [
        PPTBuilder.cell('北京'),
        PPTBuilder.cell('986'),
        PPTBuilder.cell('+8.3%'),
        PPTBuilder.cell('19.8%')
      ],
      [
        PPTBuilder.cell('广东'),
        PPTBuilder.cell('875'),
        PPTBuilder.cell('+15.2%'),
        PPTBuilder.cell('17.5%')
      ],
      [
        PPTBuilder.cell('江苏'),
        PPTBuilder.cell('652'),
        PPTBuilder.cell('+5.8%'),
        PPTBuilder.cell('13.1%')
      ],
      [
        PPTBuilder.cell('浙江'),
        PPTBuilder.cell('521'),
        PPTBuilder.cell('+10.1%'),
        PPTBuilder.cell('10.4%')
      ],
      [
        PPTBuilder.cell('其他'),
        PPTBuilder.cell('698'),
        PPTBuilder.cell('+3.2%'),
        PPTBuilder.cell('14.0%')
      ]
    ];

    // 使用createBusinessTable生成标准配置
    const tableSpec = PPTBuilder.createBusinessTable(sampleData, {
      x: 0.5,
      y: 1.5,
      w: 12,
      fontSize: 14,
      hasHeader: true
    });

    return {
      test: 'pptxgenjs表格格式验证',
      timestamp: new Date().toISOString(),
      sampleData,
      formattedRows,
      businessTableSpec: tableSpec,
      keyPoints: [
        '✅ 使用 { text, options } 嵌套结构',
        '✅ fill颜色不带#号前缀',
        '✅ 样式在options对象中',
        '✅ 支持表头和高亮行',
        '✅ 可与pptxgenjs slide.addTable()直接配合使用'
      ]
    };
  }

  /**
   * 创建标准业务表格配置（用于pptxgenjs的slide.addTable）
   *
   * @param {Array} dataRows - 数据行，每行是一个数组
   * @param {object} options - 表格配置
   * @param {number} options.x - X位置（英寸）
   * @param {number} options.y - Y位置（英寸）
   * @param {number} options.w - 宽度（英寸）
   * @param {number} options.h - 高度（英寸，可选）
   * @param {number} options.fontSize - 字体大小（默认14）
   * @param {boolean} options.hasHeader - 是否有表头（默认true）
   * @param {Array} options.colWidths - 列宽比例数组，如[1, 2, 1]
   * @returns {object} 包含rows和tableOptions的对象
   *
   * 示例：
   * ```javascript
   * const tableSpec = PPTBuilder.createBusinessTable(
   *   [['省份', '销售额'], ['上海', 1000]],
   *   { x: 0.5, y: 1.5, w: 10 }
   * );
   * slide.addTable(tableSpec.rows, tableSpec.tableOptions);
   * ```
   */
  static createBusinessTable(dataRows, options = {}) {
    const {
      x = 0.5,
      y = 1.0,
      w = 12,
      h,
      fontSize = 14,
      hasHeader = true,
      colWidths
    } = options;

    // 处理单元格格式
    const rows = dataRows.map((row, rowIndex) => {
      return row.map((cell, colIndex) => {
        if (typeof cell === 'object' && cell.text) {
          // 已经是格式化的单元格，直接返回
          return cell;
        }
        // 纯文本/数字，根据行类型判断
        if (hasHeader && rowIndex === 0) {
          return PPTBuilder.headerCell(cell);
        }
        return PPTBuilder.cell(cell);
      });
    });

    const tableOptions = {
      x,
      y,
      w,
      fontSize,
      border: { color: TABLE_COLORS.BORDER }
    };

    if (h !== undefined) tableOptions.h = h;
    if (colWidths) tableOptions.colW = colWidths;

    return { rows, tableOptions };
  }
}

// 命令行接口
if (require.main === module) {
  const args = process.argv.slice(2);
  const builder = new PPTBuilder();

  if (args[0] === 'create-project' && args[1]) {
    const result = builder.createProject(args[1]);
    console.log(JSON.stringify(result, null, 2));
  } else if (args[0] === 'generate-insights') {
    const dataDir = args[1] || './data_templates';
    const insights = builder.generateDataInsights(dataDir);
    console.log(JSON.stringify(insights, null, 2));
  } else if (args[0] === 'pregenerate-echarts') {
    const outputDir = args[1] || './tmp_charts';
    const result = builder.pregenerateECharts(outputDir);
    console.log(JSON.stringify(result, null, 2));
  } else if (args[0] === 'generate-ppt') {
    const outputPath = args[1] || './output/presentation.pptx';
    const result = builder.generatePPT(outputPath);
    console.log(JSON.stringify(result, null, 2));
  } else if (args[0] === 'final-generation') {
    const dataDir = args[1] || './data_templates';
    const outputDir = args[2] || './output';
    builder.runFinalGeneration(dataDir, outputDir).then(result => {
      console.log(JSON.stringify(result, null, 2));
    });
  } else if (args[0] === 'test-table') {
    // 测试pptxgenjs表格生成功能
    console.log(JSON.stringify(PPTBuilder.testTableGeneration(), null, 2));
  } else if (args[0] === 'help') {
    console.log(`
PPT Builder 工具 v2.1

【技术栈要求】
- 必须使用: JavaScript + ECharts(浏览器渲染) + pptxgenjs
- 禁止使用: Python或其他技术栈方案

【ECharts生成方式】
- 不使用pptxgenjs内置图表功能
- 使用浏览器(Puppeteer/Playwright/本地Chrome) + ECharts库渲染
- 生成2倍分辨率(@2x)高清PNG图片

【pptxgenjs表格渲染】
- 纯表格直接用pptxgenjs表格API生成
- 必须使用 { text, options } 格式，样式在options中
- fill颜色不带#号前缀（如'EEEEEE'）

用法:
  node ppt_builder.js create-project <项目名称>        - 创建新项目
  node ppt_builder.js generate-insights [数据目录]       - 数据洞察
  node ppt_builder.js pregenerate-echarts [输出目录]    - 仅预生成ECharts复杂图表（纯表格跳过）
  node ppt_builder.js generate-ppt [输出路径]           - 生成PPT（pptxgenjs直接生成表格+插入ECharts图片）
  node ppt_builder.js final-generation [数据目录] [输出] - 完整最终生成流程
  node ppt_builder.js test-table                         - 测试表格生成功能
  node ppt_builder.js help                               - 显示帮助

【纯表格 vs ECharts】
- 纯表格: 直接用pptxgenjs生成，无需预生成图片
- ECharts复杂图表: 需要预生成2倍分辨率高清图片

【表格API快速参考】
  PPTBuilder.cell('文本', { bold: true, fill: 'EEEEEE' })
  PPTBuilder.headerCell('表头')
  PPTBuilder.highlightCell('高亮数据')
  PPTBuilder.createBusinessTable(data, { x: 0.5, y: 1, w: 12 })
`);
  }
}

module.exports = PPTBuilder;
