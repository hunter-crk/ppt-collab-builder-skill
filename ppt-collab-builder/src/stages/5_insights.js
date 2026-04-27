/**
 * 阶段5: 数据洞察
 * 读取已填充的数据，生成关键结论，确认数据质量
 */

const readline = require('readline');
const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');

class InsightsStage {
  constructor(state) {
    this.state = state;
    this.rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout
    });
    this.dataCache = {};
  }

  async run() {
    console.log('\n' + '='.repeat(80));
    console.log('阶段5: 数据洞察');
    console.log('='.repeat(80) + '\n');

    await this.loadFilledData();
    await this.analyzeData();
    await this.generateInsights();
    await this.confirmInsights();

    this.state.setStage(6);
    console.log('\n✅ 阶段5完成！进入阶段6: 最终生成\n');
  }

  ask(question) {
    return new Promise((resolve) => {
      this.rl.question(question + ' ', (answer) => {
        resolve(answer.trim());
      });
    });
  }

  async loadFilledData() {
    console.log('📂 加载已填充的数据...');

    const templates = this.state.resources.dataTemplates || [];
    let loadedCount = 0;

    for (const template of templates) {
      if (template.type === 'chart' && template.templatePath) {
        if (fs.existsSync(template.templatePath)) {
          try {
            const wb = xlsx.readFile(template.templatePath);
            const dataWs = wb.Sheets['数据'];
            if (dataWs) {
              const data = xlsx.utils.sheet_to_json(dataWs, { header: 1 });
              const key = `${template.pageId}_${template.chartIndex}`;
              this.dataCache[key] = {
                raw: data,
                parsed: this.parseChartData(data),
                template: template
              };
              loadedCount++;
            }
          } catch (e) {
            console.log(`   ⚠️  无法读取 ${template.templatePath}: ${e.message}`);
          }
        }
      }
    }

    console.log(`✅ 已加载 ${loadedCount} 个图表数据\n`);
  }

  parseChartData(rawData) {
    if (!rawData || rawData.length < 2) return null;

    const headers = rawData[0];
    const rows = rawData.slice(1).filter(r => r.length > 0 && r[0]);

    return {
      headers: headers,
      xAxis: headers[0],
      series: headers.slice(1),
      dataPoints: rows.map(row => ({
        x: row[0],
        values: row.slice(1)
      }))
    };
  }

  async analyzeData() {
    console.log('🔍 分析数据...');

    const insights = [];

    for (const [key, chartData] of Object.entries(this.dataCache)) {
      if (!chartData.parsed) continue;

      const parsed = chartData.parsed;
      const template = chartData.template;

      const chartInsights = this.analyzeSingleChart(parsed, template);
      insights.push({
        pageId: template.pageId,
        pageTitle: template.pageTitle,
        chartIndex: template.chartIndex,
        chartType: template.chartType,
        insights: chartInsights
      });
    }

    this.state.setEnhancedOutline({
      ...this.state.detailedOutline,
      chartInsights: insights,
      dataCache: this.dataCache
    });

    console.log(`✅ 已分析 ${insights.length} 个图表\n`);
  }

  analyzeSingleChart(parsed, template) {
    const insights = [];
    const { dataPoints, series } = parsed;

    if (!dataPoints || dataPoints.length === 0) return insights;

    for (let s = 0; s < series.length; s++) {
      const values = dataPoints.map(d => Number(d.values[s]) || 0);
      const validValues = values.filter(v => !isNaN(v));

      if (validValues.length === 0) continue;

      const max = Math.max(...validValues);
      const min = Math.min(...validValues);
      const avg = validValues.reduce((a, b) => a + b, 0) / validValues.length;
      const first = validValues[0];
      const last = validValues[validValues.length - 1];
      const growth = first !== 0 ? ((last - first) / first * 100) : 0;

      const maxIndex = values.indexOf(max);
      const minIndex = values.indexOf(min);

      insights.push({
        series: series[s],
        max: { value: max, period: dataPoints[maxIndex]?.x },
        min: { value: min, period: dataPoints[minIndex]?.x },
        avg: avg,
        growth: growth,
        trend: growth > 5 ? '上升' : growth < -5 ? '下降' : '平稳'
      });
    }

    return insights;
  }

  async generateInsights() {
    console.log('💡 生成数据洞察...');

    const enhancedOutline = this.state.enhancedOutline;
    if (!enhancedOutline || !enhancedOutline.chartInsights) return;

    console.log('\n📊 数据概览:');
    console.log('-'.repeat(80));

    for (const chart of enhancedOutline.chartInsights) {
      console.log(`\n[${chart.pageTitle}] 图表${chart.chartIndex}:`);

      for (const insight of chart.insights) {
        const trendIcon = insight.trend === '上升' ? '📈' : insight.trend === '下降' ? '📉' : '➡️';
        console.log(`  ${trendIcon} ${insight.series}:`);
        console.log(`     峰值: ${insight.max.value.toFixed(1)} (${insight.max.period})`);
        console.log(`     谷值: ${insight.min.value.toFixed(1)} (${insight.min.period})`);
        console.log(`     平均: ${insight.avg.toFixed(1)}`);
        console.log(`     趋势: ${insight.trend} (${insight.growth >= 0 ? '+' : ''}${insight.growth.toFixed(1)}%)`);
      }
    }

    console.log('\n' + '-'.repeat(80) + '\n');
  }

  async confirmInsights() {
    console.log('📝 您可以补充业务洞察文案（可选）:');
    console.log('这些文案将显示在PPT对应页面上。\n');

    const enhancedOutline = this.state.enhancedOutline;
    if (!enhancedOutline || !enhancedOutline.chartInsights) return;

    for (const chart of enhancedOutline.chartInsights) {
      const customInsight = await this.ask(
        `[${chart.pageTitle}] 补充业务洞察（留空跳过）:`
      );
      if (customInsight) {
        chart.customInsight = customInsight;
      }
    }

    this.state.setEnhancedOutline(enhancedOutline);
  }

  close() {
    this.rl.close();
  }
}

module.exports = InsightsStage;
