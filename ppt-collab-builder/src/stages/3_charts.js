/**
 * 阶段3: 图表规划
 * 识别需要图片/图表的页面，逐图确认图表类型、维度、数据定义
 */

const readline = require('readline');
const defaults = require('../config/defaults');

class ChartsStage {
  constructor(state) {
    this.state = state;
    this.rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout
    });
  }

  async run() {
    console.log('\n' + '='.repeat(80));
    console.log('阶段3: 图表规划');
    console.log('='.repeat(80) + '\n');

    await this.planCharts();
    await this.confirmChartPlan();

    this.state.setStage(4);
    console.log('\n✅ 阶段3完成！进入阶段4: 资源准备\n');
  }

  ask(question) {
    return new Promise((resolve) => {
      this.rl.question(question + ' ', (answer) => {
        resolve(answer.trim());
      });
    });
  }

  async planCharts() {
    const pages = this.state.detailedOutline.pages || [];
    const chartPlan = [];

    console.log('📊 让我们逐个规划需要图表的页面:\n');

    for (const page of pages) {
      if (!page.hasChart && !page.hasTable && !page.hasImages) {
        continue;
      }

      console.log(`--- 第${page.id}页: ${page.title} ---`);

      if (page.hasChart) {
        const chartCount = parseInt(await this.ask('  这页有几个图表？(默认1):') || '1');

        for (let i = 0; i < chartCount; i++) {
          console.log(`\n  📈 图表${i + 1}:`);

          const chartType = await this.chooseChartType();
          const xAxis = await this.ask('  X轴是什么？(例如: 月份、渠道):');
          const yAxis = await this.ask('  Y轴是什么？(例如: 销售额、店铺数):');
          const series = await this.ask('  有哪些数据系列？(用逗号分隔，例如: 小柴购,TOP5平均):');
          const dataSource = await this.ask('  数据来源说明？(可选):');

          chartPlan.push({
            pageId: page.id,
            pageTitle: page.title,
            chartIndex: i + 1,
            chartType: chartType,
            xAxis: xAxis,
            yAxis: yAxis,
            series: series.split(',').map(s => s.trim()),
            dataSource: dataSource
          });
        }
      }

      if (page.hasImages) {
        console.log('  🖼️  图片资源:');
        const imageDesc = await this.ask('  需要哪些图片？请描述一下:');
        chartPlan.push({
          pageId: page.id,
          pageTitle: page.title,
          type: 'images',
          description: imageDesc
        });
      }

      console.log('');
    }

    this.state.setChartPlan(chartPlan);
    console.log(`✅ 已规划 ${chartPlan.length} 个图表/图片资源\n`);
  }

  async chooseChartType() {
    console.log('  请选择图表类型:');
    defaults.chartTypes.forEach((type, index) => {
      console.log(`    ${index + 1}. ${type.name} - ${type.description}`);
    });

    const choice = await this.ask('  请输入选项编号（默认1）:');
    const typeIndex = Math.max(0, Math.min(defaults.chartTypes.length - 1, (parseInt(choice) || 1) - 1));
    return defaults.chartTypes[typeIndex].id;
  }

  async confirmChartPlan() {
    console.log('\n' + '-'.repeat(80));
    console.log('📊 图表规划确认:');
    console.log('-'.repeat(80));

    this.state.chartPlan.forEach((item, index) => {
      if (item.type === 'images') {
        console.log(`  ${index + 1}. [${item.pageTitle}] 图片: ${item.description}`);
      } else {
        console.log(`  ${index + 1}. [${item.pageTitle}] 图表${item.chartIndex}: ${item.chartType}`);
        console.log(`     X轴: ${item.xAxis}, Y轴: ${item.yAxis}`);
        console.log(`     系列: ${item.series.join(', ')}`);
      }
    });
    console.log('-'.repeat(80) + '\n');

    const confirm = await this.ask('✅ 这个图表规划是否正确？(y/n):');
    if (confirm.toLowerCase() !== 'y' && confirm.toLowerCase() !== 'yes') {
      console.log('\n🔄 让我们重新规划图表...\n');
      return this.planCharts();
    }
  }

  close() {
    this.rl.close();
  }
}

module.exports = ChartsStage;
