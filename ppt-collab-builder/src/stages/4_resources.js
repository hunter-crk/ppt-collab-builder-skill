/**
 * 阶段4: 资源准备
 * 生成数据模板Excel文件、创建图片文件夹结构
 */

const readline = require('readline');
const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');

class ResourcesStage {
  constructor(state) {
    this.state = state;
    this.rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout
    });
  }

  async run() {
    console.log('\n' + '='.repeat(80));
    console.log('阶段4: 资源准备');
    console.log('='.repeat(80) + '\n');

    await this.setupProjectDirs();
    await this.generateDataTemplates();
    await this.confirmResourceStructure();

    this.state.setStage(5);
    console.log('\n✅ 阶段4完成！进入阶段5: 数据洞察\n');
  }

  ask(question) {
    return new Promise((resolve) => {
      this.rl.question(question + ' ', (answer) => {
        resolve(answer.trim());
      });
    });
  }

  async setupProjectDirs() {
    console.log('📁 设置项目目录结构...');

    const projectName = this.state.requirements.projectName.replace(/[^\w\u4e00-\u9fa5]/g, '_');
    const baseDir = path.join(process.cwd(), 'output', projectName);
    const dataDir = path.join(baseDir, 'data_templates');
    const imageDir = path.join(baseDir, 'images');
    const outputDir = path.join(baseDir, 'output');

    [baseDir, dataDir, imageDir, outputDir].forEach(dir => {
      if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
      }
    });

    this.state.setPaths({ dataDir, imageDir, outputDir, baseDir });
    console.log(`✅ 项目目录已创建: ${baseDir}`);
    console.log(`   - ${dataDir}`);
    console.log(`   - ${imageDir}`);
    console.log(`   - ${outputDir}\n`);
  }

  async generateDataTemplates() {
    console.log('📊 生成数据模板...');

    const chartPlan = this.state.chartPlan || [];
    const dataTemplates = [];

    for (const item of chartPlan) {
      if (item.type === 'images') {
        const imageFolder = path.join(this.state.paths.imageDir, `page_${item.pageId}`);
        if (!fs.existsSync(imageFolder)) {
          fs.mkdirSync(imageFolder, { recursive: true });
        }
        dataTemplates.push({
          type: 'images',
          pageId: item.pageId,
          pageTitle: item.pageTitle,
          folder: imageFolder,
          description: item.description
        });
      } else {
        const templatePath = await this.createChartTemplate(item);
        dataTemplates.push({
          type: 'chart',
          pageId: item.pageId,
          pageTitle: item.pageTitle,
          chartIndex: item.chartIndex,
          chartType: item.chartType,
          templatePath: templatePath
        });
      }
    }

    this.state.setResources({ dataTemplates, imageFolders: [] });
    console.log(`✅ 已生成 ${dataTemplates.length} 个数据模板\n`);
  }

  async createChartTemplate(chartItem) {
    const fileName = `CHART_${String(chartItem.pageId).padStart(2, '0')}_${String(chartItem.chartIndex).padStart(3, '0')}.xlsx`;
    const filePath = path.join(this.state.paths.dataDir, fileName);

    const wb = xlsx.utils.book_new();

    // 创建数据说明页
    const infoData = [
      ['图表信息', ''],
      ['页面', `${chartItem.pageId} - ${chartItem.pageTitle}`],
      ['图表序号', chartItem.chartIndex],
      ['图表类型', chartItem.chartType],
      ['X轴', chartItem.xAxis],
      ['Y轴', chartItem.yAxis],
      ['数据系列', chartItem.series.join(', ')],
      ['', ''],
      ['使用说明', ''],
      ['1. 在"数据"页填入你的数据'],
      ['2. 保持表头格式不变'],
      ['3. 第一列为X轴数据'],
      ['4. 后续各列为数据系列']
    ];
    const infoWs = xlsx.utils.aoa_to_sheet(infoData);
    xlsx.utils.book_append_sheet(wb, infoWs, '说明');

    // 创建数据页（带示例数据）
    const sampleData = this.generateSampleData(chartItem);
    const dataWs = xlsx.utils.aoa_to_sheet(sampleData);
    xlsx.utils.book_append_sheet(wb, dataWs, '数据');

    xlsx.writeFile(wb, filePath);
    return filePath;
  }

  generateSampleData(chartItem) {
    const headers = [chartItem.xAxis, ...chartItem.series];
    const samplePeriods = ['1月', '2月', '3月', '4月', '5月', '6月'];

    const data = [headers];
    for (const period of samplePeriods) {
      const row = [period];
      for (let i = 0; i < chartItem.series.length; i++) {
        row.push(Math.floor(Math.random() * 1000) + 100);
      }
      data.push(row);
    }

    return data;
  }

  async confirmResourceStructure() {
    console.log('\n' + '-'.repeat(80));
    console.log('📦 资源结构确认:');
    console.log('-'.repeat(80));
    console.log(`项目目录: ${this.state.paths.baseDir}`);
    console.log('\n数据模板:');

    this.state.resources.dataTemplates.forEach((item, index) => {
      if (item.type === 'images') {
        console.log(`  ${index + 1}. [${item.pageTitle}] 图片文件夹: ${item.folder}`);
      } else {
        console.log(`  ${index + 1}. [${item.pageTitle}] 图表${item.chartIndex}: ${item.templatePath}`);
      }
    });

    console.log('\n下一步:');
    console.log('  1. 填充 data_templates/ 目录下的Excel文件');
    console.log('  2. 将图片放入 images/ 对应文件夹');
    console.log('  3. 完成后继续下一阶段\n');
    console.log('-'.repeat(80) + '\n');

    const confirm = await this.ask('✅ 资源结构确认完成？准备好后输入 y 继续 (y):');
    if (confirm && confirm.toLowerCase() !== 'y' && confirm.toLowerCase() !== 'yes') {
      console.log('\n⏸️  暂停中，请完成资源准备后再继续...\n');
      return this.confirmResourceStructure();
    }
  }

  close() {
    this.rl.close();
  }
}

module.exports = ResourcesStage;
