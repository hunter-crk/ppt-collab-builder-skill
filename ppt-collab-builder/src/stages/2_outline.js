/**
 * 阶段2: 大纲完善
 * 确定视觉规范、风格、章节内容、排版比例
 */

const readline = require('readline');
const themes = require('../config/themes');
const defaults = require('../config/defaults');

class OutlineStage {
  constructor(state) {
    this.state = state;
    this.rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout
    });
  }

  async run() {
    console.log('\n' + '='.repeat(80));
    console.log('阶段2: 大纲完善');
    console.log('='.repeat(80) + '\n');

    await this.chooseTheme();
    await this.confirmVisualSpec();
    await this.refinePageDetails();
    await this.confirmDetailedOutline();

    this.state.setStage(3);
    console.log('\n✅ 阶段2完成！进入阶段3: 图表规划\n');
  }

  ask(question) {
    return new Promise((resolve) => {
      this.rl.question(question + ' ', (answer) => {
        resolve(answer.trim());
      });
    });
  }

  async chooseTheme() {
    console.log('🎨 请选择PPT主题风格:');
    const themeList = Object.entries(themes);
    themeList.forEach(([key, theme], index) => {
      console.log(`  ${index + 1}. ${theme.name}`);
    });

    const choice = await this.ask('\n请输入选项编号（默认1）:');
    const themeIndex = Math.max(0, Math.min(themeList.length - 1, (parseInt(choice) || 1) - 1));
    const selectedTheme = themeList[themeIndex];

    this.state.setDetailedOutline({
      theme: selectedTheme[0],
      visualSpec: {
        colors: selectedTheme[1].colors,
        fonts: defaults.theme.fonts,
        fontSize: defaults.theme.fontSize
      }
    });

    console.log(`✅ 已选择主题: ${selectedTheme[1].name}\n`);
  }

  async confirmVisualSpec() {
    console.log('📐 当前视觉规范:');
    console.log('  PPT尺寸: 13.33 × 7.5 英寸（宽屏）');
    console.log('  边距: 0.5 英寸');
    console.log('  标题栏高度: 0.8 英寸');
    console.log('  标题字体: 思源黑体 Bold 24pt');
    console.log('  正文字体: 微软雅黑 Regular 14pt\n');

    const confirm = await this.ask('是否使用默认视觉规范？(y/n，默认y):');
    if (confirm.toLowerCase() === 'n' || confirm.toLowerCase() === 'no') {
      console.log('⚠️ 自定义视觉规范功能开发中...\n');
    }
  }

  async refinePageDetails() {
    console.log('📄 现在让我们细化每个页面的内容:');
    const pages = this.state.requirements.initialOutline || [];

    const detailedPages = [];
    for (const page of pages) {
      console.log(`\n--- 第${page.id}页: ${page.title} ---`);

      const layout = await this.ask('  请选择页面类型: [overview/efficiency/structure/essential/super_3d/super_performance/targets/growth/investment]:');
      const hasChart = await this.ask('  这页需要图表吗？(y/n):');
      const hasTable = await this.ask('  这页需要表格吗？(y/n):');
      const hasImages = await this.ask('  这页需要图片吗？(y/n):');
      const layoutRatio = await this.ask('  内容排版比例（例如: 左60%图表，右40%数据，默认自动):');

      detailedPages.push({
        id: page.id,
        title: page.title,
        purpose: page.purpose,
        type: layout || 'overview',
        hasChart: hasChart.toLowerCase() === 'y',
        hasTable: hasTable.toLowerCase() === 'y',
        hasImages: hasImages.toLowerCase() === 'y',
        layoutRatio: layoutRatio || 'auto'
      });
    }

    this.state.setDetailedOutline({ pages: detailedPages });
    console.log(`\n✅ 已完成 ${detailedPages.length} 个页面的详细配置\n`);
  }

  async confirmDetailedOutline() {
    console.log('\n' + '-'.repeat(80));
    console.log('📋 详细内容大纲确认:');
    console.log('-'.repeat(80));
    const themeName = themes[this.state.detailedOutline.theme]?.name || 'business';
    console.log(`主题: ${themeName}\n`);
    console.log('页面详情:');
    this.state.detailedOutline.pages?.forEach((page) => {
      console.log(`  ${page.id}. ${page.title} [${page.type}]`);
      const features = [];
      if (page.hasChart) features.push('图表');
      if (page.hasTable) features.push('表格');
      if (page.hasImages) features.push('图片');
      if (features.length > 0) console.log(`     包含: ${features.join(', ')}`);
    });
    console.log('-'.repeat(80) + '\n');

    const confirm = await this.ask('✅ 这个详细大纲是否正确？(y/n):');
    if (confirm.toLowerCase() !== 'y' && confirm.toLowerCase() !== 'yes') {
      console.log('\n🔄 让我们重新细化页面详情...\n');
      return this.refinePageDetails();
    }
  }

  close() {
    this.rl.close();
  }
}

module.exports = OutlineStage;
