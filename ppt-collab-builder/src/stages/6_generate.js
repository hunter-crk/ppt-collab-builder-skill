/**
 * 阶段6: 最终生成
 * 生成ECharts图表、合成最终PPT文件
 */

const readline = require('readline');
const path = require('path');
const fs = require('fs');
const PptxGenJS = require('pptxgenjs');

class GenerateStage {
  constructor(state) {
    this.state = state;
    this.rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout
    });
  }

  async run() {
    console.log('\n' + '='.repeat(80));
    console.log('阶段6: 最终生成');
    console.log('='.repeat(80) + '\n');

    await this.renderCharts();
    await this.collectImages();
    await this.generateFinalPPT();

    console.log('\n🎉 全部完成！');
    this.state.setStage(7);
  }

  ask(question) {
    return new Promise((resolve) => {
      this.rl.question(question + ' ', (answer) => {
        resolve(answer.trim());
      });
    });
  }

  async renderCharts() {
    console.log('📈 渲染ECharts图表 (此功能需要配置Playwright)...');
    console.log('   ⚠️  图表渲染功能需要在完整项目中配置\n');

    const chartsDir = path.join(this.state.paths.baseDir, 'charts');
    if (!fs.existsSync(chartsDir)) {
      fs.mkdirSync(chartsDir, { recursive: true });
    }

    this.chartsDir = chartsDir;
    console.log(`✅ 图表输出目录: ${chartsDir}\n`);
  }

  async collectImages() {
    console.log('🖼️  收集图片资源...');

    const images = [];
    const imageDir = this.state.paths.imageDir;

    if (fs.existsSync(imageDir)) {
      const pageDirs = fs.readdirSync(imageDir).filter(f =>
        fs.statSync(path.join(imageDir, f)).isDirectory()
      );

      for (const pageDir of pageDirs) {
        const pagePath = path.join(imageDir, pageDir);
        const files = fs.readdirSync(pagePath).filter(f =>
          /\.(png|jpg|jpeg|gif|webp)$/i.test(f)
        );
        for (const file of files) {
          const pageIdMatch = pageDir.match(/page_(\d+)/);
          images.push({
            pageId: pageIdMatch ? parseInt(pageIdMatch[1]) : null,
            path: path.join(pagePath, file),
            fileName: file
          });
        }
      }
    }

    this.collectedImages = images;
    console.log(`✅ 已收集 ${images.length} 张图片\n`);
  }

  async generateFinalPPT() {
    console.log('📊 生成PPT文件...');

    const pptx = new PptxGenJS();
    pptx.layout = 'WIDE';
    pptx.author = 'PPT协作构建器';
    pptx.subject = this.state.requirements.projectName;
    pptx.title = this.state.requirements.projectName;

    const theme = this.state.detailedOutline?.theme || 'business';
    const themes = require('../config/themes');
    const colors = themes[theme]?.colors || themes.business.colors;

    const pages = this.state.detailedOutline?.pages || [];

    for (const page of pages) {
      this.addPage(pptx, page, colors);
    }

    const outputPath = path.join(
      this.state.paths.outputDir,
      `${this.state.requirements.projectName.replace(/[^\w\u4e00-\u9fa5]/g, '_')}.pptx`
    );

    await pptx.writeFile({ fileName: outputPath });

    console.log(`\n✅ PPT已生成: ${outputPath}`);

    await this.generateProjectSummary(outputPath);
  }

  addPage(pptx, page, colors) {
    const slide = pptx.addSlide();

    slide.background = { color: colors.pageBg };

    slide.addShape(pptx.ShapeType.rect, {
      x: 0, y: 0, w: 13.33, h: 0.8,
      fill: { color: colors.titleBar }
    });

    slide.addText(page.title, {
      x: 0.5, y: 0.15, w: 12.33, h: 0.5,
      fontFace: 'Microsoft YaHei',
      fontSize: 24,
      bold: true,
      color: colors.titleText
    });

    slide.addText(this.state.requirements.projectName, {
      x: 0.5, y: 7, w: 10, h: 0.4,
      fontFace: 'Microsoft YaHei',
      fontSize: 10,
      color: colors.textLight
    });

    slide.addText(`${page.id}`, {
      x: 12.5, y: 7, w: 0.5, h: 0.4,
      fontFace: 'Microsoft YaHei',
      fontSize: 10,
      color: colors.textLight,
      align: 'right'
    });

    slide.addShape(pptx.ShapeType.rect, {
      x: 0.5, y: 1, w: 12.33, h: 5.8,
      fill: { color: colors.cardBg },
      line: { color: colors.border }
    });

    slide.addText('内容区域', {
      x: 1, y: 3, w: 11.33, h: 2,
      fontFace: 'Microsoft YaHei',
      fontSize: 18,
      color: colors.textLight,
      align: 'center',
      valign: 'middle'
    });
  }

  async generateProjectSummary(pptPath) {
    const summaryPath = path.join(this.state.paths.baseDir, '项目摘要.json');

    const summary = {
      projectName: this.state.requirements.projectName,
      projectPurpose: this.state.requirements.projectPurpose,
      targetAudience: this.state.requirements.targetAudience,
      generatedAt: new Date().toISOString(),
      theme: this.state.detailedOutline?.theme,
      totalPages: this.state.detailedOutline?.pages?.length || 0,
      totalCharts: this.state.chartPlan?.filter(c => !c.type || c.type !== 'images').length || 0,
      totalImages: this.state.chartPlan?.filter(c => c.type === 'images').length || 0,
      files: {
        ppt: pptPath,
        dataDir: this.state.paths.dataDir,
        imageDir: this.state.paths.imageDir
      }
    };

    fs.writeFileSync(summaryPath, JSON.stringify(summary, null, 2), 'utf-8');
    console.log(`✅ 项目摘要: ${summaryPath}\n`);

    console.log('='.repeat(80));
    console.log('📦 项目交付物:');
    console.log('='.repeat(80));
    console.log(`   📊 PPT文件: ${pptPath}`);
    console.log(`   📂 数据模板: ${this.state.paths.dataDir}`);
    console.log(`   🖼️  图片目录: ${this.state.paths.imageDir}`);
    console.log(`   📝 项目摘要: ${summaryPath}`);
    console.log('='.repeat(80));
  }

  close() {
    this.rl.close();
  }
}

module.exports = GenerateStage;
