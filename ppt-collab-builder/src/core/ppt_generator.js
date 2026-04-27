/**
 * PPT生成核心模块
 * 支持多种页面模板，数据驱动的内容生成
 */

const PptxGenJS = require('pptxgenjs');
const themes = require('../config/themes');

class PPTGenerator {
  constructor(state) {
    this.state = state;
    this.pptx = new PptxGenJS();
    this.pptx.layout = 'WIDE';
    this.pptx.author = 'PPT协作构建器';
    this.pptx.subject = state.requirements.projectName;
    this.pptx.title = state.requirements.projectName;

    const themeKey = state.detailedOutline?.theme || 'business';
    this.colors = themes[themeKey]?.colors || themes.business.colors;
  }

  getThemeColors() {
    return this.colors;
  }

  addTitleSlide(title, subtitle = '') {
    const slide = this.pptx.addSlide();
    slide.background = { color: this.colors.pageBg };

    slide.addShape(this.pptx.ShapeType.rect, {
      x: 0, y: 0, w: 13.33, h: 7.5,
      fill: { color: this.colors.primary }
    });

    slide.addText(title, {
      x: 0.5, y: 2.5, w: 12.33, h: 1.5,
      fontFace: 'Microsoft YaHei',
      fontSize: 44,
      bold: true,
      color: this.colors.white,
      align: 'center',
      valign: 'middle'
    });

    if (subtitle) {
      slide.addText(subtitle, {
        x: 0.5, y: 4.2, w: 12.33, h: 0.8,
        fontFace: 'Microsoft YaHei',
        fontSize: 20,
        color: this.colors.white,
        align: 'center',
        valign: 'middle'
      });
    }

    return slide;
  }

  addContentSlide(title, contentOptions = {}) {
    const slide = this.pptx.addSlide();
    slide.background = { color: this.colors.pageBg };

    this.addStandardHeader(slide, title);

    const card = slide.addShape(this.pptx.ShapeType.rect, {
      x: 0.5, y: 1, w: 12.33, h: 5.8,
      fill: { color: this.colors.cardBg },
      line: { color: this.colors.border }
    });

    this.addStandardFooter(slide);

    return { slide, card };
  }

  addStandardHeader(slide, title) {
    slide.addShape(this.pptx.ShapeType.rect, {
      x: 0, y: 0, w: 13.33, h: 0.8,
      fill: { color: this.colors.titleBar }
    });

    slide.addText(title, {
      x: 0.5, y: 0.15, w: 12.33, h: 0.5,
      fontFace: 'Microsoft YaHei',
      fontSize: 24,
      bold: true,
      color: this.colors.titleText
    });
  }

  addStandardFooter(slide) {
    slide.addText(this.state.requirements.projectName, {
      x: 0.5, y: 7, w: 10, h: 0.4,
      fontFace: 'Microsoft YaHei',
      fontSize: 10,
      color: this.colors.textLight
    });
  }

  addOverviewSlide(title, metrics = []) {
    const { slide } = this.addContentSlide(title);

    let x = 0.8;
    let y = 1.3;
    const cardWidth = 3.8;
    const cardHeight = 2.5;
    const gap = 0.5;

    metrics.forEach((metric, index) => {
      const col = index % 3;
      const row = Math.floor(index / 3);

      const cx = x + col * (cardWidth + gap);
      const cy = y + row * (cardHeight + gap);

      slide.addShape(this.pptx.ShapeType.rect, {
        x: cx, y: cy, w: cardWidth, h: cardHeight,
        fill: { color: this.colors.bgLight },
        line: { color: this.colors.border }
      });

      slide.addText(metric.label, {
        x: cx + 0.3, y: cy + 0.3, w: cardWidth - 0.6, h: 0.6,
        fontFace: 'Microsoft YaHei',
        fontSize: 14,
        color: this.colors.textLight
      });

      slide.addText(metric.value, {
        x: cx + 0.3, y: cy + 0.9, w: cardWidth - 0.6, h: 0.9,
        fontFace: 'Microsoft YaHei',
        fontSize: 32,
        bold: true,
        color: metric.isPositive ? this.colors.success :
               metric.isNegative ? this.colors.danger : this.colors.primary
      });

      if (metric.change) {
        const changeColor = metric.change.startsWith('+') ? this.colors.success : this.colors.danger;
        slide.addText(metric.change, {
          x: cx + 0.3, y: cy + 1.8, w: cardWidth - 0.6, h: 0.5,
          fontFace: 'Microsoft YaHei',
          fontSize: 14,
          color: changeColor
        });
      }
    });

    return slide;
  }

  addChartSlide(title, chartImagePath, insight = '') {
    const { slide } = this.addContentSlide(title);

    if (chartImagePath) {
      slide.addImage({
        path: chartImagePath,
        x: 0.8, y: 1.3, w: 8, h: 5.2
      });
    }

    if (insight) {
      slide.addShape(this.pptx.ShapeType.rect, {
        x: 9, y: 1.3, w: 3.6, h: 5.2,
        fill: { color: this.colors.bgLight },
        line: { color: this.colors.border }
      });

      slide.addText('💡 数据洞察', {
        x: 9.2, y: 1.5, w: 3.2, h: 0.5,
        fontFace: 'Microsoft YaHei',
        fontSize: 14,
        bold: true,
        color: this.colors.primary
      });

      slide.addText(insight, {
        x: 9.2, y: 2.1, w: 3.2, h: 4.2,
        fontFace: 'Microsoft YaHei',
        fontSize: 12,
        color: this.colors.text
      });
    }

    return slide;
  }

  addGridSlide(title, items = [], cols = 3) {
    const { slide } = this.addContentSlide(title);

    const cardWidth = (11.8 - (cols - 1) * 0.4) / cols;
    const cardHeight = 2.5;
    let x = 0.8;
    let y = 1.3;

    items.forEach((item, index) => {
      const col = index % cols;
      const row = Math.floor(index / cols);

      const cx = x + col * (cardWidth + 0.4);
      const cy = y + row * (cardHeight + 0.3);

      slide.addShape(this.pptx.ShapeType.rect, {
        x: cx, y: cy, w: cardWidth, h: cardHeight,
        fill: { color: item.bgColor || this.colors.cardBg },
        line: { color: this.colors.border }
      });

      if (item.title) {
        slide.addText(item.title, {
          x: cx + 0.2, y: cy + 0.2, w: cardWidth - 0.4, h: 0.5,
          fontFace: 'Microsoft YaHei',
          fontSize: 14,
          bold: true,
          color: this.colors.text
        });
      }

      if (item.value) {
        slide.addText(item.value, {
          x: cx + 0.2, y: cy + 0.7, w: cardWidth - 0.4, h: 0.8,
          fontFace: 'Microsoft YaHei',
          fontSize: 24,
          bold: true,
          color: this.colors.primary
        });
      }

      if (item.description) {
        slide.addText(item.description, {
          x: cx + 0.2, y: cy + 1.6, w: cardWidth - 0.4, h: 0.7,
          fontFace: 'Microsoft YaHei',
          fontSize: 10,
          color: this.colors.textLight
        });
      }
    });

    return slide;
  }

  async save(outputPath) {
    await this.pptx.writeFile({ fileName: outputPath });
    return outputPath;
  }
}

module.exports = PPTGenerator;
