/**
 * ECharts图表渲染模块
 * 使用Playwright在浏览器中渲染高清图表
 */

const path = require('path');
const fs = require('fs');

class ChartRenderer {
  constructor(outputDir) {
    this.outputDir = outputDir;
    this.browser = null;
    this.page = null;
  }

  async init() {
    try {
      const { chromium } = require('playwright');
      this.browser = await chromium.launch({ headless: true });
      this.page = await this.browser.newPage({
        viewport: { width: 1600, height: 1000 }
      });
    } catch (e) {
      console.log('⚠️  Playwright未安装，图表渲染功能将不可用');
      console.log('   如需使用图表渲染，请运行: npm install playwright && npx playwright install chromium');
    }
  }

  async close() {
    if (this.browser) {
      await this.browser.close();
    }
  }

  getChartHTML(chartOptions) {
    const {
      type = 'line',
      title = '',
      xAxis = 'X轴',
      yAxis = 'Y轴',
      series = [],
      data = [],
      theme = 'business'
    } = chartOptions;

    const colors = this.getThemeColors(theme);
    const chartColors = colors.chart || ['#165DFF', '#00B42A', '#FF7D00', '#F53F3F', '#722ED1'];

    const echartSeries = series.map((s, i) => ({
      name: s,
      type: type,
      data: data.map(d => d.values[i]),
      smooth: true,
      itemStyle: { color: chartColors[i % chartColors.length] },
      areaStyle: type === 'line' ? {
        color: {
          type: 'linear',
          x: 0, y: 0, x2: 0, y2: 1,
          colorStops: [
            { offset: 0, color: chartColors[i % chartColors.length] + '40' },
            { offset: 1, color: chartColors[i % chartColors.length] + '05' }
          ]
        }
      } : undefined
    }));

    const echartsOption = {
      title: {
        text: title,
        left: 'center',
        top: 10,
        textStyle: {
          fontSize: 20,
          fontFamily: 'Microsoft YaHei',
          fontWeight: 'bold',
          color: '#1D2129'
        }
      },
      tooltip: {
        trigger: 'axis',
        backgroundColor: 'rgba(255,255,255,0.95)',
        borderColor: '#E5E6EB',
        textStyle: { color: '#1D2129', fontSize: 14 }
      },
      legend: {
        top: 50,
        textStyle: { fontSize: 14, fontFamily: 'Microsoft YaHei' }
      },
      grid: {
        left: '3%',
        right: '4%',
        bottom: '3%',
        top: 100,
        containLabel: true
      },
      xAxis: {
        type: 'category',
        boundaryGap: type !== 'line',
        data: data.map(d => d.x),
        axisLabel: { fontSize: 14, fontFamily: 'Microsoft YaHei' },
        axisLine: { lineStyle: { color: '#E5E6EB' } },
        name: xAxis,
        nameTextStyle: { fontSize: 14, fontFamily: 'Microsoft YaHei' }
      },
      yAxis: {
        type: 'value',
        axisLabel: { fontSize: 14, fontFamily: 'Microsoft YaHei' },
        axisLine: { lineStyle: { color: '#E5E6EB' } },
        splitLine: { lineStyle: { color: '#F2F3F5' } },
        name: yAxis,
        nameTextStyle: { fontSize: 14, fontFamily: 'Microsoft YaHei' }
      },
      series: echartSeries
    };

    return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <script src="https://cdn.jsdelivr.net/npm/echarts@5.4.3/dist/echarts.min.js"></script>
  <style>
    body { margin: 0; padding: 0; background: #fff; }
    #chart { width: 1600px; height: 1000px; }
  </style>
</head>
<body>
  <div id="chart"></div>
  <script>
    const chart = echarts.init(document.getElementById('chart'));
    chart.setOption(${JSON.stringify(echartsOption)});
  </script>
</body>
</html>`;
  }

  getThemeColors(theme) {
    const themes = {
      business: {
        chart: ['#165DFF', '#00B42A', '#FF7D00', '#F53F3F', '#722ED1']
      },
      simple: {
        chart: ['#1D2129', '#4E5969', '#86909C', '#C9CDD4', '#F2F3F5']
      },
      tech: {
        chart: ['#00B42A', '#0FC6C2', '#165DFF', '#722ED1', '#FF7D00']
      }
    };
    return themes[theme] || themes.business;
  }

  async renderChart(chartOptions, fileName) {
    if (!this.page) {
      console.log('⚠️  图表渲染不可用，跳过: ' + fileName);
      return null;
    }

    const html = this.getChartHTML(chartOptions);
    const outputPath = path.join(this.outputDir, fileName);

    await this.page.setContent(html, { waitUntil: 'networkidle' });
    await this.page.waitForTimeout(500);

    const chartElement = await this.page.$('#chart');
    await chartElement.screenshot({
      path: outputPath,
      scale: 2
    });

    return outputPath;
  }

  async renderBarChart(options, fileName) {
    return this.renderChart({ ...options, type: 'bar' }, fileName);
  }

  async renderLineChart(options, fileName) {
    return this.renderChart({ ...options, type: 'line' }, fileName);
  }

  async renderPieChart(options, fileName) {
    const { title = '', data = [], theme = 'business' } = options;
    const colors = this.getThemeColors(theme);

    const echartsOption = {
      title: {
        text: title,
        left: 'center',
        top: 10,
        textStyle: {
          fontSize: 20,
          fontFamily: 'Microsoft YaHei',
          fontWeight: 'bold',
          color: '#1D2129'
        }
      },
      tooltip: {
        trigger: 'item',
        backgroundColor: 'rgba(255,255,255,0.95)',
        borderColor: '#E5E6EB',
        textStyle: { color: '#1D2129', fontSize: 14 },
        formatter: '{b}: {c} ({d}%)'
      },
      legend: {
        orient: 'vertical',
        right: '5%',
        top: 'center',
        textStyle: { fontSize: 14, fontFamily: 'Microsoft YaHei' }
      },
      series: [{
        type: 'pie',
        radius: ['35%', '55%'],
        center: ['35%', '50%'],
        data: data.map((d, i) => ({
          name: d.name,
          value: d.value,
          itemStyle: { color: colors.chart[i % colors.chart.length] }
        })),
        label: {
          show: true,
          fontSize: 12,
          fontFamily: 'Microsoft YaHei'
        },
        emphasis: {
          itemStyle: {
            shadowBlur: 10,
            shadowOffsetX: 0,
            shadowColor: 'rgba(0, 0, 0, 0.5)'
          }
        }
      }]
    };

    const html = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <script src="https://cdn.jsdelivr.net/npm/echarts@5.4.3/dist/echarts.min.js"></script>
  <style>
    body { margin: 0; padding: 0; background: #fff; }
    #chart { width: 1600px; height: 1000px; }
  </style>
</head>
<body>
  <div id="chart"></div>
  <script>
    const chart = echarts.init(document.getElementById('chart'));
    chart.setOption(${JSON.stringify(echartsOption)});
  </script>
</body>
</html>`;

    if (!this.page) {
      console.log('⚠️  图表渲染不可用，跳过: ' + fileName);
      return null;
    }

    const outputPath = path.join(this.outputDir, fileName);
    await this.page.setContent(html, { waitUntil: 'networkidle' });
    await this.page.waitForTimeout(500);

    const chartElement = await this.page.$('#chart');
    await chartElement.screenshot({ path: outputPath, scale: 2 });

    return outputPath;
  }
}

module.exports = ChartRenderer;
