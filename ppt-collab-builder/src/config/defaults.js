/**
 * 默认配置
 */

module.exports = {
  // PPT默认配置
  ppt: {
    width: 13.33,
    height: 7.5,
    margin: 0.5,
    titleBarHeight: 0.8
  },

  // 默认主题
  theme: {
    colors: {
      titleBar: '165DFF',
      titleText: 'FFFFFF',
      pageBg: 'F5F7FA',
      cardBg: 'FFFFFF',
      primary: '1A73E8',
      success: '00B42A',
      danger: 'F53F3F',
      text: '1D2129',
      textLight: '4E5969',
      bgLight: 'E8F3FF',
      bgLightGreen: 'E8FFEA',
      border: 'E5E6EB',
      headerBg: 'F2F3F5',
      white: 'FFFFFF'
    },
    fonts: {
      title: 'Microsoft YaHei',
      body: 'Microsoft YaHei'
    },
    fontSize: {
      title: 28,
      subtitle: 16,
      body: 14,
      small: 10
    }
  },

  // 图表类型
  chartTypes: [
    { id: 'line', name: '折线图', description: '趋势展示' },
    { id: 'bar', name: '柱状图', description: '对比展示' },
    { id: 'pie', name: '饼图/环形图', description: '占比展示' },
    { id: 'combo', name: '组合图', description: '双轴展示' }
  ],

  // 页面模板
  pageTemplates: [
    { id: 'overview', name: '生意概览', hasChart: true, hasTable: false },
    { id: 'efficiency', name: '单店效能', hasChart: true, hasTable: false },
    { id: 'structure', name: '单品结构', hasChart: true, hasTable: true },
    { id: 'essential', name: '必备品全景', hasChart: false, hasTable: false, hasImages: true },
    { id: 'super_3d', name: '超级单品三维论证', hasChart: true, hasTable: false },
    { id: 'super_performance', name: '超级单品表现', hasChart: true, hasTable: false },
    { id: 'targets', name: '必备品目标', hasChart: false, hasTable: true },
    { id: 'growth', name: '增长路径', hasChart: false, hasTable: false },
    { id: 'investment', name: '品牌配套投入', hasChart: false, hasTable: false }
  ]
};
