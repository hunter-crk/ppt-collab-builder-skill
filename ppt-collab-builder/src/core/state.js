/**
 * 项目状态管理
 */

class ProjectState {
  constructor() {
    // 阶段1: 需求分析
    this.requirements = {
      projectName: '',
      projectPurpose: '',
      targetAudience: '',
      initialOutline: []
    };

    // 阶段2: 大纲完善
    this.detailedOutline = {
      theme: 'business',
      visualSpec: {
        colors: null,
        fonts: null,
        layout: null
      },
      pages: []
    };

    // 阶段3: 图表规划
    this.chartPlan = [];

    // 阶段4: 资源准备
    this.resources = {
      dataTemplates: [],
      imageFolders: []
    };

    // 阶段5: 数据洞察
    this.enhancedOutline = null;

    // 当前阶段
    this.currentStage = 1;

    // 项目文件路径
    this.paths = {
      dataDir: null,
      imageDir: null,
      outputDir: null
    };
  }

  setRequirements(requirements) {
    this.requirements = { ...this.requirements, ...requirements };
  }

  setDetailedOutline(outline) {
    this.detailedOutline = { ...this.detailedOutline, ...outline };
  }

  setChartPlan(chartPlan) {
    this.chartPlan = chartPlan;
  }

  setResources(resources) {
    this.resources = { ...this.resources, ...resources };
  }

  setEnhancedOutline(outline) {
    this.enhancedOutline = outline;
  }

  setStage(stage) {
    this.currentStage = stage;
  }

  setPaths(paths) {
    this.paths = { ...this.paths, ...paths };
  }

  toJSON() {
    return {
      requirements: this.requirements,
      detailedOutline: this.detailedOutline,
      chartPlan: this.chartPlan,
      resources: this.resources,
      enhancedOutline: this.enhancedOutline,
      currentStage: this.currentStage,
      paths: this.paths
    };
  }

  fromJSON(json) {
    this.requirements = json.requirements || this.requirements;
    this.detailedOutline = json.detailedOutline || this.detailedOutline;
    this.chartPlan = json.chartPlan || this.chartPlan;
    this.resources = json.resources || this.resources;
    this.enhancedOutline = json.enhancedOutline || this.enhancedOutline;
    this.currentStage = json.currentStage || this.currentStage;
    this.paths = json.paths || this.paths;
  }
}

module.exports = ProjectState;
