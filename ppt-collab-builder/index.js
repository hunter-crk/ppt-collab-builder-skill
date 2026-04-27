#!/usr/bin/env node

/**
 * PPT协作构建器
 * 6阶段协作式数据型PPT生成工具
 */

const ProjectState = require('./src/core/state');
const RequirementsStage = require('./src/stages/1_requirements');
const OutlineStage = require('./src/stages/2_outline');
const ChartsStage = require('./src/stages/3_charts');
const ResourcesStage = require('./src/stages/4_resources');
const InsightsStage = require('./src/stages/5_insights');
const GenerateStage = require('./src/stages/6_generate');

console.log(`
╔═══════════════════════════════════════════════════════════════╗
║                    PPT 协作构建器                               ║
║            6阶段协作式数据型PPT生成工具                         ║
╚═══════════════════════════════════════════════════════════════╝
`);

async function main() {
  const state = new ProjectState();

  const stages = [
    RequirementsStage,
    OutlineStage,
    ChartsStage,
    ResourcesStage,
    InsightsStage,
    GenerateStage
  ];

  const stageInstances = [];

  try {
    for (let i = 0; i < stages.length; i++) {
      const StageClass = stages[i];
      const stage = new StageClass(state);
      stageInstances.push(stage);

      if (state.currentStage > i + 1) {
        console.log(`⏭️  跳过阶段${i + 1} (已完成)\n`);
        continue;
      }

      await stage.run();
      stage.close();
    }

    console.log(`
╔═══════════════════════════════════════════════════════════════╗
║                        🎉 项目完成！                             ║
╚═══════════════════════════════════════════════════════════════╝
`);

  } catch (error) {
    console.error('\n❌ 发生错误:', error.message);
    console.error(error.stack);
  } finally {
    stageInstances.forEach(stage => {
      try { stage.close(); } catch (e) {}
    });
  }
}

main();
