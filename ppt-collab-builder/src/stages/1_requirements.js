/**
 * 阶段1: 需求分析
 * 通过多轮对话收集用户需求，生成初步内容大纲
 */

const readline = require('readline');

class RequirementsStage {
  constructor(state) {
    this.state = state;
    this.rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout
    });
  }

  async run() {
    console.log('\n' + '='.repeat(80));
    console.log('阶段1: 需求分析');
    console.log('='.repeat(80) + '\n');

    await this.askProjectName();
    await this.askProjectPurpose();
    await this.askTargetAudience();
    await this.askInitialStructure();
    await this.confirmPreliminaryOutline();

    this.state.setStage(2);
    console.log('\n✅ 阶段1完成！进入阶段2: 大纲完善\n');
  }

  ask(question) {
    return new Promise((resolve) => {
      this.rl.question(question + ' ', (answer) => {
        resolve(answer.trim());
      });
    });
  }

  async askProjectName() {
    const name = await this.ask('📋 请输入项目名称（例如：差旅用品经营分析）:');
    if (!name) {
      console.log('⚠️ 项目名称不能为空，请重新输入');
      return this.askProjectName();
    }
    this.state.setRequirements({ projectName: name });
    console.log(`✅ 项目名称已设置: ${name}\n`);
  }

  async askProjectPurpose() {
    const purpose = await this.ask('🎯 这个PPT的主要目的是什么？（例如：向客户展示经营数据、内部汇报、竞品分析等）:');
    if (!purpose) {
      console.log('⚠️ 请描述一下PPT的目的');
      return this.askProjectPurpose();
    }
    this.state.setRequirements({ projectPurpose: purpose });
    console.log(`✅ 项目目的已记录\n`);
  }

  async askTargetAudience() {
    const audience = await this.ask('👥 这个PPT的目标受众是谁？（例如：客户、管理层、销售团队等）:');
    if (!audience) {
      console.log('⚠️ 请说明一下目标受众');
      return this.askTargetAudience();
    }
    this.state.setRequirements({ targetAudience: audience });
    console.log(`✅ 目标受众已记录\n`);
  }

  async askInitialStructure() {
    console.log('\n📑 现在让我们规划PPT的页面结构。');
    console.log('请依次输入每个页面的标题，输入空行结束。');
    console.log('示例：');
    console.log('  1. 生意概览');
    console.log('  2. 单店效能');
    console.log('  3. ...\n');

    const pages = [];
    let index = 1;

    while (true) {
      const pageTitle = await this.ask(`第${index}页标题（留空结束）:`);
      if (!pageTitle) {
        if (pages.length === 0) {
          console.log('⚠️ 至少需要一个页面，请重新输入');
          continue;
        }
        break;
      }
      pages.push({
        id: index,
        title: pageTitle,
        purpose: await this.ask(`  🔹 这一页的作用是什么？:`)
      });
      index++;
    }

    this.state.setRequirements({ initialOutline: pages });
    console.log(`\n✅ 已规划 ${pages.length} 个页面\n`);
  }

  async confirmPreliminaryOutline() {
    console.log('\n' + '-'.repeat(80));
    console.log('📋 初步内容大纲确认:');
    console.log('-'.repeat(80));
    console.log(`项目名称: ${this.state.requirements.projectName}`);
    console.log(`项目目的: ${this.state.requirements.projectPurpose}`);
    console.log(`目标受众: ${this.state.requirements.targetAudience}`);
    console.log('\n页面结构:');
    this.state.requirements.initialOutline.forEach((page) => {
      console.log(`  ${page.id}. ${page.title}`);
      console.log(`     作用: ${page.purpose}`);
    });
    console.log('-'.repeat(80) + '\n');

    const confirm = await this.ask('✅ 这个初步大纲是否正确？(y/n):');
    if (confirm.toLowerCase() !== 'y' && confirm.toLowerCase() !== 'yes') {
      console.log('\n🔄 让我们重新开始...\n');
      return this.askInitialStructure();
    }
  }

  close() {
    this.rl.close();
  }
}

module.exports = RequirementsStage;
