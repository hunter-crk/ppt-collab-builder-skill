# pptxgenjs 表格渲染经验说明

## 问题背景

在使用 pptxgenjs 生成 PPT 表格时，直接设置 `rowsFill` 或在单元格对象顶层设置 `fill` 属性都无法正确渲染背景色。经过多次测试验证，找到了正确的 API 调用方式。

## 正确 API 格式

### 单元格数据结构

```javascript
// 错误方式（不生效）
{ text: 'Header', bold: true, fill: 'EEEEEE' }

// 正确方式
{
  text: 'Header',
  options: {
    bold: true,
    fill: 'EEEEEE'  // 注意：不需要 # 号前缀
  }
}
```

### 完整表格示例

```javascript
const pptxgen = require('pptxgenjs');
const ppt = new pptxgen();
ppt.defineLayout({ name: 'WIDE', width: 13.33, height: 7.5 });
ppt.layout = 'WIDE';

const slide = ppt.addSlide();

const rows = [
  [
    { text: 'Header 1', options: { bold: true, fill: 'EEEEEE' } },
    { text: 'Header 2', options: { bold: true, fill: 'EEEEEE' } },
    { text: 'Header 3', options: { bold: true, fill: 'EEEEEE' } }
  ],
  [
    { text: 'Data 1', options: { fill: 'E8F5E9' } },
    { text: 'Data 2', options: { fill: 'E8F5E9' } },
    'Data 3'  // 纯字符串也可以，无样式
  ],
  [
    'Data 4',
    'Data 5',
    'Data 6'
  ]
];

slide.addTable(rows, {
  x: 0.5,
  y: 0.5,
  w: 4,
  h: 2,
  fontSize: 14,
  border: { color: 'DDDDDD' }
});

ppt.writeFile({ fileName: 'test.pptx' });
```

## 辅助函数封装

为了方便使用，可以封装一个 `cell()` 辅助函数：

```javascript
function cell(text, opts = {}) {
  const result = { text: String(text) };
  if (Object.keys(opts).length > 0) {
    result.options = {};
    if (opts.bold !== undefined) result.options.bold = opts.bold;
    if (opts.fill && opts.fill.color) {
      // 自动去掉颜色中的 # 号
      result.options.fill = opts.fill.color.replace('#', '');
    }
  }
  return result;
}

// 使用示例
const tableData = [
  [
    cell('省份', { bold: true, fill: { color: '#EEEEEE' } }),
    cell('占比', { bold: true, fill: { color: '#EEEEEE' } })
  ],
  [
    cell('上海', { fill: { color: '#E8F5E9' } }),
    cell('12.5%', { fill: { color: '#E8F5E9' } })
  ]
];
```

## 关键要点

1. **必须使用 `options` 嵌套对象**：样式属性（`bold`、`fill` 等）不能直接放在单元格对象顶层，必须放在 `options` 子对象中。

2. **`fill` 颜色不带 `#` 前缀**：只需要十六进制字符串，如 `'EEEEEE'`，不需要 `'#EEEEEE'`。

3. **可以混合使用**：同一表格中可以混用 `{text, options}` 对象和纯字符串，纯字符串单元格无样式。

4. **表格级别配置**：`slide.addTable()` 的第二个参数可以设置全局样式（`fontSize`、`border` 等），会应用到所有没有单独设置样式的单元格。

## 配色规范（本项目）

| 用途 | 颜色值 | 说明 |
|------|--------|------|
| 表头背景 | `EEEEEE` | 浅灰色 |
| 重点高亮行 | `E8F5E9` | 浅绿色 |
| 主色（标题栏） | `1A73E8` | 蓝色 |
| 辅助色（增长） | `34A853` | 绿色 |
| 边框 | `DDDDDD` | 浅灰色 |

## 测试验证文件

项目中包含测试验证文件：
- `scripts/test-table-official.js` - 官方文档正确用法测试

运行测试：
```bash
cd scripts
node test-table-official.js
```

## 常见错误

### ❌ 错误1：样式放在顶层
```javascript
// 不生效！
{ text: 'Header', bold: true, fill: 'EEEEEE' }
```

### ❌ 错误2：fill 带 # 号
```javascript
// 可能不生效！
{ text: 'Header', options: { fill: '#EEEEEE' } }
```

### ❌ 错误3：使用 rowsFill
```javascript
// 不生效！
slide.addTable(rows, { rowsFill: ['EEEEEE', null] });
```

### ✅ 正确方式
```javascript
{ text: 'Header', options: { bold: true, fill: 'EEEEEE' } }
```

## 参考文档

- pptxgenjs 官方文档：https://gitbrent.github.io/PptxGenJS/
- 表格 API 文档：https://gitbrent.github.io/PptxGenJS/docs/api-tables.html
