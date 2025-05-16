// 完整的Markdown解析器测试工具

import fs from 'fs';
import { Md2Json } from './js/md2json.js';

// 读取测试文件
try {
  const markdownText = fs.readFileSync('./test.md', 'utf8');

  // 初始化解析器
  const parser = new Md2Json();

  // 解析整个文档
  console.log('开始解析整个Markdown文档...');
  const document = parser.convert(markdownText);

  // 输出文档结构概览
  console.log('\n文档结构概览:');
  console.log(`总计 ${document.children.length} 个顶级节点`);

  // 按类型统计节点
  const nodeTypes = {};
  document.children.forEach(node => {
    if (!nodeTypes[node.type]) {
      nodeTypes[node.type] = 0;
    }
    nodeTypes[node.type]++;
  });

  console.log('\n节点类型统计:');
  Object.keys(nodeTypes).forEach(type => {
    console.log(`- ${type}: ${nodeTypes[type]}个`);
  });

  // 输出前5个节点的详细信息作为示例
  console.log('\n前5个节点示例:');
  for (let i = 0; i < Math.min(5, document.children.length); i++) {
    console.log(`\n节点 ${i+1} (类型: ${document.children[i].type}):`);
    console.log(JSON.stringify(document.children[i], null, 2));
  }

  // 保存完整解析结果到文件
  fs.writeFileSync('./parsed-output.json', JSON.stringify(document, null, 2));
  console.log('\n完整解析结果已保存到 parsed-output.json');

  // 测试特定格式的解析
  console.log('\n测试特定格式解析:');

  // 1. 测试加粗斜体组合
  console.log('\n1. 测试加粗斜体组合:');
  const boldItalicTest = parser.parseInlineStyles('这是***加粗斜体***文本');
  console.log(JSON.stringify(boldItalicTest, null, 2));

  // 2. 测试嵌套样式
  console.log('\n2. 测试嵌套样式:');
  const nestedStyleTest = parser.parseInlineStyles('这是**加粗的*斜体*文本**');
  console.log(JSON.stringify(nestedStyleTest, null, 2));

  // 3. 测试复杂组合
  console.log('\n3. 测试复杂组合:');
  const complexTest = parser.parseInlineStyles('这是**加粗**和*斜体*还有`代码`以及~~删除线~~的组合');
  console.log(JSON.stringify(complexTest, null, 2));

} catch (error) {
  console.error('解析过程中出错:', error);
}
