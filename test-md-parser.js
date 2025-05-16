// 测试Markdown解析器

// 导入解析器
import { Md2Json } from './js/md2json.js';

// 创建测试用例
const testCases = [
  "***这是加粗斜体文本***",
  "**这是加粗文本**",
  "*这是斜体文本*",
  "这是**加粗**和*斜体*混合文本",
  "这是***加粗斜体***混合文本",
  "~~这是删除线文本~~"
];

// 初始化解析器
const parser = new Md2Json();

// 运行测试
testCases.forEach(test => {
  console.log(`\n测试输入: "${test}"`);

  // 创建一个临时段落对象
  const paragraph = {
    type: 'paragraph',
    text: test
  };

  // 解析样式
  const parsedStyles = parser.parseInlineStyles(test);

  // 输出解析结果
  console.log("解析结果:");
  console.log(JSON.stringify({
    fullContent: parsedStyles.fullContent,
    styles: parsedStyles.styles
  }, null, 2));
});

// 直接测试parseTextStyle函数
console.log("\n直接测试parseTextStyle函数:");
[
  "***这是加粗斜体文本***",
  "**这是加粗文本**",
  "*这是斜体文本*"
].forEach(test => {
  console.log(`\n输入: "${test}"`);
  console.log("结果:", JSON.stringify(parser.parseTextStyle(test), null, 2));
});
