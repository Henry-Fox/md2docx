import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  LevelFormat,
  AlignmentType,
  HeadingLevel,
  SectionType,
  LineRuleType,
  BorderStyle,
  ExternalHyperlink,
  Table,
  TableRow,
  TableCell,
  WidthType,
  CheckBox,
  ImageRun,
  TextWrappingType,
  TextWrappingSide,
  HorizontalPositionRelativeFrom,
  VerticalPositionRelativeFrom,
  HorizontalPositionAlign,
  VerticalPositionAlign,
} from "docx";
import { saveAs } from "file-saver";

/**
 * @async
 * @function loadImage
 * @description 加载图片并转换为ArrayBuffer
 * @param {string} url - 图片URL
 * @returns {Promise<{buffer: ArrayBuffer, width: number, height: number}>}
 */
async function loadImage(url) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.crossOrigin = "anonymous"; // 处理跨域问题

    img.onload = async () => {
      try {
        // 创建canvas来获取图片数据
        const canvas = document.createElement("canvas");
        canvas.width = img.naturalWidth;
        canvas.height = img.naturalHeight;

        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0);

        // 将canvas转换为blob
        const blob = await new Promise((resolve) => canvas.toBlob(resolve));

        // 将blob转换为ArrayBuffer
        const buffer = await blob.arrayBuffer();

        resolve({
          buffer,
          width: img.naturalWidth,
          height: img.naturalHeight,
        });
      } catch (error) {
        reject(error);
      }
    };

    img.onerror = () => {
      reject(new Error(`Failed to load image: ${url}`));
    };

    img.src = url;
  });
}

/**
 * @async
 * @function processImages
 * @description 处理文档中的所有图片
 * @param {Array} json - 文档JSON数据
 * @returns {Promise<Array>} 处理后的图片信息数组
 */
async function processImages(json) {
  const imageInfos = [];

  // 遍历JSON数据查找图片
  for (const child of json.children) {
    if (child.type === "image") {
      try {
        const imageData = await loadImage(child.url);
        imageInfos.push({
          ...child,
          ...imageData,
        });
      } catch (error) {
        console.error(`Error processing image: ${child.url}`, error);
        // 如果图片加载失败，使用占位符
        imageInfos.push({
          ...child,
          buffer: null,
          width: 400,
          height: 300,
        });
      }
    }
  }

  return imageInfos;
}

// Documents contain sections, you can have multiple sections per document, go here to learn more about sections
// This simple example will only contain one section

// 浏览器环境中使用Packer.toBlob和saveAs
export default async function runTest() {
  console.log("开始生成文档...");

  // 创建一个数组用于收集所有段落
  const paragraphs = [];

  //设置一个json
  const json = {
    type: "document",
    children: [
      {
        type: "quote",
        rawText: "> 这是一段引用文本",
        fullContent: "这是一段引用文本",
        inlineStyles: [
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "这是一段引用文本",
          },
        ],
      },
      {
        type: "paragraph",
        rawText: "这是一个普通段落，用于分隔引用",
        hasNumber: false,
        number: "",
        fullContent: "这是一个普通段落，用于分隔引用",
        inlineStyles: [
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "这是一个普通段落，用于分隔引用",
          },
        ],
      },
      {
        type: "quote",
        rawText: "> 这是另一段引用文本，包含**粗体**和*斜体*",
        fullContent: "这是另一段引用文本，包含粗体和斜体",
        inlineStyles: [
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "这是另一段引用文本，包含",
          },
          {
            bold: true,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "粗体",
          },
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "和",
          },
          {
            bold: false,
            italics: true,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "斜体",
          },
        ],
      },
      {
        type: "list",
        rawText: "1. 这是第一个有序列表项",
        hasNumber: true,
        level: 0,
        fullContent: "这是第一个有序列表项",
        inlineStyles: [
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "这是第一个有序列表项",
          },
        ],
      },
      {
        type: "list",
        rawText: "2. 这是第二个有序列表项",
        level: 0,
        hasNumber: true,
        fullContent: "这是第二个有序列表项",
        inlineStyles: [
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "这是第二个有序列表项",
          },
        ],
      },
      {
        type: "list",
        rawText: "3. 这是第三个有序列表项",
        level: 0,
        hasNumber: true,
        fullContent: "这是第三个有序列表项",
        inlineStyles: [
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "这是第三个有序列表项",
          },
        ],
      },
      {
        type: "paragraph",
        rawText: "这是一个普通段落，用于分隔列表",
        hasNumber: false,
        number: "",
        fullContent: "这是一个普通段落，用于分隔列表",
        inlineStyles: [
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "这是一个普通段落，用于分隔列表",
          },
        ],
      },
      {
        type: "list",
        rawText: "- 这是第一个无序列表项",
        level: 0,
        hasNumber: false,
        fullContent: "这是第一个无序列表项",
        inlineStyles: [
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "这是第一个无序列表项",
          },
        ],
      },
      {
        type: "list",
        rawText: "- 这是第二个无序列表项",
        level: 0,
        hasNumber: false,
        fullContent: "这是第二个无序列表项",
        inlineStyles: [
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "这是第二个无序列表项",
          },
        ],
      },
      {
        type: "list",
        rawText: "- 这是第三个无序列表项",
        level: 0,
        hasNumber: false,
        fullContent: "这是第三个无序列表项",
        inlineStyles: [
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "这是第三个无序列表项",
          },
        ],
      },
      {
        type: "heading",
        rawText: "## 代码块测试",
        level: 2,
        hasNumber: true,
        number: "1",
        fullContent: "代码块测试",
        inlineStyles: [
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "代码块测试",
          },
        ],
      },
      {
        type: "paragraph",
        rawText: "下面是一个简单的JavaScript代码块：",
        hasNumber: false,
        number: "",
        fullContent: "下面是一个简单的JavaScript代码块：",
        inlineStyles: [
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "下面是一个简单的JavaScript代码块：",
          },
        ],
      },
      {
        type: "code",
        rawText:
          "```javascript\nfunction hello() {\n    console.log('Hello, World!');\n}\n```",
        fullContent: "function hello() {\n    console.log('Hello, World!');\n}",
        inlineStyles: [
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "function hello() {\n    console.log('Hello, World!');\n}",
          },
        ],
      },
      {
        type: "paragraph",
        rawText: "下面是一个HTML代码块：",
        hasNumber: false,
        number: "",
        fullContent: "下面是一个HTML代码块：",
        inlineStyles: [
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "下面是一个HTML代码块：",
          },
        ],
      },
      {
        type: "code",
        rawText:
          '```html\n<div class="container">\n    <h1>标题</h1>\n    <p>段落内容</p>\n</div>\n```',
        fullContent:
          '<div class="container">\n    <h1>标题</h1>\n    <p>段落内容</p>\n</div>',
        inlineStyles: [
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content:
              '<div class="container">\n    <h1>标题</h1>\n    <p>段落内容</p>\n</div>',
          },
        ],
      },
      {
        type: "paragraph",
        rawText: "下面是一个CSS代码块：",
        hasNumber: false,
        number: "",
        fullContent: "下面是一个CSS代码块：",
        inlineStyles: [
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "下面是一个CSS代码块：",
          },
        ],
      },
      {
        type: "code",
        rawText:
          "```css\n.container {\n    width: 100%;\n    max-width: 1200px;\n    margin: 0 auto;\n    padding: 20px;\n}\n```",
        fullContent:
          ".container {\n    width: 100%;\n    max-width: 1200px;\n    margin: 0 auto;\n    padding: 20px;\n}",
        inlineStyles: [
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content:
              ".container {\n    width: 100%;\n    max-width: 1200px;\n    margin: 0 auto;\n    padding: 20px;\n}",
          },
        ],
      },
      {
        type: "paragraph",
        rawText: "下面是一些超链接示例：",
        hasNumber: false,
        number: "",
        fullContent: "下面是一些超链接示例：",
        inlineStyles: [
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "下面是一些超链接示例：",
          },
        ],
      },
      {
        type: "hyperlink",
        text: "这是一个链接",
        url: "https://www.example.com",
        title: "",
        rawText: "[这是一个链接](https://www.example.com)",
      },
      {
        type: "hyperlink",
        text: "带有标题的链接",
        url: "https://www.example.com",
        title: "链接标题",
        rawText: '[带有标题的链接](https://www.example.com "链接标题")',
      },
      {
        type: "hyperlink",
        text: "https://www.example.com",
        url: "https://www.example.com",
        title: "",
        rawText: "<https://www.example.com> (自动链接)",
      },
      {
        type: "heading",
        rawText: "### 人员信息表格",
        level: 3,
        hasNumber: false,
        number: "",
        fullContent: "人员信息表格",
        inlineStyles: [
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "人员信息表格",
          },
        ],
      },
      {
        type: "table",
        headers: [
          {
            rawText: "姓名",
            fullContent: "姓名",
          },
          {
            rawText: "年龄",
            fullContent: "年龄",
          },
          {
            rawText: "职业",
            fullContent: "职业",
          },
        ],
        alignments: ["left", "left", "left"],
        rows: [
          [
            {
              rawText: "张三",
              fullContent: "张三",
            },
            {
              rawText: "25",
              fullContent: "25",
            },
            {
              rawText: "工程师",
              fullContent: "工程师",
            },
          ],
          [
            {
              rawText: "李四",
              fullContent: "李四",
            },
            {
              rawText: "30",
              fullContent: "30",
            },
            {
              rawText: "设计师",
              fullContent: "设计师",
            },
          ],
          [
            {
              rawText: "王五",
              fullContent: "王五",
            },
            {
              rawText: "28",
              fullContent: "28",
            },
            {
              rawText: "产品经理",
              fullContent: "产品经理",
            },
          ],
        ],
      },
      {
        type: "heading",
        rawText: "### 对齐方式表格",
        level: 3,
        hasNumber: false,
        number: "",
        fullContent: "对齐方式表格",
        inlineStyles: [
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "对齐方式表格",
          },
        ],
      },
      {
        type: "table",
        headers: [
          {
            rawText: "左对齐",
            fullContent: "左对齐",
          },
          {
            rawText: "居中对齐",
            fullContent: "居中对齐",
          },
          {
            rawText: "右对齐",
            fullContent: "右对齐",
          },
        ],
        alignments: ["left", "center", "right"],
        rows: [
          [
            {
              rawText: "内容",
              fullContent: "内容",
            },
            {
              rawText: "内容",
              fullContent: "内容",
            },
            {
              rawText: "内容",
              fullContent: "内容",
            },
          ],
          [
            {
              rawText: "文本",
              fullContent: "文本",
            },
            {
              rawText: "文本",
              fullContent: "文本",
            },
            {
              rawText: "文本",
              fullContent: "文本",
            },
          ],
        ],
      },
      {
        type: "horizontal_rule",
        rawText: "---",
      },
      {
        type: "heading",
        rawText: "## 任务列表测试",
        level: 2,
        hasNumber: false,
        number: "",
        fullContent: "任务列表测试",
        inlineStyles: [
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "任务列表测试",
          },
        ],
      },
      {
        type: "task",
        rawText: "- [ ] 未完成任务1",
        isChecked: false,
        fullContent: "未完成任务1",
        inlineStyles: [
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "未完成任务1",
          },
        ],
      },
      {
        type: "task",
        rawText: "- [x] 已完成任务1",
        isChecked: true,
        fullContent: "已完成任务1",
        inlineStyles: [
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "已完成任务1",
          },
        ],
      },
      {
        type: "task",
        rawText: "  - [ ] 嵌套的未完成任务",
        isChecked: false,
        fullContent: "嵌套的未完成任务",
        inlineStyles: [
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "嵌套的未完成任务",
          },
        ],
      },
      {
        type: "task",
        rawText: "  - [x] 嵌套的已完成任务",
        isChecked: true,
        fullContent: "嵌套的已完成任务",
        inlineStyles: [
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "嵌套的已完成任务",
          },
        ],
      },
      {
        type: "paragraph",
        rawText: "这是一个普通段落，用于分隔任务列表",
        hasNumber: false,
        number: "",
        fullContent: "这是一个普通段落，用于分隔任务列表",
        inlineStyles: [
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "这是一个普通段落，用于分隔任务列表",
          },
        ],
      },
      {
        type: "task",
        rawText: "- [ ] 带**粗体**的任务",
        isChecked: false,
        fullContent: "带粗体的任务",
        inlineStyles: [
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "带",
          },
          {
            bold: true,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "粗体",
          },
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "的任务",
          },
        ],
      },
      {
        type: "task",
        rawText: "- [x] 带*斜体*的任务",
        isChecked: true,
        fullContent: "带斜体的任务",
        inlineStyles: [
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "带",
          },
          {
            bold: false,
            italics: true,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "斜体",
          },
          {
            bold: false,
            italics: false,
            strike: false,
            code: false,
            underline: false,
            superScript: false,
            subScript: false,
            content: "的任务",
          },
        ],
      },
      {
        type: "image",
        rawText:
          '[![示例图片](./img/test.jpg "示例图片标题")](https://www.baidu.com "点击链接")',
        alt: "示例图片",
        url: "./img/test.jpg",
        title: "示例图片标题",
        clickUrl: "https://www.baidu.com",
        clickTitle: "点击链接",
      },
    ],
  };

  // 处理所有图片
  console.log("开始处理图片...");
  const imageInfos = await processImages(json);
  console.log(`图片处理完成，共处理 ${imageInfos.length} 张图片`);

  // 辅助函数：根据字号计算首行缩进（2个汉字宽度）
  function calculateFirstLineIndent(fontSize) {
    // fontSize是半磅单位，1个汉字宽度约等于字号，2个汉字宽度为字号的2倍
    // 1 twip = 1/20 pt，所以需要乘以20将点转换为twip
    return fontSize * 2 * 20;
  }

  // 转换字符到twip的辅助函数
  function convertCharesToTwip(inches) {
    return Math.round(inches * 180);
  }

  // 创建一个numbering配置
  const numberingConfig = [
    {
      reference: "my-heading-style",
      levels: [
        {
          level: 2,
          format: LevelFormat.CHINESE_COUNTING, // 数字格式
          text: "%3、",
          alignment: AlignmentType.START,
          start: 1, // 明确设置从1开始
          style: {
            run: {
              size: 32, // 三号字→16磅
              font: "黑体", // 黑体
              color: "000000", // 黑色
              bold: true, // 加粗
            },
            paragraph: {
              alignment: AlignmentType.START,
              indent: {
                left: convertCharesToTwip(0),
              },
            },
          },
        },
        {
          level: 3,
          format: LevelFormat.DECIMAL,
          text: "%4.",
          alignment: AlignmentType.START,
          style: {
            paragraph: {
              indent: {
                left: convertCharesToTwip(0.5),
              },
            },
          },
        },
        {
          level: 4,
          format: LevelFormat.LOWER_LETTER,
          text: "%5)",
          alignment: AlignmentType.START,
          style: {
            paragraph: {
              indent: {
                left: convertCharesToTwip(1),
              },
            },
          },
        },
        {
          level: 5,
          format: LevelFormat.UPPER_LETTER,
          text: "%6)",
          alignment: AlignmentType.START,
          style: {
            paragraph: {
              indent: {
                left: convertCharesToTwip(1.5),
              },
            },
          },
        },
        {
          level: 6,
          format: LevelFormat.UPPER_LETTER,
          text: "%7)",
          alignment: AlignmentType.START,
          style: {
            paragraph: {
              indent: {
                left: convertCharesToTwip(2),
              },
            },
          },
        },
      ],
    },
    {
      reference: "my-paragraph-style",
      levels: [
        {
          level: 0,
          format: LevelFormat.DECIMAL,
          text: "%1.",
          alignment: AlignmentType.START,
          style: {
            run: {
              size: 24, // 小四号→12磅
              font: "仿宋", // 仿宋
              color: "000000", // 黑色
            },
            paragraph: {
              alignment: AlignmentType.JUSTIFIED, // 两端对齐
              indent: {
                left: convertCharesToTwip(0), // 左缩进2个汉字宽度，字号是24磅
                hanging: convertCharesToTwip(0), // 悬挂缩进2个汉字宽度，字号是24磅
              },
              spacing: {
                before: 0,
                after: 0,
                line: 600, // 30磅 = 600 twip
                lineRule: LineRuleType.EXACT,
              },
            },
          },
        },
      ],
    },
    {
      reference: "my-Unordered-list",
      levels: [
        {
          level: 0,
          format: LevelFormat.BULLET,
          text: "\u25CF",
          alignment: AlignmentType.LEFT,
          style: {
            paragraph: {
              indent: {
                left: convertCharesToTwip(0),
                hanging: convertCharesToTwip(0),
              },
            },
          },
        },
      ],
    },
    {
      reference: "my-task-list",
      levels: [
        {
          level: 0,
          format: LevelFormat.BULLET,
          text: "\u25A0", // 使用方块作为任务列表的标记
          alignment: AlignmentType.LEFT,
          style: {
            paragraph: {
              indent: {
                left: convertCharesToTwip(0),
                hanging: convertCharesToTwip(0),
              },
            },
          },
        },
      ],
    },
  ];

  //heading的识别方法
  function headingRecognition(level, hasNumber, inlineStyles) {
    console.log(level, hasNumber, inlineStyles);
    //准备一个包含TextRun的数组
    const textRuns = [];
    //根据标题级别设置不同的字体样式
    let fontSize, fontFamily, isBold;
    let hangingIndent;

    switch (level) {
      case 2: // 一级标题
        fontSize = 32; // 16磅
        fontFamily = "黑体";
        isBold = true;
        hangingIndent = convertCharesToTwip(1.5);
        break;
      case 3: // 二级标题
        fontSize = 32; // 16磅
        fontFamily = "楷体";
        isBold = true;
        hangingIndent = convertCharesToTwip(1.5);
        break;
      case 4: // 三级标题
        fontSize = 28; // 14磅
        fontFamily = "仿宋";
        isBold = true;
        hangingIndent = convertCharesToTwip(1.5);
        break;
      case 5: // 四级标题
        fontSize = 24; // 12磅
        fontFamily = "仿宋";
        isBold = true;
        hangingIndent = convertCharesToTwip(1.5);
        break;
      case 6: // 五级标题
        fontSize = 21; // 10.5磅
        fontFamily = "仿宋";
        isBold = true;
        hangingIndent = convertCharesToTwip(1.5);
        break;
      default:
        fontSize = 32;
        fontFamily = "黑体";
        isBold = true;
        hangingIndent = convertCharesToTwip(1.5);
    }
    //遍历inlineStyles
    inlineStyles.forEach((word) => {
      textRuns.push(
        new TextRun({
          text: word.content,
          size: fontSize,
          font: fontFamily,
          bold: isBold,
          indent: {
            hanging: hangingIndent,
          },
        })
      );
    });

    //判断标题是否需要编号
    if (hasNumber) {
      //有标题编号
      return new Paragraph({
        numbering: {
          reference: "my-heading-style",
          level: level,
        },
        heading: setHeading(level),
        children: textRuns,
      });
    } else {
      //无标题编号
      return new Paragraph({
        heading: setHeading(level),
        children: textRuns,
        style: `heading${level}`,
      });
    }
  }

  //根据level设置heading的值
  function setHeading(level) {
    switch (level) {
      case 1:
        return HeadingLevel.TITLE;
      case 2:
        return HeadingLevel.HEADING_1;
      case 3:
        return HeadingLevel.HEADING_2;
      case 4:
        return HeadingLevel.HEADING_3;
      case 5:
        return HeadingLevel.HEADING_4;
      case 6:
        return HeadingLevel.HEADING_5;
      default:
        return HeadingLevel.HEADING_1;
    }
  }

  //paragraph的识别方法
  function paragraphRecognition(hasNumber, inlineStyles) {
    if (hasNumber) {
      return new Paragraph({
        numbering: {
          reference: "my-paragraph-style",
          level: 0,
        },
        children: inlineStyles.map(
          (style) =>
            new TextRun({
              text: style.content,
              bold: style.bold,
              italics: style.italics,
              strike: style.strike,
              underline: style.underline,
              superScript: style.superScript,
              subScript: style.subScript,
              size: 24, // 小四号→12磅
              font: "仿宋", // 仿宋
              color: "000000", // 黑色
            })
        ),
      });
    } else {
      return new Paragraph({
        children: inlineStyles.map(
          (style) =>
            new TextRun({
              text: style.content,
              bold: style.bold,
              italics: style.italics,
              strike: style.strike,
              underline: style.underline,
              superScript: style.superScript,
              subScript: style.subScript,
              size: 24, // 小四号→12磅
              font: "仿宋", // 仿宋
              color: "000000", // 黑色
            })
        ),
      });
    }
  }

  //list的识别方法
  function listRecognition(hasNumber, inlineStyles, level) {
    return new Paragraph({
      numbering: {
        reference: hasNumber ? "my-paragraph-style" : "my-Unordered-list",
        level: level,
      },
      children: inlineStyles.map(
        (style) =>
          new TextRun({
            text: style.content,
            bold: style.bold,
            italics: style.italics,
            strike: style.strike,
            underline: style.underline,
            superScript: style.superScript,
            subScript: style.subScript,
            size: 24,
            font: "仿宋",
            color: "000000",
          })
      ),
    });
  }

  //task的识别方法
  function taskRecognition(inlineStyles, isChecked) {
    return new Paragraph({
      children: [
        // 添加复选框
        new CheckBox({
          checked: isChecked,
          size: 24,
        }),
        // 添加空格
        new TextRun({
          text: " ",
          size: 24,
        }),
        // 添加任务内容
        ...inlineStyles.map(
          (style) =>
            new TextRun({
              text: style.content,
              bold: style.bold,
              italics: style.italics,
              strike: style.strike,
              underline: style.underline,
              superScript: style.superScript,
              subScript: style.subScript,
              size: 24,
              font: "仿宋",
              color: "000000",
            })
        ),
      ],
    });
  }

  //quote的识别方法
  function quoteRecognition(inlineStyles) {
    return new Paragraph({
      style: "Intense Quote",
      alignment: AlignmentType.CENTER,
      spacing: {
        before: 360,
        after: 360,
      },
      indent: {
        left: 864,
        right: 864,
      },
      border: {
        top: { style: BorderStyle.SINGLE, size: 4, color: "0F4761" },
        bottom: { style: BorderStyle.SINGLE, size: 4, color: "0F4761" },
      },
      children: inlineStyles.map(
        (style) =>
          new TextRun({
            text: style.content,
            size: 24,
            font: "仿宋",
            color: "0F4761",
            italics: true,
          })
      ),
    });
  }

  //代码块的识别方法
  function codeBlockRecognition(inlineStyles) {
    return new Paragraph({
      style: "Code Block",
      alignment: AlignmentType.LEFT,
      spacing: {
        before: 120,
        after: 120,
        line: 360, // 18磅行距
        lineRule: LineRuleType.EXACT,
      },
      indent: {
        left: 720, // 36磅左缩进
        right: 720, // 36磅右缩进
      },
      border: {
        top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
        left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
        right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
      },
      shading: {
        fill: "F5F5F5", // 浅灰色背景
      },
      children: inlineStyles.map(
        (style) =>
          new TextRun({
            text: style.content,
            size: 24, // 12磅
            font: "等线", // 等宽字体
            color: "000000", // 黑色
          })
      ),
    });
  }

  //hyperlink的识别方法
  function hyperlinkRecognition(text, url, title) {
    return new Paragraph({
      children: [
        new ExternalHyperlink({
          children: [
            new TextRun({
              text: text,
              color: "0000FF", // 蓝色
              underline: {
                type: "single",
                color: "0000FF",
              },
            }),
          ],
          link: url,
        }),
      ],
    });
  }

  // 添加表格处理函数
  function tableRecognition(headers, alignments, rows) {
    // 创建表头行
    const headerRow = new TableRow({
      children: headers.map((header, index) => {
        return new TableCell({
          width: {
            size: 3000,
            type: WidthType.DXA,
          },
          children: [
            new Paragraph({
              alignment: getAlignmentType(alignments[index]),
              indent: { firstLine: 0 },
              children: [
                new TextRun({
                  text: header.fullContent,
                  bold: true,
                  size: 24,
                  font: "仿宋",
                  color: "000000",
                }),
              ],
            }),
          ],
        });
      }),
    });

    // 创建数据行
    const dataRows = rows.map((row) => {
      return new TableRow({
        children: row.map((cell, index) => {
          return new TableCell({
            width: {
              size: 3000,
              type: WidthType.DXA,
            },
            children: [
              new Paragraph({
                alignment: getAlignmentType(alignments[index]),
                indent: { firstLine: 0 },
                children: [
                  new TextRun({
                    text: cell.fullContent,
                    size: 24,
                    font: "仿宋",
                    color: "000000",
                  }),
                ],
              }),
            ],
          });
        }),
      });
    });

    // 创建表格
    return new Table({
      width: {
        size: 100,
        type: WidthType.PERCENTAGE,
      },
      alignment: AlignmentType.CENTER, // 表格整体居中对齐
      margins: {
        top: 0,
        bottom: 0,
        left: 0,
        right: 0,
      },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        insideHorizontal: {
          style: BorderStyle.SINGLE,
          size: 1,
          color: "000000",
        },
        insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
      },
      rows: [headerRow, ...dataRows],
    });
  }

  // 辅助函数：将字符串对齐方式转换为docx.js的AlignmentType
  function getAlignmentType(alignment) {
    switch (alignment) {
      case "left":
        return AlignmentType.LEFT;
      case "center":
        return AlignmentType.CENTER;
      case "right":
        return AlignmentType.RIGHT;
      default:
        return AlignmentType.LEFT;
    }
  }

  // 水平线的识别方法
  function horizontalRuleRecognition() {
    return new Paragraph({
      spacing: {
        before: 200, // 段前10磅
        after: 200, // 段后10磅
      },
      border: {
        bottom: {
          color: "#CCCCCC", // 灰色
          space: 1, // 间距
          style: BorderStyle.SINGLE, // 单线
          size: 6, // 线宽
        },
      },
    });
  }

  // 处理单个图片的函数
  async function processImage(url) {
    try {
      // 获取图片文件
      const response = await fetch(url);
      const blob = await response.blob();

      // 获取图片尺寸
      const img = new Image();
      await new Promise((resolve, reject) => {
        img.onload = resolve;
        img.onerror = reject;
        img.src = URL.createObjectURL(blob);
      });

      // 转换为ArrayBuffer
      const buffer = await blob.arrayBuffer();

      // 设置图片尺寸
      const maxWidth = 800; // 最大宽度
      const maxHeight = 600; // 最大高度
      let width = img.naturalWidth;
      let height = img.naturalHeight;

      // 调整图片尺寸
      if (width > maxWidth) {
        height = (maxWidth / width) * height;
        width = maxWidth;
      }
      if (height > maxHeight) {
        width = (maxHeight / height) * width;
        height = maxHeight;
      }

      // 从 URL 中推断图片类型
      const imageType = url.split(".").pop().toLowerCase(); // 获取文件扩展名

      return {
        buffer,
        width,
        height,
        type: imageType, // 返回图片类型
      };
    } catch (error) {
      console.error("加载图片失败:", error);
      throw error;
    }
  }

  // 处理json数据，将内容添加到paragraphs数组
  for (const child of json.children) {
    console.log(child);
    //判断child的type
    if (child.type === "heading") {
      //使用heading的识别方法
      const headingParagraph = headingRecognition(
        child.level,
        child.hasNumber,
        child.inlineStyles
      );
      console.log(headingParagraph);
      paragraphs.push(headingParagraph);
    }
    if (child.type === "paragraph") {
      //使用paragraph的识别方法
      const textParagraph = paragraphRecognition(
        child.hasNumber,
        child.inlineStyles
      );
      paragraphs.push(textParagraph);
    }
    //处理列表
    if (child.type === "list") {
      //使用list的识别方法
      const listParagraph = listRecognition(
        child.hasNumber,
        child.inlineStyles,
        child.level
      );
      paragraphs.push(listParagraph);
    }
    //处理任务列表
    if (child.type === "task") {
      //使用task的识别方法
      const taskParagraph = taskRecognition(
        child.inlineStyles,
        child.isChecked
      );
      paragraphs.push(taskParagraph);
    }
    //处理引用
    if (child.type === "quote") {
      //使用quote的识别方法
      const quoteParagraph = quoteRecognition(child.inlineStyles);
      paragraphs.push(quoteParagraph);
    }
    //处理代码块
    if (child.type === "code") {
      //使用codeBlock的识别方法
      const codeBlockParagraph = codeBlockRecognition(child.inlineStyles);
      paragraphs.push(codeBlockParagraph);
    }
    //处理链接
    if (child.type === "hyperlink") {
      //使用hyperlink的识别方法
      const hyperlinkParagraph = hyperlinkRecognition(
        child.text,
        child.url,
        child.title
      );
      paragraphs.push(hyperlinkParagraph);
    }
    //处理表格
    if (child.type === "table") {
      //使用table的识别方法
      const table = tableRecognition(
        child.headers,
        child.alignments,
        child.rows
      );
      paragraphs.push(table);
    }
    //处理水平线
    if (child.type === "horizontal_rule") {
      //使用horizontalRule的识别方法
      const horizontalRuleParagraph = horizontalRuleRecognition();
      paragraphs.push(horizontalRuleParagraph);
    }
    //处理图片
    if (child.type === "image") {
      try {
        // 处理图片并等待完成
        const imageData = await processImage(child.url);
        const imageParagraph = new Paragraph({
          children: [
            new ImageRun({
              type: imageData.type, // 使用推断的图片类型
              data: imageData.buffer,
              transformation: {
                width: 200,
                height: 200,
              },
              floating: {
                horizontalPosition: {
                  relative: HorizontalPositionRelativeFrom.PAGE,
                  align: HorizontalPositionAlign.CENTER,
                },
                verticalPosition: {
                  relative: VerticalPositionRelativeFrom.LINE,
                  align: VerticalPositionAlign.CENTER,
                },
                wrap: {
                  type: TextWrappingType.SQUARE,
                  side: TextWrappingSide.BOTH_SIDES,
                },
              },
            }),
          ],
          alignment: AlignmentType.CENTER,
        });
        paragraphs.push(imageParagraph);
      } catch (error) {
        // 如果图片加载失败，使用占位符
        const placeholderParagraph = new Paragraph({
          children: [
            new TextRun({
              text: `[图片: ${child.alt || "加载失败"}]`,
              color: "FF0000",
            }),
          ],
          alignment: AlignmentType.CENTER,
        });
        paragraphs.push(placeholderParagraph);
      }
    }
  }

  // 创建文档（一次性创建，包含所有内容）
  const doc = new Document({
    styles: {
      default: {
        document: {
          run: {
            size: 24, // 小四号→12磅
            font: "仿宋", // 仿宋
            color: "000000", // 黑色
          },
          paragraph: {
            alignment: AlignmentType.JUSTIFIED, // 两端对齐
            spacing: {
              before: 0,
              after: 0,
              line: 600, // 30磅 = 600 twip
              lineRule: LineRuleType.EXACT,
            },
            indent: {
              firstLine: 480, // 首行缩进2个汉字宽度（12磅 × 2 = 24磅 = 480 twip）
            },
          },
        },
        title: {
          run: {
            size: 44, // 二号字→22磅
            font: "黑体", // 黑体
            color: "000000", // 黑色
            bold: true, // 加粗
          },
          paragraph: {
            alignment: AlignmentType.CENTER, // 居中对齐
            spacing: {
              before: 480, // 段前24磅（约合480 twip）
              after: 480, // 段前24磅（约合480 twip）
            },
            lineSpacing: { before: 560, lineRule: LineRuleType.EXACT }, //行距设置为固定值28磅，约合560 twip
            indent: {
              left: convertCharesToTwip(0), // 左缩进0
              firstLine: convertCharesToTwip(0), // 首行缩进
            },
          },
        },
        heading1: {
          run: {
            size: 32, // 三号字→16磅
            font: "黑体", // 黑体
            color: "000000", // 黑色
            bold: true, // 加粗
          },
          paragraph: {
            alignment: AlignmentType.START, // 居左对齐
            spacing: {
              before: 240, // 段前12磅（约合240 twip）
              after: 120, // 段后6磅（约合120 twip）
            },
            lineSpacing: { before: 440, lineRule: LineRuleType.EXACT }, //行距设置为固定值22磅，约合440 twip
            indent: {
              left: convertCharesToTwip(0), // 左缩进0
            },
          },
        },
        heading2: {
          run: {
            size: 32, //三号字→16磅（与一级标题同字号，通过字体区分层级）
            font: "楷体", // 楷体
            color: "000000", // 黑色
            bold: true, // 加粗
          },
          paragraph: {
            alignment: AlignmentType.START, // 居左对齐
            spacing: {
              before: 120, // 段前6磅（约合120 twip）
              after: 60, // 段后3磅（约合60 twip）
            },
            lineSpacing: { before: 360, lineRule: LineRuleType.EXACT }, //行距设置为固定值18磅，约合360 twip
            indent: {
              left: convertCharesToTwip(0), // 左缩进0
            },
          },
        },
        heading3: {
          run: {
            size: 28, // 四号字→14磅
            font: "仿宋", // 仿宋
            color: "000000", // 黑色
            bold: true, // 加粗
          },
          paragraph: {
            alignment: AlignmentType.START, // 居左对齐
            spacing: {
              before: 60, // 段前3磅（约合60 twip）
              after: 60, // 段后3磅（约合60 twip）
            },
            lineSpacing: { before: 320, lineRule: LineRuleType.EXACT }, //行距设置为固定值16磅，约合320 twip
            indent: {
              left: convertCharesToTwip(0), // 左缩进0
            },
          },
        },
        heading4: {
          run: {
            size: 24, // 小四号→12磅
            font: "仿宋", // 仿宋
            color: "000000", // 黑色
            bold: true, // 加粗
          },
          paragraph: {
            alignment: AlignmentType.START, // 居左对齐
            spacing: {
              before: 30, // 段前1.5磅（约合30 twip）
              after: 30, // 段后1.5磅（约合30 twip）
            },
            lineSpacing: { before: 280, lineRule: LineRuleType.EXACT }, //行距设置为固定值14磅，约合280 twip
            indent: {
              left: convertCharesToTwip(0), // 左缩进0
            },
          },
        },
        heading5: {
          run: {
            size: 21, // 小五号→10.5磅
            font: "仿宋", // 仿宋
            color: "000000", // 黑色
            bold: true, // 加粗
          },
          paragraph: {
            alignment: AlignmentType.START, // 居左对齐
            spacing: {
              before: 0, // 无段前间距
              after: 0, // 无段前间距
            },
            lineSpacing: { before: 240, lineRule: LineRuleType.EXACT }, //行距设置为固定值12磅，约合240 twip
            indent: {
              left: convertCharesToTwip(0), // 左缩进0
            },
          },
        },
      },
    },
    numbering: {
      config: numberingConfig,
    },
    sections: [
      {
        properties: {
          page: {
            size: {
              width: 12240, // A4宽度（595pt × 20 = 11900twip，考虑四舍五入设为12240）
              height: 15840, // A4高度（842pt × 20 = 16840twip，考虑装订边距设为15840）
            },
            orientation: "portrait", // 纵向（默认）
            margin: {
              top: 1440, // 上边距2.54cm（1440twip）
              bottom: 1440, // 下边距2.54cm（1440twip）
              left: 1800, // 左边距3.18cm（1800twip）
              right: 1800, // 右边距3.18cm（1800twip）
              gutter: 0, // 装订线间距
            },
          },
        },
        children: paragraphs,
      },
    ],
  });

  console.log("文档创建完成");
  console.log(doc);

  // 使用toBlob而不是toBuffer，适用于浏览器环境
  return Packer.toBlob(doc)
    .then((blob) => {
      saveAs(blob, "Test Document.docx");
      console.log("文档已生成并保存");
    })
    .catch((error) => {
      console.error("生成文档时出错:", error);
      throw error;
    });
}
