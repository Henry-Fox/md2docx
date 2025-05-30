// 引入需要的库
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  LineRuleType,
  LevelFormat,
} from "docx";
import { saveAs } from "file-saver";
import { Md2Json } from "./md2json.js";
import runTest from "./json2docx.js";
import { marked } from "marked";

/**
 * 简化版的Markdown到Docx转换器
 */
class SimpleMd2Docx {
  constructor() {
    console.log("SimpleMd2Docx 初始化");
    this.md2json = new Md2Json();
  }

  /**
   * 转换Markdown为Word文档
   * @param {string} markdown - Markdown文本
   * @returns {Promise<boolean>} 转换是否成功
   */
  async convertToDocx(markdown) {
    try {
      console.log("开始转换Markdown到Word文档...");

      // 使用Md2Json将markdown转换为json
      const jsonData = this.md2json.convert(markdown);
      console.log("Markdown转换为JSON成功:", jsonData);

      // 使用json2docx.js中的runTest函数生成docx
      await runTest(jsonData);

      return true;
    } catch (error) {
      console.error("转换过程中出错:", error);
      throw error;
    }
  }

  /**
   * 直接使用marked解析结果转换为Word文档
   * @param {string} markdown - Markdown文本
   * @returns {Promise<boolean>} 转换是否成功
   */
  async convertToDocxDirect(markdown) {
    try {
      console.log("开始直接转换Markdown到Word文档...");

      // 保存当前markdown内容
      this.currentMarkdown = markdown;

      // 配置marked选项
      // marked.setOptions({
      //   gfm: true,  // 启用GitHub风格的Markdown
      //   breaks: true,  // 启用换行符转换
      //   headerIds: true,  // 为标题添加id
      //   mangle: false,  // 不转义标题中的特殊字符
      //   pedantic: true,  // 严格遵循markdown规范
      //   smartLists: true,  // 使用更智能的列表行为
      //   smartypants: false,  // 不使用更智能的标点符号
      //   xhtml: false  // 不使用xhtml
      // });

      // 使用marked解析Markdown，得到token队列
      const tokens = marked.lexer(markdown);
      console.log("Markdown解析结果:", tokens);

      // 创建一个数组用于收集所有段落
      const paragraphs = [];

      // 遍历token队列，生成对应的docx段落
      for (const token of tokens) {
        switch (token.type) {
          case "heading":
            paragraphs.push(this.createHeadingFromMarked(token));
            break;
          case "paragraph":
            paragraphs.push(this.createParagraphFromMarked(token));
            break;
          case "code":
            paragraphs.push(this.createCodeBlockFromMarked(token));
            break;
          case "list":
            paragraphs.push(...this.createListFromMarked(token));
            break;
          case "table":
            paragraphs.push(this.createTableFromMarked(token));
            break;
          case "blockquote":
            paragraphs.push(this.createBlockquoteFromMarked(token));
            break;
          case "hr":
            paragraphs.push(this.createHorizontalRule());
            break;
          case "image":
            paragraphs.push(this.createImageFromMarked(token));
            break;
          case "space":
            paragraphs.push(this.createSpaceFromMarked());
            break;
          default:
            console.warn(`未支持的token类型: ${token.type}`);
            break;
        }
      }

      // 创建一个numbering配置
      const numberingConfig = [
        {
          reference: "my-heading-style",
          levels: [
            {
              level: 0, // 对应Markdown的##（depth:2）
              format: LevelFormat.DECIMAL,
              text: "%1.",
              alignment: AlignmentType.START,
              start: 1,
              style: {
                run: {
                  size: 32,
                  font: "黑体",
                  color: "000000",
                  bold: true,
                },
                paragraph: {
                  alignment: AlignmentType.START,
                  indent: {
                    left: 0,
                    firstLine: 0,
                  },
                },
              },
            },
            {
              level: 1, // 对应Markdown的###（depth:3）
              format: LevelFormat.DECIMAL,
              text: "%2.",
              alignment: AlignmentType.START,
              style: {
                paragraph: {
                  indent: {
                    left: 0,
                    firstLine: 0,
                  },
                },
              },
            },
            {
              level: 2, // 对应Markdown的####（depth:4）
              format: LevelFormat.LOWER_LETTER,
              text: "%3)",
              alignment: AlignmentType.START,
              style: {
                paragraph: {
                  indent: {
                    left: 0,
                    firstLine: 0,
                  },
                },
              },
            },
            {
              level: 3, // 对应Markdown的#####（depth:5）
              format: LevelFormat.UPPER_LETTER,
              text: "%4)",
              alignment: AlignmentType.START,
              style: {
                paragraph: {
                  indent: {
                    left: 0,
                    firstLine: 0,
                  },
                },
              },
            },
            {
              level: 4, // 对应Markdown的######（depth:6）
              format: LevelFormat.UPPER_LETTER,
              text: "%5)",
              alignment: AlignmentType.START,
              style: {
                paragraph: {
                  indent: {
                    left: 0,
                    firstLine: 0,
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
                  size: 24,
                  font: "仿宋",
                  color: "000000",
                },
                paragraph: {
                  alignment: AlignmentType.JUSTIFIED,
                  indent: {
                    left: 0,
                    hanging: 0,
                  },
                  spacing: {
                    before: 0,
                    after: 0,
                    line: 600,
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
              format: LevelFormat.JUSTIFIED,
              text: "\u25CF",
              alignment: AlignmentType.LEFT,
              style: {
                paragraph: {
                  indent: {
                    left: 0, // 左缩进0
                    hanging: 0, // 悬挂缩进0
                    firstLine: 480, // 首行缩进2个汉字宽度（12磅 × 2 = 24磅 = 480 twip）
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
                    left: 0,
                    hanging: 0,
                  },
                },
              },
            },
          ],
        },
      ];

      // 创建文档
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
                  left: 0, // 左缩进0
                  firstLine: 0, // 首行缩进
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
                  left: 0, // 左缩进0
                  firstLine: 0, // 首行缩进
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
                  left: 0, // 左缩进0
                  firstLine: 0, // 首行缩进
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
                  left: 0, // 左缩进0
                  firstLine: 0, // 首行缩进
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
                  left: 0, // 左缩进0
                  firstLine: 0, // 首行缩进
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
                  left: 0, // 左缩进0
                  firstLine: 0, // 首行缩进
                },
              },
            },
          },
          footnote: {
            run: {
              size: 20, // 10磅
              font: "仿宋",
              color: "000000",
            },
            paragraph: {
              alignment: AlignmentType.JUSTIFIED,
              spacing: {
                before: 120,
                after: 120,
                line: 400,
              },
              indent: {
                left: 720, // 36磅左缩进
                hanging: 360, // 18磅悬挂缩进
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

      // 生成并保存文档
      const blob = await Packer.toBlob(doc);
      saveAs(blob, "document.docx");

      return true;
    } catch (error) {
      console.error("直接转换过程中出错:", error);
      throw error;
    }
  }

  /**
   * 处理marked解析的单个token
   * @param {Object} token - marked解析的token
   * @returns {Promise<Object|Array|null>} 处理结果
   */
  async processMarkedToken(token) {
    switch (token.type) {
      case "heading":
        // 检查是否真的是标题（不应该包含换行符）
        if (token.text.includes("\n")) {
          // 如果包含换行符，说明是段落而不是标题
          return this.createParagraphFromMarked({
            type: "paragraph",
            text: token.text,
          });
        }
        return this.createHeadingFromMarked(token);
      case "paragraph":
        return this.createParagraphFromMarked(token);
      case "code":
        return this.createCodeBlockFromMarked(token);
      case "list":
        return this.createListFromMarked(token);
      case "table":
        return this.createTableFromMarked(token);
      case "blockquote":
        return this.createBlockquoteFromMarked(token);
      case "hr":
        return this.createHorizontalRule();
      case "image":
        return this.createImageFromMarked(token);
      case "space":
        return this.createSpaceFromMarked();
      default:
        console.warn(`未支持的token类型: ${token.type}`);
        return null;
    }
  }

  // 创建空行
  createSpaceFromMarked() {
    //返回一个空段落
    return new Paragraph({
      text: "",
      alignment: AlignmentType.START,
    });
  }

  /**
   * 从marked token创建标题
   * @param {Object} token - marked解析的标题token
   * @returns {Object} docx Paragraph对象
   */
  createHeadingFromMarked(token) {
    const level = token.depth;
    const match = token.text.match(/^\d+\./);
    const hasNumber = match !== null;
    const content = hasNumber ? token.text.replace(/^\d+\.\s*/, "") : token.text;

    // 设置标题样式
    let headingLevel;
    switch (level) {
      case 1: headingLevel = HeadingLevel.TITLE; break;
      case 2: headingLevel = HeadingLevel.HEADING_1; break;
      case 3: headingLevel = HeadingLevel.HEADING_2; break;
      case 4: headingLevel = HeadingLevel.HEADING_3; break;
      case 5: headingLevel = HeadingLevel.HEADING_4; break;
      case 6: headingLevel = HeadingLevel.HEADING_5; break;
      default: headingLevel = HeadingLevel.TITLE;
    }

    const textRuns = [
      new TextRun({
        text: content,
        bold: true,
      })
    ];

    if (level === 1) {
      // 一级标题永远不用编号
      return new Paragraph({
        heading: headingLevel,
        children: textRuns,
      });
    } else if (hasNumber) {
      // 二级及以下标题，有序号前缀时加自动编号
      return new Paragraph({
        numbering: {
          reference: "my-heading-style",
          level: level - 2, // Markdown的##对应Word的level:0
        },
        heading: headingLevel,
        children: textRuns,
      });
    } else {
      // 二级及以下标题，无序号前缀时只用样式
      return new Paragraph({
        heading: headingLevel,
        children: textRuns,
      });
    }
  }

  /**
   * 递归处理 marked tokens，生成 docx 的 TextRun 数组
   * @param {Array} tokens - marked 的 tokens
   * @param {Object} style - 当前叠加的样式
   * @returns {Array} TextRun[]
   */
  parseTokens(tokens, style = {}) {
    let runs = [];
    for (const t of tokens) {
      if (t.type === 'text') {
        runs.push(new TextRun({ text: t.text, ...style }));
      } else if (t.type === 'strong') {
        runs = runs.concat(this.parseTokens(t.tokens || [], { ...style, bold: true }));
      } else if (t.type === 'em') {
        runs = runs.concat(this.parseTokens(t.tokens || [], { ...style, italics: true }));
      } else if (t.type === 'del') {
        runs = runs.concat(this.parseTokens(t.tokens || [], { ...style, strike: true }));
      } else {
        // 其他类型直接递归
        runs = runs.concat(this.parseTokens(t.tokens || [], style));
      }
    }
    return runs;
  }

  /**
   * 从marked token创建段落
   * @param {Object} token - marked解析的段落token
   * @returns {Object} docx Paragraph对象
   */
  createParagraphFromMarked(token) {
    // 检查是否包含序号
    const hasNumber = /^\d+\.\s/.test(token.text);
    // 如果有序号，去掉序号部分
    const cleanText = hasNumber
      ? token.text.replace(/^\d+\.\s*/, "")
      : token.text;

    let textRuns = [];
    if (token.tokens) {
      textRuns = this.parseTokens(token.tokens);
    } else {
      textRuns = [
        new TextRun({
          text: cleanText,
          bold: false,
          italics: false,
          strike: false,
        })
      ];
    }

    const paragraph = new Paragraph({
      children: textRuns,
      alignment: AlignmentType.JUSTIFIED,
    });

    // 如果有序号，添加编号
    if (hasNumber) {
      paragraph.numbering = {
        reference: "my-paragraph-style",
        level: 0,
      };
    }

    return paragraph;
  }

  /**
   * 从marked token创建代码块
   * @param {Object} token - marked解析的代码块token
   * @returns {Object} docx Paragraph对象
   */
  createCodeBlockFromMarked(token) {
    return new Paragraph({
      text: token.text,
      alignment: AlignmentType.LEFT,
      indent: {
        left: 720,
      },
    });
  }

  /**
   * 从marked token创建列表
   * @param {Object} token - marked解析的列表token
   * @returns {Array} docx Paragraph对象数组
   */
  createListFromMarked(token) {
    const paragraphs = [];

    // 递归处理列表项
    const processListItem = (item, level = 0) => {
      // 创建当前列表项
      const paragraph = new Paragraph({
        numbering: {
          reference: token.ordered ? "my-paragraph-style" : "my-Unordered-list",
          level: level,
        },
        children: item.tokens ? this.parseTokens(item.tokens) : [
          new TextRun({
            text: item.text,
            size: 24,
            font: "仿宋",
            color: "000000",
          })
        ],
      });
      paragraphs.push(paragraph);

      // 递归处理子列表
      if (item.items && item.items.length > 0) {
        item.items.forEach(subItem => {
          processListItem(subItem, level + 1);
        });
      }
    };

    // 处理所有顶级列表项
    token.items.forEach(item => {
      processListItem(item);
    });

    return paragraphs;
  }

  /**
   * 从marked token创建表格
   * @param {Object} token - marked解析的表格token
   * @returns {Object} docx Table对象
   */
  createTableFromMarked(token) {
    const rows = [];

    // 添加表头
    const headerRow = new TableRow({
      children: token.header.map(
        (cell) =>
          new TableCell({
            children: [new Paragraph({ text: cell.text })],
          })
      ),
    });
    rows.push(headerRow);

    // 添加数据行
    token.rows.forEach((row) => {
      const tableRow = new TableRow({
        children: row.map(
          (cell) =>
            new TableCell({
              children: [new Paragraph({ text: cell.text })],
            })
        ),
      });
      rows.push(tableRow);
    });

    return new Table({
      rows: rows,
      width: {
        size: 100,
        type: WidthType.PERCENTAGE,
      },
    });
  }

  /**
   * 从marked token创建引用块
   * @param {Object} token - marked解析的引用块token
   * @returns {Object} docx Paragraph对象
   */
  createBlockquoteFromMarked(token) {
    return new Paragraph({
      text: token.text,
      indent: {
        left: 720,
      },
      border: {
        left: {
          color: "CCCCCC",
          size: 4,
          style: BorderStyle.SINGLE,
        },
      },
    });
  }

  /**
   * 创建水平线
   * @returns {Object} docx Paragraph对象
   */
  createHorizontalRule() {
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

  /**
   * 从marked token创建图片
   * @param {Object} token - marked解析的图片token
   * @returns {Object} docx Paragraph对象
   */
  createImageFromMarked(token) {
    return new Paragraph({
      text: `[图片: ${token.text}]`,
      alignment: AlignmentType.CENTER,
    });
  }

  /**
   * 处理文档中的所有图片
   * @param {Array} tokens - marked解析的tokens
   * @returns {Promise<Array>} 处理后的图片信息数组
   */
  async processImages(tokens) {
    const imageInfos = [];
    for (const token of tokens) {
      if (token.type === "image") {
        try {
          const imageData = await this.loadImage(token.href);
          imageInfos.push({
            ...token,
            ...imageData,
          });
        } catch (error) {
          console.error(`处理图片失败: ${token.href}`, error);
        }
      }
    }
    return imageInfos;
  }

  /**
   * 加载图片
   * @param {string} url - 图片URL
   * @returns {Promise<Object>} 图片数据
   */
  async loadImage(url) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.crossOrigin = "anonymous";

      img.onload = async () => {
        try {
          const canvas = document.createElement("canvas");
          canvas.width = img.naturalWidth;
          canvas.height = img.naturalHeight;

          const ctx = canvas.getContext("2d");
          ctx.drawImage(img, 0, 0);

          const blob = await new Promise((resolve) => canvas.toBlob(resolve));
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
        reject(new Error(`加载图片失败: ${url}`));
      };

      img.src = url;
    });
  }

  /**
   * 处理单个元素
   * @param {Object} element - 元素对象
   * @returns {Object|Array} docx文档元素或元素数组
   */
  processElement(element) {
    if (!element) return null;

    switch (element.type) {
      case "heading":
        return this.createHeading(element);

      case "paragraph":
        return this.createParagraph(element);

      case "code_block":
        return this.createCodeBlock(element);

      case "list":
        return this.createList(element);

      case "table":
        return this.createTable(element);

      case "image":
        return this.createImagePlaceholder(element);

      case "horizontal_rule":
        return this.createHorizontalRule();

      case "blockquote":
        return this.createBlockquote(element);

      default:
        console.warn(`未支持的元素类型: ${element.type}`);
        return new Paragraph({
          children: [
            new TextRun({
              text: `[未支持的元素: ${element.type}]`,
              color: "999999",
            }),
          ],
        });
    }
  }

  /**
   * 创建标题
   * @param {Object} element - 标题元素
   * @returns {Object} docx Paragraph对象
   */
  createHeading(element) {
    console.log("创建标题:", element);

    const level = element.level || 1;
    let headingLevel;

    // 将Markdown的level转换为docx的HeadingLevel
    switch (level) {
      case 1:
        headingLevel = HeadingLevel.HEADING_1;
        break;
      case 2:
        headingLevel = HeadingLevel.HEADING_2;
        break;
      case 3:
        headingLevel = HeadingLevel.HEADING_3;
        break;
      case 4:
        headingLevel = HeadingLevel.HEADING_4;
        break;
      case 5:
        headingLevel = HeadingLevel.HEADING_5;
        break;
      case 6:
        headingLevel = HeadingLevel.HEADING_6;
        break;
      default:
        headingLevel = HeadingLevel.HEADING_1;
    }

    // 处理内联样式
    if (
      element.inlineStyles &&
      Array.isArray(element.inlineStyles) &&
      element.inlineStyles.length > 0
    ) {
      const children = [];

      element.inlineStyles.forEach((style) => {
        if (style.content) {
          children.push(
            new TextRun({
              text: style.content,
              bold: style.bold === true,
              italic: style.italic === true,
              strike: style.strike === true,
              underline: style.underline === true ? {} : undefined,
              superScript: style.superscript === true,
              subScript: style.subscript === true,
            })
          );
        }
      });

      if (children.length > 0) {
        return new Paragraph({
          heading: headingLevel,
          children,
        });
      }
    }

    // 提取标题文本（如果没有可用的内联样式）
    const titleText =
      element.fullContent || element.rawText || element.text || "";
    // 去掉Markdown的#符号
    const cleanTitle = titleText.replace(/^#+\s+/, "");

    console.log(`创建标题: '${cleanTitle}', 级别: ${level}`);

    // 创建标题段落
    return new Paragraph({
      text: cleanTitle,
      heading: headingLevel,
    });
  }

  /**
   * 创建段落
   * @param {Object} element - 段落元素
   * @returns {Object} docx Paragraph对象
   */
  createParagraph(element) {
    console.log("创建段落:", element);

    // 如果有内联样式
    if (
      element.inlineStyles &&
      Array.isArray(element.inlineStyles) &&
      element.inlineStyles.length > 0
    ) {
      const children = [];

      element.inlineStyles.forEach((style) => {
        if (style.content) {
          children.push(
            new TextRun({
              text: style.content,
              bold: style.bold === true,
              italic: style.italic === true,
              strike: style.strike === true,
              underline: style.underline === true ? {} : undefined,
              superScript: style.superscript === true,
              subScript: style.subscript === true,
            })
          );
        }
      });

      if (children.length > 0) {
        return new Paragraph({ children });
      }
    }

    // 如果没有内联样式但有fullContent
    if (element.fullContent) {
      return new Paragraph({ text: element.fullContent });
    }

    // 简单段落，使用任何可用的文本属性
    const text = element.rawText || element.text || "";
    return new Paragraph({ text });
  }

  /**
   * 创建代码块
   * @param {Object} element - 代码块元素
   * @returns {Object} docx Paragraph对象
   */
  createCodeBlock(element) {
    console.log("创建代码块:", element);

    const code = element.fullContent || element.content || element.text || "";
    const language = element.language || "";

    return new Paragraph({
      children: [
        new TextRun({
          text: `${language ? `[${language}] ` : ""}${code}`,
          font: "Courier New",
        }),
      ],
      indent: {
        left: 720, // 缩进量，720 = 0.5英寸
      },
    });
  }

  /**
   * 创建列表
   * @param {Object} element - 列表元素
   * @returns {Array} Paragraph对象数组
   */
  createList(element) {
    console.log("创建列表:", element);

    const items = element.items || [];
    const isOrdered = element.ordered === true;
    const result = [];

    // 对每个列表项创建一个段落
    items.forEach((item, index) => {
      const prefix = isOrdered ? `${index + 1}. ` : "• ";
      const itemText = item.fullContent || item.text || "";

      result.push(
        new Paragraph({
          text: `${prefix}${itemText}`,
          indent: {
            left: 720, // 缩进量
          },
        })
      );

      // 处理嵌套列表
      if (
        item.children &&
        Array.isArray(item.children) &&
        item.children.length > 0
      ) {
        item.children.forEach((child) => {
          const nestedItems = this.processElement(child);
          if (Array.isArray(nestedItems)) {
            result.push(...nestedItems);
          } else if (nestedItems) {
            result.push(nestedItems);
          }
        });
      }
    });

    return result;
  }

  /**
   * 创建表格
   * @param {Object} element - 表格元素
   * @returns {Object} docx Table对象
   */
  createTable(element) {
    console.log("创建表格:", element);

    // 检查表格数据
    if (
      !element.data ||
      !Array.isArray(element.data) ||
      element.data.length === 0
    ) {
      return new Paragraph({ text: "[空表格]" });
    }

    // 创建表格行
    const rows = [];
    element.data.forEach((rowData, rowIndex) => {
      if (!Array.isArray(rowData)) {
        console.warn(`表格行${rowIndex}不是数组`);
        return;
      }

      const cells = [];
      rowData.forEach((cellData, cellIndex) => {
        const cellText = cellData.fullContent || cellData.text || "";

        cells.push(
          new TableCell({
            children: [new Paragraph({ text: cellText })],
          })
        );
      });

      if (cells.length > 0) {
        rows.push(new TableRow({ children: cells }));
      }
    });

    if (rows.length === 0) {
      return new Paragraph({ text: "[空表格]" });
    }

    // 创建表格
    return new Table({
      rows: rows,
      width: {
        size: 100,
        type: WidthType.PERCENTAGE,
      },
    });
  }

  /**
   * 创建图片占位符
   * @param {Object} element - 图片元素
   * @returns {Object} docx Paragraph对象
   */
  createImagePlaceholder(element) {
    console.log("创建图片占位符:", element);

    const altText = element.alt || "[图片]";
    const url = element.url || "";

    return new Paragraph({
      children: [
        new TextRun({
          text: `[图片: ${altText}]${url ? ` ${url}` : ""}`,
          italic: true,
          color: "0000FF",
        }),
      ],
      alignment: AlignmentType.CENTER,
    });
  }

  /**
   * 创建引用块
   * @param {Object} element - 引用块元素
   * @returns {Object} docx Paragraph对象
   */
  createBlockquote(element) {
    console.log("创建引用块:", element);

    const text = element.fullContent || element.text || "";

    return new Paragraph({
      text: text,
      indent: {
        left: 720, // 缩进量
      },
      border: {
        left: {
          color: "CCCCCC",
          size: 4,
          style: BorderStyle.SINGLE,
        },
      },
      spacing: {
        before: 240, // 前间距
        after: 240, // 后间距
      },
    });
  }

  /**
   * 创建错误文档
   * @param {string} errorMessage - 错误信息
   * @returns {Promise<Blob>} 错误文档的Blob对象
   */
  async createErrorDocument(errorMessage) {
    try {
      const doc = new Document({
        sections: [
          {
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "文档生成失败",
                    bold: true,
                    color: "FF0000",
                    size: 36, // 18pt
                  }),
                ],
                alignment: AlignmentType.CENTER,
              }),
              new Paragraph({
                text: errorMessage,
                alignment: AlignmentType.CENTER,
              }),
              new Paragraph({
                text: `生成时间: ${new Date().toLocaleString()}`,
                alignment: AlignmentType.CENTER,
              }),
            ],
          },
        ],
      });

      return await Packer.toBlob(doc);
    } catch (error) {
      console.error("创建错误文档时出错:", error);

      // 创建最小的错误文档
      const minimalDoc = new Document({
        sections: [
          {
            children: [
              new Paragraph({
                text: `错误: ${errorMessage}`,
              }),
            ],
          },
        ],
      });

      return await Packer.toBlob(minimalDoc);
    }
  }

  /**
   * 保存为docx文件
   * @param {Blob} blob - docx文件的Blob对象
   * @param {string} filename - 文件名
   */
  saveAsDocx(blob, filename = "document.docx") {
    try {
      saveAs(blob, filename);
      console.log(`文档已保存为 ${filename}`);
    } catch (error) {
      console.error("保存文档时出错:", error);
      alert(`保存文档失败: ${error.message}`);
    }
  }
}

export default SimpleMd2Docx;
