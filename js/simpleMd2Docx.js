// 引入需要的库
import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType, BorderStyle, LineRuleType, LevelFormat } from "docx";
import { saveAs } from "file-saver";
import { Md2Json } from './md2json.js';
import runTest from './json2docx.js';
import { marked } from 'marked';

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
      console.error('转换过程中出错:', error);
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

      // 使用marked解析Markdown
      const tokens = marked.lexer(markdown);
      console.log("Markdown解析结果:", tokens);

      // 创建一个数组用于收集所有段落
      const paragraphs = [];

      // 处理所有图片
      console.log("开始处理图片...");
      const imageInfos = await this.processImages(tokens);
      console.log(`图片处理完成，共处理 ${imageInfos.length} 张图片`);

      // 处理每个token
      for (const token of tokens) {
        const result = await this.processMarkedToken(token);
        if (Array.isArray(result)) {
          paragraphs.push(...result);
        } else if (result) {
          paragraphs.push(result);
        }
      }

      // 转换字符到twip的辅助函数
      function convertCharesToTwip(inches) {
        return Math.round(inches * 180);
      }

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
                  before: 480, // 段前24磅
                  after: 480, // 段后24磅
                },
                lineSpacing: { before: 560, lineRule: LineRuleType.EXACT }, //行距设置为固定值28磅
                indent: {
                  left: convertCharesToTwip(0),
                  firstLine: convertCharesToTwip(0),
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
                  before: 240, // 段前12磅
                  after: 120, // 段后6磅
                },
                lineSpacing: { before: 440, lineRule: LineRuleType.EXACT }, //行距设置为固定值22磅
                indent: {
                  left: convertCharesToTwip(0),
                  firstLine: convertCharesToTwip(0),
                },
              },
            },
            heading2: {
              run: {
                size: 32, //三号字→16磅
                font: "楷体", // 楷体
                color: "000000", // 黑色
                bold: true, // 加粗
              },
              paragraph: {
                alignment: AlignmentType.START, // 居左对齐
                spacing: {
                  before: 120, // 段前6磅
                  after: 60, // 段后3磅
                },
                lineSpacing: { before: 360, lineRule: LineRuleType.EXACT }, //行距设置为固定值18磅
                indent: {
                  left: convertCharesToTwip(0),
                  firstLine: convertCharesToTwip(0),
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
                  before: 60, // 段前3磅
                  after: 60, // 段后3磅
                },
                lineSpacing: { before: 320, lineRule: LineRuleType.EXACT }, //行距设置为固定值16磅
                indent: {
                  left: convertCharesToTwip(0),
                  firstLine: convertCharesToTwip(0),
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
                  before: 30, // 段前1.5磅
                  after: 30, // 段后1.5磅
                },
                lineSpacing: { before: 280, lineRule: LineRuleType.EXACT }, //行距设置为固定值14磅
                indent: {
                  left: convertCharesToTwip(0),
                  firstLine: convertCharesToTwip(0),
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
                  after: 0, // 无段后间距
                },
                lineSpacing: { before: 240, lineRule: LineRuleType.EXACT }, //行距设置为固定值12磅
                indent: {
                  left: convertCharesToTwip(0),
                  firstLine: convertCharesToTwip(0),
                },
              },
            },
          },
        },
        numbering: {
          config: [
            {
              reference: "my-heading-style",
              levels: [
                {
                  level: 2,
                  format: LevelFormat.CHINESE_COUNTING,
                  text: "%3、",
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
                        left: convertCharesToTwip(0),
                        firstLine: convertCharesToTwip(0),
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
                        left: convertCharesToTwip(0),
                        firstLine: convertCharesToTwip(0),
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
                        left: convertCharesToTwip(0),
                        firstLine: convertCharesToTwip(0),
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
                        left: convertCharesToTwip(0),
                        firstLine: convertCharesToTwip(0),
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
                        left: convertCharesToTwip(0),
                        firstLine: convertCharesToTwip(0),
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
                        left: convertCharesToTwip(0),
                        hanging: convertCharesToTwip(0),
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
                        left: convertCharesToTwip(0),
                        hanging: convertCharesToTwip(0),
                        firstLine: 480,
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
                  text: "\u25A0",
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
          ],
        },
        sections: [
          {
            properties: {
              page: {
                size: {
                  width: 12240, // A4宽度
                  height: 15840, // A4高度
                },
                orientation: "portrait",
                margin: {
                  top: 1440,
                  bottom: 1440,
                  left: 1800,
                  right: 1800,
                  gutter: 0,
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
      console.error('直接转换过程中出错:', error);
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
      case 'heading':
        return this.createHeadingFromMarked(token);
      case 'paragraph':
        return this.createParagraphFromMarked(token);
      case 'code':
        return this.createCodeBlockFromMarked(token);
      case 'list':
        return this.createListFromMarked(token);
      case 'table':
        return this.createTableFromMarked(token);
      case 'blockquote':
        return this.createBlockquoteFromMarked(token);
      case 'hr':
        return this.createHorizontalRule();
      case 'image':
        return this.createImageFromMarked(token);
      case 'space':
        return this.createSpaceFromMarked();
      default:
        console.warn(`未支持的token类型: ${token.type}`);
        return null;
    }
  }

  // 创建空行
  createSpaceFromMarked() {
    return new Paragraph({
      text: '',
      alignment: AlignmentType.START
    });
  }

  /**
   * 从marked token创建标题
   * @param {Object} token - marked解析的标题token
   * @returns {Object} docx Paragraph对象
   */
  createHeadingFromMarked(token) {
    const level = token.depth;

    // 使用正则表达式匹配序号和文本内容
    const match = token.text.match(/^(\d+)\.(.*)/);
    const hasNumber = match !== null;
    const number = hasNumber ? match[1] : null;
    const content = hasNumber ? match[2].trim() : token.text;

    // 准备TextRun数组
    const textRuns = [];

    // 根据标题级别设置不同的字体样式
    let fontSize, fontFamily, isBold;

    switch (level) {
      case 1: // 标题
        fontSize = 44; // 2号
        fontFamily = "黑体";
        isBold = true;
        break;
      case 2: // 一级标题
        fontSize = 32; // 三号
        fontFamily = "黑体";
        isBold = true;
        break;
      case 3: // 二级标题
        fontSize = 32; // 16磅
        fontFamily = "楷体";
        isBold = true;
        break;
      case 4: // 三级标题
        fontSize = 28; // 14磅
        fontFamily = "仿宋";
        isBold = true;
        break;
      case 5: // 四级标题
        fontSize = 24; // 12磅
        fontFamily = "仿宋";
        isBold = true;
        break;
      case 6: // 五级标题
        fontSize = 21; // 10.5磅
        fontFamily = "仿宋";
        isBold = true;
        break;
      default:
        fontSize = 32;
        fontFamily = "黑体";
        isBold = true;
    }

    textRuns.push(
      new TextRun({
        text: content,
        size: fontSize,
        font: fontFamily,
        bold: isBold,
      })
    );

    // 设置标题级别
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

    // 判断是否需要编号
    if (hasNumber) {
      return new Paragraph({
        numbering: {
          reference: "my-heading-style",
          level: level,
        },
        heading: headingLevel,
        children: textRuns,
      });
    } else {
      return new Paragraph({
        heading: headingLevel,
        children: textRuns,
        style: `heading${level}`,
      });
    }
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
    const cleanText = hasNumber ? token.text.replace(/^\d+\.\s*/, '') : token.text;

    const paragraph = new Paragraph({
      text: cleanText,
      alignment: AlignmentType.JUSTIFIED
    });

    // 如果有序号，添加编号
    if (hasNumber) {
      paragraph.numbering = {
        reference: "my-paragraph-style",
        level: 0
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
        left: 720
      }
    });
  }

  /**
   * 从marked token创建列表
   * @param {Object} token - marked解析的列表token
   * @returns {Array} docx Paragraph对象数组
   */
  createListFromMarked(token) {
    const paragraphs = [];
    token.items.forEach(item => {
      paragraphs.push(new Paragraph({
        text: item.text,
        bullet: {
          level: 0
        }
      }));
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
      children: token.header.map(cell => new TableCell({
        children: [new Paragraph({ text: cell.text })]
      }))
    });
    rows.push(headerRow);

    // 添加数据行
    token.rows.forEach(row => {
      const tableRow = new TableRow({
        children: row.map(cell => new TableCell({
          children: [new Paragraph({ text: cell.text })]
        }))
      });
      rows.push(tableRow);
    });

    return new Table({
      rows: rows,
      width: {
        size: 100,
        type: WidthType.PERCENTAGE
      }
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
        left: 720
      },
      border: {
        left: {
          color: "CCCCCC",
          size: 4,
          style: BorderStyle.SINGLE
        }
      }
    });
  }

  /**
   * 创建水平线
   * @returns {Object} docx Paragraph对象
   */
  createHorizontalRule() {
    return new Paragraph({
      text: "──────────────────────────────────────",
      alignment: AlignmentType.CENTER
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
      alignment: AlignmentType.CENTER
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
      if (token.type === 'image') {
        try {
          const imageData = await this.loadImage(token.href);
          imageInfos.push({
            ...token,
            ...imageData
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

          const blob = await new Promise(resolve => canvas.toBlob(resolve));
          const buffer = await blob.arrayBuffer();

          resolve({
            buffer,
            width: img.naturalWidth,
            height: img.naturalHeight
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
              color: "999999"
            })
          ]
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
      case 1: headingLevel = HeadingLevel.HEADING_1; break;
      case 2: headingLevel = HeadingLevel.HEADING_2; break;
      case 3: headingLevel = HeadingLevel.HEADING_3; break;
      case 4: headingLevel = HeadingLevel.HEADING_4; break;
      case 5: headingLevel = HeadingLevel.HEADING_5; break;
      case 6: headingLevel = HeadingLevel.HEADING_6; break;
      default: headingLevel = HeadingLevel.HEADING_1;
    }

    // 处理内联样式
    if (element.inlineStyles && Array.isArray(element.inlineStyles) && element.inlineStyles.length > 0) {
      const children = [];

      element.inlineStyles.forEach(style => {
        if (style.content) {
          children.push(new TextRun({
            text: style.content,
            bold: style.bold === true,
            italic: style.italic === true,
            strike: style.strike === true,
            underline: style.underline === true ? {} : undefined,
            superScript: style.superscript === true,
            subScript: style.subscript === true
          }));
        }
      });

      if (children.length > 0) {
        return new Paragraph({
          heading: headingLevel,
          children
        });
      }
    }

    // 提取标题文本（如果没有可用的内联样式）
    const titleText = element.fullContent || element.rawText || element.text || '';
    // 去掉Markdown的#符号
    const cleanTitle = titleText.replace(/^#+\s+/, '');

    console.log(`创建标题: '${cleanTitle}', 级别: ${level}`);

    // 创建标题段落
    return new Paragraph({
      text: cleanTitle,
      heading: headingLevel
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
    if (element.inlineStyles && Array.isArray(element.inlineStyles) && element.inlineStyles.length > 0) {
      const children = [];

      element.inlineStyles.forEach(style => {
        if (style.content) {
          children.push(new TextRun({
            text: style.content,
            bold: style.bold === true,
            italic: style.italic === true,
            strike: style.strike === true,
            underline: style.underline === true ? {} : undefined,
            superScript: style.superscript === true,
            subScript: style.subscript === true
          }));
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
    const text = element.rawText || element.text || '';
    return new Paragraph({ text });
  }

  /**
   * 创建代码块
   * @param {Object} element - 代码块元素
   * @returns {Object} docx Paragraph对象
   */
  createCodeBlock(element) {
    console.log("创建代码块:", element);

    const code = element.fullContent || element.content || element.text || '';
    const language = element.language || '';

    return new Paragraph({
      children: [
        new TextRun({
          text: `${language ? `[${language}] ` : ''}${code}`,
          font: "Courier New"
        })
      ],
      indent: {
        left: 720 // 缩进量，720 = 0.5英寸
      }
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
      const prefix = isOrdered ? `${index + 1}. ` : '• ';
      const itemText = item.fullContent || item.text || '';

      result.push(new Paragraph({
        text: `${prefix}${itemText}`,
        indent: {
          left: 720 // 缩进量
        }
      }));

      // 处理嵌套列表
      if (item.children && Array.isArray(item.children) && item.children.length > 0) {
        item.children.forEach(child => {
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
    if (!element.data || !Array.isArray(element.data) || element.data.length === 0) {
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
        const cellText = cellData.fullContent || cellData.text || '';

        cells.push(new TableCell({
          children: [new Paragraph({ text: cellText })]
        }));
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
        type: WidthType.PERCENTAGE
      }
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
          text: `[图片: ${altText}]${url ? ` ${url}` : ''}`,
          italic: true,
          color: "0000FF"
        })
      ],
      alignment: AlignmentType.CENTER
    });
  }

  /**
   * 创建引用块
   * @param {Object} element - 引用块元素
   * @returns {Object} docx Paragraph对象
   */
  createBlockquote(element) {
    console.log("创建引用块:", element);

    const text = element.fullContent || element.text || '';

    return new Paragraph({
      text: text,
      indent: {
        left: 720 // 缩进量
      },
      border: {
        left: {
          color: "CCCCCC",
          size: 4,
          style: BorderStyle.SINGLE
        }
      },
      spacing: {
        before: 240, // 前间距
        after: 240   // 后间距
      }
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
                    size: 36 // 18pt
                  })
                ],
                alignment: AlignmentType.CENTER
              }),
              new Paragraph({
                text: errorMessage,
                alignment: AlignmentType.CENTER
              }),
              new Paragraph({
                text: `生成时间: ${new Date().toLocaleString()}`,
                alignment: AlignmentType.CENTER
              })
            ]
          }
        ]
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
                text: `错误: ${errorMessage}`
              })
            ]
          }
        ]
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
