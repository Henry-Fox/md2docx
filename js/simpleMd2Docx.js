// 引入需要的库
import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType, BorderStyle } from "docx";
import { saveAs } from "file-saver";
import { Md2Json } from './md2json.js';
import runTest from './json2docx.js';

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
   * 创建水平线
   * @returns {Object} docx Paragraph对象
   */
  createHorizontalRule() {
    console.log("创建水平线");

    return new Paragraph({
      text: "──────────────────────────────────────",
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
