import { Document, Paragraph, TextRun, HeadingLevel, AlignmentType, BorderStyle, Table, TableRow, TableCell, Packer, WidthType, TableOfContents, NumberingLevel, LevelFormat, convertInchesToTwip, Tab, TabStopType, TabStopPosition, HorizontalPositionAlign, TableLayoutType, ImageRun, ExternalHyperlink } from 'docx';
import { marked } from './marked.esm.js';
import { saveAs } from 'file-saver';
import defaultStyles from '../styles/default-styles.json';

/**
 * @class Md2Docx
 * @description Markdown转Word文档的转换器类
 */
class Md2Docx {
  /**
   * @constructor
   * @param {Object} customStyles - 自定义样式配置
   */
  constructor(customStyles = {}) {
    // 合并样式设置，先使用默认样式，再覆盖自定义样式
    this.styles = this.mergeStyles(defaultStyles, customStyles);
    this.doc = null;
    this.footnotes = new Map(); // 存储脚注的集合

    // 为有序列表和无序列表创建编号定义
    const listStyles = this.styles.list || {};
    const unorderedListStyles = listStyles.unordered || {};
    const orderedListStyles = listStyles.ordered || {};

    this.bulletListNumbering = {
      reference: "bulletList",
      levels: [
        {
          level: 0,
          format: LevelFormat.BULLET,
          text: unorderedListStyles.bulletChars?.[0] || "●",
          alignment: AlignmentType.LEFT,
          style: {
            paragraph: {
              indent: { left: unorderedListStyles.indentLevel || 720, hanging: 360 }
            },
            run: {
              font: unorderedListStyles.font || this.styles.paragraph.font
            }
          }
        },
        {
          level: 1,
          format: LevelFormat.BULLET,
          text: unorderedListStyles.bulletChars?.[1] || "○",
          alignment: AlignmentType.LEFT,
          style: {
            paragraph: {
              indent: { left: (unorderedListStyles.indentLevel || 720) * 2, hanging: 360 }
            },
            run: {
              font: unorderedListStyles.font || this.styles.paragraph.font
            }
          }
        },
        {
          level: 2,
          format: LevelFormat.BULLET,
          text: unorderedListStyles.bulletChars?.[2] || "■",
          alignment: AlignmentType.LEFT,
          style: {
            paragraph: {
              indent: { left: (unorderedListStyles.indentLevel || 720) * 3, hanging: 360 }
            },
            run: {
              font: unorderedListStyles.font || this.styles.paragraph.font
            }
          }
        }
      ]
    };

    this.orderedListNumbering = {
      reference: "orderedList",
      levels: [
        {
          level: 0,
          format: LevelFormat.DECIMAL,
          text: orderedListStyles.numberFormats?.[0] || "%1.",
          alignment: AlignmentType.LEFT,
          style: {
            paragraph: {
              indent: { left: orderedListStyles.indentLevel || 720, hanging: 360 }
            },
            run: {
              font: orderedListStyles.font || this.styles.paragraph.font
            }
          }
        },
        {
          level: 1,
          format: LevelFormat.DECIMAL,
          text: orderedListStyles.numberFormats?.[1] || "%2.",
          alignment: AlignmentType.LEFT,
          style: {
            paragraph: {
              indent: { left: (orderedListStyles.indentLevel || 720) * 2, hanging: 360 }
            },
            run: {
              font: orderedListStyles.font || this.styles.paragraph.font
            }
          }
        },
        {
          level: 2,
          format: LevelFormat.DECIMAL,
          text: orderedListStyles.numberFormats?.[2] || "%3.",
          alignment: AlignmentType.LEFT,
          style: {
            paragraph: {
              indent: { left: (orderedListStyles.indentLevel || 720) * 3, hanging: 360 }
            },
            run: {
              font: orderedListStyles.font || this.styles.paragraph.font
            }
          }
        }
      ]
    };
  }

  /**
   * @method mergeStyles
   * @description 深度合并两个样式对象
   * @param {Object} target - 目标对象
   * @param {Object} source - 源对象
   * @returns {Object} 合并后的对象
   */
  mergeStyles(target, source) {
    if (!source) return target;
    const result = { ...target };

    for (const key in source) {
      if (Object.prototype.hasOwnProperty.call(source, key)) {
        if (
          source[key] &&
          typeof source[key] === 'object' &&
          !Array.isArray(source[key]) &&
          target[key] &&
          typeof target[key] === 'object' &&
          !Array.isArray(target[key])
        ) {
          result[key] = this.mergeStyles(target[key], source[key]);
        } else {
          result[key] = source[key];
        }
      }
    }

    return result;
  }

  /**
   * @method getDefaultStyles
   * @returns {Object} 默认样式对象
   */
  getDefaultStyles() {
    try {
      // 直接返回defaultStyles模块引入的对象
      // 注意：在使用webpack等工具时，这会被正确导入
      const styles = JSON.parse(JSON.stringify(defaultStyles));

      // 输出加载的样式用于调试
      console.log("从JSON文件加载的默认样式:", styles);

      return styles;
    } catch (error) {
      console.error('加载默认样式出错:', error);

      // 返回硬编码的基本样式，确保应用不会崩溃
      return {
        document: {
          pageSize: "A4",
          pageOrientation: "portrait",
          margins: {
            top: 2099,
            right: 1474,
            bottom: 1984,
            left: 1587
          }
        },
        heading: {
          font: "方正小标宋简体",
          color: "#000000",
          colors: {
            h1: "#000000",
            h2: "#000000",
            h3: "#000000",
            h4: "#000000",
            h5: "#000000",
            h6: "#000000"
          },
          bold: {
            h1: false,
            h2: true,
            h3: false,
            h4: false,
            h5: false,
            h6: false
          },
          sizes: {
            h1: 22,
            h2: 16,
            h3: 16,
            h4: 16,
            h5: 16,
            h6: 10.5
          },
          alignment: {
            h1: "center",
            h2: "left",
            h3: "left",
            h4: "left",
            h5: "left",
            h6: "left"
          },
          fonts: {
            h1: "方正小标宋简体",
            h2: "黑体",
            h3: "楷体",
            h4: "仿宋_GB2312",
            h5: "仿宋_GB2312",
            h6: "仿宋_GB2312"
          },
          indent: {
            h1: 0,
            h2: 0,
            h3: 0,
            h4: 800,
            h5: 800,
            h6: 0
          },
          prefix: {
            h1: "",
            h2: "一、",
            h3: "(一)",
            h4: "1.",
            h5: "(1)",
            h6: ""
          },
          usePrefix: {
            h1: false,
            h2: true,
            h3: true,
            h4: true,
            h5: true,
            h6: false
          }
        },
        paragraph: {
          font: "仿宋_GB2312",
          size: 16,
          color: "#000000",
          firstLineIndent: 800,
          alignment: "justified",
          lineSpacingRule: "auto",
          lineSpacing: 1.5,
          spacing: 0
        }
      };
    }
  }

  /**
   * @method setStyles
   * @description 设置文档样式
   * @param {Object} styles - 样式对象
   */
  setStyles(styles) {
    this.styles = { ...styles };
  }

  /**
   * @method convert
   * @param {string} markdown - Markdown文本内容
   * @returns {Document} - 返回docx文档对象
   */
  convert(markdown) {
    // 获取页面设置
    const docStyles = this.styles.document;
    const headingStyles = this.styles.heading;
    // 直接使用paragraph样式
    const paragraphStyles = this.styles.paragraph || {};
    // 获取表格样式
    const tableStyles = this.styles.table || {};

    // 处理颜色值，确保没有#符号
    const paragraphColor = paragraphStyles.color.replace('#', '');
    const tableBorderColor = (tableStyles.borderColor || "000000").replace('#', '');

    // 处理表格默认设置
    if (!this.styles.table) {
      this.styles.table = {
        borderColor: "000000",
        borderWidth: 1,
        headerBackground: "E6E6E6",
        headerFont: "仿宋_GB2312",
        fontSize: 16,
        alignment: "center"
      };
    }

    // 定义标题样式配置
    const headingStyleDefinitions = [];

    // 处理行距设置
    // 如果是固定行距（磅值），直接使用数值；如果是倍数行距，则需要乘以240转换
    const lineSpacingValue = paragraphStyles.lineSpacingRule === 'exact' ?
                            paragraphStyles.lineSpacing * 20 : // 磅值转twip (1磅 = 20 twip)
                            paragraphStyles.lineSpacing * 240; // 倍数行距

    const lineSpacingRule = paragraphStyles.lineSpacingRule === 'exact' ?
                           'exact' : 'auto';

    // 主样式 - Normal
    const normalStyle = {
      id: "Normal",
      name: "Normal",
      run: {
        size: paragraphStyles.size * 2,
        font: paragraphStyles.font,
        color: paragraphColor
      },
      paragraph: {
        spacing: {
          line: lineSpacingValue,
          lineRule: lineSpacingRule,
          before: 0,
          after: paragraphStyles.spacing || 0
        },
        indent: {
          firstLine: paragraphStyles.firstLineIndent || 800
        },
        alignment: this.getAlignmentType(paragraphStyles.alignment) || AlignmentType.JUSTIFIED
      }
    };

    // 为每个标题级别创建样式定义
    for (let i = 1; i <= 6; i++) {
      const headingKey = `h${i}`;

      // 获取该标题级别的特定字体设置
      const headingFont = headingStyles.fonts?.[headingKey] || headingStyles.font;

      // 获取该标题级别的特定颜色设置
      const headingColor = headingStyles.colors?.[headingKey] || headingStyles.color;
      const headingColorValue = typeof headingColor === 'string' ? headingColor.replace('#', '') : '000000';

      // 获取该标题级别的特定加粗设置
      const isBold = headingStyles.bold?.[headingKey] !== undefined
                   ? headingStyles.bold[headingKey]
                   : (i === 1 || headingStyles.bold === true);

      // 获取对齐方式
      const alignmentSetting = headingStyles.alignment?.[headingKey] || 'left';
      const alignment = this.getAlignmentType(alignmentSetting);

      // 获取缩进设置
      const leftIndent = headingStyles.indent?.[headingKey] || 0;

      // 创建标题样式定义
      const headingStyle = {
        id: `Heading${i}`,
        name: `Heading ${i}`,
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: {
          size: headingStyles.sizes[headingKey] * 2,
          bold: isBold,
          font: headingFont,
          color: headingColorValue
        },
        paragraph: {
          spacing: {
            before: i === 1 ? 240 : 200 - (i-2) * 20, // 根据级别递减间距
            after: i === 1 ? 240 : 120 - (i-2) * 20,
            line: lineSpacingValue,
            lineRule: lineSpacingRule
          },
          indent: {
            firstLine: 0, // 标题不缩进首行
            left: leftIndent // 应用左侧缩进
          },
          alignment: alignment
        }
      };

      headingStyleDefinitions.push(headingStyle);
    }

    // 创建一个符合公文格式标准的文档结构
    this.doc = new Document({
      features: {
        update: {
          fields: true
        }
      },
      styles: {
        default: {
          document: {
            run: {
              font: paragraphStyles.font,
              size: paragraphStyles.size * 2,
              color: paragraphColor
            },
            paragraph: {
              spacing: {
                line: lineSpacingValue,
                lineRule: lineSpacingRule,
                before: 0,
                after: paragraphStyles.spacing || 0
              },
              alignment: this.getAlignmentType(paragraphStyles.alignment) || AlignmentType.JUSTIFIED
            }
          }
        },
        paragraphStyles: [
          normalStyle,
          ...headingStyleDefinitions
        ]
      },
      sections: [{
        properties: {
          page: {
            size: {
              width: 11906, // A4纸宽度
              height: 16838 // A4纸高度
            },
            margin: {
              top: docStyles.margins.top,
              right: docStyles.margins.right,
              bottom: docStyles.margins.bottom,
              left: docStyles.margins.left,
            },
            orientation: docStyles.pageOrientation
          },
          // 如果定义了文档网格，则设置
          docGrid: docStyles.grid ? {
            linePitch: 360, // 行网格基本单位
            charSpace: Math.floor(360 / (docStyles.grid.charPerLine || 28)) // 字符网格间距，根据每行字符数计算
          } : undefined
        },
        children: this.parseMarkdown(markdown),
      }],
      // 编号引用配置
      numbering: {
        config: [
          {
            reference: "bulletList",
            levels: [
              {
                level: 0,
                format: LevelFormat.BULLET,
                text: this.styles.list?.unordered?.bulletChars?.[0] || "●",
                alignment: AlignmentType.LEFT,
                style: {
                  paragraph: {
                    indent: { left: this.styles.list?.unordered?.indentLevel || 800, hanging: 400 }
                  },
                  run: {
                    font: this.styles.list?.unordered?.font || paragraphStyles.font
                  }
                }
              },
              {
                level: 1,
                format: LevelFormat.BULLET,
                text: this.styles.list?.unordered?.bulletChars?.[1] || "○",
                alignment: AlignmentType.LEFT,
                style: {
                  paragraph: {
                    indent: { left: (this.styles.list?.unordered?.indentLevel || 800) * 2, hanging: 400 }
                  },
                  run: {
                    font: this.styles.list?.unordered?.font || paragraphStyles.font
                  }
                }
              },
              {
                level: 2,
                format: LevelFormat.BULLET,
                text: this.styles.list?.unordered?.bulletChars?.[2] || "■",
                alignment: AlignmentType.LEFT,
                style: {
                  paragraph: {
                    indent: { left: (this.styles.list?.unordered?.indentLevel || 800) * 3, hanging: 400 }
                  },
                  run: {
                    font: this.styles.list?.unordered?.font || paragraphStyles.font
                  }
                }
              }
            ]
          },
          {
            reference: "orderedList",
            levels: [
              {
                level: 0,
                format: LevelFormat.DECIMAL,
                text: this.styles.list?.ordered?.numberFormats?.[0] || "%1、",
                alignment: AlignmentType.LEFT,
                style: {
                  paragraph: {
                    indent: { left: this.styles.list?.ordered?.indentLevel || 800, hanging: 400 }
                  },
                  run: {
                    font: this.styles.list?.ordered?.font || paragraphStyles.font
                  }
                }
              },
              {
                level: 1,
                format: LevelFormat.DECIMAL,
                text: this.styles.list?.ordered?.numberFormats?.[1] || "（%2）",
                alignment: AlignmentType.LEFT,
                style: {
                  paragraph: {
                    indent: { left: (this.styles.list?.ordered?.indentLevel || 800) * 2, hanging: 400 }
                  },
                  run: {
                    font: this.styles.list?.ordered?.font || paragraphStyles.font
                  }
                }
              },
              {
                level: 2,
                format: LevelFormat.DECIMAL,
                text: this.styles.list?.ordered?.numberFormats?.[2] || "%3.",
                alignment: AlignmentType.LEFT,
                style: {
                  paragraph: {
                    indent: { left: (this.styles.list?.ordered?.indentLevel || 800) * 3, hanging: 400 }
                  },
                  run: {
                    font: this.styles.list?.ordered?.font || paragraphStyles.font
                  }
                }
              }
            ]
          }
        ]
      }
    });
    return this.doc;
  }

  /**
   * @method parseMarkdown
   * @param {string} markdown - Markdown文本内容
   * @returns {Array} - docx段落对象数组
   */
  parseMarkdown(markdown) {
    // 配置marked解析选项
    marked.setOptions({
      gfm: true, // 启用GitHub风格Markdown
      breaks: true, // 启用换行符
      tables: true, // 启用表格
      smartLists: true, // 智能列表
      xhtml: false, // 不使用XHTML
      silent: true // 忽略解析错误
    });

    // 解析Markdown标记
    const tokens = marked.lexer(markdown);

    // 预处理复杂表格数据
    tokens.forEach(token => {
      if (token.type === 'table') {
        // 检查表格单元格是否包含JSON字符串
        if (token.header && token.header.length > 0) {
          const firstHeader = token.header[0];
          if (typeof firstHeader === 'string' && firstHeader.includes('{"text":')) {
            // 需要预处理
            console.log('检测到JSON格式表格数据，预处理中...');
          }
        }
      }
    });

    // 预处理，提取所有脚注定义
    this.extractFootnotes(tokens);

    const elements = [];
    for (const token of tokens) {
      // 跳过已经处理过的脚注定义
      if (token.type === 'footnoteDefinition') {
        continue;
      }

      const el = this.convertToken(token);
      if (Array.isArray(el)) {
        elements.push(...el);
      } else if (el) {
        elements.push(el);
      }
    }

    // 如果有脚注，则添加脚注部分
    if (this.footnotes.size > 0) {
      // 添加一个水平线作为脚注分隔符
      elements.push(this.createHorizontalRule());

      // 添加脚注标题
      elements.push(new Paragraph({
        text: "脚注",
        style: "Heading3"
      }));

      // 添加所有脚注内容
      const footnoteLabels = Array.from(this.footnotes.keys()).sort();
      for (const label of footnoteLabels) {
        const content = this.footnotes.get(label);
        elements.push(this.createFootnote({
          label: label,
          text: content
        }));
      }
    }

    return elements;
  }

  /**
   * @method extractFootnotes
   * @description 从标记中提取所有脚注定义
   * @param {Array} tokens - Markdown标记数组
   */
  extractFootnotes(tokens) {
    // 清空脚注集合
    this.footnotes.clear();

    // 提取脚注定义
    for (const token of tokens) {
      if (token.type === 'footnoteDefinition') {
        this.footnotes.set(token.label, token.text);
      }
    }
  }

  /**
   * @method createInlineFootnoteReference
   * @description 创建内联脚注引用
   * @param {string} label - 脚注标签
   * @returns {TextRun} 文本运行对象
   */
  createInlineFootnoteReference(label) {
    const footnoteStyles = this.styles.footnote || {};
    const paragraphStyles = this.styles.paragraph;

    return new TextRun({
      text: `[${label}]`,
      size: (footnoteStyles.size || paragraphStyles.size - 3) * 2,
      superScript: true,
      color: (footnoteStyles.color || paragraphStyles.color).replace('#', '')
    });
  }

  /**
   * @method processTableData
   * @description 处理可能是JSON格式的表格数据
   * @param {Array|String} data - 表格数据，可能是字符串或数组
   * @returns {Array} 处理后的数据数组
   */
  processTableData(data) {
    if (!data) return [];

    // 如果已经是数组，处理数组中的每个元素
    if (Array.isArray(data)) {
      return data.map(item => this.processTableCell(item));
    }

    // 如果是字符串，可能是制表符分隔的单元格数据
    if (typeof data === 'string') {
      // 检查是否包含制表符，如果是则按制表符拆分
      if (data.includes('\t')) {
        return data.split('\t').map(cell => this.processTableCell(cell));
      }

      // 单个单元格处理
      return [this.processTableCell(data)];
    }

    return [];
  }

  /**
   * @method processTableCell
   * @description 处理单个表格单元格数据
   * @param {any} cell - 单元格数据
   * @returns {String} 处理后的单元格文本
   */
  processTableCell(cell) {
    // 如果是null或undefined，返回空字符串
    if (cell === null || cell === undefined) {
      return '';
    }

    // 如果是字符串类型
    if (typeof cell === 'string') {
      // 尝试解析JSON格式
      if (cell.startsWith('{') && cell.includes('"text"')) {
        try {
          const parsed = JSON.parse(cell);
          return parsed.text || '';
        } catch (e) {
          // 解析失败，返回原始字符串
          return cell;
        }
      }
      return cell;
    }

    // 如果是对象类型，尝试获取text属性
    if (typeof cell === 'object') {
      return cell.text || '';
    }

    // 其他类型转为字符串
    return String(cell);
  }

  /**
   * @method convertToken
   * @param {Object} token - markdown token
   * @returns {Paragraph|Array<Paragraph>|Table|null}
   */
  convertToken(token) {
    switch (token.type) {
      case 'heading':
        return this.createHeading(token);
      case 'paragraph':
        // 检查段落是否只包含图片
        if (token.text && token.text.match(/^!\[.*?\]\(.*?\)$/)) {
          // 提取图片信息
          const match = token.text.match(/^!\[(.*?)\]\((.*?)\)$/);
          if (match) {
            const [_, altText, url] = match;
            // 创建图片token并调用图片处理方法
            return this.createImagePlaceholder({
              type: 'image',
              text: altText,
              href: url
            });
          }
        }
        return this.createParagraph(token);
      case 'list':
        // 检查是否是任务列表
        if (token.items && token.items[0] &&
            (token.items[0].text.startsWith('[x]') || token.items[0].text.startsWith('[ ]'))) {
          return this.createTaskList(token);
        }
        return this.createList(token, 0); // 从0级开始
      case 'code':
        return this.createCode(token);
      case 'table':
        // 预处理表格数据
        if (token.header && typeof token.header === 'string' && token.header.includes('\t')) {
          // 处理制表符分隔的表格数据
          const rows = token.text.split('\n').filter(row => row.trim() !== '');
          if (rows.length > 0) {
            token.header = this.processTableData(rows[0]);
            token.rows = rows.slice(1).map(row => this.processTableData(row));
          }
        }
        return this.createTable(token);
      case 'blockquote':
        return this.createBlockquote(token);
      case 'hr':
        return this.createHorizontalRule();
      case 'space':
        return new Paragraph({});
      case 'image':
        return this.createImagePlaceholder(token);
      case 'footnote':
        return this.createFootnote(token);
      case 'link':
        // 使用ExternalHyperlink创建超链接
        const linkText = new TextRun({
          text: token.text,
          ...baseProps,
          color: linkStyles.color?.replace('#', '') || '0066CC',
          underline: linkStyles.underline || {},
        });

        runs.push(new ExternalHyperlink({
          children: [linkText],
          link: token.url
        }));
        break;
      default:
        // 如果是未知类型，尝试作为普通文本处理
        if (token.text) {
          return this.createParagraph({ type: 'paragraph', text: token.text });
        }
        return null;
    }
  }

  /**
   * @method parseInlineContent
   * @param {string} text - 要解析的内联文本
   * @returns {Array<TextRun>} - TextRun数组
   */
  parseInlineContent(text) {
    // 确保输入是字符串类型
    if (typeof text !== 'string') {
      text = text?.toString() || '';
    }

    // 获取正文样式
    const paragraphStyles = this.styles.paragraph || {};
    const textStyles = this.styles.textStyles || {};
    const linkStyles = this.styles.link || {};
    const codeStyles = this.styles.code || {};

    // 标准化换行符
    text = text.replace(/\r\n/g, '\n');

    // 使用更复杂的解析方法来处理各种文本格式
    const tokens = this.tokenizeInlineText(text);
    const runs = [];

    tokens.forEach(token => {
      const baseProps = {
        size: paragraphStyles.size * 2,
        font: paragraphStyles.font,
        color: paragraphStyles.color.replace('#', '')
      };

      switch (token.type) {
        case 'text':
          runs.push(new TextRun({
            text: token.text,
            ...baseProps
          }));
          break;

        case 'bold':
          runs.push(new TextRun({
            text: token.text,
            ...baseProps,
            bold: true
          }));
          break;

        case 'italic':
          runs.push(new TextRun({
            text: token.text,
            ...baseProps,
            italics: true
          }));
          break;

        case 'bold_italic':
          runs.push(new TextRun({
            text: token.text,
            ...baseProps,
            bold: true,
            italics: true
          }));
          break;

        case 'strike':
          runs.push(new TextRun({
            text: token.text,
            ...baseProps,
            strike: true
          }));
          break;

        case 'underline':
          runs.push(new TextRun({
            text: token.text,
            ...baseProps,
            underline: {}
          }));
          break;

        case 'code':
          runs.push(new TextRun({
            text: token.text,
            size: codeStyles.size * 2,
            font: codeStyles.font,
            color: codeStyles.color.replace('#', '')
          }));
          break;

        case 'link':
          // 使用ExternalHyperlink创建超链接
          const linkText = new TextRun({
            text: token.text,
            ...baseProps,
            color: linkStyles.color?.replace('#', '') || '0066CC',
            underline: linkStyles.underline || {},
          });

          runs.push(new ExternalHyperlink({
            children: [linkText],
            link: token.url
          }));
          break;

        case 'footnote_ref':
          // 处理脚注引用
          const footnoteStyles = this.styles.footnote || {};
          runs.push(new TextRun({
            text: `[${token.label}]`,
            size: (footnoteStyles.size || paragraphStyles.size - 3) * 2,
            superScript: true,
            color: (footnoteStyles.color || paragraphStyles.color).replace('#', '')
          }));
          break;
      }
    });

    return runs.length > 0 ? runs : [
      new TextRun({
        text: text,
        size: paragraphStyles.size * 2,
        font: paragraphStyles.font,
        color: paragraphStyles.color.replace('#', '')
      })
    ];
  }

  /**
   * @method tokenizeInlineText
   * @description 将内联文本分解为格式化的标记
   * @param {string} text - 要分解的文本
   * @returns {Array<Object>} - 标记数组
   */
  tokenizeInlineText(text) {
    const tokens = [];
    let tempText = text;

    // 正则表达式匹配不同的格式
    const boldItalicRegex = /\*\*\*(.*?)\*\*\*/g;
    const boldRegex = /\*\*(.*?)\*\*/g;
    const italicRegex = /\*(.*?)\*/g;
    const strikeRegex = /~~(.*?)~~/g;
    const underlineRegex = /<u>(.*?)<\/u>/g;
    const codeRegex = /`(.*?)`/g;
    const linkRegex = /\[(.*?)\]\((.*?)\)/g;
    const footnoteRefRegex = /\[\^(.*?)\]/g;  // 脚注引用格式 [^1]

    // 为了跟踪已处理的部分，创建一个标记数组
    const allTokens = [];

    // 函数：添加标记并在tempText中标记为已处理
    function addToken(match, type, text, url = null) {
      const start = match.index;
      const end = start + match[0].length;

      allTokens.push({
        type,
        text: text,
        url: url,
        start,
        end,
        original: match[0]
      });

      // 在tempText中用空格替换已处理的部分
      // 这样后续的匹配不会影响这部分内容
      tempText = tempText.substring(0, start) + ' '.repeat(match[0].length) + tempText.substring(end);
    }

    // 按顺序处理不同的格式化标记
    // 粗斜体
    let match;
    while ((match = boldItalicRegex.exec(text)) !== null) {
      addToken(match, 'bold_italic', match[1]);
    }

    // 粗体
    boldRegex.lastIndex = 0; // 重置正则索引
    while ((match = boldRegex.exec(text)) !== null) {
      // 检查这部分是否已经被处理（在粗斜体处理中）
      const alreadyProcessed = allTokens.some(token =>
        token.start <= match.index && token.end >= match.index + match[0].length
      );

      if (!alreadyProcessed) {
        addToken(match, 'bold', match[1]);
      }
    }

    // 斜体
    italicRegex.lastIndex = 0;
    while ((match = italicRegex.exec(text)) !== null) {
      const alreadyProcessed = allTokens.some(token =>
        token.start <= match.index && token.end >= match.index + match[0].length
      );

      if (!alreadyProcessed) {
        addToken(match, 'italic', match[1]);
      }
    }

    // 删除线
    strikeRegex.lastIndex = 0;
    while ((match = strikeRegex.exec(text)) !== null) {
      addToken(match, 'strike', match[1]);
    }

    // 下划线
    underlineRegex.lastIndex = 0;
    while ((match = underlineRegex.exec(text)) !== null) {
      addToken(match, 'underline', match[1]);
    }

    // 代码
    codeRegex.lastIndex = 0;
    while ((match = codeRegex.exec(text)) !== null) {
      addToken(match, 'code', match[1]);
    }

    // 链接 - 特别处理为链接类型
    linkRegex.lastIndex = 0;
    while ((match = linkRegex.exec(text)) !== null) {
      addToken(match, 'link', match[1], match[2]);
    }

    // 脚注引用
    footnoteRefRegex.lastIndex = 0;
    while ((match = footnoteRefRegex.exec(text)) !== null) {
      addToken(match, 'footnote_ref', match[1]);
    }

    // 排序标记，确保按文本顺序处理
    allTokens.sort((a, b) => a.start - b.start);

    // 处理未标记的文本
    // 从原始文本开始，逐段添加标记和未标记的文本
    let lastEnd = 0;
    const finalTokens = [];

    for (const token of allTokens) {
      // 如果当前标记前有未处理的文本，添加为普通文本
      if (token.start > lastEnd) {
        finalTokens.push({
          type: 'text',
          text: text.substring(lastEnd, token.start)
        });
      }

      // 添加当前标记
      finalTokens.push({
        type: token.type,
        text: token.text,
        url: token.url
      });

      lastEnd = token.end;
    }

    // 处理最后一个标记后的文本
    if (lastEnd < text.length) {
      finalTokens.push({
        type: 'text',
        text: text.substring(lastEnd)
      });
    }

    return finalTokens.length > 0 ? finalTokens : [{ type: 'text', text }];
  }

  /**
   * @method createHeading
   * @param {Object} token - 标题token
   * @returns {Paragraph}
   */
  createHeading(token) {
    // 确保文本是字符串
    if (typeof token.text !== 'string') {
      token.text = token.text?.toString() || '';
    }

    const headingStyles = this.styles.heading;

    // 确定标题级别
    const level = token.depth || 1;
    const headingKey = `h${level}`;

    // 获取字体和颜色
    const headingFont = headingStyles.fonts?.[headingKey] || headingStyles.font;

    // 获取颜色 - 优先使用级别特定颜色，如果没有则使用通用颜色
    let headingColor = headingStyles.colors?.[headingKey] || headingStyles.color;
    if (typeof headingColor === 'string') {
      headingColor = headingColor.replace('#', '');
    }

    // 获取标题加粗设置
    const isBold = headingStyles.bold?.[headingKey] !== undefined
                 ? headingStyles.bold[headingKey]
                 : (level === 1 || headingStyles.bold === true);

    // 获取对齐方式
    const alignmentSetting = headingStyles.alignment?.[headingKey] || 'left';
    const alignment = this.getAlignmentType(alignmentSetting);

    // 获取标题缩进设置 (4级和5级标题需要左空2字符)
    const leftIndent = headingStyles.indent?.[headingKey] || 0;

    // 标题包含前缀？
    let titleText = token.text;
    const usePrefix = headingStyles.usePrefix?.[headingKey] || false;
    const prefix = headingStyles.prefix?.[headingKey] || '';

    if (usePrefix && prefix) {
      titleText = prefix + titleText;
    }

    // 基于标题级别和样式创建段落格式
    return new Paragraph({
      heading: level,
      style: `Heading${level}`,
      spacing: {
        before: level === 1 ? 240 : 120,
        after: level === 1 ? 120 : 120
      },
      alignment: alignment,
      indent: {
        left: leftIndent
      },
      children: [
        new TextRun({
          text: titleText,
          bold: isBold,
          font: headingFont,
          size: (headingStyles.sizes?.[headingKey] || 24) * 2, // 转换为半点单位
          color: headingColor
        })
      ]
    });
  }

  /**
   * @method createParagraph
   * @param {Object} token - 段落token
   * @returns {Paragraph} 段落对象
   */
  createParagraph(token) {
    // 确保文本是字符串
    if (typeof token.text !== 'string') {
      token.text = token.text?.toString() || '';
    }

    // 使用正文样式替代段落样式
    const paragraphStyles = this.styles.paragraph || {};

    // 处理行距设置
    // 如果是固定行距（磅值），直接使用数值；如果是倍数行距，则需要乘以240转换
    const lineSpacingValue = paragraphStyles.lineSpacingRule === 'exact' ?
                             paragraphStyles.lineSpacing * 20 : // 磅值转twip (1磅 = 20 twip)
                             paragraphStyles.lineSpacing * 240; // 倍数行距

    const lineSpacingRule = paragraphStyles.lineSpacingRule === 'exact' ?
                            'exact' : 'auto';

    return new Paragraph({
      spacing: {
        after: paragraphStyles.spacing || 0,
        line: lineSpacingValue,
        lineRule: lineSpacingRule
      },
      indent: {
        firstLine: paragraphStyles.firstLineIndent || 800 // 首行缩进2字符，约为800 twip
      },
      alignment: this.getAlignmentType(paragraphStyles.alignment) || AlignmentType.JUSTIFIED, // 两端对齐
      children: this.parseInlineContent(token.text)
    });
  }

  /**
   * @method createList
   * @param {Object} token - 列表token
   * @param {number} level - 嵌套级别
   * @returns {Array<Paragraph>}
   */
  createList(token, level = 0) {
    const paragraphs = [];
    const paragraphStyles = this.styles.paragraph;
    const listStyles = this.styles.list || {};

    // 处理行距设置
    // 如果是固定行距（磅值），直接使用数值；如果是倍数行距，则需要乘以240转换
    const lineSpacingValue = paragraphStyles.lineSpacingRule === 'exact' ?
                            paragraphStyles.lineSpacing * 20 : // 磅值转twip (1磅 = 20 twip)
                            paragraphStyles.lineSpacing * 240; // 倍数行距

    const lineSpacingRule = paragraphStyles.lineSpacingRule === 'exact' ?
                           'exact' : 'auto';

    // 根据GB/T 9704-2012标准，公文中的条款序号格式有严格规定：
    // 一级条款："一、二、三、..." (黑体)
    // 二级条款："(一)(二)(三)..." (楷体)
    // 三级条款："1. 2. 3. ..." (仿宋)
    // 四级条款："(1)(2)(3)..." (仿宋)

    // 序号字体设置
    const levelFonts = [
      "黑体",         // 一级条款：黑体
      "楷体",         // 二级条款：楷体
      "仿宋_GB2312",  // 三级条款：仿宋
      "仿宋_GB2312"   // 四级条款：仿宋
    ];

    // 是否加粗，一般只有一级序号（黑体）需要加粗
    const levelBold = [true, false, false, false];

    token.items.forEach(item => {
      // 处理从Word粘贴的内容
      let itemText = '';
      if (typeof item.text === 'string') {
        // 处理可能包含子项目的情况（如：项目二 * 子项目A * 子项目B）
        if (item.text.includes(' * ')) {
          const parts = item.text.split(' * ');
          itemText = parts[0];

          // 为子项目创建嵌套列表项
          const subItems = parts.slice(1).map(subText => {
            return { text: subText, items: [] };
          });

          // 如果之前没有嵌套列表，则添加
          if (!item.items) {
            item.items = [];
          }

          // 将解析出的子项目添加到嵌套列表中
          item.items.push(...subItems);
        } else {
          itemText = item.text;
        }
      } else {
        // 处理非字符串类型
        itemText = String(item.text || '');
      }

      // 清理列表项前缀
      // 如果列表项已经有中文序号或特定格式，保留原文本
      let cleanedText = itemText;

      // 清理可能的前导序号和符号
      if (token.ordered) {
        cleanedText = cleanedText.replace(/^[\d一二三四五六七八九十]+(\.|\、|\）|\))\s*/, '');
      } else {
        cleanedText = cleanedText.replace(/^[●○■•◦▪▫□▹▻➢➣➤◆◇◈⦿⦾⚫⚪✦✧✩✪✫✬✭✮✯✰✱✲✳✴✵✶✷✸✹✺✻✼❉❊❋⁕⁑⁂✽✾✿❀❁❂❃❄❅❆❇❈❉❊❋☙☯✡✢✣✤✥✦✧✩✪✫✬✭✮✯✰✱✲✳✴✵✶✷✸✹✺✻✼⚜❧\t]+\s*/, '');
      }

      // 获取当前级别的字体和加粗设置
      const levelFont = levelFonts[Math.min(level, levelFonts.length - 1)];
      const isLevelBold = levelBold[Math.min(level, levelBold.length - 1)];

      // 创建段落，应用公文规范中的对应格式
      const para = new Paragraph({
        spacing: {
          before: 120,
          after: 120,
          line: lineSpacingValue,
          lineRule: lineSpacingRule
        },
        bullet: token.ordered ? undefined : {
          level: level
        },
        numbering: token.ordered ? {
          reference: "orderedList",
          level: level
        } : undefined,
        children: [
          // 使用自定义格式处理列表项内容
          new TextRun({
            text: cleanedText,
            font: levelFont,
            size: paragraphStyles.size * 2, // 3号字体，与正文一致
            bold: isLevelBold,
            color: paragraphStyles.color.replace('#', '')
          })
        ],
        alignment: AlignmentType.JUSTIFIED // 公文要求正文两端对齐
      });

      paragraphs.push(para);

      // 处理嵌套列表
      if (item.items && item.items.length > 0) {
        // 递归处理嵌套列表
        const nestedToken = {
          type: 'list',
          ordered: token.ordered,
          items: item.items
        };
        const subList = this.createList(nestedToken, Math.min(level + 1, 3)); // 最多只支持到4级嵌套
        paragraphs.push(...subList);
      }
    });

    return paragraphs;
  }

  /**
   * @method createCode
   * @param {Object} token - 代码token
   * @returns {Paragraph}
   */
  createCode(token) {
    // 处理代码文本，确保为字符串
    let codeText = token.text;
    if (typeof codeText !== 'string') {
      codeText = String(codeText || '');
    }

    // 处理从Word粘贴的代码块，它们可能已经丢失换行符
    // 查找像"function xx() {"这样的模式后面应该有换行
    codeText = codeText.replace(/\{(?!\s*\n)/g, '{\n');
    // 为每个分号后面添加换行（如果没有的话）
    codeText = codeText.replace(/;(?!\s*\n)/g, ';\n');
    // 分行处理，保持缩进
    const lines = codeText.split(/\n|\r\n/);

    // 检测并调整缩进
    const codeLines = lines.map(line => {
      // 移除行首的过多空格但保持适当缩进
      const trimmedLine = line.trimStart();
      // 如果当前行看起来是缩进代码，添加适当的缩进
      if (trimmedLine.startsWith('}') ||
          trimmedLine.startsWith('else') ||
          trimmedLine.startsWith('catch')) {
        return '  ' + trimmedLine;
      }
      return trimmedLine;
    });

    // 获取代码样式 - 根据GB/T 9704-2012标准，附件（如代码）可以使用等线字体
    const codeStyles = this.styles.code;
    const codeFont = codeStyles.font || "等线";
    const codeFontSize = (codeStyles.size || 16) * 2; // 5号字体，适合附件代码
    const codeColor = codeStyles.color.replace('#', '');
    const codeBackgroundColor = codeStyles.backgroundColor.replace('#', '');

    // 构建代码块段落 - 公文附件格式
    // 附件标识和内容之间应有一空行
    const codeBlock = new Paragraph({
      spacing: {
        before: 240,
        after: 240,
        line: 360 // 固定行距，适合代码
      },
      indent: {
        left: 600, // 缩进调整为公文附件要求
        right: 600
      },
      shading: {
        type: 'clear',
        fill: codeBackgroundColor
      },
      border: {
        top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
      },
      children: [
        new TextRun({
          text: codeLines.join('\n'),
          font: codeFont,
          size: codeFontSize,
          color: codeColor
        })
      ],
      alignment: AlignmentType.LEFT // 代码通常左对齐
    });

    // 根据公文规范，附件应有标识（如：附件：代码示例）
    // 实际应用中可在Markdown中明确标明"附件"，此处仅作参考
    return codeBlock;
  }

  /**
   * @method createBlockquote
   * @param {Object} token - 引用块token
   * @returns {Paragraph}
   */
  createBlockquote(token) {
    // 处理引用文本，确保为字符串
    let quoteText = token.text;
    if (typeof quoteText !== 'string') {
      quoteText = String(quoteText || '');
    }

    // 处理从Word粘贴的引用块
    // 清理可能的">"前缀和多余空格
    quoteText = quoteText.replace(/^>\s*/mg, '');
    quoteText = quoteText.replace(/^\s+/mg, '');

    // 获取样式配置
    const blockquoteStyles = this.styles.blockquote || {};
    const paragraphStyles = this.styles.paragraph;
    const blockquoteColor = (blockquoteStyles.color || paragraphStyles.color).replace('#', '');
    const borderColor = (blockquoteStyles.borderColor || "#000000").replace('#', '');

    // 处理行距设置
    // 如果是固定行距（磅值），直接使用数值；如果是倍数行距，则需要乘以240转换
    const lineSpacingValue = paragraphStyles.lineSpacingRule === 'exact' ?
                             paragraphStyles.lineSpacing * 20 : // 磅值转twip (1磅 = 20 twip)
                             paragraphStyles.lineSpacing * 240; // 倍数行距

    const lineSpacingRule = paragraphStyles.lineSpacingRule === 'exact' ?
                            'exact' : 'auto';

    // 根据GB/T 9704-2012标准，引用（如领导批示、文件引述）使用仿宋_GB2312字体
    // 首行缩进2字符，具有一定的标识性
    return new Paragraph({
      spacing: {
        before: 120,
        after: 120,
        line: lineSpacingValue,
        lineRule: lineSpacingRule
      },
      indent: {
        left: blockquoteStyles.leftIndent || 800, // 左侧缩进
        firstLine: blockquoteStyles.firstLineIndent || 800 // 首行缩进2字符（约800twip）
      },
      border: {
        left: {
          style: BorderStyle.SINGLE,
          size: 12,
          color: borderColor,
          space: 15
        }
      },
      alignment: AlignmentType.JUSTIFIED, // 两端对齐，符合公文规范
      children: this.parseInlineContent(quoteText).map(run => {
        // 确保引用文本使用仿宋_GB2312字体
        return new TextRun({
          ...run,
          font: "仿宋_GB2312",
          size: paragraphStyles.size * 2, // 与正文字号相同，通常为3号（16pt）
          color: blockquoteColor
        });
      })
    });
  }

  /**
   * @method createHorizontalRule
   * @returns {Paragraph}
   */
  createHorizontalRule() {
    return new Paragraph({
      text: "",
      border: {
        bottom: {
          style: BorderStyle.SINGLE,
          size: 1,
          color: 'AAAAAA'
        }
      },
      spacing: { before: 240, after: 240 }
    });
  }

  /**
   * @method createTable
   * @param {Object} token - 表格token
   * @returns {Table} 表格对象
   */
  createTable(token) {
    try {
      // 获取样式设置
      const tableStyles = this.styles.table || {};
      const paragraphStyles = this.styles.paragraph || {};

      // 处理颜色值（确保没有#前缀）
      const tableBorderColor = (tableStyles.borderColor || "000000").replace('#', '');
      const tableHeaderBg = (tableStyles.headerBackground || "E6E6E6").replace('#', '');
      const textColor = (paragraphStyles.color || "000000").replace('#', '');

      // 设置字体和字号
      const headerFont = tableStyles.headerFont || "仿宋_GB2312";
      const fontSize = (tableStyles.fontSize || 16) * 2; // 转换为twip单位

      // 处理表格数据
      let headerCells = [];
      let rowsData = [];

      // 处理表头 - 处理可能是JSON字符串的情况
      if (token.header) {
        headerCells = Array.isArray(token.header)
          ? token.header.map(cell => this.processTableCell(cell))
          : this.processTableData(token.header);
      }

      // 处理表格行 - 处理可能是JSON字符串的情况
      if (token.rows) {
        if (Array.isArray(token.rows)) {
          rowsData = token.rows.map(row =>
            Array.isArray(row)
              ? row.map(cell => this.processTableCell(cell))
              : this.processTableData(row)
          );
        } else if (typeof token.rows === 'string') {
          // 可能是制表符分隔的行数据，按行分割
          const rowLines = token.rows.split('\n').filter(line => line.trim());
          rowsData = rowLines.map(line => this.processTableData(line));
        }
      }

      // 如果没有表头但有文本，尝试从文本中解析表格
      if ((!headerCells || headerCells.length === 0) && token.text) {
        const tableText = token.text.trim();
        if (tableText) {
          const lines = tableText.split('\n').filter(line => line.trim());
          if (lines.length > 0) {
            // 第一行作为表头
            headerCells = this.processTableData(lines[0]);
            // 剩余行作为数据行
            if (lines.length > 1) {
              rowsData = lines.slice(1).map(line => this.processTableData(line));
            }
          }
        }
      }

      // 确保表格数据有效
      if (!headerCells || headerCells.length === 0) {
        console.warn("创建表格失败：无效的表头数据");
        return new Paragraph({ text: "无效的表格数据" });
      }

      // 计算表格列数和列宽
      const columnCount = headerCells.length;
      const tableWidth = 8000; // 表格总宽度，约为页面宽度的80%
      const columnWidth = Math.floor(tableWidth / columnCount);

      // 创建表头行
      const headerRow = new TableRow({
        tableHeader: true, // 指定这是表头行
        height: { value: 400, rule: 'atLeast' }, // 设置最小行高
        children: headerCells.map(cellText => {
          return new TableCell({
            shading: { fill: tableHeaderBg, type: 'clear' },
            verticalAlign: 'center',
            margins: { top: 100, bottom: 100, left: 100, right: 100 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 60, after: 60 },
                children: [
                  new TextRun({
                    text: cellText,
                    font: headerFont,
                    size: fontSize,
                    bold: true
                  })
                ]
              })
            ]
          });
        })
      });

      // 创建数据行
      const dataRows = rowsData.map(rowData => {
        return new TableRow({
          height: { value: 400, rule: 'atLeast' }, // 设置最小行高
          children: rowData.map(cellText => {
            return new TableCell({
              verticalAlign: 'center',
              margins: { top: 100, bottom: 100, left: 100, right: 100 },
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { before: 60, after: 60 },
                  children: [
                    new TextRun({
                      text: cellText,
                      font: paragraphStyles.font || "仿宋_GB2312",
                      size: fontSize,
                      color: textColor
                    })
                  ]
                })
              ]
            });
          })
        });
      });

      // 创建表格
      return new Table({
        // 表格宽度设置
        width: { size: tableWidth, type: WidthType.DXA },

        // 表格对齐方式
        alignment: AlignmentType.CENTER,

        // 表格边框设置
        borders: {
          top: { style: BorderStyle.SINGLE, size: tableStyles.borderWidth || 1, color: tableBorderColor },
          bottom: { style: BorderStyle.SINGLE, size: tableStyles.borderWidth || 1, color: tableBorderColor },
          left: { style: BorderStyle.SINGLE, size: tableStyles.borderWidth || 1, color: tableBorderColor },
          right: { style: BorderStyle.SINGLE, size: tableStyles.borderWidth || 1, color: tableBorderColor },
          insideHorizontal: { style: BorderStyle.SINGLE, size: tableStyles.borderWidth || 1, color: tableBorderColor },
          insideVertical: { style: BorderStyle.SINGLE, size: tableStyles.borderWidth || 1, color: tableBorderColor }
        },

        // 表格行数据
        rows: [headerRow, ...dataRows],

        // 设置列宽
        columnWidths: Array(columnCount).fill(columnWidth),

        // 表格布局类型：固定宽度布局
        layout: TableLayoutType.FIXED
      });
    } catch (error) {
      console.error("创建表格时出错:", error);
      return new Paragraph({ text: `表格创建失败: ${error.message}` });
    }
  }

  /**
   * @method createImagePlaceholder
   * @description 创建图片占位符或插入真实图片
   * @param {Object} token - 图片token
   * @returns {Paragraph}
   */
  createImagePlaceholder(token) {
    try {
      // 获取图片样式
      const imageStyles = this.styles.image || {};
      const paragraphStyles = this.styles.paragraph;

      // 获取图片描述信息
      let altText = token.text || token.title || "图片";
      if (typeof altText === 'string' && altText.match(/!\[(.+?)\]/)) {
        altText = altText.match(/!\[(.+?)\]/)[1];
      }

      // 获取图片路径
      const imgSrc = token.href || '';

      // 确定对齐方式
      let alignment = AlignmentType.CENTER;
      if (imageStyles.alignment) {
        switch (imageStyles.alignment.toLowerCase()) {
          case 'left':
            alignment = AlignmentType.LEFT;
            break;
          case 'right':
            alignment = AlignmentType.RIGHT;
            break;
          default:
            alignment = AlignmentType.CENTER;
        }
      }

      // 设置图片最大宽度和高度
      const maxWidth = imageStyles.maxWidth || 6000; // 约为页面宽度的60%
      const aspectRatio = 0.75; // 按照4:3的默认比例
      const maxHeight = Math.round(maxWidth * aspectRatio);

      // 检查是否为Base64数据
      if (imgSrc && imgSrc.startsWith('data:')) {
        try {
          // 提取MIME类型和Base64数据
          const matches = imgSrc.match(/^data:([A-Za-z-+/]+);base64,(.+)$/);

          if (!matches || matches.length !== 3) {
            throw new Error('无效的Base64图片数据');
          }

          const mimeType = matches[1]; // 如 "image/png"
          const base64Data = matches[2]; // 实际Base64编码数据，不带前缀

          // 验证支持的MIME类型
          if (!mimeType.startsWith('image/')) {
            throw new Error(`不支持的MIME类型: ${mimeType}`);
          }

          // 检查webp格式 (docx.js不支持webp)
          if (mimeType === 'image/webp') {
            throw new Error('不支持webp格式图片，请转换为PNG或JPEG格式');
          }

          console.log(`处理Base64图片, MIME类型: ${mimeType}, 数据长度: ${base64Data.length}`);

          try {
            // 创建包含图片的段落
            return new Paragraph({
              alignment: alignment,
              spacing: { before: 200, after: 200 },
              children: [
                new ImageRun({
                  data: base64Data, // 直接使用原始Base64字符串，不包含data:...前缀
                  transformation: {
                    width: maxWidth,
                    height: maxHeight,
                    rotation: 0,
                  },
                  altText: altText,
                  type: mimeType // 明确指定MIME类型
                })
              ]
            });
          } catch (imageError) {
            console.error('创建ImageRun时出错:', imageError);
            throw imageError;
          }
        } catch (base64Error) {
          console.error('处理Base64图片失败:', base64Error);
          // 转入下面的占位符处理
        }
      } else if (imgSrc && (imgSrc.startsWith('http') || imgSrc.startsWith('https'))) {
        // 对于URL图片，我们需要使用预处理过的Base64数据
        console.log('远程图片URL应该已被预处理为Base64格式:', imgSrc);

        return new Paragraph({
          alignment: alignment,
          spacing: { before: 200, after: 200 },
          children: [
            new TextRun({
              text: `图片URL: ${imgSrc} (需要先将远程图片转换为Base64)`,
              italic: true,
              color: paragraphStyles.color.replace('#', '')
            })
          ],
          border: {
            top: { style: BorderStyle.DOTTED, size: 1, color: "AAAAAA" },
            bottom: { style: BorderStyle.DOTTED, size: 1, color: "AAAAAA" },
            left: { style: BorderStyle.DOTTED, size: 1, color: "AAAAAA" },
            right: { style: BorderStyle.DOTTED, size: 1, color: "AAAAAA" }
          },
          shading: {
            type: 'clear',
            fill: "F5F5F5"
          }
        });
      }

      // 返回占位符作为fallback
      const imageDescription = altText ? `图片: ${altText}` : '图片';
      const paragraphColor = paragraphStyles.color.replace('#', '');

      return new Paragraph({
        alignment: alignment,
        spacing: {
          before: 200,
          after: 200
        },
        border: {
          top: { style: BorderStyle.DOTTED, size: 1, color: "AAAAAA" },
          bottom: { style: BorderStyle.DOTTED, size: 1, color: "AAAAAA" },
          left: { style: BorderStyle.DOTTED, size: 1, color: "AAAAAA" },
          right: { style: BorderStyle.DOTTED, size: 1, color: "AAAAAA" }
        },
        shading: {
          type: 'clear',
          fill: "F5F5F5"
        },
        children: [
          new TextRun({
            text: imageDescription,
            italic: true,
            color: paragraphColor
          })
        ]
      });
    } catch (error) {
      console.error("创建图片时出错:", error);
      return new Paragraph({ text: `图片创建失败: ${error.message}` });
    }
  }

  /**
   * @method createTaskList
   * @param {Object} token - 任务列表token
   * @returns {Array<Paragraph>} 任务列表段落数组
   */
  createTaskList(token) {
    const paragraphs = [];
    const listStyles = this.styles.list?.task || {};
    const paragraphStyles = this.styles.paragraph;

    token.items.forEach(item => {
      // 检查任务项是否已完成（格式为：- [x] 或 - [ ]）
      const isCompleted = item.text.startsWith('[x]');
      let itemText = item.text;

      // 移除任务标记，保留实际文本内容
      if (isCompleted) {
        itemText = itemText.replace(/^\[x\]\s+/, '');
      } else {
        itemText = itemText.replace(/^\[\s*\]\s+/, '');
      }

      // 确保文本是字符串
      if (typeof itemText !== 'string') {
        itemText = itemText?.toString() || '';
      }

      // 创建任务列表段落
      const para = new Paragraph({
        spacing: {
          before: 120,
          after: 120
        },
        indent: {
          left: listStyles.indentLevel || 720
        },
        children: [
          new TextRun({
            text: isCompleted ?
              (listStyles.completedChar || '☑') :
              (listStyles.uncompletedChar || '☐'),
            size: paragraphStyles.size * 2,
            font: paragraphStyles.font
          }),
          new TextRun({
            text: ' ' + itemText,
            size: paragraphStyles.size * 2,
            font: paragraphStyles.font,
            color: paragraphStyles.color.replace('#', '')
          })
        ]
      });

      paragraphs.push(para);
    });

    return paragraphs;
  }

  /**
   * @method createFootnote
   * @param {Object} token - 脚注token
   * @returns {Paragraph} 脚注段落
   */
  createFootnote(token) {
    // 获取脚注样式
    const footnoteStyles = this.styles.footnote || {};
    const paragraphStyles = this.styles.paragraph;

    // 确保文本是字符串
    const footnoteLabel = typeof token.label === 'string' ? token.label : String(token.label || '');
    const footnoteText = typeof token.text === 'string' ? token.text : String(token.text || '');

    // 创建脚注段落
    return new Paragraph({
      spacing: {
        before: 120,
        after: 120,
        line: paragraphStyles.lineSpacing * 240
      },
      children: [
        new TextRun({
          text: `[${footnoteLabel}] `,
          size: (footnoteStyles.size || paragraphStyles.size - 3) * 2,
          font: footnoteStyles.font || paragraphStyles.font,
          color: (footnoteStyles.color || paragraphStyles.color).replace('#', ''),
          superScript: true
        }),
        new TextRun({
          text: footnoteText,
          size: (footnoteStyles.size || paragraphStyles.size - 3) * 2,
          font: footnoteStyles.font || paragraphStyles.font,
          color: (footnoteStyles.color || paragraphStyles.color).replace('#', '')
        })
      ]
    });
  }

  /**
   * @method saveAsDocx
   * @description 保存文档为docx文件
   * @param {string} filename - 文件名
   */
  saveAsDocx(filename = 'document.docx') {
    if (!this.doc) {
      throw new Error('请先调用convert方法生成文档');
    }

    Packer.toBlob(this.doc).then(blob => {
      saveAs(blob, filename);
    });
  }

  /**
   * @method getAlignmentType
   * @description 将字符串对齐方式转换为docx库的AlignmentType
   * @param {string} alignment - 对齐方式字符串
   * @returns {AlignmentType} docx对齐方式枚举
   */
  getAlignmentType(alignment) {
    switch (alignment) {
      case 'center':
        return AlignmentType.CENTER;
      case 'right':
        return AlignmentType.RIGHT;
      case 'justified':
        return AlignmentType.JUSTIFIED;
      case 'left':
      default:
        return AlignmentType.LEFT;
    }
  }
}

export { Md2Docx };
