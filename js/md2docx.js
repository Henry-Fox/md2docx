// 导入所需模块
import { marked } from './marked.esm.js';
import { saveAs } from 'file-saver';
import defaultStyles from '../styles/default-styles.json';
import { Md2Json } from './md2json.js';
import {
  Document,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  BorderStyle,
  Table,
  TableRow,
  TableCell,
  Packer,
  WidthType,
  TableLayoutType,
  ImageRun,
  ExternalHyperlink,
  LevelFormat
} from 'docx';

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
    // 初始化解析器
    this.md2json = new Md2Json();

    // 合并样式设置
    this.styles = this.mergeStyles(defaultStyles, customStyles);

    // 初始化其他属性
    this.doc = null;
    this.footnotes = new Map();
    this.imageInfos = [];
  }

  /**
   * @method mergeStyles
   * @description 深度合并两个样式对象
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
   * @method previewMarkdown
   * @description 使用marked库将Markdown转换为HTML用于预览
   * @param {string} markdown - Markdown文本
   * @returns {string} HTML字符串
   */
  previewMarkdown(markdown) {
    return marked.parse(markdown);
  }

  /**
   * @method convert
   * @description 将Markdown转换为Word文档
   * @param {string} markdown - Markdown文本
   * @returns {Promise<Blob>} - 返回Word文档的Blob对象
   */
  async convert(markdown) {
    try {
      console.log('开始转换Markdown为Word文档...');

      // 1. 使用md2json解析Markdown为JSON
      console.log('使用md2json解析Markdown...');
      let jsonData;
      try {
        jsonData = await this.md2json.convert(markdown);
        console.log('Markdown解析为JSON完成:', jsonData);
      } catch (parseError) {
        console.error('Markdown解析失败:', parseError);
        // 创建一个最小的有效JSON结构
        jsonData = {
          type: 'document',
          children: [
            {
              type: 'heading',
              level: 1,
              fullContent: '解析错误',
              inlineStyles: [{ content: '解析错误', bold: true }]
            },
            {
              type: 'paragraph',
              fullContent: `无法解析Markdown: ${parseError.message}`,
              inlineStyles: [{ content: `无法解析Markdown: ${parseError.message}` }]
            }
          ]
        };
      }

      // 确保JSON结构有效
      if (!jsonData || !jsonData.children) {
        console.warn('JSON结构无效，创建默认结构');
        jsonData = {
          type: 'document',
          children: [
            {
              type: 'paragraph',
              fullContent: '无效的文档结构',
              inlineStyles: [{ content: '无效的文档结构' }]
            }
          ]
        };
      }

      // 2. 使用JSON数据创建Word文档
      console.log('开始创建Word文档...');
      const doc = this.createDocument(jsonData);

      // 3. 生成docx文件
      console.log('生成docx文件...');
      const blob = await this.generateDocx(doc);

      console.log('文档转换完成');
      return blob;
    } catch (error) {
      console.error('转换过程中出错:', error);
      throw new Error(`转换失败: ${error.message}`);
    }
  }

  /**
   * @method createDocument
   * @description 根据JSON数据创建Word文档
   * @param {Object} jsonData - 从md2json解析得到的JSON数据
   * @returns {Document} - 返回docx.js的Document对象
   */
  createDocument(jsonData) {
    try {
      console.log('开始创建Document对象...');

      // 创建一个空的子元素数组，用于存放所有文档元素
      const children = [];

      // 处理JSON数据中的每个元素
      if (jsonData && jsonData.children && Array.isArray(jsonData.children)) {
        console.log(`开始处理jsonData中的${jsonData.children.length}个元素`);

        jsonData.children.forEach((item, index) => {
          console.log(`处理第${index+1}个元素，类型: ${item.type}`);
          let element = null;

          switch (item.type) {
            case 'heading':
              element = this.createHeading(item);
              break;
            case 'paragraph':
              element = this.createParagraph(item);
              break;
            case 'code_block':
              element = this.createCodeBlock(item);
              break;
            case 'image':
              element = this.createImageElement(item);
              break;
            case 'list':
              // 列表的特殊处理
              this.createListElements(item, children);
              console.log('列表元素已直接添加到children数组');
              break;
            case 'table':
              element = this.createTableElement(item);
              break;
            case 'blockquote':
              element = this.createBlockquote(item);
              break;
            case 'horizontal_rule':
              element = this.createHorizontalRule();
              break;
            default:
              console.warn(`未知的token类型: ${item.type}`);
              break;
          }

          if (element) {
            // 直接添加到children数组
            children.push(element);
            console.log(`成功添加元素 ${item.type} 到文档，children数组现在有${children.length}个元素`);
          }
        });
      } else {
        console.warn('无效的文档结构，没有children数组');
        // 添加一个默认段落，确保文档不为空
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: "无效的文档结构",
                color: "FF0000",
                bold: true
              })
            ]
          })
        );
      }

      // 确保有内容
      if (children.length === 0) {
        console.warn('文档没有内容，添加默认段落');
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: "空文档",
                color: "FF0000",
                bold: true
              })
            ]
          })
        );
      }

      console.log(`准备创建文档对象，包含 ${children.length} 个元素`);

      // 保存到this.children用于备份
      this.children = children;

      // 使用标准的docx.js文档创建方式
      try {
        // 先创建一个Document对象
        const doc = new Document({
          creator: "Md2Docx",
          title: "Markdown转Word文档",
          description: "由Md2Docx自动生成的Word文档"
        });

        // 确保document对象有sections属性
        if (!doc.sections) {
          console.warn('创建的Document对象没有sections属性，手动初始化');
          doc.sections = [];
        }

        // 添加一个section
        doc.sections.push({
          properties: {
            page: {
              size: {
                orientation: this.styles.document.pageOrientation,
                width: this.styles.document.pageSize === "A4" ? 11906 : 12240,
                height: this.styles.document.pageSize === "A4" ? 16838 : 15840
              },
              margin: {
                top: this.styles.document.margins.top,
                right: this.styles.document.margins.right,
                bottom: this.styles.document.margins.bottom,
                left: this.styles.document.margins.left
              }
            }
          },
          children: children
        });

        console.log('Document对象创建完成，检查结构:',
          doc.sections ? `sections: ${doc.sections.length}个` : 'sections: 无效',
          doc.sections && doc.sections.length > 0 && doc.sections[0].children ?
            `children: ${doc.sections[0].children.length}个` : 'children: 无效'
        );

        return doc;
      } catch (docError) {
        console.error('创建Document对象失败，使用后备方法:', docError);

        // 使用直接的方式创建文档
        return {
          creator: "Md2Docx",
          title: "Markdown转Word文档",
          description: "由Md2Docx自动生成的Word文档",
          children: children, // 保存在根级别，后续generateDocx会处理
          sections: [
            {
              properties: {
                page: {
                  size: {
                    orientation: this.styles.document.pageOrientation,
                    width: this.styles.document.pageSize === "A4" ? 11906 : 12240,
                    height: this.styles.document.pageSize === "A4" ? 16838 : 15840
                  },
                  margin: {
                    top: this.styles.document.margins.top,
                    right: this.styles.document.margins.right,
                    bottom: this.styles.document.margins.bottom,
                    left: this.styles.document.margins.left
                  }
                }
              },
              children: children
            }
          ]
        };
      }
    } catch (error) {
      console.error('创建Document对象时出错:', error);

      // 返回一个最小可用的文档对象
      const children = [
        new Paragraph({
          children: [
            new TextRun({
              text: `文档创建失败: ${error.message}`,
              color: "FF0000",
              bold: true
            })
          ]
        })
      ];

      // 保存children的引用，以便generateDocx可以使用
      this.children = children;

      return {
        creator: "Md2Docx",
        title: "错误文档",
        description: "文档创建过程中发生错误",
        children: children,
        sections: [
          {
            children: children
          }
        ]
      };
    }
  }

  /**
   * @method generateDocx
   * @description 生成docx文件
   * @param {Object} doc - 文档对象
   * @returns {Promise<Blob>} docx文件的Blob对象
   */
  async generateDocx(doc) {
    try {
      console.log('开始生成DOCX文件...');

      // 检查文档对象
      if (!doc) {
        console.error('无效的文档对象，无法生成DOCX');
        throw new Error('无效的文档对象');
      }

      console.log('文档对象:', doc);
      console.log('Document结构:',
        doc.sections ? `有sections(${doc.sections.length}个)` : '无sections',
        doc.document ? '有document属性' : '无document属性'
      );

      // 创建一个全新的文档对象，将内容复制过去
      const newDoc = new Document({
        creator: "Md2Docx",
        title: "Markdown转Word文档",
        description: "由Md2Docx自动生成的Word文档",
        sections: [
          {
            properties: {
              page: {
                size: {
                  orientation: this.styles.document.pageOrientation,
                  width: this.styles.document.pageSize === "A4" ? 11906 : 12240,
                  height: this.styles.document.pageSize === "A4" ? 16838 : 15840
                },
                margin: {
                  top: this.styles.document.margins.top,
                  right: this.styles.document.margins.right,
                  bottom: this.styles.document.margins.bottom,
                  left: this.styles.document.margins.left
                }
              }
            },
            children: []
          }
        ]
      });

      // 获取原始doc的内容
      let contentChildren = [];

      if (doc.sections && Array.isArray(doc.sections) && doc.sections.length > 0 &&
          doc.sections[0].children && Array.isArray(doc.sections[0].children)) {
        // 如果原文档有效，使用其内容
        contentChildren = doc.sections[0].children;
        console.log(`从原文档复制${contentChildren.length}个元素`);
      } else if (doc.document && doc.document.sections &&
                 Array.isArray(doc.document.sections) &&
                 doc.document.sections.length > 0 &&
                 doc.document.sections[0].children) {
        // 兼容某些docx.js版本可能使用的结构
        contentChildren = doc.document.sections[0].children;
        console.log(`从document.sections复制${contentChildren.length}个元素`);
      } else if (doc.children && Array.isArray(doc.children)) {
        // 兼容直接存储在doc.children的情况
        contentChildren = doc.children;
        console.log(`从doc.children复制${contentChildren.length}个元素`);
      }

      // 确保有内容
      if (contentChildren.length === 0) {
        console.warn('文档内容为空，添加一个默认段落');
        contentChildren.push(
          new Paragraph({
            children: [
              new TextRun({
                text: "文档内容为空",
                color: "FF0000",
                bold: true,
                size: 24 * 2
              })
            ]
          })
        );
      }

      // 将内容复制到新文档
      newDoc.sections[0].children = contentChildren;

      console.log(`新文档创建完成，包含${newDoc.sections[0].children.length}个元素`);

      console.log(`准备生成DOCX文件，文档包含 ${newDoc.sections.length} 个节和 ${newDoc.sections[0].children.length} 个元素`);

      try {
        // 使用最简单的方式创建blob对象
        return await Packer.toBlob(newDoc);
      } catch (error) {
        console.error('使用Packer.toBlob生成文档时错误:', error);

        // 创建最简单的测试文档
        const testDoc = new Document({
          sections: [
            {
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: "测试文档 - 如果你能看到这段文字，则说明docx.js基本功能正常",
                      bold: true,
                      color: "FF0000",
                      font: "宋体"
                    })
                  ]
                })
              ]
            }
          ]
        });

        console.log('尝试生成测试文档...');
        return await Packer.toBlob(testDoc);
      }
    } catch (error) {
      console.error('生成DOCX文件时发生错误:', error);

      // 创建内容为错误消息的最小文档
      const minimalDoc = new Document({
        sections: [
          {
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: `生成文档时出错: ${error.message}`,
                    bold: true,
                    color: "FF0000"
                  })
                ]
              })
            ]
          }
        ]
      });

      try {
        console.log('生成错误信息文档...');
        return await Packer.toBlob(minimalDoc);
      } catch (finalError) {
        console.error('创建错误文档也失败:', finalError);
        throw error; // 抛出原始错误
      }
    }
  }

  /**
   * @method saveAsDocx
   * @description 保存docx文件
   * @param {string} filename - 文件名
   */
  saveAsDocx(filename = 'document.docx') {
    try {
      console.log(`开始保存文档为 ${filename}...`);

      if (!this.doc) {
        console.error('没有可保存的文档，document对象为空');
        throw new Error('没有可保存的文档');
      }

      // 检查document对象是否有效
      if (!this.doc.sections || !Array.isArray(this.doc.sections) || this.doc.sections.length === 0) {
        console.error('文档sections无效，无法保存');
        throw new Error('文档sections无效');
      }

      if (!this.doc.sections[0].children || !Array.isArray(this.doc.sections[0].children)) {
        console.warn('文档没有内容，将保存空文档');
        this.doc.sections[0].children = [
          new Paragraph({
            children: [
              new TextRun({
                text: "空文档",
                color: "FF0000",
                bold: true
              })
            ]
          })
        ];
      }

      console.log(`文档包含 ${this.doc.sections.length} 个section和 ${this.doc.sections[0].children.length} 个元素`);

      // 严格按照docx.js官方文档保存文档
      // 确保创建一个新的Document对象，以避免可能的引用问题
      const docToSave = this.doc;

      // 使用标准的Promise方式处理
      Packer.toBlob(docToSave)
        .then(blob => {
          console.log(`文档生成成功，大小: ${Math.round(blob.size / 1024)} KB`);

          try {
            // 使用FileSaver.js的saveAs函数保存文件
            saveAs(blob, filename);
            console.log(`文档已保存为 ${filename}`);
          } catch (saveError) {
            console.error('保存文件时出错:', saveError);

            // 尝试使用替代方法保存文件
            try {
              // 创建URL并触发下载
              const url = URL.createObjectURL(blob);
              const a = document.createElement('a');
              a.href = url;
              a.download = filename;
              document.body.appendChild(a);
              a.click();

              // 清理
              setTimeout(() => {
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
              }, 0);

              console.log(`已使用替代方法保存文档: ${filename}`);
            } catch (altSaveError) {
              console.error('替代保存方法也失败:', altSaveError);
              throw new Error(`无法保存文档: ${saveError.message}`);
            }
          }
        })
        .catch(error => {
          console.error('生成文档Blob时出错:', error);

          // 尝试创建最小文档
          const minimalDoc = new Document({
            sections: [{
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: `保存文档时出错: ${error.message}`,
                      color: "FF0000",
                      bold: true
                    })
                  ]
                })
              ]
            }]
          });

          Packer.toBlob(minimalDoc)
            .then(fallbackBlob => {
              try {
                saveAs(fallbackBlob, filename);
                console.log(`已保存错误提示文档: ${filename}`);
              } catch (saveError) {
                console.error('保存错误提示文档也失败:', saveError);
              }
            })
            .catch(fallbackError => {
              console.error('创建最小文档也失败:', fallbackError);
            });
        });
    } catch (error) {
      console.error('保存文档过程中发生错误:', error);
      throw error;
    }
  }

  /**
   * @method getDocumentStyles
   * @description 获取文档样式
   * @return {Object} 文档样式配置
   */
  getDocumentStyles() {
    return {
        paragraphStyles: [
        {
          id: 'CodeBlock',
          name: 'Code Block',
          basedOn: 'Normal',
      run: {
            font: 'Courier New',
            size: 20,
            color: '333333'
      },
      paragraph: {
            spacing: { before: 120, after: 120 },
            indent: { left: 720 }
          }
        },
        {
          id: 'Blockquote',
          name: 'Block Quote',
          basedOn: 'Normal',
                  paragraph: {
            spacing: { before: 120, after: 120 },
            indent: { left: 720 }
                  },
                  run: {
            italics: true,
            color: '666666'
          }
        }
      ]
    };
  }

  /**
   * @method createHeading
   * @description 创建标题段落
   * @param {Object} token - 标题标记
   * @return {Object} 标题段落对象
   */
  createHeading(token) {
    console.log("处理标题:", token);

    const headingStyles = this.styles.heading;

    // 确定标题级别
    const level = token.level || token.depth || 1;
    console.log(`标题级别: ${level}`);

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

    // 获取标题缩进设置
      const leftIndent = headingStyles.indent?.[headingKey] || 0;

    // 创建文本运行数组
    const children = [];

    // 添加前缀（如果需要）
    const usePrefix = headingStyles.usePrefix?.[headingKey] || false;
    const prefix = headingStyles.prefix?.[headingKey] || '';

    if (usePrefix && prefix) {
      children.push(
        new TextRun({
          text: prefix,
          bold: isBold,
          font: headingFont,
          size: (headingStyles.sizes?.[headingKey] || 24) * 2,
          color: headingColor
        })
      );
    }

    // 处理inlineStyles中的文本和格式
    if (token.inlineStyles && token.inlineStyles.length > 0) {
      token.inlineStyles.forEach(style => {
        if (style.content) {
          children.push(
            new TextRun({
              text: style.content,
              bold: style.bold || isBold, // 合并Markdown样式和标题样式
              italic: style.italic,
              strike: style.strike,
              font: headingFont,
              size: (headingStyles.sizes?.[headingKey] || 24) * 2,
              color: headingColor,
              underline: style.underline ? {} : undefined,
              superScript: style.superscript,
              subScript: style.subscript
            })
          );
        }
      });
    } else if (token.fullContent) {
      // 兜底方案，使用fullContent
      children.push(
        new TextRun({
          text: token.fullContent,
          bold: isBold,
          font: headingFont,
          size: (headingStyles.sizes?.[headingKey] || 24) * 2,
          color: headingColor
        })
      );
    } else {
      // 极端情况，生成空标题
      children.push(
        new TextRun({
          text: `标题 ${level}`,
          bold: isBold,
          font: headingFont,
          size: (headingStyles.sizes?.[headingKey] || 24) * 2,
          color: headingColor
        })
      );
    }

    console.log(`标题样式: 字体=${headingFont}, 颜色=${headingColor}, 加粗=${isBold}, 对齐=${alignmentSetting}`);

    // 创建段落对象
    const paragraph = new Paragraph({
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
      children: children
    });

    console.log("创建标题段落:", paragraph);
    return paragraph;
  }

  /**
   * @method createParagraph
   * @description 创建段落
   * @param {Object} token - 段落标记
   * @return {Object} 段落对象
   */
  createParagraph(token) {
    // 使用正文样式
    const paragraphStyles = this.styles.paragraph || {};

    // 处理行距设置
    const lineSpacingValue = paragraphStyles.lineSpacingRule === 'exact' ?
                           paragraphStyles.lineSpacing * 20 : // 磅值转twip
                           paragraphStyles.lineSpacing * 240; // 倍数行距

    const lineSpacingRule = paragraphStyles.lineSpacingRule === 'exact' ?
                          'exact' : 'auto';

    // 创建文本运行数组
    const children = [];

    // 从inlineStyles获取格式信息
    if (token.inlineStyles && token.inlineStyles.length > 0) {
      token.inlineStyles.forEach(style => {
        if (style.content) {
          children.push(
            new TextRun({
              text: style.content,
              size: paragraphStyles.size * 2,
              font: paragraphStyles.font,
              color: paragraphStyles.color.replace('#', ''),
              bold: style.bold,
              italic: style.italic,
              strike: style.strike,
              underline: style.underline ? {} : undefined,
              superScript: style.superscript,
              subScript: style.subscript
            })
          );
        }
      });
    } else if (token.fullContent) {
      // 兜底方案，使用fullContent
      children.push(
        new TextRun({
          text: token.fullContent,
          size: paragraphStyles.size * 2,
              font: paragraphStyles.font,
          color: paragraphStyles.color.replace('#', '')
        })
      );
    } else if (token.rawText) {
      // 备选方案
      children.push(
        new TextRun({
          text: token.rawText,
              size: paragraphStyles.size * 2,
          font: paragraphStyles.font,
          color: paragraphStyles.color.replace('#', '')
        })
      );
    } else {
      // 极端情况
      children.push(
        new TextRun({
          text: "",
          size: paragraphStyles.size * 2,
          font: paragraphStyles.font,
          color: paragraphStyles.color.replace('#', '')
        })
      );
    }

    // 创建段落对象
    return new Paragraph({
              spacing: {
        after: paragraphStyles.spacing || 0,
                line: lineSpacingValue,
        lineRule: lineSpacingRule
      },
      indent: {
        firstLine: paragraphStyles.firstLineIndent || 800 // 首行缩进2字符
      },
      alignment: this.getAlignmentType(paragraphStyles.alignment) || AlignmentType.JUSTIFIED,
      children: children
    });
  }

  /**
   * @method createCodeBlock
   * @description 创建代码块
   * @param {Object} token - 代码块标记
   * @return {Object} 代码块段落
   */
  createCodeBlock(token) {
    // 获取代码文本
    let codeText = '';

    // 直接使用fullContent作为代码内容
    if (token.fullContent) {
      codeText = token.fullContent;
    } else if (token.content) {
      codeText = token.content;
    } else if (token.text) {
      codeText = token.text;
    } else {
      codeText = '';
    }

    // 获取语言信息（如果有）
    const language = token.language || '';
    if (language) {
      console.log(`代码块语言: ${language}`);
    }

    // 获取代码样式
    const codeStyles = this.styles.code || {};

    // 创建代码块段落
    return new Paragraph({
      style: 'CodeBlock',
      spacing: { before: 200, after: 200 },
      children: [
        new TextRun({
          text: codeText,
          font: codeStyles.font || 'Courier New',
          size: (codeStyles.size || 10) * 2,
          color: codeStyles.color || '333333'
        })
      ]
    });
  }

  /**
   * @method createImageElement
   * @description 创建图片元素
   * @param {Object} token - 图片标记
   * @return {Object} 图片段落
   */
  createImageElement(token) {
    // 检查token是否有效
    if (!token) {
      console.warn('无效的图片标记');
      return null;
    }

    // 检查是否有图片信息
    if (!this.imageInfos || !Array.isArray(this.imageInfos) || this.imageInfos.length === 0) {
      console.warn('没有可用的图片信息');
      return this.createImagePlaceholder(token.text || '图片');
    }

    // 查找匹配的图片
    const imageInfo = this.imageInfos.find(info =>
      info && (info.src === token.href || info.alt === token.text)
    );

    if (!imageInfo || !imageInfo.buffer) {
      console.warn(`未找到图片信息: ${token.href || 'N/A'}`);
      return this.createImagePlaceholder(token.text || token.href || '图片');
    }

    try {
      // 确保AlignmentType可用
      const alignment = AlignmentType ? AlignmentType.CENTER : 'center';

      // 创建图片段落
      return new Paragraph({
        children: [
          new ImageRun({
            data: imageInfo.buffer,
            transformation: {
              width: imageInfo.width || 400,
              height: imageInfo.height || 300
            }
          })
        ],
        alignment: alignment,
        spacing: { before: 160, after: 160 }
      });
    } catch (error) {
      console.error('创建图片元素时出错:', error);
      return this.createImagePlaceholder(token.text || '图片(创建失败)');
    }
  }

  /**
   * @method createBlockquote
   * @description 创建引用块
   * @param {Object} token - 引用块标记
   * @return {Object} 引用块段落
   */
  createBlockquote(token) {
    const blockquoteStyles = this.styles.blockquote || {};
    const paragraphStyles = this.styles.paragraph || {};

    // 引用块缩进级别
    const level = token.level || 1;
    const indentSize = level * 720; // 根据引用层级增加缩进

    // 创建文本运行数组
    const children = [];

    // 使用inlineStyles处理格式
    if (token.inlineStyles && token.inlineStyles.length > 0) {
      token.inlineStyles.forEach(style => {
        if (style.content) {
          children.push(
            new TextRun({
              text: style.content,
              size: (blockquoteStyles.size || paragraphStyles.size) * 2,
              font: blockquoteStyles.font || paragraphStyles.font,
              color: (blockquoteStyles.color || paragraphStyles.color).replace('#', ''),
              italics: blockquoteStyles.italic !== undefined ? blockquoteStyles.italic : true,
              bold: style.bold,
              strike: style.strike,
              underline: style.underline ? {} : undefined
            })
          );
        }
      });
    } else if (token.fullContent) {
      // 使用fullContent
      children.push(
        new TextRun({
          text: token.fullContent,
          size: (blockquoteStyles.size || paragraphStyles.size) * 2,
          font: blockquoteStyles.font || paragraphStyles.font,
          color: (blockquoteStyles.color || paragraphStyles.color).replace('#', ''),
          italics: blockquoteStyles.italic !== undefined ? blockquoteStyles.italic : true
        })
      );
    } else {
      // 兜底方案
      const content = token.content || token.text || '';
      children.push(
        new TextRun({
          text: content,
          size: (blockquoteStyles.size || paragraphStyles.size) * 2,
          font: blockquoteStyles.font || paragraphStyles.font,
          color: (blockquoteStyles.color || paragraphStyles.color).replace('#', ''),
          italics: blockquoteStyles.italic !== undefined ? blockquoteStyles.italic : true
        })
      );
    }

    // 创建引用块段落
    return new Paragraph({
      style: 'Blockquote',
      indent: { left: indentSize },
      spacing: { before: 120, after: 120 },
      border: {
        left: {
          color: blockquoteStyles.borderColor || '#CCCCCC',
          space: 12,
          style: BorderStyle.SINGLE,
          size: 4
        }
      },
      children: children
    });
  }

  /**
   * @method createHorizontalRule
   * @description 创建水平分隔线
   * @return {Object} 水平分隔线段落
   */
  createHorizontalRule() {
    return new Paragraph({
      text: '',
      border: {
        bottom: {
          color: '#CCCCCC',
          space: 1,
          style: 'single',
          size: 6
        }
      },
      spacing: { before: 200, after: 200 }
    });
  }

  /**
   * @method createListElements
   * @description 创建列表元素
   * @param {Object} token - 列表标记
   * @param {Array} children - 段落数组
   */
  createListElements(token, children) {
    // 处理列表项
    if (!token || !token.items) {
      console.warn('无效的列表结构:', token);
      return;
    }

    // 确保children是有效的数组
    if (!children || !Array.isArray(children)) {
      console.warn('传递给createListElements的children不是数组');
      return;
    }

    console.log(`处理列表元素，包含 ${token.items.length} 个项目`);

    // 获取列表项和样式
    const listStyles = this.styles.list || {};
    const paragraphStyles = this.styles.paragraph || {};
    const isOrdered = token.listType === 'ordered';
    console.log(`列表类型: ${isOrdered ? '有序' : '无序'}`);

    const childrenCountBefore = children.length;

    if (Array.isArray(token.items)) {
      token.items.forEach((item, index) => {
        // 跳过无效的列表项
        if (!item) {
          console.warn(`列表项 #${index} 无效`);
          return;
        }

        // 获取列表级别
        const level = item.level || 0;

        // 创建列表项文本运行数组
        const textRuns = [];

        // 优先使用inlineStyles处理格式
        if (item.inlineStyles && Array.isArray(item.inlineStyles) && item.inlineStyles.length > 0) {
          console.log(`处理列表项 #${index} 的 ${item.inlineStyles.length} 个内联样式`);
          item.inlineStyles.forEach(style => {
            if (style && style.content) {
              textRuns.push(
                new TextRun({
                  text: style.content,
                  size: paragraphStyles.size * 2,
                  font: paragraphStyles.font,
                  color: paragraphStyles.color?.replace('#', '') || '000000',
                  bold: style.bold,
                  italic: style.italic,
                  strike: style.strike,
                  underline: style.underline ? {} : undefined,
                  superScript: style.superscript,
                  subScript: style.subscript
                })
              );
            }
          });
        } else if (item.fullContent) {
          // 使用fullContent
          console.log(`列表项 #${index} 使用fullContent: ${item.fullContent.substring(0, 30)}${item.fullContent.length > 30 ? '...' : ''}`);
          textRuns.push(
            new TextRun({
              text: item.fullContent,
              size: paragraphStyles.size * 2,
              font: paragraphStyles.font,
              color: paragraphStyles.color?.replace('#', '') || '000000'
            })
          );
        } else if (typeof item === 'string') {
          // 兼容旧版本
          console.log(`列表项 #${index} 是字符串: ${item.substring(0, 30)}${item.length > 30 ? '...' : ''}`);
          textRuns.push(
            new TextRun({
              text: item,
              size: paragraphStyles.size * 2,
              font: paragraphStyles.font,
              color: paragraphStyles.color?.replace('#', '') || '000000'
            })
          );
        } else {
          // 当item是对象但没有处理的属性时
          console.log(`列表项 #${index} 使用替代内容`);
          const content = item.text || item.content || (typeof item === 'object' ? JSON.stringify(item) : String(item || ''));
          textRuns.push(
            new TextRun({
              text: content,
              size: paragraphStyles.size * 2,
              font: paragraphStyles.font,
              color: paragraphStyles.color?.replace('#', '') || '000000'
            })
          );
        }

        // 只有在textRuns不为空时才添加段落
        if (textRuns.length > 0) {
          // 添加列表项到children数组
          const paragraph = new Paragraph({
            text: '',
            indent: { left: level * 720 }, // 根据层级增加缩进
            spacing: { before: 80, after: 80 },
            bullet: {
              level: level
            },
            children: textRuns
          });
          children.push(paragraph);
          console.log(`已添加列表项 #${index} 到children数组`);
        } else {
          console.warn(`列表项 #${index} 的textRuns为空，不添加到文档`);
        }
      });
    }

    const childrenCountAfter = children.length;
    console.log(`列表处理前children数组长度: ${childrenCountBefore}, 处理后: ${childrenCountAfter}, 新增: ${childrenCountAfter - childrenCountBefore}`);
  }

  /**
   * @method createTableElement
   * @description 创建表格元素
   * @param {Object} token - 表格标记
   * @return {Object} 表格对象
   */
  createTableElement(token) {
    // 如果没有表头或行数据，返回null
    if (!token.headers || !token.rows) {
      console.warn('表格缺少表头或行数据');
      return null;
    }

    // 创建表头行
    const headerRow = new TableRow({
      children: token.headers.map(header => {
        // 处理表头单元格内容
        const content = header.fullContent || header.rawText || '';

        return new TableCell({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: content,
                  bold: true
                })
              ]
            })
          ],
          shading: {
            fill: "EEEEEE"
          }
        });
      })
    });

    // 获取表格对齐方式
    const alignments = token.alignments || [];

    // 创建数据行
    const rows = token.rows.map(rowData => {
      return new TableRow({
        children: rowData.map((cell, cellIndex) => {
          // 处理单元格内容
          const content = cell.fullContent || cell.rawText || '';

          // 获取该列的对齐方式
          const cellAlignment = alignments[cellIndex]
            ? this.getAlignmentType(alignments[cellIndex])
            : AlignmentType.LEFT;

          return new TableCell({
            children: [
              new Paragraph({
                alignment: cellAlignment,
                children: [
                  new TextRun({
                    text: content
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
      rows: [headerRow, ...rows],
      width: {
        size: 100,
        type: WidthType.PERCENTAGE
      },
      layout: TableLayoutType.FIXED
    });
  }

  /**
   * @method setImageInfos
   * @description 设置图片信息
   * @param {Array} imageInfos - 图片信息数组
   */
  setImageInfos(imageInfos) {
    this.imageInfos = Array.isArray(imageInfos) ? imageInfos : [];
    console.log(`设置了${this.imageInfos.length}个图片信息`);
  }

  /**
   * @method createImagePlaceholder
   * @description 当图片无法插入时创建占位符
   * @param {string} altText - 替代文本
   * @return {Object} 占位符段落
   */
  createImagePlaceholder(altText) {
    // 确保AlignmentType可用
    const alignment = AlignmentType ? AlignmentType.CENTER : 'center';

    return new Paragraph({
      children: [
        new TextRun({
          text: `[图片: ${altText || '无可用描述'}]`,
          italics: true,
          color: '999999'
        })
      ],
      alignment: alignment,
      spacing: { before: 120, after: 120 }
    });
  }

  /**
   * @method parse
   * @description 使用marked解析Markdown文本
   * @param {string} markdown - Markdown文本内容
   * @return {Array} 解析后的标记数组
   */
  parse(markdown) {
    try {
      if (!markdown || typeof markdown !== 'string') {
        console.warn('无效的Markdown内容:', markdown);
        return [{type: 'paragraph', text: '文档为空或格式错误', raw: ''}];
      }

      console.log('开始解析Markdown文本...');

      // 预处理Markdown文本，修复常见的格式问题
      const fixedMarkdown = this.preprocessMarkdown(markdown);
      console.log('预处理后的Markdown (截取前100字符):', fixedMarkdown.substring(0, 100) + '...');

      // 使用marked.lexer解析Markdown
      let tokens = [];
      try {
        tokens = marked.lexer(fixedMarkdown);
        console.log(`使用marked.lexer成功解析出${tokens.length}个标记`);
      } catch (error) {
        console.warn('使用marked.lexer解析失败，尝试手动解析:', error);

        // 如果marked解析失败，尝试手动解析基本元素
        const lines = fixedMarkdown.split(/\r?\n/);
        tokens = this.manualParsing(lines);
      }

      // 检查是否成功提取了内容
      if (!tokens || tokens.length === 0) {
        console.warn('未能提取任何内容，返回默认段落');
        return [{type: 'paragraph', text: '无法解析文档内容', raw: ''}];
      }

      // 后处理标记，修复潜在问题
      const processedTokens = this.postprocessTokens(tokens);
      console.log(`解析完成，返回${processedTokens.length}个处理后的标记`);

      return processedTokens;
    } catch (error) {
      console.error('解析Markdown时发生错误:', error);
      return [{type: 'paragraph', text: `解析错误: ${error.message}`, raw: ''}];
    }
  }

  /**
   * @method manualParsing
   * @description 手动解析Markdown基本元素（当marked解析失败时使用）
   * @param {Array} lines - Markdown文本按行分割的数组
   * @return {Array} 解析后的标记数组
   */
  manualParsing(lines) {
    const tokens = [];
    let codeBlock = null;

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();

      // 处理代码块
      if (line.startsWith('```')) {
        if (codeBlock) {
          // 结束代码块
          tokens.push({
            type: 'code',
            text: codeBlock.content.join('\n'),
            lang: codeBlock.lang,
            raw: codeBlock.raw.join('\n') + '\n```'
          });
          codeBlock = null;
        } else {
          // 开始代码块
          const lang = line.substring(3).trim();
          codeBlock = {
            lang: lang,
            content: [],
            raw: [line]
          };
        }
        continue;
      }

      // 如果在代码块内，继续收集内容
      if (codeBlock) {
        codeBlock.content.push(line);
        codeBlock.raw.push(line);
        continue;
      }

      // 处理标题
      if (line.startsWith('#')) {
        const match = line.match(/^(#{1,6})\s+(.+)$/);
        if (match) {
          tokens.push({
            type: 'heading',
            depth: match[1].length,
            text: match[2],
            raw: line
          });
          continue;
        }
      }

      // 处理普通段落
      if (line.length > 0) {
        tokens.push({
          type: 'paragraph',
          text: line,
          raw: line
        });
      }
    }

    return tokens;
  }

  /**
   * @method postprocessTokens
   * @description 后处理标记，修复潜在问题
   * @param {Array} tokens - 需要后处理的标记数组
   * @return {Array} 处理后的标记数组
   */
  postprocessTokens(tokens) {
    // 目前只是简单返回，后续可以增加更多处理逻辑
    return tokens;
  }

  /**
   * @method parseMarkdown
   * @param {string} markdown - Markdown文本内容
   * @returns {Promise<Array>} - docx段落对象数组
   */
  async parseMarkdown(markdown) {
    try {
      // 使用自定义解析方法
      const tokens = this.parse(markdown);
      console.log("解析得到的tokens数量:", tokens.length);

    // 预处理，提取所有脚注定义
    this.extractFootnotes(tokens);

    const elements = [];
      // 安全处理所有token
      for (let i = 0; i < tokens.length; i++) {
        const token = tokens[i];

      // 跳过已经处理过的脚注定义
        if (token?.type === 'footnoteDefinition') {
        continue;
      }

        // 安全处理每个token
      try {
        const el = await this.convertToken(token);
        if (Array.isArray(el)) {
          elements.push(...el);
        } else if (el) {
          elements.push(el);
        }
      } catch (error) {
          console.error(`处理token失败 (索引 ${i}):`, error, token);

          // 添加错误信息段落，但继续处理其他token
        elements.push(new Paragraph({
            style: "Normal",
            children: [
              new TextRun({
                text: `处理错误: ${error.message}`,
                color: "FF0000"
              })
            ]
        }));
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
    } catch (error) {
      console.error("解析Markdown失败:", error);
      // 返回一个包含错误信息的段落，让文档仍然能够生成
      return [
        new Paragraph({
          text: "Markdown解析失败",
          style: "Heading1"
        }),
        new Paragraph({
          text: `错误信息: ${error.message}`,
          style: "Normal"
        }),
        new Paragraph({
          text: "请检查您的Markdown语法是否正确。",
          style: "Normal"
        })
      ];
    }
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
   * @returns {Promise<Paragraph|Array<Paragraph>|Table|null>}
   */
  async convertToken(token) {
    // 检查无效token
    if (!token || token.type === undefined) {
      console.warn('遇到无效token:', token);
      return null;
    }

    // 捕获任何可能出现的错误，避免整个转换过程因单个token失败而中断
    try {
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
            return await this.createImagePlaceholder({
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
        return await this.createImagePlaceholder(token);
      case 'footnote':
        return this.createFootnote(token);
      case 'link':
          // 链接应该被处理为内联元素，直接返回包含链接的段落
          return this.createParagraph(token);
        case 'html':
          // 对HTML内容进行基本处理，转为普通文本
          if (token.text) {
            return this.createParagraph({ type: 'paragraph', text: token.text.replace(/<[^>]*>/g, '') });
          }
          return null;
      default:
        // 如果是未知类型，尝试作为普通文本处理
          console.warn(`遇到未处理的token类型: ${token.type}`, token);
        if (token.text) {
          return this.createParagraph({ type: 'paragraph', text: token.text });
        }
        return null;
      }
    } catch (error) {
      console.error(`处理token时出错 (类型: ${token?.type}):`, error);
      // 返回一个错误段落，而不是中断整个转换过程
      return new Paragraph({
        style: "Normal",
        children: [
          new TextRun({
            text: `转换错误: ${error.message}`,
            color: "FF0000"
          })
        ]
      });
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
    // 确保文本是字符串
    if (typeof text !== 'string') {
      text = String(text || '');
    }

    const tokens = [];
    let currentText = text;

    // 调试信息
    console.log('处理内联文本:', text.slice(0, 50) + (text.length > 50 ? '...' : ''));

    // 正则表达式匹配不同的格式
    // 更宽松的格式匹配规则，不强制要求前后有空格
    const boldItalicRegex = /\*\*\*(.*?)\*\*\*/g;
    const boldRegex = /\*\*(.*?)\*\*/g;
    const italicRegex = /(?<!\*)\*(.*?)\*(?!\*)/g;  // 避免与**冲突
    const strikeRegex = /~~(.*?)~~/g;
    const underlineRegex = /<u>(.*?)<\/u>/g;
    const codeRegex = /`(.*?)`/g;
    const linkRegex = /\[(.*?)\]\((.*?)\)/g;
    const footnoteRefRegex = /\[\^(.*?)\]/g;  // 脚注引用格式 [^1]

    // 为了跟踪已处理的部分，创建一个标记数组
    const allTokens = [];

    // 函数：添加标记并在tempText中标记为已处理
    function addToken(match, type, text, url = null) {
      // 确保文本内容不包含前后空格或标记
      let cleanText = text;
      const start = match.index;
      const end = start + match[0].length;

      allTokens.push({
        type,
        text: cleanText,
        url: url,
        start,
        end,
        original: match[0]
      });

      // 在currentText中用空格替换已处理的部分以避免重复处理
      currentText = currentText.substring(0, start) + ' '.repeat(match[0].length) + currentText.substring(end);
    }

    // 按顺序处理不同的格式化标记
    // 粗斜体
    let match;
    while ((match = boldItalicRegex.exec(text)) !== null) {
      addToken(match, 'bold_italic', match[1]);
    }

    // 粗体
    boldRegex.lastIndex = 0;
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

    // 检查有无任何标记被识别
    if (allTokens.length === 0) {
      // 如果没有识别到任何Markdown格式，直接返回纯文本
      return [{
        type: 'text',
        text: text
      }];
    }

    for (const token of allTokens) {
      // 如果当前标记前有未处理的文本，先添加为普通文本
      if (token.start > lastEnd) {
        const plainText = text.substring(lastEnd, token.start);
        if (plainText.trim() !== '') {
        finalTokens.push({
          type: 'text',
            text: plainText
        });
        }
      }

      // 添加当前标记
      finalTokens.push({
        type: token.type,
        text: token.text,
        url: token.url
      });

      // 更新lastEnd
      lastEnd = token.end;
    }

    // 添加最后一个标记后的文本（如果有）
    if (lastEnd < text.length) {
      const plainText = text.substring(lastEnd);
      if (plainText.trim() !== '') {
      finalTokens.push({
        type: 'text',
          text: plainText
      });
      }
    }

    // 输出处理结果数量
    console.log(`处理结果: 识别了 ${allTokens.length} 个格式标记, 生成了 ${finalTokens.length} 个文本段落`);

    return finalTokens;
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

  /**
   * @method preprocessMarkdown
   * @description 预处理Markdown文本，但保留原始的标记格式
   * @param {string} text - 原始Markdown文本
   * @return {string} 修正后的Markdown文本
   */
  preprocessMarkdown(text) {
    if (!text) return '';

    // 标准化行尾并移除可能的BOM和零宽字符
    let fixed = text
      .replace(/\r\n?/g, '\n')
      .replace(/^\ufeff/, '')
      .replace(/\u200b/g, '');

    // 替换中文引号为英文引号（避免解析问题）
    fixed = fixed
      .replace(/[""]|"/g, '"')
      .replace(/['']|'/g, "'");

    // 确保标题#后有空格
    fixed = fixed.replace(/^(#{1,6})([^#\s])/gm, '$1 $2');

    // 确保无序列表项有空格
    fixed = fixed.replace(/^(\s*)([*+-])([^\s])/gm, '$1$2 $3');

    // 确保有序列表项有空格
    fixed = fixed.replace(/^(\s*)(\d+\.)([^\s])/gm, '$1$2 $3');

    // 修复表格格式
    fixed = fixed.replace(/\|(\S)/g, '| $1');
    fixed = fixed.replace(/(\S)\|/g, '$1 |');

    // 记录预处理结果
    console.log(`预处理Markdown：${fixed.substring(0, 100)}...`);

    return fixed;
  }

  /**
   * @method init
   * @description 初始化Md2Docx转换器
   */
  init() {
    // 由于我们已经通过npm安装并正确导入了docx库，不再需要检查
    console.log('docx.js库已正确导入');

    // 检查marked库是否已加载
    if (typeof marked === 'undefined') {
      console.error('错误: marked库未加载！请确保在使用Md2Docx前先加载marked库。');
      return false;
    }

    console.log(`正在初始化Md2Docx转换器，使用marked版本: ${marked.version || '未知'}`);

    // 配置marked选项
    const markedOptions = {
      gfm: true,            // 启用GitHub风格Markdown
      breaks: true,         // 将换行符转换为<br>
      pedantic: false,      // 不使用原始markdown.pl的bug
      mangle: false,        // 不转义自动链接的邮箱地址
      headerIds: true,      // 为标题生成id
      silent: false,        // 不忽略错误
      smartLists: true,     // 使用更智能的列表行为
      smartypants: false    // 不使用更智能的标点符号（引号、破折号等）
    };

    // 设置marked选项
    if (marked.setOptions) {
      marked.setOptions(markedOptions);
    }

    // 简单的测试，确保marked正常工作
    try {
      const testMarkdown = '**粗体文本** *斜体文本*';
      const testHtml = marked.parse(testMarkdown);
      console.log('Marked解析测试成功:', testHtml);
    } catch (error) {
      console.error('测试marked解析时出错:', error);
      return false;
    }

    console.log('Md2Docx转换器初始化完成');
    return true;
  }

  /**
   * @method convertMdToTokens
   * @description 将Markdown文本转换为标记数组
   * @param {string} markdown - Markdown文本
   * @return {Array} 解析后的标记数组
   */
  convertMdToTokens(markdown) {
    // 确保marked库已正确加载
    if (typeof marked === 'undefined') {
      console.error('错误: marked库未加载！');
      return [];
    }

    try {
      // 如果尚未初始化，则进行初始化
      if (!this._initialized) {
        const initResult = this.init();
        if (!initResult) {
          console.error('Md2Docx初始化失败，无法转换文档');
          return [];
        }
        this._initialized = true;
      }

      // 解析Markdown文本
      const tokens = this.parse(markdown);

      // 转换标记为内部格式
      return this.transformTokens(tokens);
    } catch (error) {
      console.error('转换Markdown为标记时发生错误:', error);
      return [{
        type: 'paragraph',
        content: `错误: ${error.message}`,
        raw: ''
      }];
    }
  }

  /**
   * @method transformTokens
   * @description 将marked的标记转换为内部标记格式
   * @param {Array} tokens - marked解析的标记
   * @return {Array} 转换后的内部标记格式
   */
  transformTokens(tokens) {
    if (!tokens || !Array.isArray(tokens)) {
      return [];
    }

    // 统一标记格式
    return tokens.map(token => {
      // 转换标记类型
      switch (token.type) {
        case 'heading':
          return {
            type: 'heading',
            depth: token.depth,
            content: token.text,
            raw: token.raw
          };

        case 'paragraph':
          return {
            type: 'paragraph',
            content: token.text,
            raw: token.raw
          };

        case 'code':
          return {
            type: 'code',
            language: token.lang || '',
            content: token.text,
            raw: token.raw
          };

        case 'blockquote':
          return {
            type: 'blockquote',
            content: token.text,
            raw: token.raw
          };

        case 'list':
          return {
            type: 'list',
            ordered: token.ordered,
            items: token.items.map(item => item.text),
            raw: token.raw
          };

        case 'list_item':
          return {
            type: 'list_item',
            content: token.text,
            raw: token.raw
          };

        case 'table':
          return {
            type: 'table',
            header: token.header,
            rows: token.rows || [],
            raw: token.raw
          };

        case 'html':
          return {
            type: 'html',
            content: token.text,
            raw: token.raw
          };

        case 'hr':
          return {
            type: 'hr',
            raw: token.raw || '---'
          };

        case 'image':
          return {
            type: 'image',
            href: token.href,
            title: token.title || '',
            text: token.text || '',
            raw: token.raw
          };

        case 'link':
          return {
            type: 'link',
            href: token.href,
            title: token.title || '',
            content: token.text,
            raw: token.raw
          };

        case 'text':
          return {
            type: 'text',
            content: token.text,
            raw: token.raw
          };

        default:
          console.warn(`未知标记类型: ${token.type}`, token);
          return {
            type: 'unknown',
            content: token.text || token.raw || '',
            raw: token.raw || ''
          };
      }
    }).filter(token => token !== null);
  }

  /**
   * @method getDefaultStyles
   * @description 获取默认样式配置
   * @returns {Object} 默认样式对象
   */
  getDefaultStyles() {
    try {
      // 直接返回defaultStyles模块引入的对象
      const styles = JSON.parse(JSON.stringify(defaultStyles));
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
}

// 导出Md2Docx类
export { Md2Docx };
