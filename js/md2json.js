// Markdown转DOCX Token JSON工具

/**
 * @class Md2Json
 * @description 将Markdown转换为DOCX Token的JSON格式
 */
class Md2Json {
  constructor() {
    this.tokens = [];
    this.footnotes = new Map(); // 存储脚注定义映射
  }

  /**
   * @method parseTextStyle
   * @description 解析文本中的样式，如加粗、斜体等
   * @param {string} text - 要解析的文本
   * @returns {Object} 包含样式和纯文本的对象
   */
  parseTextStyle(text) {
    // 初始化样式对象
    const style = {
      bold: false,
      italic: false,
      strike: false,
      code: false,
      underline: false,
      superscript: false,
      subscript: false,
      content: text // 保存对应的纯文本内容
    };

    // 检查是否是加粗斜体文本 (*** 或 ___)
    if ((text.startsWith('***') && text.endsWith('***')) ||
        (text.startsWith('___') && text.endsWith('___'))) {
      style.bold = true;
      style.italic = true;
      style.content = text.slice(3, -3);
    }
    // 检查是否是加粗文本 (** 或 __)
    else if ((text.startsWith('**') && text.endsWith('**')) ||
        (text.startsWith('__') && text.endsWith('__'))) {
      style.bold = true;
      style.content = text.slice(2, -2);
    }
    // 检查是否是斜体文本 (* 或 _)
    else if ((text.startsWith('*') && text.endsWith('*') && !text.startsWith('**')) ||
             (text.startsWith('_') && text.endsWith('_') && !text.startsWith('__'))) {
      style.italic = true;
      style.content = text.slice(1, -1);
    }
    // 检查是否是删除线文本 (~~)
    else if (text.startsWith('~~') && text.endsWith('~~')) {
      style.strike = true;
      style.content = text.slice(2, -2);
    }
    // 检查是否是行内代码 (`)
    else if (text.startsWith('`') && text.endsWith('`')) {
      style.code = true;
      style.content = text.slice(1, -1);
    }
    // 检查是否是下划线文本 (<ins></ins>)
    else if (text.startsWith('<ins>') && text.endsWith('</ins>')) {
      style.underline = true;
      style.content = text.slice(5, -6);
    }
    // 检查是否是上标文本 (<sup></sup>)
    else if (text.startsWith('<sup>') && text.endsWith('</sup>')) {
      style.superscript = true;
      style.content = text.slice(5, -6);
    }
    // 检查是否是下标文本 (<sub></sub>)
    else if (text.startsWith('<sub>') && text.endsWith('</sub>')) {
      style.subscript = true;
      style.content = text.slice(5, -6);
    }

    return style;
  }

  /**
   * @method parseInlineStyles
   * @description 解析段落中的内联样式
   * @param {string} text - 段落文本
   * @returns {Object} 包含纯文本内容和样式数组的对象
   */
  parseInlineStyles(text) {
    // 使用正则表达式匹配各种标记，确保先匹配更长的模式
    const regex = /(\*\*\*.*?\*\*\*|\*\*.*?\*\*|\*.*?\*|___.*?___|__.*?__|_.*?_|~~.*?~~|`.*?`|<ins>.*?<\/ins>|<sup>.*?<\/sup>|<sub>.*?<\/sub>)|([^*_~`<]+)/g;

    const matches = text.match(regex) || [text];
    const styleSegments = matches.map(part => this.parseTextStyle(part));

    // 提取纯文本内容，不添加额外空格
    const fullContent = styleSegments.map(segment => segment.content).join('');

    return {
      fullContent: fullContent,
      styles: styleSegments
    };
  }

  /**
   * @method parseHeadingWithNumber
   * @description 解析标题文本中的数字序号
   * @param {string} headingText - 标题文本
   * @returns {Object} 包含序号和内容的对象
   */
  parseHeadingWithNumber(headingText) {
    // 匹配开头的数字+点+空格模式
    const match = headingText.match(/^(\d+)(\.\s*)(.+)$/);

    if (match) {
      return {
        number: match[1], // 只保留数字部分，如 "1"
        content: match[3].trim(), // 实际标题内容
        hasNumber: true
      };
    }

    return {
      number: "",
      content: headingText,
      hasNumber: false
    };
  }

  /**
   * @method parseBlockquoteLevel
   * @description 解析引用的嵌套级别
   * @param {string} text - 引用文本
   * @returns {Object} 包含嵌套级别和内容的对象
   */
  parseBlockquoteLevel(text) {
    let level = 0;
    let content = text;

    // 计算开头有多少个 '>' 符号
    while (content.startsWith('>')) {
      level++;
      content = content.substring(1).trim();
    }

    return {
      level: level,
      content: content
    };
  }

  /**
   * @method parseTableAlignment
   * @description 解析表格对齐方式
   * @param {string} line - 表格分隔行
   * @returns {Array} 列对齐方式数组
   */
  parseTableAlignment(line) {
    const columns = line.split('|').filter(col => col.trim() !== '');
    const alignments = columns.map(col => {
      col = col.trim();
      // 支持所有标准对齐符号
      if (col.match(/^:?-+:?$/)) {
        if (col.startsWith(':') && col.endsWith(':')) {
          return 'center';
        } else if (col.startsWith(':')) {
          return 'left';
        } else if (col.endsWith(':')) {
          return 'right';
        }
      }
      return 'left';  // 默认左对齐
    });

    return alignments;
  }

  /**
   * @method parseTableCell
   * @description 解析表格单元格内容
   * @param {string} cellText - 单元格文本
   * @returns {Object} 解析后的单元格对象
   */
  parseTableCell(cellText) {
    // 转义特殊字符
    const escapedText = cellText.replace(/\|/g, '\\|').replace(/`/g, '\\`');

    // 解析内联样式
    const parsedStyles = this.parseInlineStyles(cellText);

    // 如果没有任何样式，则不包含inlineStyles字段
    const result = {
      rawText: cellText,
      fullContent: parsedStyles.fullContent
    };

    // 只有当存在样式时才添加inlineStyles
    if (parsedStyles.styles.some(style =>
      style.bold || style.italic || style.strike ||
      style.code || style.underline ||
      style.superscript || style.subscript
    )) {
      result.inlineStyles = parsedStyles.styles;
    }

    return result;
  }

  /**
   * @method createTable
   * @description 创建表格对象
   * @param {Array} headers - 表头数组
   * @param {Array} alignments - 对齐方式数组
   * @param {Array} rows - 数据行数组
   * @returns {Object} 表格对象
   */
  createTable(headers, alignments, rows) {
    return {
      type: 'table',
      headers: headers.map(header => this.parseTableCell(header)),
      alignments: alignments,
      rows: rows.map(row => row.map(cell => this.parseTableCell(cell)))
    };
  }

  /**
   * @method parseLinkOrImage
   * @description 解析链接或图片
   * @param {string} text - 链接或图片文本
   * @returns {Object} 解析后的链接或图片对象
   */
  parseLinkOrImage(text) {
    // 嵌套的图片链接语法: [![alt](url "title")](clickUrl "title")
    const nestedImgRegex = /\[!\[(.*?)\]\((.*?)(?:\s+"(.*?)")?\)\]\((.*?)(?:\s+"(.*?)")?\)/;
    // 图片语法: ![alt](url "title")
    const imgRegex = /!\[(.*?)\]\((.*?)(?:\s+"(.*?)")?\)/;
    // 链接语法: [text](url "title")
    const linkRegex = /\[(.*?)\]\((.*?)(?:\s+"(.*?)")?\)/;
    // 自动链接语法: <url>
    const autoLinkRegex = /<(https?:\/\/.*?)>/;

    let match;

    if (match = text.match(nestedImgRegex)) {
      return {
        type: 'image',
        alt: match[1] || '',
        url: match[2] || '',
        clickUrl: match[4] || '',
        title: match[5] || match[3] || '', // 优先使用外层链接的title，如果没有则使用内层图片的title
        rawText: text
      };
    } else if (match = text.match(imgRegex)) {
      return {
        type: 'image',
        alt: match[1] || '',
        url: match[2] || '',
        clickUrl: '',
        title: match[3] || '',
        rawText: text
      };
    } else if (match = text.match(linkRegex)) {
      return {
        type: 'hyperlink',
        text: match[1] || '',
        url: match[2] || '',
        title: match[3] || '',
        rawText: text
      };
    } else if (match = text.match(autoLinkRegex)) {
      return {
        type: 'hyperlink',
        text: match[1],
        url: match[1],
        title: '',
        rawText: text
      };
    }

    return null;
  }

  /**
   * @method parseFootnote
   * @description 解析段落中的脚注引用
   * @param {string} text - 段落文本
   * @returns {Object|null} 包含脚注信息的对象，如果没有脚注则返回null
   */
  parseFootnote(text) {
    // 匹配脚注引用格式 [^id]
    const footnoteRegex = /\[\^(.*?)\]/;
    const match = text.match(footnoteRegex);

    if (match) {
      const footnoteSign = match[1]; // 脚注标识符
      const fullContent = text.replace(match[0], '').trim(); // 去除脚注标记的内容

      return {
        hasFootnote: true,
        footnoteSign: footnoteSign,
        fullContent: fullContent,
        footnoteMatch: match[0] // 完整的脚注标记
      };
    }

    return {
      hasFootnote: false
    };
  }

  /**
   * @method isFootnoteDefinition
   * @description 检查一行是否是脚注定义行
   * @param {string} line - 要检查的行
   * @returns {Object|null} 包含脚注定义信息的对象，如果不是脚注定义则返回null
   */
  isFootnoteDefinition(line) {
    // 匹配脚注定义格式 [^id]: content
    const match = line.match(/^\[\^(.*?)\]:\s*(.*?)$/);

    if (match) {
      return {
        isDefinition: true,
        footnoteSign: match[1], // 脚注标识符
        footnoteContent: match[2].trim() // 脚注内容
      };
    }

    return {
      isDefinition: false
    };
  }

  /**
   * @method convert
   * @description 将Markdown文本转换为JSON格式的tokens
   * @param {string} markdownText - Markdown文本
   * @returns {Object} tokens的JSON对象
   */
  convert(markdownText) {
    this.tokens = [];
    this.footnotes.clear();

    // 解析Markdown文本
    const lines = markdownText.split('\n');

    // 创建文档结构
    const document = {
      type: 'document',
      children: []
    };

    // 首先收集所有脚注定义
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      const footnoteDefInfo = this.isFootnoteDefinition(line);

      if (footnoteDefInfo.isDefinition) {
        this.footnotes.set(footnoteDefInfo.footnoteSign, footnoteDefInfo.footnoteContent);
      }
    }

    // 解析每一行，生成相应的tokens
    let currentHeadingLevel = 0;
    let inCodeBlock = false;
    let codeBlockLanguage = '';
    let codeBlockContent = [];
    let inList = false;
    let listItems = [];
    let listType = '';
    let currentListItem = null;
    let currentListLevel = -1;
    let tableHeaders = [];
    let tableAlignments = [];
    let inTable = false;
    let tableRows = [];

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];

      // 处理代码块
      if (line.startsWith('```')) {
        if (!inCodeBlock) {
          // 开始代码块
          inCodeBlock = true;
          codeBlockLanguage = line.slice(3).trim();
          codeBlockContent = [];
        } else {
          // 结束代码块
          const codeContent = codeBlockContent.join('\n');

          // 如果在列表项中，将代码块添加到当前列表项的blocks中
          if (inList && currentListItem) {
            currentListItem.blocks = currentListItem.blocks || [];
            currentListItem.blocks.push({
              type: 'code_block',
              rawText: '```' + codeBlockLanguage + '\n' + codeContent + '\n```',
              language: codeBlockLanguage,
              fullContent: codeContent
            });
          } else {
            // 否则添加到文档中
            document.children.push({
              type: 'code_block',
              rawText: '```' + codeBlockLanguage + '\n' + codeContent + '\n```',
              language: codeBlockLanguage,
              fullContent: codeContent
            });
          }

          inCodeBlock = false;
          codeBlockLanguage = '';
          codeBlockContent = [];
        }
        continue;
      }

      // 在代码块内
      if (inCodeBlock) {
        codeBlockContent.push(line);
        continue;
      }

      // 处理标题
      if (line.startsWith('#')) {
        const match = line.match(/^(#{1,6})\s+(.+)$/);
        if (match) {
          // 如果之前有列表，添加到文档
          if (inList) {
            document.children.push({
              type: 'list',
              listType: listType,
              items: listItems
            });
            inList = false;
            listItems = [];
            listType = '';
            currentListItem = null;
            currentListLevel = -1;
          }

          // 如果之前有表格，添加到文档
          if (inTable) {
            document.children.push(this.createTable(tableHeaders, tableAlignments, tableRows));
            inTable = false;
            tableHeaders = [];
            tableAlignments = [];
            tableRows = [];
          }

          const level = match[1].length;
          const rawText = match[2].trim();

          // 解析标题中的数字序号
          const parsedHeading = this.parseHeadingWithNumber(rawText);
          // 解析标题中的内联样式
          const parsedStyles = this.parseInlineStyles(parsedHeading.content);

          document.children.push({
            type: 'heading',
            rawText: line.trim(),
            level: level,
            hasNumber: parsedHeading.hasNumber,
            number: parsedHeading.number,
            fullContent: parsedStyles.fullContent,
            inlineStyles: parsedStyles.styles
          });

          currentHeadingLevel = level;
          continue;
        }
      }

      // 处理无序列表
      if (line.match(/^\s*[-*+]\s+.+/)) {
        const rawText = line.trim();
        const indentMatch = line.match(/^(\s*)/);
        const indentLevel = indentMatch ? Math.floor(indentMatch[0].length / 2) : 0; // 转为整数
        const text = line.replace(/^\s*[-*+]\s+/, '').trim();
        // 解析列表项中的内联样式
        const parsedStyles = this.parseInlineStyles(text);

        if (!inList || listType !== 'unordered') {
          // 开始新的无序列表
          if (inList) {
            // 先添加之前的列表
            document.children.push({
              type: 'list',
              listType: listType,
              items: listItems
            });
          }

          inList = true;
          listType = 'unordered';

          // 创建新的列表项
          const newItem = {
            type: 'listItem',
            rawText: rawText,
            fullContent: parsedStyles.fullContent,
            level: indentLevel,
            inlineStyles: parsedStyles.styles,
            blocks: [] // 添加blocks数组存放附加内容
          };

          listItems = [newItem];
          currentListItem = newItem;
          currentListLevel = indentMatch ? indentMatch[0].length : 0;
        } else {
          // 继续添加到当前列表
          const newItem = {
            type: 'listItem',
            rawText: rawText,
            fullContent: parsedStyles.fullContent,
            level: indentLevel,
            inlineStyles: parsedStyles.styles,
            blocks: [] // 添加blocks数组存放附加内容
          };

          listItems.push(newItem);
          currentListItem = newItem;
          currentListLevel = indentMatch ? indentMatch[0].length : 0;
        }
        continue;
      }

      // 处理有序列表
      if (line.match(/^\s*\d+\.\s+.+/)) {
        const listItemMatch = line.match(/^(\s*)(\d+)(\.\s+)(.+)$/);
        if (listItemMatch) {
          const rawText = line.trim();
          const indentLevel = Math.floor((listItemMatch[1].length) / 2); // 转为整数
          const number = listItemMatch[2]; // 只保留数字部分
          const text = listItemMatch[4].trim();
          // 解析列表项中的内联样式
          const parsedStyles = this.parseInlineStyles(text);

          if (!inList || listType !== 'ordered') {
            // 开始新的有序列表
            if (inList) {
              // 先添加之前的列表
              document.children.push({
                type: 'list',
                listType: listType,
                items: listItems
              });
            }

            inList = true;
            listType = 'ordered';

            // 创建新的列表项
            const newItem = {
              type: 'listItem',
              rawText: rawText,
              fullContent: parsedStyles.fullContent,
              level: indentLevel,
              number: number,
              inlineStyles: parsedStyles.styles,
              blocks: [] // 添加blocks数组存放附加内容
            };

            listItems = [newItem];
            currentListItem = newItem;
            currentListLevel = listItemMatch[1].length;
          } else {
            // 继续添加到当前列表
            const newItem = {
              type: 'listItem',
              rawText: rawText,
              fullContent: parsedStyles.fullContent,
              level: indentLevel,
              number: number,
              inlineStyles: parsedStyles.styles,
              blocks: [] // 添加blocks数组存放附加内容
            };

            listItems.push(newItem);
            currentListItem = newItem;
            currentListLevel = listItemMatch[1].length;
          }
          continue;
        }
      }

      // 处理列表项的附加内容（缩进大于当前列表项的非列表行）
      if (inList && line.trim() !== '' &&
          !line.match(/^\s*[-*+]\s+.+/) &&
          !line.match(/^\s*\d+\.\s+.+/)) {

        // 检查是否是属于列表项的内容（缩进量大于等于列表项）
        const indentMatch = line.match(/^(\s*)/);
        const indentLevel = indentMatch ? indentMatch[0].length : 0;

        // 如果缩进量大于等于当前列表项，视为附加内容
        if (indentLevel >= currentListLevel && currentListItem) {
          const trimmedLine = line.trim();

          // 检查是否是链接或图片
          const linkOrImage = this.parseLinkOrImage(trimmedLine);
          if (linkOrImage) {
            currentListItem.blocks.push(linkOrImage);
          } else {
            // 普通段落
            const parsedStyles = this.parseInlineStyles(trimmedLine);
            currentListItem.blocks.push({
              type: 'paragraph',
              rawText: trimmedLine,
              fullContent: parsedStyles.fullContent,
              inlineStyles: parsedStyles.styles
            });
          }

          continue;
        }
      }

      // 结束列表（空行或新的不同缩进的内容）
      if (inList && (line.trim() === '' ||
          (line.match(/^(\s*)/)[0].length < currentListLevel &&
           !line.match(/^\s*[-*+]\s+.+/) &&
           !line.match(/^\s*\d+\.\s+.+/)))) {
        document.children.push({
          type: 'list',
          listType: listType,
          items: listItems
        });
        inList = false;
        listItems = [];
        listType = '';
        currentListItem = null;
        currentListLevel = -1;

        // 如果不是空行，需要重新处理当前行
        if (line.trim() !== '') {
          i--; // 回退一行，在下一次循环中重新处理
          continue;
        }
      }

      // 处理引用
      if (line.startsWith('>')) {
        // 如果之前有列表，添加到文档
        if (inList) {
          document.children.push({
            type: 'list',
            listType: listType,
            items: listItems
          });
          inList = false;
          listItems = [];
          listType = '';
          currentListItem = null;
          currentListLevel = -1;
        }

        // 解析引用级别
        const parsedQuote = this.parseBlockquoteLevel(line.trim());
        const text = parsedQuote.content;
        // 解析引用中的内联样式
        const parsedStyles = this.parseInlineStyles(text);

        document.children.push({
          type: 'blockquote',
          rawText: line.trim(),
          level: parsedQuote.level,
          fullContent: parsedStyles.fullContent,
          inlineStyles: parsedStyles.styles
        });
        continue;
      }

      // 处理表格
      if (line.includes('|') && line.trim().startsWith('|')) {
        // 如果之前有列表，添加到文档
        if (inList) {
          document.children.push({
            type: 'list',
            listType: listType,
            items: listItems
          });
          inList = false;
          listItems = [];
          listType = '';
          currentListItem = null;
          currentListLevel = -1;
        }

        // 检测是否是表格分隔行
        if (line.match(/^\|\s*[-:]+[-\s:]*\|/)) {
          // 解析对齐方式
          tableAlignments = this.parseTableAlignment(line);
          continue;
        }

        const cells = line.split('|')
          .filter(cell => cell.trim() !== '')  // 过滤掉空单元格
          .map(cell => cell.trim());

        // 如果还没有表头且没有对齐信息，这是表头行
        if (!inTable) {
          inTable = true;
          tableHeaders = cells;
        } else {
          // 否则是数据行
          tableRows.push(cells);
        }
        continue;
      } else if (inTable) {
        // 如果不是表格行但之前有表格，结束表格
        document.children.push(this.createTable(tableHeaders, tableAlignments, tableRows));
        inTable = false;
        tableHeaders = [];
        tableAlignments = [];
        tableRows = [];
      }

      // 处理水平线
      if (line.match(/^\s*[-*_]{3,}\s*$/)) {
        // 如果之前有列表，添加到文档
        if (inList) {
          document.children.push({
            type: 'list',
            listType: listType,
            items: listItems
          });
          inList = false;
          listItems = [];
          listType = '';
          currentListItem = null;
          currentListLevel = -1;
        }

        document.children.push({
          type: 'horizontal_rule',
          rawText: line.trim()
        });
        continue;
      }

      // 处理链接和图片
      const linkOrImage = this.parseLinkOrImage(line.trim());
      if (linkOrImage) {
        document.children.push(linkOrImage);
        continue;
      }

      // 跳过脚注定义行，因为已经在之前的循环中处理过
      const footnoteDefInfo = this.isFootnoteDefinition(line.trim());
      if (footnoteDefInfo.isDefinition) {
        continue;
      }

      // 处理普通段落和脚注引用
      if (line.trim() !== '') {
        // 如果不在列表中或者缩进小于当前列表项
        if (!inList || (line.match(/^(\s*)/)[0].length < currentListLevel)) {
          // 如果之前有列表，添加到文档
          if (inList) {
            document.children.push({
              type: 'list',
              listType: listType,
              items: listItems
            });
            inList = false;
            listItems = [];
            listType = '';
            currentListItem = null;
            currentListLevel = -1;
          }

          const rawText = line.trim();

          // 检查是否包含脚注引用
          const footnoteInfo = this.parseFootnote(rawText);

          if (footnoteInfo.hasFootnote) {
            // 获取脚注内容
            const footnoteContent = this.footnotes.get(footnoteInfo.footnoteSign) || '';

            // 解析内联样式
            const parsedStyles = this.parseInlineStyles(footnoteInfo.fullContent);

            document.children.push({
              type: 'footnote', // 使用特定的脚注类型
              rawText: rawText,
              hasNumber: false,
              number: '',
              fullContent: parsedStyles.fullContent,
              inlineStyles: parsedStyles.styles,
              footnoteSign: footnoteInfo.footnoteSign,
              footnoteContent: footnoteContent
            });
          } else {
            // 检查段落是否以数字序号开头
            const parsedParagraph = this.parseHeadingWithNumber(rawText);
            // 解析段落中的内联样式
            const parsedStyles = this.parseInlineStyles(parsedParagraph.content);

            document.children.push({
              type: 'paragraph',
              rawText: rawText,
              hasNumber: parsedParagraph.hasNumber,
              number: parsedParagraph.number,
              fullContent: parsedStyles.fullContent,
              inlineStyles: parsedStyles.styles
            });
          }
        }
      }
    }

    // 结束时如果还有未处理的列表，添加到文档
    if (inList) {
      document.children.push({
        type: 'list',
        listType: listType,
        items: listItems
      });
    }

    // 结束时如果还有未处理的表格，添加到文档
    if (inTable) {
      document.children.push(this.createTable(tableHeaders, tableAlignments, tableRows));
    }

    return document;
  }
}

// 导出模块
export { Md2Json };
