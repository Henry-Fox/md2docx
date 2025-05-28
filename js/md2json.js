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
      italics: false,
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
      style.italics = true;
      style.content = text.slice(3, -3);
    }
    // 检查是否是加粗文本 (** 或 __)
    else if ((text.startsWith('**') && text.endsWith('**')) ||
        (text.startsWith('__') && text.endsWith('__'))) {
      style.bold = true;
      style.italics = false; // 确保加粗文本不是斜体
      style.content = text.slice(2, -2);
    }
    // 检查是否是斜体文本 (* 或 _)
    else if ((text.startsWith('*') && text.endsWith('*') && !text.startsWith('**')) ||
             (text.startsWith('_') && text.endsWith('_') && !text.startsWith('__'))) {
      style.italics = true;
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
    // 检查是否是下划线文本 (<u></u>)
    else if (text.startsWith('<u>') && text.endsWith('</u>')) {
      style.underline = true;
      style.content = text.slice(3, -4);
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
   * @description 解析行内样式，支持<u>、<sup>、<sub>等HTML标签和Markdown语法，避免内容重复，支持嵌套
   * @param {string} text - 要解析的文本
   * @returns {Array} 解析后的样式数组
   */
  parseInlineStyles(text) {
    const styles = [];
    // 先递归处理HTML标签（u、sup、sub）
    const tagRegex = /<(u|sup|sub)>([\s\S]*?)<\/\1>/ig;
    let match;
    let cursor = 0;
    let hasTag = false;
    while ((match = tagRegex.exec(text)) !== null) {
      hasTag = true;
      // 递归处理标签前的所有内容（而不是整体推入）
      if (match.index > cursor) {
        const before = text.slice(cursor, match.index);
        if (before) {
          styles.push(...this.parseInlineStyles(before));
        }
      }
      // 处理标签内的内容（递归，支持嵌套）
      const inner = this.parseInlineStyles(match[2]);
      inner.forEach(style => {
        if (match[1].toLowerCase() === 'u') style.underline = true;
        if (match[1].toLowerCase() === 'sup') style.superscript = true;
        if (match[1].toLowerCase() === 'sub') style.subscript = true;
      });
      styles.push(...inner);
      cursor = match.index + match[0].length;
    }
    // 处理最后剩余的文本
    if (cursor < text.length) {
      const after = text.slice(cursor);
      if (after) {
        styles.push(...this._parseMarkdownInlineStyles(after));
      }
    }
    // 如果没有任何HTML标签，直接用Markdown样式分割
    if (!hasTag) {
      return this._mergePlainTextRuns(this._parseMarkdownInlineStyles(text));
    }
    // 合并相邻普通文本段
    return this._mergePlainTextRuns(styles);
  }

  /**
   * @private
   * @method _parseMarkdownInlineStyles
   * @description 递归处理Markdown语法的行内样式，支持嵌套
   * @param {string} text
   * @param {Object} parentStyle - 父级样式（用于递归叠加）
   * @returns {Array}
   */
  _parseMarkdownInlineStyles(text, parentStyle = {}) {
    const styles = [];
    if (!text) return styles;
    // 支持的样式正则，顺序很重要（先处理复杂的）
    const regex = /\*\*\*([\s\S]+?)\*\*\*|___([\s\S]+?)___|\*\*([\s\S]+?)\*\*|__([\s\S]+?)__|\*([\s\S]+?)\*|_([\s\S]+?)_|~~([\s\S]+?)~~|`([^`]+?)`/g;
    let lastIndex = 0;
    let match;
    while ((match = regex.exec(text)) !== null) {
      // 普通文本
      if (match.index > lastIndex) {
        styles.push({
          bold: !!parentStyle.bold,
          italics: !!parentStyle.italics,
          strike: !!parentStyle.strike,
          code: !!parentStyle.code,
          underline: !!parentStyle.underline,
          superscript: !!parentStyle.superscript,
          subscript: !!parentStyle.subscript,
          content: text.slice(lastIndex, match.index)
        });
      }
      // 匹配到的样式，递归处理内容
      let style = {
        bold: !!parentStyle.bold,
        italics: !!parentStyle.italics,
        strike: !!parentStyle.strike,
        code: !!parentStyle.code,
        underline: !!parentStyle.underline,
        superscript: !!parentStyle.superscript,
        subscript: !!parentStyle.subscript
      };
      let innerContent = '';
      if (match[1] !== undefined) { // ***加粗斜体***
        style.bold = true;
        style.italics = true;
        innerContent = match[1];
      } else if (match[2] !== undefined) { // ___加粗斜体___
        style.bold = true;
        style.italics = true;
        innerContent = match[2];
      } else if (match[3] !== undefined) { // **加粗**
        style.bold = true;
        innerContent = match[3];
      } else if (match[4] !== undefined) { // __加粗__
        style.bold = true;
        innerContent = match[4];
      } else if (match[5] !== undefined) { // *斜体*
        style.italics = true;
        innerContent = match[5];
      } else if (match[6] !== undefined) { // _斜体_
        style.italics = true;
        innerContent = match[6];
      } else if (match[7] !== undefined) { // ~~删除线~~
        style.strike = true;
        innerContent = match[7];
      } else if (match[8] !== undefined) { // `代码`
        style.code = true;
        innerContent = match[8];
      }
      // 递归处理内容，叠加样式
      if (innerContent) {
        const inner = this._parseMarkdownInlineStyles(innerContent, style);
        styles.push(...inner);
      }
      lastIndex = match.index + match[0].length;
    }
    // 剩余普通文本
    if (lastIndex < text.length) {
      styles.push({
        bold: !!parentStyle.bold,
        italics: !!parentStyle.italics,
        strike: !!parentStyle.strike,
        code: !!parentStyle.code,
        underline: !!parentStyle.underline,
        superscript: !!parentStyle.superscript,
        subscript: !!parentStyle.subscript,
        content: text.slice(lastIndex)
      });
    }
    return styles;
  }

  /**
   * @private
   * @method _mergePlainTextRuns
   * @description 合并相邻的普通文本（所有样式属性全为false的段）
   * @param {Array} styles
   * @returns {Array}
   */
  _mergePlainTextRuns(styles) {
    if (!styles.length) return styles;
    const merged = [styles[0]];
    for (let i = 1; i < styles.length; i++) {
      const prev = merged[merged.length - 1];
      const curr = styles[i];
      // 判断所有样式属性都为false
      const isPlain = s => !s.bold && !s.italics && !s.strike && !s.code && !s.underline && !s.superscript && !s.subscript;
      if (isPlain(prev) && isPlain(curr)) {
        prev.content += curr.content;
      } else {
        merged.push(curr);
      }
    }
    return merged;
  }

  /**
   * @method parseParagraph
   * @description 解析段落
   * @param {string} text - 要解析的文本
   * @returns {Object} 解析后的段落对象
   */
  parseParagraph(text) {
    const parsedStyles = this.parseInlineStyles(text);
    return {
      type: 'paragraph',
      rawText: text,
      hasNumber: false,
      number: '',
      fullContent: parsedStyles.map(style => style.content).join(''),
      inlineStyles: parsedStyles
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
      fullContent: parsedStyles.map(style => style.content).join('')
    };

    // 只有当存在样式时才添加inlineStyles
    if (parsedStyles.some(style =>
      style.bold || style.italics || style.strike ||
      style.code || style.underline ||
      style.superscript || style.subscript
    )) {
      result.inlineStyles = parsedStyles;
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
        text: match[3] || match[1] || '', // 优先使用title，如果没有则使用text
        url: match[2] || '',
        rawText: text
      };
    } else if (match = text.match(autoLinkRegex)) {
      return {
        type: 'hyperlink',
        text: match[1],
        url: match[1],
        rawText: text
      };
    }

    return null;
  }

  /**
   * @method parseFootnote
   * @description 解析段落中的脚注引用和定义
   * @param {string} text - 段落文本
   * @param {string} nextLine - 下一行文本（用于检查脚注定义）
   * @returns {Object|null} 包含脚注信息的对象，如果没有脚注则返回null
   */
  parseFootnote(text, nextLine = '') {
    // 匹配脚注引用格式 [^id]
    const footnoteRegex = /\[\^(.*?)\]/;
    const match = text.match(footnoteRegex);

    if (match) {
      const footnoteSign = match[1]; // 脚注标识符
      const fullContent = text.replace(match[0], '').trim(); // 去除脚注标记的内容

      // 检查下一行是否是脚注定义
      const footnoteDefRegex = new RegExp(`^\\[\\^${footnoteSign}\\]:\\s*(.*?)$`);
      const defMatch = nextLine.trim().match(footnoteDefRegex);

      let footnoteContent = '';
      if (defMatch) {
        footnoteContent = defMatch[1].trim();
      }

      // 解析内联样式
      const parsedStyles = this.parseInlineStyles(fullContent);

      return {
        type: 'footnote',
        rawText: text,
        fullContent: fullContent,
        footnoteSign: footnoteSign,
        footnoteContent: footnoteContent,
        inlineStyles: parsedStyles,
        hasDefinition: !!defMatch
      };
    }

    return null;
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
      const nextLine = i + 1 < lines.length ? lines[i + 1] : '';

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
            fullContent: parsedStyles.map(style => style.content).join(''),
            inlineStyles: parsedStyles
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

        // 检查是否是任务列表
        const taskListMatch = line.match(/^\s*[-*+]\s+\[([ x])\]\s+(.+)$/i);
        if (taskListMatch) {
          const text = taskListMatch[2].trim();
          // 解析列表项中的内联样式
          const parsedStyles = this.parseInlineStyles(text);

          document.children.push({
            type: 'task',
            rawText: rawText,
            isChecked: taskListMatch[1].toLowerCase() === 'x',
            fullContent: text,
            inlineStyles: parsedStyles
          });
          continue;
        }

        const text = line.replace(/^\s*[-*+]\s+/, '').trim();
        // 解析列表项中的内联样式
        const parsedStyles = this.parseInlineStyles(text);

        document.children.push({
          type: 'list',
          rawText: rawText,
          level: indentLevel,
          hasNumber: false,
          fullContent: text,
          inlineStyles: parsedStyles
        });
        continue;
      }

      // 处理有序列表
      if (line.match(/^\s*\d+\.\s+.+/)) {
        const listItemMatch = line.match(/^(\s*)(\d+)(\.\s+)(.+)$/);
        if (listItemMatch) {
          const rawText = line.trim();
          const indentLevel = Math.floor((listItemMatch[1].length) / 2); // 转为整数
          const text = listItemMatch[4].trim();
          // 解析列表项中的内联样式
          const parsedStyles = this.parseInlineStyles(text);

          document.children.push({
            type: 'list',
            rawText: rawText,
            level: indentLevel,
            hasNumber: true,
            fullContent: text,
            inlineStyles: parsedStyles
          });
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
              fullContent: parsedStyles.map(style => style.content).join(''),
              inlineStyles: parsedStyles
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
          fullContent: parsedStyles.map(style => style.content).join(''),
          inlineStyles: parsedStyles
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
          const footnoteInfo = this.parseFootnote(rawText, nextLine);

          if (footnoteInfo) {
            document.children.push(footnoteInfo);
            // 如果下一行是脚注定义，跳过它
            if (footnoteInfo.hasDefinition) {
              i++;
            }
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
              fullContent: parsedStyles.map(style => style.content).join(''),
              inlineStyles: parsedStyles
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
