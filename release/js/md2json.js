// Markdown转DOCX Token JSON工具
import { marked } from 'marked';

/**
 * @class Md2Json
 * @description 将Markdown转换为DOCX Token的JSON格式
 */
class Md2Json {
  constructor() {
    // 配置 marked 选项
    this.markedOptions = {
      gfm: true, // 启用 GitHub Flavored Markdown
      breaks: true, // 启用换行符
      pedantic: false, // 不启用严格模式
      smartLists: true, // 启用智能列表
      smartypants: true, // 启用智能标点
      xhtml: false, // 不启用 XHTML
      headerIds: false // 禁用标题ID生成
    };
  }

  /**
   * @method convert
   * @description 将Markdown文本转换为JSON格式的tokens
   * @param {string} markdownText - Markdown文本
   * @returns {Object} tokens的JSON对象
   */
  convert(markdownText) {
    // 使用 marked 解析 Markdown
    const tokens = marked.lexer(markdownText, this.markedOptions);

    // 打印原始 tokens
    console.log('原始 Markdown Tokens:');
    console.log(JSON.stringify(tokens, null, 2));

    // 创建文档结构
    const document = {
      type: 'document',
      children: []
    };

    // 处理解析后的 tokens
    for (const token of tokens) {
      const convertedToken = this.convertToken(token);
      if (convertedToken) {
        document.children.push(convertedToken);
      }
    }

    // 打印转换后的 JSON
    console.log('转换后的 JSON:');
    console.log(JSON.stringify(document, null, 2));

    return document;
  }

  /**
   * @private
   * @method convertToken
   * @description 转换单个token
   * @param {Object} token - marked token
   * @returns {Object} 转换后的token
   */
  convertToken(token) {
    switch (token.type) {
      case 'paragraph':
        return {
          type: 'paragraph',
          rawText: token.text,
          fullContent: token.text,
          inlineStyles: this.processInlineTokens(token.tokens || [])
        };

      case 'heading':
        const parsedHeading = this.parseHeadingWithNumber(token.text);
        return {
          type: 'heading',
          level: token.depth,
          hasNumber: parsedHeading.hasNumber,
          number: parsedHeading.number,
          fullContent: parsedHeading.content,
          inlineStyles: this.processInlineTokens(token.tokens || [])
        };

      case 'list':
        return {
          type: 'list',
          listType: token.ordered ? 'ordered' : 'unordered',
          items: token.items.map(item => ({
            type: 'list_item',
            fullContent: item.text,
            inlineStyles: this.processInlineTokens(item.tokens || [])
          }))
        };

      case 'code':
        return {
          type: 'code_block',
          language: token.lang,
          fullContent: token.text
        };

      case 'table':
        return {
          type: 'table',
          headers: (token.header || []).map(cell => ({
            rawText: cell.text,
            fullContent: cell.text,
            inlineStyles: this.processInlineTokens(cell.tokens || [])
          })),
          rows: (token.cells || []).map(row =>
            (row || []).map(cell => ({
              rawText: cell.text,
              fullContent: cell.text,
              inlineStyles: this.processInlineTokens(cell.tokens || [])
            }))
          )
        };

      case 'blockquote':
        return {
          type: 'blockquote',
          fullContent: token.text,
          inlineStyles: this.processInlineTokens(token.tokens || [])
        };

      case 'hr':
        return {
          type: 'horizontal_rule'
        };

      case 'space':
        return {
          type: 'paragraph',
          rawText: '',
          fullContent: '',
          inlineStyles: []
        };

      default:
        return null;
    }
  }

  /**
   * @private
   * @method processInlineTokens
   * @description 处理内联样式的tokens
   * @param {Array} tokens - marked tokens数组
   * @returns {Array} 处理后的样式数组
   */
  processInlineTokens(tokens) {
    const styles = [];

    const processToken = (token, parentStyles = {}) => {
      const currentStyles = {
        bold: parentStyles.bold || false,
        italics: parentStyles.italics || false,
        strike: parentStyles.strike || false,
        code: parentStyles.code || false,
        underline: parentStyles.underline || false,
        superscript: parentStyles.superscript || false,
        subscript: parentStyles.subscript || false
      };

      if (token.type === 'text') {
        if (token.text.trim()) {
          styles.push({
            content: token.text,
            ...currentStyles
          });
        }
      } else if (token.type === 'strong') {
        if (token.tokens) {
          token.tokens.forEach(t => processToken(t, { ...currentStyles, bold: true }));
        } else {
          styles.push({
            content: token.text,
            ...currentStyles,
            bold: true
          });
        }
      } else if (token.type === 'em') {
        if (token.tokens) {
          token.tokens.forEach(t => processToken(t, { ...currentStyles, italics: true }));
        } else {
          styles.push({
            content: token.text,
            ...currentStyles,
            italics: true
          });
        }
      } else if (token.type === 'del') {
        if (token.tokens) {
          token.tokens.forEach(t => processToken(t, { ...currentStyles, strike: true }));
        } else {
          styles.push({
            content: token.text,
            ...currentStyles,
            strike: true
          });
        }
      } else if (token.type === 'codespan') {
        styles.push({
          content: token.text,
          ...currentStyles,
          code: true
        });
      } else if (token.type === 'html') {
        const style = {
          content: token.text.replace(/<\/?[^>]+(>|$)/g, ''),
          ...currentStyles
        };
        if (token.text.startsWith('<u>')) {
          style.underline = true;
        } else if (token.text.startsWith('<sup>')) {
          style.superscript = true;
        } else if (token.text.startsWith('<sub>')) {
          style.subscript = true;
        }
        styles.push(style);
      } else if (token.type === 'br') {
        styles.push({
          content: '\n',
          ...currentStyles
        });
      } else if (token.tokens) {
        token.tokens.forEach(t => processToken(t, currentStyles));
      }
    };

    tokens.forEach(token => processToken(token));

    // 如果没有样式，添加一个默认样式
    if (styles.length === 0) {
      styles.push({
        content: tokens.map(t => t.text).join(''),
        bold: false,
        italics: false,
        strike: false,
        code: false,
        underline: false,
        superscript: false,
        subscript: false
      });
    }

    return styles;
  }

  /**
   * @private
   * @method parseHeadingWithNumber
   * @description 解析标题文本中的数字序号
   * @param {string} headingText - 标题文本
   * @returns {Object} 包含序号和内容的对象
   */
  parseHeadingWithNumber(headingText) {
    // 匹配数字序号（包括中文数字）
    const match = headingText.match(/^(\d+)(\.\s*)(.+)$/);
    if (match) {
      return {
        number: match[1],
        content: match[3].trim(),
        hasNumber: true
      };
    }
    return {
      number: "",
      content: headingText,
      hasNumber: false
    };
  }
}

// 导出模块
export { Md2Json };
