import { PDFDocument, StandardFonts, rgb, degrees } from "pdf-lib";
import fontkit from "@pdf-lib/fontkit";
import { saveAs } from "file-saver";
import { marked } from "marked";
import { templateManager } from "./templateManager.js";

/**
 * SimpleMd2Pdf - 模板驱动的Markdown到PDF转换器
 * 
 * 中文字体策略说明:
 * - 使用 Noto Sans SC (Google 开源字体，OFL-1.1 许可证)
 * - 通过 CDN 动态加载，避免打包体积过大
 * - 字体加载后缓存在内存中，提高后续导出速度
 * - 完整支持简体中文、繁体中文及常用标点符号
 * 
 * 支持的Markdown元素:
 * - 标题(H1-H6)
 * - 段落
 * - 列表(有序/无序/任务列表)
 * - 表格
 * - 代码块
 * - 引用块
 * - 水平线
 * - 链接
 * - 文本格式(粗体/斜体/删除线)
 */
class SimpleMd2Pdf {
  constructor() {
    console.log("SimpleMd2Pdf 初始化");
    this.pdfStyles = null;
    this.pdfDoc = null;
    this.currentPage = null;
    this.currentY = 0;
    this.fonts = {};
    this.pageMargin = { top: 72, bottom: 72, left: 72, right: 72 }; // 默认1英寸边距
    this.fontCache = {}; // 字体缓存
  }

  setTemplate(template) {
    this.pdfStyles = templateManager.toPdfStyles(template);
  }

  getStyles() {
    return this.pdfStyles || templateManager.toPdfStyles(templateManager.getActive());
  }

  /**
   * 将mm转换为points (PDF标准单位)
   */
  mmToPoints(mm) {
    return mm * 2.83465;
  }

  /**
   * 获取页面可用宽度
   */
  getPageWidth() {
    const styles = this.getStyles();
    return styles.pageWidth - styles.pageMargin.left - styles.pageMargin.right;
  }

  /**
   * 获取页面可用高度
   */
  getPageHeight() {
    const styles = this.getStyles();
    return styles.pageHeight - styles.pageMargin.top - styles.pageMargin.bottom;
  }

  /**
   * 检查是否需要新页面
   */
  needsNewPage(requiredHeight) {
    const styles = this.getStyles();
    const bottomMargin = styles.pageMargin.bottom;
    return this.currentY - requiredHeight < bottomMargin;
  }

  /**
   * 添加新页面
   */
  addNewPage() {
    const styles = this.getStyles();
    this.currentPage = this.pdfDoc.addPage([styles.pageWidth, styles.pageHeight]);
    this.currentY = styles.pageHeight - styles.pageMargin.top;
    return this.currentPage;
  }

  /**
   * 获取文本宽度
   */
  getTextWidth(text, font, fontSize) {
    return font.widthOfTextAtSize(text, fontSize);
  }

  /**
   * 将文本分行以适应页面宽度
   */
  wrapText(text, font, fontSize, maxWidth) {
    const words = text.split(" ");
    const lines = [];
    let currentLine = "";

    for (const word of words) {
      const testLine = currentLine ? `${currentLine} ${word}` : word;
      const width = this.getTextWidth(testLine, font, fontSize);

      if (width > maxWidth && currentLine) {
        lines.push(currentLine);
        currentLine = word;
      } else {
        currentLine = testLine;
      }
    }

    if (currentLine) {
      lines.push(currentLine);
    }

    return lines;
  }

  /**
   * 绘制文本段落
   */
  async drawText(text, options = {}) {
    const {
      font = this.fonts.regular,
      fontSize = 12,
      color = rgb(0, 0, 0),
      bold = false,
      italic = false,
      align = "left",
      indent = 0,
      firstLineIndent = 0,  // 首行缩进（只应用于第一行）
      lineSpacing = null,   // 行距（points），null 则使用 fontSize * 1.5
    } = options;

    const styles = this.getStyles();
    const bodyStyle = styles.body || {};
    
    // 使用模板的行距，如果未指定则回退到 fontSize * 1.5
    const actualLineSpacing = lineSpacing !== null ? lineSpacing : 
                             (bodyStyle.lineSpacing || fontSize * 1.5);
    
    const actualFont = bold ? this.fonts.bold : italic ? this.fonts.italic : font;
    
    // 计算最大宽度（考虑首行缩进）
    const maxWidthFirstLine = this.getPageWidth() - indent - firstLineIndent;
    const maxWidthOtherLines = this.getPageWidth() - indent;
    
    // CJK-aware 文本换行：逐字符测量宽度
    // 中文/日文/韩文几乎无空格，需要按字符换行；拉丁文保留单词换行
    const lines = [];
    if (text.trim()) {
      let currentLine = '';
      let currentMaxWidth = maxWidthFirstLine;
      let isFirstLine = true;
      
      // 检测是否为 CJK 字符
      const isCJK = (char) => {
        const code = char.charCodeAt(0);
        return (code >= 0x4E00 && code <= 0x9FFF) ||   // CJK Unified Ideographs
               (code >= 0x3400 && code <= 0x4DBF) ||   // CJK Extension A
               (code >= 0x20000 && code <= 0x2A6DF) || // CJK Extension B
               (code >= 0x3040 && code <= 0x309F) ||   // Hiragana
               (code >= 0x30A0 && code <= 0x30FF) ||   // Katakana
               (code >= 0xAC00 && code <= 0xD7AF);     // Hangul
      };
      
      // 检测文本是否主要为 CJK（>30% CJK 字符）
      const cjkChars = text.split('').filter(isCJK).length;
      const isCJKText = (cjkChars / text.length) > 0.3;
      
      if (isCJKText) {
        // CJK 模式：逐字符换行
        for (let i = 0; i < text.length; i++) {
          const char = text[i];
          const testLine = currentLine + char;
          const width = this.getTextWidth(testLine, actualFont, fontSize);
          
          if (width > currentMaxWidth && currentLine) {
            // 当前行已满，推入并开始新行
            lines.push({ text: currentLine, isFirst: isFirstLine });
            currentLine = char;
            isFirstLine = false;
            currentMaxWidth = maxWidthOtherLines;
          } else {
            currentLine = testLine;
          }
        }
      } else {
        // 拉丁文模式：按单词换行
        const words = text.split(' ');
        for (const word of words) {
          const testLine = currentLine ? `${currentLine} ${word}` : word;
          const width = this.getTextWidth(testLine, actualFont, fontSize);
          
          if (width > currentMaxWidth && currentLine) {
            lines.push({ text: currentLine, isFirst: isFirstLine });
            currentLine = word;
            isFirstLine = false;
            currentMaxWidth = maxWidthOtherLines;
          } else {
            currentLine = testLine;
          }
        }
      }
      
      if (currentLine) {
        lines.push({ text: currentLine, isFirst: isFirstLine });
      }
    }

    for (const lineObj of lines) {
      if (this.needsNewPage(actualLineSpacing)) {
        this.addNewPage();
      }

      // 首行应用首行缩进，其他行只应用常规缩进
      const currentIndent = lineObj.isFirst ? (indent + firstLineIndent) : indent;
      let x = styles.pageMargin.left + currentIndent;
      const textWidth = this.getTextWidth(lineObj.text, actualFont, fontSize);

      if (align === "center") {
        x = styles.pageMargin.left + (this.getPageWidth() - textWidth) / 2;
      } else if (align === "right") {
        x = styles.pageMargin.left + this.getPageWidth() - textWidth;
      }

      this.currentPage.drawText(lineObj.text, {
        x,
        y: this.currentY,
        size: fontSize,
        font: actualFont,
        color,
      });

      this.currentY -= actualLineSpacing;
    }
  }

  /**
   * 绘制标题
   */
  async drawHeading(token) {
    const level = Math.min(token.depth, 6);
    const styleKey = ["title", "h1", "h2", "h3", "h4", "h5"][level - 1];
    const styles = this.getStyles();
    const headingStyle = styles[styleKey] || styles.body;

    // 标题前增加间距
    this.currentY -= headingStyle.fontSize;

    await this.drawText(token.text, {
      fontSize: headingStyle.fontSize,
      bold: headingStyle.bold,
      color: this.parseColor(headingStyle.color),
      align: headingStyle.alignment,
    });

    // 标题后增加间距
    this.currentY -= headingStyle.fontSize * 0.5;
  }

  /**
   * 解析颜色字符串为RGB
   */
  parseColor(colorStr) {
    if (!colorStr || colorStr === "000000") {
      return rgb(0, 0, 0);
    }
    const r = parseInt(colorStr.substring(0, 2), 16) / 255;
    const g = parseInt(colorStr.substring(2, 4), 16) / 255;
    const b = parseInt(colorStr.substring(4, 6), 16) / 255;
    return rgb(r, g, b);
  }

  /**
   * 绘制段落
   */
  async drawParagraph(token) {
    const styles = this.getStyles();
    const bodyStyle = styles.body;

    // 解析内联格式
    const text = this.extractPlainText(token.text);

    await this.drawText(text, {
      fontSize: bodyStyle.fontSize,
      color: this.parseColor(bodyStyle.color),
      align: bodyStyle.alignment,
      firstLineIndent: bodyStyle.firstLineIndent,  // 首行缩进
      lineSpacing: bodyStyle.lineSpacing,           // 使用模板行距
    });

    // 段落后增加间距
    this.currentY -= bodyStyle.fontSize * 0.5;
  }

  /**
   * 提取纯文本(移除HTML标签)
   */
  extractPlainText(html) {
    return html.replace(/<[^>]*>/g, "");
  }

  /**
   * 绘制列表
   */
  async drawList(token, ordered = false) {
    const styles = this.getStyles();
    const bodyStyle = styles.body;
    let index = 1;

    for (const item of token.items) {
      const bullet = ordered ? `${index}. ` : "• ";
      const text = this.extractPlainText(item.text);
      const fullText = bullet + text;

      await this.drawText(fullText, {
        fontSize: bodyStyle.fontSize,
        color: this.parseColor(bodyStyle.color),
        indent: 20,
      });

      index++;
    }

    this.currentY -= bodyStyle.fontSize * 0.5;
  }

  /**
   * 绘制表格
   */
  async drawTable(token) {
    const styles = this.getStyles();
    const bodyStyle = styles.body;
    const cellPadding = 5;
    const rowHeight = bodyStyle.fontSize * 2;
    const pageWidth = this.getPageWidth();
    
    // 计算列宽
    const numCols = token.header?.length || 0;
    if (numCols === 0) return;
    
    const colWidth = (pageWidth - cellPadding * 2 * numCols) / numCols;

    // 绘制表头
    if (token.header) {
      let x = styles.pageMargin.left;
      
      for (const cell of token.header) {
        const text = this.extractPlainText(cell.text);
        
        // 绘制单元格边框
        this.currentPage.drawRectangle({
          x,
          y: this.currentY - rowHeight,
          width: colWidth + cellPadding * 2,
          height: rowHeight,
          borderColor: rgb(0, 0, 0),
          borderWidth: 1,
        });

        // 绘制单元格文本
        this.currentPage.drawText(text, {
          x: x + cellPadding,
          y: this.currentY - rowHeight / 2 - bodyStyle.fontSize / 2,
          size: bodyStyle.fontSize,
          font: this.fonts.bold,
          color: rgb(0, 0, 0),
        });

        x += colWidth + cellPadding * 2;
      }

      this.currentY -= rowHeight;
    }

    // 绘制表格行
    if (token.rows) {
      for (const row of token.rows) {
        if (this.needsNewPage(rowHeight)) {
          this.addNewPage();
        }

        let x = styles.pageMargin.left;

        for (const cell of row) {
          const text = this.extractPlainText(cell.text);

          // 绘制单元格边框
          this.currentPage.drawRectangle({
            x,
            y: this.currentY - rowHeight,
            width: colWidth + cellPadding * 2,
            height: rowHeight,
            borderColor: rgb(0, 0, 0),
            borderWidth: 1,
          });

          // 绘制单元格文本
          this.currentPage.drawText(text, {
            x: x + cellPadding,
            y: this.currentY - rowHeight / 2 - bodyStyle.fontSize / 2,
            size: bodyStyle.fontSize,
            font: this.fonts.regular,
            color: rgb(0, 0, 0),
          });

          x += colWidth + cellPadding * 2;
        }

        this.currentY -= rowHeight;
      }
    }

    this.currentY -= bodyStyle.fontSize * 0.5;
  }

  /**
   * 绘制代码块
   */
  async drawCodeBlock(token) {
    const styles = this.getStyles();
    const codeStyle = styles.code || styles.body;
    const bgColor = rgb(0.95, 0.95, 0.95);
    const padding = 10;
    const lines = token.text.split("\n");
    const lineHeight = codeStyle.fontSize * 1.3;
    const blockHeight = lines.length * lineHeight + padding * 2;

    if (this.needsNewPage(blockHeight)) {
      this.addNewPage();
    }

    // 绘制背景
    this.currentPage.drawRectangle({
      x: styles.pageMargin.left,
      y: this.currentY - blockHeight,
      width: this.getPageWidth(),
      height: blockHeight,
      color: bgColor,
      borderColor: rgb(0.8, 0.8, 0.8),
      borderWidth: 1,
    });

    // 绘制代码文本
    let y = this.currentY - padding;
    for (const line of lines) {
      this.currentPage.drawText(line, {
        x: styles.pageMargin.left + padding,
        y: y - lineHeight,
        size: codeStyle.fontSize,
        font: this.fonts.monospace,
        color: rgb(0, 0, 0),
      });
      y -= lineHeight;
    }

    this.currentY -= blockHeight + codeStyle.fontSize * 0.5;
  }

  /**
   * 绘制引用块
   */
  async drawBlockquote(token) {
    const styles = this.getStyles();
    const bodyStyle = styles.body;
    const text = this.extractPlainText(token.text);

    // 绘制左侧竖线
    const barWidth = 4;
    const barColor = rgb(0.7, 0.7, 0.7);
    const textLines = this.wrapText(text, this.fonts.italic, bodyStyle.fontSize, this.getPageWidth() - 30);
    const blockHeight = textLines.length * bodyStyle.fontSize * 1.5;

    this.currentPage.drawRectangle({
      x: styles.pageMargin.left,
      y: this.currentY - blockHeight,
      width: barWidth,
      height: blockHeight,
      color: barColor,
    });

    // 绘制引用文本
    await this.drawText(text, {
      fontSize: bodyStyle.fontSize,
      italic: true,
      color: rgb(0.3, 0.3, 0.3),
      indent: 20,
    });

    this.currentY -= bodyStyle.fontSize * 0.5;
  }

  /**
   * 绘制水平线
   */
  async drawHorizontalRule() {
    const styles = this.getStyles();
    const lineY = this.currentY - 10;

    this.currentPage.drawLine({
      start: { x: styles.pageMargin.left, y: lineY },
      end: { x: styles.pageMargin.left + this.getPageWidth(), y: lineY },
      thickness: 1,
      color: rgb(0.8, 0.8, 0.8),
    });

    this.currentY -= 20;
  }

  /**
   * 获取输出文件名
   */
  getOutputFilename(markdownTokens) {
    const fallback = "document";
    const titleToken = markdownTokens.find((token) => token.type === "heading" && token.depth === 1);
    const rawTitle = titleToken?.text || fallback;
    const filename = this.sanitizeFilename(rawTitle) || fallback;
    return `${filename}.pdf`;
  }

  /**
   * 清理文件名
   */
  sanitizeFilename(value) {
    return value
      .replace(/[<>:"/\\|?*\x00-\x1F]/g, "")
      .replace(/\s+/g, " ")
      .trim()
      .substring(0, 200);
  }

  /**
   * 加载中文字体
   */
  async loadCJKFonts() {
    try {
      // 使用 Google Fonts CDN 加载 Noto Sans SC (v40)
      const fontUrls = {
        regular: 'https://fonts.gstatic.com/s/notosanssc/v40/k3kCo84MPvpLmixcA63oeAL7Iqp5IZJF9bmaG9_FnYw.ttf',
        bold: 'https://fonts.gstatic.com/s/notosanssc/v40/k3kCo84MPvpLmixcA63oeAL7Iqp5IZJF9bmaGzjCnYw.ttf',
      };

      console.log("正在加载中文字体...");

      // 加载常规字体
      if (!this.fontCache.regular) {
        const regularResp = await fetch(fontUrls.regular);
        if (!regularResp.ok) throw new Error('字体加载失败');
        this.fontCache.regular = await regularResp.arrayBuffer();
      }

      // 加载粗体字体
      if (!this.fontCache.bold) {
        const boldResp = await fetch(fontUrls.bold);
        if (!boldResp.ok) throw new Error('粗体字体加载失败');
        this.fontCache.bold = await boldResp.arrayBuffer();
      }

      // 注册 fontkit
      this.pdfDoc.registerFontkit(fontkit);

      // 嵌入字体
      const regular = await this.pdfDoc.embedFont(this.fontCache.regular);
      const bold = await this.pdfDoc.embedFont(this.fontCache.bold);

      console.log("中文字体加载完成");

      return {
        regular,
        bold,
        italic: regular, // 使用常规字体代替斜体
        monospace: await this.pdfDoc.embedFont(StandardFonts.Courier), // 代码使用等宽字体
      };
    } catch (error) {
      console.warn("中文字体加载失败，回退到标准字体:", error);
      return {
        regular: await this.pdfDoc.embedFont(StandardFonts.Helvetica),
        bold: await this.pdfDoc.embedFont(StandardFonts.HelveticaBold),
        italic: await this.pdfDoc.embedFont(StandardFonts.HelveticaOblique),
        monospace: await this.pdfDoc.embedFont(StandardFonts.Courier),
      };
    }
  }

  /**
   * 生成PDF字节数组（用于预览等场景，不触发下载）
   * @param {string} markdown - Markdown文本
   * @returns {Promise<Uint8Array>} PDF字节数组
   */
  async generatePdfBytes(markdown) {
    console.log("开始生成PDF字节（预览模式）");

    // 1. 解析Markdown
    const tokens = marked.lexer(markdown);

    // 2. 创建PDF文档
    this.pdfDoc = await PDFDocument.create();
    
    // 3. 加载字体（包含中文字体 Noto Sans SC）
    this.fonts = await this.loadCJKFonts();

    // 4. 获取样式配置
    const styles = this.getStyles();
    
    // 5. 添加首页
    this.addNewPage();

    // 6. 处理每个token
    for (const token of tokens) {
      try {
        switch (token.type) {
          case "heading":
            await this.drawHeading(token);
            break;
          case "paragraph":
            await this.drawParagraph(token);
            break;
          case "list":
            await this.drawList(token, token.ordered);
            break;
          case "table":
            await this.drawTable(token);
            break;
          case "code":
            await this.drawCodeBlock(token);
            break;
          case "blockquote":
            await this.drawBlockquote(token);
            break;
          case "hr":
            await this.drawHorizontalRule();
            break;
          case "space":
            this.currentY -= 10;
            break;
          default:
            console.warn(`未支持的token类型: ${token.type}`);
            break;
        }
      } catch (error) {
        console.error(`处理token时出错:`, token, error);
      }
    }

    // 7. 生成并返回PDF字节
    const pdfBytes = await this.pdfDoc.save();
    console.log("PDF字节生成完成");
    
    return pdfBytes;
  }

  /**
   * 主转换方法: Markdown → PDF（下载文件）
   */
  async convertToPdfDirect(markdown) {
    console.log("开始转换Markdown到PDF");

    // 生成PDF字节
    const pdfBytes = await this.generatePdfBytes(markdown);

    // 解析token用于获取文件名
    const tokens = marked.lexer(markdown);
    
    // 保存文件
    const filename = this.getOutputFilename(tokens);
    const blob = new Blob([pdfBytes], { type: "application/pdf" });
    saveAs(blob, filename);

    console.log("PDF生成完成:", filename);
  }
}

export default SimpleMd2Pdf;
