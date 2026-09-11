import SimpleMd2Docx from "./simpleMd2Docx.js";
import SimpleMd2Pdf from "./simpleMd2Pdf.js";
import { templateManager } from "./templateManager.js";

/**
 * ExportManager - 统一的导出管理器
 * 
 * 负责协调Markdown到DOCX和PDF的导出流程
 * 支持使用模板系统配置输出格式
 */
class ExportManager {
  constructor() {
    this.docxConverter = new SimpleMd2Docx();
    this.pdfConverter = new SimpleMd2Pdf();
  }

  /**
   * 导出为DOCX格式
   * @param {string} markdown - Markdown文本
   * @param {object} template - 模板配置(可选,默认使用当前激活模板)
   * @returns {Promise<void>}
   */
  async exportDocx(markdown, template = null) {
    if (!markdown || !markdown.trim()) {
      throw new Error("Markdown内容不能为空");
    }

    const tpl = template || templateManager.getActive();
    console.log("使用模板导出DOCX:", tpl.name);
    
    this.docxConverter.setTemplate(tpl);
    await this.docxConverter.convertToDocxDirect(markdown);
  }

  /**
   * 导出为PDF格式
   * @param {string} markdown - Markdown文本
   * @param {object} template - 模板配置(可选,默认使用当前激活模板)
   * @returns {Promise<void>}
   */
  async exportPdf(markdown, template = null) {
    if (!markdown || !markdown.trim()) {
      throw new Error("Markdown内容不能为空");
    }

    const tpl = template || templateManager.getActive();
    console.log("使用模板导出PDF:", tpl.name);
    
    this.pdfConverter.setTemplate(tpl);
    await this.pdfConverter.convertToPdfDirect(markdown);
  }

  /**
   * 获取支持的导出格式列表
   * @returns {Array<string>}
   */
  getSupportedFormats() {
    return ["docx", "pdf"];
  }

  /**
   * 检查格式是否支持
   * @param {string} format - 格式名称
   * @returns {boolean}
   */
  isFormatSupported(format) {
    return this.getSupportedFormats().includes(format.toLowerCase());
  }

  /**
   * 通用导出方法
   * @param {string} markdown - Markdown文本
   * @param {string} format - 导出格式 ('docx' 或 'pdf')
   * @param {object} template - 模板配置(可选)
   * @returns {Promise<void>}
   */
  async export(markdown, format, template = null) {
    const formatLower = format.toLowerCase();

    if (!this.isFormatSupported(formatLower)) {
      throw new Error(`不支持的导出格式: ${format}`);
    }

    switch (formatLower) {
      case "docx":
        return await this.exportDocx(markdown, template);
      case "pdf":
        return await this.exportPdf(markdown, template);
      default:
        throw new Error(`未实现的导出格式: ${format}`);
    }
  }
}

// 导出单例实例
export const exportManager = new ExportManager();
export default ExportManager;
