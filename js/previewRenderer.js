/**
 * PreviewRenderer - 真正的所见即所得预览渲染器
 * 
 * 使用与导出相同的PDF渲染管道生成预览，确保预览与导出完全一致。
 * 支持:
 * - PDF预览模式（默认，使用pdf-lib + pdf.js渲染）
 * - 加载状态、错误处理
 * - 防抖优化
 * - 模板更改时刷新预览
 */

// 使用 pdfjs-dist/webpack.mjs 零配置入口（自动设置 worker，确保版本匹配）
import * as pdfjsLib from 'pdfjs-dist/webpack.mjs';
const { getDocument } = pdfjsLib;

import SimpleMd2Pdf from './simpleMd2Pdf.js';
import { templateManager } from './templateManager.js';

class PreviewRenderer {
  constructor() {
    this.pdfConverter = new SimpleMd2Pdf();
    this.currentPreviewTask = null;
    this.lastMarkdown = null;
    this.lastTemplateId = null;
  }

  /**
   * 渲染预览
   * @param {string} markdown - Markdown文本
   * @param {HTMLElement} container - 预览容器元素
   * @param {Object} template - 模板配置（可选）
   */
  async renderPreview(markdown, container, template = null) {
    if (!container) {
      console.warn('预览容器不存在');
      return;
    }

    const tpl = template || templateManager.getActive();
    
    // 如果内容和模板都没有变化，跳过
    if (this.lastMarkdown === markdown && this.lastTemplateId === tpl.id) {
      return;
    }

    // 取消正在进行的预览任务
    if (this.currentPreviewTask) {
      this.currentPreviewTask.cancelled = true;
    }

    // 创建新的预览任务
    const previewTask = { cancelled: false };
    this.currentPreviewTask = previewTask;

    // 显示空状态
    if (!markdown || !markdown.trim()) {
      this._showEmptyState(container);
      this.lastMarkdown = markdown;
      this.lastTemplateId = tpl.id;
      return;
    }

    // 显示加载状态
    this._showLoadingState(container);

    try {
      // 生成PDF
      const pdfBytes = await this._generatePdfBytes(markdown, tpl);

      // 如果任务已取消，停止
      if (previewTask.cancelled) {
        return;
      }

      // 渲染PDF页面
      await this._renderPdfPages(pdfBytes, container);

      // 保存状态
      this.lastMarkdown = markdown;
      this.lastTemplateId = tpl.id;
    } catch (error) {
      // 如果任务已取消，不显示错误
      if (!previewTask.cancelled) {
        console.error('预览生成失败:', error);
        this._showErrorState(container, error);
      }
    } finally {
      if (this.currentPreviewTask === previewTask) {
        this.currentPreviewTask = null;
      }
    }
  }

  /**
   * 生成PDF字节数组
   * @private
   */
  async _generatePdfBytes(markdown, template) {
    // 使用与导出完全相同的PDF生成逻辑（包含CJK字体加载）
    this.pdfConverter.setTemplate(template);
    const pdfBytes = await this.pdfConverter.generatePdfBytes(markdown);
    return pdfBytes;
  }

  /**
   * 渲染PDF页面到容器
   * @private
   */
  async _renderPdfPages(pdfBytes, container) {
    // 加载PDF文档
    const loadingTask = getDocument({ data: pdfBytes });
    const pdf = await loadingTask.promise;

    // 清空容器
    container.innerHTML = '';
    container.className = 'preview-container-wysiwyg';

    // 创建页面容器
    const pagesContainer = document.createElement('div');
    pagesContainer.className = 'pdf-pages-container';
    container.appendChild(pagesContainer);

    // 渲染每一页
    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
      const page = await pdf.getPage(pageNum);
      
      // 创建页面容器
      const pageContainer = document.createElement('div');
      pageContainer.className = 'pdf-page-container';
      
      // 创建canvas
      const canvas = document.createElement('canvas');
      canvas.className = 'pdf-page-canvas';
      const context = canvas.getContext('2d');
      
      // 设置渲染比例（根据容器宽度）
      const viewport = page.getViewport({ scale: 1.0 });
      const containerWidth = container.clientWidth - 40; // 减去padding
      const scale = containerWidth / viewport.width;
      const scaledViewport = page.getViewport({ scale });
      
      canvas.width = scaledViewport.width;
      canvas.height = scaledViewport.height;
      
      // 渲染页面
      await page.render({
        canvasContext: context,
        viewport: scaledViewport,
      }).promise;
      
      // 添加到容器
      pageContainer.appendChild(canvas);
      
      // 添加页码
      const pageNumber = document.createElement('div');
      pageNumber.className = 'pdf-page-number';
      pageNumber.textContent = `第 ${pageNum} / ${pdf.numPages} 页`;
      pageContainer.appendChild(pageNumber);
      
      pagesContainer.appendChild(pageContainer);
    }
  }

  /**
   * 显示空状态
   * @private
   */
  _showEmptyState(container) {
    container.innerHTML = `
      <div class="preview-empty-state">
        <span class="material-symbols-outlined preview-empty-icon">description</span>
        <div class="preview-empty-text">在左侧输入 Markdown 内容</div>
        <div class="preview-empty-hint">预览将显示与导出文档完全一致的效果</div>
      </div>
    `;
    container.className = 'preview-container';
  }

  /**
   * 显示加载状态
   * @private
   */
  _showLoadingState(container) {
    container.innerHTML = `
      <div class="preview-loading-state">
        <div class="preview-loading-spinner"></div>
        <div class="preview-loading-text">正在生成预览...</div>
        <div class="preview-loading-hint">使用与导出相同的渲染引擎</div>
      </div>
    `;
    container.className = 'preview-container';
  }

  /**
   * 显示错误状态
   * @private
   */
  _showErrorState(container, error) {
    const errorMessage = error.message || '未知错误';
    const isFontError = errorMessage.includes('font') || errorMessage.includes('字体');
    
    container.innerHTML = `
      <div class="preview-error-state">
        <span class="material-symbols-outlined preview-error-icon">error</span>
        <div class="preview-error-title">预览生成失败</div>
        <div class="preview-error-message">${this._escapeHtml(errorMessage)}</div>
        ${isFontError ? `
          <div class="preview-error-hint">
            <strong>字体加载失败？</strong>
            <p>预览使用标准PDF字体，对中文支持有限但能显示基本内容。导出的PDF/Word文档将使用完整的中文字体。</p>
          </div>
        ` : ''}
      </div>
    `;
    container.className = 'preview-container';
  }

  /**
   * 转义HTML
   * @private
   */
  _escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  /**
   * 清除缓存
   */
  clearCache() {
    this.lastMarkdown = null;
    this.lastTemplateId = null;
  }
}

// 导出单例
export const previewRenderer = new PreviewRenderer();
export default PreviewRenderer;
