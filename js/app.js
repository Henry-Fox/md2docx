import { marked } from 'marked';
import SimpleMd2Docx from './simpleMd2Docx.js';
import { exportManager } from './exportManager.js';
import { initLanguageSwitcher, updateContent, t, tWithVars } from '../src/js/i18n.js';
import { templateManager } from './templateManager.js';
import { parseDocxStyles } from './docxParser.js';
import { checkTemplateFonts, formatMissingFontsWarning, detectOS } from './fontDetector.js';
import { checkDocumentStructure, formatCheckResult } from './documentChecker.js';
import { previewRenderer } from './previewRenderer.js';
import packageInfo from '../package.json';

class App {
  constructor() {
    this.initElements();
    initLanguageSwitcher();
    updateContent();
    this.initEventListeners();
    this.initPreview();
    this.initTemplateUI();
    this.renderVersion();
    this.loadDefaultExample();
    
    // 防抖定时器
    this.previewDebounceTimer = null;
    this.previewDebounceDelay = 500; // 500ms 防抖延迟
  }

  initElements() {
    this.markdownInput    = document.getElementById('markdown-input');
    this.fileInput        = document.getElementById('file-input');
    this.fileNameLabel    = document.getElementById('file-name-label');
    this.dragArea         = document.querySelector('.drag-area');
    this.clearBtn         = document.getElementById('clear-btn');
    this.directConvertBtn = document.getElementById('direct-convert-btn');
    this.exportPdfBtn     = document.getElementById('export-pdf-btn');
    this.previewContainer = document.getElementById('preview-container');
  }

  renderVersion() {
    const versionEl = document.getElementById('app-version');
    if (versionEl) {
      versionEl.textContent = `v${packageInfo.version}`;
    }
  }

  initEventListeners() {
    const customFileBtn = document.getElementById('custom-file-btn');
    if (customFileBtn && this.fileInput) {
      customFileBtn.addEventListener('click', () => this.fileInput.click());
    }

    this.fileInput?.addEventListener('change', (e) => this.handleFileSelect(e));

    if (this.dragArea) {
      this.dragArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        this.dragArea.classList.add('active');
      });
      this.dragArea.addEventListener('dragleave', () => this.dragArea.classList.remove('active'));
      this.dragArea.addEventListener('drop', (e) => {
        e.preventDefault();
        this.dragArea.classList.remove('active');
        const file = e.dataTransfer.files[0];
        if (file) this.readFile(file);
      });
    }

    this.clearBtn?.addEventListener('click', () => this.clearMarkdown());
    this.directConvertBtn?.addEventListener('click', () => this.directConvertToDocx());
    this.exportPdfBtn?.addEventListener('click', () => this.directConvertToPdf());
    this.markdownInput?.addEventListener('input', () => {
      this.updatePreview();
      this.toggleEmptyState();
    });

    // Toolbar formatting buttons
    document.querySelectorAll('.toolbar-btn[data-md-action]').forEach(btn => {
      btn.addEventListener('click', () => this._insertMarkdown(btn.dataset.mdAction));
    });

    // Prompt modal
    document.getElementById('show-prompt-btn')?.addEventListener('click', () => {
      const tpl = templateManager.getActive();
      this._showPromptModal(tpl.name, this._generateLLMPrompt(tpl));
    });

    // ChatGPT and Claude buttons
    document.getElementById('open-chatgpt-btn')?.addEventListener('click', () => {
      window.open('https://chat.openai.com/', '_blank');
    });
    document.getElementById('open-claude-btn')?.addEventListener('click', () => {
      window.open('https://claude.ai/', '_blank');
    });

    // Empty state example buttons
    document.querySelectorAll('[data-example]').forEach(btn => {
      btn.addEventListener('click', () => this.loadExample(btn.dataset.example));
    });

    // Donate panel toggle
    document.getElementById('sidebar-donate-toggle')?.addEventListener('click', () => {
      document.getElementById('sidebar-donate-panel')?.classList.toggle('collapsed');
    });
  }

  _insertMarkdown(action) {
    const ta = this.markdownInput;
    if (!ta) return;
    const s = ta.selectionStart, e = ta.selectionEnd;
    const sel = ta.value.substring(s, e);
    const map = {
      bold:   () => `**${sel || '粗体文字'}**`,
      italic: () => `*${sel || '斜体文字'}*`,
      h1:     () => `# ${sel || '一级标题'}`,
      h2:     () => `## ${sel || '二级标题'}`,
      h3:     () => `### ${sel || '三级标题'}`,
      link:   () => `[${sel || '链接文字'}](url)`,
      ul:     () => `- ${sel || '列表项'}`,
      ol:     () => `1. ${sel || '列表项'}`,
    };
    const fn = map[action];
    if (!fn) return;
    const text = fn();
    ta.value = ta.value.substring(0, s) + text + ta.value.substring(e);
    ta.focus();
    ta.selectionStart = ta.selectionEnd = s + text.length;
    this.updatePreview();
  }

  handleFileSelect(e) {
    const file = e.target.files[0];
    if (!file) return;
    if (this.fileNameLabel) this.fileNameLabel.textContent = file.name;
    this.readFile(file);
  }

  readFile(file) {
    if (file.type !== 'text/markdown' && file.type !== 'text/plain' && !file.name.endsWith('.md')) {
      this.showMessage(t('fileTypeError'), 'error');
      return;
    }
    const reader = new FileReader();
    reader.onload = (e) => {
      this.markdownInput.value = e.target.result;
      this.updatePreview();
    };
    reader.readAsText(file);
  }

  clearMarkdown() {
    this.markdownInput.value = '';
    if (this.fileNameLabel) this.fileNameLabel.textContent = '未选择文件';
    this.updatePreview();
    this.toggleEmptyState();
  }

  loadDefaultExample() {
    if (this.markdownInput) {
      this.markdownInput.value = '';
      this.updatePreview();
      this.toggleEmptyState();
    }
  }

  toggleEmptyState() {
    const emptyState = document.getElementById('editor-empty-state');
    if (!emptyState) return;
    const isEmpty = !this.markdownInput?.value.trim();
    emptyState.style.display = isEmpty ? 'flex' : 'none';
  }

  loadExample(exampleType) {
    const examples = {
      'official-doc': `# 关于印发《XX管理办法》的通知

各有关单位：

为进一步规范XX工作，根据相关法律法规，结合实际情况，现将《XX管理办法》印发给你们，请认真贯彻执行。

## 一、总则

### （一）目的和依据

为规范XX管理工作，根据《XX法》等法律法规，制定本办法。

### （二）适用范围

本办法适用于本市范围内的XX管理工作。

## 二、主要内容

1. 明确管理职责
2. 规范工作流程
3. 强化监督检查

特此通知。

XX单位  
2026年9月11日`,
      'thesis': `# 基于深度学习的图像识别技术研究

## 摘要

本文研究了基于深度学习的图像识别技术，提出了一种改进的卷积神经网络模型。实验结果表明，该方法在标准数据集上取得了良好的识别效果。

**关键词：** 深度学习；图像识别；卷积神经网络

## 1 引言

随着人工智能技术的快速发展，图像识别已成为计算机视觉领域的重要研究方向。本文针对传统方法的局限性，提出了一种基于深度学习的改进方案。

## 2 相关工作

### 2.1 传统图像识别方法

传统方法主要包括特征提取和分类器设计两个步骤，存在特征表达能力有限的问题。

### 2.2 深度学习方法

深度学习通过多层神经网络自动学习特征表示，在图像识别任务上取得了突破性进展。

## 3 研究方法

本文提出的方法包括以下关键技术：

1. 数据预处理和增强
2. 网络结构设计
3. 训练策略优化

## 4 实验结果

### 4.1 数据集

使用CIFAR-10和ImageNet两个标准数据集进行实验验证。

### 4.2 性能评估

| 方法 | 准确率 | 召回率 | F1值 |
|------|--------|--------|------|
| 传统方法 | 85.2% | 82.1% | 83.6% |
| 本文方法 | 92.5% | 91.3% | 91.9% |

## 5 结论

本文提出的方法在图像识别任务上取得了优异的性能，具有良好的应用前景。

## 参考文献

[1] LeCun Y, Bengio Y, Hinton G. Deep learning[J]. nature, 2015.
[2] He K, Zhang X, Ren S, et al. Deep residual learning[C]//CVPR, 2016.`,
      'weekly': `# 技术部周报（2026年第37周）

**报告人：** 张三  
**报告时间：** 2026年9月11日

## 本周工作总结

### 1. 项目开发

- 完成了用户管理模块的后端接口开发
- 实现了权限控制功能，通过了单元测试
- 修复了数据导出功能的3个bug

### 2. 技术优化

- 优化了数据库查询性能，响应时间降低30%
- 重构了文件上传模块，提升了代码可维护性
- 完成了API文档的更新

### 3. 团队协作

- 参加了产品需求评审会议
- 协助新同事解决技术问题
- 完成了代码review工作

## 下周工作计划

### 1. 功能开发

- [ ] 完成消息通知模块开发
- [ ] 实现数据统计报表功能
- [ ] 进行性能测试和优化

### 2. 技术学习

- [ ] 学习Kubernetes容器编排
- [ ] 研究微服务架构最佳实践
- [ ] 参加公司技术分享会

## 遇到的问题

1. **第三方API不稳定：** 外部支付接口偶尔超时，已联系供应商排查
2. **测试环境资源不足：** 需要申请增加服务器配置

## 备注

下周三下午参加外部技术交流活动，可能需要请假半天。`
    };

    const content = examples[exampleType];
    if (content && this.markdownInput) {
      this.markdownInput.value = content;
      this.updatePreview();
      this.toggleEmptyState();
    }
  }

  initPreview() {
    this.updatePreview();
  }

  updatePreview() {
    if (!this.previewContainer) return;
    
    // 清除之前的防抖定时器
    if (this.previewDebounceTimer) {
      clearTimeout(this.previewDebounceTimer);
    }
    
    // 设置新的防抖定时器
    this.previewDebounceTimer = setTimeout(() => {
      this._doUpdatePreview();
    }, this.previewDebounceDelay);
  }

  /**
   * 实际执行预览更新（内部方法）
   * @private
   */
  async _doUpdatePreview() {
    if (!this.previewContainer || !this.markdownInput) return;
    
    const markdown = this.markdownInput.value;
    
    try {
      // 使用新的预览渲染器
      await previewRenderer.renderPreview(
        markdown,
        this.previewContainer,
        templateManager.getActive()
      );
    } catch (error) {
      console.error('更新预览时出错:', error);
      // 错误已在 previewRenderer 中处理
    }
  }

  async directConvertToDocx() {
    const markdown = this.markdownInput.value;
    if (!markdown.trim()) {
      this.showMessage(t('emptyInput'), 'warning');
      return;
    }
    
    // 文档结构检查
    const checkResult = checkDocumentStructure(markdown);
    if (!checkResult.valid || checkResult.warnings.length > 0) {
      const shouldContinue = await this._showDocumentCheckDialog(checkResult);
      if (!shouldContinue) return;
    }
    
    try {
      this.showMessage(t('convertingSimple'), 'info');
      await exportManager.exportDocx(markdown);
      this.showMessage(t('convertSimpleSuccess'), 'success');
    } catch (error) {
      console.error('转换失败:', error);
      this.showMessage(tWithVars('convertSimpleFail', { msg: error.message }), 'error');
    }
  }

  async directConvertToPdf() {
    const markdown = this.markdownInput.value;
    if (!markdown.trim()) {
      this.showMessage(t('emptyInput'), 'warning');
      return;
    }
    
    // 文档结构检查
    const checkResult = checkDocumentStructure(markdown);
    if (!checkResult.valid || checkResult.warnings.length > 0) {
      const shouldContinue = await this._showDocumentCheckDialog(checkResult);
      if (!shouldContinue) return;
    }
    
    try {
      this.showMessage(t('exportingPdf'), 'info');
      await exportManager.exportPdf(markdown);
      this.showMessage(t('exportPdfSuccess'), 'success');
    } catch (error) {
      console.error('PDF导出失败:', error);
      this.showMessage(tWithVars('exportPdfFail', { msg: error.message }), 'error');
    }
  }

  _showDocumentCheckDialog(checkResult) {
    return new Promise((resolve) => {
      const { errors, warnings } = checkResult;
      
      let message = '';
      if (errors.length > 0) {
        message = '❌ 文档存在以下问题:\n\n';
        errors.forEach((err, i) => {
          message += `${i + 1}. ${err.message}\n   ${err.description}\n\n`;
        });
        message += '建议修改后再导出。';
        alert(message);
        resolve(false);
        return;
      }
      
      if (warnings.length > 0) {
        message = '⚠️ 文档结构建议:\n\n';
        warnings.forEach((warn, i) => {
          message += `${i + 1}. ${warn.message}\n   ${warn.description}\n\n`;
        });
        message += '是否仍然继续导出?';
        resolve(confirm(message));
        return;
      }
      
      resolve(true);
    });
  }

  showMessage(message, type = 'info') {
    console.log(`[${type.toUpperCase()}] ${message}`);
    let container = document.getElementById('message-container');
    if (!container) {
      container = document.createElement('div');
      container.id = 'message-container';
      container.style.cssText = 'position:fixed;top:20px;right:20px;z-index:1000;';
      document.body.appendChild(container);
    }
    const el = document.createElement('div');
    el.className = `message message-${type}`;
    el.innerHTML = message;
    const colors = { success: '#4caf50', error: '#f44336', warning: '#ff9800', info: '#2196f3' };
    el.style.cssText = `padding:10px 15px;margin-bottom:10px;border-radius:4px;box-shadow:0 2px 4px rgba(0,0,0,.2);font-size:14px;font-weight:bold;transition:all .3s ease;background:${colors[type] || colors.info};color:#fff;`;
    container.appendChild(el);
    setTimeout(() => {
      el.style.opacity = '0';
      setTimeout(() => el.remove(), 300);
    }, 3000);
  }

  // ── Template UI ───────────────────────────────────────────────────────────

  initTemplateUI() {
    this._renderTemplateSelector();
    this._bindTemplateModalEvents();
  }

  _renderTemplateSelector() {
    const sel = document.getElementById('template-select');
    if (!sel) return;
    sel.innerHTML = '';
    const active = templateManager.getActive();
    templateManager.getAll().forEach(tpl => {
      const opt = document.createElement('option');
      opt.value = tpl.id;
      opt.textContent = tpl.name;
      if (tpl.id === active.id) opt.selected = true;
      sel.appendChild(opt);
    });
    sel.onchange = () => {
      templateManager.setActive(sel.value);
      this._checkTemplateFonts();
      // 模板切换时清除预览缓存并刷新预览
      previewRenderer.clearCache();
      this.updatePreview();
    };
  }

  _bindTemplateModalEvents() {
    const openBtn  = document.getElementById('manage-templates-btn');
    const modal    = document.getElementById('template-modal');
    const closeBtn = document.getElementById('template-modal-close');
    if (!openBtn || !modal || !closeBtn) return;

    openBtn.addEventListener('click', () => this._openTemplateModal());
    closeBtn.addEventListener('click', () => this._closeTemplateModal());
    modal.addEventListener('click', e => { if (e.target === modal) this._closeTemplateModal(); });
    document.addEventListener('keydown', e => {
      if (e.key === 'Escape' && modal.classList.contains('active')) this._closeTemplateModal();
    });

    document.getElementById('new-template-btn')?.addEventListener('click', () => {
      const blank = templateManager.clone('general');
      blank.name = '新建模板';
      templateManager.save(blank);
      this._renderTemplateList();
      this._openTemplateEditor(blank.id);
    });

    document.getElementById('tpl-cancel-btn')?.addEventListener('click', () => this._hideTemplateEditor());

    // Config export
    document.getElementById('export-config-btn')?.addEventListener('click', () => {
      templateManager.exportConfig();
    });

    // Config import
    document.getElementById('import-config-btn')?.addEventListener('click', () => {
      document.getElementById('config-file-input')?.click();
    });
    document.getElementById('config-file-input')?.addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (ev) => {
        try {
          const count = templateManager.importConfig(ev.target.result);
          this._renderTemplateList();
          this._renderTemplateSelector();
          this._showToast(`✓ 已导入 ${count} 个模板`);
        } catch (err) {
          this._showToast(`✗ 导入失败: ${err.message}`, true);
        }
        e.target.value = '';
      };
      reader.readAsText(file);
    });

    // Import format from DOCX
    document.getElementById('import-docx-btn')?.addEventListener('click', () => {
      document.getElementById('docx-file-input')?.click();
    });
    document.getElementById('docx-file-input')?.addEventListener('change', async (e) => {
      const file = e.target.files[0];
      if (!file) return;
      try {
        this._showToast('⏳ 正在提取格式…');
        const tpl = await parseDocxStyles(file);
        templateManager.save(tpl);
        templateManager.setActive(tpl.id);
        this._renderTemplateList();
        this._renderTemplateSelector();
        this._openTemplateEditor(tpl.id);
        this._showTemplateImportSummary(tpl);
        await this._checkTemplateFonts();
      } catch (err) {
        this._showToast(`✗ 提取失败: ${err.message}`, true);
        console.error('DOCX import error:', err);
      }
      e.target.value = '';
    });

    // Prompt modal bindings
    const promptModal = document.getElementById('prompt-modal');
    document.getElementById('prompt-modal-close')?.addEventListener('click', () => {
      promptModal?.classList.remove('active');
    });
    promptModal?.addEventListener('click', (e) => {
      if (e.target === promptModal) promptModal.classList.remove('active');
    });
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && promptModal?.classList.contains('active')) {
        promptModal.classList.remove('active');
      }
    });

    const doCopy = () => {
      const ta = document.getElementById('prompt-textarea');
      if (!ta) return;
      navigator.clipboard.writeText(ta.value).then(() => {
        const hint = document.getElementById('prompt-copy-hint');
        if (hint) { hint.textContent = '✓ 已复制'; setTimeout(() => { hint.textContent = ''; }, 2000); }
        const icon = document.getElementById('prompt-copy-icon');
        if (icon) { icon.textContent = 'check'; icon.style.color = '#4ade80'; setTimeout(() => { icon.textContent = 'content_copy'; icon.style.color = ''; }, 2000); }
      }).catch(() => { ta.select(); document.execCommand('copy'); });
    };
    document.getElementById('copy-prompt-btn')?.addEventListener('click', doCopy);
    document.getElementById('prompt-copy-icon-btn')?.addEventListener('click', doCopy);
  }

  _openTemplateModal() {
    const modal = document.getElementById('template-modal');
    if (!modal) return;
    this._renderTemplateList();
    this._hideTemplateEditor();
    this._showEmptyTemplateDetail();
    modal.classList.add('active');
  }

  _closeTemplateModal() {
    const modal = document.getElementById('template-modal');
    if (modal) modal.classList.remove('active');
    this._renderTemplateSelector();
  }

  _showEmptyTemplateDetail() {
    const editor = document.getElementById('template-editor');
    if (!editor) return;
    
    editor.style.display = 'flex';
    editor.innerHTML = `
      <div class="tpl-empty-detail">
        <span class="material-symbols-outlined">description</span>
        <div style="font-size: 16px; font-weight: 600; margin-bottom: 8px;">选择一个模板</div>
        <div style="font-size: 14px;">点击左侧列表中的模板查看详情</div>
      </div>`;
  }

  _showTemplateDetail(id) {
    const tpl = templateManager.get(id);
    const editor = document.getElementById('template-editor');
    if (!editor) return;

    const alignLabel = (a) => ({ justified: '两端对齐', left: '左对齐', center: '居中', right: '右对齐' })[a] || a;
    
    editor.style.display = 'flex';
    editor.innerHTML = `
      <div class="tpl-detail-panel">
        <div class="tpl-detail-header">
          <div class="tpl-detail-title">${tpl.name}</div>
          <div class="tpl-detail-desc">${tpl.description || '无描述'}</div>
        </div>

        <div class="tpl-detail-section">
          <div class="tpl-detail-section-title">页面设置</div>
          <div class="tpl-detail-row">
            <div class="tpl-detail-label">纸张大小</div>
            <div class="tpl-detail-value">${tpl.page.size} (${tpl.page.orientation === 'landscape' ? '横向' : '纵向'})</div>
          </div>
          <div class="tpl-detail-row">
            <div class="tpl-detail-label">页边距</div>
            <div class="tpl-detail-value">上${tpl.page.marginTop}mm · 下${tpl.page.marginBottom}mm · 左${tpl.page.marginLeft}mm · 右${tpl.page.marginRight}mm</div>
          </div>
        </div>

        <div class="tpl-detail-section">
          <div class="tpl-detail-section-title">正文格式</div>
          <div class="tpl-detail-row">
            <div class="tpl-detail-label">字体</div>
            <div class="tpl-detail-value">${tpl.body.font}</div>
          </div>
          <div class="tpl-detail-row">
            <div class="tpl-detail-label">字号</div>
            <div class="tpl-detail-value">${tpl.body.fontSize}pt</div>
          </div>
          <div class="tpl-detail-row">
            <div class="tpl-detail-label">行距</div>
            <div class="tpl-detail-value">${tpl.body.lineSpacing}pt（固定值）</div>
          </div>
          <div class="tpl-detail-row">
            <div class="tpl-detail-label">对齐方式</div>
            <div class="tpl-detail-value">${alignLabel(tpl.body.alignment)}</div>
          </div>
        </div>

        <div class="tpl-detail-section">
          <div class="tpl-detail-section-title">标题格式</div>
          <div class="tpl-detail-row">
            <div class="tpl-detail-label">文档标题</div>
            <div class="tpl-detail-value">${tpl.title.font} · ${tpl.title.fontSize}pt${tpl.title.bold ? ' · 加粗' : ''}</div>
          </div>
          <div class="tpl-detail-row">
            <div class="tpl-detail-label">一级标题</div>
            <div class="tpl-detail-value">${tpl.h1.font} · ${tpl.h1.fontSize}pt${tpl.h1.bold ? ' · 加粗' : ''}</div>
          </div>
          <div class="tpl-detail-row">
            <div class="tpl-detail-label">二级标题</div>
            <div class="tpl-detail-value">${tpl.h2.font} · ${tpl.h2.fontSize}pt${tpl.h2.bold ? ' · 加粗' : ''}</div>
          </div>
        </div>

        <div class="tpl-detail-actions">
          <button class="btn btn-primary" onclick="document.querySelector('[data-id=\\"${id}\\"].tpl-use-btn').click()">使用此模板</button>
          <button class="btn btn-outline" onclick="document.querySelector('[data-id=\\"${id}\\"].tpl-copy-btn').click()">复制模板</button>
          ${!tpl.readonly ? `<button class="btn btn-outline" onclick="document.querySelector('[data-id=\\"${id}\\"].tpl-edit-btn').click()">编辑模板</button>` : ''}
        </div>
      </div>`;
  }

  _renderTemplateList() {
    const sysList  = document.getElementById('template-list');
    const userList = document.getElementById('user-template-list');
    const userGrp  = document.getElementById('user-templates-group');
    if (!sysList) return;

    const active   = templateManager.getActive();
    const all      = templateManager.getAll();
    const sysTpls  = all.filter(t =>  t.readonly);
    const userTpls = all.filter(t => !t.readonly);

    sysList.innerHTML  = '';
    if (userList) userList.innerHTML = '';
    if (userGrp)  userGrp.style.display = userTpls.length ? '' : 'none';

    const makeItem = (tpl) => {
      const item = document.createElement('div');
      item.className = 'tpl-list-item' + (tpl.id === active.id ? ' selected' : '');
      item.dataset.id = tpl.id;
      item.innerHTML = `
        <div class="tpl-item-name">${tpl.name}${tpl.readonly ? ' <span class="tpl-badge">内置</span>' : ''}</div>
        <div class="tpl-item-desc">${tpl.description || ''}</div>
        <div class="tpl-item-actions">
          <button class="btn btn-xs tpl-use-btn"  data-id="${tpl.id}">使用</button>
          <button class="btn btn-xs tpl-copy-btn" data-id="${tpl.id}">复制</button>
          <button class="btn btn-xs tpl-view-btn" data-id="${tpl.id}">查看</button>
          ${!tpl.readonly ? `<button class="btn btn-xs tpl-edit-btn" data-id="${tpl.id}">编辑</button>` : ''}
          ${!tpl.readonly ? `<button class="btn btn-xs btn-danger tpl-del-btn" data-id="${tpl.id}">删除</button>` : ''}
        </div>`;
      return item;
    };

    sysTpls.forEach(t  => sysList.appendChild(makeItem(t)));
    if (userList) userTpls.forEach(t => userList.appendChild(makeItem(t)));

    const attachHandlers = (container) => {
      container.querySelectorAll('.tpl-use-btn').forEach(btn => btn.addEventListener('click', e => {
        templateManager.setActive(e.currentTarget.dataset.id);
        this._renderTemplateList();
        this._renderTemplateSelector();
        this._closeTemplateModal();
        this._checkTemplateFonts();
      }));
      container.querySelectorAll('.tpl-copy-btn').forEach(btn => btn.addEventListener('click', e => {
        const cloned = templateManager.clone(e.currentTarget.dataset.id);
        this._renderTemplateList();
        this._openTemplateEditor(cloned.id);
      }));
      container.querySelectorAll('.tpl-view-btn').forEach(btn => btn.addEventListener('click', e => {
        this._showTemplateDetail(e.currentTarget.dataset.id);
      }));
      container.querySelectorAll('.tpl-edit-btn').forEach(btn => btn.addEventListener('click', e => {
        this._openTemplateEditor(e.currentTarget.dataset.id);
      }));
      container.querySelectorAll('.tpl-del-btn').forEach(btn => btn.addEventListener('click', e => {
        if (confirm('确定要删除这个模板吗？')) {
          templateManager.delete(e.currentTarget.dataset.id);
          this._renderTemplateList();
          this._showEmptyTemplateDetail();
        }
      }));
    };
    attachHandlers(sysList);
    if (userList) attachHandlers(userList);
  }

  _openTemplateEditor(id) {
    const tpl    = templateManager.get(id);
    const editor = document.getElementById('template-editor');
    if (!editor) return;
    editor.style.display = 'flex';
    editor.dataset.editingId = id;
    this._fillEditorForm(tpl);
  }

  _hideTemplateEditor() {
    const editor = document.getElementById('template-editor');
    if (editor) editor.style.display = 'none';
  }

  _fillEditorForm(tpl) {
    const f   = id => document.getElementById(id);
    const set = (id, val) => { const el = f(id); if (el) el.value = String(val ?? ''); };

    set('tpl-name', tpl.name);
    set('tpl-desc', tpl.description || '');

    set('tpl-page-size',        tpl.page.size);
    set('tpl-page-orientation', tpl.page.orientation);
    // Sync orientation toggle buttons
    const orientVal = tpl.page.orientation || 'portrait';
    document.querySelectorAll('#orientation-toggle .orient-btn').forEach(b => {
      b.classList.toggle('active', b.dataset.value === orientVal);
    });
    // Update paper thumbnail
    window.updatePaperPreview?.();
    set('tpl-margin-top',    tpl.page.marginTop);
    set('tpl-margin-bottom', tpl.page.marginBottom);
    set('tpl-margin-left',   tpl.page.marginLeft);
    set('tpl-margin-right',  tpl.page.marginRight);

    set('tpl-body-font',    tpl.body.font);
    set('tpl-body-size',    tpl.body.fontSize);
    set('tpl-body-spacing', tpl.body.lineSpacing);
    set('tpl-body-indent',  tpl.body.firstLineIndent);
    set('tpl-body-align',   tpl.body.alignment);

    const FONTS  = ['仿宋_GB2312','楷体_GB2312','黑体','宋体','微软雅黑','方正小标宋_GBK','等线','Arial','Times New Roman'];
    const ALIGNS = [['justified','两端对齐'],['left','左对齐'],['center','居中'],['right','右对齐']];
    const LEVELS = [['title','标题(#)'],['h1','一级(##)'],['h2','二级(###)'],['h3','三级(####)'],['h4','四级(#####)'],['h5','五级(######)']];

    const rowsContainer = document.getElementById('tpl-headings-rows');
    if (rowsContainer) {
      rowsContainer.innerHTML = '';
      LEVELS.forEach(([lvl, label]) => {
        const h = tpl[lvl];
        const row = document.createElement('div');
        row.className = 'tpl-heading-row';
        const fontOpts  = FONTS.map(fn => `<option value="${fn}"${h.font === fn ? ' selected' : ''}>${fn}</option>`).join('');
        const alignOpts = ALIGNS.map(([v, l]) => `<option value="${v}"${h.alignment === v ? ' selected' : ''}>${l}</option>`).join('');
        row.innerHTML = `
          <span>${label}</span>
          <select id="tpl-${lvl}-font">${fontOpts}</select>
          <input  id="tpl-${lvl}-size"  type="number" value="${h.fontSize}" min="6" max="80" step="0.5">
          <label class="bold-pill-wrap"><input id="tpl-${lvl}-bold" type="checkbox"${h.bold ? ' checked' : ''}><span class="bold-pill"></span></label>
          <select id="tpl-${lvl}-align">${alignOpts}</select>`;
        rowsContainer.appendChild(row);
      });
    }

    // Show/hide readonly alert
    const alert = document.getElementById('tpl-readonly-alert');
    if (alert) alert.style.display = tpl.readonly ? 'flex' : 'none';

    const form = document.getElementById('template-editor-form');
    if (form) {
      form.querySelectorAll('input, select, textarea').forEach(el => {
        el.disabled = !!tpl.readonly;
      });
      // Keep orient-btn elements interactive (they update hidden input)
      form.querySelectorAll('.orient-btn').forEach(el => { el.disabled = !!tpl.readonly; });
    }
    const saveBtn = document.getElementById('tpl-save-btn');
    if (saveBtn) {
      saveBtn.style.display = tpl.readonly ? 'none' : '';
      // Replace node to clear any stale listeners
      const newBtn = saveBtn.cloneNode(true);
      saveBtn.parentNode.replaceChild(newBtn, saveBtn);
      newBtn.addEventListener('click', () => this._saveEditorForm());
    }
  }

  _saveEditorForm() {
    const f  = id => document.getElementById(id);
    const gv = id => { const el = f(id); return el ? el.value : ''; };
    const gc = id => { const el = f(id); return el ? el.checked : false; };
    const gn = id => parseFloat(gv(id)) || 0;

    const editor = document.getElementById('template-editor');
    const id = editor?.dataset.editingId;
    if (!id) return;

    const existing = templateManager.get(id);
    const updated = {
      ...existing,
      name:        gv('tpl-name').trim() || '未命名模板',
      description: gv('tpl-desc'),
      page: {
        size:         gv('tpl-page-size')        || 'A4',
        orientation:  gv('tpl-page-orientation') || 'portrait',
        marginTop:    gn('tpl-margin-top'),
        marginBottom: gn('tpl-margin-bottom'),
        marginLeft:   gn('tpl-margin-left'),
        marginRight:  gn('tpl-margin-right'),
      },
      body: {
        font:            gv('tpl-body-font')    || '宋体',
        fontSize:        gn('tpl-body-size')    || 12,
        lineSpacing:     gn('tpl-body-spacing') || 24,
        firstLineIndent: gn('tpl-body-indent'),
        alignment:       gv('tpl-body-align')   || 'justified',
      },
    };
    ['title', 'h1', 'h2', 'h3', 'h4', 'h5'].forEach(lvl => {
      updated[lvl] = {
        font:      gv(`tpl-${lvl}-font`)  || '宋体',
        fontSize:  gn(`tpl-${lvl}-size`)  || 12,
        bold:      gc(`tpl-${lvl}-bold`),
        alignment: gv(`tpl-${lvl}-align`) || 'left',
        color:     existing[lvl]?.color   || '000000',
      };
    });

    templateManager.save(updated);
    this._renderTemplateList();
    this._showToast(`✓ 模板"${updated.name}"已保存`);
    
    // 如果保存的是当前激活的模板，刷新预览
    if (templateManager.getActive().id === updated.id) {
      previewRenderer.clearCache();
      this.updatePreview();
    }
  }

  _generateLLMPrompt(template) {
    const { name, body, title, h1, h2, h3, h4, h5, page } = template;
    const alignLabel = (a) => ({ justified: '两端对齐', left: '左对齐', center: '居中', right: '右对齐' })[a] || a;
    const hLine = (mark, label, h) =>
      `- **${label}**（\`${mark}\`）：${h.font} · ${h.fontSize}pt${h.bold ? ' · 加粗' : ''} · ${alignLabel(h.alignment)}`;

    return `你是一个专业的文档撰写助手。请严格按照以下格式规范编写 Markdown 文档，以便使用「${name}」模板转换为标准 Word 文档。

## 页面设置
- 纸张：${page.size}（${page.orientation === 'landscape' ? '横向' : '纵向'}）
- 页边距：上 ${page.marginTop}mm · 下 ${page.marginBottom}mm · 左 ${page.marginLeft}mm · 右 ${page.marginRight}mm

## 正文格式
- 字体：${body.font} · ${body.fontSize}pt
- 行距：${body.lineSpacing}pt（固定值）
- 首行缩进：${body.firstLineIndent} 字符
- 段落对齐：${alignLabel(body.alignment)}

## 标题层级规范
${hLine('#',      '文档主标题（唯一）', title)}
${hLine('##',     '一级标题', h1)}
${hLine('###',    '二级标题', h2)}
${hLine('####',   '三级标题', h3)}
${hLine('#####',  '四级标题', h4)}
${hLine('######', '五级标题', h5)}

## 写作要求
1. 直接输出 Markdown 内容，不要添加任何说明或代码块包裹
2. 文档主标题（\`#\`）只出现一次，置于文档开头
3. 段落之间空一行，不要连续空多行
4. 不要使用 HTML 标签
5. 列表、表格、引用等按标准 Markdown 语法正常使用

请根据用户要求开始撰写：`;
  }

  _showPromptModal(templateName, prompt) {
    const modal = document.getElementById('prompt-modal');
    const ta    = document.getElementById('prompt-textarea');
    if (!modal || !ta) return;
    ta.value = prompt;
    const meta = document.getElementById('prompt-meta-line');
    if (meta) meta.textContent = `为模板「${templateName}」生成，共 ${prompt.length} 字`;
    modal.classList.add('active');
    
    // 自动复制到剪贴板
    navigator.clipboard.writeText(prompt).then(() => {
      const hint = document.getElementById('prompt-copy-hint');
      if (hint) {
        hint.textContent = '✓ 提示词已自动复制';
        hint.style.color = '#4ade80';
      }
      this._showToast('✓ 提示词已复制到剪贴板，可直接粘贴到 AI 工具使用', false, 3000);
    }).catch(() => {
      const hint = document.getElementById('prompt-copy-hint');
      if (hint) hint.textContent = '';
    });
  }

  _showTemplateImportSummary(tpl) {
    const alignLabel = (a) => ({ justified: '两端对齐', left: '左对齐', center: '居中', right: '右对齐' })[a] || a;
    
    const summary = `📄 模板提取成功

📏 页面设置
• 纸张:${tpl.page.size} ${tpl.page.orientation === 'portrait' ? '纵向' : '横向'}
• 边距:上${tpl.page.marginTop}mm 下${tpl.page.marginBottom}mm 左${tpl.page.marginLeft}mm 右${tpl.page.marginRight}mm

✏️ 正文格式
• 字体:${tpl.body.font} ${tpl.body.fontSize}pt
• 行距:${tpl.body.lineSpacing}pt
• 首行缩进:${tpl.body.firstLineIndent}字 · ${alignLabel(tpl.body.alignment)}

📑 标题格式
• 主标题:${tpl.title.font} ${tpl.title.fontSize}pt ${tpl.title.bold ? '加粗' : ''} ${alignLabel(tpl.title.alignment)}
• H1:${tpl.h1.font} ${tpl.h1.fontSize}pt ${tpl.h1.bold ? '加粗' : ''} ${alignLabel(tpl.h1.alignment)}
• H2:${tpl.h2.font} ${tpl.h2.fontSize}pt ${tpl.h2.bold ? '加粗' : ''} ${alignLabel(tpl.h2.alignment)}

⚠️ 注意:页眉页脚需手动配置(当前版本不支持自动提取)`;
    
    this._showToast(summary, false, 6000);
  }

  async _checkTemplateFonts() {
    const tpl = templateManager.getActive();
    if (!tpl) return;
    
    try {
      const { missing } = await checkTemplateFonts(tpl);
      
      if (missing.length > 0) {
        const os = detectOS();
        const fontList = missing.map(f => `"${f}"`).join('、');
        
        let message = `⚠️ 字体检测:当前系统未安装 ${fontList}\n`;
        
        if (os === 'Windows') {
          message += `导出的 Word 文档可能使用替代字体`;
        } else if (os === 'macOS') {
          message += `macOS 用户可使用等效字体(如"黑体-简"/"楷体-简")`;
        } else {
          message += `建议安装对应字体或在导出后手动调整`;
        }
        
        // 使用更长的显示时间以便用户阅读
        this._showToast(message, false, 4500);
      }
    } catch (err) {
      console.warn('字体检测失败:', err);
    }
  }

  _showToast(message, isError = false, duration = 2400) {
    const toast = document.createElement('div');
    toast.className = 'app-toast';
    toast.textContent = message;
    if (isError) toast.style.background = '#ba1a1a';
    document.body.appendChild(toast);
    setTimeout(() => toast.remove(), duration);
  }
}

document.addEventListener('DOMContentLoaded', () => new App());
