# md2docx 产品与工程优化规划

> **文档版本**：v1.0  
> **创建日期**：2026-09-11  
> **适用版本**：md2docx v1.1.2  
> **状态**：规划阶段（Phase 0）

## 目录

1. [现状摘要](#1-现状摘要)
2. [PDF 导出方案对比](#2-pdf-导出方案对比)
3. [建议的接口/模块边界](#3-建议的接口模块边界)
4. [功能优化优先级清单](#4-功能优化优先级清单)
5. [工程改造清单](#5-工程改造清单)
6. [分阶段落地](#6-分阶段落地)
7. [风险与不做清单](#7-风险与不做清单)

---

## 1. 现状摘要

### 1.1 架构概览

**md2docx** 是一个纯前端的 Markdown → Word DOCX 文档转换工具，具有以下特点：

- **核心定位**：Browser-only，无需服务器，所有转换在客户端完成
- **当前版本**：v1.1.2
- **在线演示**：https://henry-fox.github.io/md2docx/
- **许可协议**：MIT

### 1.2 技术栈

| 类型 | 技术 | 版本 | 用途 |
|------|------|------|------|
| Markdown 解析 | marked | 15.0.11 | 核心解析引擎 |
| Markdown 插件 | marked-extended-tables | 2.0.1 | 扩展表格支持 |
| Markdown 插件 | marked-gfm-heading-id | 4.1.1 | 标题ID生成 |
| Markdown 插件 | marked-mangle | 1.1.10 | 混淆处理 |
| DOCX 生成 | docx | 9.5.0 | Word文档生成 |
| 文件保存 | file-saver | 2.0.5 | 客户端文件下载 |
| 国际化 | i18next | 25.2.1 | 多语言支持 |
| 国际化 | react-i18next | 15.5.2 | i18n React绑定 |
| 构建工具 | webpack | 5.91.0 | 打包构建 |
| 开发服务器 | webpack-dev-server | 5.0.4 | 本地开发 |

### 1.3 核心流程

```
用户输入 Markdown
    ↓
marked.parse() → Token AST
    ↓
SimpleMd2Docx.convertToDocxDirect()
    ↓
应用 Template (templateManager)
    ↓
docx.js → Document 对象
    ↓
Packer.toBlob() → file-saver
    ↓
下载 .docx 文件
```

**关键模块**：

1. **`js/app.js`**：UI 控制器，处理文件上传、预览、按钮事件
2. **`js/simpleMd2Docx.js`**（核心）：
   - 调用 `marked.lexer()` 解析 Markdown
   - 遍历 token 生成 docx.js 段落对象
   - 处理标题、列表、表格、代码块、图片、链接等
3. **`js/templateManager.js`**：
   - 管理内置模板（党政机关公文、学术论文、毕业论文、司法文书、通用文档）
   - 用户自定义模板（LocalStorage）
   - 模板包含页面设置、字体、行距、缩进、对齐等样式
4. **`js/md2json.js`**：早期解析方案，目前 SimpleMd2Docx 直接使用 marked
5. **`js/json2docx.js`**：配合 md2json 的早期方案
6. **`js/docxParser.js`**：从已有 DOCX 文件导入格式，生成模板

### 1.4 产品亮点

✅ **模板系统完善**：内置 5 套专业模板，支持导入 DOCX 格式  
✅ **实时预览**：基于 marked.parse() 的 HTML 预览  
✅ **国际化**：支持中文、英文、法语、西班牙语、俄语、阿拉伯语  
✅ **LLM 集成**：可生成模板配置提示词，让 AI 生成符合格式的文档  
✅ **离线可用**：纯静态页面，下载后本地即可使用  
✅ **无隐私风险**：所有数据处理在浏览器本地，不上传服务器

### 1.5 已知约束

⚠️ **浏览器限制**：无法直接读取本地文件系统图片（需 base64 或网络 URL）  
⚠️ **字体依赖**：中文字体（仿宋_GB2312、楷体_GB2312、方正小标宋_GBK）需系统支持  
⚠️ **单文件处理**：不支持批量转换  
⚠️ **预览与导出不一致**：预览使用浏览器渲染，导出使用 docx.js，样式可能有差异

### 1.6 工程现状

**代码结构**：
- 前端入口：`index.html`
- 核心逻辑：`js/` 目录
- 样式：`css/`、`styles/`
- 国际化：`src/locales/`
- 构建配置：`webpack.*.js`

**潜在问题**：
- ❌ 根目录有散落的测试文件（test-*.js, test-*.md, test.json）
- ❌ 根目录有重复文件（simpleMd2Docx.js 同时存在于 `./` 和 `js/`）
- ❌ 有大型参考文件（党政机关公文格式.pdf，8MB）
- ❌ 有示例文件（My Document.docx）
- ❌ `package.json` 中 test script 为空占位
- ❌ 缺少自动化测试
- ❌ 缺少 CI/CD 流程

---

## 2. PDF 导出方案对比

### 2.1 方案 A：HTML 预览 → PDF（前端转换）

**实现方式**：
```javascript
// 使用 html2canvas + jsPDF 或 html2pdf.js
import html2pdf from 'html2pdf.js';

const element = document.getElementById('preview-container');
const opt = {
  margin: [25.4, 31.8, 25.4, 31.8],  // mm
  filename: 'document.pdf',
  image: { type: 'jpeg', quality: 0.98 },
  html2canvas: { scale: 2, useCORS: true },
  jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
};
html2pdf().set(opt).from(element).save();
```

**优点**：
- ✅ 实现快速（集成库即可）
- ✅ 与当前预览完全一致（所见即所得）
- ✅ 保持纯前端架构
- ✅ 可复用现有 marked 预览流程

**缺点**：
- ❌ **保真度风险**：HTML → Canvas → PDF，字体渲染、分页、中文支持可能有问题
- ❌ **分页控制弱**：难以精确控制页眉页脚、页码
- ❌ **代码块/表格溢出**：长表格或代码可能被截断
- ❌ **文件体积较大**：图片质量与体积难以平衡
- ❌ **不支持模板**：难以复用 templateManager 的专业格式

**适用场景**：快速 MVP，对格式要求不高的个人使用

**库选择**：
- **html2pdf.js**（推荐）：基于 html2canvas + jsPDF
- **jsPDF + html2canvas**：手动组合，灵活度更高
- **Puppeteer**（需服务器）：不适用

**中文支持风险**：
- html2canvas 对中文字体渲染可能有问题
- 需确保 jsPDF 包含中文字体（或嵌入自定义字体）

---

### 2.2 方案 B：模板驱动 PDF 生成（pdf-lib）

**实现方式**：
```javascript
// 基于 pdf-lib，类似 SimpleMd2Docx 的架构
import { PDFDocument, rgb, StandardFonts } from 'pdf-lib';

class SimpleMd2Pdf {
  async convertToPdfDirect(markdown) {
    const pdfDoc = await PDFDocument.create();
    const tokens = marked.lexer(markdown);
    
    for (const token of tokens) {
      switch (token.type) {
        case 'heading': this.createHeading(pdfDoc, token); break;
        case 'paragraph': this.createParagraph(pdfDoc, token); break;
        // ...
      }
    }
    
    const pdfBytes = await pdfDoc.save();
    saveAs(new Blob([pdfBytes], { type: 'application/pdf' }), 'document.pdf');
  }
}
```

**优点**：
- ✅ **与 DOCX 逻辑并行**：可复用 marked token 解析和 templateManager
- ✅ **格式可控**：精确控制页边距、页眉页脚、页码、分页
- ✅ **产品一致性**：PDF 与 DOCX 使用相同模板配置
- ✅ **中文字体可控**：可嵌入自定义中文字体（需注意许可证）
- ✅ **纯前端**：保持无服务器架构

**缺点**：
- ❌ **开发成本高**：需重新实现 Markdown → PDF 转换逻辑（类似 SimpleMd2Docx）
- ❌ **中文字体体积大**：嵌入完整中文字体可能增加 10MB+ 体积（可按需子集化）
- ❌ **表格布局复杂**：pdf-lib 对表格支持较弱，需手动计算行列
- ❌ **数学公式困难**：不像 DOCX 有现成的 MathML 支持

**适用场景**：对格式要求高，需要专业模板系统，长期维护

**库选择**：
- **pdf-lib**（推荐）：纯 JS，功能全面，MIT 许可
- **pdfmake**：基于声明式配置，但对 Markdown 支持不佳
- **pdfkit**（Node.js）：不适用于浏览器

**中文字体方案**：
1. **嵌入子集字体**：使用 fontmin 提取常用汉字（约 3-5MB）
2. **外部字体 URL**：引用 CDN 字体（需网络，不离线）
3. **系统字体映射**：使用 PDF 标准中文字体（保真度低）

---

### 2.3 方案 C：服务端转换（不推荐）

**实现方式**：
```javascript
// 上传 Markdown 到服务器，调用 Pandoc/LibreOffice
const formData = new FormData();
formData.append('markdown', markdown);
const response = await fetch('/api/convert/pdf', {
  method: 'POST',
  body: formData
});
const blob = await response.blob();
saveAs(blob, 'document.pdf');
```

**优点**：
- ✅ 使用成熟工具（Pandoc、LibreOffice）
- ✅ 保真度高
- ✅ 支持复杂排版

**缺点**：
- ❌ **违背产品定位**：需要服务器，破坏 "Browser-only" 核心价值
- ❌ **隐私风险**：用户数据上传到服务器
- ❌ **运维成本**：需要维护服务器、处理并发、防止滥用
- ❌ **离线不可用**：依赖网络

**建议**：**不采用此方案**，除非产品定位变更为 SaaS 服务

---

### 2.4 推荐方案

**阶段性采用策略**：

| 阶段 | 方案 | 理由 |
|------|------|------|
| **Phase 1 MVP** | 方案 A（html2pdf.js） | 快速验证需求，2-3天即可上线 |
| **Phase 2 产品化** | 方案 B（pdf-lib） | 复用模板系统，提供专业格式 |
| **长期** | 双引擎并存 | 快速导出用 A，专业导出用 B，用户可选 |

**关键技术风险**：

1. **中文字体**：
   - 方案 A：测试 html2canvas 中文渲染，必要时用 Web Font 覆盖
   - 方案 B：子集化字体（仅包含 GB2312 字符集），控制在 3-5MB

2. **分页处理**：
   - 方案 A：使用 `page-break-before/after` CSS，依赖 jsPDF 自动分页
   - 方案 B：手动计算每页内容高度，精确控制分页

3. **代码块渲染**：
   - 方案 A：预览样式直接转换
   - 方案 B：使用等宽字体，手动处理自动换行

---

## 3. 建议的接口/模块边界

### 3.1 模块职责划分

```
┌─────────────────────────────────────────────────────┐
│                     App.js (UI层)                    │
│  - 文件上传/拖拽                                        │
│  - 工具栏操作                                          │
│  - 模板选择                                           │
│  - 导出按钮事件                                        │
└──────────────┬──────────────────────┬────────────────┘
               │                      │
               ↓                      ↓
    ┌──────────────────┐   ┌──────────────────┐
    │  PreviewManager  │   │  ExportManager   │
    │  - updatePreview │   │  - exportDocx    │
    │  - 使用 marked   │   │  - exportPdf     │
    └──────────────────┘   └────────┬─────────┘
                                    │
                 ┌──────────────────┴──────────────────┐
                 ↓                                     ↓
      ┌──────────────────┐                 ┌──────────────────┐
      │ SimpleMd2Docx    │                 │  SimpleMd2Pdf    │
      │ - convertDirect  │                 │  - convertDirect │
      │ - createHeading  │                 │  - createHeading │
      │ - createTable    │                 │  - createTable   │
      └────────┬─────────┘                 └────────┬─────────┘
               │                                     │
               └────────────┬────────────────────────┘
                            ↓
                   ┌──────────────────┐
                   │ TemplateManager  │
                   │ - getActive()    │
                   │ - toDocxStyles() │
                   │ - toPdfStyles()  │← 新增
                   └──────────────────┘
```

### 3.2 核心接口设计

#### 3.2.1 统一导出接口

```javascript
// js/exportManager.js (新建)
export class ExportManager {
  constructor() {
    this.docxConverter = new SimpleMd2Docx();
    this.pdfConverter = new SimpleMd2Pdf();  // Phase 2 添加
  }

  /**
   * 导出为 DOCX
   * @param {string} markdown - Markdown 文本
   * @param {object} template - 模板配置
   * @returns {Promise<void>}
   */
  async exportDocx(markdown, template = null) {
    const tpl = template || templateManager.getActive();
    this.docxConverter.setTemplate(tpl);
    await this.docxConverter.convertToDocxDirect(markdown);
  }

  /**
   * 导出为 PDF
   * @param {string} markdown - Markdown 文本
   * @param {object} template - 模板配置
   * @param {string} method - 'html' | 'native' (Phase 2)
   * @returns {Promise<void>}
   */
  async exportPdf(markdown, template = null, method = 'html') {
    const tpl = template || templateManager.getActive();
    
    if (method === 'html') {
      // Phase 1: 基于预览 HTML 转换
      await this._exportPdfFromHtml(markdown, tpl);
    } else {
      // Phase 2: 原生 PDF 生成
      this.pdfConverter.setTemplate(tpl);
      await this.pdfConverter.convertToPdfDirect(markdown);
    }
  }

  async _exportPdfFromHtml(markdown, template) {
    // 实现 html2pdf 转换
  }
}
```

#### 3.2.2 模板管理扩展

```javascript
// js/templateManager.js (扩展现有代码)
class TemplateManager {
  // ... 现有方法 ...

  /**
   * 将模板转换为 PDF 样式配置（Phase 2 新增）
   * @param {object} template - 模板对象
   * @returns {object} PDF 样式配置
   */
  toPdfStyles(template) {
    const size = PAGE_SIZES[template.page.size];
    return {
      pageWidth: size.width,
      pageHeight: size.height,
      pageOrientation: template.page.orientation === 'landscape' ? 'landscape' : 'portrait',
      pageMargin: {
        top: mmToPoints(template.page.marginTop),
        bottom: mmToPoints(template.page.marginBottom),
        left: mmToPoints(template.page.marginLeft),
        right: mmToPoints(template.page.marginRight),
      },
      body: {
        font: this._mapPdfFont(template.body.font),
        fontSize: template.body.fontSize,
        lineHeight: template.body.lineSpacing / template.body.fontSize,
        alignment: template.body.alignment,
        firstLineIndent: template.body.firstLineIndent,
      },
      // ... 类似 toDocxStyles 的映射
    };
  }

  /**
   * 将中文字体映射到 PDF 字体
   */
  _mapPdfFont(fontName) {
    const fontMap = {
      '仿宋_GB2312': 'FangSong',
      '宋体': 'SimSun',
      '黑体': 'SimHei',
      '楷体_GB2312': 'KaiTi',
      '微软雅黑': 'Microsoft YaHei',
      // ... 其他映射
    };
    return fontMap[fontName] || fontName;
  }
}

function mmToPoints(mm) {
  return mm * 2.83465;  // 1mm = 2.83465pt
}
```

#### 3.2.3 共享 Token 解析层

```javascript
// js/markdownTokenizer.js (新建，抽象共享逻辑)
export class MarkdownTokenizer {
  constructor() {
    this.tokens = [];
  }

  /**
   * 解析 Markdown 为 Token 树
   * @param {string} markdown
   * @returns {Array<Token>}
   */
  parse(markdown) {
    return marked.lexer(markdown);
  }

  /**
   * 遍历 Token 树，应用处理器
   * @param {Array<Token>} tokens
   * @param {object} handlers - { heading, paragraph, table, ... }
   * @returns {Array<any>}
   */
  async traverse(tokens, handlers) {
    const results = [];
    for (const token of tokens) {
      const handler = handlers[token.type];
      if (handler) {
        const result = await handler(token);
        results.push(result);
      } else {
        console.warn(`No handler for token type: ${token.type}`);
      }
    }
    return results;
  }
}
```

**SimpleMd2Docx 和 SimpleMd2Pdf 都使用这个统一的 tokenizer**：

```javascript
// 在 SimpleMd2Docx 中
const tokenizer = new MarkdownTokenizer();
const tokens = tokenizer.parse(markdown);
const paragraphs = await tokenizer.traverse(tokens, {
  heading: (token) => this.createHeadingFromMarked(token),
  paragraph: (token) => this.createParagraphFromMarked(token),
  // ...
});
```

### 3.3 UI 层改造

```javascript
// js/app.js (改造导出按钮逻辑)
import { ExportManager } from './exportManager.js';

class App {
  constructor() {
    this.exportManager = new ExportManager();
    // ...
  }

  async directConvertToDocx() {
    const markdown = this.markdownInput.value;
    if (!markdown.trim()) {
      this.showMessage(t('emptyInput'), 'warning');
      return;
    }
    try {
      this.showMessage(t('convertingSimple'), 'info');
      await this.exportManager.exportDocx(markdown);
      this.showMessage(t('convertSimpleSuccess'), 'success');
    } catch (error) {
      console.error('转换失败:', error);
      this.showMessage(tWithVars('convertSimpleFail', { msg: error.message }), 'error');
    }
  }

  // Phase 1 新增
  async directConvertToPdf() {
    const markdown = this.markdownInput.value;
    if (!markdown.trim()) {
      this.showMessage(t('emptyInput'), 'warning');
      return;
    }
    try {
      this.showMessage('正在生成 PDF...', 'info');
      await this.exportManager.exportPdf(markdown, null, 'html');
      this.showMessage('PDF 生成成功', 'success');
    } catch (error) {
      console.error('PDF 生成失败:', error);
      this.showMessage(`PDF 生成失败: ${error.message}`, 'error');
    }
  }
}
```

**UI 按钮改造**：

```html
<!-- index.html (改造导出按钮区域) -->
<div class="export-buttons">
  <button id="direct-convert-btn" class="btn btn-primary">
    <span class="material-icons">description</span>
    <span data-i18n="exportDocx">导出 Word</span>
  </button>
  
  <!-- Phase 1 新增 -->
  <button id="export-pdf-btn" class="btn btn-secondary">
    <span class="material-icons">picture_as_pdf</span>
    <span data-i18n="exportPdf">导出 PDF</span>
  </button>
  
  <!-- Phase 2 新增：高级选项 -->
  <button id="export-options-btn" class="btn btn-outline">
    <span class="material-icons">settings</span>
    <span data-i18n="exportOptions">导出选项</span>
  </button>
</div>
```

### 3.4 模块边界原则

**严格遵守的边界**：

1. **UI 层不直接操作 docx/pdf 对象**：
   - ❌ 错误：`app.js` 中直接 `new Document()`
   - ✅ 正确：`app.js` 调用 `exportManager.exportDocx()`

2. **转换器不直接访问 DOM**：
   - ❌ 错误：`simpleMd2Docx.js` 中 `document.getElementById()`
   - ✅ 正确：转换器接收纯数据参数

3. **模板管理器不依赖转换器**：
   - ❌ 错误：`templateManager.js` 中 `import { SimpleMd2Docx } from ...`
   - ✅ 正确：模板提供通用配置对象，转换器自行解释

4. **预览与导出解耦**：
   - ❌ 错误：导出时读取 `preview-container` 的 HTML
   - ✅ 正确：预览和导出都从原始 Markdown 重新解析

---

## 4. 功能优化优先级清单

### P0 - 核心缺陷修复（必须完成）

| ID | 功能 | 现状 | 目标 | 工作量 |
|----|------|------|------|--------|
| P0-1 | **预览与导出一致性** | 预览用 marked.parse()，导出用 docx.js，样式可能不一致 | 统一样式规则，添加"预览即导出"模式 | 3天 |
| P0-2 | **错误处理** | 转换失败时信息不明确，部分错误未捕获 | 完善错误提示，添加详细错误日志 | 2天 |
| P0-3 | **中文字体缺失提示** | 系统缺少中文字体时静默失败 | 检测字体可用性，提示用户安装 | 1天 |

### P1 - 重要功能增强（强烈建议）

| ID | 功能 | 现状 | 目标 | 工作量 |
|----|------|------|------|--------|
| P1-1 | **PDF 导出 MVP** | 无 | 基于 html2pdf.js 实现快速 PDF 导出 | 3天 |
| P1-2 | **模板页眉页脚** | 模板只包含页面设置和字体，无页眉页脚 | 支持页眉页脚、页码配置 | 5天 |
| P1-3 | **Markdown 扩展** | 不支持数学公式、脚注、Mermaid 图表 | 支持 KaTeX 数学公式、脚注 | 5天 |
| P1-4 | **导出文件名自动化** | 当前根据 `#` 标题生成，但可能重复 | 支持自定义文件名模板，如 `{title}_{date}` | 1天 |
| P1-5 | **批量转换** | 只能单文件转换 | 支持多文件选择和批量转换 | 3天 |

### P2 - 体验优化（长期改进）

| ID | 功能 | 现状 | 目标 | 工作量 |
|----|------|------|------|--------|
| P2-1 | **实时保存草稿** | 刷新页面后内容丢失 | 自动保存到 LocalStorage | 1天 |
| P2-2 | **代码块语法高亮** | DOCX 中代码无高亮 | 使用颜色标记关键字（可选） | 3天 |
| P2-3 | **表格样式模板** | 表格样式固定 | 提供多种表格主题（简约、专业、学术） | 2天 |
| P2-4 | **目录（TOC）生成** | 无 | 自动生成文档目录 | 5天 |
| P2-5 | **封面页** | 无 | 支持模板定义封面页 | 3天 |
| P2-6 | **图片处理增强** | 网络图片需 CORS，可能加载失败 | 添加图片代理、缓存 | 3天 |
| P2-7 | **导出历史记录** | 无 | 记录最近导出的文件，支持重新下载 | 2天 |
| P2-8 | **快捷键支持** | 无 | 添加 Ctrl+S 保存、Ctrl+E 导出等 | 1天 |
| P2-9 | **暗色模式** | 仅亮色 | 支持暗色主题 | 2天 |

### 功能覆盖率分析

| Markdown 特性 | 当前支持 | 优先级 | 备注 |
|---------------|----------|--------|------|
| 标题 (H1-H6) | ✅ 完整支持 | - | 已实现 |
| 段落 | ✅ 完整支持 | - | 已实现 |
| 粗体/斜体/删除线 | ✅ 完整支持 | - | 已实现 |
| 链接 | ✅ 完整支持 | - | 已实现 |
| 图片 | ⚠️ 仅网络图片 | P1 | 本地图片需 base64 |
| 代码块 | ✅ 完整支持 | - | 无语法高亮（P2 优化） |
| 行内代码 | ✅ 完整支持 | - | 已实现 |
| 有序列表 | ✅ 完整支持 | - | 已实现 |
| 无序列表 | ✅ 完整支持 | - | 已实现 |
| 任务列表 | ✅ 完整支持 | - | 已实现 |
| 表格 | ✅ 完整支持 | - | 支持对齐方式 |
| 引用块 | ✅ 完整支持 | - | 已实现 |
| 水平线 | ✅ 完整支持 | - | 已实现 |
| HTML | ⚠️ 部分支持 | P2 | 仅 `<div>`+`<img>` |
| 数学公式 | ❌ 不支持 | P1 | 建议用 KaTeX |
| 脚注 | ❌ 不支持 | P1 | Markdown 扩展 |
| Mermaid 图表 | ❌ 不支持 | P2 | 需转 SVG 嵌入 |
| 目录 | ❌ 不支持 | P2 | 自动生成 TOC |

---

## 5. 工程改造清单

### 5.1 测试体系建设

**现状**：
- ❌ `package.json` 中 `test` script 为空
- ❌ 根目录有 `test-*.js` 手工测试脚本，但不是自动化测试
- ❌ 无 CI 流程

**目标**：
- ✅ 建立单元测试框架（Vitest 或 Jest）
- ✅ 核心模块测试覆盖率 > 80%
- ✅ 集成测试（E2E）覆盖主要转换场景
- ✅ CI 自动运行测试

**工作量**：5天

**推荐方案**：

```json
// package.json
{
  "scripts": {
    "test": "vitest",
    "test:ui": "vitest --ui",
    "test:coverage": "vitest run --coverage"
  },
  "devDependencies": {
    "vitest": "^1.0.0",
    "@vitest/ui": "^1.0.0",
    "jsdom": "^23.0.0"
  }
}
```

**测试结构**：

```
tests/
├── unit/
│   ├── templateManager.test.js       # 模板管理测试
│   ├── simpleMd2Docx.test.js         # DOCX 转换测试
│   ├── simpleMd2Pdf.test.js          # PDF 转换测试
│   └── exportManager.test.js         # 导出管理测试
├── integration/
│   ├── markdown-to-docx.test.js      # 端到端 DOCX 测试
│   └── markdown-to-pdf.test.js       # 端到端 PDF 测试
└── fixtures/
    ├── sample.md                      # 测试用 Markdown
    ├── expected-output.docx           # 预期输出（用于对比）
    └── templates/                     # 测试模板
```

**核心测试用例**：

```javascript
// tests/unit/simpleMd2Docx.test.js
import { describe, it, expect } from 'vitest';
import SimpleMd2Docx from '../js/simpleMd2Docx.js';

describe('SimpleMd2Docx', () => {
  it('should convert heading correctly', async () => {
    const converter = new SimpleMd2Docx();
    const token = { type: 'heading', depth: 1, text: '测试标题' };
    const result = await converter.createHeadingFromMarked(token);
    expect(result[0].heading).toBe('TITLE');
  });

  it('should handle Chinese fonts', () => {
    const converter = new SimpleMd2Docx();
    converter.setTemplate({ body: { font: '仿宋_GB2312' } });
    const styles = converter.getStyles();
    expect(styles.body.font).toBe('仿宋_GB2312');
  });

  // 更多测试...
});
```

### 5.2 代码清理

**现状问题**：

| 问题 | 文件 | 影响 | 处理方案 |
|------|------|------|----------|
| 根目录重复文件 | `simpleMd2Docx.js` | 混淆，可能引入错误版本 | 删除根目录版本，统一用 `js/simpleMd2Docx.js` |
| 测试文件散落 | `test-*.js`, `test-*.md`, `test.json` | 目录混乱 | 移动到 `tests/fixtures/` |
| 大型二进制文件 | `党政机关公文格式.pdf` (8MB) | 仓库体积过大 | 移到 `docs/references/` 或用 Git LFS |
| 示例文件 | `My Document.docx` | 无实际作用 | 移到 `examples/` |
| 空测试脚本 | `package.json`: `test` | 误导开发者 | 替换为实际测试命令 |

**清理步骤**：

```bash
# 1. 删除根目录重复文件
rm simpleMd2Docx.js

# 2. 整理测试文件
mkdir -p tests/fixtures
mv test-*.js test-*.md test.json tests/fixtures/

# 3. 整理文档和示例
mkdir -p docs/references examples
mv 党政机关公文格式.pdf docs/references/
mv "My Document.docx" examples/

# 4. 更新 .gitignore
echo "tests/fixtures/*.local.*" >> .gitignore
```

**工作量**：0.5天

### 5.3 构建优化

**现状**：
- ✅ Webpack 5 已配置
- ⚠️ 无生产构建优化（Tree Shaking、代码分割）
- ⚠️ 无 Source Map

**目标**：
- 减少构建产物体积（目前 `dist/main.js` 约 2MB）
- 支持按需加载（懒加载 PDF 转换库）
- 添加 Source Map 用于调试

**优化方案**：

```javascript
// webpack.config.prod.js (优化版)
const { merge } = require('webpack-merge');
const common = require('./webpack.common.js');
const TerserPlugin = require('terser-webpack-plugin');
const { BundleAnalyzerPlugin } = require('webpack-bundle-analyzer');

module.exports = merge(common, {
  mode: 'production',
  devtool: 'source-map',
  optimization: {
    minimize: true,
    minimizer: [
      new TerserPlugin({
        terserOptions: {
          compress: { drop_console: true },
        },
      }),
    ],
    splitChunks: {
      chunks: 'all',
      cacheGroups: {
        vendor: {
          test: /[\\/]node_modules[\\/]/,
          name: 'vendors',
          priority: 10,
        },
        docx: {
          test: /[\\/]node_modules[\\/]docx/,
          name: 'docx-lib',
          priority: 20,
        },
      },
    },
  },
  plugins: [
    new BundleAnalyzerPlugin({
      analyzerMode: 'static',
      openAnalyzer: false,
    }),
  ],
});
```

**预期效果**：
- 主包 < 500KB（gzip 后）
- 第三方库独立缓存
- 首次加载 < 2s（4G 网络）

**工作量**：2天

### 5.4 CI/CD 流程

**目标**：
- 自动运行测试
- 自动构建静态页面
- 自动发布到 GitHub Pages

**GitHub Actions 配置**：

```yaml
# .github/workflows/ci.yml
name: CI

on:
  push:
    branches: [main, develop]
  pull_request:
    branches: [main]

jobs:
  test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v3
      - uses: actions/setup-node@v3
        with:
          node-version: 20
          cache: 'npm'
      - run: npm ci
      - run: npm test
      - run: npm run build

  deploy:
    needs: test
    if: github.ref == 'refs/heads/main'
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v3
      - uses: actions/setup-node@v3
        with:
          node-version: 20
      - run: npm ci
      - run: npm run build
      - uses: peaceiris/actions-gh-pages@v3
        with:
          github_token: ${{ secrets.GITHUB_TOKEN }}
          publish_dir: ./dist
```

**工作量**：1天

### 5.5 版本发布流程

**现状**：
- `package.json` 中有 `prepare-release` 和 `release` 脚本
- 手动创建 zip 包

**改进**：
- 使用 GitHub Releases 自动打包
- 添加 CHANGELOG.md
- 语义化版本（Semantic Versioning）

**Release 工作流**：

```yaml
# .github/workflows/release.yml
name: Release

on:
  push:
    tags:
      - 'v*'

jobs:
  release:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v3
      - uses: actions/setup-node@v3
        with:
          node-version: 20
      - run: npm ci
      - run: npm run build
      - name: Create Release Archive
        run: |
          cd dist
          zip -r ../md2docx-${{ github.ref_name }}.zip .
      - uses: softprops/action-gh-release@v1
        with:
          files: md2docx-${{ github.ref_name }}.zip
          generate_release_notes: true
```

**工作量**：1天

### 5.6 文档完善

**需要补充的文档**：

| 文档 | 内容 | 优先级 |
|------|------|--------|
| `CONTRIBUTING.md` | 贡献指南，代码规范，提交规范 | P1 |
| `ARCHITECTURE.md` | 架构设计，模块关系 | P1 |
| `API.md` | 核心 API 文档 | P2 |
| `CHANGELOG.md` | 版本更新日志 | P1 |
| `docs/TEMPLATE_GUIDE.md` | 模板开发指南 | P2 |
| `docs/TROUBLESHOOTING.md` | 常见问题排查 | P2 |

**工作量**：3天

---

## 6. 分阶段落地

### Phase 0：文档与规划（当前阶段）

**目标**：完成规划文档，团队对齐

**交付物**：
- ✅ `docs/OPTIMIZATION_PLAN.md`（本文档）
- ✅ README 添加路线图链接

**工作量**：1天

**验收标准**：
- 维护者 Review 通过
- 无重大技术分歧

---

### Phase 1：PDF MVP + 工程基础（2-3周）

**目标**：快速上线 PDF 导出，建立测试体系

**子任务**：

| 任务 | 描述 | 工作量 | 优先级 |
|------|------|--------|--------|
| 1.1 代码清理 | 移除重复文件、整理目录结构 | 0.5天 | P0 |
| 1.2 错误处理增强 | 完善异常捕获和用户提示 | 2天 | P0 |
| 1.3 测试框架搭建 | 安装 Vitest，编写核心测试用例 | 3天 | P0 |
| 1.4 PDF 导出（html2pdf） | 基于预览实现快速 PDF 导出 | 3天 | P1 |
| 1.5 UI 改造 | 添加 PDF 导出按钮和选项 | 1天 | P1 |
| 1.6 中文字体检测 | 提示用户安装缺失字体 | 1天 | P0 |
| 1.7 CI/CD 配置 | GitHub Actions 自动测试和部署 | 1天 | P1 |

**验收标准**：
- ✅ 用户可点击"导出 PDF"按钮，下载 PDF 文件
- ✅ 测试覆盖率 > 60%
- ✅ CI 通过，自动部署到 gh-pages
- ✅ 无已知 P0 缺陷

**里程碑版本**：v1.2.0

---

### Phase 2：PDF 产品化 + 模板增强（4-6周）

**目标**：基于 pdf-lib 实现专业 PDF 导出，模板支持页眉页脚

**子任务**：

| 任务 | 描述 | 工作量 | 优先级 |
|------|------|--------|--------|
| 2.1 SimpleMd2Pdf 实现 | 类似 SimpleMd2Docx，基于 pdf-lib | 7天 | P1 |
| 2.2 中文字体子集化 | 提取常用汉字，减小体积 | 2天 | P1 |
| 2.3 模板扩展 | 支持页眉、页脚、页码配置 | 5天 | P1 |
| 2.4 ExportManager 实现 | 统一导出接口 | 2天 | P1 |
| 2.5 双引擎切换 | UI 支持选择"快速导出"或"专业导出" | 1天 | P2 |
| 2.6 数学公式支持 | 集成 KaTeX（可选功能） | 3天 | P1 |
| 2.7 脚注支持 | Markdown 扩展 + DOCX/PDF 实现 | 3天 | P1 |
| 2.8 测试覆盖率提升 | 补充集成测试 | 2天 | P1 |

**验收标准**：
- ✅ PDF 输出样式与 DOCX 一致
- ✅ 支持自定义页眉页脚和页码
- ✅ 测试覆盖率 > 80%
- ✅ PDF 文件体积合理（< 5MB 含字体）

**里程碑版本**：v1.3.0

---

### Phase 3：产品打磨 + 体验优化（持续迭代）

**目标**：提升易用性，扩展高级功能

**子任务**：

| 任务 | 描述 | 工作量 | 优先级 |
|------|------|--------|--------|
| 3.1 批量转换 | 多文件上传和批量导出 | 3天 | P1 |
| 3.2 实时保存草稿 | LocalStorage 自动保存 | 1天 | P2 |
| 3.3 目录生成 | 自动生成 TOC | 5天 | P2 |
| 3.4 封面页 | 模板支持封面配置 | 3天 | P2 |
| 3.5 导出历史 | 记录最近导出的文件 | 2天 | P2 |
| 3.6 暗色模式 | UI 支持暗色主题 | 2天 | P2 |
| 3.7 快捷键 | 添加常用快捷键 | 1天 | P2 |
| 3.8 性能优化 | 大文件转换优化（分块处理） | 3天 | P2 |
| 3.9 预览导出一致性 | 解决预览与导出样式差异 | 3天 | P0（遗留） |

**验收标准**：
- ✅ 支持 100+ 页文档快速转换（< 5s）
- ✅ 用户体验流畅，无明显卡顿
- ✅ 完善的错误提示和帮助文档

**里程碑版本**：v1.4.0+

---

### 阶段总结

| 阶段 | 核心价值 | 时间 | 版本 |
|------|----------|------|------|
| Phase 0 | 规划对齐 | 1天 | v1.1.2 |
| Phase 1 | 快速 MVP + 工程基础 | 2-3周 | v1.2.0 |
| Phase 2 | 专业 PDF + 模板增强 | 4-6周 | v1.3.0 |
| Phase 3 | 体验打磨（持续） | 长期 | v1.4.0+ |

---

## 7. 风险与不做清单

### 7.1 技术风险

| 风险 | 可能性 | 影响 | 缓解措施 |
|------|--------|------|----------|
| **中文字体渲染问题** | 高 | 高 | Phase 1 充分测试，准备字体降级方案 |
| **PDF 分页不准确** | 中 | 中 | Phase 2 手动计算页高，添加分页标记 |
| **大文件性能问题** | 中 | 中 | 分块处理，添加进度条，限制文件大小 |
| **浏览器兼容性** | 低 | 中 | 明确支持的浏览器版本（Chrome/Edge/Firefox 最新版） |
| **依赖库更新破坏** | 低 | 高 | 锁定主要依赖版本，定期测试更新 |

### 7.2 产品风险

| 风险 | 可能性 | 影响 | 缓解措施 |
|------|--------|------|----------|
| **用户期望 PDF 完美对齐 DOCX** | 高 | 中 | 文档明确说明两者差异，提供双引擎选择 |
| **模板过于复杂难以配置** | 中 | 中 | 提供预设模板，简化自定义 UI |
| **离线使用受限** | 低 | 低 | 文档说明网络图片需本地化 |

### 7.3 工程风险

| 风险 | 可能性 | 影响 | 缓解措施 |
|------|--------|------|----------|
| **测试覆盖不足** | 中 | 高 | 强制 CI 覆盖率检查，核心模块 > 80% |
| **文档滞后** | 高 | 中 | 每个 Phase 结束强制更新文档 |
| **代码重构导致回归** | 中 | 高 | 完善测试用例，重构前先增加测试 |

### 7.4 不做清单（明确边界）

为保持产品定位和开发聚焦，以下功能**明确不做**：

#### 7.4.1 不改变核心定位

| 功能 | 理由 |
|------|------|
| **服务端转换** | 违背 "Browser-only" 核心价值，引入隐私风险和运维成本 |
| **用户账户系统** | 纯工具类应用，无需登录，保持简洁 |
| **云存储同步** | 增加复杂度，不符合离线使用定位 |
| **在线协作编辑** | 定位是单人转换工具，非协作平台 |

#### 7.4.2 不破坏现有产品逻辑

| 功能 | 理由 |
|------|------|
| **重构 DOCX 生成逻辑** | 当前逻辑已稳定，避免引入回归风险 |
| **移除内置模板** | 用户可能依赖，保持兼容 |
| **改变模板配置格式** | 需提供迁移方案，成本高 |

#### 7.4.3 不做低价值功能

| 功能 | 理由 |
|------|------|
| **DOCX → Markdown 逆向转换** | 需求低，技术难度高 |
| **PDF 编辑功能** | 超出转换工具定位，已有专业工具 |
| **富文本所见即所得编辑器** | 定位是 Markdown 工具，非富文本编辑器 |
| **Markdown 语法检查** | 已有成熟工具（如 markdownlint），无需重复造轮子 |

#### 7.4.4 不做过度依赖功能

| 功能 | 理由 |
|------|------|
| **AI 自动生成内容** | 超出工具定位，可集成外部 AI 服务但不内置 |
| **图片 OCR** | 需第三方 API，破坏离线使用 |
| **实时翻译** | 超出范围，用户可用专业翻译工具 |

#### 7.4.5 不做技术风险过高功能

| 功能 | 理由 |
|------|------|
| **完整数学公式编辑器** | 维护成本极高，推荐用户用 LaTeX 预处理 |
| **复杂图表交互编辑** | 超出纯文档转换范围 |
| **视频/音频嵌入** | DOCX/PDF 支持有限，体验差 |

### 7.5 关键不变量（必须保持）

以下特性是产品的核心价值，**任何改动都必须保持**：

1. ✅ **Browser-only**：所有转换在客户端完成，无需服务器
2. ✅ **离线可用**：下载后可完全离线使用
3. ✅ **MIT 许可**：保持开源和免费
4. ✅ **隐私保护**：不上传用户数据
5. ✅ **国际化**：保持多语言支持
6. ✅ **模板系统**：用户可自定义和导出模板
7. ✅ **向后兼容**：新版本不破坏旧模板和用户数据

---

## 附录

### A. 技术栈版本锁定

```json
{
  "dependencies": {
    "marked": "^15.0.11",
    "docx": "^9.5.0",
    "file-saver": "^2.0.5",
    "i18next": "^25.2.1"
  },
  "devDependencies": {
    "webpack": "^5.91.0",
    "vitest": "^1.0.0"
  }
}
```

**升级原则**：
- 小版本自动更新（^）
- 大版本手动评估，需回归测试
- 每季度检查依赖安全更新

### B. 参考资源

**标准文档**：
- GB/T 9704-2012 《党政机关公文格式》
- GB/T 7713.1-2006 《学位论文编写规则》
- GB/T 7714-2015 《文后参考文献著录规则》

**技术文档**：
- [marked.js 官方文档](https://marked.js.org/)
- [docx.js 官方文档](https://docx.js.org/)
- [pdf-lib 官方文档](https://pdf-lib.js.org/)
- [html2pdf.js GitHub](https://github.com/eKoopmans/html2pdf.js)

**社区资源**：
- [Markdown Guide](https://www.markdownguide.org/)
- [CommonMark Spec](https://commonmark.org/)

### C. 更新日志

| 版本 | 日期 | 变更 |
|------|------|------|
| 1.0 | 2026-09-11 | 初版规划文档 |

### D. 维护者

**文档维护**：AI 辅助生成，需人工审核  
**反馈渠道**：GitHub Issues  
**更新频率**：每个 Phase 结束后更新

---

**文档结束**
