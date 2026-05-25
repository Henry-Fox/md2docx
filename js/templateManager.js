// Template manager — stores user-defined and built-in format templates

export const PAGE_SIZES = {
  A4:     { width: 11907, height: 16839 },
  A3:     { width: 16839, height: 23814 },
  Letter: { width: 12240, height: 15840 },
  B5:     { width: 10319, height: 14600 },
};

export function mmToTwips(mm)              { return Math.round(mm * 1440 / 25.4); }
export function ptToHalfPoints(pt)         { return Math.round(pt * 2); }
export function ptToTwips(pt)              { return Math.round(pt * 20); }
export function charsToTwips(n, fontSizePt){ return Math.round(n * fontSizePt * 20); }

// ── Built-in templates ───────────────────────────────────────────────────────

const BUILT_IN_TEMPLATES = [
  {
    id: 'official',
    name: '党政机关公文',
    description: 'GB/T 9704-2012 · 版心156×225mm · 22行×28字 · 仿宋三号',
    readonly: true,
    // 纸张 A4, 上37/下35/左28/右26 mm (版心156×225mm)
    page: { size: 'A4', orientation: 'portrait', marginTop: 37, marginBottom: 35, marginLeft: 28, marginRight: 26 },
    // 正文: 仿宋_GB2312 三号(16pt), 行距28pt固定, 首行缩进2字
    body:  { font: '仿宋_GB2312', fontSize: 16, lineSpacing: 28, firstLineIndent: 2, alignment: 'justified' },
    title: { font: '方正小标宋_GBK', fontSize: 22, bold: true,  alignment: 'center', color: '000000' },
    h1:    { font: '黑体',           fontSize: 16, bold: true,  alignment: 'left',   color: '000000' },
    h2:    { font: '楷体_GB2312',    fontSize: 16, bold: true,  alignment: 'left',   color: '000000' },
    h3:    { font: '仿宋_GB2312',    fontSize: 14, bold: true,  alignment: 'left',   color: '000000' },
    h4:    { font: '仿宋_GB2312',    fontSize: 12, bold: true,  alignment: 'left',   color: '000000' },
    h5:    { font: '仿宋_GB2312',    fontSize: 10.5, bold: true, alignment: 'left',  color: '000000' },
  },
  {
    id: 'academic',
    name: '学术论文',
    description: 'GB/T 7713.1-2006 · 适合期刊/学术报告 · 宋体小四',
    readonly: true,
    // 上下25mm, 左30mm, 右25mm
    page: { size: 'A4', orientation: 'portrait', marginTop: 25, marginBottom: 25, marginLeft: 30, marginRight: 25 },
    // 正文: 宋体 小四(12pt), 行距20pt, 首行缩进2字
    body:  { font: '宋体', fontSize: 12, lineSpacing: 20, firstLineIndent: 2, alignment: 'justified' },
    // 大标题: 黑体 小二(18pt), 居中
    title: { font: '黑体',   fontSize: 18, bold: true,  alignment: 'center', color: '000000' },
    // 章: 黑体 小三(15pt), 居中
    h1:    { font: '黑体',   fontSize: 15, bold: true,  alignment: 'center', color: '000000' },
    // 节: 黑体 四号(14pt), 左对齐
    h2:    { font: '黑体',   fontSize: 14, bold: true,  alignment: 'left',   color: '000000' },
    // 条: 黑体 小四(12pt)
    h3:    { font: '黑体',   fontSize: 12, bold: true,  alignment: 'left',   color: '000000' },
    h4:    { font: '宋体',   fontSize: 12, bold: true,  alignment: 'left',   color: '000000' },
    h5:    { font: '宋体',   fontSize: 10.5, bold: false, alignment: 'left', color: '000000' },
  },
  {
    id: 'thesis',
    name: '毕业论文',
    description: '教育部标准 · 左30右25mm · 宋体小四 · 1.5倍行距',
    readonly: true,
    // 左30/右25, 上下25mm
    page: { size: 'A4', orientation: 'portrait', marginTop: 25, marginBottom: 25, marginLeft: 30, marginRight: 25 },
    // 正文: 宋体 12pt, 行距22pt (约1.5倍), 首行缩进2字
    body:  { font: '宋体', fontSize: 12, lineSpacing: 22, firstLineIndent: 2, alignment: 'justified' },
    // 标题: 黑体 16pt, 居中
    title: { font: '黑体',   fontSize: 16, bold: true,  alignment: 'center', color: '000000' },
    // 一级章节: 黑体 14pt
    h1:    { font: '黑体',   fontSize: 14, bold: true,  alignment: 'left',   color: '000000' },
    // 二级节: 黑体 12pt
    h2:    { font: '黑体',   fontSize: 12, bold: true,  alignment: 'left',   color: '000000' },
    // 三级小节: 宋体 12pt 加粗
    h3:    { font: '宋体',   fontSize: 12, bold: true,  alignment: 'left',   color: '000000' },
    h4:    { font: '宋体',   fontSize: 12, bold: false, alignment: 'left',   color: '000000' },
    h5:    { font: '宋体',   fontSize: 10.5, bold: false, alignment: 'left', color: '000000' },
  },
  {
    id: 'legal',
    name: '司法文书',
    description: '法院/检察院文书格式 · 仿宋三号 · 版心同党政公文',
    readonly: true,
    // 同党政机关公文版心
    page: { size: 'A4', orientation: 'portrait', marginTop: 37, marginBottom: 35, marginLeft: 28, marginRight: 26 },
    // 正文: 仿宋_GB2312 三号(16pt), 行距28pt
    body:  { font: '仿宋_GB2312', fontSize: 16, lineSpacing: 28, firstLineIndent: 2, alignment: 'justified' },
    // 标题: 黑体 小二(18pt), 居中
    title: { font: '黑体',        fontSize: 18, bold: true,  alignment: 'center', color: '000000' },
    h1:    { font: '黑体',        fontSize: 16, bold: true,  alignment: 'left',   color: '000000' },
    h2:    { font: '楷体_GB2312', fontSize: 16, bold: false, alignment: 'left',   color: '000000' },
    h3:    { font: '仿宋_GB2312', fontSize: 14, bold: false, alignment: 'left',   color: '000000' },
    h4:    { font: '仿宋_GB2312', fontSize: 12, bold: false, alignment: 'left',   color: '000000' },
    h5:    { font: '仿宋_GB2312', fontSize: 10.5, bold: false, alignment: 'left', color: '000000' },
  },
  {
    id: 'general',
    name: '通用文档',
    description: '适合日常使用的通用格式 · 宋体小四 · 标准边距',
    readonly: true,
    page: { size: 'A4', orientation: 'portrait', marginTop: 25.4, marginBottom: 25.4, marginLeft: 31.8, marginRight: 31.8 },
    body:  { font: '宋体', fontSize: 12, lineSpacing: 24, firstLineIndent: 2, alignment: 'justified' },
    title: { font: '黑体', fontSize: 22, bold: true,  alignment: 'center', color: '000000' },
    h1:    { font: '黑体', fontSize: 16, bold: true,  alignment: 'left',   color: '000000' },
    h2:    { font: '黑体', fontSize: 14, bold: true,  alignment: 'left',   color: '000000' },
    h3:    { font: '宋体', fontSize: 12, bold: true,  alignment: 'left',   color: '000000' },
    h4:    { font: '宋体', fontSize: 10.5, bold: true, alignment: 'left',  color: '000000' },
    h5:    { font: '宋体', fontSize: 9,   bold: true,  alignment: 'left',  color: '000000' },
  },
];

// ── Storage keys ─────────────────────────────────────────────────────────────

const STORAGE_KEY = 'md2docx_user_templates';
const ACTIVE_KEY  = 'md2docx_active_template';
const CONFIG_VERSION = '1.1';

// ── TemplateManager ───────────────────────────────────────────────────────────

class TemplateManager {
  getAll()  { return [...BUILT_IN_TEMPLATES, ...this._load()]; }
  get(id)   { return this.getAll().find(t => t.id === id) || BUILT_IN_TEMPLATES[0]; }

  getActive() {
    const id = localStorage.getItem(ACTIVE_KEY) || 'official';
    return this.get(id);
  }
  setActive(id) { localStorage.setItem(ACTIVE_KEY, id); }

  save(template) {
    const builtinIds = BUILT_IN_TEMPLATES.map(t => t.id);
    if (!template.id || builtinIds.includes(template.id)) {
      template.id = 'user_' + Date.now();
    }
    template.readonly = false;
    const list = this._load();
    const idx  = list.findIndex(t => t.id === template.id);
    if (idx >= 0) list[idx] = template; else list.push(template);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(list));
    return template;
  }

  delete(id) {
    const list = this._load().filter(t => t.id !== id);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(list));
  }

  clone(id) {
    const src  = this.get(id);
    const copy = JSON.parse(JSON.stringify(src));
    copy.id       = 'user_' + Date.now();
    copy.name     = src.name + ' (副本)';
    copy.readonly = false;
    return this.save(copy);
  }

  // ── Config file export / import ──────────────────────────────────────────

  exportConfig() {
    const config = {
      version:        CONFIG_VERSION,
      exportDate:     new Date().toISOString(),
      activeTemplate: localStorage.getItem(ACTIVE_KEY) || 'official',
      userTemplates:  this._load(),
    };
    const blob = new Blob([JSON.stringify(config, null, 2)], { type: 'application/json' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = 'md2docx-config.json';
    a.click();
    URL.revokeObjectURL(url);
  }

  importConfig(jsonText) {
    let config;
    try {
      config = JSON.parse(jsonText);
    } catch {
      throw new Error('配置文件格式错误，请确认是有效的 JSON 文件');
    }
    if (!config.userTemplates || !Array.isArray(config.userTemplates)) {
      throw new Error('配置文件缺少 userTemplates 字段');
    }
    // Validate each template has minimum required fields
    const required = ['id','name','page','body','title','h1','h2','h3','h4','h5'];
    for (const tpl of config.userTemplates) {
      for (const key of required) {
        if (!(key in tpl)) throw new Error(`模板 "${tpl.name || tpl.id}" 缺少字段: ${key}`);
      }
      tpl.readonly = false;
    }
    // Store user templates (replace existing)
    localStorage.setItem(STORAGE_KEY, JSON.stringify(config.userTemplates));
    // Restore active template if it still exists after import
    if (config.activeTemplate) {
      const ids = [...BUILT_IN_TEMPLATES.map(t => t.id), ...config.userTemplates.map(t => t.id)];
      if (ids.includes(config.activeTemplate)) {
        localStorage.setItem(ACTIVE_KEY, config.activeTemplate);
      }
    }
    return config.userTemplates.length;
  }

  // ── Convert template to docx.js units ────────────────────────────────────

  toDocxStyles(template) {
    const { page, body, title, h1, h2, h3, h4, h5 } = template;
    const ps      = PAGE_SIZES[page.size] || PAGE_SIZES.A4;
    const portrait = page.orientation !== 'landscape';
    return {
      pageWidth:  portrait ? ps.width  : ps.height,
      pageHeight: portrait ? ps.height : ps.width,
      pageMargin: {
        top:    mmToTwips(page.marginTop),
        bottom: mmToTwips(page.marginBottom),
        left:   mmToTwips(page.marginLeft),
        right:  mmToTwips(page.marginRight),
      },
      body: {
        font:            body.font,
        fontSize:        ptToHalfPoints(body.fontSize),
        lineSpacing:     ptToTwips(body.lineSpacing),
        firstLineIndent: charsToTwips(body.firstLineIndent, body.fontSize),
        alignment:       body.alignment,
      },
      title: _hStyle(title),
      h1:    _hStyle(h1),
      h2:    _hStyle(h2),
      h3:    _hStyle(h3),
      h4:    _hStyle(h4),
      h5:    _hStyle(h5),
    };
  }

  _load() {
    try { return JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]'); }
    catch { return []; }
  }
}

function _hStyle(h) {
  return {
    font:      h.font,
    fontSize:  ptToHalfPoints(h.fontSize),
    bold:      h.bold,
    alignment: h.alignment,
    color:     h.color || '000000',
  };
}

export const templateManager = new TemplateManager();
export { BUILT_IN_TEMPLATES };
