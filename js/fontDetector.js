/**
 * 字体检测工具 - 检测系统是否安装模板所需的中文字体
 * Font Detector - Detect if required Chinese fonts are installed on the system
 */

const COMMON_CHINESE_FONTS = {
  '仿宋_GB2312': { name: '仿宋_GB2312', aliases: ['FangSong', '仿宋'], family: 'serif' },
  '仿宋': { name: '仿宋', aliases: ['FangSong', '仿宋_GB2312'], family: 'serif' },
  '黑体': { name: '黑体', aliases: ['SimHei', 'Heiti SC', 'Heiti TC'], family: 'sans-serif' },
  '楷体_GB2312': { name: '楷体_GB2312', aliases: ['KaiTi', '楷体'], family: 'serif' },
  '楷体': { name: '楷体', aliases: ['KaiTi', '楷体_GB2312'], family: 'serif' },
  '宋体': { name: '宋体', aliases: ['SimSun', 'Songti SC', 'Songti TC'], family: 'serif' },
  '微软雅黑': { name: '微软雅黑', aliases: ['Microsoft YaHei', 'Microsoft YaHei UI'], family: 'sans-serif' },
  '方正小标宋_GBK': { name: '方正小标宋_GBK', aliases: ['FZXiaoBiaoSong-B05S', '方正小标宋'], family: 'serif' },
};

/**
 * 使用 Canvas 检测字体是否可用
 * @param {string} fontName - 字体名称
 * @param {string} fallbackFamily - 回退字体族 (serif/sans-serif)
 * @returns {boolean} - 是否可用
 */
function isFontAvailableCanvas(fontName, fallbackFamily = 'serif') {
  const canvas = document.createElement('canvas');
  const context = canvas.getContext('2d');
  
  const testString = '龍国國龙';  // 使用繁简混合的汉字测试
  const fontSize = 72;
  
  // 测试回退字体渲染
  context.font = `${fontSize}px ${fallbackFamily}`;
  const fallbackWidth = context.measureText(testString).width;
  
  // 测试目标字体渲染
  context.font = `${fontSize}px "${fontName}", ${fallbackFamily}`;
  const targetWidth = context.measureText(testString).width;
  
  // 如果宽度不同,说明字体可用
  return Math.abs(targetWidth - fallbackWidth) > 0.5;
}

/**
 * 使用 Font Loading API 检测字体
 * @param {string} fontName - 字体名称
 * @returns {Promise<boolean>} - 是否可用
 */
async function isFontAvailableFontAPI(fontName) {
  if (!document.fonts || !document.fonts.check) {
    return false;
  }
  
  try {
    // 检查字体是否已加载或可用
    return document.fonts.check(`16px "${fontName}"`);
  } catch (e) {
    return false;
  }
}

/**
 * 综合检测字体可用性(优先使用 Font API,回退到 Canvas)
 * @param {string} fontName - 字体名称
 * @param {string[]} aliases - 字体别名列表
 * @param {string} fallbackFamily - 回退字体族
 * @returns {Promise<boolean>} - 是否可用
 */
async function detectFont(fontName, aliases = [], fallbackFamily = 'serif') {
  // 1. 尝试主字体名
  if (await isFontAvailableFontAPI(fontName)) {
    return true;
  }
  
  // 2. 尝试别名
  for (const alias of aliases) {
    if (await isFontAvailableFontAPI(alias)) {
      return true;
    }
  }
  
  // 3. Canvas 回退检测
  if (isFontAvailableCanvas(fontName, fallbackFamily)) {
    return true;
  }
  
  for (const alias of aliases) {
    if (isFontAvailableCanvas(alias, fallbackFamily)) {
      return true;
    }
  }
  
  return false;
}

/**
 * 提取模板中使用的字体列表
 * @param {Object} template - 模板对象
 * @returns {Set<string>} - 字体名称集合
 */
export function extractTemplateFonts(template) {
  const fonts = new Set();
  
  if (template.body?.font) fonts.add(template.body.font);
  if (template.title?.font) fonts.add(template.title.font);
  if (template.h1?.font) fonts.add(template.h1.font);
  if (template.h2?.font) fonts.add(template.h2.font);
  if (template.h3?.font) fonts.add(template.h3.font);
  if (template.h4?.font) fonts.add(template.h4.font);
  if (template.h5?.font) fonts.add(template.h5.font);
  
  return fonts;
}

/**
 * 检测模板所需字体的可用性
 * @param {Object} template - 模板对象
 * @returns {Promise<Object>} - { available: string[], missing: string[] }
 */
export async function checkTemplateFonts(template) {
  const requiredFonts = extractTemplateFonts(template);
  const available = [];
  const missing = [];
  
  for (const fontName of requiredFonts) {
    const fontInfo = COMMON_CHINESE_FONTS[fontName];
    
    if (!fontInfo) {
      // 未知字体,假设可用(可能是英文字体或通用字体)
      available.push(fontName);
      continue;
    }
    
    const isAvailable = await detectFont(
      fontInfo.name,
      fontInfo.aliases,
      fontInfo.family
    );
    
    if (isAvailable) {
      available.push(fontName);
    } else {
      missing.push(fontName);
    }
  }
  
  return { available, missing };
}

/**
 * 获取字体下载/安装建议
 * @param {string} fontName - 字体名称
 * @returns {Object} - { message: string, links: Array }
 */
export function getFontInstallSuggestion(fontName) {
  const suggestions = {
    '仿宋_GB2312': {
      message: '仿宋_GB2312 是 Windows 系统预装字体,macOS/Linux 用户需要自行安装',
      links: [
        { text: '中标仿宋下载', url: 'https://www.fontke.com/font/' },
        { text: '字体安装教程', url: 'https://support.microsoft.com/zh-cn/office/添加字体-b7c5f17c-4426-4b53-967f-455339c564c1' }
      ]
    },
    '仿宋': {
      message: '仿宋是常用中文字体,macOS 可使用"华文仿宋",Linux 可安装开源替代字体',
      links: [
        { text: 'macOS 字体册', url: 'https://support.apple.com/zh-cn/guide/font-book/' },
        { text: 'Linux 字体安装', url: 'https://wiki.archlinux.org/title/Fonts' }
      ]
    },
    '黑体': {
      message: '黑体在 macOS 系统中为"黑体-简",Linux 可使用 WenQuanYi 或 Noto Sans CJK',
      links: [
        { text: 'WenQuanYi Zen Hei', url: 'http://wenq.org/wqy2/index.cgi?ZenHei' },
        { text: 'Noto Sans CJK', url: 'https://github.com/notofonts/noto-cjk' }
      ]
    },
    '楷体_GB2312': {
      message: '楷体_GB2312 是 Windows 系统预装字体,其他系统可使用"楷体"或"华文楷体"',
      links: [
        { text: '楷体下载', url: 'https://www.fontke.com/font/' }
      ]
    },
    '楷体': {
      message: 'macOS 用户可使用"楷体-简"或"华文楷体",Linux 可安装 AR PL UKai',
      links: [
        { text: 'AR PL UKai', url: 'https://www.freedesktop.org/wiki/Software/CJKUnifonts/' }
      ]
    },
    '宋体': {
      message: '宋体是最常见的中文字体,macOS 使用"宋体-简",Linux 可使用 WenQuanYi',
      links: []
    },
    '方正小标宋_GBK': {
      message: '方正小标宋是商业字体,常用于公文标题,需要购买授权或使用替代字体',
      links: [
        { text: '方正字库官网', url: 'https://www.foundertype.com/' }
      ]
    }
  };
  
  return suggestions[fontName] || {
    message: `字体"${fontName}"在当前系统中未找到,导出的文档可能使用替代字体`,
    links: []
  };
}

/**
 * 格式化缺失字体警告消息
 * @param {string[]} missingFonts - 缺失的字体列表
 * @returns {string} - 格式化的警告消息
 */
export function formatMissingFontsWarning(missingFonts) {
  if (missingFonts.length === 0) return '';
  
  const fontList = missingFonts.map(f => `"${f}"`).join('、');
  
  return `当前系统未检测到以下字体:${fontList}\n\n` +
         `导出的 Word 文档可能使用替代字体显示,建议:\n` +
         `1. Windows 用户:安装对应字体后重启浏览器\n` +
         `2. macOS/Linux 用户:使用等效字体(如"黑体-简"替代"黑体")\n` +
         `3. 如无法安装,可在导出后用 Word 手动调整字体`;
}

/**
 * 检测操作系统类型
 * @returns {string} - 'Windows' | 'macOS' | 'Linux' | 'Other'
 */
export function detectOS() {
  const ua = navigator.userAgent.toLowerCase();
  if (ua.includes('win')) return 'Windows';
  if (ua.includes('mac')) return 'macOS';
  if (ua.includes('linux')) return 'Linux';
  return 'Other';
}
