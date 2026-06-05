// Parse a .docx file and extract its formatting as a template-compatible object.
import JSZip from 'jszip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

// Get attribute value prefixed with 'w:' (works for both namespaced and plain XML)
function wa(el, name) {
  if (!el) return null;
  return el.getAttribute('w:' + name) ?? el.getAttributeNS(W, name) ?? null;
}

// Get first direct child element with the given local name
function wc(parent, tag) {
  if (!parent) return null;
  for (const child of parent.childNodes) {
    if (child.nodeType === 1 && child.localName === tag) return child;
  }
  return null;
}

// Twips (1/1440 inch) to mm, 1 decimal place
const twipsToMm = (v) => parseFloat((parseInt(v, 10) * 25.4 / 1440).toFixed(1));
// Half-points to pt
const hpToPt = (v) => parseInt(v, 10) / 2;

function extractRPr(rPr) {
  const r = {};
  if (!rPr) return r;
  const fonts = wc(rPr, 'rFonts');
  if (fonts) {
    r.font = wa(fonts, 'eastAsia') || wa(fonts, 'ascii') || wa(fonts, 'hAnsi') || null;
  }
  const sz = wc(rPr, 'sz');
  if (sz) { const v = wa(sz, 'val'); if (v) r.fontSize = hpToPt(v); }
  const b = wc(rPr, 'b');
  if (b) { const v = wa(b, 'val'); r.bold = v !== '0' && v !== 'false'; }
  return r;
}

function extractPPr(pPr) {
  const r = {};
  if (!pPr) return r;
  const jc = wc(pPr, 'jc');
  if (jc) { const v = wa(jc, 'val'); r.alignment = v === 'both' ? 'justified' : v; }
  const sp = wc(pPr, 'spacing');
  if (sp) {
    const line = wa(sp, 'line');
    const rule = wa(sp, 'lineRule');
    if (line && rule === 'exact') r.lineSpacing = parseInt(line, 10) / 20;
  }
  const ind = wc(pPr, 'ind');
  if (ind) { const fl = wa(ind, 'firstLine'); if (fl) r.firstLineTwips = parseInt(fl, 10); }
  return r;
}

function directText(el) {
  return Array.from(el.getElementsByTagNameNS(W, 't')).map(t => t.textContent || '').join('').trim();
}

function mode(values) {
  const counts = new Map();
  for (const value of values.filter(v => v !== undefined && v !== null && v !== '')) {
    counts.set(value, (counts.get(value) || 0) + 1);
  }
  let best = null;
  for (const [value, count] of counts) {
    if (!best || count > best.count) best = { value, count };
  }
  return best ? best.value : undefined;
}

function firstDefined(...values) {
  return values.find(v => v !== undefined && v !== null && v !== '');
}

// Follow basedOn chain and merge properties (child overrides parent)
function resolveStyle(styleMap, id, depth = 0) {
  if (depth > 6 || !id || !styleMap[id]) return {};
  const el = styleMap[id];
  const basedOn = wc(el, 'basedOn');
  const parentId = basedOn ? wa(basedOn, 'val') : null;
  const parent = parentId ? resolveStyle(styleMap, parentId, depth + 1) : {};
  return { ...parent, ...extractRPr(wc(el, 'rPr')), ...extractPPr(wc(el, 'pPr')) };
}

function inferPageSize(w, h) {
  const SIZES = [['A4', 11907, 16839], ['A3', 16839, 23814], ['Letter', 12240, 15840], ['B5', 10319, 14600]];
  for (const [name, sw, sh] of SIZES) {
    if (Math.abs(w - sw) < 150 && Math.abs(h - sh) < 150) return name;
  }
  return 'A4';
}

export async function parseDocxStyles(file) {
  const zip = await JSZip.loadAsync(file);
  const parser = new DOMParser();

  const load = async (path) => {
    const entry = zip.file(path);
    if (!entry) throw new Error(`缺少 ${path}，请确认是有效的 .docx 文件`);
    return parser.parseFromString(await entry.async('string'), 'text/xml');
  };

  const [stylesDoc, docDoc] = await Promise.all([load('word/styles.xml'), load('word/document.xml')]);

  // Build style lookup maps
  const styleMap  = {};  // id → element
  const nameToId  = {};  // lowercase name → id
  for (const s of stylesDoc.getElementsByTagNameNS(W, 'style')) {
    const id = wa(s, 'styleId');
    if (!id) continue;
    styleMap[id] = s;
    const nameEl = wc(s, 'name');
    if (nameEl) { const n = wa(nameEl, 'val'); if (n) nameToId[n.toLowerCase()] = id; }
  }

  // Resolve a candidate id: try direct id first, then name lookup
  const resolve = (id) => styleMap[id] ? id : (nameToId[id.toLowerCase()] || id);
  const resolveList = (candidates) => candidates.map(resolve);
  const paragraphStyle = (p) => {
    const pPr = wc(p, 'pPr');
    const pStyle = wc(pPr, 'pStyle');
    const styleId = pStyle ? wa(pStyle, 'val') : null;
    const base = styleId ? resolveStyle(styleMap, styleId) : {};
    const run = wc(p, 'r');
    return { ...base, ...extractPPr(pPr), ...extractRPr(run ? wc(run, 'rPr') : null) };
  };
  const sampleParagraph = (candidates) => {
    const ids = new Set(resolveList(candidates));
    for (const p of docDoc.getElementsByTagNameNS(W, 'p')) {
      const pPr = wc(p, 'pPr');
      const pStyle = wc(pPr, 'pStyle');
      const styleId = pStyle ? wa(pStyle, 'val') : null;
      if (styleId && ids.has(styleId)) return paragraphStyle(p);
    }
    return {};
  };
  const paragraphs = Array.from(docDoc.getElementsByTagNameNS(W, 'p'))
    .map((p, index) => ({ index, text: directText(p), style: paragraphStyle(p) }))
    .filter(p => p.text);
  function toHeading(raw, fallbackSize, fallbackAlign = 'left') {
    return {
    font:      raw.font      || body.font,
    fontSize:  raw.fontSize  || fallbackSize,
    bold:      raw.bold      !== undefined ? raw.bold : true,
    alignment: raw.alignment || fallbackAlign,
    color:     '000000',
    };
  }

  // Page settings from last sectPr (body-level)
  const sectPrs = docDoc.getElementsByTagNameNS(W, 'sectPr');
  const sectPr  = sectPrs[sectPrs.length - 1] || null;
  const pgSz    = wc(sectPr, 'pgSz');
  const pgMar   = wc(sectPr, 'pgMar');

  const rawW = pgSz ? parseInt(wa(pgSz, 'w'), 10) : 11907;
  const rawH = pgSz ? parseInt(wa(pgSz, 'h'), 10) : 16839;
  const landscape = rawW > rawH;
  const [normW, normH] = landscape ? [Math.min(rawW, rawH), Math.max(rawW, rawH)] : [rawW, rawH];

  const page = {
    size:         inferPageSize(normW, normH),
    orientation:  landscape ? 'landscape' : 'portrait',
    marginTop:    pgMar ? twipsToMm(wa(pgMar, 'top'))    : 25.4,
    marginBottom: pgMar ? twipsToMm(wa(pgMar, 'bottom')) : 25.4,
    marginLeft:   pgMar ? twipsToMm(wa(pgMar, 'left'))   : 31.8,
    marginRight:  pgMar ? twipsToMm(wa(pgMar, 'right'))  : 31.8,
  };

  // Body style from Normal
  const bodyRaw  = resolveStyle(styleMap, resolve('Normal'));
  const bodySamples = paragraphs.filter(p =>
    p.style.fontSize &&
    !p.style.bold &&
    p.style.alignment !== 'center' &&
    p.text.length >= 12
  );
  const bodySample = bodySamples.length ? {
    font: mode(bodySamples.map(p => p.style.font)),
    fontSize: mode(bodySamples.map(p => p.style.fontSize)),
    lineSpacing: mode(bodySamples.map(p => p.style.lineSpacing)),
    firstLineTwips: mode(bodySamples.map(p => p.style.firstLineTwips)),
    alignment: mode(bodySamples.map(p => p.style.alignment)),
  } : {};
  const bodySize = firstDefined(bodySample.fontSize, bodyRaw.fontSize, 12);
  const body = {
    font:            bodyRaw.font        || '宋体',
    fontSize:        bodySize,
    lineSpacing:     firstDefined(bodySample.lineSpacing, bodyRaw.lineSpacing, 24),
    firstLineIndent: firstDefined(bodySample.firstLineTwips, bodyRaw.firstLineTwips)
      ? Math.max(0, Math.round(firstDefined(bodySample.firstLineTwips, bodyRaw.firstLineTwips) / (bodySize * 20)))
      : 2,
    alignment:       firstDefined(bodySample.alignment, bodyRaw.alignment, 'justified'),
  };
  body.font = firstDefined(bodySample.font, bodyRaw.font, body.font);
  const directTitle = paragraphs.find(p =>
    p.index < 5 &&
    p.style.fontSize >= Math.max(18, body.fontSize + 2) &&
    p.style.alignment === 'center'
  );
  const directHeadings = paragraphs.filter(p =>
    p.style.fontSize &&
    p.style.bold &&
    p.style.alignment !== 'center' &&
    p.text.length <= 80
  );
  const directHeadingByRank = (rank) => {
    const groups = [...new Set(directHeadings.map(p => `${p.style.font || ''}|${p.style.fontSize || ''}|${p.style.alignment || ''}`))];
    const key = groups[rank];
    if (!key) return {};
    return directHeadings.find(p => `${p.style.font || ''}|${p.style.fontSize || ''}|${p.style.alignment || ''}` === key)?.style || {};
  };

  // Build a heading style, trying each candidate id in order
  const heading = (candidates, fallbackSize, fallbackAlign = 'left') => {
    const direct = fallbackAlign === 'center'
      ? directTitle?.style
      : fallbackSize === 16
        ? directHeadingByRank(0)
        : fallbackSize === 14
          ? directHeadingByRank(1)
          : null;
    if (direct && (direct.font || direct.fontSize)) {
      return toHeading(direct, fallbackSize, fallbackAlign);
    }
    const sampled = sampleParagraph(candidates);
    if (sampled.font || sampled.fontSize) {
      return toHeading(sampled, fallbackSize, fallbackAlign);
    }
    for (const cand of candidates) {
      const id  = resolve(cand);
      if (!styleMap[id]) continue;
      const raw = resolveStyle(styleMap, id);
      if (raw.font || raw.fontSize) {
        return toHeading(raw, fallbackSize, fallbackAlign);
      }
    }
    return { font: body.font, fontSize: fallbackSize, bold: true, alignment: fallbackAlign, color: '000000' };
  };

  return {
    id:          'imported_' + Date.now(),
    name:        '从Word导入',
    description: `从Word文档提取 · ${page.size} · ${body.font} ${body.fontSize}pt`,
    readonly:    false,
    page,
    body,
    title: heading(['Title', 'title', '标题'],                              22, 'center'),
    h1:    heading(['Heading1', 'heading 1', '标题 1', '标题1'],            16, 'left'),
    h2:    heading(['Heading2', 'heading 2', '标题 2', '标题2'],            14, 'left'),
    h3:    heading(['Heading3', 'heading 3', '标题 3', '标题3'],            12, 'left'),
    h4:    heading(['Heading4', 'heading 4', '标题 4', '标题4'],          10.5, 'left'),
    h5:    heading(['Heading5', 'heading 5', '标题 5', '标题5'],             9, 'left'),
  };
}
