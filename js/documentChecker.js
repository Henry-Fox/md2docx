/**
 * 文档结构检查工具 - 检查 Markdown 文档的结构完整性
 * Document Structure Checker - Validate Markdown document structure
 */

import { marked } from 'marked';

/**
 * 检查文档结构问题
 * @param {string} markdown - Markdown 文本
 * @returns {Object} - { valid: boolean, warnings: Array, errors: Array }
 */
export function checkDocumentStructure(markdown) {
  const warnings = [];
  const errors = [];
  
  // 1. 检查空文档
  if (!markdown || markdown.trim().length === 0) {
    errors.push({
      type: 'empty_document',
      level: 'error',
      message: '文档为空',
      description: '请至少添加一些内容后再导出'
    });
    return { valid: false, warnings, errors };
  }
  
  // 2. 解析 Markdown 为 tokens
  let tokens;
  try {
    tokens = marked.lexer(markdown);
  } catch (err) {
    errors.push({
      type: 'parse_error',
      level: 'error',
      message: 'Markdown 解析失败',
      description: `解析错误:${err.message}`
    });
    return { valid: false, warnings, errors };
  }
  
  // 3. 提取所有标题
  const headings = [];
  for (const token of tokens) {
    if (token.type === 'heading') {
      headings.push({
        depth: token.depth,
        text: token.text,
        raw: token.raw
      });
    }
  }
  
  // 4. 检查是否有任何内容(非空白 tokens)
  const hasContent = tokens.some(t => 
    (t.type === 'paragraph' && t.text.trim().length > 0) ||
    t.type === 'heading' ||
    t.type === 'list' ||
    t.type === 'table' ||
    t.type === 'code'
  );
  
  if (!hasContent) {
    warnings.push({
      type: 'no_content',
      level: 'warning',
      message: '文档内容过少',
      description: '文档似乎没有实质性内容,导出的文档可能为空'
    });
  }
  
  // 5. 检查是否缺少一级标题 (H1)
  const hasH1 = headings.some(h => h.depth === 1);
  if (headings.length > 0 && !hasH1) {
    warnings.push({
      type: 'missing_h1',
      level: 'warning',
      message: '缺少一级标题',
      description: '建议使用一级标题(# 标题)作为文档主标题'
    });
  }
  
  // 6. 检查标题层级跳跃和嵌套问题
  if (headings.length > 0) {
    let prevDepth = 0;
    let firstH1Index = headings.findIndex(h => h.depth === 1);
    
    for (let i = 0; i < headings.length; i++) {
      const h = headings[i];
      
      // 6a. 检查是否有标题出现在第一个 H1 之前
      if (firstH1Index > 0 && i < firstH1Index && h.depth > 1) {
        warnings.push({
          type: 'heading_before_h1',
          level: 'warning',
          message: `${h.depth}级标题出现在文档主标题之前`,
          description: `标题"${h.text}"(H${h.depth})出现在第一个一级标题之前,可能导致文档结构混乱`
        });
      }
      
      // 6b. 检查标题层级跳跃 (例如 H1 -> H3, 跳过 H2)
      if (prevDepth > 0 && h.depth > prevDepth + 1) {
        warnings.push({
          type: 'heading_level_skip',
          level: 'warning',
          message: `标题层级跳跃:H${prevDepth} → H${h.depth}`,
          description: `标题"${h.text}"跳过了${h.depth - prevDepth - 1}级,建议按顺序使用标题层级`
        });
      }
      
      prevDepth = h.depth;
    }
  }
  
  // 7. 检查是否有过深的标题嵌套 (一般不超过 5 级)
  const deepHeadings = headings.filter(h => h.depth > 5);
  if (deepHeadings.length > 0) {
    warnings.push({
      type: 'deep_heading',
      level: 'warning',
      message: `使用了 ${deepHeadings.length} 个超过 5 级的标题`,
      description: '标题层级过深可能导致文档结构复杂,建议简化'
    });
  }
  
  // 8. 检查是否有重复的一级标题
  const h1s = headings.filter(h => h.depth === 1);
  if (h1s.length > 1) {
    warnings.push({
      type: 'multiple_h1',
      level: 'warning',
      message: `文档有 ${h1s.length} 个一级标题`,
      description: '建议只使用一个一级标题作为文档主标题,其他章节使用二级标题'
    });
  }
  
  // 9. 检查文档长度(如果非常短,可能不完整)
  const wordCount = markdown.trim().split(/\s+/).length;
  if (wordCount < 10 && warnings.length === 0) {
    warnings.push({
      type: 'short_document',
      level: 'info',
      message: '文档内容较短',
      description: `当前文档约 ${wordCount} 个词,请确认是否已完成编辑`
    });
  }
  
  const valid = errors.length === 0;
  return { valid, warnings, errors, headings, wordCount };
}

/**
 * 格式化检查结果为用户友好的消息
 * @param {Object} result - checkDocumentStructure 返回的结果
 * @returns {string} - 格式化的消息
 */
export function formatCheckResult(result) {
  const { valid, warnings, errors } = result;
  
  if (valid && warnings.length === 0) {
    return '✓ 文档结构检查通过';
  }
  
  let message = '';
  
  if (errors.length > 0) {
    message += '❌ 发现以下错误:\n';
    errors.forEach((err, i) => {
      message += `${i + 1}. ${err.message}: ${err.description}\n`;
    });
  }
  
  if (warnings.length > 0) {
    if (errors.length > 0) message += '\n';
    message += '⚠️ 发现以下建议:\n';
    warnings.forEach((warn, i) => {
      message += `${i + 1}. ${warn.message}\n   ${warn.description}\n`;
    });
  }
  
  return message.trim();
}

/**
 * 获取检查结果的简短摘要
 * @param {Object} result - checkDocumentStructure 返回的结果
 * @returns {string} - 简短摘要
 */
export function getCheckSummary(result) {
  const { valid, warnings, errors } = result;
  
  if (valid && warnings.length === 0) {
    return '文档结构良好';
  }
  
  if (errors.length > 0) {
    return `发现 ${errors.length} 个错误`;
  }
  
  if (warnings.length > 0) {
    return `发现 ${warnings.length} 个建议`;
  }
  
  return '未知状态';
}
