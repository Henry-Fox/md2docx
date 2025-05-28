/**
 * marked.esm.js - 导出marked及其扩展
 */
import { marked } from 'marked';
import { gfmHeadingId } from 'marked-gfm-heading-id';
import { mangle } from 'marked-mangle';

// 导入 marked-extended-tables
// 由于模块导出的不是函数，我们应该直接使用它而不是调用它
const extendedTables = require('marked-extended-tables');

// 配置marked的基本选项
marked.setOptions({
  gfm: true,  // 启用 GitHub Flavored Markdown
  breaks: true,  // 启用换行符转换
  headerIds: true,  // 启用标题 ID
  mangle: false,  // 禁用链接引用混淆（对中文更友好）
  sanitize: false,  // 允许 HTML 标签
  smartLists: true,  // 使用更智能的列表行为
  smartypants: true,  // 使用更智能的标点符号
  xhtml: false  // 不使用 xhtml 规范
});

// 配置中文标题处理
marked.use(gfmHeadingId({
  prefix: 'heading-',
  transform: (id) => {
    // 处理中文标题 ID，保留中文字符
    return id.toLowerCase().replace(/[^\w\u4e00-\u9fa5]+/g, '-');
  }
}));

// 配置其他扩展
marked.use(mangle());
marked.use(extendedTables);

// 导出配置好的marked
export { marked };
