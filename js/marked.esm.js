/**
 * marked.esm.js - 导出marked及其扩展
 */
import { marked } from 'marked';
import { gfmHeadingId } from 'marked-gfm-heading-id';
import { mangle } from 'marked-mangle';

// 导入 marked-extended-tables
// 由于模块导出的不是函数，我们应该直接使用它而不是调用它
const extendedTables = require('marked-extended-tables');

// 配置marked使用扩展
marked.use(gfmHeadingId(), mangle(), extendedTables);

// 导出配置好的marked
export { marked };
