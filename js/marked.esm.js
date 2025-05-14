/**
 * 从node_modules重导出marked库
 */
import * as markedModule from 'marked';
export const marked = markedModule.marked;
