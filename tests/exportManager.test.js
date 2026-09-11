/**
 * ExportManager 单元测试
 */
import { describe, it, expect, vi } from 'vitest';
import { exportManager } from '../js/exportManager.js';

describe('ExportManager', () => {
  describe('constructor', () => {
    it('应该正确初始化', () => {
      expect(exportManager).toBeDefined();
      expect(exportManager.docxConverter).toBeDefined();
      expect(exportManager.pdfConverter).toBeDefined();
    });
  });

  describe('getSupportedFormats', () => {
    it('应该返回支持的格式列表', () => {
      const formats = exportManager.getSupportedFormats();
      expect(formats).toContain('docx');
      expect(formats).toContain('pdf');
    });
  });

  describe('isFormatSupported', () => {
    it('应该正确判断DOCX格式', () => {
      expect(exportManager.isFormatSupported('docx')).toBe(true);
      expect(exportManager.isFormatSupported('DOCX')).toBe(true);
    });

    it('应该正确判断PDF格式', () => {
      expect(exportManager.isFormatSupported('pdf')).toBe(true);
      expect(exportManager.isFormatSupported('PDF')).toBe(true);
    });

    it('应该拒绝不支持的格式', () => {
      expect(exportManager.isFormatSupported('txt')).toBe(false);
      expect(exportManager.isFormatSupported('html')).toBe(false);
    });
  });

  describe('export', () => {
    it('空Markdown应该抛出错误', async () => {
      await expect(exportManager.export('', 'docx')).rejects.toThrow('Markdown内容不能为空');
      await expect(exportManager.export('   ', 'pdf')).rejects.toThrow('Markdown内容不能为空');
    });

    it('不支持的格式应该抛出错误', async () => {
      await expect(exportManager.export('# Test', 'txt')).rejects.toThrow('不支持的导出格式');
    });
  });
});
