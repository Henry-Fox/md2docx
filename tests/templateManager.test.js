/**
 * TemplateManager 单元测试
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { templateManager, BUILT_IN_TEMPLATES } from '../js/templateManager.js';

describe('TemplateManager', () => {
  beforeEach(() => {
    localStorage.clear();
  });

  describe('getAll', () => {
    it('应该返回所有内置模板', () => {
      const templates = templateManager.getAll();
      expect(templates.length).toBeGreaterThanOrEqual(BUILT_IN_TEMPLATES.length);
    });

    it('内置模板应该包含党政机关公文', () => {
      const templates = templateManager.getAll();
      const official = templates.find(t => t.id === 'official');
      expect(official).toBeDefined();
      expect(official.name).toBe('党政机关公文');
    });
  });

  describe('get', () => {
    it('应该通过ID获取模板', () => {
      const template = templateManager.get('official');
      expect(template).toBeDefined();
      expect(template.id).toBe('official');
    });

    it('获取不存在的模板时应该返回默认模板', () => {
      const template = templateManager.get('non-existent');
      expect(template).toBeDefined();
      expect(template.id).toBe('official');
    });
  });

  describe('getActive', () => {
    it('应该返回激活的模板', () => {
      const active = templateManager.getActive();
      expect(active).toBeDefined();
      expect(active.id).toBe('official'); // 默认激活
    });
  });

  describe('setActive', () => {
    it('应该设置激活的模板', () => {
      templateManager.setActive('academic');
      const active = templateManager.getActive();
      expect(active.id).toBe('academic');
    });
  });

  describe('save', () => {
    it('应该保存自定义模板', () => {
      const customTemplate = {
        id: 'custom_test',
        name: '测试模板',
        description: '用于测试',
        page: { size: 'A4', orientation: 'portrait', marginTop: 25, marginBottom: 25, marginLeft: 25, marginRight: 25 },
        body: { font: '宋体', fontSize: 12, lineSpacing: 20, firstLineIndent: 2, alignment: 'justified' },
        title: { font: '黑体', fontSize: 18, bold: true, alignment: 'center', color: '000000' },
        h1: { font: '黑体', fontSize: 16, bold: true, alignment: 'left', color: '000000' },
        h2: { font: '黑体', fontSize: 14, bold: true, alignment: 'left', color: '000000' },
        h3: { font: '黑体', fontSize: 12, bold: true, alignment: 'left', color: '000000' },
        h4: { font: '宋体', fontSize: 12, bold: true, alignment: 'left', color: '000000' },
        h5: { font: '宋体', fontSize: 10.5, bold: false, alignment: 'left', color: '000000' },
      };

      const saved = templateManager.save(customTemplate);
      expect(saved).toBeDefined();
      expect(saved.readonly).toBe(false);
    });
  });

  describe('toDocxStyles', () => {
    it('应该将模板转换为DOCX样式', () => {
      const template = templateManager.get('official');
      const styles = templateManager.toDocxStyles(template);

      expect(styles).toBeDefined();
      expect(styles.pageWidth).toBeDefined();
      expect(styles.pageHeight).toBeDefined();
      expect(styles.pageMargin).toBeDefined();
      expect(styles.body).toBeDefined();
      expect(styles.title).toBeDefined();
    });

    it('DOCX样式应该包含正确的页边距', () => {
      const template = templateManager.get('official');
      const styles = templateManager.toDocxStyles(template);

      expect(styles.pageMargin.top).toBeGreaterThan(0);
      expect(styles.pageMargin.bottom).toBeGreaterThan(0);
      expect(styles.pageMargin.left).toBeGreaterThan(0);
      expect(styles.pageMargin.right).toBeGreaterThan(0);
    });
  });

  describe('toPdfStyles', () => {
    it('应该将模板转换为PDF样式', () => {
      const template = templateManager.get('official');
      const styles = templateManager.toPdfStyles(template);

      expect(styles).toBeDefined();
      expect(styles.pageWidth).toBeDefined();
      expect(styles.pageHeight).toBeDefined();
      expect(styles.pageMargin).toBeDefined();
      expect(styles.body).toBeDefined();
      expect(styles.title).toBeDefined();
    });

    it('PDF样式应该包含字体信息', () => {
      const template = templateManager.get('academic');
      const styles = templateManager.toPdfStyles(template);

      expect(styles.body.font).toBeDefined();
      expect(styles.body.fontSize).toBeDefined();
      expect(styles.title.font).toBeDefined();
    });

    it('应该正确映射中文字体到PDF标准字体', () => {
      const template = templateManager.get('official');
      const styles = templateManager.toPdfStyles(template);

      // 仿宋 -> Helvetica
      expect(styles.body.font).toBe('Helvetica');
    });
  });

  describe('clone', () => {
    it('应该克隆模板', () => {
      const cloned = templateManager.clone('official');
      expect(cloned).toBeDefined();
      expect(cloned.id).not.toBe('official');
      expect(cloned.name).toContain('副本');
      expect(cloned.readonly).toBe(false);
    });
  });

  describe('delete', () => {
    it('应该删除自定义模板', () => {
      const custom = {
        id: 'test_delete',
        name: '待删除模板',
        page: { size: 'A4', orientation: 'portrait', marginTop: 25, marginBottom: 25, marginLeft: 25, marginRight: 25 },
        body: { font: '宋体', fontSize: 12, lineSpacing: 20, firstLineIndent: 2, alignment: 'justified' },
        title: { font: '黑体', fontSize: 18, bold: true, alignment: 'center', color: '000000' },
        h1: { font: '黑体', fontSize: 16, bold: true, alignment: 'left', color: '000000' },
        h2: { font: '黑体', fontSize: 14, bold: true, alignment: 'left', color: '000000' },
        h3: { font: '黑体', fontSize: 12, bold: true, alignment: 'left', color: '000000' },
        h4: { font: '宋体', fontSize: 12, bold: true, alignment: 'left', color: '000000' },
        h5: { font: '宋体', fontSize: 10.5, bold: false, alignment: 'left', color: '000000' },
      };

      templateManager.save(custom);
      templateManager.delete('test_delete');

      const templates = templateManager.getAll();
      const deleted = templates.find(t => t.id === 'test_delete');
      expect(deleted).toBeUndefined();
    });
  });
});
