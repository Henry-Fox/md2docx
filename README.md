# Md2Docx

一个将 Markdown 文档转换为 Word 文档（DOCX）的工具。

## 项目介绍

Md2Docx 是一个用于将 Markdown 文档转换为 Word 文档的工具。它采用两步转换的方式：
1. 首先将 Markdown 解析为结构化的 JSON
2. 然后基于 JSON 生成 Word 文档

这种方式使得转换过程更加灵活和可维护，同时也便于后续扩展其他格式的支持。

## 功能特性

### 已实现功能
- [x] Markdown 转 JSON
  - 支持基本文本格式（粗体、斜体、删除线等）
  - 支持标题（H1-H6）
  - 支持列表（有序、无序）
  - 支持引用块
  - 支持代码块（带语言标识）
  - 支持表格（带对齐方式）
  - 支持链接和图片
  - 支持水平线
  - 支持脚注
  - 支持嵌套的图片链接

### 计划功能
- [ ] JSON 转 DOCX
  - 实现基本文本格式转换
  - 实现表格转换
  - 实现图片插入
  - 实现链接处理
  - 实现脚注转换
  - 实现样式定制

## 技术栈

- JavaScript/Node.js
- docx.js（计划使用）

## 使用方法

### 安装

```bash
npm install md2docx
```

### 基本使用

```javascript
import { Md2Json } from './js/md2json.js';

// 创建解析器实例
const parser = new Md2Json();

// 解析 Markdown 文本
const json = parser.convert(markdownText);

// 输出 JSON
console.log(JSON.stringify(json, null, 2));
```

### JSON 结构示例

```json
{
  "type": "document",
  "children": [
    {
      "type": "heading",
      "level": 1,
      "rawText": "# 标题",
      "fullContent": "标题",
      "inlineStyles": [...]
    },
    {
      "type": "paragraph",
      "rawText": "这是一个**加粗**的段落",
      "fullContent": "这是一个加粗的段落",
      "inlineStyles": [...]
    }
  ]
}
```

## 开发计划

1. ✅ Markdown 转 JSON（已完成）
   - [x] 基本文本格式解析
   - [x] 复杂结构解析（表格、列表等）
   - [x] 样式信息提取
   - [x] 特殊元素处理（脚注、图片链接等）

2. 🔄 JSON 转 DOCX（进行中）
   - [ ] 基础文档结构生成
   - [ ] 文本样式应用
   - [ ] 表格生成
   - [ ] 图片处理
   - [ ] 链接处理
   - [ ] 脚注转换
   - [ ] 样式定制支持

3. 📝 后续计划
   - [ ] 命令行工具支持
   - [ ] 批量转换功能
   - [ ] 自定义样式模板
   - [ ] 其他格式支持（如 PDF）

## 贡献指南

欢迎提交 Issue 和 Pull Request！

1. Fork 本仓库
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启 Pull Request

## 许可证

本项目采用 MIT 许可证 - 详见 [LICENSE](LICENSE) 文件

## 联系方式

如有问题或建议，欢迎提交 Issue。
