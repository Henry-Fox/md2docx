# Md2Docx

[English](#english) | [中文](#中文) | [Français](#français) | [Español](#español) | [Русский](#русский) | [العربية](#العربية)

## English

A simple and efficient web-based tool for converting Markdown files to Word documents (DOCX format).

### Features
- Drag and drop support
- Real-time preview
- Multiple language support
- Pure frontend implementation
- No server required

### Usage
1. Open the webpage
2. Input Markdown text or upload a Markdown file
3. Click "Convert to DOCX" to download the converted file

### Technologies
- docx.js
- marked.js
- Pure HTML/CSS/JavaScript

## 中文

一个简单高效的基于网页的 Markdown 转 Word 文档（DOCX 格式）工具。

### 功能特点
- 支持拖放文件
- 实时预览
- 多语言支持
- 纯前端实现
- 无需服务器

### 使用方法
1. 打开网页
2. 输入 Markdown 文本或上传 Markdown 文件
3. 点击"转换为 DOCX"下载转换后的文件

### 技术栈
- docx.js
- marked.js
- 纯 HTML/CSS/JavaScript

## Français

Un outil web simple et efficace pour convertir des fichiers Markdown en documents Word (format DOCX).

### Fonctionnalités
- Support du glisser-déposer
- Aperçu en temps réel
- Support multilingue
- Implémentation frontend pure
- Pas de serveur requis

### Utilisation
1. Ouvrez la page web
2. Saisissez du texte Markdown ou téléchargez un fichier Markdown
3. Cliquez sur "Convertir en DOCX" pour télécharger le fichier converti

### Technologies
- docx.js
- marked.js
- HTML/CSS/JavaScript pur

## Español

Una herramienta web simple y eficiente para convertir archivos Markdown a documentos Word (formato DOCX).

### Características
- Soporte para arrastrar y soltar
- Vista previa en tiempo real
- Soporte multilingüe
- Implementación frontend pura
- No requiere servidor

### Uso
1. Abra la página web
2. Ingrese texto Markdown o cargue un archivo Markdown
3. Haga clic en "Convertir a DOCX" para descargar el archivo convertido

### Tecnologías
- docx.js
- marked.js
- HTML/CSS/JavaScript puro

## Русский

Простой и эффективный веб-инструмент для преобразования файлов Markdown в документы Word (формат DOCX).

### Особенности
- Поддержка перетаскивания
- Предварительный просмотр в реальном времени
- Многоязычная поддержка
- Чистая фронтенд-реализация
- Не требует сервера

### Использование
1. Откройте веб-страницу
2. Введите текст Markdown или загрузите файл Markdown
3. Нажмите "Конвертировать в DOCX" для загрузки преобразованного файла

### Технологии
- docx.js
- marked.js
- Чистый HTML/CSS/JavaScript

## العربية

أداة ويب بسيطة وفعالة لتحويل ملفات Markdown إلى مستندات Word (تنسيق DOCX).

### الميزات
- دعم السحب والإفلات
- معاينة في الوقت الفعلي
- دعم متعدد اللغات
- تنفيذ واجهة أمامية خالصة
- لا يتطلب خادم

### الاستخدام
1. افتح صفحة الويب
2. أدخل نص Markdown أو قم بتحميل ملف Markdown
3. انقر على "تحويل إلى DOCX" لتنزيل الملف المحول

### التقنيات
- docx.js
- marked.js
- HTML/CSS/JavaScript خالص

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
