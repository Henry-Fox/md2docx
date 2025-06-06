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
- Support for various Markdown elements:
  - Headings (H1-H6)
  - Text styles (bold, italic, strikethrough)
  - Lists (ordered, unordered, task lists)
  - Blockquotes
  - Code blocks with syntax highlighting
  - Tables with alignment
  - Links and images
  - Horizontal rules
  - And more...

### Usage

#### Easiest Way (Recommended)
1. Download the latest release from [Releases](https://github.com/Henry-Fox/md2docx/releases)
2. Extract the downloaded file
3. Open `index.html` in a modern browser

#### Development Version
1. Clone the repository:
```bash
git clone https://github.com/Henry-Fox/md2docx.git
```
2. Install dependencies:
```bash
npm install
```
3. Build the project:
```bash
npm run build
```
4. Open `dist/index.html` in a modern browser

#### Development with Live Server
1. Clone the repository
2. Install dependencies
3. Run development server:
```bash
npm start
```

### Technologies
- marked.js (Markdown parsing)
- docx.js (Word document generation)
- Pure HTML/CSS/JavaScript

### Support the Project
If you find this tool helpful, consider buying me a coffee! Your support helps me continue improving this project.

![Support QR Code](./img/support-qr.png)

Thank you for your support! Your generosity helps make this project better for everyone. 🙏

## 中文

一个简单高效的基于网页的 Markdown 转 Word 文档（DOCX 格式）工具。

### 功能特点
- 支持拖放文件
- 实时预览
- 多语言支持
- 纯前端实现
- 无需服务器
- 支持多种 Markdown 元素：
  - 标题（H1-H6）
  - 文本样式（粗体、斜体、删除线）
  - 列表（有序、无序、任务列表）
  - 引用块
  - 代码块（支持语法高亮）
  - 表格（支持对齐方式）
  - 链接和图片
  - 水平线
  - 更多...

### 使用方法

#### 最简单的方式（推荐）
1. 从 [Releases](https://github.com/Henry-Fox/md2docx/releases) 下载最新版本
2. 解压下载的文件
3. 用现代浏览器打开 `index.html`

#### 开发版本
1. 克隆仓库：
```bash
git clone https://github.com/Henry-Fox/md2docx.git
```
2. 安装依赖：
```bash
npm install
```
3. 构建项目：
```bash
npm run build
```
4. 用现代浏览器打开 `dist/index.html`

#### 开发环境运行
1. 克隆仓库
2. 安装依赖
3. 运行开发服务器：
```bash
npm start
```

### 技术栈
- marked.js（Markdown 解析）
- docx.js（Word 文档生成）
- 纯 HTML/CSS/JavaScript

### 支持项目
如果您觉得这个工具对您有帮助，欢迎请我喝杯咖啡！您的支持是我继续改进这个项目的动力。

![支持二维码](./img/support-qr.png)

感谢您的支持！您的慷慨帮助让这个项目变得更好。🙏

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
