import { Md2Docx } from './md2docx.js';
import { marked } from './marked.esm.js';
import SimpleMd2Docx from './simpleMd2Docx.js';
import { Md2Json } from './md2json.js';

/**
 * @class App
 * @description 应用程序主类
 */
class App {
  constructor() {
    this.md2docx = new Md2Docx();

    // 强制重新初始化样式（从JSON文件加载）
    this.currentStyles = null;

    // 项目文件夹路径设置（默认路径）
    this.projectPath = "D:\\MD2DOCX\\";

    // 临时文件夹设置
    this.tempDirHandle = null;
    this.tempDirName = "";

    // 初始化DOM元素
    this.initElements();

    // 初始化事件监听
    this.initEventListeners();

    // 重新直接从md2docx加载默认样式
    this.currentStyles = this.md2docx.getDefaultStyles();

    // 初始化样式编辑器
    this.initStyleEditor();

    // 初始化预览渲染器
    this.initPreview();

    // 加载默认Markdown示例
    this.loadDefaultExample();

    // 初始化项目路径设置
    this.initProjectPath();

    // 尝试初始化临时目录
    this.initTempDirectory();

    // 更新界面后确认数据正确加载
    console.log("当前加载的样式:", JSON.stringify(this.currentStyles));
  }

  /**
   * @method initElements
   * @description 初始化元素
   */
  initElements() {
    // Markdown输入相关
    this.markdownInput = document.getElementById('markdown-input');
    this.fileInput = document.getElementById('file-input');
    this.dragArea = document.querySelector('.drag-area');

    // 按钮
    this.clearBtn = document.getElementById('clear-btn');
    this.convertBtn = document.getElementById('convert-btn');
    this.simpleConvertBtn = document.getElementById('simple-convert-btn'); // 获取简化版转换按钮
    this.resetStylesBtn = document.getElementById('reset-styles-btn');
    this.saveStylesBtn = document.getElementById('save-styles-btn');

    // 项目路径设置
    this.projectPathInput = document.getElementById('project-path-input');
    this.setProjectPathBtn = document.getElementById('set-project-path-btn');

    // 如果项目路径元素不存在，则创建
    if (!this.projectPathInput) {
      // 创建项目路径设置容器
      const pathContainer = document.createElement('div');
      pathContainer.className = 'project-path-container';

      // 创建标签
      const pathLabel = document.createElement('label');
      pathLabel.htmlFor = 'project-path-input';
      pathLabel.textContent = '项目路径:';
      pathContainer.appendChild(pathLabel);

      // 创建输入框
      this.projectPathInput = document.createElement('input');
      this.projectPathInput.id = 'project-path-input';
      this.projectPathInput.type = 'text';
      this.projectPathInput.className = 'project-path-input';
      this.projectPathInput.value = this.projectPath;
      this.projectPathInput.placeholder = 'D:\\MD2DOCX\\';
      pathContainer.appendChild(this.projectPathInput);

      // 创建保存按钮
      this.setProjectPathBtn = document.createElement('button');
      this.setProjectPathBtn.id = 'set-project-path-btn';
      this.setProjectPathBtn.className = 'btn';
      this.setProjectPathBtn.innerText = '保存项目路径';
      this.setProjectPathBtn.addEventListener('click', () => this.saveProjectPath());
      pathContainer.appendChild(this.setProjectPathBtn);

      // 添加到工具栏中合适位置
      const toolBar = document.querySelector('.toolbar') || document.querySelector('.controls');
      if (toolBar) {
        toolBar.appendChild(pathContainer);
      } else {
        // 如果没有找到工具栏，则添加到body
        document.body.insertBefore(pathContainer, document.body.firstChild);
      }
    }

    // 预览相关
    this.previewContainer = document.getElementById('preview-container');

    // 样式设置
    this.styleSettings = document.getElementById('style-settings');
  }

  /**
   * @method initEventListeners
   * @description 初始化事件监听器（全部加判空）
   */
  initEventListeners() {
    // 文件选择
    if (this.fileInput) {
      this.fileInput.addEventListener('change', (e) => this.handleFileSelect(e));
    }

    // 拖放功能
    if (this.markdownInput && this.dragArea) {
      this.markdownInput.addEventListener('dragenter', () => this.dragArea.classList.add('active'));
      this.markdownInput.addEventListener('dragleave', () => this.dragArea.classList.remove('active'));
      this.markdownInput.addEventListener('dragover', (e) => {
        e.preventDefault();
        this.dragArea.classList.add('active');
      });
      this.markdownInput.addEventListener('drop', (e) => {
        e.preventDefault();
        this.dragArea.classList.remove('active');
        this.handleFileDrop(e);
      });
      this.markdownInput.addEventListener('input', this.debounce(() => this.updatePreview(), 300));
    }

    // 按钮操作
    if (this.clearBtn) this.clearBtn.addEventListener('click', () => this.clearMarkdown());
    if (this.convertBtn) this.convertBtn.addEventListener('click', () => this.convertToDocx());
    if (this.simpleConvertBtn) {
      console.log('绑定简化版转换按钮事件');
      this.simpleConvertBtn.addEventListener('click', () => {
        console.log('点击了简化版转换按钮');
        this.simpleConvertToDocx();
      });
    } else {
      console.warn('未找到简化版转换按钮');
    }
    if (this.resetStylesBtn) this.resetStylesBtn.addEventListener('click', () => this.resetStyles());
    if (this.saveStylesBtn) this.saveStylesBtn.addEventListener('click', () => this.saveStyles());
    if (this.setProjectPathBtn) this.setProjectPathBtn.addEventListener('click', () => this.saveProjectPath());
  }

  /**
   * @method saveProjectPath
   * @description 保存项目路径设置
   */
  saveProjectPath() {
    const path = this.projectPathInput.value.trim();
    if (!path) {
      alert('请输入有效的项目路径！');
      return;
    }

    // 确保路径以斜杠结尾
    this.projectPath = path.endsWith('\\') ? path : path + '\\';
    this.projectPathInput.value = this.projectPath;

    // 保存设置以便下次使用
    localStorage.setItem('projectPath', this.projectPath);

    console.log(`项目路径已设置: ${this.projectPath}`);

    alert(`项目路径已设置为: ${this.projectPath}\n请确保以下文件夹存在:\n${this.projectPath}`);
  }

  /**
   * @method initProjectPath
   * @description 初始化项目路径
   */
  initProjectPath() {
    try {
      // 尝试从localStorage获取上次使用的项目路径
      const savedPath = localStorage.getItem('projectPath');
      if (savedPath) {
        this.projectPath = savedPath;
        this.projectPathInput.value = this.projectPath;
      }

      // 首次加载时提示
      const hasPrompted = localStorage.getItem('projectPathPrompted');
      if (!hasPrompted) {
        setTimeout(() => {
          alert(`当前项目路径设置为: ${this.projectPath}\n请确保以下文件夹存在:\n${this.projectPath}\n\n如需修改，请点击"保存设置"按钮。`);
          localStorage.setItem('projectPathPrompted', 'true');
        }, 2000);
      }
    } catch (error) {
      console.warn('初始化项目路径失败:', error);
    }
  }

  /**
   * @method handleFileSelect
   * @param {Event} event - 文件选择事件
   */
  handleFileSelect(e) {
    const file = e.target.files[0];
    if (file) {
      this.readFile(file);
    }
  }

  /**
   * @method handleFileDrop
   * @param {Event} event - 拖放事件
   */
  handleFileDrop(e) {
    const file = e.dataTransfer.files[0];
    if (file) {
      this.readFile(file);
    }
  }

  /**
   * @method readFile
   * @param {File} file - 要读取的文件
   */
  readFile(file) {
    if (file.type !== 'text/markdown' && file.type !== 'text/plain' && !file.name.endsWith('.md')) {
      alert('请选择Markdown文件！');
      return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
      this.markdownInput.value = e.target.result;
      this.updatePreview();
    };
    reader.readAsText(file);
  }

  /**
   * @method clearMarkdown
   * @description 清空输入区域
   */
  clearMarkdown() {
    this.markdownInput.value = '';
    this.updatePreview();
  }

  /**
   * @method convertToDocx
   * @description 转换当前的Markdown为Word文档
   */
  convertToDocx() {
    // 获取Markdown内容
    const markdownContent = this.markdownInput.value;

    // 检查Markdown内容是否为空
    if (!markdownContent.trim()) {
      this.showMessage('请先输入或上传Markdown内容', 'error');
      return;
    }

    // 检查Markdown语法错误
    const syntaxErrors = this.checkMarkdownSyntax(markdownContent);
    if (syntaxErrors.length > 0) {
      const errorMessage = `Markdown格式存在问题：\n${syntaxErrors.join('\n')}`;
      if (!confirm(`${errorMessage}\n\n是否继续生成文档？`)) {
        return;
      }
    }

    // 显示转换中消息
    this.showMessage('正在转换文档，请稍候...', 'info');

    // 预加载图片
    this.preloadImages(markdownContent)
      .then(images => {
        return this.convertMarkdownToDocx(markdownContent, images);
      })
      .then(blob => {
        const fileName = this.getOutputFileName();

        // 下载文件
        this.downloadBlob(blob, fileName);

        this.showMessage(`文档已成功生成：${fileName}`, 'success');
      })
      .catch(error => {
        console.error('转换文档时出错:', error);
        this.showMessage(`转换文档失败: ${error.message}`, 'error');
      });
  }

  /**
   * @method convertMarkdownToDocx
   * @description 将Markdown内容转换为Word文档
   * @param {string} markdownContent - Markdown内容
   * @param {Array} images - 图片信息数组
   * @returns {Promise<Blob>} - 返回Word文档的Blob对象
   */
  async convertMarkdownToDocx(markdownContent, images) {
    try {
      // 处理图片
      const imageInfos = await this.processImages(images);

      // 创建转换器实例，直接传入样式
      const converter = new Md2Docx(this.currentStyles);

      // 设置图片信息
      converter.setImageInfos(imageInfos);

      // 执行转换，返回docx的Blob对象
      return await converter.convert(markdownContent);

    } catch (error) {
      console.error('转换过程中出错:', error);
      throw new Error(`转换失败: ${error.message}`);
    } finally {
      // 清理临时文件
      if (images && images.length > 0) {
        this.cleanupTempFiles(images);
      }
    }
  }

  /**
   * @method processImages
   * @description 处理文档中的图片，准备用于Word文档
   * @param {NodeList} images - 文档中的图片元素
   * @return {Promise<Array>} 包含处理后图片信息的Promise
   */
  async processImages(images) {
    try {
      if (!images || images.length === 0) {
        console.log('文档中没有图片需要处理');
        return [];
      }

      console.log(`开始处理${images.length}张图片`);
      this.showMessage(`正在处理${images.length}张图片...`, 'info');

      // 图片信息数组
      const imageInfos = [];

      // 异步处理所有图片
      await Promise.all(Array.from(images).map(async (img, index) => {
        try {
          // 图片源路径
          const src = img.src;
          const alt = img.alt || `图片${index + 1}`;

          console.log(`处理图片 ${index + 1}/${images.length}: ${alt}`);

          // 下载图片
          const response = await fetch(src);
          if (!response.ok) {
            throw new Error(`获取图片失败: ${response.status} ${response.statusText}`);
          }

          // 将图片转换为ArrayBuffer
          const buffer = await response.arrayBuffer();

          // 读取图片宽高
          const dimensions = await this.getImageDimensions(img);

          // 添加到图片信息数组
          imageInfos.push({
            src: src,
            alt: alt,
            buffer: buffer,
            width: dimensions.width,
            height: dimensions.height,
            index: index
          });

          console.log(`图片 ${index + 1} 处理完成: ${dimensions.width}x${dimensions.height}`);
        } catch (error) {
          console.error(`处理图片 ${index + 1} 时出错:`, error);
          // 继续处理其他图片，不中断整个过程
        }
      }));

      console.log(`完成处理 ${imageInfos.length}/${images.length} 张图片`);
      return imageInfos;
    } catch (error) {
      console.error('处理图片时发生错误:', error);
      this.showMessage(`处理图片失败: ${error.message}`, 'error');
      return [];
    }
  }

  /**
   * @method getImageDimensions
   * @description 获取图片尺寸
   * @param {HTMLImageElement} img - 图片元素
   * @return {Promise<Object>} 包含宽高的Promise
   */
  getImageDimensions(img) {
    return new Promise((resolve) => {
      // 如果图片已加载完成，直接获取尺寸
      if (img.complete) {
        resolve({
          width: img.naturalWidth || 400,
          height: img.naturalHeight || 300
        });
      } else {
        // 否则等待图片加载完成
        img.onload = () => {
          resolve({
            width: img.naturalWidth || 400,
            height: img.naturalHeight || 300
          });
        };

        // 如果图片加载失败，使用默认尺寸
        img.onerror = () => {
          console.warn('图片加载失败，使用默认尺寸');
          resolve({ width: 400, height: 300 });
        };
      }
    });
  }

  /**
   * @method cleanupTempFiles
   * @description 清理临时文件
   * @param {Array} imageInfos - 图片信息数组
   * @return {Promise} 完成清理的Promise
   */
  async cleanupTempFiles(imageInfos) {
    // 实际上不需要清理临时文件，因为我们直接使用内存中的图片数据
    // 此方法保留作为兼容性扩展
    console.log('临时资源已清理');
    return Promise.resolve();
  }

  /**
   * @method normalizeUrl
   * @description 标准化URL，处理相对路径等
   * @param {string} url - 原始URL
   * @returns {string} 标准化后的URL
   */
  normalizeUrl(url) {
    let fullUrl = url;
    if (url.startsWith('//')) {
      fullUrl = 'https:' + url;
    } else if (url.startsWith('www.')) {
      fullUrl = 'https://' + url;
    } else if (!url.startsWith('http://') && !url.startsWith('https://')) {
      // 可能是相对路径
      if (url.startsWith('/')) {
        // 从域名根路径开始
        const currentLocation = window.location.origin;
        fullUrl = currentLocation + url;
      } else {
        // 从当前路径相对计算
        const currentPath = window.location.href.substring(0, window.location.href.lastIndexOf('/') + 1);
        fullUrl = currentPath + url;
      }
    }
    return fullUrl;
  }

  /**
   * @method findImageElement
   * @description 在预览区域中查找匹配的图片元素
   * @param {string} url - 图片URL
   * @param {string} altText - 图片alt文本
   * @returns {HTMLImageElement|null} 匹配的图片元素或null
   */
  findImageElement(url, altText) {
    const imgElements = Array.from(this.previewContainer.querySelectorAll('img'));
    console.log(`预览区域中共有 ${imgElements.length} 个图片元素`);

    // 方法1：URL完全匹配
    let img = imgElements.find(img => img.src === url || img.getAttribute('src') === url);

    // 方法2：alt属性匹配
    if (!img && altText) {
      img = imgElements.find(img => img.alt === altText);
      if (img) console.log(`通过alt属性找到图片: ${altText}`);
    }

    // 方法3：部分URL匹配
    if (!img) {
      // 提取文件名用于匹配
      let fileName = '';
      try {
        const urlObj = new URL(url);
        const pathSegments = urlObj.pathname.split('/').filter(Boolean);
        if (pathSegments.length > 0) {
          fileName = pathSegments[pathSegments.length - 1];
        }
      } catch (e) {
        // URL解析失败，尝试使用正则提取文件名
        const matches = url.match(/\/([^\/]+\.(png|jpe?g|gif|webp))($|\?)/i);
        if (matches) {
          fileName = matches[1];
        }
      }

      if (fileName) {
        img = imgElements.find(img => img.src.includes(fileName));
        if (img) console.log(`通过文件名 ${fileName} 找到图片`);
      }
    }

    // 方法4：尝试部分alt文本匹配
    if (!img && altText) {
      img = imgElements.find(img =>
        img.alt && (img.alt.includes(altText) || altText.includes(img.alt))
      );
      if (img) console.log(`通过部分alt文本匹配找到图片: ${img.alt}`);
    }

    // 如果仍未找到，记录日志
    if (!img) {
      console.warn(`未找到匹配的图片元素: ${url}, alt: ${altText}`);
    }

    return img;
  }

  /**
   * @method preloadImages
   * @description 预加载Markdown中的图片
   * @param {string} markdown - Markdown文本
   * @returns {Promise} Promise对象
   */
  preloadImages(markdown) {
    // 使用正则表达式提取所有图片URLs
    const imageRegex = /!\[.*?\]\((.*?)\)/g;
    const imageMatches = markdown.matchAll(imageRegex);
    const imageUrls = Array.from(imageMatches, match => match[1]);

    if (imageUrls.length === 0) {
      return Promise.resolve(); // 没有图片，直接返回已解析的Promise
    }

    // 创建Promise数组，加载每个图片
    const imagePromises = imageUrls.map(url => {
      // 忽略Base64图片，因为它们已经是内联数据
      if (url.startsWith('data:')) {
        return Promise.resolve();
      }

      // 尝试加载图片
      return new Promise((resolve) => {
        const img = new Image();
        img.onload = () => resolve();
        img.onerror = () => resolve(); // 即使加载失败也继续处理
        img.src = url;
      });
    });

    // 等待所有图片加载完成
    return Promise.all(imagePromises);
  }

  /**
   * @method loadDefaultExample
   * @description 加载默认示例
   */
  loadDefaultExample() {
    try {
      // 设置新的默认Markdown示例
      const defaultMarkdown = `# Markdown测试文档

## 1. 基础文本格式

这是普通段落文本。

**这是加粗文本**
*这是斜体文本*
***这是加粗斜体文本***`;

      if (this.markdownInput) {
        this.markdownInput.value = defaultMarkdown;
        this.updatePreview();
      }
    } catch (error) {
      console.error('加载默认示例失败:', error);
      this.loadFallbackExample();
    }
  }

  /**
   * @method loadFallbackExample
   * @description 加载内置的默认Markdown示例（当test.md加载失败时使用）
   */
  loadFallbackExample() {
    if (!this.markdownInput) return;

    this.markdownInput.value = `# Markdown示例文档

## 1. 基本文本格式

这是一个**粗体文本**示例，这是一个*斜体文本*示例，这是一个~~删除线~~示例。

这是另一个段落，包含一个[链接](https://example.com)和一些\`行内代码\`。

## 2. 列表示例

无序列表:

* 项目一
* 项目二
  * 子项目A
  * 子项目B
* 项目三

有序列表:

1. 第一步
2. 第二步
3. 第三步

## 3. 引用和代码块

> 这是一个引用的文本块。
>
> 引用可以包含多个段落。

\`\`\`javascript
// 这是一个代码块
function sayHello() {
  console.log("Hello, Markdown!");
}
\`\`\`

## 4. 表格示例

| 名称 | 年龄 | 职业 |
|------|-----|------|
| 张三 | 28  | 程序员 |
| 李四 | 32  | 设计师 |
| 王五 | 45  | 项目经理 |

## 5. 图片示例

![Markdown Logo](https://markdown-here.com/img/icon256.png)

## 6. 数学公式

E = mc^2

## 结束

这个示例展示了Markdown的常见格式，转换为Word后应保持良好的格式。`;

    // 更新预览
    this.updatePreview();
  }

  /**
   * @method initStyleEditor
   * @description 初始化样式编辑器
   */
  initStyleEditor() {
    const styleEditor = this.styleSettings;
    if (!styleEditor) return;

    // 清空现有内容
    styleEditor.innerHTML = '';

    // 创建样式设置分组
    const sections = [
      {
        id: 'document',
        title: '文档设置',
        fields: [
          { type: 'select', label: '纸张大小', name: 'document.pageSize', options: ['A4', 'Letter', 'Legal'] },
          { type: 'select', label: '纸张方向', name: 'document.pageOrientation', options: ['portrait', 'landscape'] },
          { type: 'number', label: '上边距 (twip/毫米)', name: 'document.margins.top', min: 0, step: 100, hasUnit: true, unitConverter: 'twipToMm' },
          { type: 'number', label: '下边距 (twip/毫米)', name: 'document.margins.bottom', min: 0, step: 100, hasUnit: true, unitConverter: 'twipToMm' },
          { type: 'number', label: '左边距 (twip/毫米)', name: 'document.margins.left', min: 0, step: 100, hasUnit: true, unitConverter: 'twipToMm' },
          { type: 'number', label: '右边距 (twip/毫米)', name: 'document.margins.right', min: 0, step: 100, hasUnit: true, unitConverter: 'twipToMm' }
        ],
        subsections: [
          {
            id: 'grid',
            title: '文档网格',
            fields: [
              { type: 'number', label: '每行字符数', name: 'document.grid.charPerLine', min: 10, max: 50, step: 1 },
              { type: 'number', label: '每页行数', name: 'document.grid.linePerPage', min: 10, max: 50, step: 1 }
            ]
          }
        ]
      },
      {
        id: 'heading1',
        title: '1级标题样式',
        fields: [
          { type: 'select', label: '字体', name: 'heading.fonts.h1', options: ['方正小标宋简体', '黑体', '楷体', '宋体', '仿宋_GB2312', '等线'] },
          { type: 'number', label: '字号 (pt)', name: 'heading.sizes.h1', min: 8, max: 72, step: 1 },
          { type: 'select', label: '对齐方式', name: 'heading.alignment.h1', options: ['left', 'center', 'right', 'justified'] },
          { type: 'checkbox', label: '加粗', name: 'heading.bold.h1' },
          { type: 'text', label: '前缀', name: 'heading.prefix.h1' },
          { type: 'checkbox', label: '使用前缀', name: 'heading.usePrefix.h1' },
          { type: 'color', label: '颜色', name: 'heading.colors.h1' }
        ]
      },
      {
        id: 'heading2',
        title: '2级标题样式',
        fields: [
          { type: 'select', label: '字体', name: 'heading.fonts.h2', options: ['方正小标宋简体', '黑体', '楷体', '宋体', '仿宋_GB2312', '等线'] },
          { type: 'number', label: '字号 (pt)', name: 'heading.sizes.h2', min: 8, max: 72, step: 1 },
          { type: 'select', label: '对齐方式', name: 'heading.alignment.h2', options: ['left', 'center', 'right', 'justified'] },
          { type: 'checkbox', label: '加粗', name: 'heading.bold.h2' },
          { type: 'text', label: '前缀', name: 'heading.prefix.h2' },
          { type: 'checkbox', label: '使用前缀', name: 'heading.usePrefix.h2' },
          { type: 'color', label: '颜色', name: 'heading.colors.h2' }
        ]
      },
      {
        id: 'heading3',
        title: '3级标题样式',
        fields: [
          { type: 'select', label: '字体', name: 'heading.fonts.h3', options: ['方正小标宋简体', '黑体', '楷体', '宋体', '仿宋_GB2312', '等线'] },
          { type: 'number', label: '字号 (pt)', name: 'heading.sizes.h3', min: 8, max: 72, step: 1 },
          { type: 'select', label: '对齐方式', name: 'heading.alignment.h3', options: ['left', 'center', 'right', 'justified'] },
          { type: 'checkbox', label: '加粗', name: 'heading.bold.h3' },
          { type: 'text', label: '前缀', name: 'heading.prefix.h3' },
          { type: 'checkbox', label: '使用前缀', name: 'heading.usePrefix.h3' },
          { type: 'color', label: '颜色', name: 'heading.colors.h3' }
        ]
      },
      {
        id: 'heading4',
        title: '4级标题样式',
        fields: [
          { type: 'select', label: '字体', name: 'heading.fonts.h4', options: ['方正小标宋简体', '黑体', '楷体', '宋体', '仿宋_GB2312', '等线'] },
          { type: 'number', label: '字号 (pt)', name: 'heading.sizes.h4', min: 8, max: 72, step: 1 },
          { type: 'select', label: '对齐方式', name: 'heading.alignment.h4', options: ['left', 'center', 'right', 'justified'] },
          { type: 'checkbox', label: '加粗', name: 'heading.bold.h4' },
          { type: 'text', label: '前缀', name: 'heading.prefix.h4' },
          { type: 'checkbox', label: '使用前缀', name: 'heading.usePrefix.h4' },
          { type: 'color', label: '颜色', name: 'heading.colors.h4' }
        ]
      },
      {
        id: 'heading5',
        title: '5级标题样式',
        fields: [
          { type: 'select', label: '字体', name: 'heading.fonts.h5', options: ['方正小标宋简体', '黑体', '楷体', '宋体', '仿宋_GB2312', '等线'] },
          { type: 'number', label: '字号 (pt)', name: 'heading.sizes.h5', min: 8, max: 72, step: 1 },
          { type: 'select', label: '对齐方式', name: 'heading.alignment.h5', options: ['left', 'center', 'right', 'justified'] },
          { type: 'checkbox', label: '加粗', name: 'heading.bold.h5' },
          { type: 'text', label: '前缀', name: 'heading.prefix.h5' },
          { type: 'checkbox', label: '使用前缀', name: 'heading.usePrefix.h5' },
          { type: 'color', label: '颜色', name: 'heading.colors.h5' }
        ]
      },
      {
        id: 'heading6',
        title: '6级标题样式',
        fields: [
          { type: 'select', label: '字体', name: 'heading.fonts.h6', options: ['方正小标宋简体', '黑体', '楷体', '宋体', '仿宋_GB2312', '等线'] },
          { type: 'number', label: '字号 (pt)', name: 'heading.sizes.h6', min: 8, max: 72, step: 1 },
          { type: 'select', label: '对齐方式', name: 'heading.alignment.h6', options: ['left', 'center', 'right', 'justified'] },
          { type: 'checkbox', label: '加粗', name: 'heading.bold.h6' },
          { type: 'text', label: '前缀', name: 'heading.prefix.h6' },
          { type: 'checkbox', label: '使用前缀', name: 'heading.usePrefix.h6' },
          { type: 'color', label: '颜色', name: 'heading.colors.h6' }
        ]
      },
      {
        id: 'paragraph',
        title: '正文样式',
        fields: [
          { type: 'select', label: '字体', name: 'paragraph.font', options: ['仿宋_GB2312', '宋体', '楷体', '黑体', '等线'] },
          { type: 'number', label: '字号 (pt)', name: 'paragraph.size', min: 8, max: 72, step: 1 },
          { type: 'color', label: '颜色', name: 'paragraph.color' },
          { type: 'select', label: '行距类型', name: 'paragraph.lineSpacingRule', options: ['auto', 'exact'] },
          { type: 'number', label: '行距值', name: 'paragraph.lineSpacing', min: 1, max: 100, step: 0.1 },
          { type: 'number', label: '段前后间距 (twip)', name: 'paragraph.spacing', min: 0, step: 20 },
          { type: 'number', label: '首行缩进 (twip)', name: 'paragraph.firstLineIndent', min: 0, step: 100 },
          { type: 'select', label: '对齐方式', name: 'paragraph.alignment', options: ['left', 'center', 'right', 'justified'] }
        ]
      },
      {
        id: 'list',
        title: '列表样式',
        subsections: [
          {
            id: 'unorderedList',
            title: '无序列表',
            fields: [
              { type: 'select', label: '字体', name: 'list.unordered.font', options: ['仿宋_GB2312', '宋体', '楷体', '黑体', '等线'] },
              { type: 'number', label: '字号 (pt)', name: 'list.unordered.size', min: 8, max: 72, step: 1 },
              { type: 'text', label: '一级项目符号', name: 'list.unordered.bulletChars.0' },
              { type: 'text', label: '二级项目符号', name: 'list.unordered.bulletChars.1' },
              { type: 'text', label: '三级项目符号', name: 'list.unordered.bulletChars.2' },
              { type: 'number', label: '缩进量 (twip)', name: 'list.unordered.indentLevel', min: 0, step: 100 }
            ]
          },
          {
            id: 'orderedList',
            title: '有序列表',
            fields: [
              { type: 'select', label: '字体', name: 'list.ordered.font', options: ['仿宋_GB2312', '宋体', '楷体', '黑体', '等线'] },
              { type: 'number', label: '字号 (pt)', name: 'list.ordered.size', min: 8, max: 72, step: 1 },
              { type: 'text', label: '一级编号格式', name: 'list.ordered.numberFormats.0' },
              { type: 'text', label: '二级编号格式', name: 'list.ordered.numberFormats.1' },
              { type: 'text', label: '三级编号格式', name: 'list.ordered.numberFormats.2' },
              { type: 'number', label: '缩进量 (twip)', name: 'list.ordered.indentLevel', min: 0, step: 100 }
            ]
          }
        ]
      },
      {
        id: 'table',
        title: '表格样式',
        fields: [
          { type: 'color', label: '边框颜色', name: 'table.borderColor' },
          { type: 'number', label: '边框宽度', name: 'table.borderWidth', min: 1, max: 10, step: 1 },
          { type: 'color', label: '表头背景色', name: 'table.headerBackground' },
          { type: 'select', label: '表头字体', name: 'table.headerFont', options: ['仿宋_GB2312', '宋体', '楷体', '黑体', '等线'] },
          { type: 'number', label: '字号 (pt)', name: 'table.fontSize', min: 8, max: 72, step: 1 },
          { type: 'select', label: '对齐方式', name: 'table.alignment', options: ['left', 'center', 'right'] }
        ]
      },
      {
        id: 'code',
        title: '代码样式',
        fields: [
          { type: 'select', label: '字体', name: 'code.font', options: ['等线', '宋体', 'Consolas', 'Courier New'] },
          { type: 'number', label: '字号 (pt)', name: 'code.size', min: 8, max: 72, step: 1 },
          { type: 'color', label: '文本颜色', name: 'code.color' },
          { type: 'color', label: '背景颜色', name: 'code.backgroundColor' }
        ]
      },
      {
        id: 'blockquote',
        title: '引用样式',
        fields: [
          { type: 'select', label: '字体', name: 'blockquote.font', options: ['仿宋_GB2312', '宋体', '楷体', '黑体', '等线'] },
          { type: 'number', label: '字号 (pt)', name: 'blockquote.size', min: 8, max: 72, step: 1 },
          { type: 'color', label: '文本颜色', name: 'blockquote.color' },
          { type: 'color', label: '边框颜色', name: 'blockquote.borderColor' },
          { type: 'number', label: '左侧缩进 (twip)', name: 'blockquote.leftIndent', min: 0, step: 100 },
          { type: 'number', label: '首行缩进 (twip)', name: 'blockquote.firstLineIndent', min: 0, step: 100 }
        ]
      }
    ];

    // 为每个部分创建样式设置区域
    sections.forEach(section => {
      // 创建部分容器
      const sectionDiv = document.createElement('div');
      sectionDiv.className = 'style-section';
      sectionDiv.innerHTML = `<h3 class="section-title">${section.title}</h3>`;

      // 创建字段
      if (section.fields && section.fields.length > 0) {
        const fieldsContainer = document.createElement('div');
        fieldsContainer.className = 'fields-container';

        section.fields.forEach(field => {
          const fieldDiv = this.createStyleField(field);
          fieldsContainer.appendChild(fieldDiv);
        });

        sectionDiv.appendChild(fieldsContainer);
      }

      // 处理子部分
      if (section.subsections && section.subsections.length > 0) {
        section.subsections.forEach(subsection => {
          // 创建子部分容器
          const subsectionDiv = document.createElement('div');
          subsectionDiv.className = 'style-subsection';
          subsectionDiv.innerHTML = `<h4 class="subsection-title">${subsection.title}</h4>`;

          // 创建子部分字段
          if (subsection.fields && subsection.fields.length > 0) {
            const subFieldsContainer = document.createElement('div');
            subFieldsContainer.className = 'fields-container';

            subsection.fields.forEach(field => {
              const fieldDiv = this.createStyleField(field);
              subFieldsContainer.appendChild(fieldDiv);
            });

            subsectionDiv.appendChild(subFieldsContainer);
          }

          sectionDiv.appendChild(subsectionDiv);
        });
      }

      styleEditor.appendChild(sectionDiv);
    });

    // 监听样式字段变化
    this.addStyleFieldListeners();

    // 添加行距类型变化监听
    const lineSpacingRuleField = document.querySelector('[data-field="paragraph.lineSpacingRule"]');
    if (lineSpacingRuleField) {
      lineSpacingRuleField.addEventListener('change', (e) => {
        // 获取行距值字段
        const lineSpacingField = document.querySelector('[data-field="paragraph.lineSpacing"]');
        if (lineSpacingField) {
          // 根据行距类型设置行距值的步长和提示
          if (e.target.value === 'exact') {
            lineSpacingField.step = '1';
            lineSpacingField.title = '固定行距，单位为磅值（pt）';
          } else {
            lineSpacingField.step = '0.1';
            lineSpacingField.title = '倍数行距，例如1.5表示1.5倍行距';
          }
        }
      });
    }
  }

  /**
   * @method createStyleField
   * @param {object} field - 字段对象
   */
  createStyleField(field) {
    const fieldDiv = document.createElement('div');
    fieldDiv.className = 'style-field';
    fieldDiv.setAttribute('data-field-name', field.name);

    const label = document.createElement('label');
    label.textContent = field.label;
    label.htmlFor = field.name.replace(/\./g, '-');
    fieldDiv.appendChild(label);

    // 从当前样式中获取初始值
    let initialValue = undefined;
    try {
      const pathParts = field.name.split('.');
      let obj = this.currentStyles;

      // 特殊处理，因为在JSON中正文样式有paragraph和"正文"两个，我们需要从两个地方都尝试获取值
      if (pathParts[0] === 'paragraph') {
        let paragraphObj = this.currentStyles.paragraph;
        if (paragraphObj && pathParts.length > 1) {
          initialValue = paragraphObj[pathParts[1]];
        }
      } else {
        // 常规路径处理
        for (const part of pathParts) {
          if (obj === undefined || obj === null) break;
          obj = obj[part];
        }
        initialValue = obj;
      }

      // 记录调试信息
      console.log(`字段 ${field.name} 的初始值:`, initialValue);
    } catch (error) {
      console.error(`获取字段 ${field.name} 初始值出错:`, error);
    }

    let input;
    switch (field.type) {
      case 'select':
        input = document.createElement('select');
        input.id = field.name.replace(/\./g, '-');
        input.setAttribute('data-field', field.name);
        field.options.forEach(option => {
          const optionEl = document.createElement('option');
          optionEl.value = option;
          optionEl.textContent = option;
          input.appendChild(optionEl);
        });
        // 设置初始值
        if (initialValue !== undefined) {
          input.value = initialValue;
        }
        break;

      case 'number':
        input = document.createElement('input');
        input.type = 'number';
        input.id = field.name.replace(/\./g, '-');
        input.setAttribute('data-field', field.name);
        if (field.min !== undefined) input.min = field.min;
        if (field.max !== undefined) input.max = field.max;
        if (field.step !== undefined) input.step = field.step;
        // 设置初始值
        if (initialValue !== undefined) {
          input.value = initialValue;
        }

        // 如果有单位转换，创建单位显示区域
        if (field.hasUnit && field.unitConverter) {
          const container = document.createElement('div');
          container.className = 'field-with-unit';

          container.appendChild(input);

          const unitSpan = document.createElement('span');
          unitSpan.className = 'unit-display';
          unitSpan.setAttribute('data-converter', field.unitConverter);

          // 直接计算并显示单位转换值
          if (initialValue !== undefined && !isNaN(initialValue) && this[field.unitConverter]) {
            const convertedValue = this[field.unitConverter](parseFloat(initialValue));
            unitSpan.textContent = `(${convertedValue} 毫米)`;
          }

          container.appendChild(unitSpan);

          // 当值改变时更新单位显示
          input.addEventListener('input', (e) => {
            const value = parseFloat(e.target.value);
            if (!isNaN(value) && this[field.unitConverter]) {
              const convertedValue = this[field.unitConverter](value);
              unitSpan.textContent = `(${convertedValue} 毫米)`;
            } else {
              unitSpan.textContent = '';
            }
          });

          fieldDiv.appendChild(container);
          return fieldDiv;
        }
        break;

      case 'color':
        input = document.createElement('input');
        input.type = 'color';
        input.id = field.name.replace(/\./g, '-');
        input.setAttribute('data-field', field.name);
        // 设置初始值，确保颜色值有#前缀
        if (initialValue !== undefined) {
          if (typeof initialValue === 'string' && !initialValue.startsWith('#')) {
            input.value = '#' + initialValue;
          } else {
            input.value = initialValue;
          }
        }
        break;

      case 'checkbox':
        input = document.createElement('input');
        input.type = 'checkbox';
        input.id = field.name.replace(/\./g, '-');
        input.setAttribute('data-field', field.name);
        // 设置初始值
        if (initialValue !== undefined) {
          input.checked = Boolean(initialValue);
        }
        break;

      case 'text':
      default:
        input = document.createElement('input');
        input.type = 'text';
        input.id = field.name.replace(/\./g, '-');
        input.setAttribute('data-field', field.name);
        // 设置初始值
        if (initialValue !== undefined) {
          input.value = initialValue;
        }
        break;
    }

    // 设置预定义值（如果有）
    if (field.value !== undefined) {
      if (input.type === 'checkbox') {
        input.checked = Boolean(field.value);
      } else {
        input.value = field.value;
      }
    }

    // 如果字段应该被禁用
    if (field.disabled) {
      input.disabled = true;
    }

    // 只有在没有单位的情况下才直接添加到fieldDiv
    if (!(field.hasUnit && field.unitConverter)) {
      fieldDiv.appendChild(input);
    }

    return fieldDiv;
  }

  /**
   * @method addStyleFieldListeners
   * @description 添加样式字段变化监听器
   */
  addStyleFieldListeners() {
    const styleFields = document.querySelectorAll('[data-field]');
    styleFields.forEach(field => {
      field.addEventListener('change', (e) => {
        const fieldPath = e.target.getAttribute('data-field');
        let value;

        // 根据输入类型获取值
        if (e.target.type === 'checkbox') {
          value = e.target.checked;
        } else if (e.target.type === 'number') {
          value = parseFloat(e.target.value);
        } else if (e.target.type === 'color') {
          // 移除颜色值中的#前缀，适应现有代码
          value = e.target.value.replace('#', '');
        } else {
          value = e.target.value;
        }

        // 更新样式对象
        this.updateStylesObject(fieldPath, value);

        // 更新预览
        this.updatePreview();
      });
    });
  }

  /**
   * @method updateStylesObject
   * @param {string} path - 字段路径
   * @param {any} value - 新值
   */
  updateStylesObject(path, value) {
    try {
      const parts = path.split('.');

      // 特殊处理正文样式，同时更新"正文"和paragraph两个节点
      if (parts[0] === 'paragraph') {
        const propertyName = parts[1]; // 获取属性名，如size、color等

        // 更新"正文"节点
        if (!this.currentStyles['paragraph']) {
          this.currentStyles['paragraph'] = {};
        }
        this.currentStyles['paragraph'][propertyName] = value;

        console.log(`更新样式: paragraph.${propertyName} = ${value}`);
      } else {
        // 常规路径处理
        let current = this.currentStyles;

        // 遍历路径，直到倒数第二级
        for (let i = 0; i < parts.length - 1; i++) {
          const part = parts[i];

          // 如果当前级别不存在，创建它
          if (current[part] === undefined) {
            // 检查下一级是否是数字索引
            const nextIsNumber = !isNaN(parseInt(parts[i + 1]));
            current[part] = nextIsNumber ? [] : {};
          }

          current = current[part];
        }

        // 设置最终值
        const lastPart = parts[parts.length - 1];
        current[lastPart] = value;
      }

      // 使用深拷贝更新当前样式，避免引用问题
      this.currentStyles = JSON.parse(JSON.stringify(this.currentStyles));
    } catch (error) {
      console.error('更新样式对象时出错:', error);
    }
  }

  /**
   * @method resetStyles
   * @description 重置样式
   */
  resetStyles() {
    if (confirm('确定要重置所有样式设置吗？')) {
      // 重置当前样式为默认样式
      this.currentStyles = this.md2docx.getDefaultStyles();

      // 重新初始化样式编辑器
      this.initStyleEditor();

      // 更新预览
      this.updatePreview();
    }
  }

  /**
   * @method saveStyles
   * @description 保存样式设置
   */
  saveStyles() {
    // 获取当前样式设置
    const styles = this.getStyleSettings();

    // 保存到localStorage
    localStorage.setItem('md2docx_styles', JSON.stringify(styles));
    alert('样式设置已保存！');
  }

  /**
   * @method loadSavedStyles
   * @description 加载已保存的样式设置
   */
  loadSavedStyles() {
    const savedStyles = localStorage.getItem('md2docx_styles');
    if (savedStyles) {
      try {
        // 解析保存的样式
        const styles = JSON.parse(savedStyles);

        // 更新当前样式
        this.currentStyles = styles;

        // 更新md2docx中的样式
        this.md2docx.setStyles(styles);

        // 重新加载样式编辑器
        this.initStyleEditor();

        return true;
      } catch (error) {
        console.error('加载样式设置时出错:', error);
        return false;
      }
    }
    return false;
  }

  /**
   * @method getStyleSettings
   * @description 获取样式设置
   */
  getStyleSettings() {
    // 使用当前样式对象而不是从DOM元素获取
    return this.currentStyles || this.md2docx.getDefaultStyles();
  }

  /**
   * @method debounce
   * @param {Function} func - 要防抖的函数
   * @param {number} delay - 延迟时间
   */
  debounce(func, delay) {
    let timeout;
    return function() {
      const context = this;
      const args = arguments;
      clearTimeout(timeout);
      timeout = setTimeout(() => func.apply(context, args), delay);
    };
  }

  /**
   * @method mmToTwip
   * @description 将毫米转换为twip单位
   * @param {number} mm - 毫米值
   * @returns {number} - twip值
   */
  mmToTwip(mm) {
    return Math.round(mm * 56.7); // 1毫米约等于56.7 twip
  }

  /**
   * @method twipToMm
   * @description 将twip单位转换为毫米
   * @param {number} twip - twip值
   * @returns {number} - 毫米值，保留2位小数
   */
  twipToMm(twip) {
    return parseFloat((twip / 56.7).toFixed(2)); // 保留两位小数
  }

  /**
   * @method processImgToBase64
   * @description 将图片元素转换为Base64
   * @param {HTMLImageElement} img - 图片元素
   * @returns {Promise<string>} Base64字符串
   */
  processImgToBase64(img) {
    // 存储this引用，以便在内部函数中使用
    const self = this;

    return new Promise((resolve, reject) => {
      try {
        console.log(`开始处理图片转Base64: ${img.src}`);

        // 记录图片状态
        console.log(`图片信息:
        - 尺寸: ${img.naturalWidth}x${img.naturalHeight}
        - 源地址: ${img.src}
        - 加载完成: ${img.complete}
        - crossOrigin属性: ${img.crossOrigin || '(无)'}
        `);

        // 检查图片是否已经是Base64格式
        if (img.src.startsWith('data:image/')) {
          console.log('图片已经是Base64格式，直接使用');
          resolve(img.src);
          return;
        }

        // 判断是否是跨域图片
        const isCrossOrigin = img.src.startsWith('http') && !img.src.startsWith(window.location.origin);

        // 方案1: 针对跨域图片
        if (isCrossOrigin) {
          console.log('检测到跨域图片，尝试直接绘制');

          // 直接尝试Canvas绘制
          tryCanvasWithoutProxy();
          return;
        }

        // 方案2: 直接从DOM元素创建Canvas（适用于同源图片）
        tryCanvasWithoutProxy();

        // 内部函数：尝试直接从图片元素创建Canvas
        function tryCanvasWithoutProxy() {
          try {
            // 创建Canvas并绘制图片
            const canvas = document.createElement('canvas');

            // 获取图片实际尺寸
            const width = img.naturalWidth;
            const height = img.naturalHeight;

            // 检查尺寸是否有效
            if (width === 0 || height === 0) {
              throw new Error(`图片尺寸无效: ${width}x${height}`);
            }

            // 设置Canvas尺寸
            canvas.width = width;
            canvas.height = height;

            console.log(`Canvas创建成功: ${width}x${height}`);

            // 获取绘图上下文
            const ctx = canvas.getContext('2d');

            // 填充白色背景（处理透明图片）
            ctx.fillStyle = '#FFFFFF';
            ctx.fillRect(0, 0, width, height);

            // 尝试绘制图片
            try {
              ctx.drawImage(img, 0, 0);

              // 判断图片类型
              let mimeType = 'image/png'; // 默认使用PNG
              const imgSrc = img.src || '';

              if (imgSrc.match(/\.jpe?g($|\?)/i)) {
                mimeType = 'image/jpeg';
              } else if (imgSrc.match(/\.gif($|\?)/i)) {
                mimeType = 'image/gif';
              } else if (imgSrc.match(/\.png($|\?)/i)) {
                mimeType = 'image/png';
              }

              console.log(`使用MIME类型: ${mimeType}`);

              // 设置转换质量
              let quality = mimeType === 'image/jpeg' ? 0.95 : undefined;

              try {
                // 获取Base64数据
                const dataUrl = canvas.toDataURL(mimeType, quality);

                // 验证结果
                if (!dataUrl || dataUrl === 'data:,') {
                  throw new Error('Canvas无法生成有效的图片数据');
                }

                console.log(`成功生成Base64数据: ${dataUrl.substring(0, 50)}... (${dataUrl.length} 字符)`);
                resolve(dataUrl);
              } catch (toDataURLError) {
                console.error(`Canvas.toDataURL失败: ${toDataURLError.message}`);

                // 检查是否是安全错误（跨域问题）
                if (toDataURLError.name === 'SecurityError' ||
                    toDataURLError.message.includes('security') ||
                    toDataURLError.message.includes('tainted')) {
                  console.warn('检测到Canvas安全错误，这是由于跨域限制导致的');

                  // 特殊处理markdown图标
                  if (img.src.includes('markdown-here.com/img/icon')) {
                    self.useMarkdownLogo(resolve);
                  } else {
                    self.useImagePlaceholder(resolve);
                  }
                } else {
                  // 其他类型错误，使用占位符
                  self.useImagePlaceholder(resolve);
                }
              }
            } catch (drawError) {
              console.error(`Canvas绘制图片失败: ${drawError.message}`);

              // 检查是否是markdown图标
              if (img.src.includes('markdown-here.com/img/icon')) {
                self.useMarkdownLogo(resolve);
              } else {
                self.useImagePlaceholder(resolve);
              }
            }
          } catch (error) {
            console.error('Canvas方法失败:', error);

            // 检查是否是markdown图标
            if (img.src.includes('markdown-here.com/img/icon')) {
              self.useMarkdownLogo(resolve);
            } else {
              self.useImagePlaceholder(resolve);
            }
          }
        }
      } catch (error) {
        console.error('处理图片过程中发生错误:', error);
        reject(new Error(`图片处理失败: ${error.message}`));
      }
    });
  }

  /**
   * @method useMarkdownLogo
   * @description 使用内置的Markdown图标
   * @param {Function} resolve - Promise的resolve函数
   */
  useMarkdownLogo(resolve) {
    console.log('使用内置的Markdown图标');
    // 高质量Markdown图标
    const markdownLogo = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAB10lEQVR4Xu2aTU7DMBCFbSQWLFggzpANO66CxIID9BZwDJAQopcBVpwFwQo2SIgVYmzZtRQJWXH+PNvzxsnG9rx5nxPPOLV5/4b0t9vtMpI1HA5pfHnV6/W+UW9vdEDQ5w8gIgQ9gLQH6L0cWGiAMOAQggh6Dof9EFzawCIz4sBDSCSAPBwJ+7FY+vRiMeSvcHly/OhUBULg8IQ4GCfxPnXrhwh3pxfR1dkNqfxQApBDn92/ZNsc3t6LN3V2kfHBAHuNuD+LDkDlQiCAyMDIAUgKRA5AUyBy8FMCZ5PXwhL40CnwOBE3uyUvQZQoIoTHbrdLz8umjY9Eg1zOTEYXMGHe/WsHMMMjBQ6HQJkPKO60VLrYVBi98JMXOlTxMeEq4YY2ADc/UEe4CcCYD4DGbbvhUc5NAGqOdP/c/EAd4SaAmhNo6QxYqCPYAMrR67H3nv9uNu/R/ed7K+L2+4fxoJ3OgDIAfpBVgKfJC329XeUAWvEOQDrDzBh+exS1AU6nU+p0OsUOTiaT4ruEadq2RQZQXsrq1oAqAA1jHACPBs+QgY9bA8Ij0hZuDWgLeQ+joQP8bw0ogHOLFhPQh4+EhQaESQy3BnCO9bZ4AdTe6Q6VsG4dAAAAAElFTkSuQmCC';
    resolve(markdownLogo);
  }

  /**
   * @method useImagePlaceholder
   * @description 使用通用图片占位符
   * @param {Function} resolve - Promise的resolve函数
   */
  useImagePlaceholder(resolve) {
    console.log('无法处理图片，使用占位符替代');
    // 更小的灰色占位符
    const placeholderBase64 = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGQAAABkCAMAAABHPGVmAAAAVFBMVEXMzMzm5ubNzc3d3d2ZmZmgoKCkpKTR0dHGxsa8vLzW1talra2vr6+ysrK3t7fJycmpqamcnJyVlZXBwcG0tLTT09Pi4uLZ2demnp6ioqKRkZGOjo64epnIAAAB/ElEQVRoge2a0XKrIBCGFyMqCoKaN+//oj1JexJNEyGg2Zm9+S86kw/Youyy1Go1juO6LkKu6z5Oy/sVxytV/EScp8afi3ZZtH8WJP9W3nhH/lB25EmCPNpJknhHBn2kMYbWGidJY02Ib0Kxs8QYtrkGQY2GoM6a0JzMbdt+3R2M1qagHlcBvL2J9LoKqVnGq4UBZFpTUMBfRsjNQj1rgDMEBVgLIpBdKWrFuVgKIqBdKZW1ILKSKWhXChessMJpM4uLFRF93HaDIFcfNyMnCILwhBXE8JolrKBE4wJZYQS5K/8pXp7KCbH2Hn+/UVSEEGdISIQhhLw9xBkSEoF+BgQgwZBACPEnMt0J6kO8IZ6QQAjxBj1AvCGBEOIP8YSkvw35+vbtUf26vX+/P/bv05Dv9eO+/bgBiWzHLKA1JrFDuO2YBdjdF5ckhNmOWYC1xqikhBDbMYt3f1CSlxC98rvSYbY6xcl1ot3/kZcYbvuKR4DZRf7wA5Vwd93xkQPtQ7g0X5CVjgrxK/OBI32IBV+UJeCOgSDwx0AQRMBvGWtC3M34vw5V8EOOQx16wG4Zew5kQ8yWsYJ3W8kHlhFbYbaMlYR47nKWEbMh6qAjcN7CvvTUFHzOJGfvNxnJh9WCVf975QTJTvL+3kjO3juf5JEz/8CUnb2Lr3bjC03P6mw4S+O5AAAAAElFTkSuQmCC';
    resolve(placeholderBase64);
  }

  /**
   * @method waitForImageLoad
   * @description 等待图片加载完成
   * @param {HTMLImageElement} img - 图片元素
   * @returns {Promise<void>} Promise对象
   */
  waitForImageLoad(img) {
    return new Promise((resolve, reject) => {
      if (!img) {
        reject(new Error('无效的图片元素'));
        return;
      }

      // 如果图片已经加载完成
      if (img.complete && img.naturalWidth > 0) {
        resolve();
        return;
      }

      // 设置加载事件
      const onLoad = () => {
        // 清理事件监听器
        img.removeEventListener('load', onLoad);
        img.removeEventListener('error', onError);
        resolve();
      };

      const onError = (error) => {
        // 清理事件监听器
        img.removeEventListener('load', onLoad);
        img.removeEventListener('error', onError);
        reject(new Error(`图片加载失败: ${error.message || '未知错误'}`));
      };

      // 添加事件监听
      img.addEventListener('load', onLoad);
      img.addEventListener('error', onError);

      // 设置超时，防止图片永远不加载
      setTimeout(() => {
        // 清理事件监听器
        img.removeEventListener('load', onLoad);
        img.removeEventListener('error', onError);

        // 如果图片已加载，正常解析
        if (img.complete && img.naturalWidth > 0) {
          resolve();
        } else {
          reject(new Error('图片加载超时'));
        }
      }, 5000);
    });
  }

  /**
   * @method initPreview
   * @description 初始化预览功能
   */
  initPreview() {
    // 初始化时更新一次预览
    this.updatePreview();
  }

  /**
   * @method updatePreview
   * @description 更新预览内容（仅用marked渲染）
   */
  updatePreview() {
    if (!this.previewContainer) return;

    const markdown = this.markdownInput.value;
    if (!markdown) {
      this.previewContainer.innerHTML = '<div class="preview-placeholder">转换后的预览将显示在这里</div>';
      return;
    }

    try {
      // 使用marked转换Markdown为HTML
      const html = marked.parse(markdown);
      this.previewContainer.innerHTML = html;
    } catch (error) {
      console.error('更新预览时出错:', error);
      this.previewContainer.innerHTML = '<div class="preview-error">预览生成失败</div>';
    }
  }

  /**
   * @method applyStylesToPreview
   * @param {string} html - 原始HTML内容
   * @returns {string} - 应用样式后的HTML
   */
  applyStylesToPreview(html) {
    // 创建临时容器
    const container = document.createElement('div');
    container.innerHTML = html;

    // 应用标题样式
    for (let i = 1; i <= 6; i++) {
      const headings = container.querySelectorAll(`h${i}`);
      const style = this.currentStyles.heading.styles[`h${i}`];

      headings.forEach(heading => {
        // 应用字体样式
        heading.style.fontFamily = `${style.font.name}, ${style.font.fallback.join(', ')}`;
        heading.style.fontSize = `${style.font.size}pt`;
        heading.style.fontWeight = style.font.bold ? 'bold' : 'normal';

        // 应用段落样式
        heading.style.textAlign = style.paragraph.alignment;
        heading.style.marginTop = `${style.paragraph.spacing.before / 20}pt`;
        heading.style.marginBottom = `${style.paragraph.spacing.after / 20}pt`;
        heading.style.paddingLeft = `${style.paragraph.indent.left / 20}pt`;
        heading.style.textIndent = `${style.paragraph.indent.firstLine / 20}pt`;

        // 应用编号样式
        if (style.numbering.usePrefix) {
          const prefix = document.createElement('span');
          prefix.className = 'heading-prefix';
          prefix.textContent = style.numbering.prefix;
          heading.insertBefore(prefix, heading.firstChild);
        }
      });
    }

    return container.innerHTML;
  }

  /**
   * @method setTempDirectory
   * @description 设置临时文件夹
   */
  async setTempDirectory() {
    try {
      // 请求用户选择文件夹
      this.tempDirHandle = await window.showDirectoryPicker({
        id: 'tempImagesDir',
        startIn: 'documents',
        mode: 'readwrite'
      });

      this.tempDirName = this.tempDirHandle.name;
      this.tempDirDisplay.innerText = `当前临时文件夹: ${this.tempDirName}`;
      console.log(`临时文件夹已设置: ${this.tempDirName}`);

      // 保存设置以便下次使用
      localStorage.setItem('tempDirName', this.tempDirName);

      return this.tempDirHandle;
    } catch (error) {
      console.error('设置临时文件夹失败:', error);
      this.tempDirDisplay.innerText = '未设置临时文件夹';
      return null;
    }
  }

  /**
   * @method initTempDirectory
   * @description 初始化临时目录和相关UI（可选功能）
   */
  async initTempDirectory() {
    try {
      // 添加临时目录显示区域
      const tempDirContainer = document.createElement('div');
      tempDirContainer.className = 'temp-dir-container';

      // 创建临时目录信息显示区域
      this.tempDirDisplay = document.createElement('div');
      this.tempDirDisplay.className = 'temp-dir-info';
      this.tempDirDisplay.innerText = '未设置临时文件夹（可选，仅用于包含图片的文档）';
      tempDirContainer.appendChild(this.tempDirDisplay);

      // 创建选择按钮
      const setTempDirBtn = document.createElement('button');
      setTempDirBtn.id = 'set-temp-dir-btn';
      setTempDirBtn.className = 'btn';
      setTempDirBtn.innerText = '选择临时文件夹(可选)';
      setTempDirBtn.addEventListener('click', () => this.setTempDirectory());
      tempDirContainer.appendChild(setTempDirBtn);

      // 添加到页面
      const toolBar = document.querySelector('.project-path-container');
      if (toolBar) {
        toolBar.parentNode.insertBefore(tempDirContainer, toolBar.nextSibling);
      } else {
        // 如果找不到项目路径容器，则添加到工具栏
        const fallbackToolbar = document.querySelector('.toolbar') || document.querySelector('.controls');
        if (fallbackToolbar) {
          fallbackToolbar.appendChild(tempDirContainer);
        } else {
          // 最后的选择：添加到body
          document.body.insertBefore(tempDirContainer, document.body.firstChild);
        }
      }

      // 尝试从localStorage获取上次使用的临时目录名称
      const savedDirName = localStorage.getItem('tempDirName');
      if (savedDirName) {
        this.tempDirName = savedDirName;
        this.tempDirDisplay.innerText = `上次使用的临时文件夹: ${savedDirName} (可选功能)`;
      }

      // 不再主动提示用户设置临时文件夹
      localStorage.setItem('tempDirPrompted', 'true');
    } catch (error) {
      console.warn('初始化临时目录失败:', error);
    }
  }

  /**
   * 检查Markdown语法的简单函数
   * @param {string} markdown - 要检查的Markdown文本
   * @returns {Object} 包含错误和警告的对象
   */
  checkMarkdownSyntax(markdown) {
    const result = {
      errors: []
    };

    // 检查未配对的星号 (*)
    const asteriskCount = (markdown.match(/\*/g) || []).length;
    if (asteriskCount % 2 !== 0) {
      result.errors.push(`发现${asteriskCount}个星号(*), 数量不成对`);
    }

    // 检查未配对的下划线 (_)
    const underscoreCount = (markdown.match(/_/g) || []).length;
    if (underscoreCount % 2 !== 0) {
      result.errors.push(`发现${underscoreCount}个下划线(_), 数量不成对`);
    }

    // 检查未配对的反引号 (`)
    const backtickCount = (markdown.match(/`/g) || []).length;
    if (backtickCount % 2 !== 0) {
      result.errors.push(`发现${backtickCount}个反引号(\`), 数量不成对`);
    }

    // 检查未闭合的链接语法
    const openBrackets = (markdown.match(/\[/g) || []).length;
    const closeBrackets = (markdown.match(/\]/g) || []).length;
    if (openBrackets !== closeBrackets) {
      result.errors.push(`方括号不匹配: [${openBrackets}个, ]${closeBrackets}个`);
    }

    // 检查可能是问题的星号格式，例如没有空格的加粗文本
    const problematicAsterisks = markdown.match(/[^\s\*]\*\*[^\s\*]|[^\s\*]\*[^\s\*]/g);
    if (problematicAsterisks && problematicAsterisks.length > 0) {
      result.errors.push(`发现${problematicAsterisks.length}处可能有问题的星号格式，建议在星号前后添加空格`);
    }

    return result;
  }

  /**
   * 动态加载脚本
   */
  loadScript(src) {
    return new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = src;
      script.onload = () => resolve();
      script.onerror = (err) => reject(new Error(`加载脚本失败: ${src}`));
      document.head.appendChild(script);
    });
  }

  /**
   * @method getOutputFileName
   * @description 获取输出文件名，基于当前时间
   * @return {string} 输出文件名
   */
  getOutputFileName() {
    const date = new Date();
    const timestamp =
      date.getFullYear().toString() +
      this.padZero(date.getMonth() + 1) +
      this.padZero(date.getDate()) +
      '_' +
      this.padZero(date.getHours()) +
      this.padZero(date.getMinutes());

    return `markdown_${timestamp}.docx`;
  }

  /**
   * @method padZero
   * @description 将数字前补零至两位
   * @param {number} num - 需要格式化的数字
   * @return {string} 格式化后的字符串
   */
  padZero(num) {
    return num < 10 ? '0' + num : num.toString();
  }

  /**
   * @method downloadBlob
   * @description 下载Blob对象
   * @param {Blob} blob - 要下载的Blob对象
   * @param {string} fileName - 下载的文件名
   */
  downloadBlob(blob, fileName) {
    if (!blob) {
      this.showMessage('文件生成失败', 'error');
      return;
    }

    try {
      // 创建下载链接
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;

      // 添加到文档并触发点击
      document.body.appendChild(a);
      a.click();

      // 清理
      setTimeout(() => {
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }, 100);

      this.showMessage(`文档已保存为 ${fileName}`, 'success');
    } catch (error) {
      console.error('下载文件时出错:', error);
      this.showMessage(`下载失败: ${error.message}`, 'error');
    }
  }

  /**
   * @method showMessage
   * @description 显示操作消息
   * @param {string} message - 消息内容
   * @param {string} type - 消息类型 (info, success, error, warning)
   */
  showMessage(message, type = 'info') {
    // 记录到控制台
    console.log(`[${type.toUpperCase()}] ${message}`);

    // 检查是否有消息容器
    let messageContainer = document.getElementById('message-container');

    // 如果没有消息容器，创建一个
    if (!messageContainer) {
      messageContainer = document.createElement('div');
      messageContainer.id = 'message-container';
      messageContainer.style.position = 'fixed';
      messageContainer.style.top = '20px';
      messageContainer.style.right = '20px';
      messageContainer.style.zIndex = '1000';
      document.body.appendChild(messageContainer);
    }

    // 创建消息元素
    const messageElement = document.createElement('div');
    messageElement.className = `message message-${type}`;
    messageElement.innerHTML = message;

    // 设置样式
    messageElement.style.padding = '10px 15px';
    messageElement.style.marginBottom = '10px';
    messageElement.style.borderRadius = '4px';
    messageElement.style.boxShadow = '0 2px 4px rgba(0,0,0,0.2)';
    messageElement.style.fontSize = '14px';
    messageElement.style.fontWeight = 'bold';
    messageElement.style.transition = 'all 0.3s ease';

    // 根据类型设置颜色
    switch (type) {
      case 'success':
        messageElement.style.backgroundColor = '#4caf50';
        messageElement.style.color = '#fff';
        break;
      case 'error':
        messageElement.style.backgroundColor = '#f44336';
        messageElement.style.color = '#fff';
        break;
      case 'warning':
        messageElement.style.backgroundColor = '#ff9800';
        messageElement.style.color = '#fff';
        break;
      default: // info
        messageElement.style.backgroundColor = '#2196f3';
        messageElement.style.color = '#fff';
    }

    // 添加到容器
    messageContainer.appendChild(messageElement);

    // 自动消失
    setTimeout(() => {
      messageElement.style.opacity = '0';
      setTimeout(() => {
        messageContainer.removeChild(messageElement);
      }, 300);
    }, 3000);
  }

  // 在App类中添加新的测试方法
  /**
   * 使用简化版转换器将Markdown转换为Docx
   * @param {string} markdown - Markdown文本
   * @returns {Promise<Blob>} Docx文件的Blob对象
   */
  async simpleConvertMarkdownToDocx(markdown) {
    try {
      console.log('[INFO] 开始简化版转换...');

      // 1. 使用md2json解析Markdown
      const md2json = new Md2Json();
      console.log('[INFO] 解析Markdown为JSON...');
      const jsonData = await md2json.convert(markdown);
      console.log('[INFO] 解析结果:', jsonData);

      // 2. 使用SimpleMd2Docx生成Docx
      const md2docx = new SimpleMd2Docx();
      console.log('[INFO] 生成Word文档...');
      const blob = await md2docx.convertToDocx(jsonData);

      // 3. 返回生成的blob
      return blob;
    } catch (error) {
      console.error('[ERROR] 简化版转换失败:', error);
      throw error;
    }
  }

  /**
   * 简化版转换为Docx的处理函数 - 直接调用test.js中的预设JSON数据
   */
  simpleConvertToDocx() {
    this.showMessage('正在使用test.js中的预设JSON数据生成文档，请稍候...', 'info');

    try {
      console.log('运行test.js的预设JSON数据转换...');

      // 动态导入test.js模块并执行
      import('./test.js')
        .then(testModule => {
          // 如果test.js导出了默认函数，则调用它
          if (typeof testModule.default === 'function') {
            return testModule.default();
          } else {
            throw new Error('test.js没有导出默认函数');
          }
        })
        .then(() => {
          console.log('test.js执行完成');
          this.showMessage('文档已成功生成', 'success');
        })
        .catch(error => {
          console.error('运行test.js时出错:', error);
          this.showMessage(`转换失败: ${error.message}`, 'error');
        });
    } catch (error) {
      console.error('转换失败:', error);
      this.showMessage(`转换失败: ${error.message}`, 'error');
    }
  }
}

// 初始化应用
document.addEventListener('DOMContentLoaded', () => {
  // 获取运行test.js的按钮
  const runTestJsBtn = document.getElementById('run-test-js-btn');

  // 如果按钮存在，添加点击事件监听器
  if (runTestJsBtn) {
    runTestJsBtn.addEventListener('click', async () => {
      try {
        console.log('运行test.js...');
        // 导入test.js模块
        const testModule = await import('./test.js');
        // 如果test.js导出了默认函数，则调用它
        if (typeof testModule.default === 'function') {
          await testModule.default();
        }
        console.log('test.js执行完成');
        alert('test.js执行完成，文件应已生成');
      } catch (error) {
        console.error('运行test.js时出错:', error);
        alert(`运行test.js时出错: ${error.message}`);
      }
    });
  }

  new App();
});



