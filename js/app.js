import { Md2Docx } from './md2docx.js';
import { marked } from './marked.esm.js';

/**
 * @class App
 * @description 应用程序主类
 */
class App {
  constructor() {
    this.md2docx = new Md2Docx();

    // 强制重新初始化样式（从JSON文件加载）
    this.currentStyles = null;

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
    this.resetStylesBtn = document.getElementById('reset-styles-btn');
    this.saveStylesBtn = document.getElementById('save-styles-btn');

    // 预览相关
    this.previewContainer = document.getElementById('preview-container');

    // 样式设置
    this.styleSettings = document.getElementById('style-settings');
  }

  /**
   * @method initEventListeners
   * @description 初始化事件监听器
   */
  initEventListeners() {
    // 文件选择
    this.fileInput.addEventListener('change', (e) => this.handleFileSelect(e));

    // 拖放功能
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

    // 按钮操作
    this.clearBtn.addEventListener('click', () => this.clearMarkdown());
    this.convertBtn.addEventListener('click', () => this.convertToDocx());
    this.resetStylesBtn.addEventListener('click', () => this.resetStyles());
    this.saveStylesBtn.addEventListener('click', () => this.saveStyles());

    // Markdown输入变化时更新预览
    this.markdownInput.addEventListener('input', this.debounce(() => this.updatePreview(), 300));
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
   * @description 将Markdown转换为Word文档
   */
  convertToDocx() {
    const markdownContent = this.markdownInput.value.trim();
    if (!markdownContent) {
      alert('请先输入Markdown内容！');
      return;
    }

    try {
      // 显示加载提示
      const loadingMsg = document.createElement('div');
      loadingMsg.className = 'loading-message';
      loadingMsg.innerHTML = `
        <div class="loading-content">
          <div class="loading-step">1/3 正在处理图片...</div>
          <div class="loading-progress"></div>
        </div>
      `;
      document.body.appendChild(loadingMsg);

      // 更新加载进度
      const updateLoadingStep = (step, text) => {
        const stepElement = loadingMsg.querySelector('.loading-step');
        if (stepElement) {
          stepElement.textContent = `${step}/3 ${text}`;
        }
      };

      // 获取当前样式设置
      const styles = this.getStyleSettings();

      // 设置样式
      this.md2docx.setStyles(styles);

      // 处理图片
      updateLoadingStep(1, '正在处理和转换图片...');
      this.processImages(markdownContent)
        .then(processedMarkdown => {
          console.log("图片处理完成，开始生成docx...");
          updateLoadingStep(2, '正在转换Markdown为Word格式...');

          // 转换
          setTimeout(() => {
            try {
              this.md2docx.convert(processedMarkdown);
              updateLoadingStep(3, '正在生成Word文档...');

              // 保存文档
              setTimeout(() => {
                try {
                  this.md2docx.saveAsDocx();
                  // 移除加载提示
                  document.body.removeChild(loadingMsg);
                } catch (saveError) {
                  console.error('保存文档失败:', saveError);
                  alert(`保存文档失败: ${saveError.message}`);
                  document.body.removeChild(loadingMsg);
                }
              }, 300);
            } catch (convertError) {
              console.error('转换失败:', convertError);
              alert(`转换失败: ${convertError.message}`);
              document.body.removeChild(loadingMsg);
            }
          }, 300);
        })
        .catch(error => {
          console.error('图片处理失败:', error);

          if (confirm('图片处理过程中遇到问题，是否继续生成文档？(部分图片可能无法显示)')) {
            // 即使图片处理失败也尝试转换
            updateLoadingStep(2, '正在转换Markdown为Word格式...');

            try {
              this.md2docx.convert(markdownContent);
              updateLoadingStep(3, '正在生成Word文档...');

              setTimeout(() => {
                try {
                  this.md2docx.saveAsDocx();
                } catch (saveError) {
                  console.error('保存文档失败:', saveError);
                  alert(`保存文档失败: ${saveError.message}`);
                }
                document.body.removeChild(loadingMsg);
              }, 300);
            } catch (convertError) {
              console.error('转换失败:', convertError);
              alert(`转换失败: ${convertError.message}`);
              document.body.removeChild(loadingMsg);
            }
          } else {
            document.body.removeChild(loadingMsg);
          }
        });
    } catch (error) {
      console.error('转换失败:', error);
      alert(`转换失败: ${error.message}`);
      // 确保加载提示被移除
      const loadingMsg = document.querySelector('.loading-message');
      if (loadingMsg) {
        document.body.removeChild(loadingMsg);
      }
    }
  }

  /**
   * @method initPreview
   * @description 初始化预览渲染器
   */
  initPreview() {
    // 初始化预览区域
    this.previewContainer.innerHTML = '<div class="preview-placeholder">输入Markdown内容后将在此处显示预览</div>';
    marked.setOptions({
      renderer: new marked.Renderer(),
      highlight: function(code, lang) {
        return code;
      },
      gfm: true,
      breaks: true,
      headerIds: true,
      mangle: false
    });
  }

  /**
   * @method updatePreview
   * @description 更新预览
   */
  updatePreview() {
    const markdownContent = this.markdownInput.value.trim();

    if (!markdownContent) {
      this.previewContainer.innerHTML = '<div class="preview-placeholder">输入Markdown内容后将在此处显示预览</div>';
      return;
    }

    try {
      // 更新预览前预加载图片
      this.preloadImages(markdownContent).then(() => {
        this.previewContainer.innerHTML = marked.parse(markdownContent);
      }).catch(error => {
        console.error('预览生成失败:', error);
        this.previewContainer.innerHTML = marked.parse(markdownContent);
      });
    } catch (error) {
      console.error('预览生成失败:', error);
      this.previewContainer.innerHTML = '<div class="preview-error">预览生成失败，请检查Markdown格式！</div>';
    }
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
   * @method processImages
   * @description 处理Markdown中的图片，转换为Base64格式
   * @param {string} markdown - Markdown文本
   * @returns {Promise<string>} 处理后的Markdown文本
   */
  processImages(markdown) {
    return new Promise(async (resolve, reject) => {
      try {
        // 使用正则表达式提取所有图片URLs
        const imageRegex = /!\[(.*?)\]\((.*?)\)/g;
        let processedMarkdown = markdown;
        const imageMatches = Array.from(markdown.matchAll(imageRegex));

        if (imageMatches.length === 0) {
          return resolve(markdown); // 没有图片，直接返回原文本
        }

        console.log(`发现 ${imageMatches.length} 个图片，开始处理...`);

        // 计数器，用于追踪处理完成的图片数
        let processedCount = 0;
        let successCount = 0;

        // 处理每个图片URL
        for (const match of imageMatches) {
          const [fullMatch, altText, url] = match;
          console.log(`处理图片: ${url}`);

          // 如果已经是Base64格式，验证其格式
          if (url.startsWith('data:')) {
            // 验证Base64格式是否正确
            const isValidBase64 = /^data:image\/[a-zA-Z+.-]+;base64,[A-Za-z0-9+/=]+$/.test(url);
            if (isValidBase64) {
              console.log(`图片已经是有效的Base64格式，跳过处理`);
              processedCount++;
              successCount++;
            } else {
              console.warn(`检测到无效的Base64数据格式，尝试修复...`);
              try {
                // 尝试修复格式 - 提取MIME类型和数据部分
                const parts = url.split(',');
                if (parts.length >= 2) {
                  const mimeMatch = parts[0].match(/^data:(image\/[a-zA-Z+.-]+);base64$/);
                  if (mimeMatch && mimeMatch[1]) {
                    const mimeType = mimeMatch[1];
                    const base64Data = parts[1].trim();
                    // 重新构建正确格式的Base64字符串
                    const fixedBase64 = `data:${mimeType};base64,${base64Data}`;
                    processedMarkdown = processedMarkdown.replace(
                      fullMatch,
                      `![${altText}](${fixedBase64})`
                    );
                    console.log(`Base64格式已修复`);
                  }
                }
              } catch (e) {
                console.error(`无法修复Base64格式:`, e);
              }
              processedCount++;
            }

            // 检查是否所有图片都已处理完毕
            if (processedCount === imageMatches.length) {
              console.log(`所有图片处理完成，成功: ${successCount}/${imageMatches.length}`);
              resolve(processedMarkdown);
            }
            continue;
          }

          // 处理远程URL
          try {
            // 对于远程URL，需要使用fetch API获取
            if (url.startsWith('http') || url.startsWith('https') ||
                url.startsWith('//') || url.startsWith('www.')) {
              // 确保完整URL
              const fullUrl = url.startsWith('//') ? 'https:' + url :
                             url.startsWith('www.') ? 'https://' + url : url;

              console.log(`获取远程图片: ${fullUrl}`);

              try {
                const response = await fetch(fullUrl, {
                  method: 'GET',
                  mode: 'cors', // 尝试跨域获取
                  cache: 'no-cache', // 不使用缓存
                  headers: {
                    'Accept': 'image/*'
                  }
                });

                if (!response.ok) {
                  throw new Error(`获取图片失败: ${response.status} ${response.statusText}`);
                }

                const contentType = response.headers.get('content-type');
                if (!contentType || !contentType.startsWith('image/')) {
                  throw new Error(`返回的内容不是图片: ${contentType}`);
                }

                const blob = await response.blob();
                const base64 = await this.blobToBase64(blob);

                // 验证生成的Base64字符串
                if (!base64 || !base64.startsWith('data:image/')) {
                  throw new Error('无效的Base64图片数据');
                }

                // 替换原始URL为Base64
                processedMarkdown = processedMarkdown.replace(
                  fullMatch,
                  `![${altText}](${base64})`
                );

                console.log(`图片转换成功: ${fullUrl.substring(0, 30)}...`);
                successCount++;
              } catch (fetchError) {
                console.error(`获取图片失败: ${fullUrl}`, fetchError);
                // 继续处理其他图片
              }
            } else {
              // 本地文件路径处理尝试
              console.log(`尝试处理本地图片路径: ${url}`);
              // 对于demo示例中的相对路径，我们可以尝试使用fetch加载
              try {
                const response = await fetch(url, {
                  method: 'GET',
                  cache: 'no-cache'
                });

                if (response.ok) {
                  const contentType = response.headers.get('content-type');
                  if (contentType && contentType.startsWith('image/')) {
                    const blob = await response.blob();
                    const base64 = await this.blobToBase64(blob);

                    // 替换原始URL为Base64
                    processedMarkdown = processedMarkdown.replace(
                      fullMatch,
                      `![${altText}](${base64})`
                    );

                    console.log(`本地图片转换成功: ${url}`);
                    successCount++;
                  }
                } else {
                  console.warn(`无法加载本地图片: ${url}`);
                }
              } catch (e) {
                console.warn(`处理本地图片失败: ${url}`, e);
              }
            }
          } catch (error) {
            console.warn(`处理图片 ${url} 失败:`, error);
          }

          // 更新计数器
          processedCount++;
          // 检查是否所有图片都已处理完毕
          if (processedCount === imageMatches.length) {
            console.log(`所有图片处理完成，成功: ${successCount}/${imageMatches.length}`);
            resolve(processedMarkdown);
          }
        }
      } catch (error) {
        console.error('处理图片过程中出错:', error);
        reject(error);
      }
    });
  }

  /**
   * @method blobToBase64
   * @description 将Blob对象转换为Base64字符串
   * @param {Blob} blob - Blob对象
   * @returns {Promise<string>} Base64字符串
   */
  blobToBase64(blob) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => {
        try {
          const result = reader.result;
          // 验证Base64结果格式
          if (!result || typeof result !== 'string' || !result.startsWith('data:')) {
            throw new Error('转换结果不是有效的Base64数据');
          }

          // 确保MIME类型正确
          const mimeType = blob.type || 'image/png';
          if (!result.includes(mimeType)) {
            // 需要修正MIME类型
            const base64Data = result.split(',')[1];
            const correctedBase64 = `data:${mimeType};base64,${base64Data}`;
            console.log(`已修正Base64 MIME类型为 ${mimeType}`);
            console.log(`Base64转换完成，数据长度: ${correctedBase64.length}`);
            resolve(correctedBase64);
          } else {
            console.log(`Base64转换完成，数据长度: ${result.length}`);
            resolve(result);
          }
        } catch (error) {
          console.error('处理Base64结果时出错:', error);
          reject(error);
        }
      };
      reader.onerror = (error) => {
        console.error('Base64转换失败:', error);
        reject(error);
      };
      reader.readAsDataURL(blob);
    });
  }

  /**
   * @method loadDefaultExample
   * @description 加载默认Markdown示例
   */
  loadDefaultExample() {
    // 加载默认的Markdown示例
    this.markdownInput.value = `# Markdown转Word示例文档

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
}

// 初始化应用
document.addEventListener('DOMContentLoaded', () => {
  new App();
});

