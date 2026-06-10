// 引入需要的库
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  LineRuleType,
  LevelFormat,
  CheckBox,
  ExternalHyperlink,
  UnderlineType,
  ImageRun,
  TableLayoutType,
} from "docx";
import { saveAs } from "file-saver";
import { Md2Json } from "./md2json.js";
import { marked } from "marked";
import runTest from "./json2docx.js";
import { templateManager } from "./templateManager.js";

/**
 * 简化版的Markdown到Docx转换器
 */
class SimpleMd2Docx {
  constructor() {
    console.log("SimpleMd2Docx 初始化");
    this.md2json = new Md2Json();
    this.docxStyles = null; // set via setTemplate() before converting
  }

  setTemplate(template) {
    this.docxStyles = templateManager.toDocxStyles(template);
  }

  getStyles() {
    return this.docxStyles || templateManager.toDocxStyles(templateManager.getActive());
  }

  getAlignment(value, fallback = AlignmentType.JUSTIFIED) {
    const map = {
      left: AlignmentType.LEFT,
      center: AlignmentType.CENTER,
      right: AlignmentType.RIGHT,
      justified: AlignmentType.JUSTIFIED,
      both: AlignmentType.JUSTIFIED,
      start: AlignmentType.START,
    };
    return map[value] || fallback;
  }

  getTextStyle(kind = "body", overrides = {}) {
    const S = this.getStyles();
    const src = S[kind] || S.body;
    return {
      font: src.font,
      size: src.fontSize,
      color: src.color || "000000",
      bold: src.bold,
      ...overrides,
    };
  }

  getParagraphStyle(kind = "body", overrides = {}) {
    const S = this.getStyles();
    const src = S[kind] || S.body;
    const isBody = kind === "body";
    return {
      alignment: this.getAlignment(src.alignment, isBody ? AlignmentType.JUSTIFIED : AlignmentType.LEFT),
      spacing: {
        before: 0,
        after: 0,
        line: isBody ? S.body.lineSpacing : Math.round(src.fontSize * 1.3 * 10),
        lineRule: LineRuleType.EXACT,
      },
      indent: isBody ? { firstLine: S.body.firstLineIndent } : { left: 0, firstLine: 0 },
      ...overrides,
    };
  }

  getHeadingStyleKey(level) {
    return ["title", "h1", "h2", "h3", "h4", "h5"][Math.max(1, Math.min(level, 6)) - 1];
  }

  getOutputFilename(markdownTokens) {
    const fallback = "document";
    const titleToken = markdownTokens.find((token) => token.type === "heading" && token.depth === 1);
    const rawTitle = titleToken?.text || fallback;
    const filename = this.sanitizeWindowsFilename(rawTitle) || fallback;
    return `${filename}.docx`;
  }

  sanitizeWindowsFilename(value) {
    const reservedNames = /^(con|prn|aux|nul|com[1-9]|lpt[1-9])$/i;
    const cleaned = String(value)
      .replace(/<[^>]*>/g, "")
      .replace(/[`*_~#\[\]()]/g, "")
      .replace(/[<>:"/\\|?*\x00-\x1F]/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .replace(/[. ]+$/g, "");

    if (!cleaned || reservedNames.test(cleaned)) {
      return "";
    }

    return cleaned.slice(0, 120).replace(/[. ]+$/g, "");
  }

  /**
   * 转换Markdown为Word文档
   * @param {string} markdown - Markdown文本
   * @returns {Promise<boolean>} 转换是否成功
   */
  async convertToDocx(markdown) {
    try {
      console.log("开始转换Markdown到Word文档...");

      // 使用Md2Json将markdown转换为json
      const jsonData = this.md2json.convert(markdown);
      console.log("Markdown转换为JSON成功:", jsonData);

      // 使用json2docx.js中的runTest函数生成docx
      await runTest(jsonData);

      return true;
    } catch (error) {
      console.error("转换过程中出错:", error);
      throw error;
    }
  }

  /**
   * 直接使用marked解析结果转换为Word文档
   * @param {string} markdown - Markdown文本
   * @returns {Promise<boolean>} 转换是否成功
   */
  async convertToDocxDirect(markdown) {
    try {
      console.log("开始直接转换Markdown到Word文档...");

      // 保存当前markdown内容
      this.currentMarkdown = markdown;

      // 使用marked解析Markdown，得到token队列
      const markdownArray = marked.lexer(markdown);
      const outputFilename = this.getOutputFilename(markdownArray);
      console.log("Markdown解析结果:", markdownArray);

      // 创建一个数组用于收集所有段落
      let paragraphs = [];

      // 创建一个Promise数组来跟踪所有异步操作
      const asyncOperations = [];

      // 处理markdown的数组
      const processMarkedArray = async (markdownArray) => {
        for (const eachMarkdown of markdownArray) {
          let paragraphPromise;

          // 检查token本身
          switch (eachMarkdown.type) {
            case "heading":
              paragraphPromise = Promise.resolve(
                this.createHeadingFromMarked(eachMarkdown)
              );
              break;
            case "paragraph":
              paragraphPromise = Promise.resolve(
                this.createParagraphFromMarked(eachMarkdown)
              );
              break;
            case "code":
              paragraphPromise = Promise.resolve(
                this.createCodeBlockFromMarked(eachMarkdown)
              );
              break;
            case "list":
              paragraphPromise = this.createListFromMarked(eachMarkdown);
              break;
            case "table":
              paragraphPromise = Promise.resolve(
                this.createTableFromMarked(eachMarkdown)
              );
              break;
            case "blockquote":
              paragraphPromise = Promise.resolve(
                this.createBlockquoteFromMarked(eachMarkdown)
              );
              break;
            case "hr":
              paragraphPromise = Promise.resolve(this.createHorizontalRule());
              break;
            case "html":
              console.log(
                "【convertToDocxDirect】发现HTML token:",
                eachMarkdown
              );
              paragraphPromise = this.processHtmlToken(eachMarkdown);
              break;
            case "space":
              paragraphPromise = Promise.resolve(
                this.createSpaceFromMarked(eachMarkdown)
              );
              break;
            default:
              console.warn(`未支持的token类型: ${eachMarkdown.type}`);
              paragraphPromise = Promise.resolve([]);
              break;
          }

          // 将每个paragraphPromise添加到asyncOperations数组中
          asyncOperations.push(paragraphPromise);
        }
      };

      // 等待所有Markdown数组处理完成
      await processMarkedArray(markdownArray);

      // 等待所有异步操作完成并获取结果
      const results = await Promise.all(asyncOperations);

      // 将所有结果数组展平并过滤掉空值
      paragraphs = results.flat().filter(Boolean);

      console.log("所有Markdown数组处理完成");

      // 创建一个numbering配置
      const numberingConfig = [
        {
          reference: "my-heading-style",
          levels: [
            {
              level: 0, // 对应Markdown的##（depth:2）
              format: LevelFormat.DECIMAL,
              text: "%1.",
              alignment: AlignmentType.START,
              start: 1,
              style: {
                run: {
                  size: 32,
                  font: "黑体",
                  color: "000000",
                  bold: true,
                },
                paragraph: {
                  alignment: AlignmentType.START,
                  indent: {
                    left: 0,
                    firstLine: 0,
                  },
                },
              },
            },
            {
              level: 1, // 对应Markdown的###（depth:3）
              format: LevelFormat.DECIMAL,
              text: "%2.",
              alignment: AlignmentType.START,
              style: {
                paragraph: {
                  indent: {
                    left: 0,
                    firstLine: 0,
                  },
                },
              },
            },
            {
              level: 2, // 对应Markdown的####（depth:4）
              format: LevelFormat.LOWER_LETTER,
              text: "%3)",
              alignment: AlignmentType.START,
              style: {
                paragraph: {
                  indent: {
                    left: 0,
                    firstLine: 0,
                  },
                },
              },
            },
            {
              level: 3, // 对应Markdown的#####（depth:5）
              format: LevelFormat.UPPER_LETTER,
              text: "%4)",
              alignment: AlignmentType.START,
              style: {
                paragraph: {
                  indent: {
                    left: 0,
                    firstLine: 0,
                  },
                },
              },
            },
            {
              level: 4, // 对应Markdown的######（depth:6）
              format: LevelFormat.UPPER_LETTER,
              text: "%5)",
              alignment: AlignmentType.START,
              style: {
                paragraph: {
                  indent: {
                    left: 0,
                    firstLine: 0,
                  },
                },
              },
            },
          ],
        },
        {
          reference: "my-paragraph-style",
          levels: [
            {
              level: 0,
              format: LevelFormat.DECIMAL,
              text: "%1.",
              alignment: AlignmentType.LEFT,
              style: {
                run: {
                  size: 24,
                  font: "仿宋",
                  color: "000000",
                },
                paragraph: {
                  indent: {
                    left: 720, // 36磅左缩进
                    hanging: 360, // 18磅悬挂缩进
                  },
                },
              },
            },
            {
              level: 1,
              format: LevelFormat.DECIMAL,
              text: "%1.%2.",
              alignment: AlignmentType.LEFT,
              style: {
                run: {
                  size: 24,
                  font: "仿宋",
                  color: "000000",
                },
                paragraph: {
                  indent: {
                    left: 1440, // 72磅左缩进
                    hanging: 360, // 18磅悬挂缩进
                  },
                },
              },
            },
          ],
        },
        {
          reference: "my-Unordered-list",
          levels: [
            {
              level: 0,
              format: LevelFormat.BULLET,
              text: "\u25CF", // 实心圆点
              alignment: AlignmentType.LEFT,
              style: {
                paragraph: {
                  indent: {
                    left: 720, // 36磅左缩进
                    hanging: 360, // 18磅悬挂缩进
                  },
                },
              },
            },
            {
              level: 1,
              format: LevelFormat.BULLET,
              text: "\u25CB", // 空心圆点
              alignment: AlignmentType.LEFT,
              style: {
                paragraph: {
                  indent: {
                    left: 1440, // 72磅左缩进
                    hanging: 360, // 18磅悬挂缩进
                  },
                },
              },
            },
          ],
        },
        {
          reference: "my-task-list",
          levels: [
            {
              level: 0,
              format: LevelFormat.BULLET,
              text: "", // 不使用任何标记，因为我们会使用CheckBox
              alignment: AlignmentType.LEFT,
              style: {
                run: {
                  size: 24,
                  font: "仿宋",
                  color: "000000",
                },
                paragraph: {
                  indent: {
                    left: 720, // 36磅左缩进
                    hanging: 360, // 18磅悬挂缩进
                  },
                },
              },
            },
            {
              level: 1,
              format: LevelFormat.BULLET,
              text: "", // 不使用任何标记，因为我们会使用CheckBox
              alignment: AlignmentType.LEFT,
              style: {
                run: {
                  size: 24,
                  font: "仿宋",
                  color: "000000",
                },
                paragraph: {
                  indent: {
                    left: 1440, // 72磅左缩进
                    hanging: 360, // 18磅悬挂缩进
                  },
                },
              },
            },
          ],
        },
      ];
      console.log("开始创建文档...");

      // Resolve styles — fall back to a sensible default if no template was set
      const S = this.getStyles();

      // 创建文档
      const doc = new Document({
        styles: {
          default: {
            document: {
              run: {
                size:  S.body.fontSize,
                font:  S.body.font,
                color: "000000",
              },
              paragraph: {
                alignment: this.getAlignment(S.body.alignment),
                spacing: { before: 0, after: 0, line: S.body.lineSpacing, lineRule: LineRuleType.EXACT },
                indent: { firstLine: S.body.firstLineIndent },
              },
            },
            title: {
              run: {
                size:  S.title.fontSize,
                font:  S.title.font,
                color: S.title.color,
                bold:  S.title.bold,
              },
              paragraph: {
                alignment: this.getAlignment(S.title.alignment, AlignmentType.CENTER),
                spacing: { before: 480, after: 480 },
                lineSpacing: { before: Math.round(S.title.fontSize * 1.3), lineRule: LineRuleType.EXACT },
                indent: { left: 0, firstLine: 0 },
              },
            },
            heading1: {
              run: { size: S.h1.fontSize, font: S.h1.font, color: S.h1.color, bold: S.h1.bold },
              paragraph: {
                alignment: this.getAlignment(S.h1.alignment, AlignmentType.LEFT),
                spacing: { before: 240, after: 120 },
                lineSpacing: { before: Math.round(S.h1.fontSize * 1.4), lineRule: LineRuleType.EXACT },
                indent: { left: 0, firstLine: 0 },
              },
            },
            heading2: {
              run: { size: S.h2.fontSize, font: S.h2.font, color: S.h2.color, bold: S.h2.bold },
              paragraph: {
                alignment: this.getAlignment(S.h2.alignment, AlignmentType.LEFT),
                spacing: { before: 120, after: 60 },
                lineSpacing: { before: Math.round(S.h2.fontSize * 1.35), lineRule: LineRuleType.EXACT },
                indent: { left: 0, firstLine: 0 },
              },
            },
            heading3: {
              run: { size: S.h3.fontSize, font: S.h3.font, color: S.h3.color, bold: S.h3.bold },
              paragraph: {
                alignment: this.getAlignment(S.h3.alignment, AlignmentType.LEFT),
                spacing: { before: 60, after: 60 },
                lineSpacing: { before: Math.round(S.h3.fontSize * 1.3), lineRule: LineRuleType.EXACT },
                indent: { left: 0, firstLine: 0 },
              },
            },
            heading4: {
              run: { size: S.h4.fontSize, font: S.h4.font, color: S.h4.color, bold: S.h4.bold },
              paragraph: {
                alignment: this.getAlignment(S.h4.alignment, AlignmentType.LEFT),
                spacing: { before: 30, after: 30 },
                lineSpacing: { before: Math.round(S.h4.fontSize * 1.3), lineRule: LineRuleType.EXACT },
                indent: { left: 0, firstLine: 0 },
              },
            },
            heading5: {
              run: { size: S.h5.fontSize, font: S.h5.font, color: S.h5.color, bold: S.h5.bold },
              paragraph: {
                alignment: this.getAlignment(S.h5.alignment, AlignmentType.LEFT),
                spacing: { before: 0, after: 0 },
                lineSpacing: { before: Math.round(S.h5.fontSize * 1.3), lineRule: LineRuleType.EXACT },
                indent: { left: 0, firstLine: 0 },
              },
            },
          },
          footnote: {
            run: { size: Math.round(S.body.fontSize * 0.83), font: S.body.font, color: "000000" },
            paragraph: {
              alignment: AlignmentType.JUSTIFIED,
              spacing: { before: 120, after: 120, line: 400 },
              indent: { left: 720, hanging: 360 },
            },
          },
        },
        numbering: {
          config: numberingConfig,
        },
        sections: [
          {
            properties: {
              page: {
                size: { width: S.pageWidth, height: S.pageHeight },
                orientation: S.pageOrientation,
                margin: { ...S.pageMargin, gutter: 0 },
              },
            },
            children: paragraphs,
          },
        ],
      });

      console.log("开始生成文档...");
      // 生成并保存文档
      const blob = await Packer.toBlob(doc);
      saveAs(blob, outputFilename);
      console.log("文档生成完成");

      return true;
    } catch (error) {
      console.error("直接转换过程中出错:", error);
      throw error;
    }
  }

  /**
   * 处理图片
   * @param {string} url - 图片URL
   * @returns {Promise<Object>} 图片数据
   */
  async processImage(url) {
    try {
      console.log("【processImage】开始处理图片:", url);

      // 检查是否是本地图片
      if (
        url.startsWith("./") ||
        url.startsWith("../") ||
        !url.startsWith("http")
      ) {
        console.warn("【processImage】检测到本地图片路径:", url);
        console.warn(
          "【processImage】浏览器环境不支持直接读取本地文件，请将图片转换为base64或使用网络URL"
        );
        throw new Error(
          "浏览器环境不支持直接读取本地文件，请将图片转换为base64或使用网络URL"
        );
      }

      // 1. 获取图片文件
      console.log("【processImage】开始获取图片文件...");
      const response = await fetch(url);
      console.log("【processImage】获取图片响应:", {
        status: response.status,
        statusText: response.statusText,
        contentType: response.headers.get("Content-Type"),
      });

      if (!response.ok) {
        throw new Error(
          `获取图片失败: ${response.status} ${response.statusText}`
        );
      }

      const blob = await response.blob();
      console.log("【processImage】获取到blob:", {
        size: blob.size,
        type: blob.type,
      });

      // 2. 判断图片类型
      let imageType = "";
      const contentType = response.headers.get("Content-Type");
      console.log("【processImage】Content-Type:", contentType);

      if (contentType && contentType.startsWith("image/")) {
        imageType = contentType.split("/")[1].toLowerCase();
        if (imageType === "jpeg") imageType = "jpg";
      } else {
        const ext = url.split(".").pop().toLowerCase();
        if (["png", "jpg", "jpeg", "gif", "webp", "bmp"].includes(ext)) {
          imageType = ext === "jpeg" ? "jpg" : ext;
        } else {
          // 尝试从blob.type获取类型
          const blobType = blob.type.split("/")[1]?.toLowerCase();
          if (
            blobType &&
            ["png", "jpg", "jpeg", "gif", "webp", "bmp"].includes(blobType)
          ) {
            imageType = blobType === "jpeg" ? "jpg" : blobType;
          } else {
            imageType = "png"; // 默认使用png
          }
        }
      }
      console.log("【processImage】最终图片类型:", imageType);

      // 3. 获取图片尺寸
      console.log("【processImage】开始加载图片获取尺寸...");
      const img = new Image();
      await new Promise((resolve, reject) => {
        img.onload = () => {
          console.log("【processImage】图片加载成功，尺寸:", {
            width: img.naturalWidth,
            height: img.naturalHeight,
          });
          resolve();
        };
        img.onerror = (error) => {
          console.error("【processImage】图片加载失败:", error);
          reject(error);
        };
        img.src = URL.createObjectURL(blob);
      });

      // 4. 转换为ArrayBuffer
      console.log("【processImage】开始转换为ArrayBuffer...");
      const buffer = await blob.arrayBuffer();
      console.log("【processImage】ArrayBuffer大小:", buffer.byteLength);

      // 5. 设置图片尺寸（限制最大尺寸）
      const maxWidth = 800;
      const maxHeight = 600;
      let width = img.naturalWidth;
      let height = img.naturalHeight;

      // 调整图片尺寸
      if (width > maxWidth) {
        height = (maxWidth / width) * height;
        width = maxWidth;
      }
      if (height > maxHeight) {
        width = (maxHeight / height) * width;
        height = maxHeight;
      }

      console.log("【processImage】最终图片尺寸:", { width, height });
      return {
        buffer,
        width,
        height,
        type: imageType,
      };
    } catch (error) {
      console.error("【processImage】处理图片失败:", error);
      throw error;
    }
  }

  /**
   * 从marked token创建图片
   * @param {Object} token - marked解析的图片token
   * @returns {Promise<Array>} docx Paragraph对象数组
   */
  async createImageFromMarked(token) {
    try {
      console.log("【createImageFromMarked】开始创建图片段落:", token);

      // 检查是否是本地图片
      if (
        token.href.startsWith("./") ||
        token.href.startsWith("../") ||
        !token.href.startsWith("http")
      ) {
        console.warn("【createImageFromMarked】检测到本地图片:", token.href);
        return [
          new Paragraph({
            children: [
              new TextRun({
                text: `[本地图片: ${token.alt || token.href}]`,
                color: "FF0000",
              }),
              new TextRun({
                text: "\n注意：浏览器环境不支持直接读取本地文件，请将图片转换为base64或使用网络URL",
                color: "FF0000",
                size: 20,
              }),
            ],
            alignment: AlignmentType.CENTER,
          }),
        ];
      }

      // 1. 处理图片并等待完成
      console.log("【createImageFromMarked】开始处理图片...");
      const imageData = await this.processImage(token.href);
      console.log("【createImageFromMarked】图片处理完成:", imageData);

      // 2. 返回创建包含图片的段落
      console.log("【createImageFromMarked】开始创建ImageRun...");
      const imageRun = new ImageRun({
        data: imageData.buffer,
        type: imageData.type,
        transformation: {
          width: imageData.width,
          height: imageData.height,
        },
      });
      console.log("【createImageFromMarked】ImageRun创建完成");

      const paragraph = new Paragraph({
        children: [imageRun],
        alignment: AlignmentType.CENTER, // 段落居中对齐
      });
      console.log("【createImageFromMarked】段落创建完成");

      // 如果有图片说明，添加说明文字
      if (token.title) {
        console.log("【createImageFromMarked】添加图片说明:", token.title);
        paragraph.addChildElement(
          new Paragraph({
            children: [
              new TextRun({
                text: token.title,
                italics: true,
                size: 20,
              }),
            ],
            alignment: AlignmentType.CENTER,
          })
        );
      }

      console.log("【createImageFromMarked】图片段落创建完成");
      return [paragraph];
    } catch (error) {
      console.error("【createImageFromMarked】创建图片段落失败:", error);
      // 如果图片加载失败，返回占位符
      return [
        new Paragraph({
          children: [
            new TextRun({
              text: `[图片: ${token.alt || "加载失败"}]`,
              color: "FF0000",
            }),
            new TextRun({
              text: `\n错误信息: ${error.message}`,
              color: "FF0000",
              size: 20,
            }),
          ],
          alignment: AlignmentType.CENTER,
        }),
      ];
    }
  }

  /**
   * 处理HTML标签
   * @param {Object} token - marked解析的HTML token
   * @returns {Promise<Array>} docx段落数组
   */
  async processHtmlToken(token) {
    console.log("【processHtmlToken】开始处理HTML标签:", token);

    // 检查是否是div标签
    if (token.text.startsWith("<div")) {
      console.log("【processHtmlToken】检测到div标签");
      // 解析align属性
      const alignMatch = token.text.match(/align="([^"]+)"/);
      const alignment = alignMatch ? alignMatch[1] : "left";
      console.log("【processHtmlToken】对齐方式:", alignment);

      // 检查是否包含img标签
      const imgMatch = token.text.match(/<img[^>]+>/);
      if (imgMatch) {
        console.log("【processHtmlToken】检测到img标签:", imgMatch[0]);
        // 提取图片URL
        const srcMatch = imgMatch[0].match(/src="([^"]+)"/);
        const altMatch = imgMatch[0].match(/alt="([^"]+)"/);

        console.log("【processHtmlToken】图片信息:", {
          src: srcMatch ? srcMatch[1] : null,
          alt: altMatch ? altMatch[1] : null,
        });

        if (srcMatch) {
          // 创建图片token
          const imageToken = {
            type: "image",
            href: srcMatch[1],
            alt: altMatch ? altMatch[1] : "",
            title: "",
          };

          console.log("【processHtmlToken】创建图片token:", imageToken);
          // 创建图片段落
          const imageParagraph = await this.createImageFromMarked(imageToken);

          // 设置对齐方式
          switch (alignment.toLowerCase()) {
            case "center":
              imageParagraph[0].alignment = AlignmentType.CENTER;
              break;
            case "right":
              imageParagraph[0].alignment = AlignmentType.RIGHT;
              break;
            case "left":
            default:
              imageParagraph[0].alignment = AlignmentType.LEFT;
              break;
          }

          console.log(
            "【processHtmlToken】图片段落创建完成，对齐方式:",
            alignment
          );
          return imageParagraph;
        }
      }
    }

    console.log("【processHtmlToken】不支持的HTML标签，返回空数组");
    return [];
  }

  /**
   * 创建空行
   * @returns {Array} docx Paragraph对象数组
   */
  createSpaceFromMarked() {
    return [
      new Paragraph({
        text: "",
        alignment: AlignmentType.START,
      }),
    ];
  }

  /**
   * 从marked token创建标题
   * @param {Object} token - marked解析的标题token
   * @returns {Array} docx Paragraph对象数组
   */
  async createHeadingFromMarked(token) {
    const level = token.depth;
    const styleKey = this.getHeadingStyleKey(level);

    // 设置标题样式
    let headingLevel;
    switch (level) {
      case 1:
        headingLevel = HeadingLevel.TITLE;
        break;
      case 2:
        headingLevel = HeadingLevel.HEADING_1;
        break;
      case 3:
        headingLevel = HeadingLevel.HEADING_2;
        break;
      case 4:
        headingLevel = HeadingLevel.HEADING_3;
        break;
      case 5:
        headingLevel = HeadingLevel.HEADING_4;
        break;
      case 6:
        headingLevel = HeadingLevel.HEADING_5;
        break;
      default:
        headingLevel = HeadingLevel.TITLE;
    }

    const headingTextStyle = this.getTextStyle(styleKey);
    const textRuns =
      token.tokens && token.tokens.length > 0
        ? await this.parseTokens(token.tokens, headingTextStyle)
        : this.createTextRunsWithBreaks(token.text, headingTextStyle);

    return [
      new Paragraph({
        heading: headingLevel,
        ...this.getParagraphStyle(styleKey),
        children: textRuns,
      }),
    ];
  }

  /**
   * 从marked token创建段落
   * @param {Object} token - marked解析的段落token
   * @returns {Array} docx Paragraph对象数组
   */
  async createParagraphFromMarked(token) {
    console.log("创建段落:", token);
    const bodyTextStyle = this.getTextStyle("body");
    const sourceTokens = token.tokens && token.tokens.length > 0
      ? token.tokens
      : [{ type: "text", text: token.text || "" }];
    const tokenLines = this.splitInlineTokensIntoLines(sourceTokens);

    if (tokenLines.length > 1) {
      const paragraphs = [];

      for (const lineTokens of tokenLines) {
        const children = lineTokens.length > 0
          ? await this.parseTokens(lineTokens, bodyTextStyle)
          : [];

        paragraphs.push(new Paragraph({
          ...this.getParagraphStyle("body", { alignment: AlignmentType.LEFT }),
          children,
        }));
      }

      return paragraphs;
    }

    const textRuns = tokenLines[0] && tokenLines[0].length > 0
      ? await this.parseTokens(tokenLines[0], bodyTextStyle)
      : [];

    return new Paragraph({
      ...this.getParagraphStyle("body"),
      children: textRuns,
    });
  }

  async parseInlineRunsFromMarked(token, style = {}) {
    if (token?.tokens && token.tokens.length > 0) {
      return this.parseTokens(token.tokens, style);
    }

    if (token?.text) {
      return [new TextRun({ text: token.text, ...style })];
    }

    return [];
  }

  getTextDiagramMetrics(text) {
    const lines = text.split("\n");
    const maxUnits = lines.reduce((max, line) => Math.max(max, this.measureMonospaceUnits(line)), 0);
    const looksLikeDiagram = /[┌┐└┘├┤┬┴┼─│▼▲►◄→←↑↓]/.test(text) || maxUnits > 80;
    const usableTwips = Math.max(
      3600,
      this.getStyles().pageWidth - this.getStyles().pageMargin.left - this.getStyles().pageMargin.right
    );
    const baseIndent = looksLikeDiagram ? 0 : 360;
    const contentTwips = Math.max(2400, usableTwips - baseIndent * 2 - 240);
    const maxFontPt = looksLikeDiagram ? 9 : 10;
    const minFontPt = looksLikeDiagram ? 5.5 : 8;
    // A monospace half-width character is roughly 0.55em in Word, i.e. about
    // 11 twips per point. CJK and box-drawing characters are measured as two
    // half-width units above, so this estimates whether the longest line fits.
    const estimatedPt = maxUnits > 0 ? contentTwips / (maxUnits * 11) : maxFontPt;
    const fontPt = Math.max(minFontPt, Math.min(maxFontPt, estimatedPt));

    return {
      fontSize: Math.round(fontPt * 2),
      lineHeight: Math.round(fontPt * 20 * 1.25),
      indent: baseIndent,
    };
  }

  measureMonospaceUnits(text) {
    let units = 0;
    for (const char of text) {
      units += /[^\x00-\xff]/.test(char) ? 2 : 1;
    }
    return units;
  }

  /**
   * Create a DOCX paragraph from a marked code-block token.
   */
  createCodeBlockFromMarked(token) {
    const textRuns = [];
    const metrics = this.getTextDiagramMetrics(token.text);
    const codeFont = "Consolas";

    // 如果有语言标识，添加语言标签
    if (token.lang) {
      textRuns.push(
        new TextRun({
          text: `[${token.lang}]`,
          size: metrics.fontSize,
          font: codeFont,
          color: "666666", // 灰色
          bold: true,
        })
      );
    }

    // 处理代码文本，按换行符切分
    const lines = token.text.split("\n");
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      // 添加代码行
      textRuns.push(
        new TextRun({
          text: line,
          size: metrics.fontSize,
          font: codeFont,
          color: "000000", // 黑色
          // 如果有语言标识，第一行需要break；如果没有语言标识，第一行不需要break
          break: (token.lang && i === 0) || i > 0 ? 1 : 0,
        })
      );
    }

    return [
      new Paragraph({
        style: "Code Block",
        alignment: AlignmentType.LEFT,
        spacing: {
          before: 120,
          after: 120,
          line: metrics.lineHeight,
          lineRule: LineRuleType.EXACT,
        },
        indent: {
          left: metrics.indent,
          right: metrics.indent,
          firstLine: 0, // 去掉首行缩进
        },
        border: {
          top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
          bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
          left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
          right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
        },
        shading: {
          fill: "F5F5F5", // 浅灰色背景
        },
        children: textRuns,
      }),
    ];
  }

  /**
   * 从marked token创建列表
   * @param {Object} token - marked解析的列表token
   * @returns {Array} docx Paragraph对象数组
   */
  async createListFromMarked(token, level = 0) {
    console.log("开始创建列表:", token);
    //创建一个数组用于收集所有列表项
    let listItems = [];
    //判断list类型是有序列表还是无序列表还是任务列表
    //ordered是真是有序列表，假就可能是无序或任务列表
    if (token.ordered) {
      //遍历token.items
      for (const item of token.items) {
        console.log("item:", item);
        //判断task的值是否为true，是就是任务列表，否则就是有序列表
        if (item.task) {
          console.log("识别到任务列表");
          //遍历item.tokens
          for (const token of item.tokens) {
            if (token.type === "text") {
              listItems.push(
                new CheckBox({
                  checked: item.task,
                  name: item.text,
                })
              );
            } else if (token.type === "list") {
              let tempArray = await this.createListFromMarked(token, level + 1);
              console.log("tempArray:", tempArray);
              listItems.push(...tempArray);
            }
          }
          listItems.push(
            new Paragraph({
              ...this.getParagraphStyle("body", { indent: { left: 720, hanging: 360 } }),
              children: listItems,
            })
          );
        } else {
          console.log("识别到有序列表");
          //遍历item.tokens
          for (const token of item.tokens) {
            if (token.type === "text") {
              const children = await this.parseInlineRunsFromMarked(token, this.getTextStyle("body"));
              listItems.push(
                new Paragraph({
                  children,
                  ...this.getParagraphStyle("body", { indent: { left: 720 * (level + 1), hanging: 360 } }),
                  numbering: {
                    reference: "my-paragraph-style",
                    level: level,
                  },
                })
              );
            } else if (token.type === "list") {
              let tempArray = await this.createListFromMarked(token, level + 1);
              console.log("tempArray:", tempArray);
              listItems.push(...tempArray);
            }
          }
        }
      }
    } else {
      console.log("无序列表或任务列表");
      //遍历token.items
      for (const item of token.items) {
        console.log("item:", item);
        //判断task的值是否为true，是就是任务列表，否则就是无序列表
        if (item.task) {
          console.log("识别到任务列表");
          //遍历item.tokens
          for (const token of item.tokens) {
            if (token.type === "text") {
              //创建一个ICheckboxSymbolOptions
              const  ICheckboxSymbolOptions={
                alias:item.text,
                checked:item.checked,
              }

              listItems.push(new Paragraph({
                ...this.getParagraphStyle("body", { indent: { left: 720 * (level + 1), hanging: 360 } }),
                children:[
                  new CheckBox(ICheckboxSymbolOptions),
                  ...(await this.parseInlineRunsFromMarked(token, this.getTextStyle("body"))),
                ],
                numbering: {
                  reference: "my-task-list",
                  level: level,
                },
              }));
            }else if (token.type === "list") {
              let tempArray = await this.createListFromMarked(token, level + 1);
              console.log("tempArray:", tempArray);
              listItems.push(...tempArray);
            }
          }
        }else{
          console.log("识别到无序列表");
          //遍历item.tokens
          for (const token of item.tokens) {
            if (token.type === "text") {
              const children = await this.parseInlineRunsFromMarked(token, this.getTextStyle("body"));
              listItems.push(new Paragraph({
                children,
                ...this.getParagraphStyle("body", { indent: { left: 720 * (level + 1), hanging: 360 } }),
                numbering: {
                  reference: "my-Unordered-list",
                  level: level,
                },
              }));
            }
          }
        }
      }
    }

    return listItems;
  }

  /**
   * 从marked token创建表格
   * @param {Object} token - marked解析的表格token
   * @returns {Array} docx Table对象数组
   */
  async createTableFromMarked(token) {
    console.log("创建表格:", token);
    const rows = [];

    // 解析列对齐方式（从token.align获取）
    const alignments = token.align || [];

    // 处理表头
    if (token.header) {
      const headerCells = await Promise.all(
        token.header.map(async (cell, index) => {
          const children = cell.tokens
            ? await this.parseTokens(cell.tokens, this.getTextStyle("body", { bold: true }))
            : [new TextRun({ text: cell.text, ...this.getTextStyle("body", { bold: true }) })];

          // 获取当前列的对齐方式
          const alignment = alignments[index]
            ? this.getAlignmentType(alignments[index])
            : AlignmentType.LEFT;

          return new TableCell({
            children: [
              new Paragraph({
                indent: { firstLine: 0 },
                alignment: alignment, // 动态设置对齐
                children: children,
              })
            ],
            shading: {
              fill: "F0F0F0",
            },
          });
        })
      );

      rows.push(new TableRow({ children: headerCells }));
    }

    // 处理表格内容
    for (const row of token.rows) {
      const cells = await Promise.all(
        row.map(async (cell, index) => {
          const children = cell.tokens
            ? await this.parseTokens(cell.tokens, this.getTextStyle("body"))
            : [new TextRun({ text: cell.text, ...this.getTextStyle("body") })];

          // 获取当前列的对齐方式
          const alignment = alignments[index]
            ? this.getAlignmentType(alignments[index])
            : AlignmentType.LEFT;

          return new TableCell({
            children: [
              new Paragraph({
                indent: { firstLine: 0 },
                alignment: alignment, // 动态设置对齐
                children: children,
              })
            ],
          });
        })
      );

      rows.push(new TableRow({ children: cells }));
    }

    return new Table({
      rows,
      alignment: AlignmentType.CENTER,
      width: {
        size: 100,
        type: WidthType.PERCENTAGE,
      },
      layout: TableLayoutType.FIXED,
      columnWidths: token.header ? token.header.map(() => 3000) : [],
      margins: {
        top: 100,
        bottom: 100,
        left: 100,
        right: 100,
      },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 2, color: "000000" },
        bottom: { style: BorderStyle.SINGLE, size: 2, color: "000000" },
        left: { style: BorderStyle.SINGLE, size: 2, color: "000000" },
        right: { style: BorderStyle.SINGLE, size: 2, color: "000000" },
        insideHorizontal: {
          style: BorderStyle.SINGLE,
          size: 1,
          color: "CCCCCC",
        },
        insideVertical: {
          style: BorderStyle.SINGLE,
          size: 1,
          color: "CCCCCC",
        },
      },
    });
  }

  /**
   * 获取对齐方式
   * @param {string} align - 对齐方式
   * @returns {AlignmentType} docx对齐方式
   */
  getAlignmentType(align) {
    switch (align) {
      case "left":
      case ":--":      // GFM 左对齐
        return AlignmentType.LEFT;
      case "center":
      case ":-:":      // GFM 居中对齐
        return AlignmentType.CENTER;
      case "right":
      case "--:":      // GFM 右对齐
        return AlignmentType.RIGHT;
      default:
        return AlignmentType.LEFT;
    }
  }

  /**
   * 从marked token创建引用块
   * @param {Object} token - marked解析的引用块token
   * @returns {Array} docx Paragraph对象数组
   */
  createBlockquoteFromMarked(token) {
    return [
      new Paragraph({
        style: "Intense Quote",
        alignment: AlignmentType.CENTER,
        spacing: {
          before: 360,
          after: 360,
        },
        indent: {
          left: 864,
          right: 864,
        },
        border: {
          top: { style: BorderStyle.SINGLE, size: 4, color: "0F4761" },
          bottom: { style: BorderStyle.SINGLE, size: 4, color: "0F4761" },
        },
        children: [
          new TextRun({
            text: token.text,
            size: this.docxStyles ? this.docxStyles.body.fontSize : 24,
            font: this.docxStyles ? this.docxStyles.body.font : "仿宋_GB2312",
            color: "0F4761",
            italics: true,
          }),
        ],
      }),
    ];
  }

  /**
   * 创建水平线
   * @returns {Array} docx Paragraph对象数组
   */
  createHorizontalRule() {
    return [
      new Paragraph({
        spacing: {
          before: 200, // 段前10磅
          after: 200, // 段后10磅
        },
        border: {
          bottom: {
            color: "#CCCCCC", // 灰色
            space: 1, // 间距
            style: BorderStyle.SINGLE, // 单线
            size: 6, // 线宽
          },
        },
      }),
    ];
  }

  /**
   * 处理文档中的所有图片
   * @param {Array} tokens - marked解析的tokens
   * @returns {Promise<Array>} 处理后的图片信息数组
   */
  async processImages(tokens) {
    const imageInfos = [];
    for (const token of tokens) {
      if (token.type === "image") {
        try {
          const imageData = await this.loadImage(token.href);
          imageInfos.push({
            ...token,
            ...imageData,
          });
        } catch (error) {
          console.error(`处理图片失败: ${token.href}`, error);
        }
      }
    }
    return imageInfos;
  }

  /**
   * 加载图片
   * @param {string} url - 图片URL
   * @returns {Promise<Object>} 图片数据
   */
  async loadImage(url) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.crossOrigin = "anonymous";

      img.onload = async () => {
        try {
          const canvas = document.createElement("canvas");
          canvas.width = img.naturalWidth;
          canvas.height = img.naturalHeight;

          const ctx = canvas.getContext("2d");
          ctx.drawImage(img, 0, 0);

          const blob = await new Promise((resolve) => canvas.toBlob(resolve));
          const buffer = await blob.arrayBuffer();

          resolve({
            buffer,
            width: img.naturalWidth,
            height: img.naturalHeight,
          });
        } catch (error) {
          reject(error);
        }
      };

      img.onerror = () => {
        reject(new Error(`加载图片失败: ${url}`));
      };

      img.src = url;
    });
  }

  createTextRunsWithBreaks(text = "", style = {}) {
    const lines = String(text).split(/\r?\n/);
    const runs = [];

    lines.forEach((line, index) => {
      if (index > 0) {
        runs.push(new TextRun({ break: 1 }));
      }

      if (line.length > 0) {
        runs.push(new TextRun({ text: line, ...style }));
      }
    });

    return runs.length > 0 ? runs : [new TextRun({ text: "", ...style })];
  }

  splitInlineTokensIntoLines(tokens = []) {
    const lines = [[]];

    const pushToken = (token) => {
      if (!token) {
        return;
      }

      if (token.type === "br") {
        lines.push([]);
        return;
      }

      if (typeof token.text === "string" && /\r?\n/.test(token.text) && !token.tokens) {
        const parts = token.text.split(/\r?\n/);
        parts.forEach((part, index) => {
          if (index > 0) {
            lines.push([]);
          }

          if (part.length > 0) {
            lines[lines.length - 1].push({ ...token, text: part, raw: part });
          }
        });
        return;
      }

      lines[lines.length - 1].push(token);
    };

    tokens.forEach(pushToken);

    while (lines.length > 1 && lines[lines.length - 1].length === 0) {
      lines.pop();
    }

    return lines;
  }

  /**
   * 解析tokens为docx元素
   * @param {Array} tokens - marked解析后的tokens
   * @param {Object} style - 样式对象
   * @returns {Array} docx元素数组
   */
  async parseTokens(tokens, style = {}) {
    const self = this; // 保存this引用
    const ElementHandlers = {
      // 文本类型（可嵌套）
      text: {
        async process(token, style, parseChildren) {
          if (token.tokens) {
            const children = await parseChildren(token.tokens, style);
            return children.flat();
          }
          return self.createTextRunsWithBreaks(token.text, style);
        },
      },

      br: {
        async process() {
          return [new TextRun({ break: 1 })];
        },
      },

      // 粗体类型
      strong: {
        async process(token, style, parseChildren) {
          const children = await parseChildren(token.tokens, {
            ...style,
            bold: true,
          });
          return children.flat();
        },
      },

      // 斜体类型
      em: {
        async process(token, style, parseChildren) {
          const children = await parseChildren(token.tokens, {
            ...style,
            italics: true,
          });
          return children.flat();
        },
      },

      // 删除线
      del: {
        async process(token, style, parseChildren) {
          const children = await parseChildren(token.tokens, {
            ...style,
            strike: true,
          });
          return children.flat();
        },
      },

      // 超链接
      link: {
        async process(token, style, parseChildren) {
          const linkStyle = {
            ...style,
            color: "0563C1",
            underline: { type: UnderlineType.SINGLE, color: "0563C1" },
          };

          const children = await parseChildren(token.tokens, linkStyle);

          return [
            new ExternalHyperlink({
              children: children.flat(),
              link: token.href,
              tooltip: token.title || token.href,
            }),
          ];
        },
      },

      // 代码块
      codespan: {
        async process(token, style) {
          return [
            new TextRun({
              text: token.text,
              font: "等线",
              size: 24,
              color: "0000FF",
              bold: true,
              ...style, // 允许外部样式覆盖
            }),
          ];
        },
      },

      // 图片处理
      image: {
        async process(token, style, parseChildren) {
          try {
            // 检查是否是本地图片
            if (
              token.href.startsWith("./") ||
              token.href.startsWith("../") ||
              !token.href.startsWith("http")
            ) {
              console.warn("【image.process】检测到本地图片:", token.href);
              // 创建一个1x1像素的透明PNG图片作为占位符
              const placeholderImage = new Uint8Array([
                0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00,
                0x00, 0x0d, 0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01,
                0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f,
                0x15, 0xc4, 0x89, 0x00, 0x00, 0x00, 0x0a, 0x49, 0x44, 0x41,
                0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00,
                0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49,
                0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
              ]);

              return [
                new ImageRun({
                  data: placeholderImage,
                  type: "png",
                  transformation: {
                    width: 1,
                    height: 1,
                  },
                }),
                new TextRun({
                  text: `[本地图片: ${token.text || token.href}]`,
                  color: "FF0000",
                  bold: true,
                }),
                new TextRun({
                  text: "\n注意：浏览器环境不支持直接读取本地文件，请将图片转换为base64或使用网络URL",
                  color: "FF0000",
                  size: 20,
                }),
              ];
            }

            // 解析alt文本中的markdown
            const altTokens =
              marked.lexer(token.text || "image")[0]?.tokens || [];
            const altRuns = await parseChildren(altTokens, {
              ...style,
              size: style.size ? style.size - 2 : 20,
            });

            // 获取图片数据
            const imageData = await self.processImage(token.href);
            console.log("【image.process】图片数据:", {
              type: imageData.type,
              width: imageData.width,
              height: imageData.height,
              bufferSize: imageData.buffer.byteLength,
            });

            // 确保图片类型是有效的
            const validTypes = ["png", "jpg", "jpeg", "gif", "bmp", "webp"];
            const imageType = validTypes.includes(imageData.type)
              ? imageData.type
              : "png";

            return [
              new ImageRun({
                data: imageData.buffer,
                type: imageType,
                transformation: {
                  width: imageData.width,
                  height: imageData.height,
                },
              }),
            ];
          } catch (error) {
            console.error("处理图片失败:", error);
            // 创建一个1x1像素的透明PNG图片作为占位符
            const placeholderImage = new Uint8Array([
              0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00,
              0x0d, 0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00,
              0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4, 0x89,
              0x00, 0x00, 0x00, 0x0a, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63,
              0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4,
              0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60,
              0x82,
            ]);

            return [
              new ImageRun({
                data: placeholderImage,
                type: "png",
                transformation: {
                  width: 1,
                  height: 1,
                },
              }),
              new TextRun({
                text: `[图片加载失败: ${token.href}]`,
                color: "FF0000",
                bold: true,
              }),
              new TextRun({
                text: `\n错误信息: ${error.message}`,
                color: "FF0000",
                size: 20,
              }),
            ];
          }
        },
      },
    };

    const result = [];

    for (const token of tokens) {
      try {
        const handler = ElementHandlers[token.type] || ElementHandlers.text;

        const parseChildren = async (childrenTokens, childStyle) =>
          childrenTokens ? this.parseTokens(childrenTokens, childStyle) : [];

        const elements = await handler.process(
          token,
          { ...style }, // 克隆样式对象
          parseChildren
        );

        result.push(...elements.flat());
      } catch (error) {
        console.error(`处理 ${token.type} 失败:`, error);
        result.push(
          new TextRun({
            text: `[解析错误: ${token.type}]`,
            color: "FF0000",
          })
        );
      }
    }

    return result;
  }

  /**
   * 处理单个元素
   * @param {Object} element - 元素对象
   * @returns {Object|Array} docx文档元素或元素数组
   */
  processElement(element) {
    if (!element) return null;

    switch (element.type) {
      case "heading":
        return this.createHeading(element);

      case "paragraph":
        return this.createParagraph(element);

      case "code_block":
        return this.createCodeBlock(element);

      case "list":
        return this.createList(element);

      case "table":
        return this.createTable(element);

      case "horizontal_rule":
        return this.createHorizontalRule();

      case "blockquote":
        return this.createBlockquote(element);

      default:
        console.warn(`未支持的元素类型: ${element.type}`);
        return new Paragraph({
          children: [
            new TextRun({
              text: `[未支持的元素: ${element.type}]`,
              color: "999999",
            }),
          ],
        });
    }
  }

  /**
   * 创建标题
   * @param {Object} element - 标题元素
   * @returns {Object} docx Paragraph对象
   */
  createHeading(element) {
    console.log("创建标题:", element);

    const level = element.level || 1;
    let headingLevel;

    // 将Markdown的level转换为docx的HeadingLevel
    switch (level) {
      case 1:
        headingLevel = HeadingLevel.HEADING_1;
        break;
      case 2:
        headingLevel = HeadingLevel.HEADING_2;
        break;
      case 3:
        headingLevel = HeadingLevel.HEADING_3;
        break;
      case 4:
        headingLevel = HeadingLevel.HEADING_4;
        break;
      case 5:
        headingLevel = HeadingLevel.HEADING_5;
        break;
      case 6:
        headingLevel = HeadingLevel.HEADING_6;
        break;
      default:
        headingLevel = HeadingLevel.HEADING_1;
    }

    // 处理内联样式
    if (
      element.inlineStyles &&
      Array.isArray(element.inlineStyles) &&
      element.inlineStyles.length > 0
    ) {
      const children = [];

      element.inlineStyles.forEach((style) => {
        if (style.content) {
          children.push(
            new TextRun({
              text: style.content,
              bold: style.bold === true,
              italic: style.italic === true,
              strike: style.strike === true,
              underline: style.underline === true ? {} : undefined,
              superScript: style.superscript === true,
              subScript: style.subscript === true,
            })
          );
        }
      });

      if (children.length > 0) {
        return new Paragraph({
          heading: headingLevel,
          children,
        });
      }
    }

    // 提取标题文本（如果没有可用的内联样式）
    const titleText =
      element.fullContent || element.rawText || element.text || "";
    // 去掉Markdown的#符号
    const cleanTitle = titleText.replace(/^#+\s+/, "");

    console.log(`创建标题: '${cleanTitle}', 级别: ${level}`);

    // 创建标题段落
    return new Paragraph({
      text: cleanTitle,
      heading: headingLevel,
    });
  }

  /**
   * 创建段落
   * @param {Object} element - 段落元素
   * @returns {Object} docx Paragraph对象
   */
  createParagraph(element) {
    console.log("创建段落:", element);

    // 如果有内联样式
    if (
      element.inlineStyles &&
      Array.isArray(element.inlineStyles) &&
      element.inlineStyles.length > 0
    ) {
      const children = [];

      element.inlineStyles.forEach((style) => {
        if (style.content) {
          children.push(
            new TextRun({
              text: style.content,
              bold: style.bold === true,
              italic: style.italic === true,
              strike: style.strike === true,
              underline: style.underline === true ? {} : undefined,
              superScript: style.superscript === true,
              subScript: style.subscript === true,
            })
          );
        }
      });

      if (children.length > 0) {
        return new Paragraph({ children });
      }
    }

    // 如果没有内联样式但有fullContent
    if (element.fullContent) {
      return new Paragraph({ text: element.fullContent });
    }

    // 简单段落，使用任何可用的文本属性
    const text = element.rawText || element.text || "";
    return new Paragraph({ text });
  }

  /**
   * 创建代码块
   * @param {Object} element - 代码块元素
   * @returns {Object} docx Paragraph对象
   */
  createCodeBlock(element) {
    console.log("创建代码块:", element);

    const code = element.fullContent || element.content || element.text || "";
    const language = element.language || "";

    return new Paragraph({
      children: [
        new TextRun({
          text: `${language ? `[${language}] ` : ""}${code}`,
          font: "Courier New",
        }),
      ],
      indent: {
        left: 720, // 缩进量，720 = 0.5英寸
      },
    });
  }

  /**
   * 创建列表
   * @param {Object} element - 列表元素
   * @returns {Array} Paragraph对象数组
   */
  createList(element) {
    console.log("创建列表:", element);

    const items = element.items || [];
    const isOrdered = element.ordered === true;
    const result = [];

    // 对每个列表项创建一个段落
    items.forEach((item, index) => {
      const prefix = isOrdered ? `${index + 1}. ` : "• ";
      const itemText = item.fullContent || item.text || "";

      result.push(
        new Paragraph({
          text: `${prefix}${itemText}`,
          indent: {
            left: 720, // 缩进量
          },
        })
      );

      // 处理嵌套列表
      if (
        item.children &&
        Array.isArray(item.children) &&
        item.children.length > 0
      ) {
        item.children.forEach((child) => {
          const nestedItems = this.processElement(child);
          if (Array.isArray(nestedItems)) {
            result.push(...nestedItems);
          } else if (nestedItems) {
            result.push(nestedItems);
          }
        });
      }
    });

    return result;
  }

  /**
   * 创建表格
   * @param {Object} element - 表格元素
   * @returns {Object} docx Table对象
   */
  createTable(element) {
    console.log("创建表格:", element);

    // 检查表格数据
    if (
      !element.data ||
      !Array.isArray(element.data) ||
      element.data.length === 0
    ) {
      return new Paragraph({ text: "[空表格]" });
    }

    // 创建表格行
    const rows = [];
    element.data.forEach((rowData, rowIndex) => {
      if (!Array.isArray(rowData)) {
        console.warn(`表格行${rowIndex}不是数组`);
        return;
      }

      const cells = [];
      rowData.forEach((cellData, cellIndex) => {
        const cellText = cellData.fullContent || cellData.text || "";

        cells.push(
          new TableCell({
            children: [new Paragraph({ text: cellText })],
          })
        );
      });

      if (cells.length > 0) {
        rows.push(new TableRow({ children: cells }));
      }
    });

    if (rows.length === 0) {
      return new Paragraph({ text: "[空表格]" });
    }

    // 创建表格
    return new Table({
      rows: rows,
      width: {
        size: 100,
        type: WidthType.PERCENTAGE,
      },
    });
  }

  /**
   * 创建图片占位符
   * @param {Object} element - 图片元素
   * @returns {Object} docx Paragraph对象
   */
  createImagePlaceholder(element) {
    console.log("创建图片占位符:", element);

    const altText = element.alt || "[图片]";
    const url = element.url || "";

    return new Paragraph({
      children: [
        new TextRun({
          text: `[图片: ${altText}]${url ? ` ${url}` : ""}`,
          italic: true,
          color: "0000FF",
        }),
      ],
      alignment: AlignmentType.CENTER,
    });
  }

  /**
   * 创建引用块
   * @param {Object} element - 引用块元素
   * @returns {Object} docx Paragraph对象
   */
  createBlockquote(element) {
    console.log("创建引用块:", element);

    const text = element.fullContent || element.text || "";

    return new Paragraph({
      text: text,
      indent: {
        left: 720, // 缩进量
      },
      border: {
        left: {
          color: "CCCCCC",
          size: 4,
          style: BorderStyle.SINGLE,
        },
      },
      spacing: {
        before: 240, // 前间距
        after: 240, // 后间距
      },
    });
  }

  /**
   * 创建错误文档
   * @param {string} errorMessage - 错误信息
   * @returns {Promise<Blob>} 错误文档的Blob对象
   */
  async createErrorDocument(errorMessage) {
    try {
      const doc = new Document({
        sections: [
          {
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "文档生成失败",
                    bold: true,
                    color: "FF0000",
                    size: 36, // 18pt
                  }),
                ],
                alignment: AlignmentType.CENTER,
              }),
              new Paragraph({
                text: errorMessage,
                alignment: AlignmentType.CENTER,
              }),
              new Paragraph({
                text: `生成时间: ${new Date().toLocaleString()}`,
                alignment: AlignmentType.CENTER,
              }),
            ],
          },
        ],
      });

      return await Packer.toBlob(doc);
    } catch (error) {
      console.error("创建错误文档时出错:", error);

      // 创建最小的错误文档
      const minimalDoc = new Document({
        sections: [
          {
            children: [
              new Paragraph({
                text: `错误: ${errorMessage}`,
              }),
            ],
          },
        ],
      });

      return await Packer.toBlob(minimalDoc);
    }
  }

  /**
   * 保存为docx文件
   * @param {Blob} blob - docx文件的Blob对象
   * @param {string} filename - 文件名
   */
  saveAsDocx(blob, filename = "document.docx") {
    try {
      saveAs(blob, filename);
      console.log(`文档已保存为 ${filename}`);
    } catch (error) {
      console.error("保存文档时出错:", error);
      alert(`保存文档失败: ${error.message}`);
    }
  }
}

export default SimpleMd2Docx;
