# DocDraft 产品研究 / Product Research

## 产品定位 / Product Positioning

### 核心价值主张 / Core Value Proposition

**中文:**
DocDraft 是一个**最后一公里**的文档生产工具——专注于将 Markdown 草稿(包括 AI 生成的内容)快速转化为**符合单位格式要求**的正式文档。

**English:**
DocDraft is a **last-mile** document production tool — focused on rapidly transforming Markdown drafts (including AI-generated content) into **formally formatted deliverable documents** that meet organizational standards.

---

### 胜利条件 / Win Condition

**用户能在 3 分钟内,将 AI 生成的 10 页通知文档,在浏览器中本地导出为符合本单位格式规范的 Word 文档。**

**A user can export a 10-page AI-generated notice into a properly formatted Word document matching their organization's style guide — in under 3 minutes, entirely in-browser.**

---

### 一句话描述 / One-Liner Narrative

**"AI 写初稿,DocDraft 排正式版"**  
*"AI drafts, DocDraft delivers."*

---

## 用户痛点 / User Pain Points

| 痛点描述 / Pain Point | 当前影响 / Current Impact | 优先级 / Priority |
|---------------------|-------------------------|------------------|
| AI 生成的 Markdown 很好,但需要转成单位要求的公文/论文格式 | ChatGPT 等工具输出格式简单,无法直接提交 | **P0** |
| 模板应用后预览效果与实际导出差距大,用户不知道最终效果 | 导出前心里没底,需要反复试错 | **P0** |
| 使用模板时,系统未检测是否缺少仿宋/黑体/楷体等字体 | Windows 用户没问题,但 macOS/Linux 用户导出文档打开后发现字体丢失 | **P0** |
| 从 Word 导入模板后,不知道提取了哪些样式,哪些可能丢失 | 用户只看到"导入成功",但不清楚页面设置、正文、标题等具体提取结果 | **P1** |
| 想用 AI 生成文档,但需要手动打开 ChatGPT,粘贴提示词 | 流程繁琐,提示词长,容易遗漏模板信息 | **P1** |
| 导出前没有文档结构检查,缺失H1、标题层级跳跃等问题导出后才发现 | 用户需要手动检查 Markdown,效率低 | **P1** |
| 部分公文/论文需要页眉页脚和页码,当前版本不支持 | 用户导出后需要在 Word 中手动添加 | **P1** |
| 粘贴或引用的图片在导出时可能失败(跨域、data URL 问题) | 用户发现导出的 Word/PDF 中图片丢失 | **P1** |
| 需要服务端转换或云同步功能 | **明确非目标**:保持 Browser-only 离线优先 | **N/A** |

---

## 竞品分析 / Competitor Map

| 竞品 / Competitor | 类型 / Type | 定位 / Positioning | 优势 / Strengths | 劣势 / Weaknesses | DocDraft 差异化 / Our Edge |
|-----------------|-----------|------------------|----------------|----------------|------------------------|
| **Pandoc** | 命令行工具 | 学术界标准转换器 | 格式支持全面,功能强大 | 需要安装配置,学习曲线陡峭,无 GUI | ✅ 浏览器即用,零安装,可视化模板管理 |
| **Typora / Obsidian** | Markdown 编辑器 | 写作环境+导出 | 专业编辑体验,插件丰富 | 重编辑轻导出,中文公文格式支持弱 | ✅ 专注导出最后一公里,内置公文/论文模板 |
| **Dillinger / StackEdit** | 在线 Markdown 编辑器 | 云端写作平台 | 多平台同步,实时协作 | 导出格式简单,无模板系统,需要账号 | ✅ 离线优先,模板驱动,从 Word 导入格式 |
| **md2doc / markdown2word** | 中文公文专项工具 | 针对党政机关公文 | 深度支持 GB/T 9704-2012 | 功能单一,界面老旧,无 PDF 支持 | ✅ 支持多种文档类型,现代 UI,PDF 双输出 |
| **Word 内置 Markdown** | Office 原生功能 | Microsoft 生态 | 无缝集成,企业认可度高 | 样式控制弱,Markdown 语法支持有限 | ✅ 完整 GFM 支持,模板精确控制 |

**竞争优势总结:**
1. **Browser-only**:无需安装,打开即用,离线可用
2. **模板驱动**:从 Word 导入格式 → 应用到 Markdown → 导出 DOCX/PDF
3. **中文优先**:内置公文/论文模板,字体/版式符合国标
4. **AI 友好**:提供 LLM 提示词,快速生成符合格式的初稿

---

## 产品路线图 / Roadmap

### P0 - 必须立即解决 / Must Fix Now

- [x] **中文 PDF 字体支持** (v1.4.0 已完成)
  - 使用 Noto Sans SC,通过 Google Fonts CDN 加载
  - 支持仿宋/黑体/楷体等模板字体映射
- [ ] **系统字体检测与警告**
  - 检测当前操作系统是否安装模板所需字体(仿宋/黑体/楷体等)
  - 如缺失,显示 Toast 提示:"当前系统未安装 XX 字体,导出的 Word 文档可能使用替代字体"
  - 提供字体下载链接或建议
- [ ] **Word 导入反馈优化**
  - `parseDocxStyles` 后显示提取摘要面板:
    - 页面设置:A4 纵向,上下边距 XX mm
    - 正文样式:字体 XX,字号 XX,行距 XX,首行缩进 X 字
    - 标题样式:H1~H5 的字体、字号、对齐方式
  - 标注可能不完整的字段(如未找到页眉页脚)
- [ ] **预览精度提升或免责声明加强**
  - 方案 A:根据选中模板动态调整预览区 CSS,使预览更接近导出效果
  - 方案 B:如预览仍与导出差距大,加强提示:"预览仅供参考,实际导出严格遵循模板格式"
  - 推荐方案 A(首选) + 方案 B(兜底)

### P1 - 高优先级增强 / High Priority Enhancements

- [ ] **AI 工作流优化**
  - 当前:"显示提示词"按钮 → 复制提示词
  - 改进:
    1. 点击"显示提示词"自动复制到剪贴板,显示 Toast:"提示词已复制,打开 AI 工具粘贴使用"
    2. 提供快捷按钮:"在 ChatGPT 中打开" / "在 Claude 中打开"
    3. 如 ChatGPT/Claude 支持 URL 参数预填,尝试拼接提示词(受浏览器 URL 长度限制)
- [ ] **文档结构检查**
  - 导出前自动检查:
    - ⚠️ 文档为空
    - ⚠️ 缺少一级标题(H1)
    - ⚠️ 标题层级跳跃(H1 → H3,跳过 H2)
    - ⚠️ 标题嵌套异常(H3 出现在第一个 H1 之前)
  - 显示非阻塞警告弹窗,用户可选择"仍然导出"或"返回编辑"
- [ ] **页眉页脚与页码支持**
  - 扩展模板配置:
    ```json
    {
      "header": { "text": "单位名称", "alignment": "center" },
      "footer": { "pageNumber": true, "format": "第 X 页 共 Y 页" }
    }
    ```
  - DOCX 使用 `docx.js` 的 Header/Footer API
  - PDF 使用 `pdf-lib` 在每页底部绘制页码
- [ ] **图片嵌入可靠性提升**
  - 当前问题:
    - 跨域图片无法加载
    - Data URL 或 Blob URL 可能失效
  - 改进方案:
    1. 支持 Base64 data URL 直接嵌入
    2. 对于外链图片,显示警告:"外链图片可能因网络问题加载失败,建议使用本地图片或 Base64"
    3. 对于同源图片,尝试转换为 Base64 后嵌入

### P2 - 未来考虑 / Future Considerations

- [ ] 批量转换(多文件上传)
- [ ] 目录(TOC)自动生成
- [ ] Mermaid 图表支持(如实现成本可控)
- [ ] 暗黑模式(仅 UI,不影响导出)
- [ ] 模板市场(社区分享)

---

## 明确非目标 / Explicit Non-Goals

为保持产品定位清晰,以下功能**明确不做**:

❌ **服务端转换**  
→ 坚持 Browser-only,保持离线可用和隐私安全

❌ **云同步或账号系统**  
→ 本地存储足够,避免引入服务器依赖

❌ **完整 Markdown 编辑器**  
→ 不与 Typora/Obsidian 竞争,专注导出环节

❌ **完整 Obsidian 语法支持**  
→ 支持 GFM(GitHub Flavored Markdown)即可,避免生态锁定

❌ **Mermaid 图表**(除非实现成本极低)  
→ 涉及渲染引擎和 SVG 转换,复杂度高

❌ **暗黑模式**(当前阶段)  
→ 不影响核心功能,优先级低

---

## 验收测试 / Acceptance Test

**场景:**  
用户拿到一份 ChatGPT 生成的 10 页会议通知(Markdown 格式),需要导出为符合单位公文格式的 Word 文档。

**测试步骤:**
1. 打开 DocDraft(浏览器直接访问或本地 `dist/index.html`)
2. 粘贴 Markdown 内容到编辑器
3. 选择"党政机关公文"模板
4. 如果系统提示缺少仿宋字体,查看提示信息(不阻塞)
5. 点击"导出 Word"
6. 下载的 `.docx` 文件用 Microsoft Word 或 WPS 打开:
   - ✅ 页面设置:A4 纵向,版心 156×225mm
   - ✅ 正文:仿宋_GB2312 三号(16pt),行距 28pt,首行缩进 2 字
   - ✅ 标题:方正小标宋 22pt 加粗居中
   - ✅ 一级标题:黑体 16pt 加粗左对齐
   - ✅ 表格、列表、引用等元素格式正确
   - ✅ 中文标点、英文单词、数字混排无异常

**通过标准:**
- 从粘贴到下载完成 < 3 分钟
- 导出的文档无需二次调整即可提交
- 如缺少字体,系统有明确提示

---

## 离线/CDN 字体说明 / Offline Font Limitation

**当前实现:**
- PDF 中文字体通过 Google Fonts CDN 加载(Noto Sans SC)
- 首次使用需要联网,后续会话缓存在内存中
- 如 CDN 不可用,自动回退到标准 PDF 字体(Helvetica,字形覆盖有限)

**改进建议:**
1. 在 CHANGELOG 和 README 中诚实说明:"PDF 中文字体首次使用需要联网"
2. 在 UI 中提供"离线模式"开关:
   - 开启后使用标准字体,完全离线
   - 关闭后使用 CDN 字体,首次需要联网
3. 未来可考虑:
   - 将常用字体子集打包进 `dist/`(增加约 2-4MB 体积)
   - 让用户选择是否下载离线字体包

---

## 实施优先级汇总 / Implementation Priority Summary

**本次 PR 必须完成(P0):**
1. ✅ 创建 `docs/PRODUCT_RESEARCH.md`
2. ✅ 系统字体检测与缺失警告
3. ✅ Word 导入反馈优化(显示提取摘要)
4. ✅ 预览精度提升或免责声明加强

**本次 PR 高优先级(P1,如不超预算):**
5. ✅ AI 工作流优化(自动复制 + ChatGPT/Claude 快捷打开)
6. ✅ 文档结构检查(非阻塞警告)
7. ⚠️ 页眉页脚页码支持(评估工作量后决定)
8. ⚠️ 图片嵌入可靠性(评估工作量后决定)

**后续 PR(P2):**
- 批量转换
- TOC 生成
- Mermaid 支持
- 暗黑模式

---

## 链接到主 README / Link to Main README

在主 README 的 Roadmap 部分添加:

```markdown
## Roadmap

For detailed product research, pain points, competitor analysis, and feature priorities, see:  
关于产品定位、痛点分析、竞品研究和功能优先级的详细说明,请参阅:

📄 [docs/PRODUCT_RESEARCH.md](./docs/PRODUCT_RESEARCH.md)
```

---

**文档版本:** v1.0  
**创建日期:** 2026-09-11  
**最后更新:** 2026-09-11
