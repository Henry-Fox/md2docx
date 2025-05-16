/**
 * 样式编辑器类
 */
class StyleEditor {
    constructor() {
        this.currentCategory = null;
        this.currentStyles = null;
        this.defaultStyles = null;
        this.init();
    }

    /**
     * 初始化编辑器
     */
    async init() {
        try {
            // 加载默认样式
            await this.loadDefaultStyles();

            // 初始化导航
            this.initNavigation();

            // 初始化工具栏
            this.initToolbar();

            // 初始化设置面板
            this.initSettingsPanels();

            // 初始化预览
            this.initPreview();

            // 默认显示文档设置
            this.switchCategory('document');
        } catch (error) {
            console.error('初始化失败:', error);
            alert('初始化失败，请刷新页面重试');
        }
    }

    /**
     * 加载默认样式
     */
    async loadDefaultStyles() {
        try {
            const response = await fetch('../default-styles.json');
            this.defaultStyles = await response.json();
            this.currentStyles = JSON.parse(JSON.stringify(this.defaultStyles));
        } catch (error) {
            console.error('加载默认样式失败:', error);
            throw error;
        }
    }

    /**
     * 初始化导航
     */
    initNavigation() {
        const navItems = document.querySelectorAll('.nav-item');
        navItems.forEach(item => {
            item.addEventListener('click', () => {
                const category = item.dataset.category;
                this.switchCategory(category);

                // 更新导航项状态
                navItems.forEach(navItem => navItem.classList.remove('active'));
                item.classList.add('active');
            });
        });
    }

    /**
     * 初始化工具栏
     */
    initToolbar() {
        // 保存按钮
        document.getElementById('saveBtn').addEventListener('click', () => this.saveStyles());

        // 重置按钮
        document.getElementById('resetBtn').addEventListener('click', () => this.resetStyles());

        // 导入按钮
        document.getElementById('importBtn').addEventListener('click', () => this.importStyles());

        // 导出按钮
        document.getElementById('exportBtn').addEventListener('click', () => this.exportStyles());
    }

    /**
     * 初始化设置面板
     */
    initSettingsPanels() {
        // 文档设置面板
        this.initDocumentSettings();

        // 标题设置面板
        this.initHeadingSettings();

        // 其他设置面板...
    }

    /**
     * 初始化文档设置面板
     */
    initDocumentSettings() {
        const panel = document.getElementById('documentSettings');
        const settings = this.currentStyles.document;

        // 页面大小
        this.createSelect(panel, '页面大小', 'document.pageSize', settings.pageSize, [
            { value: 'A4', text: 'A4' },
            { value: 'A3', text: 'A3' },
            { value: 'B5', text: 'B5' }
        ]);

        // 页面方向
        this.createSelect(panel, '页面方向', 'document.pageOrientation', settings.pageOrientation, [
            { value: 'portrait', text: '纵向' },
            { value: 'landscape', text: '横向' }
        ]);

        // 页边距
        this.createNumberInput(panel, '上边距(磅)', 'document.margins.top', settings.margins.top);
        this.createNumberInput(panel, '下边距(磅)', 'document.margins.bottom', settings.margins.bottom);
        this.createNumberInput(panel, '左边距(磅)', 'document.margins.left', settings.margins.left);
        this.createNumberInput(panel, '右边距(磅)', 'document.margins.right', settings.margins.right);

        // 网格设置
        this.createNumberInput(panel, '每行字符数', 'document.grid.charPerLine', settings.grid.charPerLine);
        this.createNumberInput(panel, '每页行数', 'document.grid.linePerPage', settings.grid.linePerPage);
    }

    /**
     * 初始化标题设置面板
     */
    initHeadingSettings() {
        const panel = document.getElementById('headingSettings');
        const settings = this.currentStyles.heading.styles;

        // 为每个标题级别创建设置
        for (let i = 1; i <= 6; i++) {
            const headingSettings = settings[`h${i}`];

            // 创建标题级别分组
            const group = document.createElement('div');
            group.className = 'setting-group';
            group.innerHTML = `<h3>${i}级标题样式</h3>`;
            panel.appendChild(group);

            // 字体设置
            this.createTextInput(group, '字体名称', `heading.styles.h${i}.font.name`, headingSettings.font.name);
            this.createTextInput(group, '字体回退', `heading.styles.h${i}.font.fallback`, headingSettings.font.fallback.join(', '));
            this.createNumberInput(group, '字号(磅)', `heading.styles.h${i}.font.size`, headingSettings.font.size);
            this.createSwitch(group, '加粗', `heading.styles.h${i}.font.bold`, headingSettings.font.bold);

            // 段落设置
            this.createSelect(group, '对齐方式', `heading.styles.h${i}.paragraph.alignment`, headingSettings.paragraph.alignment, [
                { value: 'left', text: '左对齐' },
                { value: 'center', text: '居中' },
                { value: 'right', text: '右对齐' },
                { value: 'justified', text: '两端对齐' }
            ]);

            this.createNumberInput(group, '段前间距(磅)', `heading.styles.h${i}.paragraph.spacing.before`, headingSettings.paragraph.spacing.before);
            this.createNumberInput(group, '段后间距(磅)', `heading.styles.h${i}.paragraph.spacing.after`, headingSettings.paragraph.spacing.after);
            this.createNumberInput(group, '左缩进(磅)', `heading.styles.h${i}.paragraph.indent.left`, headingSettings.paragraph.indent.left);
            this.createNumberInput(group, '首行缩进(磅)', `heading.styles.h${i}.paragraph.indent.firstLine`, headingSettings.paragraph.indent.firstLine);

            // 编号设置
            this.createSwitch(group, '使用前缀', `heading.styles.h${i}.numbering.usePrefix`, headingSettings.numbering.usePrefix);
            this.createTextInput(group, '前缀文本', `heading.styles.h${i}.numbering.prefix`, headingSettings.numbering.prefix);
            this.createTextInput(group, '后缀文本', `heading.styles.h${i}.numbering.suffix`, headingSettings.numbering.suffix);
            this.createNumberInput(group, '编号级别', `heading.styles.h${i}.numbering.level`, headingSettings.numbering.level);
            this.createTextInput(group, '编号模板', `heading.styles.h${i}.numbering.template`, headingSettings.numbering.template);
        }
    }

    /**
     * 创建文本输入框
     */
    createTextInput(container, label, path, value) {
        const div = document.createElement('div');
        div.className = 'setting-item';

        const labelElement = document.createElement('label');
        labelElement.className = 'setting-label';
        labelElement.textContent = label;

        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'setting-input';
        input.value = value;

        input.addEventListener('change', () => {
            this.updateStyle(path, input.value);
        });

        div.appendChild(labelElement);
        div.appendChild(input);
        container.appendChild(div);
    }

    /**
     * 创建数字输入框
     */
    createNumberInput(container, label, path, value) {
        const div = document.createElement('div');
        div.className = 'setting-item';

        const labelElement = document.createElement('label');
        labelElement.className = 'setting-label';
        labelElement.textContent = label;

        const input = document.createElement('input');
        input.type = 'number';
        input.className = 'setting-input';
        input.value = value;

        input.addEventListener('change', () => {
            this.updateStyle(path, Number(input.value));
        });

        div.appendChild(labelElement);
        div.appendChild(input);
        container.appendChild(div);
    }

    /**
     * 创建下拉选择框
     */
    createSelect(container, label, path, value, options) {
        const div = document.createElement('div');
        div.className = 'setting-item';

        const labelElement = document.createElement('label');
        labelElement.className = 'setting-label';
        labelElement.textContent = label;

        const select = document.createElement('select');
        select.className = 'setting-input';

        options.forEach(option => {
            const optionElement = document.createElement('option');
            optionElement.value = option.value;
            optionElement.textContent = option.text;
            select.appendChild(optionElement);
        });

        select.value = value;

        select.addEventListener('change', () => {
            this.updateStyle(path, select.value);
        });

        div.appendChild(labelElement);
        div.appendChild(select);
        container.appendChild(div);
    }

    /**
     * 创建开关
     */
    createSwitch(container, label, path, value) {
        const div = document.createElement('div');
        div.className = 'setting-item';

        const labelElement = document.createElement('label');
        labelElement.className = 'setting-label';
        labelElement.textContent = label;

        const switchLabel = document.createElement('label');
        switchLabel.className = 'switch';

        const input = document.createElement('input');
        input.type = 'checkbox';
        input.checked = value;

        const slider = document.createElement('span');
        slider.className = 'slider';

        input.addEventListener('change', () => {
            this.updateStyle(path, input.checked);
        });

        switchLabel.appendChild(input);
        switchLabel.appendChild(slider);

        div.appendChild(labelElement);
        div.appendChild(switchLabel);
        container.appendChild(div);
    }

    /**
     * 切换设置类别
     */
    switchCategory(category) {
        this.currentCategory = category;

        // 隐藏所有设置面板
        document.querySelectorAll('.settings-panel').forEach(panel => {
            panel.classList.remove('active');
        });

        // 显示当前类别的设置面板
        const panel = document.getElementById(`${category}Settings`);
        if (panel) {
            panel.classList.add('active');
        }

        // 更新预览
        this.updatePreview();
    }

    /**
     * 更新样式
     */
    updateStyle(path, value) {
        const parts = path.split('.');
        let current = this.currentStyles;

        for (let i = 0; i < parts.length - 1; i++) {
            current = current[parts[i]];
        }

        current[parts[parts.length - 1]] = value;

        // 更新预览
        this.updatePreview();
    }

    /**
     * 更新预览
     */
    updatePreview() {
        const previewContent = document.getElementById('previewContent');
        if (!previewContent) return;

        // 根据当前类别生成预览内容
        let previewHtml = '';

        switch (this.currentCategory) {
            case 'document':
                previewHtml = this.generateDocumentPreview();
                break;
            case 'heading':
                previewHtml = this.generateHeadingPreview();
                break;
            // 其他类别的预览...
        }

        previewContent.innerHTML = previewHtml;
    }

    /**
     * 生成文档设置预览
     */
    generateDocumentPreview() {
        const settings = this.currentStyles.document;
        return `
            <div style="
                width: ${settings.pageSize === 'A4' ? '210mm' : '297mm'};
                height: ${settings.pageSize === 'A4' ? '297mm' : '210mm'};
                margin: ${settings.margins.top}pt ${settings.margins.right}pt ${settings.margins.bottom}pt ${settings.margins.left}pt;
                border: 1px solid #ccc;
                position: relative;
            ">
                <div style="
                    position: absolute;
                    top: 0;
                    left: 0;
                    right: 0;
                    bottom: 0;
                    display: grid;
                    grid-template-columns: repeat(${settings.grid.charPerLine}, 1fr);
                    grid-template-rows: repeat(${settings.grid.linePerPage}, 1fr);
                    opacity: 0.1;
                ">
                    ${Array(settings.grid.linePerPage).fill().map(() =>
                        Array(settings.grid.charPerLine).fill().map(() =>
                            '<div style="border: 1px solid #ccc;"></div>'
                        ).join('')
                    ).join('')}
                </div>
            </div>
        `;
    }

    /**
     * 生成标题预览
     */
    generateHeadingPreview() {
        const settings = this.currentStyles.heading.styles;
        let previewHtml = '';

        for (let i = 1; i <= 6; i++) {
            const headingSettings = settings[`h${i}`];
            const prefix = headingSettings.numbering.usePrefix ?
                `${headingSettings.numbering.prefix}${i}${headingSettings.numbering.suffix}` : '';

            previewHtml += `
                <h${i} style="
                    font-family: ${headingSettings.font.name}, ${headingSettings.font.fallback.join(', ')};
                    font-size: ${headingSettings.font.size}pt;
                    font-weight: ${headingSettings.font.bold ? 'bold' : 'normal'};
                    text-align: ${headingSettings.paragraph.alignment};
                    margin-top: ${headingSettings.paragraph.spacing.before}pt;
                    margin-bottom: ${headingSettings.paragraph.spacing.after}pt;
                    margin-left: ${headingSettings.paragraph.indent.left}pt;
                    text-indent: ${headingSettings.paragraph.indent.firstLine}pt;
                ">
                    ${prefix} ${i}级标题示例
                </h${i}>
            `;
        }

        return previewHtml;
    }

    /**
     * 保存样式
     */
    async saveStyles() {
        try {
            const response = await fetch('../user-styles.json', {
                method: 'PUT',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(this.currentStyles, null, 2)
            });

            if (response.ok) {
                alert('样式保存成功');
            } else {
                throw new Error('保存失败');
            }
        } catch (error) {
            console.error('保存样式失败:', error);
            alert('保存失败，请重试');
        }
    }

    /**
     * 重置样式
     */
    resetStyles() {
        if (confirm('确定要重置所有样式吗？')) {
            this.currentStyles = JSON.parse(JSON.stringify(this.defaultStyles));
            this.initSettingsPanels();
            this.updatePreview();
        }
    }

    /**
     * 导入样式
     */
    importStyles() {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.json';

        input.addEventListener('change', async (event) => {
            try {
                const file = event.target.files[0];
                const text = await file.text();
                const styles = JSON.parse(text);

                this.currentStyles = styles;
                this.initSettingsPanels();
                this.updatePreview();

                alert('样式导入成功');
            } catch (error) {
                console.error('导入样式失败:', error);
                alert('导入失败，请确保文件格式正确');
            }
        });

        input.click();
    }

    /**
     * 导出样式
     */
    exportStyles() {
        const blob = new Blob([JSON.stringify(this.currentStyles, null, 2)], {
            type: 'application/json'
        });

        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'user-styles.json';
        a.click();

        URL.revokeObjectURL(url);
    }
}

// 初始化编辑器
document.addEventListener('DOMContentLoaded', () => {
    new StyleEditor();
});
