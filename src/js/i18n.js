// 语言资源
const resources = {
  en: {
    title: "Markdown to Word Converter",
    markdownContent: "Markdown Content",
    preview: "Preview",
    clear: "Clear",
    convertToDocx: "Convert to DOCX",
    dragMessage: "Drag and drop Markdown file here",
    previewPlaceholder: "Converted preview will appear here",
    footer: "© 2024 Markdown to Word Tool - Based on docx.js & marked.js",
    selectFile: "Select File",
    noFile: "No file selected",
    fileInputPlaceholder: "Paste or input Markdown text here...",
    pleaseInputOrUpload: "Please input or upload Markdown content",
    converting: "Converting document, please wait...",
    convertSuccess: "Document generated successfully",
    convertFail: "Failed to convert document: {{msg}}",
    downloadSuccess: "Document saved as {{fileName}}",
    downloadFail: "Download failed: {{msg}}",
    fileTypeError: "Please select a Markdown file!",
    processingImages: "Processing {{count}} images...",
    processImageFail: "Failed to process images: {{msg}}",
    emptyInput: "Please input Markdown content",
    convertingSimple: "Converting...",
    convertSimpleSuccess: "Conversion successful!",
    convertSimpleFail: "Conversion failed: {{msg}}"
  },
  zh: {
    title: "Markdown转Word转换器",
    markdownContent: "Markdown内容",
    preview: "预览",
    clear: "清空",
    convertToDocx: "转换为DOCX",
    dragMessage: "拖放Markdown文件到此处",
    previewPlaceholder: "转换后的预览将显示在这里",
    footer: "© 2024 Markdown转Word工具 - 基于docx.js & marked.js",
    selectFile: "选择文件",
    noFile: "未选择文件",
    fileInputPlaceholder: "在此处粘贴或输入Markdown文本...",
    pleaseInputOrUpload: "请先输入或上传Markdown内容",
    converting: "正在转换文档，请稍候...",
    convertSuccess: "文档已成功生成",
    convertFail: "转换文档失败: {{msg}}",
    downloadSuccess: "文档已保存为 {{fileName}}",
    downloadFail: "下载失败: {{msg}}",
    fileTypeError: "请选择Markdown文件！",
    processingImages: "正在处理{{count}}张图片...",
    processImageFail: "处理图片失败: {{msg}}",
    emptyInput: "请输入Markdown内容",
    convertingSimple: "正在转换...",
    convertSimpleSuccess: "转换成功！",
    convertSimpleFail: "转换失败: {{msg}}"
  },
  es: {
    welcome: "Bienvenido a Md2Docx",
    about: "Acerca de",
    contact: "Contacto",
    login: "Iniciar sesión",
    convert: "Convertir",
    upload: "Subir archivo",
    download: "Descargar",
    settings: "Configuración",
    help: "Ayuda",
    title: "Convertidor de Markdown a Word",
    markdownContent: "Contenido Markdown",
    preview: "Vista previa",
    clear: "Limpiar",
    convertToDocx: "Convertir a DOCX",
    dragMessage: "Arrastra y suelta el archivo Markdown aquí",
    previewPlaceholder: "La vista previa convertida aparecerá aquí",
    footer: "© 2024 Herramienta Markdown a Word - Basado en docx.js & marked.js",
    selectFile: "Seleccionar archivo",
    noFile: "Ningún archivo seleccionado",
    fileInputPlaceholder: "Pega o introduce el texto Markdown aquí...",
    pleaseInputOrUpload: "Por favor, introduce o sube contenido Markdown",
    converting: "Convirtiendo documento, por favor espera...",
    convertSuccess: "Documento generado con éxito",
    convertFail: "Error al convertir el documento: {{msg}}",
    downloadSuccess: "Documento guardado como {{fileName}}",
    downloadFail: "Error al descargar: {{msg}}",
    fileTypeError: "¡Por favor selecciona un archivo Markdown!",
    processingImages: "Procesando {{count}} imágenes...",
    processImageFail: "Error al procesar imágenes: {{msg}}",
    emptyInput: "Por favor, introduce contenido Markdown",
    convertingSimple: "Convirtiendo...",
    convertSimpleSuccess: "¡Conversión exitosa!",
    convertSimpleFail: "Error de conversión: {{msg}}"
  },
  fr: {
    welcome: "Bienvenue sur Md2Docx",
    about: "À propos",
    contact: "Contact",
    login: "Connexion",
    convert: "Convertir",
    upload: "Télécharger un fichier",
    download: "Télécharger",
    settings: "Paramètres",
    help: "Aide",
    title: "Convertisseur Markdown vers Word",
    markdownContent: "Contenu Markdown",
    preview: "Aperçu",
    clear: "Effacer",
    convertToDocx: "Convertir en DOCX",
    dragMessage: "Glissez-déposez le fichier Markdown ici",
    previewPlaceholder: "L'aperçu converti s'affichera ici",
    footer: "© 2024 Outil Markdown vers Word - Basé sur docx.js & marked.js",
    selectFile: "Sélectionner un fichier",
    noFile: "Aucun fichier sélectionné",
    fileInputPlaceholder: "Collez ou saisissez le texte Markdown ici...",
    pleaseInputOrUpload: "Veuillez saisir ou télécharger du contenu Markdown",
    converting: "Conversion du document, veuillez patienter...",
    convertSuccess: "Document généré avec succès",
    convertFail: "Échec de la conversion du document : {{msg}}",
    downloadSuccess: "Document enregistré sous {{fileName}}",
    downloadFail: "Échec du téléchargement : {{msg}}",
    fileTypeError: "Veuillez sélectionner un fichier Markdown !",
    processingImages: "Traitement de {{count}} images...",
    processImageFail: "Échec du traitement des images : {{msg}}",
    emptyInput: "Veuillez saisir du contenu Markdown",
    convertingSimple: "Conversion...",
    convertSimpleSuccess: "Conversion réussie !",
    convertSimpleFail: "Échec de la conversion : {{msg}}"
  },
  ru: {
    welcome: "Добро пожаловать в Md2Docx",
    about: "О нас",
    contact: "Контакты",
    login: "Вход",
    convert: "Конвертировать",
    upload: "Загрузить файл",
    download: "Скачать",
    settings: "Настройки",
    help: "Помощь",
    title: "Конвертер Markdown в Word",
    markdownContent: "Содержимое Markdown",
    preview: "Предпросмотр",
    clear: "Очистить",
    convertToDocx: "Преобразовать в DOCX",
    dragMessage: "Перетащите файл Markdown сюда",
    previewPlaceholder: "Преобразованный предпросмотр появится здесь",
    footer: "© 2024 Инструмент Markdown в Word - на основе docx.js & marked.js",
    selectFile: "Выбрать файл",
    noFile: "Файл не выбран",
    fileInputPlaceholder: "Вставьте или введите текст Markdown здесь...",
    pleaseInputOrUpload: "Пожалуйста, введите или загрузите содержимое Markdown",
    converting: "Преобразование документа, пожалуйста, подождите...",
    convertSuccess: "Документ успешно создан",
    convertFail: "Не удалось преобразовать документ: {{msg}}",
    downloadSuccess: "Документ сохранён как {{fileName}}",
    downloadFail: "Ошибка загрузки: {{msg}}",
    fileTypeError: "Пожалуйста, выберите файл Markdown!",
    processingImages: "Обработка {{count}} изображений...",
    processImageFail: "Ошибка обработки изображений: {{msg}}",
    emptyInput: "Пожалуйста, введите содержимое Markdown",
    convertingSimple: "Преобразование...",
    convertSimpleSuccess: "Преобразование успешно!",
    convertSimpleFail: "Ошибка преобразования: {{msg}}"
  },
  ar: {
    welcome: "مرحبًا بك في Md2Docx",
    about: "حول",
    contact: "اتصل بنا",
    login: "تسجيل الدخول",
    convert: "تحويل",
    upload: "رفع ملف",
    download: "تحميل",
    settings: "الإعدادات",
    help: "المساعدة",
    title: "محول Markdown إلى Word",
    markdownContent: "محتوى Markdown",
    preview: "المعاينة",
    clear: "مسح",
    convertToDocx: "تحويل إلى DOCX",
    dragMessage: "اسحب ملف Markdown وأفلته هنا",
    previewPlaceholder: "سيظهر المعاينة المحولة هنا",
    footer: "© 2024 أداة تحويل Markdown إلى Word - مبني على docx.js & marked.js",
    selectFile: "اختر ملفًا",
    noFile: "لم يتم اختيار ملف",
    fileInputPlaceholder: "الصق أو أدخل نص Markdown هنا...",
    pleaseInputOrUpload: "يرجى إدخال أو رفع محتوى Markdown",
    converting: "جاري تحويل المستند، يرجى الانتظار...",
    convertSuccess: "تم إنشاء المستند بنجاح",
    convertFail: "فشل في تحويل المستند: {{msg}}",
    downloadSuccess: "تم حفظ المستند باسم {{fileName}}",
    downloadFail: "فشل في التحميل: {{msg}}",
    fileTypeError: "يرجى اختيار ملف Markdown!",
    processingImages: "جاري معالجة {{count}} صورة...",
    processImageFail: "فشل في معالجة الصور: {{msg}}",
    emptyInput: "يرجى إدخال محتوى Markdown",
    convertingSimple: "جاري التحويل...",
    convertSimpleSuccess: "تم التحويل بنجاح!",
    convertSimpleFail: "فشل في التحويل: {{msg}}"
  }
};

// 当前语言
let currentLang = localStorage.getItem('language') || 'zh';

// 翻译函数
function t(key) {
  return resources[currentLang][key] || resources['en'][key] || key;
}

// 支持变量替换的翻译函数
function tWithVars(key, vars = {}) {
  let str = t(key);
  Object.keys(vars).forEach(k => {
    str = str.replace(new RegExp(`{{${k}}}`, 'g'), vars[k]);
  });
  return str;
}

// 切换语言
function changeLanguage(lang) {
  currentLang = lang;
  localStorage.setItem('language', lang);
  updateContent();
  // 同步select的显示
  const languageSelect = document.getElementById('language-select');
  if (languageSelect) languageSelect.value = lang;
  // RTL支持：阿拉伯语时设置为rtl，否则为ltr
  if (lang === 'ar') {
    document.documentElement.dir = 'rtl';
  } else {
    document.documentElement.dir = 'ltr';
  }
}

// 安全地更新元素文本
function safeUpdateElement(selector, text) {
  const element = document.querySelector(selector);
  if (element) {
    element.textContent = text;
  }
}

// 安全地更新元素属性
function safeUpdateAttribute(selector, attribute, value) {
  const element = document.querySelector(selector);
  if (element) {
    element.setAttribute(attribute, value);
  }
}

// 更新页面内容
function updateContent() {
  try {
    // 更新标题
    document.title = t('title');
    safeUpdateElement('.topbar-title h1', t('title'));

    // 更新Markdown内容区域
    safeUpdateElement('.markdown-input-section h2', t('markdownContent'));
    safeUpdateElement('#clear-btn', t('clear'));

    const dragMessage = document.querySelector('.drag-message div:last-child');
    if (dragMessage) {
      dragMessage.textContent = t('dragMessage');
    }

    // 更新自定义文件按钮和文件名显示
    safeUpdateElement('#custom-file-btn', t('selectFile'));
    // 只在未选择文件时显示"未选择文件"
    const fileInputEl = document.getElementById('file-input');
    if (fileInputEl && !fileInputEl.files[0]) {
      safeUpdateElement('#file-name-label', t('noFile'));
    }

    // 更新文本区域占位符
    const textarea = document.querySelector('#markdown-input');
    if (textarea) {
      textarea.placeholder = t('fileInputPlaceholder');
    }

    // 更新预览区域
    safeUpdateElement('.preview-section h2', t('preview'));
    safeUpdateElement('#direct-convert-btn', t('convertToDocx'));
    safeUpdateElement('.preview-placeholder', t('previewPlaceholder'));

    // 更新页脚
    safeUpdateElement('footer p', t('footer'));

    // 更新语言切换按钮状态
    document.querySelectorAll('.language-btn').forEach(btn => {
      if (btn) {
        btn.classList.toggle('active', btn.dataset.lang === currentLang);
      }
    });
  } catch (error) {
    console.error('Error updating content:', error);
  }
}

// 初始化语言切换下拉菜单
function initLanguageSwitcher() {
  const languageSelect = document.getElementById('language-select');
  if (!languageSelect) return;
  languageSelect.value = currentLang;
  languageSelect.onchange = function() {
    changeLanguage(languageSelect.value);
  };
}

// 导出函数
export { t, tWithVars, changeLanguage, initLanguageSwitcher, updateContent };
