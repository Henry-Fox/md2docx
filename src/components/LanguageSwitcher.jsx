import i18n from 'i18next';

// 语言配置数据
const languages = [
  { code: 'en', name: 'English' },
  { code: 'zh', name: '中文' },
  { code: 'fr', name: 'Français' },
  { code: 'ru', name: 'Русский' },
  { code: 'es', name: 'Español' },
  { code: 'ar', name: 'العربية' }
];

function LanguageSwitcher() {
  return (
    <div className="language-switcher">
      {languages.map((lang) => (
        <button
          key={lang.code}
          onClick={() => {
            i18n.changeLanguage(lang.code);
            document.documentElement.dir = lang.code === 'ar' ? 'rtl' : 'ltr';
          }}
          className={i18n.language === lang.code ? 'active' : ''}
        >
          {lang.name}
        </button>
      ))}
    </div>
  );
}

export default LanguageSwitcher;
