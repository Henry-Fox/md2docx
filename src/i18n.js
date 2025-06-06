import i18n from 'i18next';
import { initReactI18next } from 'react-i18next';

// 联合国官方语言配置
const resources = {
  en: { translation: require('./locales/en.json') },
  zh: { translation: require('./locales/zh.json') },
  fr: { translation: require('./locales/fr.json') },
  ru: { translation: require('./locales/ru.json') },
  es: { translation: require('./locales/es.json') },
  ar: { translation: require('./locales/ar.json') }
};

i18n
  .use(initReactI18next)
  .init({
    resources,
    lng: 'en', // 默认语言
    fallbackLng: 'en', // 回退语言
    interpolation: {
      escapeValue: false
    }
  });

export default i18n;
