import { useTranslation } from 'react-i18next';
import LanguageSwitcher from './components/LanguageSwitcher';

function App() {
  const { t } = useTranslation();

  return (
    <div className="app">
      <LanguageSwitcher />

      <h1>{t('welcome')}</h1>

      <div className="buttons">
        <button>{t('upload')}</button>
        <button>{t('convert')}</button>
        <button>{t('download')}</button>
        <button>{t('settings')}</button>
        <button>{t('help')}</button>
      </div>
    </div>
  );
}

export default App;
