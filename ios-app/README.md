# Табель для iPhone

Приложение открывает сайт https://kholmerg.github.io/Times-app/ внутри себя,
поэтому изменения на сайте сразу видны в приложении. Сверху добавлены
уведомления о долгой смене и системный запрос геолокации.

## Сборка на Mac (один раз)
1. Установить Xcode из App Store и Node.js (https://nodejs.org, кнопка LTS).
2. Скачать этот репозиторий (на GitHub: Code → Download ZIP) и распаковать.
3. В Терминале:
   cd ~/Downloads/times-app-main/ios-app
   npm install
   npx cap sync ios
   npx cap open ios
4. В Xcode: App → Signing & Capabilities → Team → выбрать свой Apple-аккаунт.
5. Проверка на своём iPhone: подключить iPhone кабелем, выбрать его вверху и нажать ▶.
6. Отправка в App Store: Product → Archive → Distribute App → App Store Connect.
