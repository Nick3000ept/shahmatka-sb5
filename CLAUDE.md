# СБ5 · Шахматка — Административная панель

## Стек
- **Frontend**: `index.html` — один файл, чистый HTML/CSS/JS, без фреймворков
- **Backend**: `script.gs` — Google Apps Script standalone, деплоится через clasp
- **Хостинг**: GitHub Pages — `https://nick3000ept.github.io/shahmatka-sb5/`
- **Репозиторий**: `https://github.com/Nick3000ept/shahmatka-sb5`

## Деплой (все команды по запросу "задеплой")
```bash
cd "d:/YandexDisk/VS_hub/СБ5_шахматка"
# 1. Google Apps Script (бэкенд)
clasp push --force
clasp deploy --description "СБ5 vN"
# Новый deploymentId вставить в index.html (BASE=) и снова clasp push --force
# 2. Git + GitHub Pages
git add index.html script.gs appsscript.json CLAUDE.md .gitignore
git commit -m "описание изменений"
git push
```

> **Важно:** каждый новый `clasp deploy` создаёт НОВЫЙ deploymentId. После деплоя нужно обновить BASE в index.html на новый URL и сделать ещё один clasp push + git push.

## Структура
```
СБ5_шахматка/
  index.html        ← вся фронтенд-логика
  script.gs         ← Apps Script бэкенд
  appsscript.json   ← настройки GAS (webapp, oauthScopes)
  CLAUDE.md         ← этот файл
  .gitignore        ← исключает .clasp.json
```

## Google Apps Script
- **GAS Script ID**: `15e7iz9E_v2mnIYbeibUZVsowCMuwJ4A-HYoHXSpQeP9WpnsKwquVLFis`
- **Текущий deployment URL**: `https://script.google.com/macros/s/AKfycbzYtH05JtG-KI8Mm5laliWtG_jRzHwjqRt49WV4hHuJwbSx1duUTo-8Ct_y_IsxFjGi/exec`
- **Google Sheet ID**: `1Vm0W09F1QNvBi3RK0pWjBa2KB19pXXex9kkd9vHq8zk`
- Лист данных: `Факт`
- Лист подрядчиков: `Подрядчики`
- Лист работ: `Факт_работы`
- Столбцы A–F содержат dropdown-валидацию — **никогда не записывать**
- Столбцы G–N — редактируемые (EDIT_START=7, EDIT_COLS=8)
- O+ не трогать никогда
- Скрипт **standalone**: читает таблицу через `openById(SPREADSHEET_ID)`

## Отличия от СБ3
- `SHEET_NAME = 'Факт'` (в СБ3 было `СБ3_ОБЩАЯ`)
- Все `getActiveSpreadsheet()` заменены на `openById(SPREADSHEET_ID)`
- localStorage/sessionStorage ключи: `sb5_*` (в СБ3 — `sb3_*`) — пресеты, зум, автор, подрядчики полностью раздельные
- Значок в шапке: **СБ5**
- `appsscript.json` содержит явные `oauthScopes` и `webapp` настройки

## Роли пользователей
- **Администратор** — вводит пароль `adminACCB3`
- **СК** — вводит пароль `priemkaCB3`
- **Подрядчик** — открывает ссылку с `?contractor=Название`
- **Просмотр** — без пароля и без параметра contractor

## Аккаунты
- GitHub: Nick3000ept
- GAS: `kuzkin@acons.group`
- clasp авторизован под `kuzkin@acons.group`

## Правила безопасности
- Редактировать ТОЛЬКО: `index.html`, `script.gs`, `CLAUDE.md`, `appsscript.json`, `.gitignore`
- Никогда не трогать файлы вне папки `СБ5_шахматка/`
- Не деплоить без явной команды "задеплой"
- НЕ менять deployment через ручной деплой в GAS-редакторе — только через clasp
