# СБ5 · Шахматка — Административная панель

## Стек
- **Frontend**: `index.html` — один файл, чистый HTML/CSS/JS, без фреймворков
- **Backend**: `script.gs` — Google Apps Script (хранится локально для версионирования, деплоится вручную)
- **Хостинг**: GitHub Pages (нужно создать отдельный репозиторий)

## Деплой (все команды по запросу "задеплой")
```bash
cd "d:/YandexDisk/VS_hub/СБ5_шахматка"
# 1. Google Apps Script (бэкенд)
clasp push --force
clasp deploy --deploymentId AKfycbxEk5colwld9ixqAdTIK2HxLO3C3TgoPKX18ec-1aFYQp68vDlmX28BO7MlWbwrHP4S
# 2. Git + GitHub Pages
git add index.html script.gs appsscript.json CLAUDE.md .gitignore
git commit -m "описание изменений"
git push
```

## Структура
```
СБ5_шахматка/
  index.html   ← вся фронтенд-логика
  script.gs    ← Apps Script бэкенд
  CLAUDE.md    ← этот файл
  .gitignore
```

## Google Apps Script
- **GAS Script ID**: `15e7iz9E_v2mnIYbeibUZVsowCMuwJ4A-HYoHXSpQeP9WpnsKwquVLFis`
- **GAS deployment URL**: `https://script.google.com/macros/s/AKfycbxEk5colwld9ixqAdTIK2HxLO3C3TgoPKX18ec-1aFYQp68vDlmX28BO7MlWbwrHP4S/exec`
- **Google Sheet ID**: `1Vm0W09F1QNvBi3RK0pWjBa2KB19pXXex9kkd9vHq8zk`
- Лист данных: `Факт`
- Лист подрядчиков: `Подрядчики`
- Лист работ: `Факт_работы`
- Столбцы A–F содержат dropdown-валидацию — **никогда не записывать**
- Столбцы G–N — редактируемые
- O+ не трогать никогда
- Скрипт **standalone** (не container-bound): читает таблицу через `openById(SPREADSHEET_ID)`

## Роли пользователей
- **Администратор** — вводит пароль `adminACCB3`
- **Подрядчик** — открывает ссылку с `?contractor=Название`
- **Просмотр** — без пароля и без параметра contractor

## Аккаунты
- GAS: аккаунт `kuzkin@acons.group`

## Правила безопасности
- Редактировать ТОЛЬКО: `index.html`, `script.gs`, `CLAUDE.md`, `appsscript.json`, `.gitignore`
- Никогда не трогать файлы вне папки `СБ5_шахматка/`
- Не деплоить без явной команды "задеплой"
