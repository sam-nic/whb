# whb-cross-section-load

Cloudflare Worker: принимает входящий вебхук от Planfix `{taskNo, fileUrl}`, скачивает файл заявки
(бланк заказа, лист «Заявка»), сопоставляет позиции со справочниками Planfix
**Типы изделий** (7084) и **Сечения изделий** (7082), и отправляет распознанные позиции
исходящим вебхуком `dt_working_drawings_items` (аналитика «Изделия РД»).

## Установка

```bash
npm install
npx wrangler secret put PLANFIX_TOKEN   # Bearer-токен Planfix REST (whb.planfix.ru)
```

Для локальной разработки — положить тот же токен в `.dev.vars` (не коммитится):

```
PLANFIX_TOKEN=...
```

## Разработка

```bash
npm run dev   # wrangler dev --config ./wrangler.toml, http://localhost:8787
```

**Важно:** в `~/Downloads` лежит чужой `wrangler.jsonc` (от другого проекта, seppra-dashboard) —
wrangler может подхватить его вместо локального `wrangler.toml`, если запускать без явного
`--config ./wrangler.toml`. Все npm-скрипты (`dev`/`deploy`/`tail`) уже используют
`wrangler` без явного конфига, поэтому запускать их нужно из этой директории или добавить
`--config` вручную, если возникнет ошибка «entry-point file... was not found».

## Деплой

```bash
npm run deploy
```

## Вход (POST /)

```json
{ "taskNo": 12345, "fileUrl": "https://whb.planfix.ru/filelink/...&auth=..." }
```

Алиасы ключей: `taskNo`/`task`/`taskNumber`; `fileUrl`/`url`/`file`/`link`/`fileLink`
(сопоставление регистронезависимое). Файл сначала скачивается без авторизации
(planfix filelink обычно уже содержит подписанный `&auth=`), при 401/403 — повторно
с заголовком `Authorization: Bearer <PLANFIX_TOKEN>`.

## Выход

Отвечает вызывающему (для отладки/теста прямо из настроек вебхука Planfix) сводкой:

```json
{
  "ok": true,
  "taskNo": 12345,
  "matchedCount": 18,
  "skippedCount": 2,
  "skipped": ["Стропила 152×246 2с — нет сечения", "..."],
  "webhookStatus": 200,
  "webhookResponse": "..."
}
```

И параллельно шлёт сам результат на `dt_working_drawings_items?taskNo=...` —
JSON-массив объектов `{taskNo, type, "cross-section", quantity, unit, description}`,
где `type`/`cross-section` — точные текстовые названия записей справочников
(вебхук сопоставляет по названию, не по id).

## Известные пропуски (ждут отдельной инструкции — сейчас просто not sent)

- Позиции без соответствия в справочнике 7084 (Тип): «Лаги террасы», «Балки двутавровые»,
  «Карнизная и лобовая планка» (3 варианта — нижний/средний/торцевой профиль, порядок неясен),
  «Каркас для имитации».
- Сечения, которых нет в справочнике 7082 при заданном сорте (напр. Стропила 152×246 2с).
- Позиции с нестандартным форматом сечения («Накладки нижние» — «на 240» вместо A×B) и раздел
  «Окна и двери» — не парсятся вовсе (другая структура таблицы).

Список несовпадений видно в ответе (`skipped`) и в исходном клиентском прототипе
[`whb-zayavka-rd/zayavka_rd.html`](../../whb-zayavka-rd/zayavka_rd.html) (там же таблица
`TYPE_MAP` — источник соответствий, скопирована 1:1 в `src/index.js`).
