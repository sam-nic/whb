import * as XLSX from 'xlsx';

const APP_VERSION = '1.0.0';

// Ключи в ответе Worker'а — каждый ключ отдаёт СВОЙ массив, значения выровнены по индексу
// (i-й элемент type / section / quantity / unit / description — одна и та же строка заявки).
// Planfix читает каждый ключ отдельным инфоблоком (JSONPath = имя ключа, «Все значения»)
// и сопоставляет type/section по ТЕКСТУ названия записи справочника, не по id.
const FIELD_KEYS = {
  type: 'type',
  section: 'section',
  quantity: 'quantity',
  unit: 'unit',
  description: 'description',
};

// Наименование в Excel -> точное «Название» записи в справочнике «Типы изделий» (7084).
// null = соответствия нет — позиция пока пропускается (список несовпадений — отдельно, по инструкции).
const TYPE_MAP = {
  'Подкладной брус': 'Подкладной брус',
  'Стеновые элементы': 'Балка стеновая',
  'Лаги террасы': null,
  'Перекрытия': 'Балка перекрытия',
  'Балки перголы': 'Балки перголы',
  'Стропила': 'Стропила',
  'Балки двутавровые': null,
  'Стропила разуклонки': 'Стропила разуклонки',
  'Контробрешетка': 'Контробрешетка',
  'Обрешетка внутренняя': 'Обрешетка внутренняя',
  'Обрешетка': 'Обрешетка',
  'Планки межстропильные': 'Планки межстропильные',
  'Столбы': 'Столбы',
  'Доска пола 1 этаж': 'Доска пола',
  'Доска пола 2 этаж': 'Доска пола',
  'Доска террасная': 'Доска террасная',
  'Черновая доска': 'Черновая доска пола',
  'Плинтус пола': 'Плинтус пола',
  'Плинтус потолка внутренний': 'Плинтус потолка внутренний',
  'Плинтус потолка наружный': 'Плинтус потолка наружный',
  'Плинтус кровельный': 'Плинтус кровельный',
  'Обшивка потолков (вагонка)': 'Обшивка потолков (вагонка)',
  'Подшивка свесов крыши (планкен)': 'Обшивка потолков (планкен)',
  'Уголок кровельный': 'Уголок кровельный',
  'Шпонка соединительная стены': 'Шпонка соединительная стены',
  'Шпонка соединительная окна': 'Шпонка соединительная окна',
  'Карнизная и лобовая планка': null, // в 7084 есть 3 варианта (нижний/средний/торцевой профиль) — ждём инструкцию
  'Имитация бруса': 'Имитация бруса',
  'Каркас для имитации': null,
  'Фальшбалки внутренние': 'Фальшбалки внутренние',
  'Фальшбалки наружные': 'Фальшбалки наружные',
  'Накладки нижние': 'Накладки нижние',
};

/* ====================== Planfix REST ====================== */
async function fetchAllEntries(env, token, dirId, fields) {
  let offset = 0;
  const out = [];
  while (true) {
    const r = await fetch(`${env.PLANFIX_REST_BASE}/directory/${dirId}/entry/list`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ offset, pageSize: 100, fields }),
    });
    if (!r.ok) throw new Error(`directory/${dirId}/entry/list -> HTTP ${r.status}`);
    const j = await r.json();
    const entries = j.directoryEntries || [];
    out.push(...entries);
    if (entries.length < 100) break;
    offset += 100;
  }
  return out;
}

function simplify(entries) {
  return entries.map((e) => {
    const vals = {};
    for (const cf of e.customFieldData || []) {
      let v = cf.value;
      if (v && typeof v === 'object') v = v.value;
      vals[cf.field.id] = v;
    }
    return { key: e.key, ...vals };
  });
}

async function loadDirectories(env, token) {
  const [typesRaw, secRaw] = await Promise.all([
    fetchAllEntries(env, token, env.DIR_TYPES_ID, `key,${env.FIELD_TYPE_NAME}`),
    fetchAllEntries(
      env,
      token,
      env.DIR_SECTIONS_ID,
      `key,${env.FIELD_SECTION_NAME},${env.FIELD_SECTION_FINISHED},${env.FIELD_SECTION_GRADE},${env.FIELD_SECTION_VID}`
    ),
  ]);
  const TYPES = simplify(typesRaw).map((v) => ({ key: v.key, name: v[env.FIELD_TYPE_NAME] || '' }));
  const SECTIONS = simplify(secRaw).map((v) => ({
    key: v.key,
    name: v[env.FIELD_SECTION_NAME] || '',
    finished: v[env.FIELD_SECTION_FINISHED] || '',
    grade: v[env.FIELD_SECTION_GRADE] || '',
    vid: v[env.FIELD_SECTION_VID] || '',
  }));
  return { TYPES, SECTIONS };
}

/* ====================== Сопоставление сечения ====================== */
function parsePair(str) {
  const nums = String(str || '').match(/\d+/g);
  if (!nums || nums.length < 2) return null;
  return [parseInt(nums[0], 10), parseInt(nums[1], 10)].sort((a, b) => a - b);
}
function gradeDigit(str) {
  const m = String(str || '').match(/\d/);
  return m ? m[0] : null;
}
function matchSection(SECTIONS, a, b, sortDigit) {
  const pair = [a, b].sort((x, y) => x - y);
  const cands = SECTIONS.filter((s) => {
    const p = parsePair(s.finished);
    return p && p[0] === pair[0] && p[1] === pair[1] && gradeDigit(s.grade) === String(sortDigit);
  });
  if (cands.length === 0) return null;
  const noVid = cands.filter((c) => !c.vid);
  if (noVid.length === 1) return noVid[0];
  const tselny = cands.filter((c) => c.vid === 'Цельный');
  if (tselny.length === 1) return tselny[0];
  if (cands.length === 1) return cands[0];
  return null; // неоднозначно — пропускаем
}

/* ====================== Разбор листа «Заявка» ====================== */
function parseWorkbook(arrayBuffer, TYPES_UNUSED, SECTIONS) {
  const wb = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
  const sheetName =
    wb.SheetNames.find((n) => n.trim().toLowerCase() === 'заявка') ||
    wb.SheetNames.find((n) => /заявк/i.test(n));
  if (!sheetName) throw new Error('Лист «Заявка» не найден в файле');
  const sheet = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });

  const matched = [];
  const skipped = [];

  for (const row of rows) {
    const name = row[1];
    if (name === null || name === undefined || String(name).trim() === '') continue;
    const nameStr = String(name).trim();
    if (/^исполнитель/i.test(nameStr)) break; // конец таблицы

    const a = row[2], b = row[4], qty = row[5], unit = row[6], sort = row[7];
    const numA = typeof a === 'number' ? a : null;
    const numB = typeof b === 'number' ? b : null;
    const numSort = typeof sort === 'number' ? sort : null;
    if (numA === null || numB === null || numSort === null) continue; // не похоже на позицию с сечением
    if (!qty) continue; // нулевое количество — нечего заказывать

    const typeName = TYPE_MAP[nameStr] !== undefined ? TYPE_MAP[nameStr] : null;
    const secEntry = matchSection(SECTIONS, numA, numB, numSort);

    if (!typeName || !secEntry) {
      skipped.push(
        `${nameStr} ${numA}×${numB} ${numSort}с${!typeName ? ' — нет типа' : ''}${!secEntry ? ' — нет сечения' : ''}`
      );
      continue;
    }

    const color = row[9] || '';
    const note = row[15] || '';
    const area = row[16];
    const areaUnit = row[17] || '';
    const descParts = [];
    if (color) descParts.push(String(color).trim());
    if (note) descParts.push(String(note).trim());
    if (typeof area === 'number' && area > 0) descParts.push(`Площадь покраски: ${area} ${areaUnit}`.trim());

    matched.push({
      type: typeName,
      'cross-section': secEntry.name,
      quantity: Math.round(qty * 1000) / 1000,
      unit: unit || '',
      description: descParts.join('; '),
    });
  }
  return { matched, skipped };
}

/* ====================== Гибкое извлечение полей входящего запроса ====================== */
function pickField(body, aliases) {
  for (const key of Object.keys(body || {})) {
    if (aliases.includes(key.toLowerCase())) return body[key];
  }
  return null;
}

/* ====================== Скачивание файла ====================== */
async function downloadFile(token, fileUrl) {
  // Сначала без авторизации (планфиксовые filelink-ссылки обычно уже содержат подписанный &auth=)
  let r = await fetch(fileUrl);
  if ((r.status === 401 || r.status === 403) && token) {
    r = await fetch(fileUrl, { headers: { Authorization: `Bearer ${token}` } });
  }
  if (!r.ok) throw new Error(`Не удалось скачать файл (HTTP ${r.status}) c ${fileUrl}`);
  return r.arrayBuffer();
}

/* ====================== Формирование параллельных массивов для аналитики ====================== */
// taskNo не возвращаем — Planfix и так знает текущую задачу (это её же вебхук).
// Каждое поле — отдельный массив, i-й элемент во всех массивах относится к одной строке заявки.
function buildColumns(matched) {
  return {
    [FIELD_KEYS.type]: matched.map((it) => it.type),
    [FIELD_KEYS.section]: matched.map((it) => it['cross-section']),
    [FIELD_KEYS.quantity]: matched.map((it) => it.quantity),
    [FIELD_KEYS.unit]: matched.map((it) => it.unit),
    [FIELD_KEYS.description]: matched.map((it) => it.description),
  };
}

/* ====================== Worker ====================== */
export default {
  async fetch(request, env) {
    const cors = {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization',
    };

    if (request.method === 'OPTIONS') return new Response(null, { headers: cors });

    if (request.method === 'GET') {
      return Response.json({ status: 'ok', app: 'whb-cross-section-load', version: APP_VERSION }, { headers: cors });
    }

    if (request.method !== 'POST') {
      return Response.json({ error: 'method not allowed' }, { status: 405, headers: cors });
    }

    let body;
    try {
      body = await request.json();
    } catch (e) {
      return Response.json({ error: 'invalid JSON body' }, { status: 400, headers: cors });
    }

    const fileUrl = pickField(body, ['fileurl', 'url', 'file', 'link', 'filelink']);
    // apiKey — ключ к доп. параметрам (Bearer-токен Planfix REST для чтения справочников
    // и скачивания файла). Можно не передавать — тогда используется env.PLANFIX_TOKEN
    // (секрет), для локальной разработки и обратной совместимости.
    // taskNo не используется — Planfix и так знает текущую задачу, это её же вебхук.
    // Обратного вебхука/callback нет — результат отдаётся прямо в ответе на этот запрос
    // (обработка достаточно быстрая, чтобы Planfix мог тут же разобрать ответ и добавить аналитику).
    const apiKey = pickField(body, ['apikey', 'key', 'token', 'planfixtoken']) || env.PLANFIX_TOKEN;

    if (!fileUrl) {
      return Response.json(
        {
          error: 'нужен fileUrl во входящем JSON',
          receivedKeys: Object.keys(body || {}),
          hint: 'поддерживаются алиасы: fileUrl/url/file/link/fileLink; apiKey/key/token',
        },
        { status: 400, headers: cors }
      );
    }
    if (!apiKey) {
      return Response.json(
        { error: 'нет apiKey — ни во входящем JSON, ни в env.PLANFIX_TOKEN' },
        { status: 400, headers: cors }
      );
    }

    try {
      const [{ TYPES, SECTIONS }, fileBuf] = await Promise.all([
        loadDirectories(env, apiKey),
        downloadFile(apiKey, fileUrl),
      ]);

      const { matched, skipped } = parseWorkbook(fileBuf, TYPES, SECTIONS);

      return Response.json(
        {
          ok: matched.length > 0,
          matchedCount: matched.length,
          skippedCount: skipped.length,
          skipped,
          ...buildColumns(matched),
        },
        { status: 200, headers: cors }
      );
    } catch (e) {
      return Response.json({ ok: false, error: e.message, stack: e.stack }, { status: 500, headers: cors });
    }
  },
};
