/**
 * DigiDocs — генерация структурированных текстов в Google Docs
 * Версия: 1.1
 * Автор: Данила Шестаков
 *
 * Хранение секретов:
 *  - OPENROUTER_KEY — ключ OpenRouter
 *  - TEXT_RU_KEY    — ключ Text.ru
 *  - OPENROUTER_MODEL — модель по умолчанию (gpt-4o)
 */

function onOpen() {
  DocumentApp.getUi()
    .createMenu('DigiDocs')
    .addItem('Новая статья по ТЗ', 'showPromptDialog')
    .addSeparator()
    .addItem('Проверить уникальность (Text.ru)', 'checkUniquenessTextRuOneClick')
    .addItem('Проверить баланс Text.ru', 'checkTextRuBalance')
    .addSeparator()
    .addItem('Настройки', 'showSettingsDialog')
    .addToUi();
}

const TEMPLATE_ID = '14TlJoMWMrylAziu22MPa6FiRbhopKczkj76no0_4bow'; // ID твоего шаблона
const OPENROUTER_BASE = 'https://openrouter.ai/api/v1/chat/completions';

// ---------- UI ----------

function showPromptDialog() {
  const html = HtmlService.createHtmlOutputFromFile('index')
    .setWidth(460)
    .setHeight(640);
  DocumentApp.getUi().showModalDialog(html, 'Генерация статьи');
}

function showSettingsDialog() {
  const html = HtmlService.createHtmlOutputFromFile('settings')
    .setWidth(520)
    .setHeight(420);
  DocumentApp.getUi().showModalDialog(html, 'Настройки DigiDocs');
}

// ---------- Настройки / свойства ----------

function getSettings() {
  const sp = PropertiesService.getScriptProperties();
  return {
    OPENROUTER_BASE: sp.getProperty('OPENROUTER_BASE') || 'https://openrouter.ai/api/v1/chat/completions',
    OPENROUTER_MODEL: sp.getProperty('OPENROUTER_MODEL') || 'gpt-4o',
    TEMPLATE_ID: sp.getProperty('TEMPLATE_ID') || '',
    // Ключи возвращаем в усеченном виде, чтобы не подсвечивать полностью
    OPENROUTER_KEY_MASKED: maskKey_(sp.getProperty('OPENROUTER_KEY')),
    TEXT_RU_KEY_MASKED: maskKey_(sp.getProperty('TEXT_RU_KEY'))
  };
}

function saveSettings(data) {
  if (!data || typeof data !== 'object') throw new Error('Некорректные данные настроек');
  const sp = PropertiesService.getScriptProperties();

  if ('OPENROUTER_KEY' in data && data.OPENROUTER_KEY) sp.setProperty('OPENROUTER_KEY', data.OPENROUTER_KEY.trim());
  if ('TEXT_RU_KEY' in data && data.TEXT_RU_KEY) sp.setProperty('TEXT_RU_KEY', data.TEXT_RU_KEY.trim());
  if ('TEMPLATE_ID' in data) sp.setProperty('TEMPLATE_ID', (data.TEMPLATE_ID || '').trim());
  if ('OPENROUTER_BASE' in data && data.OPENROUTER_BASE) sp.setProperty('OPENROUTER_BASE', data.OPENROUTER_BASE.trim());
  if ('OPENROUTER_MODEL' in data && data.OPENROUTER_MODEL) sp.setProperty('OPENROUTER_MODEL', data.OPENROUTER_MODEL.trim());

  return getSettings();
}

function maskKey_(val) {
  if (!val) return '';
  const v = String(val);
  if (v.length <= 8) return '••••';
  return v.slice(0, 4) + '••••' + v.slice(-4);
}

// ---------- Генерация статьи ----------

function generateArticleWithOptions(data) {
  const { prompt, model, format, tone, length, useSubheadings, title } = sanitizeInput_(data);
  const fullPrompt = buildPrompt(prompt, format, tone, length, useSubheadings);
  const raw = callOpenRouter(fullPrompt, model);
  const clean = enforceLength(raw, length);

  let finalText = clean;
  if (!hasConclusion(clean)) {
    const extra = callOpenRouter('Напиши короткое заключение к следующей статье:\n\n' + clean, model);
    finalText = clean.trim() + '\n\n## Заключение\n' + extra.trim();
  }

  let h1 = (title || '').trim();
  if (!h1) {
    const h1Match = finalText.match(/^#\s+(.*)/m);
    if (h1Match) h1 = h1Match[1].trim();
  }
  if (!h1) h1 = 'Сгенерированная статья';

  const url = insertMarkdownToDoc(finalText, h1);
  return url;
}

function sanitizeInput_(data) {
  if (!data || typeof data !== 'object') throw new Error('Нет данных');
  const length = Math.max(300, Math.min(20000, parseInt(data.length, 10) || 1500));
  return {
    prompt: String(data.prompt || '').trim(),
    model: String(data.model || getDefaultModel_()),
    format: String(data.format || 'статья'),
    tone: String(data.tone || 'информативный'),
    length: length,
    useSubheadings: Boolean(data.useSubheadings !== false),
    title: String(data.title || '').trim()
  };
}

function getDefaultModel_() {
  const sp = PropertiesService.getScriptProperties();
  return sp.getProperty('OPENROUTER_MODEL') || 'gpt-4o';
}

function hasConclusion(text) {
  return /##\s*(Заключение|Вывод|Итоги|Резюме)/i.test(text);
}

function buildPrompt(basePrompt, format, tone, length, useSubheadings) {
  let styleHint = '';
  switch (format) {
    case 'лонгрид':
      styleHint = 'Напиши развёрнутый лонгрид: вступление, блоки с подзаголовками, отдельное заключение. Каждый блок логически завершён.';
      break;
    case 'карточка':
      styleHint = 'Продающая карточка товара: сначала выгода, затем особенности и преимущества. Короткие абзацы, без воды.';
      break;
    case 'новость':
      styleHint = 'Новостная заметка: заголовок, короткий лид, основной текст. Чётко, фактологично, без клише.';
      break;
    case 'страница интернет-магазина':
      styleHint = 'SEO-текст для категории магазина: подзаголовки, списки, преимущества ассортимента, в конце вывод/CTA.';
      break;
    default:
      styleHint = 'Тематическая статья: ввод, 2–4 смысловых раздела и отдельное заключение с выводами.';
  }

  const heads = useSubheadings ? 'Используй подзаголовки в формате Markdown (## Название).' : 'Подзаголовки по усмотрению.';
  return [
    styleHint,
    'Тональность: ' + tone + '.',
    heads,
    'Объём: не менее ' + length + ' знаков.',
    'ТЗ: ' + basePrompt
  ].join('\n');
}

function callOpenRouter(prompt, model) {
  const sp = PropertiesService.getScriptProperties();
  const apiKey = sp.getProperty('OPENROUTER_KEY'); // ключ всё ещё хранится в настройках
  if (!apiKey) throw new Error('Не задан ключ OPENROUTER_KEY. Откройте «Настройки» и сохраните ключ.');

  const payload = {
    model: model || getDefaultModel_(),
    temperature: 0.7,
    messages: [
      { role: 'system', content: 'Ты профессиональный копирайтер. Пиши строго по ТЗ, структурировано, без воды.' },
      { role: 'user', content: prompt }
    ]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + apiKey,
      'X-Title': 'DigiDocs',
      'HTTP-Referer': 'https://docs.google.com'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    followRedirects: true
  };

  const resp = UrlFetchApp.fetch(OPENROUTER_BASE, options);
  const status = resp.getResponseCode();
  const text = resp.getContentText();

  if (status < 200 || status >= 300) {
    throw new Error('Ошибка OpenRouter (' + status + '): ' + text);
  }

  const result = JSON.parse(text);
  const content = result?.choices?.[0]?.message?.content;
  if (!content) throw new Error('Ответ от OpenRouter пуст или неверного формата: ' + text);
  return content;
}

function enforceLength(text, target) {
  const normalized = text.replace(/[ \t]+/g, ' ');
  if (normalized.length <= target) return text;

  const safeSlice = text.slice(0, target + 400);
  const markerMatch = safeSlice.match(/##\s*(Заключение|Вывод|Итоги|Резюме)/i);
  if (markerMatch) {
    const start = markerMatch.index;
    const after = safeSlice.slice(start);
    const endDot = after.indexOf('.') + 1;
    if (endDot > 0) return safeSlice.slice(0, start + endDot).trim();
  }

  const lastDot = safeSlice.lastIndexOf('.');
  const cutoff = lastDot > target * 0.8 ? lastDot + 1 : target;
  return safeSlice.slice(0, cutoff).trim() + '…';
}

// Переписанная функция вставки Markdown в Google Docs с копированием шаблона
function insertMarkdownToDoc(markdown, title) {

  const safeTitle = String(title || 'Сгенерированная статья')
    .replace(/[\r\n]+/g, ' ')
    .replace(/[\/\\:*?"<>|#]+/g, '—')
    .trim()
    .slice(0, 180); // запас по имени файла

  const md = String(markdown || '')
    .replace(/\r\n/g, '\n')      // Windows → Unix
    .replace(/\r/g, '\n')
    .replace(/[ \t]+\n/g, '\n')  // хвостовые пробелы
    .replace(/\n{3,}/g, '\n\n'); // больше двух пустых → две


  let file, doc;
  if (TEMPLATE_ID) {
    file = DriveApp.getFileById(TEMPLATE_ID).makeCopy(safeTitle);
    doc  = DocumentApp.openById(file.getId());
  } else {
    doc  = DocumentApp.create(safeTitle);
    file = DriveApp.getFileById(doc.getId());
  }

  // 3) Расшарим по ссылке на комментарии (при необходимости можно поменять права)
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.COMMENT);
  } catch (e) {
    // В доменах с ограничениями может не получиться — не валим выполнение
    Logger.log('Sharing warn: ' + e);
  }

  // 4) Очищаем документ и пишем содержимое
  const body = doc.getBody();
  body.clear();

  // Если в начале нет H1 — добавим из заголовка
  const hasH1 = /^#\s+.+/m.test(md);
  if (!hasH1 && safeTitle) {
    const p = body.appendParagraph(safeTitle);
    p.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph(''); // отступ
  }

  // 5) Построчный разбор Markdown (заголовки, списки, обычные абзацы)
  const lines = md.split('\n');
  const cleaned = [];
  for (let i = 0; i < lines.length; i++) {
    const cur  = lines[i].trim();
    const next = (lines[i + 1] || '').trim();
    const isHeading = /^#{1,6}\s+/.test(cur);

    // убираем двойной пустой перед заголовком
    if (isHeading && cleaned[cleaned.length - 1] === '') cleaned.pop();
    cleaned.push(cur);
    // у заголовков часто идёт пустая строка — пропустим лишнюю
    if (isHeading && next === '') i++;
  }

  cleaned.forEach(line => {
    if (line === '') { body.appendParagraph(''); return; }

    const mHeading  = line.match(/^(#{1,6})\s+(.*)$/);
    const mBullet   = line.match(/^(\*|-)\s+(.*)$/);
    const mNumbered = line.match(/^(\d+)\.\s+(.*)$/);

    if (mHeading) {
      const level = mHeading[1].length;
      const text  = mHeading[2];
      const p     = body.appendParagraph(text);
      switch (level) {
        case 1: p.setHeading(DocumentApp.ParagraphHeading.HEADING1); break;
        case 2: p.setHeading(DocumentApp.ParagraphHeading.HEADING2); break;
        case 3: p.setHeading(DocumentApp.ParagraphHeading.HEADING3); break;
        case 4: p.setHeading(DocumentApp.ParagraphHeading.HEADING4); break;
        case 5: p.setHeading(DocumentApp.ParagraphHeading.HEADING5); break;
        default: p.setHeading(DocumentApp.ParagraphHeading.HEADING6);
      }
      return;
    }

    if (mBullet) {
      const li = body.appendListItem('');
      li.setGlyphType(DocumentApp.GlyphType.BULLET);
      applyMarkdownStyles(li.editAsText(), mBullet[2]);
      return;
    }

    if (mNumbered) {
      const li = body.appendListItem('');
      li.setGlyphType(DocumentApp.GlyphType.NUMBER);
      applyMarkdownStyles(li.editAsText(), mNumbered[2]);
      return;
    }

    const p = body.appendParagraph('');
    applyMarkdownStyles(p.editAsText(), line);
  });

  try { doc.setName(safeTitle); } catch (e) { Logger.log('setName warn: ' + e); }
  doc.saveAndClose();

  return doc.getUrl();
}


function applyMarkdownStyles(textElement, input) {
  let plain = input;
  const styles = [];

  // [текст](url)
  const linkPattern = /\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g;
  let match;
  while ((match = linkPattern.exec(plain)) !== null) {
    const start = plain.indexOf(match[0]);
    const end = start + match[1].length;
    styles.push({ type: 'link', start, end, url: match[2] });
    plain = plain.replace(match[0], match[1]);
  }

  // **жирный**
  const boldPattern = /\*\*(.+?)\*\*/g;
  while ((match = boldPattern.exec(plain)) !== null) {
    const start = plain.indexOf(match[0]);
    const end = start + match[1].length;
    styles.push({ type: 'bold', start, end });
    plain = plain.replace(match[0], match[1]);
  }

  // *курсив*
  const italicPattern = /\*(?!\*)(.+?)\*/g;
  while ((match = italicPattern.exec(plain)) !== null) {
    const start = plain.indexOf(match[0]);
    const end = start + match[1].length;
    styles.push({ type: 'italic', start, end });
    plain = plain.replace(match[0], match[1]);
  }

  textElement.setText(plain);

  styles.forEach(s => {
    try {
      if (s.type === 'bold') textElement.setBold(s.start, s.end - 1, true);
      if (s.type === 'italic') textElement.setItalic(s.start, s.end - 1, true);
      if (s.type === 'link') textElement.setLinkUrl(s.start, s.end - 1, s.url);
    } catch (e) {
      Logger.log('Ошибка применения стиля: ' + e);
    }
  });
}

// ---------- Text.ru: асинхронная проверка ----------

function checkUniquenessTextRuOneClick() {
  const ui = DocumentApp.getUi();
  try {
    const sp = PropertiesService.getScriptProperties();
    const key = sp.getProperty('TEXT_RU_KEY');
    if (!key) { ui.alert('❌ Не задан ключ TEXT_RU_KEY. Откройте «Настройки» и сохраните ключ.'); return; }

    const doc = DocumentApp.getActiveDocument();
    const text = (doc.getBody().getText() || '').trim();
    if (!text || text.length < 200) { ui.alert('⚠️ Текст слишком короткий или пустой (минимум 200 знаков).'); return; }

    // 1) Отправляем текст на проверку
    const postResp = UrlFetchApp.fetch('https://api.text.ru/post', {
      method: 'post',
      payload: { text: text, userkey: key, visible: 'vis_on', json: '1' },
      muteHttpExceptions: true
    });
    const postCode = postResp.getResponseCode();
    const postBody = postResp.getContentText();
    if (postCode < 200 || postCode >= 300) { ui.alert('❌ Ошибка Text.ru (' + postCode + '): ' + postBody); return; }

    let postJson;
    try { postJson = JSON.parse(postBody); } catch(e) { ui.alert('❌ Некорректный JSON от Text.ru: ' + postBody); return; }
    if (postJson.error_code || !postJson.text_uid) { ui.alert('❌ Text.ru: ' + (postJson.error_desc || 'UID не получен')); return; }

    const uid = postJson.text_uid;

    // 2) Ждём результат (поллинг с backoff)
    const maxWaitMs = 300000; // 5 минут максимум
    let elapsed = 0;
    let delay = 4000; // старт 4с
    let resultJson = null;

    while (elapsed < maxWaitMs) {
      Utilities.sleep(delay);
      elapsed += delay;

      const checkResp = UrlFetchApp.fetch('https://api.text.ru/post', {
        method: 'post',
        payload: { userkey: key, uid: uid, json: '1' },
        muteHttpExceptions: true
      });
      const checkCode = checkResp.getResponseCode();
      const checkBody = checkResp.getContentText();
      if (checkCode < 200 || checkCode >= 300) {
        ui.alert('❌ Ошибка при запросе статуса (' + checkCode + '): ' + checkBody);
        return;
      }

      let checkJson;
      try { checkJson = JSON.parse(checkBody); } catch(e) { ui.alert('❌ Некорректный JSON статуса: ' + checkBody); return; }

      if (checkJson.error_code == 181) {
        // ещё в работе — увеличим задержку до 10с максимум
        delay = Math.min(delay + 2000, 10000);
        continue;
      }
      if (checkJson.error_code) {
        ui.alert('❌ Text.ru: ' + checkJson.error_desc);
        return;
      }
      if (checkJson.text_unique != null) {
        resultJson = checkJson;
        break;
      }
      // перестраховка — небольшой прирост задержки
      delay = Math.min(delay + 1000, 10000);
    }

    if (!resultJson) {
      ui.alert('⏳ Проверка не завершилась в отведённое время. Попробуйте ещё раз позже.');
      return;
    }

    const uniq = Number(resultJson.text_unique);
    const b = doc.getBody();
    b.appendParagraph('');
    b.appendParagraph('✅ Уникальность текста: ' + (isNaN(uniq) ? String(resultJson.text_unique) : uniq.toFixed(2) + '%'));
    b.appendParagraph('🔗 Отчёт: https://text.ru/antiplagiat/' + uid);

    ui.alert('✅ Готово: уникальность ' + (isNaN(uniq) ? String(resultJson.text_unique) : uniq.toFixed(2) + '%') + '. Ссылка добавлена в документ.');
  } catch (e) {
    ui.alert('❌ Ошибка: ' + e);
  }
}


function checkTextRuBalance() {
  const ui = DocumentApp.getUi();
  try {
    const sp = PropertiesService.getScriptProperties();
    const key = sp.getProperty('TEXT_RU_KEY');
    if (!key) {
      ui.alert('❌ Не задан ключ TEXT_RU_KEY. Откройте «Настройки» и сохраните ключ.');
      return;
    }

    const resp = UrlFetchApp.fetch('https://api.text.ru/account', {
      method: 'post',
      payload: { userkey: key, method: 'get_packages_info', json: '1' },
      muteHttpExceptions: true
    });

    const status = resp.getResponseCode();
    const body = resp.getContentText();
    if (status < 200 || status >= 300) {
      ui.alert('❌ Ошибка Text.ru (' + status + '): ' + body);
      return;
    }

    const result = JSON.parse(body);
    if (result.error_code) {
      ui.alert('❌ Text.ru: ' + result.error_desc);
      return;
    }

    const balance = result.size || 0;
    ui.alert('💰 Остаток символов: ' + Number(balance).toLocaleString('ru-RU'));
  } catch (e) {
    ui.alert('❌ Ошибка: ' + e);
  }
}
