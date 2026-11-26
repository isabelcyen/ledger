// 語音辨識文字 → 金額／項目／幣別 的解析工具（純函式，可在 Node 測試）

const DIGIT = {
  零: 0, 〇: 0, 一: 1, 壹: 1, 二: 2, 貳: 2, 贰: 2, 兩: 2, 两: 2, 三: 3, 參: 3, 叁: 3,
  四: 4, 肆: 4, 五: 5, 伍: 5, 六: 6, 陸: 6, 七: 7, 柒: 7, 八: 8, 捌: 8, 九: 9, 玖: 9,
};
const UNIT = { 十: 10, 拾: 10, 百: 100, 佰: 100, 千: 1000, 仟: 1000 };
const BIG = { 萬: 1e4, 万: 1e4, 億: 1e8, 亿: 1e8 };
const CN = Object.keys(DIGIT).join('') + Object.keys(UNIT).join('') + Object.keys(BIG).join('');
const CN_UNIT_RE = new RegExp(`[${Object.keys(UNIT).join('')}${Object.keys(BIG).join('')}]`);

const CURRENCY_WORDS = [
  [/新臺幣|新台幣|臺幣|台幣|NTD|TWD|NT\$/i, 'TWD'],
  [/美金|美元|USD|US\$/i, 'USD'],
  [/日幣|日圓|日元|円|JPY/i, 'JPY'],
  [/人民幣|RMB|CNY/i, 'CNY'],
  [/歐元|EUR|€/i, 'EUR'],
  [/港幣|港元|HKD/i, 'HKD'],
  [/韓元|韓幣|韓圜|KRW/i, 'KRW'],
  [/英鎊|GBP|£/i, 'GBP'],
];

/** 全形數字與符號轉半形 */
export function toHalfWidth(s) {
  return String(s).replace(/[０-９．，＋：／]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0));
}

/** 中文（或中阿混合）數字轉數值：一百二→120、兩萬三→23000、三千五百零五→3505、十二點五→12.5、3千5→3500 */
export function cnToNumber(input) {
  const s = toHalfWidth(input).replace(/[,\s]/g, '');
  if (!s) return NaN;
  if (/^\d+(\.\d+)?$/.test(s)) return parseFloat(s);
  const parts = s.split(/[點点.]/);
  if (parts.length > 2) return NaN;
  const intVal = parts[0] ? parseCnInt(parts[0]) : 0;
  if (Number.isNaN(intVal)) return NaN;
  if (parts.length === 1) return intVal;
  let dec = '';
  for (const c of parts[1]) {
    if (/\d/.test(c)) dec += c;
    else if (c in DIGIT) dec += DIGIT[c];
    else return NaN;
  }
  return dec ? parseFloat(`${intVal}.${dec}`) : intVal;
}

function parseCnInt(s) {
  // 沒有任何單位字 → 逐位讀法，例如「一二零」= 120
  if (!CN_UNIT_RE.test(s)) {
    let out = '';
    for (const c of s) {
      if (/\d/.test(c)) out += c;
      else if (c in DIGIT) out += DIGIT[c];
      else return NaN;
    }
    return out ? parseInt(out, 10) : NaN;
  }
  let total = 0, section = 0, num = null, lastUnit = 0, zero = false;
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (/\d/.test(c)) {
      let j = i;
      while (j < s.length && /\d/.test(s[j])) j++;
      num = parseInt(s.slice(i, j), 10);
      i = j - 1;
      continue;
    }
    if (c in DIGIT) {
      const d = DIGIT[c];
      if (d === 0) { zero = true; num = null; continue; }
      num = d;
      continue;
    }
    if (c in UNIT) {
      section += (num ?? 1) * UNIT[c];
      lastUnit = UNIT[c]; num = null; zero = false;
      continue;
    }
    if (c in BIG) {
      section += num ?? 0;
      if (section === 0 && total === 0) section = 1;
      if (BIG[c] === 1e8) total = (total + section) * 1e8;
      else total += section * BIG[c];
      section = 0; lastUnit = BIG[c]; num = null; zero = false;
      continue;
    }
    return NaN;
  }
  if (num != null) {
    // 口語省略：「一百二」= 120、「兩萬三」= 23000；有「零」時不套用（一百零二 = 102）
    if (lastUnit >= 10 && !zero) section += num * (lastUnit / 10);
    else section += num;
  }
  return total + section;
}

/** 從文字找出幣別 */
export function detectCurrency(text) {
  for (const [re, code] of CURRENCY_WORDS) if (re.test(text)) return code;
  return null;
}

function stripCurrencyWords(text) {
  let t = text;
  for (const [re] of CURRENCY_WORDS) t = t.replace(new RegExp(re.source, 'gi'), ' ');
  return t;
}

const NUM_TOKEN = `\\d[\\d,]*(?:\\.\\d+)?|[${CN}\\d]+(?:[點点][${CN}\\d]+)?`;
const MONEY_RE = new RegExp(
  `(${NUM_TOKEN})\\s*(?:元|塊錢|塊|圓|块钱|块|\\$)(?:([\\d${Object.keys(DIGIT).join('')}])(?:毛|角|(?=$|[\\s，,。])))?`, 'g');

function lastMatch(re, text) {
  re.lastIndex = 0;
  let m, last = null;
  while ((m = re.exec(text))) { last = m; if (m[0] === '') re.lastIndex++; }
  return last;
}

/**
 * 從一段話抽出金額，回傳 { amount, currency, rest }
 * rest = 去掉金額與幣別字詞後剩下的文字（可當作「項目」）
 */
export function extractAmount(raw) {
  const text = toHalfWidth(String(raw || '')).trim();
  const currency = detectCurrency(text);
  let t = stripCurrencyWords(text);
  let amount = NaN, span = null;

  // 1) 帶「元／塊」的金額，支援「十二塊五」= 12.5
  let m = lastMatch(MONEY_RE, t);
  if (m) {
    const main = cnToNumber(m[1].replace(/,/g, ''));
    if (!Number.isNaN(main)) {
      const cents = m[2] != null ? (/\d/.test(m[2]) ? +m[2] : DIGIT[m[2]]) : 0;
      amount = main + cents / 10;
      span = [m.index, m.index + m[0].length];
    }
  }
  // 2) 最後一個阿拉伯數字
  if (span == null) {
    m = lastMatch(/\d[\d,]*(?:\.\d+)?/g, t);
    if (m) { amount = parseFloat(m[0].replace(/,/g, '')); span = [m.index, m.index + m[0].length]; }
  }
  // 3) 最後一段含單位（十百千萬）的中文數字；單一「三」之類不算，避免誤抓「三明治」
  if (span == null) {
    const re = new RegExp(`[${CN}]+(?:[點点][${CN}]+)?`, 'g');
    let mm, best = null;
    while ((mm = re.exec(t))) {
      const whole = t.trim() === mm[0];
      if (whole || CN_UNIT_RE.test(mm[0])) {
        const v = cnToNumber(mm[0]);
        if (!Number.isNaN(v)) best = { v, idx: mm.index, len: mm[0].length };
      }
    }
    if (best) { amount = best.v; span = [best.idx, best.idx + best.len]; }
  }

  let rest = t;
  if (span) rest = t.slice(0, span[0]) + ' ' + t.slice(span[1]);
  rest = rest.replace(/[，,。.!！?？、]+/g, ' ').replace(/\s+/g, ' ').trim();
  if (!Number.isNaN(amount)) amount = Math.round(amount * 100) / 100;
  return { amount, currency, rest };
}

/** 價格／金額欄位的語音 → 數值 */
export function parseSpokenAmount(text) {
  return extractAmount(text).amount;
}

/* ============================================================
 * 日期與時間
 * ============================================================ */
const pad2 = (n) => String(n).padStart(2, '0');
const ymd = (d) => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
const NUM_S = `[\\d${Object.keys(DIGIT).join('')}十拾]{1,4}`; // 日期、時間用的小數字
const WEEKDAY = { 一: 1, 二: 2, 三: 3, 四: 4, 五: 5, 六: 6, 日: 0, 天: 0 };
const PERIOD = '凌晨|半夜|清晨|早上|上午|中午|下午|傍晚|晚上|今早|今晚|昨晚|昨天晚上';

function validDate(y, m, d) {
  const dt = new Date(y, m - 1, d);
  return dt.getFullYear() === y && dt.getMonth() === m - 1 && dt.getDate() === d ? dt : null;
}

/**
 * 從一段話抽出日期與時間。
 * 回傳 { date: 'YYYY-MM-DD' | null, time: 'HH:MM' | null, rest }
 * 支援：今天／昨天／前天／大前天／明天、上週五／星期三／禮拜天、11月20號、十一月二十日、
 *       2026年11月20日、2026/11/20、11/20、20號；
 *       早上8點、下午三點半、晚上7點15分、中午12點、19:30、昨晚8點
 */
export function extractDateTime(raw, now = new Date()) {
  let t = toHalfWidth(String(raw || ''));
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  let date = null, time = null;
  const cut = (m) => { t = t.slice(0, m.index) + ' ' + t.slice(m.index + m[0].length); };
  const n = (s) => cnToNumber(s);
  // 沒講年份時：若日期落在未來，視為去年（例：一月初補記十二月的帳）
  const pastYear = (m, d) => {
    let dt = validDate(today.getFullYear(), m, d);
    if (dt && dt > today) dt = validDate(today.getFullYear() - 1, m, d);
    return dt;
  };

  // ---- 日期 ----
  const datePatterns = [
    [/(\d{4})\s*[年/.-]\s*(\d{1,2})\s*[月/.-]\s*(\d{1,2})\s*[日號号]?/, (m) => validDate(+m[1], +m[2], +m[3])],
    [new RegExp(`(${NUM_S})\\s*月\\s*(${NUM_S})\\s*[日號号]?`), (m) => pastYear(n(m[1]), n(m[2]))],
    [/(?<![\d.])(\d{1,2})\/(\d{1,2})(?![\d/])/, (m) => pastYear(+m[1], +m[2])],
    [/(大前天|前天|昨天|昨日|昨晚|今天|今日|今早|今晚|明天|明日)/, (m) => {
      const off = { 大前天: -3, 前天: -2, 昨天: -1, 昨日: -1, 昨晚: -1, 今天: 0, 今日: 0, 今早: 0, 今晚: 0, 明天: 1, 明日: 1 }[m[1]];
      const d = new Date(today); d.setDate(d.getDate() + off); return d;
    }],
    [/(上上|上個?|這個?|这个?|本)?\s*(?:週|周|星期|禮拜|礼拜)\s*([一二三四五六日天])/, (m) => {
      const target = WEEKDAY[m[2]];
      const mondayIdx = (today.getDay() + 6) % 7; // 今天是本週第幾天（週一=0）
      const targetIdx = (target + 6) % 7;
      const d = new Date(today);
      if (!m[1]) { // 沒講「上／這」：取最近一次（不含未來）
        let back = mondayIdx - targetIdx;
        if (back < 0) back += 7;
        d.setDate(d.getDate() - back);
      } else {
        const weeks = m[1].startsWith('上上') ? 2 : m[1].startsWith('上') ? 1 : 0;
        d.setDate(d.getDate() - mondayIdx + targetIdx - 7 * weeks);
      }
      return d;
    }],
    [new RegExp(`(?<![月\\d])(${NUM_S})\\s*[號号日](?!本|幣)`), (m) => {
      const day = n(m[1]);
      let dt = validDate(today.getFullYear(), today.getMonth() + 1, day);
      if (dt && dt > today) { const p = new Date(today.getFullYear(), today.getMonth() - 1, 1); dt = validDate(p.getFullYear(), p.getMonth() + 1, day); }
      return dt;
    }],
  ];
  for (const [re, fn] of datePatterns) {
    const m = t.match(re);
    if (!m) continue;
    const d = fn(m);
    if (d) {
      date = ymd(d);
      // 「昨晚」「今晚」同時帶有時段意義，保留給時間判斷用
      if (/^(昨晚|今晚|今早)$/.test(m[1])) t = t.replace(m[0], m[1] === '今早' ? ' 早上 ' : ' 晚上 ');
      else cut(m);
      break;
    }
  }

  // ---- 時間 ----
  const timeRe = new RegExp(
    `(${PERIOD})?\\s*(?:(\\d{1,2}):(\\d{2})|(${NUM_S})\\s*(?:點鐘|點|点|時|时)\\s*(?:(半)|(${NUM_S})\\s*分?)?)(?!\\s*(?:元|塊|块|圓|${NUM_S}))`);
  const m = t.match(timeRe);
  if (m) {
    let h = m[2] != null ? +m[2] : n(m[4]);
    let mi = m[3] != null ? +m[3] : m[5] ? 30 : m[6] ? n(m[6]) : 0;
    const p = m[1] || '';
    if (/下午|傍晚|晚上|今晚|昨晚/.test(p) && h < 12) h += 12;
    else if (/中午/.test(p) && h < 6) h += 12; // 中午一點 = 13:00
    else if (/凌晨|半夜/.test(p) && h === 12) h = 0;
    if (Number.isInteger(h) && h >= 0 && h <= 24 && Number.isInteger(mi) && mi >= 0 && mi < 60) {
      time = `${pad2(h % 24)}:${pad2(mi)}`;
      cut(m);
    }
  }
  // 只講時段（如「晚上 晚餐 拉麵」）時，把時段字詞拿掉，不設時間
  t = t.replace(new RegExp(`(${PERIOD})`, 'g'), ' ');
  return { date, time, rest: t.replace(/\s+/g, ' ').trim() };
}

/**
 * 一句話記帳：「昨天晚上七點 晚餐 牛肉麵 150 元」
 * → { date:'2026-11-25', time:'19:00', category:'晚餐', item:'牛肉麵', amount:150, currency:null }
 * categories：現有類別清單，會挑出文字中出現的最長類別名稱
 */
export function parseQuickEntry(text, categories = [], now = new Date()) {
  const dt = extractDateTime(text, now);
  const { amount, currency, rest } = extractAmount(dt.rest);
  let category = null;
  for (const c of [...categories].sort((a, b) => b.length - a.length)) {
    if (c && rest.includes(c)) { category = c; break; }
  }
  let item = rest;
  if (category) {
    const without = rest.replace(category, ' ').replace(/\s+/g, ' ').trim();
    item = without || category;
  }
  return { date: dt.date, time: dt.time, category, item, amount, currency };
}
