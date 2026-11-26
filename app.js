import { parseSpokenAmount, parseQuickEntry, detectCurrency } from './voice-parse.js';

/* ============================================================
 * 常數與狀態
 * ============================================================ */
const LS = {
  config: 'ledger.config.v1',
  cache: 'ledger.cache.v1',
  demo: 'ledger.demo.v1',
  queue: 'ledger.queue.v1',
  prefs: 'ledger.prefs.v1',
};
const TYPES = {
  expense: { label: '支出', amountLabel: '價格', itemPh: '例如：牛肉麵、捷運' },
  income: { label: '收入', amountLabel: '金額', itemPh: '例如：11 月薪資' },
};
const WEEK = ['日', '一', '二', '三', '四', '五', '六'];

const defaultData = () => ({
  expense: [],
  income: [],
  categories: {
    expense: ['早餐', '午餐', '晚餐', '飲料', '交通', '購物', '日用品', '娛樂', '醫療', '住房', '水電瓦斯', '通訊', '教育', '其他'],
    income: ['薪資', '獎金', '投資', '副業', '利息', '退款', '其他'],
  },
  settings: { baseCurrency: 'TWD', rates: { TWD: 1, USD: 31, JPY: 0.2, EUR: 34, CNY: 4.3, HKD: 4 } },
  sheetUrl: '',
});

const store = {
  get(k, fallback) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : fallback; } catch { return fallback; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { console.warn('localStorage 寫入失敗', e); } },
  del(k) { try { localStorage.removeItem(k); } catch { /* ignore */ } },
};

let config = store.get(LS.config, { mode: '', apiUrl: '', token: '', lang: 'zh-TW' });
if (!config.mode) config.mode = config.apiUrl ? 'sheet' : 'demo';
let prefs = store.get(LS.prefs, { lastCurrency: {}, lastCategory: {} });
let data = loadLocalData();
let queue = store.get(LS.queue, []);
let flushing = false;
let syncError = '';
const forms = {};
const stats = { gran: 'month', anchor: todayStr(), from: '', to: '' };

function loadLocalData() {
  const key = config.mode === 'sheet' ? LS.cache : LS.demo;
  return normalizeData(store.get(key, null) || defaultData());
}
function persist() { store.set(config.mode === 'sheet' ? LS.cache : LS.demo, data); }
function normalizeData(d) {
  const base = defaultData();
  const out = {
    expense: Array.isArray(d.expense) ? d.expense : [],
    income: Array.isArray(d.income) ? d.income : [],
    categories: {
      expense: d.categories?.expense?.length ? d.categories.expense : base.categories.expense,
      income: d.categories?.income?.length ? d.categories.income : base.categories.income,
    },
    settings: d.settings?.rates ? d.settings : base.settings,
    sheetUrl: d.sheetUrl || '',
  };
  out.settings.rates[out.settings.baseCurrency] = 1;
  return out;
}

/* ============================================================
 * 小工具
 * ============================================================ */
const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];
function pad(n) { return String(n).padStart(2, '0'); }
function dateStr(d) { return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; }
function todayStr() { return dateStr(new Date()); }
function nowTime() { const d = new Date(); return `${pad(d.getHours())}:${pad(d.getMinutes())}`; }
function parseDate(s) { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); }
function addDays(s, n) { const d = parseDate(s); d.setDate(d.getDate() + n); return dateStr(d); }
function uid() {
  if (crypto.randomUUID) return crypto.randomUUID();
  return 'id-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 10);
}
function esc(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
const fmtCache = {};
function money(n, cur) {
  const key = cur || 'X';
  if (!fmtCache[key]) {
    try {
      fmtCache[key] = new Intl.NumberFormat('zh-TW', { style: 'currency', currency: cur, currencyDisplay: 'symbol', minimumFractionDigits: 0, maximumFractionDigits: 2 });
    } catch {
      fmtCache[key] = { format: (v) => `${new Intl.NumberFormat('zh-TW', { maximumFractionDigits: 2 }).format(v)} ${cur}` };
    }
  }
  return fmtCache[key].format(n);
}
function num(n) { return new Intl.NumberFormat('zh-TW', { maximumFractionDigits: 2 }).format(n); }
function sortRecs(list) {
  return [...list].sort((a, b) => (b.date + b.time + (b.createdAt || '')).localeCompare(a.date + a.time + (a.createdAt || '')));
}
function currencies() {
  const s = data.settings;
  const list = Object.keys(s.rates);
  if (!list.includes(s.baseCurrency)) list.unshift(s.baseCurrency);
  return [s.baseCurrency, ...list.filter((c) => c !== s.baseCurrency).sort()];
}
/** 換算成基準幣；無匯率時回傳 NaN */
function toBase(amount, cur) {
  const s = data.settings;
  if (cur === s.baseCurrency) return amount;
  const r = s.rates[cur];
  return r ? amount * r : NaN;
}

let toastTimer;
function toast(msg, kind = '') {
  const el = $('#toast');
  el.textContent = msg;
  el.className = `toast show ${kind}`;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { el.className = 'toast'; }, kind === 'error' ? 5000 : 2600);
}

/* ============================================================
 * 後端（Google Apps Script）與同步佇列
 * ============================================================ */
class ApiError extends Error {}

async function api(action, payload = {}) {
  if (!config.apiUrl) throw new ApiError('尚未設定 Apps Script 網址');
  // 以 text/plain 傳送 JSON（簡單請求，不會觸發 CORS 預檢，Apps Script 才收得到）
  const res = await fetch(config.apiUrl, {
    method: 'POST',
    body: JSON.stringify({ action, token: config.token, ...payload }),
    redirect: 'follow',
  });
  if (!res.ok) throw new ApiError(`伺服器回應 HTTP ${res.status}`);
  let j;
  try { j = await res.json(); } catch { throw new ApiError('回應不是 JSON：請確認部署為「網頁應用程式」且存取權為「所有人」'); }
  if (!j.ok) throw new ApiError(j.error || '未知錯誤');
  return j.data;
}

/** 所有資料異動都經過這裡：先更新畫面（樂觀更新），再排入佇列同步到試算表 */
async function mutate(action, payload) {
  applyOp(data, { action, payload });
  persist();
  renderAll();
  if (config.mode !== 'sheet') return;
  queue.push({ action, payload, at: Date.now() });
  store.set(LS.queue, queue);
  await flushQueue();
}

function applyOp(d, { action, payload: p }) {
  switch (action) {
    case 'add':
      if (!d[p.type].some((r) => r.id === p.record.id)) d[p.type].push(p.record);
      if (!d.categories[p.type].includes(p.record.category)) d.categories[p.type].push(p.record.category);
      break;
    case 'update': {
      const i = d[p.type].findIndex((r) => r.id === p.record.id);
      if (i >= 0) d[p.type][i] = p.record;
      break;
    }
    case 'delete':
      d[p.type] = d[p.type].filter((r) => r.id !== p.id);
      break;
    case 'addCategory':
      if (!d.categories[p.type].includes(p.name)) d.categories[p.type].push(p.name);
      break;
    case 'deleteCategory':
      d.categories[p.type] = d.categories[p.type].filter((c) => c !== p.name);
      break;
    case 'saveSettings':
      d.settings = structuredClone(p.settings);
      break;
    case 'importRecords':
      for (const r of p.records) {
        const { type, ...rec } = r;
        if (!d[type].some((x) => x.id === rec.id)) d[type].push(rec);
        if (!d.categories[type].includes(rec.category)) d.categories[type].push(rec.category);
      }
      break;
    default:
      console.warn('未知動作', action);
  }
}

async function flushQueue() {
  if (flushing || config.mode !== 'sheet') return updateStatus();
  flushing = true;
  updateStatus();
  let needReload = false;
  try {
    while (queue.length) {
      const op = queue[0];
      try {
        const result = await api(op.action, op.payload);
        if (op.action === 'importRecords' && result) {
          toast(`已匯入 ${result.added} 筆，略過重複 ${result.skipped} 筆${result.errors?.length ? `，${result.errors.length} 筆有誤` : ''}`, 'ok');
        }
        queue.shift();
        store.set(LS.queue, queue);
        syncError = '';
      } catch (e) {
        if (e instanceof ApiError) {
          // 伺服器明確拒絕（例如 Token 錯誤、資料格式錯）：丟棄此動作並以伺服器資料為準
          queue.shift();
          store.set(LS.queue, queue);
          syncError = e.message;
          toast(`同步失敗：${e.message}`, 'error');
          needReload = !/Token/.test(e.message);
          if (/Token/.test(e.message)) break;
        } else {
          syncError = '';
          break; // 網路問題：保留佇列，等恢復連線再送
        }
      }
    }
  } finally {
    flushing = false;
    updateStatus();
  }
  if (needReload) await reloadFromServer({ silent: true });
}

async function reloadFromServer({ silent = false } = {}) {
  if (config.mode !== 'sheet') { renderAll(); return; }
  if (queue.length) await flushQueue();
  if (queue.length) { updateStatus(); return; } // 仍離線，保留本機資料
  setStatus('同步中…');
  try {
    const d = await api('load');
    data = normalizeData(d);
    persist();
    syncError = '';
    renderAll();
    if (!silent) toast('已從試算表更新', 'ok');
  } catch (e) {
    syncError = e instanceof ApiError ? e.message : '';
    if (!silent || e instanceof ApiError) toast(e instanceof ApiError ? `讀取失敗：${e.message}` : '無法連線，先顯示本機資料', e instanceof ApiError ? 'error' : '');
  }
  updateStatus();
}

function setStatus(text, cls = '') {
  const el = $('#syncStatus');
  el.textContent = text;
  el.className = `status ${cls}`;
}
function updateStatus() {
  if (config.mode !== 'sheet') return setStatus('示範模式', 'warn');
  if (!config.apiUrl || !config.token) return setStatus('尚未設定', 'err');
  if (flushing) return setStatus('同步中…');
  if (syncError) return setStatus('同步錯誤', 'err');
  if (queue.length) return setStatus(navigator.onLine ? `${queue.length} 筆待同步` : `離線・${queue.length} 筆待同步`, 'warn');
  return setStatus(navigator.onLine ? '✓ 已同步' : '離線');
}

/* ============================================================
 * 語音輸入（Web Speech API）
 * ============================================================ */
const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
let activeRec = null;

function listen() {
  return new Promise((resolve, reject) => {
    if (!SR) return reject(new Error('此瀏覽器不支援語音辨識，請改用鍵盤上的麥克風（聽寫）輸入'));
    if (activeRec) { try { activeRec.abort(); } catch { /* ignore */ } }
    const rec = new SR();
    activeRec = rec;
    rec.lang = config.lang || 'zh-TW';
    rec.interimResults = true;
    rec.maxAlternatives = 1;
    rec.continuous = false;
    let finalText = '';
    let done = false;
    const overlay = $('#voiceOverlay');
    const hint = $('#voiceHint');
    hint.textContent = '請說話…';
    overlay.classList.remove('hidden');
    const finish = (fn, v) => {
      if (done) return;
      done = true;
      overlay.classList.add('hidden');
      activeRec = null;
      fn(v);
    };
    rec.onresult = (e) => {
      let interim = '';
      for (let i = e.resultIndex; i < e.results.length; i++) {
        const t = e.results[i][0].transcript;
        if (e.results[i].isFinal) finalText += t; else interim += t;
      }
      hint.textContent = finalText + interim || '請說話…';
    };
    rec.onerror = (e) => {
      const map = {
        'not-allowed': '麥克風權限被拒絕，請在瀏覽器設定允許使用麥克風',
        'service-not-allowed': '語音服務無法使用（請確認網址是 https）',
        'no-speech': '沒有聽到聲音，請再試一次',
        'audio-capture': '找不到麥克風',
        network: '語音辨識需要網路連線',
        aborted: '',
      };
      const msg = map[e.error] ?? `語音辨識錯誤：${e.error}`;
      finish(msg ? reject : resolve, msg ? new Error(msg) : '');
    };
    rec.onend = () => finish(resolve, finalText.trim());
    $('#voiceStop').onclick = () => { try { rec.stop(); } catch { /* ignore */ } };
    try { rec.start(); } catch (err) { finish(reject, err); }
  });
}

async function voiceInto(type, field, btn) {
  const f = forms[type];
  btn.classList.add('listening');
  try {
    const text = await listen();
    if (!text) return;
    if (field === 'item') {
      f.form.item.value = text;
    } else {
      const v = parseSpokenAmount(text);
      if (Number.isNaN(v)) { toast(`聽到「${text}」，但找不到數字`, 'error'); return; }
      f.form.amount.value = v;
      const cur = detectCurrency(text);
      if (cur) setCurrency(type, cur);
    }
  } catch (e) {
    toast(e.message, 'error');
    f.form[field].focus(); // 語音 API 不能用時，直接用鍵盤上的 🎤 聽寫
  } finally {
    btn.classList.remove('listening');
  }
}

async function quickVoice(type) {
  try {
    const text = await listen();
    if (!text) return;
    const r = parseQuickEntry(text, data.categories[type]);
    const f = forms[type].form;
    if (r.category) f.category.value = r.category;
    if (r.item) f.item.value = r.item;
    if (!Number.isNaN(r.amount)) f.amount.value = r.amount;
    if (r.currency) setCurrency(type, r.currency);
    toast(`辨識：「${text}」，確認後按儲存`);
    (Number.isNaN(r.amount) ? f.amount : forms[type].submit).focus();
  } catch (e) {
    toast(e.message, 'error');
  }
}

function setCurrency(type, cur) {
  const sel = forms[type].form.currency;
  if (![...sel.options].some((o) => o.value === cur)) {
    fillCurrencySelect(type, cur);
    toast(`尚未設定 ${cur} 匯率，統計時不會換算；可到 ⚙️ 設定加入`);
  }
  sel.value = cur;
}

/* ============================================================
 * 支出／收入表單
 * ============================================================ */
function buildForms() {
  const tpl = $('#entryTpl');
  for (const type of Object.keys(TYPES)) {
    const sec = $(`#tab-${type}`);
    sec.appendChild(tpl.content.cloneNode(true));
    const form = $('form', sec);
    form.dataset.type = type;
    const f = {
      sec, form,
      title: $('[data-role=title]', sec),
      submit: $('[data-role=submit]', sec),
      cancelEdit: $('[data-role=cancelEdit]', sec),
      recent: $('[data-role=recent]', sec),
      todaySum: $('[data-role=todaySum]', sec),
      newCatBox: $('[data-role=newCatBox]', sec),
      newCatInput: $('[data-role=newCatInput]', sec),
      editingId: null,
    };
    forms[type] = f;
    f.title.textContent = `記一筆${TYPES[type].label}`;
    $('[data-role=amountLabel]', sec).textContent = TYPES[type].amountLabel;
    form.item.placeholder = TYPES[type].itemPh;
    resetForm(type, true);

    $('[data-role=addCat]', sec).onclick = () => {
      f.newCatBox.classList.toggle('hidden');
      if (!f.newCatBox.classList.contains('hidden')) f.newCatInput.focus();
    };
    $('[data-role=newCatCancel]', sec).onclick = () => { f.newCatBox.classList.add('hidden'); f.newCatInput.value = ''; };
    const addCat = async () => {
      const name = f.newCatInput.value.trim();
      if (!name) return f.newCatInput.focus();
      if (data.categories[type].includes(name)) { form.category.value = name; f.newCatBox.classList.add('hidden'); return; }
      f.newCatInput.value = '';
      f.newCatBox.classList.add('hidden');
      await mutate('addCategory', { type, name });
      form.category.value = name;
      toast(`已新增類別「${name}」`, 'ok');
    };
    $('[data-role=newCatOk]', sec).onclick = addCat;
    f.newCatInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); addCat(); } });

    $$('[data-voice]', sec).forEach((btn) => { btn.onclick = () => voiceInto(type, btn.dataset.voice, btn); });
    $('[data-role=quickVoice]', sec).onclick = () => quickVoice(type);
    f.cancelEdit.onclick = () => resetForm(type);
    form.addEventListener('submit', (e) => { e.preventDefault(); submitForm(type); });

    f.recent.addEventListener('click', (e) => onListClick(e, type));
  }
}

function resetForm(type, initial = false) {
  const f = forms[type];
  const form = f.form;
  const keepDate = !initial && form.date.value;
  form.reset();
  form.date.value = keepDate || todayStr();
  form.time.value = nowTime();
  fillCategorySelect(type);
  fillCurrencySelect(type);
  f.editingId = null;
  form.classList.remove('editing');
  f.submit.textContent = '儲存';
  f.cancelEdit.classList.add('hidden');
}

async function submitForm(type) {
  const f = forms[type];
  const form = f.form;
  const amount = Number(String(form.amount.value).replace(/[,\s]/g, ''));
  const item = form.item.value.trim();
  if (!form.date.value) return toast('請選擇日期', 'error');
  if (!form.category.value) return toast('請選擇類別', 'error');
  if (!item) { form.item.focus(); return toast('請輸入項目', 'error'); }
  if (!Number.isFinite(amount) || amount <= 0) { form.amount.focus(); return toast(`請輸入正確的${TYPES[type].amountLabel}`, 'error'); }

  const old = f.editingId ? data[type].find((r) => r.id === f.editingId) : null;
  const record = {
    id: old?.id || uid(),
    date: form.date.value,
    time: form.time.value || '00:00',
    category: form.category.value,
    item,
    amount: Math.round(amount * 100) / 100,
    currency: form.currency.value,
    note: form.note.value.trim(),
    createdAt: old?.createdAt || new Date().toISOString(),
  };
  prefs.lastCurrency[type] = record.currency;
  prefs.lastCategory[type] = record.category;
  store.set(LS.prefs, prefs);

  const wasEdit = !!old;
  const wasDate = form.date.value;
  resetForm(type);
  if (wasDate !== todayStr()) form.date.value = wasDate; // 補記過去的帳時，保留日期方便連續輸入
  form.category.value = record.category;
  form.currency.value = record.currency;
  await mutate(wasEdit ? 'update' : 'add', { type, record });
  toast(`${wasEdit ? '已更新' : '已記錄'}：${record.item} ${money(record.amount, record.currency)}`, 'ok');
  form.item.focus();
}

function startEdit(type, id) {
  const r = data[type].find((x) => x.id === id);
  if (!r) return;
  const f = forms[type];
  showTab(type);
  fillCategorySelect(type, r.category);
  fillCurrencySelect(type, r.currency);
  const form = f.form;
  form.date.value = r.date;
  form.time.value = r.time;
  form.category.value = r.category;
  form.item.value = r.item;
  form.amount.value = r.amount;
  form.currency.value = r.currency;
  form.note.value = r.note || '';
  f.editingId = id;
  form.classList.add('editing');
  f.submit.textContent = '更新';
  f.cancelEdit.classList.remove('hidden');
  form.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

async function deleteRecord(type, id) {
  const r = data[type].find((x) => x.id === id);
  if (!r) return;
  if (!confirm(`確定刪除這筆${TYPES[type].label}？\n${r.date} ${r.item} ${money(r.amount, r.currency)}`)) return;
  if (forms[type].editingId === id) resetForm(type);
  await mutate('delete', { type, id });
  toast('已刪除');
}

function onListClick(e, fallbackType) {
  const li = e.target.closest('li[data-id]');
  if (!li) return;
  const type = li.dataset.type || fallbackType;
  if (e.target.closest('.rec-del')) deleteRecord(type, li.dataset.id);
  else startEdit(type, li.dataset.id);
}

function fillCategorySelect(type, ensure) {
  const sel = forms[type].form.category;
  const cur = sel.value || prefs.lastCategory[type];
  const cats = [...data.categories[type]];
  if (ensure && !cats.includes(ensure)) cats.push(ensure);
  sel.innerHTML = cats.map((c) => `<option>${esc(c)}</option>`).join('');
  if (cur && cats.includes(cur)) sel.value = cur;
}
function fillCurrencySelect(type, ensure) {
  const sel = forms[type].form.currency;
  const cur = sel.value || prefs.lastCurrency[type] || data.settings.baseCurrency;
  const list = currencies();
  if (ensure && !list.includes(ensure)) list.push(ensure);
  sel.innerHTML = list.map((c) => `<option>${esc(c)}</option>`).join('');
  sel.value = list.includes(cur) ? cur : data.settings.baseCurrency;
}
function fillCurrencySelects() { Object.keys(TYPES).forEach((t) => fillCurrencySelect(t)); }

function recItemHTML(r, type, { showType = false, pending = false } = {}) {
  const sign = type === 'expense' ? '−' : '+';
  const d = parseDate(r.date);
  return `<li data-id="${esc(r.id)}" data-type="${type}">
    <div class="rec-main">
      <div class="rec-title">${showType ? `<span class="muted">${TYPES[type].label}・</span>` : ''}${esc(r.item)}${pending ? '<span class="pending-dot" title="尚未同步"></span>' : ''}</div>
      <div class="rec-sub">${d.getMonth() + 1}/${d.getDate()}（${WEEK[d.getDay()]}）${esc(r.time)}・${esc(r.category)}${r.note ? `・${esc(r.note)}` : ''}</div>
    </div>
    <div class="rec-amt ${type}">${sign}${money(r.amount, r.currency)}</div>
    <button type="button" class="rec-del" aria-label="刪除">✕</button>
  </li>`;
}

function pendingIds() {
  const s = new Set();
  for (const op of queue) {
    if (op.payload?.record?.id) s.add(op.payload.record.id);
  }
  return s;
}

function renderRecent(type) {
  const f = forms[type];
  const list = sortRecs(data[type]).slice(0, 30);
  const pend = pendingIds();
  f.recent.innerHTML = list.length
    ? list.map((r) => recItemHTML(r, type, { pending: pend.has(r.id) })).join('')
    : `<li class="empty">還沒有${TYPES[type].label}記錄，試試上方的 🎤 一句話記帳</li>`;
  const today = todayStr();
  const month = today.slice(0, 7);
  let t = 0, m = 0;
  for (const r of data[type]) {
    const v = toBase(r.amount, r.currency);
    if (Number.isNaN(v)) continue;
    if (r.date === today) t += v;
    if (r.date.startsWith(month)) m += v;
  }
  const b = data.settings.baseCurrency;
  f.todaySum.textContent = `今日 ${money(t, b)}・本月 ${money(m, b)}`;
}

/* ============================================================
 * 統計
 * ============================================================ */
function periodRange() {
  const a = stats.anchor;
  const d = parseDate(a);
  switch (stats.gran) {
    case 'day': return [a, a];
    case 'week': {
      const dow = (d.getDay() + 6) % 7; // 週一為一週開始
      const start = addDays(a, -dow);
      return [start, addDays(start, 6)];
    }
    case 'month': {
      const s = new Date(d.getFullYear(), d.getMonth(), 1);
      const e = new Date(d.getFullYear(), d.getMonth() + 1, 0);
      return [dateStr(s), dateStr(e)];
    }
    case 'year': return [`${d.getFullYear()}-01-01`, `${d.getFullYear()}-12-31`];
    case 'custom': {
      let from = stats.from || addDays(todayStr(), -29);
      let to = stats.to || todayStr();
      if (from > to) [from, to] = [to, from];
      return [from, to];
    }
    default: return [a, a];
  }
}
function periodLabel([from, to]) {
  const d = parseDate(from);
  switch (stats.gran) {
    case 'day': return `${from}（${WEEK[d.getDay()]}）`;
    case 'week': return `${from} ～ ${to.slice(5)}`;
    case 'month': return `${d.getFullYear()} 年 ${d.getMonth() + 1} 月`;
    case 'year': return `${d.getFullYear()} 年`;
    default: return `${from} ～ ${to}`;
  }
}
function shiftPeriod(dir) {
  const d = parseDate(stats.anchor);
  if (stats.gran === 'day') d.setDate(d.getDate() + dir);
  else if (stats.gran === 'week') d.setDate(d.getDate() + 7 * dir);
  else if (stats.gran === 'month') { d.setDate(1); d.setMonth(d.getMonth() + dir); }
  else if (stats.gran === 'year') d.setFullYear(d.getFullYear() + dir);
  stats.anchor = dateStr(d);
  renderStats();
}

function recordsIn([from, to]) {
  const pick = (type) => data[type].filter((r) => r.date >= from && r.date <= to).map((r) => ({ ...r, type }));
  return { expense: pick('expense'), income: pick('income') };
}

function renderStats() {
  const range = periodRange();
  const [from, to] = range;
  $('#periodLabel').textContent = periodLabel(range);
  $$('#granSeg button').forEach((b) => b.classList.toggle('active', b.dataset.gran === stats.gran));
  $('#customRange').classList.toggle('hidden', stats.gran !== 'custom');
  $('#prevPeriod').disabled = $('#nextPeriod').disabled = stats.gran === 'custom';
  if (stats.gran === 'custom') { $('#rangeFrom').value = from; $('#rangeTo').value = to; }

  const base = data.settings.baseCurrency;
  const recs = recordsIn(range);
  const missing = new Set();
  const sum = (list) => list.reduce((s, r) => {
    const v = toBase(r.amount, r.currency);
    if (Number.isNaN(v)) { missing.add(r.currency); return s; }
    return s + v;
  }, 0);
  const inc = sum(recs.income);
  const exp = sum(recs.expense);
  const days = Math.round((parseDate(to) - parseDate(from)) / 86400000) + 1;
  const elapsed = Math.min(days, Math.max(1, Math.round((parseDate(todayStr()) - parseDate(from)) / 86400000) + 1));
  const avgDays = to >= todayStr() && from <= todayStr() ? elapsed : days;

  $('#summary').innerHTML = `
    <div class="sum-card income"><div class="k">收入</div><div class="v">${money(inc, base)}</div></div>
    <div class="sum-card expense"><div class="k">支出</div><div class="v">${money(exp, base)}</div></div>
    <div class="sum-card"><div class="k">結餘</div><div class="v" style="color:${inc - exp >= 0 ? 'var(--income)' : 'var(--expense)'}">${money(inc - exp, base)}</div></div>
    <div class="sum-card"><div class="k">${stats.gran === 'day' ? '支出筆數' : '日均支出'}</div><div class="v">${stats.gran === 'day' ? recs.expense.length + ' 筆' : money(exp / avgDays, base)}</div></div>`;
  $('#rateNote').textContent = (missing.size ? `⚠️ ${[...missing].join('、')} 沒有匯率，未計入合計（請到設定填入）。` : '')
    + (currencies().length > 1 ? `外幣以設定中的匯率換算為 ${base}。` : '');

  renderTrend(range, recs);
  renderCats('#expCats', '#expCatTotal', recs.expense, exp, base);
  renderCats('#incCats', '#incCatTotal', recs.income, inc, base);
  renderByCurrency(recs);

  const all = sortRecs([...recs.expense, ...recs.income]);
  $('#detailCount').textContent = `${all.length} 筆`;
  const shown = all.slice(0, 300);
  let html = '';
  let lastDate = '';
  for (const r of shown) {
    if (r.date !== lastDate && stats.gran !== 'day') {
      const d = parseDate(r.date);
      html += `<li class="date-sep">${r.date}（${WEEK[d.getDay()]}）</li>`;
      lastDate = r.date;
    }
    html += recItemHTML(r, r.type, { showType: true });
  }
  if (all.length > shown.length) html += `<li class="empty">只顯示最新 ${shown.length} 筆，完整資料請匯出或開啟試算表</li>`;
  $('#detailList').innerHTML = html || '<li class="empty">這段期間沒有記錄</li>';
}

function renderTrend([from, to], recs) {
  const el = $('#trendChart');
  const days = Math.round((parseDate(to) - parseDate(from)) / 86400000) + 1;
  let buckets = [];
  let keyOf;
  if (stats.gran === 'day') {
    for (let h = 0; h < 24; h++) buckets.push({ key: pad(h), label: `${h}` });
    keyOf = (r) => r.time.slice(0, 2);
  } else if (days <= 62) {
    for (let i = 0; i < days; i++) {
      const s = addDays(from, i);
      const d = parseDate(s);
      buckets.push({ key: s, label: stats.gran === 'week' ? WEEK[d.getDay()] : `${d.getDate()}` });
    }
    keyOf = (r) => r.date;
  } else {
    let y = +from.slice(0, 4), m = +from.slice(5, 7);
    const ey = +to.slice(0, 4), em = +to.slice(5, 7);
    while (y < ey || (y === ey && m <= em)) {
      buckets.push({ key: `${y}-${pad(m)}`, label: `${m}月` });
      m++; if (m > 12) { m = 1; y++; }
    }
    keyOf = (r) => r.date.slice(0, 7);
  }
  const idx = new Map(buckets.map((b, i) => [b.key, i]));
  buckets = buckets.map((b) => ({ ...b, inc: 0, exp: 0 }));
  for (const [type, list] of [['inc', recs.income], ['exp', recs.expense]]) {
    for (const r of list) {
      const i = idx.get(keyOf(r));
      const v = toBase(r.amount, r.currency);
      if (i != null && !Number.isNaN(v)) buckets[i][type] += v;
    }
  }
  const max = Math.max(1, ...buckets.map((b) => Math.max(b.inc, b.exp)));
  if (buckets.every((b) => !b.inc && !b.exp)) { el.innerHTML = '<p class="empty">沒有資料</p>'; return; }

  const W = 360, H = 170, padL = 34, padB = 18, padT = 6;
  const cw = (W - padL) / buckets.length;
  const bw = Math.max(1.5, Math.min(10, cw * 0.38));
  const y = (v) => H - padB - (v / max) * (H - padB - padT);
  const step = Math.ceil(buckets.length / 10);
  const base = data.settings.baseCurrency;
  let svg = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="收支趨勢圖">`;
  for (const t of [0, 0.5, 1]) {
    const yy = y(max * t);
    svg += `<line class="axis" x1="${padL}" x2="${W}" y1="${yy}" y2="${yy}" />`;
    svg += `<text x="${padL - 6}" y="${yy + 3}" text-anchor="end">${shortNum(max * t)}</text>`;
  }
  buckets.forEach((b, i) => {
    const cx = padL + cw * i + cw / 2;
    if (b.inc) svg += `<rect class="bar-inc" x="${cx - bw}" y="${y(b.inc)}" width="${bw}" height="${H - padB - y(b.inc)}" rx="1.5"><title>${b.key} 收入 ${money(b.inc, base)}</title></rect>`;
    if (b.exp) svg += `<rect class="bar-exp" x="${cx}" y="${y(b.exp)}" width="${bw}" height="${H - padB - y(b.exp)}" rx="1.5"><title>${b.key} 支出 ${money(b.exp, base)}</title></rect>`;
    if (i % step === 0) svg += `<text x="${cx}" y="${H - 6}" text-anchor="middle">${esc(b.label)}</text>`;
  });
  svg += '</svg>';
  el.innerHTML = svg + `<div class="legend"><span><i style="background:var(--income)"></i>收入</span><span><i style="background:var(--expense)"></i>支出</span><span>單位：${base}</span></div>`;
}
function shortNum(v) {
  if (v >= 1e8) return `${+(v / 1e8).toFixed(1)}億`;
  if (v >= 1e4) return `${+(v / 1e4).toFixed(1)}萬`;
  return num(Math.round(v));
}

function renderCats(sel, totalSel, list, total, base) {
  const m = new Map();
  for (const r of list) {
    const v = toBase(r.amount, r.currency);
    if (Number.isNaN(v)) continue;
    const c = m.get(r.category) || { v: 0, n: 0 };
    c.v += v; c.n++;
    m.set(r.category, c);
  }
  const rows = [...m.entries()].sort((a, b) => b[1].v - a[1].v);
  $(totalSel).textContent = rows.length ? `${rows.length} 類・${list.length} 筆` : '';
  $(sel).innerHTML = rows.length ? rows.map(([cat, { v, n }]) => {
    const pct = total ? (v / total) * 100 : 0;
    return `<div class="cb"><div class="cb-top"><span>${esc(cat)} <span class="muted small">${n} 筆</span></span><span>${money(v, base)}・${pct.toFixed(1)}%</span></div>
      <div class="cb-track"><div class="cb-fill" style="width:${pct.toFixed(2)}%"></div></div></div>`;
  }).join('') : '<p class="empty">沒有資料</p>';
}

function renderByCurrency(recs) {
  const m = new Map();
  for (const [k, list] of [['inc', recs.income], ['exp', recs.expense]]) {
    for (const r of list) {
      const c = m.get(r.currency) || { inc: 0, exp: 0 };
      c[k] += r.amount;
      m.set(r.currency, c);
    }
  }
  if (!m.size) { $('#byCurrency').innerHTML = '<p class="empty">沒有資料</p>'; return; }
  $('#byCurrency').innerHTML = `<table class="tbl"><thead><tr><th>幣別</th><th>收入</th><th>支出</th><th>結餘</th></tr></thead><tbody>${
    [...m.entries()].map(([c, v]) => `<tr><td>${esc(c)}</td><td>${money(v.inc, c)}</td><td>${money(v.exp, c)}</td><td>${money(v.inc - v.exp, c)}</td></tr>`).join('')
  }</tbody></table>`;
}

function bindStats() {
  $('#granSeg').addEventListener('click', (e) => {
    const b = e.target.closest('button[data-gran]');
    if (!b) return;
    stats.gran = b.dataset.gran;
    if (stats.gran === 'custom' && !stats.from) {
      const [f, t] = (() => { const g = stats.gran; stats.gran = 'month'; const r = periodRange(); stats.gran = g; return r; })();
      stats.from = f; stats.to = t;
    }
    renderStats();
  });
  $('#prevPeriod').onclick = () => shiftPeriod(-1);
  $('#nextPeriod').onclick = () => shiftPeriod(1);
  $('#periodLabel').onclick = () => { stats.anchor = todayStr(); renderStats(); };
  $('#rangeFrom').onchange = (e) => { stats.from = e.target.value; renderStats(); };
  $('#rangeTo').onchange = (e) => { stats.to = e.target.value; renderStats(); };
  $('#detailList').addEventListener('click', (e) => onListClick(e));
  // 手機左右滑動切換期間
  let sx = null, sy = null;
  const sec = $('#tab-stats');
  sec.addEventListener('touchstart', (e) => { sx = e.touches[0].clientX; sy = e.touches[0].clientY; }, { passive: true });
  sec.addEventListener('touchend', (e) => {
    if (sx == null || stats.gran === 'custom') return;
    const dx = e.changedTouches[0].clientX - sx, dy = e.changedTouches[0].clientY - sy;
    if (Math.abs(dx) > 70 && Math.abs(dx) > Math.abs(dy) * 2) shiftPeriod(dx < 0 ? 1 : -1);
    sx = null;
  }, { passive: true });
}

/* ============================================================
 * 匯出／匯入
 * ============================================================ */
const CSV_HEAD = ['類型', '日期', '時間', '類別', '項目', '金額', '幣別', '備註', 'ID'];

function csvCell(v) {
  const s = String(v ?? '');
  return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}
/** 手機優先用分享選單（存到「檔案」、傳 LINE／Email），不支援時才用一般下載 */
async function download(name, text, mime) {
  const blob = new Blob([text], { type: mime });
  try {
    const file = new File([blob], name, { type: mime });
    if (navigator.canShare && navigator.canShare({ files: [file] }) && matchMedia('(pointer: coarse)').matches) {
      await navigator.share({ files: [file], title: name });
      return;
    }
  } catch (e) {
    if (e.name === 'AbortError') return; // 使用者取消分享
  }
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}
function exportRecords() {
  const scope = $('#exportScope').value;
  const range = scope === 'period' ? periodRange() : ['0000-00-00', '9999-99-99'];
  const r = recordsIn(range);
  return { list: [...sortRecs(r.expense), ...sortRecs(r.income)], range, scope };
}
function stamp() { return todayStr().replace(/-/g, ''); }

async function exportCsv() {
  const { list, range, scope } = exportRecords();
  if (!list.length) return toast('沒有可匯出的資料', 'error');
  const lines = [CSV_HEAD.join(',')];
  for (const r of list) {
    lines.push([TYPES[r.type].label, r.date, r.time, r.category, r.item, r.amount, r.currency, r.note, r.id].map(csvCell).join(','));
  }
  const suffix = scope === 'period' ? `${range[0]}_${range[1]}` : `全部_${stamp()}`;
  await download(`記帳_${suffix}.csv`, '﻿' + lines.join('\r\n'), 'text/csv;charset=utf-8');
  toast(`已匯出 ${list.length} 筆`, 'ok');
}
async function exportJson() {
  const { list, scope, range } = exportRecords();
  const payload = {
    app: 'personal-ledger', version: 1, exportedAt: new Date().toISOString(),
    range: scope === 'period' ? range : null,
    expense: list.filter((r) => r.type === 'expense').map(({ type, ...r }) => r),
    income: list.filter((r) => r.type === 'income').map(({ type, ...r }) => r),
    categories: data.categories, settings: data.settings,
  };
  await download(`記帳備份_${stamp()}.json`, JSON.stringify(payload, null, 2), 'application/json');
  toast(`已匯出 ${list.length} 筆`, 'ok');
}
function csvTemplate() {
  const rows = [CSV_HEAD.slice(0, 8).join(','), `支出,${todayStr()},12:30,午餐,牛肉麵,150,TWD,`, `收入,${todayStr()},09:00,薪資,11 月薪資,50000,TWD,公司轉帳`];
  download('記帳_匯入範本.csv', '﻿' + rows.join('\r\n'), 'text/csv;charset=utf-8');
}

/** RFC 4180 CSV 解析（支援引號、逗號、換行） */
function parseCsv(text) {
  const rows = [];
  let row = [], cell = '', q = false;
  text = text.replace(/^﻿/, '');
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) {
      if (c === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; }
      else cell += c;
    } else if (c === '"') q = true;
    else if (c === ',') { row.push(cell); cell = ''; }
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && text[i + 1] === '\n') i++;
      row.push(cell); rows.push(row); row = []; cell = '';
    } else cell += c;
  }
  if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
  return rows.filter((r) => r.some((v) => v.trim() !== ''));
}

const HEAD_MAP = {
  type: ['類型', '收支', 'type'], date: ['日期', 'date'], time: ['時間', 'time'], category: ['類別', '分類', 'category'],
  item: ['項目', '名稱', '說明', 'item', 'description'], amount: ['金額', '價格', 'amount', 'price'],
  currency: ['幣別', '貨幣', 'currency'], note: ['備註', 'note', 'memo'], id: ['id', '編號'],
};
function normDate(s) {
  const m = String(s || '').trim().match(/^(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})/);
  return m ? `${m[1]}-${pad(m[2])}-${pad(m[3])}` : '';
}
function normTime(s) {
  const m = String(s || '').trim().match(/^(\d{1,2}):(\d{2})/);
  return m ? `${pad(m[1])}:${m[2]}` : '00:00';
}
function normType(s) {
  const t = String(s || '').trim().toLowerCase();
  if (['支出', 'expense', '出', '-'].includes(t)) return 'expense';
  if (['收入', 'income', '入', '+'].includes(t)) return 'income';
  return null;
}

function recordsFromCsv(text) {
  const rows = parseCsv(text);
  if (rows.length < 2) throw new Error('CSV 沒有資料列');
  const head = rows[0].map((h) => h.trim().toLowerCase());
  const col = {};
  for (const [k, names] of Object.entries(HEAD_MAP)) col[k] = head.findIndex((h) => names.map((n) => n.toLowerCase()).includes(h));
  if (col.date < 0 || col.amount < 0) throw new Error('找不到「日期」或「金額」欄位');
  const out = [], errors = [];
  rows.slice(1).forEach((r, i) => {
    const get = (k) => (col[k] >= 0 ? (r[col[k]] ?? '').trim() : '');
    let amount = Number(get('amount').replace(/[,\s$]/g, ''));
    let type = col.type >= 0 ? normType(get('type')) : null;
    if (!type) type = amount < 0 ? 'expense' : (col.type >= 0 ? null : 'expense');
    const date = normDate(get('date'));
    if (!type || !date || !Number.isFinite(amount)) { errors.push(`第 ${i + 2} 列`); return; }
    amount = Math.abs(amount);
    out.push({
      type, id: get('id') || uid(), date, time: normTime(get('time')),
      category: get('category') || '其他', item: get('item') || get('category') || '（未命名）',
      amount, currency: (get('currency') || data.settings.baseCurrency).toUpperCase(), note: get('note'),
      createdAt: new Date().toISOString(),
    });
  });
  return { records: out, errors };
}

function recordsFromJson(text) {
  const j = JSON.parse(text);
  const out = [];
  const take = (list, type) => (list || []).forEach((r) => out.push({ ...r, type, id: r.id || uid(), date: normDate(r.date), time: normTime(r.time), currency: String(r.currency || data.settings.baseCurrency).toUpperCase() }));
  if (Array.isArray(j)) j.forEach((r) => { const t = normType(r.type) || r.type; if (TYPES[t]) take([r], t); });
  else { take(j.expense, 'expense'); take(j.income, 'income'); }
  return { records: out.filter((r) => r.date && Number.isFinite(Number(r.amount))).map((r) => ({ ...r, amount: Number(r.amount) })), errors: [] };
}

let pendingImport = null;
async function onImportFile(e) {
  const file = e.target.files[0];
  const box = $('#importPreview');
  const btn = $('#btnDoImport');
  pendingImport = null;
  btn.classList.add('hidden');
  if (!file) { box.classList.add('hidden'); return; }
  try {
    const text = await file.text();
    const parsed = /\.json$/i.test(file.name) || text.trim().startsWith('{') || text.trim().startsWith('[') ? recordsFromJson(text) : recordsFromCsv(text);
    const dup = parsed.records.filter((r) => data[r.type].some((x) => x.id === r.id)).length;
    const nExp = parsed.records.filter((r) => r.type === 'expense').length;
    const nInc = parsed.records.length - nExp;
    const newCats = new Set(parsed.records.filter((r) => !data.categories[r.type].includes(r.category)).map((r) => r.category));
    box.innerHTML = `<b>${esc(file.name)}</b>：解析到 ${parsed.records.length} 筆（支出 ${nExp}、收入 ${nInc}）
      <ul>${dup ? `<li>${dup} 筆 ID 已存在，會略過</li>` : ''}
      ${newCats.size ? `<li>將新增類別：${[...newCats].map(esc).join('、')}</li>` : ''}
      ${parsed.errors.length ? `<li style="color:var(--danger)">${parsed.errors.length} 列格式有誤已忽略（${parsed.errors.slice(0, 5).join('、')}${parsed.errors.length > 5 ? '…' : ''}）</li>` : ''}
      <li>只會「新增」資料，不會刪除或覆蓋現有記錄</li></ul>`;
    box.classList.remove('hidden');
    const fresh = parsed.records.filter((r) => !data[r.type].some((x) => x.id === r.id));
    if (fresh.length) { pendingImport = fresh; btn.textContent = `確認匯入 ${fresh.length} 筆`; btn.classList.remove('hidden'); }
  } catch (err) {
    box.innerHTML = `<span style="color:var(--danger)">無法讀取：${esc(err.message)}</span>`;
    box.classList.remove('hidden');
  }
}
async function doImport() {
  if (!pendingImport?.length) return;
  const records = pendingImport;
  pendingImport = null;
  $('#btnDoImport').classList.add('hidden');
  $('#importPreview').classList.add('hidden');
  $('#importFile').value = '';
  // 分批送出，避免單次請求過大
  for (let i = 0; i < records.length; i += 500) {
    await mutate('importRecords', { records: records.slice(i, i + 500) });
  }
  if (config.mode !== 'sheet') toast(`已匯入 ${records.length} 筆`, 'ok');
}

function bindIO() {
  $('#btnExportCsv').onclick = exportCsv;
  $('#btnExportJson').onclick = exportJson;
  $('#btnCsvTemplate').onclick = csvTemplate;
  $('#importFile').onchange = onImportFile;
  $('#btnDoImport').onclick = doImport;
}
function renderSheetLink() {
  const el = $('#sheetLinkWrap');
  el.innerHTML = config.mode === 'sheet' && data.sheetUrl
    ? `📄 <a href="${esc(data.sheetUrl)}" target="_blank" rel="noopener">開啟 Google 試算表</a>（可用「檔案 → 版本記錄」還原、「檔案 → 下載」存成 Excel）`
    : '';
}

/* ============================================================
 * 設定
 * ============================================================ */
let draftSettings = null;

function openSettings() {
  const dlg = $('#settingsDlg');
  $$('input[name=mode]', dlg).forEach((r) => { r.checked = r.value === config.mode; });
  $('#cfgUrl').value = config.apiUrl || '';
  $('#cfgToken').value = config.token || '';
  $('#cfgLang').value = config.lang || 'zh-TW';
  $('#testResult').textContent = '';
  $('#rateFetchInfo').textContent = '';
  $('#voiceSupport').textContent = SR
    ? '✓ 此瀏覽器支援語音辨識（需 https 與麥克風權限）。'
    : '✗ 此瀏覽器不支援語音辨識 API；可改用手機鍵盤上的麥克風（聽寫）輸入。';
  draftSettings = structuredClone(data.settings);
  toggleSheetFields();
  renderRateTable();
  renderCatManager();
  dlg.showModal();
}
function toggleSheetFields() {
  const mode = $('input[name=mode]:checked')?.value;
  $('#sheetFields').classList.toggle('hidden', mode !== 'sheet');
}
function renderRateTable() {
  const s = draftSettings;
  const codes = [s.baseCurrency, ...Object.keys(s.rates).filter((c) => c !== s.baseCurrency).sort()];
  $('#cfgBase').innerHTML = codes.map((c) => `<option>${esc(c)}</option>`).join('');
  $('#cfgBase').value = s.baseCurrency;
  $('#rateTable').innerHTML = codes.filter((c) => c !== s.baseCurrency).map((c) => `
    <div class="rate-row" data-code="${esc(c)}">
      <span class="code">${esc(c)}</span>
      <input type="number" step="any" min="0" value="${s.rates[c] || ''}" placeholder="1 ${esc(c)} = ? ${esc(s.baseCurrency)}">
      <button type="button" class="btn icon" aria-label="移除">✕</button>
    </div>`).join('') || '<p class="muted small">目前只有基準幣別</p>';
}
function readRateTable() {
  $$('#rateTable .rate-row').forEach((row) => {
    const v = Number($('input', row).value);
    draftSettings.rates[row.dataset.code] = Number.isFinite(v) && v > 0 ? v : 0;
  });
}
function renderCatManager() {
  for (const type of Object.keys(TYPES)) {
    const ul = $(type === 'expense' ? '#catListExpense' : '#catListIncome');
    ul.innerHTML = data.categories[type].map((c) => `<li>${esc(c)}<button type="button" data-type="${type}" data-name="${esc(c)}" aria-label="刪除 ${esc(c)}">✕</button></li>`).join('');
  }
}

async function fetchRates() {
  const base = draftSettings.baseCurrency;
  const info = $('#rateFetchInfo');
  info.textContent = '取得中…';
  try {
    const res = await fetch(`https://open.er-api.com/v6/latest/${encodeURIComponent(base)}`);
    const j = await res.json();
    if (j.result !== 'success') throw new Error(j['error-type'] || '服務回應錯誤');
    readRateTable();
    let n = 0;
    for (const c of Object.keys(draftSettings.rates)) {
      if (c !== base && j.rates[c]) { draftSettings.rates[c] = +(1 / j.rates[c]).toPrecision(6); n++; }
    }
    renderRateTable();
    const t = j.time_last_update_unix ? new Date(j.time_last_update_unix * 1000).toLocaleDateString('zh-TW') : '';
    info.textContent = `已更新 ${n} 種（資料日期 ${t}，按「儲存設定」生效）`;
  } catch (e) {
    info.textContent = `更新失敗：${e.message}，請手動輸入`;
  }
}

async function testConnection() {
  const out = $('#testResult');
  const url = $('#cfgUrl').value.trim();
  const token = $('#cfgToken').value.trim();
  if (!/^https:\/\/script\.google(usercontent)?\.com\//.test(url)) { out.textContent = '✗ 網址應以 https://script.google.com/ 開頭、/exec 結尾'; return; }
  out.textContent = '連線中…';
  const saved = { ...config };
  config.apiUrl = url; config.token = token;
  try {
    const r = await api('ping');
    out.innerHTML = `✓ 連線成功：<a href="${esc(r.sheetUrl)}" target="_blank" rel="noopener">${esc(r.name)}</a>`;
  } catch (e) {
    out.textContent = `✗ ${e instanceof ApiError ? e.message : '無法連線（確認網址正確、部署存取權為「所有人」）'}`;
  } finally {
    config.apiUrl = saved.apiUrl; config.token = saved.token;
  }
}

async function saveSettings() {
  readRateTable();
  const mode = $('input[name=mode]:checked')?.value || 'demo';
  const url = $('#cfgUrl').value.trim();
  const token = $('#cfgToken').value.trim();
  if (mode === 'sheet' && (!url || !token)) return toast('請填入 Apps Script 網址與 Token', 'error');
  const modeChanged = mode !== config.mode || url !== config.apiUrl || token !== config.token;
  const demoData = config.mode === 'demo' ? data : null;

  draftSettings.baseCurrency = $('#cfgBase').value;
  const cleanRates = {};
  for (const [c, v] of Object.entries(draftSettings.rates)) if (v > 0 || c === draftSettings.baseCurrency) cleanRates[c] = c === draftSettings.baseCurrency ? 1 : v;
  // 沒填匯率（0）的幣別不儲存
  const newSettings = { baseCurrency: draftSettings.baseCurrency, rates: cleanRates };
  const settingsChanged = JSON.stringify(newSettings) !== JSON.stringify({ baseCurrency: data.settings.baseCurrency, rates: Object.fromEntries(Object.entries(data.settings.rates).filter(([, v]) => v > 0)) });

  config = { ...config, mode, apiUrl: url, token, lang: $('#cfgLang').value };
  store.set(LS.config, config);
  $('#settingsDlg').close();

  if (modeChanged) {
    queue = mode === 'sheet' ? queue : [];
    store.set(LS.queue, queue);
    data = loadLocalData();
    renderAll();
    if (mode === 'sheet') {
      await reloadFromServer();
      const n = demoData ? demoData.expense.length + demoData.income.length : 0;
      if (n && confirm(`示範模式中有 ${n} 筆資料，要匯入到 Google 試算表嗎？（重複的會自動略過）`)) {
        const records = [...demoData.expense.map((r) => ({ ...r, type: 'expense' })), ...demoData.income.map((r) => ({ ...r, type: 'income' }))];
        await mutate('importRecords', { records });
      }
    }
  }
  if (settingsChanged) await mutate('saveSettings', { settings: newSettings });
  renderAll();
  toast('設定已儲存', 'ok');
}

function bindSettings() {
  $('#btnSettings').onclick = openSettings;
  $$('input[name=mode]').forEach((r) => { r.onchange = toggleSheetFields; });
  $('#btnTest').onclick = testConnection;
  $('#btnSaveSettings').onclick = saveSettings;
  $('#btnFetchRates').onclick = fetchRates;
  $('#cfgBase').onchange = (e) => {
    readRateTable();
    const nb = e.target.value;
    const old = draftSettings.baseCurrency;
    const k = draftSettings.rates[nb];
    if (k > 0) { // 以新基準重新計算其他匯率
      for (const c of Object.keys(draftSettings.rates)) draftSettings.rates[c] = c === nb ? 1 : +((draftSettings.rates[c] || 0) / k).toPrecision(6);
      draftSettings.rates[old] = +(1 / k).toPrecision(6);
    }
    draftSettings.baseCurrency = nb;
    renderRateTable();
  };
  $('#btnAddCurrency').onclick = () => {
    const c = $('#newCurrency').value.trim().toUpperCase();
    if (!/^[A-Z]{3}$/.test(c)) return toast('請輸入 3 碼幣別代碼，例如 THB', 'error');
    readRateTable();
    if (!(c in draftSettings.rates)) draftSettings.rates[c] = 0;
    $('#newCurrency').value = '';
    renderRateTable();
  };
  $('#rateTable').addEventListener('click', (e) => {
    const b = e.target.closest('button');
    if (!b) return;
    readRateTable();
    delete draftSettings.rates[b.closest('.rate-row').dataset.code];
    renderRateTable();
  });
  $('#settingsDlg').addEventListener('click', async (e) => {
    const b = e.target.closest('.chip-list button');
    if (!b) return;
    const { type, name } = b.dataset;
    if (data.categories[type].length <= 1) return toast('至少要保留一個類別', 'error');
    if (!confirm(`刪除${TYPES[type].label}類別「${name}」？既有記錄不受影響。`)) return;
    await mutate('deleteCategory', { type, name });
    renderCatManager();
  });
  $('#btnClearLocal').onclick = () => {
    if (!confirm(config.mode === 'demo' ? '這會刪除示範模式的所有資料，確定？' : '清除本機快取後會重新從試算表讀取（未同步的資料會遺失），確定？')) return;
    store.del(config.mode === 'sheet' ? LS.cache : LS.demo);
    queue = []; store.set(LS.queue, queue);
    data = loadLocalData();
    $('#settingsDlg').close();
    renderAll();
    reloadFromServer();
  };
}

/* ============================================================
 * 分頁、整體渲染、啟動
 * ============================================================ */
function showTab(name) {
  $$('.tabs button').forEach((b) => b.classList.toggle('active', b.dataset.tab === name));
  $$('.tab').forEach((s) => s.classList.toggle('active', s.id === `tab-${name}`));
  if (name === 'stats') renderStats();
  if (name === 'io') renderSheetLink();
  try { history.replaceState(null, '', `#${name}`); } catch { /* ignore */ }
  window.scrollTo({ top: 0 });
}

function renderAll() {
  for (const type of Object.keys(TYPES)) {
    if (!forms[type].editingId) { fillCategorySelect(type); fillCurrencySelect(type); }
    renderRecent(type);
  }
  if ($('#tab-stats').classList.contains('active')) renderStats();
  renderSheetLink();
  updateStatus();
}

function init() {
  buildForms();
  bindStats();
  bindIO();
  bindSettings();
  $('.tabs').addEventListener('click', (e) => { const b = e.target.closest('button[data-tab]'); if (b) showTab(b.dataset.tab); });
  $('#syncStatus').onclick = () => (config.mode === 'sheet' ? reloadFromServer() : openSettings());
  window.addEventListener('online', () => flushQueue().then(() => reloadFromServer({ silent: true })));
  window.addEventListener('offline', updateStatus);
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'visible' && config.mode === 'sheet') reloadFromServer({ silent: true });
  });

  const initial = location.hash.slice(1);
  showTab(['expense', 'income', 'stats', 'io'].includes(initial) ? initial : 'expense');
  renderAll();

  if (config.mode === 'sheet') {
    if (config.apiUrl && config.token) reloadFromServer({ silent: true });
    else openSettings();
  } else if (!store.get(LS.demo, null)) {
    toast('目前為示範模式，資料只存在本機。按右上角 ⚙️ 連接 Google 試算表');
  }

  if ('serviceWorker' in navigator && location.protocol === 'https:') {
    navigator.serviceWorker.register('sw.js').catch((e) => console.warn('SW 註冊失敗', e));
  }
}

init();
