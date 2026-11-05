// Excel import/export (offline, client-side via SheetJS)

/* =============================================================
   UTILS
============================================================= */
(function(){
  function uuid(){
    // RFC4122 v4-ish
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c){
      const r = crypto.getRandomValues(new Uint8Array(1))[0] & 15;
      const v = c === 'x' ? r : (r & 0x3 | 0x8);
      return v.toString(16);
    });
  }
  // Date format: DD-MM-YYYY or YYYY only
  function isValidDate(str){
    if(!str) return false;
    const t = String(str).trim();
    if(/^\d{4}$/.test(t)) return true; // year-only allowed
    if(/^\d{2}-\d{2}-\d{4}$/.test(t)){
      const [dd,mm,yyyy] = t.split('-').map(Number);
      const d = new Date(Date.UTC(yyyy, mm-1, dd));
      return d.getUTCFullYear()===yyyy && (d.getUTCMonth()+1)===mm && d.getUTCDate()===dd;
    }
    return false;
  }
  function normalizeDate(str){
    if(!str) return '';
    const t = String(str).trim();
    if(/^\d{4}$/.test(t)) return t;
    if(/^\d{2}-\d{2}-\d{4}$/.test(t)) return t;
    // try to parse other inputs, output as DD-MM-YYYY if possible
    const d = new Date(t);
    if(Number.isNaN(d.getTime())) return '';
    const yyyy = d.getUTCFullYear();
    const mm = String(d.getUTCMonth()+1).padStart(2,'0');
    const dd = String(d.getUTCDate()).padStart(2,'0');
    return `${dd}-${mm}-${yyyy}`;
  }
  function toMMDD(dateStr){
    if(!isValidDate(dateStr)) return '';
    if(/^\d{4}$/.test(dateStr)) return '';
    const [dd,mm] = dateStr.split('-');
    return `${mm}-${dd}`;
  }
  function parseTags(str){
    if(!str) return [];
    return String(str).split(',').map(s=>s.trim()).filter(Boolean);
  }
  // Parse date from various formats and return normalized DD-MM-YYYY or YYYY
  function parseSearchDate(query){
    if(!query) return null;
    const q = String(query).trim();
    
    // YYYY only
    if(/^\d{4}$/.test(q)) return q;
    
    // DD-MM-YYYY
    if(/^\d{1,2}-\d{1,2}-\d{4}$/.test(q)){
      const parts = q.split('-');
      const dd = String(Number(parts[0])).padStart(2,'0');
      const mm = String(Number(parts[1])).padStart(2,'0');
      const yyyy = parts[2];
      return `${dd}-${mm}-${yyyy}`;
    }
    
    // YYYY-MM-DD
    if(/^\d{4}-\d{1,2}-\d{1,2}$/.test(q)){
      const parts = q.split('-');
      const yyyy = parts[0];
      const mm = String(Number(parts[1])).padStart(2,'0');
      const dd = String(Number(parts[2])).padStart(2,'0');
      return `${dd}-${mm}-${yyyy}`;
    }
    
    // DD/MM/YYYY
    if(/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(q)){
      const parts = q.split('/');
      const dd = String(Number(parts[0])).padStart(2,'0');
      const mm = String(Number(parts[1])).padStart(2,'0');
      const yyyy = parts[2];
      return `${dd}-${mm}-${yyyy}`;
    }
    
    // YYYY/MM/DD
    if(/^\d{4}\/\d{1,2}\/\d{1,2}$/.test(q)){
      const parts = q.split('/');
      const yyyy = parts[0];
      const mm = String(Number(parts[1])).padStart(2,'0');
      const dd = String(Number(parts[2])).padStart(2,'0');
      return `${dd}-${mm}-${yyyy}`;
    }
    
    // DD-MM (assume current year)
    if(/^\d{1,2}-\d{1,2}$/.test(q)){
      const parts = q.split('-');
      const dd = String(Number(parts[0])).padStart(2,'0');
      const mm = String(Number(parts[1])).padStart(2,'0');
      // Return pattern for matching (not full date)
      return `DD-MM-${dd}-${mm}`;
    }
    
    // Try parsing as natural date
    const d = new Date(q);
    if(!Number.isNaN(d.getTime())){
      const yyyy = d.getUTCFullYear();
      const mm = String(d.getUTCMonth()+1).padStart(2,'0');
      const dd = String(d.getUTCDate()).padStart(2,'0');
      return `${dd}-${mm}-${yyyy}`;
    }
    
    return null;
  }
  // Check if event date matches search date
  function eventMatchesDate(evtDate, searchDate){
    if(!evtDate || !searchDate) return false;
    
    // If search is year only
    if(/^\d{4}$/.test(searchDate)){
      if(/^\d{4}$/.test(evtDate)) return evtDate === searchDate;
      if(/^\d{2}-\d{2}-\d{4}$/.test(evtDate)) return evtDate.slice(-4) === searchDate;
      return false;
    }
    
    // If search is DD-MM pattern (without year)
    if(searchDate.startsWith('DD-MM-')){
      const parts = searchDate.split('-');
      const searchDD = parts[2];
      const searchMM = parts[3];
      if(/^\d{2}-\d{2}-\d{4}$/.test(evtDate)){
        const [dd, mm] = evtDate.split('-');
        return dd === searchDD && mm === searchMM;
      }
      return false;
    }
    
    // Full date match
    return evtDate === searchDate;
  }
  function extractYear(dateStr){
    if(!dateStr) return null;
    const d = String(dateStr).trim();
    if(/^\d{4}$/.test(d)) return Number(d);
    if(/^\d{2}-\d{2}-\d{4}$/.test(d)) return Number(d.slice(-4));
    return null;
  }
  function downloadBlob(filename, blob){
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename; a.click();
    setTimeout(()=>URL.revokeObjectURL(url), 1000);
  }
  window.AppUtils = { uuid, isValidDate, normalizeDate, toMMDD, parseTags, downloadBlob, parseSearchDate, eventMatchesDate, extractYear };
})();

/* =============================================================
   STORE (eventsStore)
============================================================= */
const eventsStore = (function(){
  let events = [];

  function validateEvent(evt){
    if(!evt) return false;
    if(!evt.date) return false;
    
    // Try to validate the date - first check if it's already valid
    if(AppUtils.isValidDate(evt.date)) return true;
    
    // If not valid, try to normalize it and check again
    const normalized = AppUtils.normalizeDate(evt.date);
    if(!normalized || normalized === '') return false;
    if(!AppUtils.isValidDate(normalized)) return false;
    
    if(!evt.title || !String(evt.title).trim()) return false;
    return true;
  }
  function normalizeEvent(evt){
    const copy = {
      id: evt.id || AppUtils.uuid(),
      date: AppUtils.normalizeDate(evt.date) || evt.date, // Keep original if normalization fails
      title: (evt.title||'').trim(),
      description: (evt.description||'').trim(),
      category: (evt.category||'').trim(),
      tags: typeof evt.tags === 'string' ? evt.tags : (Array.isArray(evt.tags) ? evt.tags.join(',') : ''),
      mediaSource: (evt.mediaSource||'').trim(),
      webSource: (evt.webSource||'').trim(),
      officialDocument: (evt.officialDocument||'').trim(),
      status: evt.status || 'public', // 'public' or 'WaitingToApprove'
      submitterName: (evt.submitterName||'').trim(), // Name of person who suggested the event
      submitterId: (evt.submitterId||'').trim() // Employee ID of person who suggested the event
    };
    return copy;
  }
  function loadFromArray(arr){
    const loaded = [];
    for(const raw of arr){
      // First validate the raw date (before normalization)
      const normalizedDate = AppUtils.normalizeDate(raw.date);
      if(!normalizedDate || normalizedDate === ''){
        toast(`שגיאה: פורמט תאריך לא תקין (${raw.date||''})`, 'error');
        continue;
      }
      if(!AppUtils.isValidDate(normalizedDate)){
        toast(`שגיאה: פורמט תאריך לא תקין (${raw.date||''})`, 'error');
        continue;
      }
      
      const evt = normalizeEvent(raw);
      if(!validateEvent(evt)) {
        toast(`שגיאה: כותרת ריקה או תאריך לא תקין (${raw.date||''})`, 'error');
        continue;
      }
      loaded.push(evt);
    }
    events = loaded;
  }
  function toArray(){ return [...events]; }
  function saveToLocalStorage(){
    try {
      const data = toArray();
      localStorage.setItem('eventsData', JSON.stringify(data, null, 2));
    } catch(err) {
      console.error('Failed to save to localStorage:', err);
    }
  }
  function loadFromLocalStorage(){
    try {
      const saved = localStorage.getItem('eventsData');
      if(saved){
        const data = JSON.parse(saved);
        if(Array.isArray(data) && data.length > 0){
          return data;
        }
      }
    } catch(err) {
      console.error('Failed to load from localStorage:', err);
    }
    return null;
  }
  function add(evt){
    // First normalize the date before creating the event
    const normalizedDate = AppUtils.normalizeDate(evt.date);
    if(!normalizedDate || normalizedDate === ''){
      throw new Error('Invalid date format');
    }
    if(!AppUtils.isValidDate(normalizedDate)){
      throw new Error('Invalid date format');
    }
    // Use normalized date
    const evtWithNormalizedDate = { ...evt, date: normalizedDate };
    const e = normalizeEvent(evtWithNormalizedDate);
    if(!validateEvent(e)) throw new Error('Invalid event');
    events.push(e);
    saveToLocalStorage(); // Auto-save after add
    return e;
  }
  function update(id, patch){
    const idx = events.findIndex(e=>e.id===id);
    if(idx===-1) throw new Error('Not found');
    // If date is being updated, normalize it first
    if(patch.date){
      const normalizedDate = AppUtils.normalizeDate(patch.date);
      if(!normalizedDate || normalizedDate === ''){
        throw new Error('Invalid date format');
      }
      if(!AppUtils.isValidDate(normalizedDate)){
        throw new Error('Invalid date format');
      }
      patch.date = normalizedDate;
    }
    const merged = normalizeEvent({...events[idx], ...patch, id});
    if(!validateEvent(merged)) throw new Error('Invalid');
    events[idx]=merged;
    saveToLocalStorage(); // Auto-save after update
    return merged;
  }
  function remove(id){ const idx = events.findIndex(e=>e.id===id); if(idx===-1) return false; events.splice(idx,1); saveToLocalStorage(); // Auto-save after remove
 return true; }
  function search(query, {category, tags, showPending = false}){
    const q = (query||'').trim();
    const tagList = AppUtils.parseTags(tags||'');
    const filterCat = category ? category.trim() : '';
    
    // Check if query is a date
    const searchDate = AppUtils.parseSearchDate(q);
    const isDateSearch = searchDate !== null;
    const qLower = q.toLowerCase();
    
    return events.filter(e=>{
      // Filter by status (only show public in viewer, all in admin)
      if(!showPending && (e.status || 'public') !== 'public') return false;
      
      // Filter by category
      if(filterCat){
        const evtCat = (e.category || '').trim();
        if(evtCat !== filterCat) return false;
      }
      // Filter by tags
      if(tagList.length > 0){
        const evtTags = AppUtils.parseTags(e.tags||'');
        if(!tagList.every(t=>evtTags.includes(t))) return false;
      }
      // Filter by date or text query
      if(q){
        if(isDateSearch){
          // Date search: check if event date matches
          if(!AppUtils.eventMatchesDate(e.date, searchDate)) return false;
        } else {
          // Text search: check in all fields
          const evtTags = AppUtils.parseTags(e.tags||'');
          const hay = [e.title, e.description, e.category, e.tags, e.mediaSource, e.webSource, e.officialDocument, e.date].join(' ').toLowerCase();
          if(!hay.includes(qLower) && !evtTags.some(t=>t.toLowerCase().includes(qLower))) return false;
        }
      }
      // If we passed all filters, include this event
      return true;
    });
  }
  function byDayMonth(mmdd){
    return events.filter(e=>AppUtils.toMMDD(e.date)===mmdd)
      .sort((a,b)=> a.date.localeCompare(b.date)); // ascending by year
  }
  function stats(){
    const total = events.length;
    const byCat = {};
    for(const e of events){ byCat[e.category||''] = (byCat[e.category||'']||0)+1; }
    return { total, byCat };
  }
  return { loadFromArray, toArray, add, update, remove, search, byDayMonth, stats, saveToLocalStorage, loadFromLocalStorage };
})();

/* =============================================================
   EXCEL (import/export) + CSV fallback
============================================================= */
const ExcelIO = (function(){
  const REQUIRED = ['date','title'];
  const COLUMN_ALIASES = {
    'תאריך': 'date',
    'כותרת': 'title',
    'תיאור': 'description',
    'קטגוריה': 'category',
    'תגיות': 'tags',
    'מקור מדיה': 'mediaSource',
    'מקור אתר': 'webSource',
    'מסמך רשמי': 'officialDocument',
    'מזהה': 'id',
    'סטטוס': 'status'
  };

  function mapRowKeys(raw){
    const o = {};
    for(const [k, v] of Object.entries(raw)){
      const key = String(k).trim();
      const mapped = COLUMN_ALIASES[key] || key;
      o[mapped] = v;
    }
    return o;
  }
  function excelSerialToDate(serial){
    if(typeof serial !== 'number' || !Number.isFinite(serial)) return null;
    const utc = Math.round((serial - 25569) * 86400 * 1000);
    const d = new Date(utc);
    if(Number.isNaN(d.getTime())) return null;
    const yyyy = d.getUTCFullYear();
    const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
    const dd = String(d.getUTCDate()).padStart(2, '0');
    return `${dd}-${mm}-${yyyy}`;
  }
  function normalizeImportedRow(raw){
    const row = mapRowKeys(raw);
    if(row.date !== undefined && row.date !== null && row.date !== ''){
      if(row.date instanceof Date){
        const d = row.date;
        const yyyy = d.getUTCFullYear();
        const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
        const dd = String(d.getUTCDate()).padStart(2, '0');
        row.date = `${dd}-${mm}-${yyyy}`;
      } else if(typeof row.date === 'number'){
        const converted = excelSerialToDate(row.date);
        if(converted) row.date = converted;
      } else {
        row.date = String(row.date).trim();
      }
    }
    for(const key of ['title','description','category','tags','mediaSource','webSource','officialDocument']){
      if(row[key] !== undefined && row[key] !== null) row[key] = String(row[key]).trim();
    }
    return row;
  }

  function validateColumns(cols){
    const missing = REQUIRED.filter(c=>!cols.includes(c));
    if(missing.length) throw new Error(`חסרות עמודות חובה: ${missing.join(', ')} (או בעברית: תאריך, כותרת)`);
  }
  async function importExcel(file){
    if(typeof XLSX === 'undefined') throw new Error('ספריית Excel לא נטענה. ודאו שקובץ xlsx.full.min.js נמצא בתיקייה.');
    const name = file.name.toLowerCase();
    const buf = await file.arrayBuffer();
    let rows = [];
    if(name.endsWith('.csv')){
      const text = new TextDecoder().decode(buf);
      const parsed = parseCSV(text);
      if(!parsed.length) throw new Error('קובץ לא תקין');
      const headers = parsed[0].map(h => String(h).trim());
      validateColumns(headers.map(h => COLUMN_ALIASES[h] || h));
      rows = parsed.slice(1).map(r => normalizeImportedRow(objFromRow(headers, r)));
    } else {
      const wb = XLSX.read(buf, { type: 'array', cellDates: true });
      let ws = null;
      if(wb.SheetNames && wb.SheetNames.length){
        const sheetName = wb.SheetNames.find(n => n === 'events') || wb.SheetNames[0];
        ws = wb.Sheets[sheetName];
      }
      if(!ws) throw new Error('קובץ לא תקין');
      const json = XLSX.utils.sheet_to_json(ws, { defval: '' });
      if(!json.length) throw new Error('אין נתונים בקובץ');
      const sample = normalizeImportedRow(json[0]);
      const cols = Object.keys(sample);
      validateColumns(cols);
      rows = json.map(normalizeImportedRow);
    }
    const count = eventsStore.toArray().length;
    if(count > 0){
      const ok = confirm(`הטעינה תחליף את ${count} האירועים הקיימים ב-${rows.length} אירועים מהקובץ.\nלהמשיך?`);
      if(!ok) return { imported: 0, cancelled: true };
    }
    eventsStore.loadFromArray(rows);
    eventsStore.saveToLocalStorage();
    return { imported: eventsStore.toArray().length, cancelled: false };
  }
  function exportExcel(){
    if(typeof XLSX === 'undefined') throw new Error('ספריית Excel לא נטענה. ודאו שקובץ xlsx.full.min.js נמצא בתיקייה.');
    const data = eventsStore.toArray().map(e => ({
      id: e.id,
      date: e.date,
      title: e.title,
      description: e.description || '',
      category: e.category || '',
      tags: e.tags || '',
      mediaSource: e.mediaSource || '',
      webSource: e.webSource || '',
      officialDocument: e.officialDocument || '',
      status: e.status || 'public'
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'events');
    const ab = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
    const blob = new Blob([ab], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    AppUtils.downloadBlob('events.xlsx', blob);
  }
  function objFromRow(headers, row){
    const o = {};
    for(let i=0;i<headers.length;i++) o[String(headers[i]).trim()] = row[i] ?? '';
    return o;
  }
  function parseCSV(text){
    const rows = [];
    let cur = [];
    let val = '';
    let inQ = false;
    for(let i=0;i<text.length;i++){
      const ch = text[i];
      if(inQ){
        if(ch==='"'){
          if(text[i+1]==='"'){ val+='"'; i++; }
          else inQ=false;
        } else val+=ch;
      } else {
        if(ch==='"') inQ=true;
        else if(ch===','){ cur.push(val); val=''; }
        else if(ch==='\n'){ cur.push(val); rows.push(cur); cur=[]; val=''; }
        else if(ch==='\r'){ /* ignore */ }
        else val+=ch;
      }
    }
    if(val!=='' || cur.length) { cur.push(val); rows.push(cur); }
    return rows;
  }
  return { importExcel, exportExcel };
})();

/* =============================================================
   STATE/UI (Router + Components)
============================================================= */
(function(){
  const monthsHe = ['ינואר','פברואר','מרץ','אפריל','מאי','יוני','יולי','אוגוסט','ספטמבר','אוקטובר','נובמבר','דצמבר'];

  const els = {
    yearInput: document.getElementById('yearInput'),
    currentYearDisplay: document.getElementById('currentYearDisplay'),
    headerInfo: document.getElementById('headerInfo'),
    monthsContainer: document.querySelector('.months-container'),
    monthView: document.getElementById('monthView'),
    yearView: document.getElementById('yearView'),
    backToYear: document.getElementById('backToYear'),
    prevMonth: document.getElementById('prevMonth'),
    nextMonth: document.getElementById('nextMonth'),
    currentMonthDisplay: document.getElementById('currentMonthDisplay'),
    daysContainer: document.querySelector('.days-container'),
    eventsModal: document.getElementById('eventsModal'),
    eventsList: document.getElementById('eventsList'),
    noEventsInDate: document.getElementById('noEventsInDate'),
    modalCloseBtns: document.querySelectorAll('.btn-close'),
    viewerSection: document.getElementById('viewer'),
    adminSection: document.getElementById('admin'),
    viewerBtn: document.getElementById('viewerBtn'),
    adminBtn: document.getElementById('adminBtn'),
    
    globalSearch: document.getElementById('globalSearch'),
    categoryFilter: document.getElementById('categoryFilter'),
    eventsTableBody: document.querySelector('#eventsTable tbody'),
    pendingEventsSection: document.getElementById('pendingEventsSection'),
    pendingEventsTableBody: document.querySelector('#pendingEventsTableBody'),
    addEventBtn: document.getElementById('addEventBtn'),
    assistantLinkBtn: document.getElementById('assistantLinkBtn'),
    assistantLinkSettingsBtn: document.getElementById('assistantLinkSettingsBtn'),
    assistantLinkModal: document.getElementById('assistantLinkModal'),
    assistantLinkLabelInput: document.getElementById('assistantLinkLabelInput'),
    assistantLinkUrlInput: document.getElementById('assistantLinkUrlInput'),
    saveAssistantLinkBtn: document.getElementById('saveAssistantLinkBtn'),
    cancelAssistantLinkBtn: document.getElementById('cancelAssistantLinkBtn'),
    themeAccentSettingsBtn: document.getElementById('themeAccentSettingsBtn'),
    themeAccentModal: document.getElementById('themeAccentModal'),
    themeAccentInput: document.getElementById('themeAccentInput'),
    themeAccentHexInput: document.getElementById('themeAccentHexInput'),
    siteTitleInput: document.getElementById('siteTitleInput'),
    siteTitleDisplay: document.getElementById('siteTitleDisplay'),
    saveThemeAccentBtn: document.getElementById('saveThemeAccentBtn'),
    cancelThemeAccentBtn: document.getElementById('cancelThemeAccentBtn'),
    resetThemeAccentBtn: document.getElementById('resetThemeAccentBtn'),
    eventFormModal: document.getElementById('eventFormModal'),
    eventForm: document.getElementById('eventForm'),
    cancelForm: document.getElementById('cancelForm'),
    categoriesList: document.getElementById('categoriesList'),
    noEventsMsg: document.getElementById('noEventsMsg'),
    prevMonth: document.getElementById('prevMonth'),
    nextMonth: document.getElementById('nextMonth'),
    passwordModal: document.getElementById('passwordModal'),
    passwordForm: document.getElementById('passwordForm'),
    passwordInput: document.getElementById('passwordInput'),
    passwordError: document.getElementById('passwordError'),
    cancelPassword: document.getElementById('cancelPassword'),
    excelFileInput: document.getElementById('excelFileInput'),
    exportExcelBtn: document.getElementById('exportExcelBtn'),
    jsonFileInput: document.getElementById('jsonFileInput'),
    downloadJsonBtn: document.getElementById('downloadJsonBtn')
  };
  // Chronological list element
  els.eventsChronoList = document.getElementById('eventsChronoList');
  els.eventsChrono = document.querySelector('.events-chrono');
  els.chronoTitle = document.getElementById('chronoTitle');
  els.chronoCount = document.getElementById('chronoCount');
  els.allYearsToggle = document.getElementById('allYearsToggle');
  els.yearSelectorWrap = document.getElementById('yearSelectorWrap');
  els.yearOnlyWrap = document.getElementById('yearOnlyWrap');

  function updateYearInputState(){
    if(!els.allYearsToggle) return;
    const all = !!els.allYearsToggle.checked;
    els.yearInput.disabled = all;
    if(els.yearOnlyWrap){ els.yearOnlyWrap.style.display = all ? 'none' : ''; }
  }

  // Month view state
  let currentViewYear = Number(new Date().getFullYear());
  let currentViewMonth = 0; // 0-11
  
  // Table sorting state
  let sortColumn = 'date'; // Default to sorting by date
  let sortDirection = 'desc'; // Default to descending (newest first)

  const ASSISTANT_LINK_DEFAULTS = {
    label: 'האסיסטנט שלנו',
    url: ''
  };
  const THEME_ACCENT_DEFAULT = '#e07020';
  const SITE_TITLE_DEFAULT = 'חודשים מספרים';
  let themeAccentPreviewBeforeOpen = THEME_ACCENT_DEFAULT;
  let siteTitlePreviewBeforeOpen = SITE_TITLE_DEFAULT;

  function clampByte(n){
    return Math.max(0, Math.min(255, Math.round(n)));
  }
  function normalizeHexColor(raw, fallback){
    const fb = fallback || THEME_ACCENT_DEFAULT;
    if(!raw || typeof raw !== 'string') return fb;
    let h = raw.trim().toLowerCase();
    if(/^#[0-9a-f]{3}$/.test(h)){
      h = '#' + h[1]+h[1]+h[2]+h[2]+h[3]+h[3];
    }
    if(!/^#[0-9a-f]{6}$/.test(h)) return fb;
    return h;
  }
  function hexToRgb(hex){
    const h = normalizeHexColor(hex, THEME_ACCENT_DEFAULT).slice(1);
    return {
      r: parseInt(h.slice(0, 2), 16),
      g: parseInt(h.slice(2, 4), 16),
      b: parseInt(h.slice(4, 6), 16)
    };
  }
  function rgbToHex(r, g, b){
    return '#' + [r, g, b].map(v => clampByte(v).toString(16).padStart(2, '0')).join('');
  }
  function mixToward(hex, target, amount){
    const a = hexToRgb(hex);
    const t = hexToRgb(target);
    return rgbToHex(
      a.r + (t.r - a.r) * amount,
      a.g + (t.g - a.g) * amount,
      a.b + (t.b - a.b) * amount
    );
  }
  function loadThemeAccent(){
    try {
      const raw = localStorage.getItem('themeAccent');
      if(raw) return normalizeHexColor(raw, THEME_ACCENT_DEFAULT);
    } catch(_) {}
    return THEME_ACCENT_DEFAULT;
  }
  function saveThemeAccent(hex){
    localStorage.setItem('themeAccent', normalizeHexColor(hex, THEME_ACCENT_DEFAULT));
  }
  function applyThemeAccent(hex){
    const color = normalizeHexColor(hex, THEME_ACCENT_DEFAULT);
    const { r, g, b } = hexToRgb(color);
    const root = document.documentElement;
    root.style.setProperty('--primary', color);
    root.style.setProperty('--primary-rgb', `${r}, ${g}, ${b}`);
    root.style.setProperty('--primary-dark', mixToward(color, '#000000', 0.28));
    root.style.setProperty('--primary-deeper', mixToward(color, '#000000', 0.18));
    root.style.setProperty('--primary-bright', mixToward(color, '#ffffff', 0.22));
    root.style.setProperty('--primary-soft', mixToward(color, '#ffffff', 0.35));
    root.style.setProperty('--primary-soft-2', mixToward(color, '#ffffff', 0.48));
    root.style.setProperty('--primary-light', `rgba(${r}, ${g}, ${b}, 0.14)`);
  }
  function applyThemeAccentFromFile(raw){
    if(raw == null) return;
    const hex = normalizeHexColor(raw, THEME_ACCENT_DEFAULT);
    saveThemeAccent(hex);
    applyThemeAccent(hex);
  }
  function fillThemeAccentAdminForm(){
    const hex = loadThemeAccent();
    if(els.themeAccentInput) els.themeAccentInput.value = hex;
    if(els.themeAccentHexInput) els.themeAccentHexInput.value = hex.slice(1);
    if(els.siteTitleInput) els.siteTitleInput.value = loadSiteTitle();
  }

  function normalizeSiteTitle(raw){
    const t = (raw == null ? '' : String(raw)).trim();
    return t || SITE_TITLE_DEFAULT;
  }
  function loadSiteTitle(){
    try {
      const raw = localStorage.getItem('siteTitle');
      if(raw != null) return normalizeSiteTitle(raw);
    } catch(_) {}
    return SITE_TITLE_DEFAULT;
  }
  function saveSiteTitle(title){
    localStorage.setItem('siteTitle', normalizeSiteTitle(title));
  }
  function applySiteTitle(title){
    const name = normalizeSiteTitle(title);
    if(els.siteTitleDisplay){
      els.siteTitleDisplay.textContent = '📅 ' + name;
    }
    document.title = 'לוח שנה היסטורי - ' + name;
  }
  function applySiteTitleFromFile(raw){
    if(raw == null) return;
    const name = normalizeSiteTitle(raw);
    saveSiteTitle(name);
    applySiteTitle(name);
  }

  function setThemeAccentFormColor(hex){
    const color = normalizeHexColor(hex, THEME_ACCENT_DEFAULT);
    if(els.themeAccentInput) els.themeAccentInput.value = color;
    if(els.themeAccentHexInput) els.themeAccentHexInput.value = color.slice(1);
    applyThemeAccent(color);
  }

  function readThemeAccentFromHexField(){
    let raw = (els.themeAccentHexInput?.value || '').trim().replace(/^#/, '');
    raw = raw.replace(/[^0-9a-fA-F]/g, '').slice(0, 6);
    if(els.themeAccentHexInput && els.themeAccentHexInput.value !== raw){
      els.themeAccentHexInput.value = raw;
    }
    if(raw.length === 3 || raw.length === 6){
      return normalizeHexColor('#' + raw, null);
    }
    return null;
  }

  function normalizeAssistantLinkSettings(raw){
    if(!raw || typeof raw !== 'object') return { ...ASSISTANT_LINK_DEFAULTS };
    return {
      label: (raw.label || ASSISTANT_LINK_DEFAULTS.label).trim() || ASSISTANT_LINK_DEFAULTS.label,
      url: (raw.url || '').trim()
    };
  }

  function loadAssistantLinkSettings(){
    try {
      const raw = localStorage.getItem('assistantLinkSettings');
      if(raw) return normalizeAssistantLinkSettings(JSON.parse(raw));
    } catch(_) {}
    return { ...ASSISTANT_LINK_DEFAULTS };
  }

  function saveAssistantLinkSettings(settings){
    localStorage.setItem('assistantLinkSettings', JSON.stringify(normalizeAssistantLinkSettings(settings)));
  }

  function applyAssistantLinkFromFile(raw){
    if(raw == null) return;
    saveAssistantLinkSettings(raw);
    applyAssistantLinkButton();
  }

  function parseEventsFileData(data){
    if(Array.isArray(data)){
      return { events: data, assistantLink: null, themeAccent: null, siteTitle: null };
    }
    if(data && typeof data === 'object'){
      return {
        events: Array.isArray(data.events) ? data.events : [],
        assistantLink: data.assistantLink != null ? data.assistantLink : null,
        themeAccent: data.themeAccent != null ? data.themeAccent : null,
        siteTitle: data.siteTitle != null ? data.siteTitle : null
      };
    }
    return { events: [], assistantLink: null, themeAccent: null, siteTitle: null };
  }

  function buildEventsFilePayload(){
    return {
      assistantLink: loadAssistantLinkSettings(),
      themeAccent: loadThemeAccent(),
      siteTitle: loadSiteTitle(),
      events: eventsStore.toArray()
    };
  }

  function downloadEventsJsonFile(){
    const json = JSON.stringify(buildEventsFilePayload(), null, 2);
    const blob = new Blob([json], { type: 'application/json' });
    AppUtils.downloadBlob('events.json', blob);
  }

  function applyAssistantLinkButton(){
    const settings = loadAssistantLinkSettings();
    const btn = els.assistantLinkBtn;
    if(!btn) return;
    btn.textContent = settings.label;
    const hasUrl = !!settings.url;
    if(hasUrl){
      btn.href = settings.url;
      btn.classList.remove('is-disabled');
      btn.setAttribute('aria-disabled', 'false');
    } else {
      btn.href = '#';
      btn.classList.add('is-disabled');
      btn.setAttribute('aria-disabled', 'true');
    }
  }

  function fillAssistantLinkAdminForm(){
    const settings = loadAssistantLinkSettings();
    if(els.assistantLinkLabelInput) els.assistantLinkLabelInput.value = settings.label;
    if(els.assistantLinkUrlInput) els.assistantLinkUrlInput.value = settings.url;
  }

  // Router
  function navigate(hash){
    if(hash === '#/admin'){
      els.viewerSection.classList.remove('active');
      els.adminSection.classList.add('active');
      els.viewerBtn.classList.remove('active');
      els.viewerBtn.style.display = 'block';
      els.adminBtn.classList.add('active');
      if(els.assistantLinkBtn) els.assistantLinkBtn.style.display = 'none';
      if(els.addEventBtn) els.addEventBtn.style.display = 'block';
      renderAdminTable();
    } else {
      els.adminSection.classList.remove('active');
      els.viewerSection.classList.add('active');
      els.adminBtn.classList.remove('active');
      els.viewerBtn.classList.add('active');
      els.viewerBtn.style.display = 'none';
      if(els.assistantLinkBtn) els.assistantLinkBtn.style.display = 'inline-flex';
      if(els.addEventBtn) els.addEventBtn.style.display = 'none';
      applyAssistantLinkButton();
      renderYearGrid(Number(els.yearInput.value));
    }
  }

  window.addEventListener('hashchange', ()=>navigate(location.hash));

  // Viewer
  function getMonthDays(year, month){ // month 0-based
    const first = new Date(Date.UTC(year, month, 1));
    const startDay = (first.getUTCDay()+1)%7; // make Sunday=6, Monday=0 ... so grid A-ש
    const daysInMonth = new Date(Date.UTC(year, month+1, 0)).getUTCDate();
    return { startDay, daysInMonth };
  }
  function renderYearGrid(year){
    // Show explanation in header when "all years" is checked, otherwise show year
    const allYears = els.allYearsToggle && els.allYearsToggle.checked;
    if(els.headerInfo){
      if(allYears){
        els.headerInfo.style.display = 'flex';
      } else {
        els.headerInfo.style.display = 'none';
      }
    }
    if(els.currentYearDisplay){
      // Always hide the year display above months - year is shown in header selector
      els.currentYearDisplay.style.display = 'none';
    }
    els.monthsContainer.innerHTML = '';
    const q = els.globalSearch ? els.globalSearch.value : '';
    const cat = els.categoryFilter ? els.categoryFilter.value : '';
    let baseEvents = eventsStore.search(q, { category: cat, tags: '' });
    for(let m=0;m<12;m++){
      const card = document.createElement('button');
      card.className = 'month-card month-' + m;
      card.setAttribute('aria-label', `פתח ${monthsHe[m]} ${year}`);

      const side = document.createElement('div');
      side.className = 'month-side';
      const num = document.createElement('div');
      num.className = 'month-num';
      num.textContent = String(m + 1);
      const title = document.createElement('div');
      title.className = 'month-title';
      title.textContent = monthsHe[m];
      side.appendChild(num);
      side.appendChild(title);

      const grid = document.createElement('div'); grid.className='days-mini';
      // show correct number of days for the month (handles leap years)
      const { daysInMonth } = getMonthDays(year, m);
      // collect days with events for this month (DD-MM-YYYY format only),
      // respect year filter if "all years" is unchecked
      const allYears = !(els.allYearsToggle) || !!els.allYearsToggle.checked;
      const selectedYear = Number(els.yearInput.value);
      const hasByDay = new Set(
        baseEvents
          .filter(e=>/^\d{2}-\d{2}-\d{4}$/.test(e.date)
            && Number(e.date.slice(3,5))===m+1
            && (allYears || Number(e.date.slice(-4))===selectedYear))
          .map(e=>Number(e.date.slice(0,2)))
      );
      for(let i=1;i<=daysInMonth;i++){
        const d = document.createElement('div');
        const has = hasByDay.has(i);
        d.className = 'day-dot' + (has ? ' has-events' : '');
        d.textContent = i;
        if(has){
          const mm = String(m+1).padStart(2,'0');
          const dd = String(i).padStart(2,'0');
          const mmdd = `${mm}-${dd}`;
          let events = baseEvents.filter(ev=>AppUtils.toMMDD(ev.date)===mmdd);
          if(!allYears){ events = events.filter(ev => /^\d{2}-\d{2}-\d{4}$/.test(ev.date) && Number(ev.date.slice(-4))===selectedYear); }
          // Always show tooltip for days with events - show all events for this date
          if(events.length > 0){
            const lines = events.map(ev => `${ev.title} (${/^\d{4}$/.test(ev.date)?ev.date:ev.date.slice(-4)})`);
            attachTooltip(d, lines);
          }
        }
        grid.appendChild(d);
      }

      card.appendChild(side);
      card.appendChild(grid);
      card.addEventListener('click', ()=>openMonthView(year, m));
      els.monthsContainer.appendChild(card);
    }
    populateCategoryOptions();
    renderChronoList();
    alignChronoTop();
  }

  // Chronological list
  function renderChronoList(){
    if(!els.eventsChronoList) return;
    const q = els.globalSearch ? els.globalSearch.value.trim() : '';
    const cat = els.categoryFilter ? els.categoryFilter.value.trim() : '';
    let data = eventsStore.search(q, { category: cat, tags: '' });
    
    // Check if we're in month view - if so, filter by month
    const isMonthView = els.monthView && els.monthView.style.display !== 'none';
    if(isMonthView){
      const targetMonth = currentViewMonth + 1; // month is 0-based, but date month is 1-based
      data = data.filter(e=>{
        if(/^\d{2}-\d{2}-\d{4}$/.test(e.date)){
          const month = Number(e.date.slice(3,5));
          return month === targetMonth;
        }
        // Year-only events don't have a month, so exclude them in month view
        return false;
      });
    }
    
    if(els.allYearsToggle && !els.allYearsToggle.checked){
      const y = Number(els.yearInput.value);
      data = data.filter(e=>{
        if(/^\d{2}-\d{2}-\d{4}$/.test(e.date)) return Number(e.date.slice(-4))===y;
        if(/^\d{4}$/.test(e.date)) return Number(e.date)===y;
        return true;
      });
    }
    // Update title based on filters and view
    if(els.chronoTitle){
      let title = 'כל האירועים';
      const hasSearch = q.length > 0;
      const hasCategory = cat.length > 0;
      const allYears = !(els.allYearsToggle) || !!els.allYearsToggle.checked;
      const selectedYear = Number(els.yearInput?.value);
      if(isMonthView){
        title = `אירועים ב${monthsHe[currentViewMonth]}`;
        if(!allYears && selectedYear) title = `אירועי ${selectedYear} ב${monthsHe[currentViewMonth]}`;
        if(hasSearch || hasCategory){
          title += ' (מסונן)';
        }
      } else if(hasSearch && hasCategory){
        title = `תוצאות חיפוש: ${cat}`;
      } else if(hasSearch){
        title = 'תוצאות חיפוש';
      } else if(hasCategory){
        title = `אירועים: ${cat}`;
      } else if(!allYears && selectedYear){
        title = `אירועי ${selectedYear}`;
      }
      els.chronoTitle.textContent = title;
    }
    if(els.chronoCount){
      // ספירה לפי היקף השנה (כל המערכת / שנה נבחרת), בלי חיפוש/קטגוריה
      let countSource = eventsStore.search('', { category: '', tags: '' });
      const allYears = !(els.allYearsToggle) || !!els.allYearsToggle.checked;
      if(!allYears){
        const y = Number(els.yearInput.value);
        countSource = countSource.filter(e=>{
          if(/^\d{2}-\d{2}-\d{4}$/.test(e.date)) return Number(e.date.slice(-4))===y;
          if(/^\d{4}$/.test(e.date)) return Number(e.date)===y;
          return false;
        });
      }
      const count = countSource.length;
      els.chronoCount.textContent = count === 1 ? 'אירוע אחד' : `${count} אירועים`;
    }
    const withKey = data.map(e=>{
      let y=0,m=0,d=0;
      if(/^\d{2}-\d{2}-\d{4}$/.test(e.date)){
        const [dd,mm,yyyy] = e.date.split('-').map(Number);
        y=yyyy; m=mm; d=dd;
      } else if(/^\d{4}$/.test(e.date)){
        y=Number(e.date); m=0; d=0;
      }
      return { y,m,d,e };
    });
    withKey.sort((a,b)=> a.y-b.y || a.m-b.m || a.d-b.d);
    els.eventsChronoList.innerHTML = '';
    for(const {e} of withKey){
      const li = document.createElement('li');
      const dateSpan = document.createElement('span'); dateSpan.className='evt-date'; dateSpan.textContent = e.date;
      const titleSpan = document.createElement('span'); titleSpan.textContent = e.title;
      li.appendChild(dateSpan); li.appendChild(titleSpan);
      // open the day modal on click if full date exists
      if(/^\d{2}-\d{2}-\d{4}$/.test(e.date)){
        const [dd,mm] = e.date.split('-');
        const mmdd = `${mm}-${dd}`;
        li.style.cursor = 'pointer';
        li.addEventListener('click', ()=> openEventsPanel(mmdd));
      }
      els.eventsChronoList.appendChild(li);
    }
    // Trigger highlight animation
    if(els.eventsChrono){
      // Remove class if it exists to restart animation
      els.eventsChrono.classList.remove('updated');
      // Force reflow to restart animation
      void els.eventsChrono.offsetWidth;
      // Add class to trigger animation
      els.eventsChrono.classList.add('updated');
      // Remove class after animation completes
      setTimeout(() => {
        if(els.eventsChrono) els.eventsChrono.classList.remove('updated');
      }, 1200);
    }
  }

  function alignChronoTop(){
    // No longer needed - chrono is on the right side now
  }
  function openMonthView(year, month){
    currentViewYear = year;
    currentViewMonth = month;
    els.yearView.style.display='none';
    els.monthView.style.display='block';
    if(els.eventsChrono) els.eventsChrono.style.display = '';
    // Add month class for color styling
    els.monthView.className = 'month-grid month-' + month;
    // Check if "all years" filter is enabled
    const allYears = els.allYearsToggle && els.allYearsToggle.checked;
    if(allYears){
      els.currentMonthDisplay.textContent = monthsHe[month];
    } else {
      els.currentMonthDisplay.textContent = `${monthsHe[month]} ${year}`;
    }
    renderMonthView(year, month);
    renderChronoList();
  }
  function navigateMonth(direction){
    // direction: -1 for previous, 1 for next
    let newMonth = currentViewMonth + direction;
    let newYear = currentViewYear;
    
    if(newMonth < 0){
      newMonth = 11;
      newYear--;
    } else if(newMonth > 11){
      newMonth = 0;
      newYear++;
    }
    
    // Update year input if needed
    if(els.yearInput && Number(els.yearInput.value) !== newYear){
      els.yearInput.value = String(newYear);
    }
    
    openMonthView(newYear, newMonth);
  }
  function backToYear(){ els.monthView.style.display='none'; els.yearView.style.display='block'; }
  const _origBackToYear = backToYear;
  backToYear = function(){
    els.monthView.style.display='none';
    els.yearView.style.display='block';
    if(els.eventsChrono) els.eventsChrono.style.display = '';
  }
  function renderMonthView(year, month){
    const { startDay, daysInMonth } = getMonthDays(year, month);
    els.daysContainer.innerHTML='';
    // Apply search and filter
    const q = els.globalSearch ? els.globalSearch.value : '';
    const cat = els.categoryFilter ? els.categoryFilter.value : '';
    const filteredEvents = eventsStore.search(q, { category: cat, tags: '' });
    
    // Check if "all years" filter is enabled
    const allYears = els.allYearsToggle && els.allYearsToggle.checked;
    
    // fill leading blanks
    for(let i=0;i<startDay;i++){ const b=document.createElement('div'); b.className='day-cell blank'; els.daysContainer.appendChild(b); }
    for(let day=1; day<=daysInMonth; day++){
      const cell = document.createElement('div'); cell.className='day-cell';
      const mm = String(month+1).padStart(2,'0'); const dd = String(day).padStart(2,'0');
      const mmdd = `${mm}-${dd}`;
      
      // If all years is enabled, show events from all years for this month-day
      // Otherwise, filter by year as well
      let events;
      if(allYears){
        events = filteredEvents.filter(e=>AppUtils.toMMDD(e.date)===mmdd);
      } else {
        events = filteredEvents.filter(e=>{
          const eMMDD = AppUtils.toMMDD(e.date);
          if(eMMDD !== mmdd) return false;
          // Check year match
          const eYear = AppUtils.extractYear(e.date);
          return eYear === year;
        });
      }
      
      const num = document.createElement('div'); num.className='num'; num.textContent = String(day);
      // Mark cell if it has events
      if(events.length > 0){
        cell.classList.add('has-events');
      }
      cell.appendChild(num);
      
      // Add events list inside cell
      const eventsContainer = document.createElement('div'); eventsContainer.className='events-in-cell';
      for(const evt of events){
        const eventDiv = document.createElement('div'); eventDiv.className='event-item-cell';
        // Extract date parts - show full date DD-MM-YYYY or year if year-only
        let dateDisplay;
        if(/^\d{4}$/.test(evt.date)){
          dateDisplay = evt.date; // Year only
        } else if(/^\d{2}-\d{2}-\d{4}$/.test(evt.date)){
          dateDisplay = evt.date; // Full date DD-MM-YYYY
        } else {
          const dayMonth = evt.date.slice(0,5); // DD-MM as fallback
          const year = AppUtils.extractYear(evt.date) || '';
          dateDisplay = year ? `${dayMonth}-${year}` : dayMonth;
        }
        eventDiv.textContent = `${evt.title} · ${dateDisplay}`;
        // Add tooltip to each individual event item in month view
        const eventTooltip = `${evt.title}\nתאריך: ${evt.date}${evt.description ? '\n' + evt.description : ''}`;
        attachTooltip(eventDiv, [eventTooltip]);
        eventDiv.addEventListener('click', (e)=> {
          e.stopPropagation(); // don't trigger cell click
          openEventsPanel(mmdd);
        });
        eventsContainer.appendChild(eventDiv);
      }
      cell.appendChild(eventsContainer);
      
      cell.addEventListener('click', ()=>openEventsPanel(mmdd));
      els.daysContainer.appendChild(cell);
    }
    // Update chronological list to show only this month's events
    renderChronoList();
  }
  function openEventsPanel(mmdd){
    // Apply search and filter
    const q = els.globalSearch ? els.globalSearch.value : '';
    const cat = els.categoryFilter ? els.categoryFilter.value : '';
    const filteredEvents = eventsStore.search(q, { category: cat, tags: '' });
    const list = filteredEvents.filter(e=>AppUtils.toMMDD(e.date)===mmdd);
    els.eventsList.innerHTML='';
    if(!list.length){ els.noEventsInDate.style.display='block'; }
    else { els.noEventsInDate.style.display='none'; }
    for(const e of list){
      const li = document.createElement('li'); li.className='event-item';
      // Render strictly by object key order (as constructed in normalizeEvent)
      const labels = {
        date: 'תאריך',
        title: 'כותרת',
        description: 'תיאור',
        category: 'קטגוריה',
        tags: 'תגיות',
        mediaSource: 'מקור מדיה',
        webSource: 'מקור אתר',
        officialDocument: 'מסמך רשמי'
      };
      for(const [key, value] of Object.entries(e)){
        if(key === 'id' || key === 'status') continue; // Skip id and status fields
        if(value === undefined || value === null || String(value).trim() === '') continue;
        const div = document.createElement('div');
        if(key === 'title') { div.className = 'title'; }
        else if(key === 'description') { div.className = 'description'; }
        else if(key === 'category' || key === 'tags' || key.endsWith('Source') || key === 'officialDocument') { div.className = 'meta'; }
        const label = labels[key];
        if(label){
          // For title and description, show only the value without label
          if(key === 'title' || key === 'description'){
            div.innerHTML = escapeHTML(String(value));
          }
          // Check if this is a URL field and make it a clickable link
          else if((key === 'webSource' || key === 'mediaSource' || key === 'officialDocument') && String(value).trim() !== ''){
            const url = String(value).trim();
            // Check if it's already a full URL, if not, try to make it one
            let href = url;
            if(!url.match(/^https?:\/\//i)){
              href = 'https://' + url;
            }
            div.innerHTML = `${label}: <a href="${escapeHTML(href)}" target="_blank" rel="noopener noreferrer" class="event-link">${escapeHTML(url)}</a>`;
          } else {
            div.innerHTML = `${label}: <span>${escapeHTML(String(value))}</span>`;
          }
        } else {
          div.textContent = String(value);
        }
        li.appendChild(div);
      }
      els.eventsList.appendChild(li);
    }
    const [m,d] = mmdd.split('-');
    const titleStr = `אירועים בתאריך ${Number(d)}-${Number(m)}`;
    document.getElementById('modalTitle').textContent = titleStr;
    openDialog(els.eventsModal);
  }

  // Admin
  function renderPendingEvents(){
    if(!els.pendingEventsTableBody || !els.pendingEventsSection) return;
    const pendingEvents = eventsStore.toArray().filter(e => (e.status || 'public') === 'WaitingToApprove');
    els.pendingEventsTableBody.innerHTML = '';
    if(pendingEvents.length === 0){
      els.pendingEventsSection.style.display = 'none';
      return;
    }
    els.pendingEventsSection.style.display = 'block';
    for(const e of pendingEvents){
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${e.date}</td>
        <td>${escapeHTML(e.title)}</td>
        <td>${escapeHTML(e.description||'')}</td>
        <td>${escapeHTML(e.category||'')}</td>
        <td>${escapeHTML(e.submitterName || '')}<br>${escapeHTML(e.submitterId || '')}</td>
        <td>
          <div class="table-actions">
            <button class="btn btn-secondary" data-act="edit">✏️ עריכה</button>
            <button class="btn btn-success" data-act="approve">✅ אישור</button>
            <button class="btn" data-act="reject">❌ דחייה</button>
          </div>
        </td>`;
      tr.querySelector('[data-act="edit"]').addEventListener('click', ()=>{
        openEventForm('edit', e);
      });
      tr.querySelector('[data-act="approve"]').addEventListener('click', ()=>{
        eventsStore.update(e.id, { status: 'public' });
        // Download updated events.json automatically
        try {
          downloadEventsJsonFile();
          toast('האירוע אושר וזמין לכולם. הקובץ events.json הורד - העלה אותו לשרת', 'success');
        } catch(err) {
          console.error('Failed to download events.json:', err);
          toast('האירוע אושר וזמין לכולם', 'success');
        }
        renderPendingEvents();
        renderAdminTable();
        renderYearGrid(Number(els.yearInput.value));
        renderChronoList();
      });
      tr.querySelector('[data-act="reject"]').addEventListener('click', ()=>{
        if(confirm('האם למחוק את האירוע המוצע?')){
          eventsStore.remove(e.id);
          // Download updated events.json automatically
          try {
            downloadEventsJsonFile();
            toast('האירוע נדחה ונמחק. הקובץ events.json הורד - העלה אותו לשרת', 'success');
          } catch(err) {
            console.error('Failed to download events.json:', err);
            toast('האירוע נדחה ונמחק', 'success');
          }
          renderPendingEvents();
          renderAdminTable();
          renderYearGrid(Number(els.yearInput.value));
          renderChronoList();
        }
      });
      els.pendingEventsTableBody.appendChild(tr);
    }
  }
  function sortTableData(data, column, direction){
    if(!column) return data;
    const sorted = [...data].sort((a, b)=>{
      let valA = (a[column] || '').toString().trim();
      let valB = (b[column] || '').toString().trim();
      
      // Special handling for date column
      if(column === 'date'){
        // Extract year for comparison
        let yearA = 0, yearB = 0;
        if(/^\d{4}$/.test(valA)) yearA = Number(valA);
        else if(/^\d{2}-\d{2}-\d{4}$/.test(valA)) yearA = Number(valA.slice(-4));
        if(/^\d{4}$/.test(valB)) yearB = Number(valB);
        else if(/^\d{2}-\d{2}-\d{4}$/.test(valB)) yearB = Number(valB.slice(-4));
        
        if(yearA !== yearB) return direction === 'asc' ? yearA - yearB : yearB - yearA;
        // If same year, compare full date string
        return direction === 'asc' ? valA.localeCompare(valB) : valB.localeCompare(valA);
      }
      
      // String comparison for other columns
      const comparison = valA.localeCompare(valB, 'he');
      return direction === 'asc' ? comparison : -comparison;
    });
    return sorted;
  }
  function updateSortIndicators(){
    const headers = document.querySelectorAll('#eventsTable thead th.sortable');
    headers.forEach(th => {
      const arrow = th.querySelector('.sort-arrow');
      if(!arrow) return;
      const column = th.getAttribute('data-sort');
      if(column === sortColumn){
        // Use Unicode arrows: ▲ for ascending, ▼ for descending
        arrow.textContent = sortDirection === 'asc' ? ' ▲' : ' ▼';
        arrow.style.display = 'inline';
        arrow.style.opacity = '1';
      } else {
        arrow.textContent = '';
        arrow.style.display = 'none';
      }
    });
  }
  function handleSortClick(column){
    if(sortColumn === column){
      // Toggle direction if clicking same column
      sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
    } else {
      // New column, start with ascending
      sortColumn = column;
      sortDirection = 'asc';
    }
    updateSortIndicators();
    renderAdminTable();
  }
  function renderAdminTable(){
    const q = els.globalSearch ? els.globalSearch.value : '';
    const cat = els.categoryFilter.value;
    const tags = '';
    // Show all approved events
    let data = eventsStore.search(q, {category: cat, tags, showPending: true}).filter(e => (e.status || 'public') === 'public');
    if(els.allYearsToggle && !els.allYearsToggle.checked){
      const y = Number(els.yearInput.value);
      data = data.filter(e=>{
        if(/^\d{2}-\d{2}-\d{4}$/.test(e.date)) return Number(e.date.slice(-4))===y;
        if(/^\d{4}$/.test(e.date)) return Number(e.date)===y;
        return true;
      });
    }
    // Sort data if sort column is set
    if(sortColumn){
      data = sortTableData(data, sortColumn, sortDirection);
    }
    els.eventsTableBody.innerHTML='';
    if(!data.length){ els.noEventsMsg.style.display='block'; }
    else { els.noEventsMsg.style.display='none'; }
    for(const e of data){
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${e.date}</td>
        <td>${escapeHTML(e.title)}</td>
        <td>${escapeHTML(e.description||'')}</td>
        <td>${escapeHTML(e.category||'')}</td>
        <td>${escapeHTML(e.tags||'')}</td>
        <td>${escapeHTML(e.mediaSource||'')}</td>
        <td>${escapeHTML(e.webSource||'')}</td>
        <td>${escapeHTML(e.officialDocument||'')}</td>
        <td>
          <div class="table-actions">
            <button class="btn btn-secondary" data-act="edit">ערוך</button>
            <button class="btn" data-act="del">מחק</button>
          </div>
        </td>`;
      tr.querySelector('[data-act="edit"]').addEventListener('click', ()=>openEventForm('edit', e));
      tr.querySelector('[data-act="del"]').addEventListener('click', ()=>confirmDelete(e.id));
      els.eventsTableBody.appendChild(tr);
    }
    updateSortIndicators();
    renderPendingEvents();
    populateCategoryOptions();
    updateExportButtons();
  }
  function downloadEventsJSON(){
    try {
      downloadEventsJsonFile();
      toast('הקובץ events.json הורד בהצלחה', 'success');
    } catch(err) {
      toast('שגיאה בהורדת הקובץ', 'error');
      console.error(err);
    }
  }
  async function importEventsJSON(file){
    const text = await file.text();
    const data = JSON.parse(text);
    const parsed = parseEventsFileData(data);
    const arr = parsed.events;
    if(!arr.length) throw new Error('קובץ JSON ריק או לא תקין');
    const count = eventsStore.toArray().length;
    if(count > 0){
      const ok = confirm(`הטעינה תחליף את ${count} האירועים הקיימים ב-${arr.length} אירועים מהקובץ.\nלהמשיך?`);
      if(!ok) return { imported: 0, cancelled: true };
    }
    eventsStore.loadFromArray(arr);
    eventsStore.saveToLocalStorage();
    applyAssistantLinkFromFile(parsed.assistantLink);
    applyThemeAccentFromFile(parsed.themeAccent);
    applySiteTitleFromFile(parsed.siteTitle);
    return { imported: eventsStore.toArray().length, cancelled: false };
  }
  function refreshAllViews(){
    renderAdminTable();
    renderYearGrid(Number(els.yearInput.value));
    if(els.monthView.style.display === 'block') renderMonthView(currentViewYear, currentViewMonth);
    renderChronoList();
    populateCategoryOptions();
    updateExportButtons();
  }
  function populateCategoryOptions(){
    const cats = Array.from(new Set(eventsStore.toArray().map(e=>e.category).filter(Boolean))).sort();
    const sel = els.categoryFilter;
    const previous = sel.value;
    sel.innerHTML = '<option value="">כל הקטגוריות</option>' + cats.map(c=>`<option value="${escapeHTML(c)}">${escapeHTML(c)}</option>`).join('');
    sel.value = previous;
    // datalist for form
    els.categoriesList.innerHTML = cats.map(c=>`<option value="${escapeHTML(c)}"></option>`).join('');
  }
  let currentFormMode = 'add'; // 'add', 'edit', 'suggest'
  function openEventForm(mode, data){
    const isEdit = mode==='edit';
    const isSuggest = mode==='suggest';
    currentFormMode = mode;
    document.getElementById('formTitle').textContent = isEdit ? 'ערוך אירוע' : (isSuggest ? 'הצע אירוע חדש' : 'הוסף אירוע');
    els.eventForm.reset();
    // Clear error messages
    const dateError = document.getElementById('formDateError');
    const titleError = document.getElementById('formTitleError');
    if(dateError) { dateError.style.display = 'none'; dateError.textContent = ''; }
    if(titleError) { titleError.style.display = 'none'; titleError.textContent = ''; }
    
    // Show/hide submitter info section
    const submitterSection = document.getElementById('submitterInfoSection');
    const submitterNameGroup = document.getElementById('submitterNameGroup');
    const submitterIdGroup = document.getElementById('submitterIdGroup');
    if(submitterSection){
      submitterSection.style.display = isSuggest ? 'block' : 'none';
    }
    if(submitterNameGroup){
      submitterNameGroup.style.display = isSuggest ? 'flex' : 'none';
    }
    if(submitterIdGroup){
      submitterIdGroup.style.display = isSuggest ? 'flex' : 'none';
    }
    // Make fields required only for suggest mode
    const submitterName = document.getElementById('formSubmitterName');
    const submitterId = document.getElementById('formSubmitterId');
    if(submitterName){
      submitterName.required = isSuggest;
      if(isSuggest && !submitterName.value) submitterName.value = '';
    }
    if(submitterId){
      submitterId.required = isSuggest;
      if(isSuggest && !submitterId.value) submitterId.value = '';
    }
    
    document.getElementById('eventId').value = isEdit ? data.id : '';
    document.getElementById('formDate').value = isEdit ? data.date : '';
    document.getElementById('formTitleInput').value = isEdit ? data.title : '';
    document.getElementById('formDescription').value = isEdit ? (data.description||'') : '';
    document.getElementById('formCategory').value = isEdit ? (data.category||'') : '';
    document.getElementById('formTags').value = isEdit ? (data.tags||'') : '';
    document.getElementById('formMediaSource').value = isEdit ? (data.mediaSource||'') : '';
    document.getElementById('formWebSource').value = isEdit ? (data.webSource||'') : '';
    document.getElementById('formOfficialDocument').value = isEdit ? (data.officialDocument||'') : '';
    
    // Load submitter info if editing a pending event
    if(isEdit && data.submitterName){
      if(submitterSection) submitterSection.style.display = 'block';
      if(submitterNameGroup) submitterNameGroup.style.display = 'block';
      if(submitterIdGroup) submitterIdGroup.style.display = 'block';
      const submitterNameEl = document.getElementById('formSubmitterName');
      const submitterIdEl = document.getElementById('formSubmitterId');
      if(submitterNameEl) submitterNameEl.value = data.submitterName || '';
      if(submitterIdEl) submitterIdEl.value = data.submitterId || '';
    } else if(isSuggest){
      // Clear submitter fields for new suggest
      const submitterNameEl = document.getElementById('formSubmitterName');
      const submitterIdEl = document.getElementById('formSubmitterId');
      if(submitterNameEl) submitterNameEl.value = '';
      if(submitterIdEl) submitterIdEl.value = '';
    }
    
    // Validate form (this will disable submit button if form is invalid)
    validateEventForm();
    
    openDialog(els.eventFormModal);
  }
  // Validate event form
  function validateEventForm(){
    const dateInput = document.getElementById('formDate');
    const titleInput = document.getElementById('formTitleInput');
    const dateError = document.getElementById('formDateError');
    const titleError = document.getElementById('formTitleError');
    const submitBtn = els.eventForm ? els.eventForm.querySelector('button[type="submit"]') : null;
    
    let isValid = true;
    let dateValid = true;
    let titleValid = true;
    
    // Validate date
    const dateValue = dateInput ? dateInput.value.trim() : '';
    if(!dateValue){
      if(dateError){
        dateError.textContent = 'תאריך הוא שדה חובה';
        dateError.style.display = 'block';
      }
      dateValid = false;
      isValid = false;
    } else {
      // Check if date is valid - try parsing it first
      let normalizedDate = '';
      try {
        // Try to normalize the date (supports multiple formats)
        normalizedDate = AppUtils.normalizeDate(dateValue);
        // If normalizeDate returns empty, the date is invalid
        if(!normalizedDate || normalizedDate === ''){
          if(dateError){
            dateError.textContent = 'תאריך לא תקין';
            dateError.style.display = 'block';
          }
          dateValid = false;
          isValid = false;
        } else {
          // Validate the normalized date
          if(!AppUtils.isValidDate(normalizedDate)){
            if(dateError){
              dateError.textContent = 'תאריך לא תקין';
              dateError.style.display = 'block';
            }
            dateValid = false;
            isValid = false;
          } else {
            if(dateError){
              dateError.style.display = 'none';
              dateError.textContent = '';
            }
          }
        }
      } catch(err) {
        if(dateError){
          dateError.textContent = 'תאריך לא תקין';
          dateError.style.display = 'block';
        }
        dateValid = false;
        isValid = false;
      }
    }
    
    // Validate submitter info if in suggest mode
    if(currentFormMode === 'suggest'){
      const submitterName = document.getElementById('formSubmitterName');
      const submitterId = document.getElementById('formSubmitterId');
      if(submitterName && (!submitterName.value || !submitterName.value.trim())){
        isValid = false;
      }
      if(submitterId && (!submitterId.value || !submitterId.value.trim())){
        isValid = false;
      }
    }
    
    // Validate title
    const titleValue = titleInput ? titleInput.value.trim() : '';
    if(!titleValue){
      if(titleError){
        titleError.textContent = 'כותרת היא שדה חובה';
        titleError.style.display = 'block';
      }
      titleValid = false;
      isValid = false;
    } else {
      if(titleError){
        titleError.style.display = 'none';
        titleError.textContent = '';
      }
    }
    
    // Update input styles
    if(dateInput){
      if(dateValid) {
        dateInput.style.borderColor = '';
      } else {
        dateInput.style.borderColor = 'var(--danger)';
      }
    }
    if(titleInput){
      if(titleValid) {
        titleInput.style.borderColor = '';
      } else {
        titleInput.style.borderColor = 'var(--danger)';
      }
    }
    
    // Enable/disable submit button
    if(submitBtn){
      submitBtn.disabled = !isValid;
      if(isValid){
        submitBtn.style.opacity = '1';
        submitBtn.style.cursor = 'pointer';
      } else {
        submitBtn.style.opacity = '0.6';
        submitBtn.style.cursor = 'not-allowed';
      }
    }
    
    return isValid;
  }
  function confirmDelete(id){
    if(!confirm('האם למחוק את האירוע?')) return;
    if(!confirm('לא ניתן לשחזר. בטוח/ה?')) return;
    eventsStore.remove(id);
    // Download updated events.json automatically when deleting event
    try {
      downloadEventsJsonFile();
      toast('האירוע נמחק. הקובץ events.json הורד - העלה אותו לשרת', 'success');
    } catch(err) {
      console.error('Failed to download events.json:', err);
      toast('האירוע נמחק', 'success');
    }
    renderAdminTable();
    renderYearGrid(Number(els.yearInput.value));
    renderChronoList();
  }

  // Dialog helpers
  function openDialog(dlg){ if(typeof dlg.showModal === 'function') dlg.showModal(); else dlg.setAttribute('open',''); }
  function closeDialog(dlg){ if(typeof dlg.close === 'function') dlg.close(); else dlg.removeAttribute('open'); }
  function escapeHTML(s){ return String(s).replace(/[&<>"']/g, (c)=>({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[c])); }

  // Toasts
  const toastBox = document.getElementById('toastContainer');
  function toast(msg, type){
    const t = document.createElement('div'); t.className = 'toast'+(type?` ${type}`:''); t.textContent = msg;
    toastBox.appendChild(t);
    setTimeout(()=>{ t.remove(); }, 3500);
  }
  window.toast = toast;

  // Tooltip (sticky while hovering a cell/dot)
  const tip = document.createElement('div');
  tip.className = 'hover-tip';
  document.body.appendChild(tip);
  let tipHideTimer = null;
  function showTip(text, x, y){
    if(tipHideTimer){
      clearTimeout(tipHideTimer);
      tipHideTimer = null;
    }
    tip.textContent = text;
    positionTip(x, y);
    tip.classList.add('show');
  }
  function moveTip(x, y){ positionTip(x, y); }
  function hideTip(){
    if(tipHideTimer) clearTimeout(tipHideTimer);
    tipHideTimer = setTimeout(()=>{ tip.classList.remove('show'); }, 50);
  }
  function positionTip(x, y){
    const pad = 12;
    const rect = { w: tip.offsetWidth, h: tip.offsetHeight };
    let left = x + pad;
    let top = y + pad;
    if(left + rect.w > window.innerWidth - 8) left = x - rect.w - pad;
    if(top + rect.h > window.innerHeight - 8) top = y - rect.h - pad;
    tip.style.left = left + 'px';
    tip.style.top = top + 'px';
  }
  function attachTooltip(el, lines){
    if(!lines || !lines.length) return;
    const text = lines.join('\n');
    const onOver = (e)=> showTip(text, e.clientX, e.clientY);
    const onMove = (e)=> moveTip(e.clientX, e.clientY);
    const onOut = (e)=> {
      const toEl = e.relatedTarget;
      if(toEl && el.contains(toEl)) return; // still inside the same cell
      // tolerate small movement outside tiny elements (e.g., dots)
      const rect = el.getBoundingClientRect();
      const x = e.clientX, y = e.clientY;
      const withinPadding = x >= rect.left-8 && x <= rect.right+8 && y >= rect.top-8 && y <= rect.bottom+8;
      if(withinPadding){
        if(tipHideTimer) clearTimeout(tipHideTimer);
        tipHideTimer = setTimeout(()=>{ tip.classList.remove('show'); }, 200);
        return;
      }
      hideTip();
    };
    el.addEventListener('mouseenter', onOver);
    el.addEventListener('mousemove', onMove);
    el.addEventListener('mouseleave', onOut);
  }

  // Password check
  const ADMIN_PASSWORD = 'internet';
  function openPasswordModal(){
    if(els.passwordModal && els.passwordInput){
      els.passwordInput.value = '';
      els.passwordError.style.display = 'none';
      els.passwordError.textContent = '';
      openDialog(els.passwordModal);
      setTimeout(() => els.passwordInput.focus(), 100);
    }
  }
  function checkPassword(){
    const password = els.passwordInput ? els.passwordInput.value.trim() : '';
    if(password === ADMIN_PASSWORD){
      closeDialog(els.passwordModal);
      location.hash = '#/admin';
      return true;
    } else {
      els.passwordError.textContent = 'סיסמה שגויה. נסה שוב.';
      els.passwordError.style.display = 'block';
      els.passwordInput.value = '';
      els.passwordInput.focus();
      return false;
    }
  }

  // Events
  els.backToYear.addEventListener('click', backToYear);
  if(els.prevMonth) els.prevMonth.addEventListener('click', ()=>navigateMonth(-1));
  if(els.nextMonth) els.nextMonth.addEventListener('click', ()=>navigateMonth(1));
  els.viewerBtn.addEventListener('click', ()=>{ location.hash = '#/viewer'; });
  els.adminBtn.addEventListener('click', ()=>{ openPasswordModal(); });
  els.yearInput.addEventListener('change', ()=>{ renderYearGrid(Number(els.yearInput.value)); if(els.monthView.style.display==='block'){ currentViewYear = Number(els.yearInput.value); const allYears = els.allYearsToggle && els.allYearsToggle.checked; els.currentMonthDisplay.textContent = allYears ? monthsHe[currentViewMonth] : `${monthsHe[currentViewMonth]} ${currentViewYear}`; renderMonthView(currentViewYear, currentViewMonth);} renderChronoList(); renderAdminTable(); });
  if(els.allYearsToggle) els.allYearsToggle.addEventListener('change', ()=>{
    updateYearInputState();
    if(!els.allYearsToggle.checked){
      els.yearInput.value = '2025';
      localStorage.setItem('prefYear', '2025');
      renderYearGrid(2025);
      if(els.monthView.style.display==='block'){
        currentViewYear = 2025;
        els.currentMonthDisplay.textContent = `${monthsHe[currentViewMonth]} ${currentViewYear}`;
        renderMonthView(currentViewYear, currentViewMonth);
      }
    } else {
      // If "all years" is enabled, refresh month view if active
      if(els.monthView.style.display==='block'){
        const allYears = els.allYearsToggle && els.allYearsToggle.checked;
        els.currentMonthDisplay.textContent = allYears ? monthsHe[currentViewMonth] : `${monthsHe[currentViewMonth]} ${currentViewYear}`;
        renderMonthView(currentViewYear, currentViewMonth);
      }
      renderYearGrid(Number(els.yearInput.value));
      renderChronoList();
    }
    // Always refresh the year grid to update event markings according to filter
    renderYearGrid(Number(els.yearInput.value));
    renderChronoList();
    renderAdminTable();
  });
  // Function to trigger filter animation
  function triggerFilterAnimation(){
    const headerEl = document.querySelector('header');
    const chronoTitleEl = els.chronoTitle;
    
    if(headerEl){
      headerEl.classList.remove('animate-filter');
      // Force reflow
      void headerEl.offsetWidth;
      headerEl.classList.add('animate-filter');
      
      // Remove class after animation completes
      setTimeout(() => {
        headerEl.classList.remove('animate-filter');
      }, 800);
    }
    
    if(chronoTitleEl){
      chronoTitleEl.classList.remove('animate-filter');
      // Force reflow
      void chronoTitleEl.offsetWidth;
      chronoTitleEl.classList.add('animate-filter');
      
      // Remove class after animation completes
      setTimeout(() => {
        chronoTitleEl.classList.remove('animate-filter');
      }, 1000);
    }
  }

  els.globalSearch.addEventListener('input', ()=>{
    triggerFilterAnimation();
    const isViewer = els.viewerSection.classList.contains('active');
    const isAdmin = els.adminSection.classList.contains('active');
    if(isViewer){
      const isMonthView = els.monthView && els.monthView.style.display === 'block';
      if(isMonthView){
        // If in month view, update it immediately
        renderMonthView(currentViewYear, currentViewMonth);
      } else {
        // If in year view, update year grid and chrono list
        renderYearGrid(Number(els.yearInput.value));
        renderChronoList();
      }
      // Always update yearGrid in background so it's ready when user returns
      if(isMonthView){
        renderYearGrid(Number(els.yearInput.value));
      }
    }
    if(isAdmin){
      renderAdminTable();
    }
  });
  els.categoryFilter.addEventListener('change', ()=>{
    triggerFilterAnimation();
    const isViewer = els.viewerSection.classList.contains('active');
    const isAdmin = els.adminSection.classList.contains('active');
    if(isViewer){
      const isMonthView = els.monthView && els.monthView.style.display === 'block';
      if(isMonthView){
        // If in month view, update it immediately
        renderMonthView(currentViewYear, currentViewMonth);
      } else {
        // If in year view, update year grid and chrono list
        renderYearGrid(Number(els.yearInput.value));
        renderChronoList();
      }
      // Always update yearGrid in background so it's ready when user returns
      if(isMonthView){
        renderYearGrid(Number(els.yearInput.value));
      }
    }
    if(isAdmin){
      renderAdminTable();
    }
  });
  
  
  els.addEventBtn.addEventListener('click', ()=>openEventForm('add', {}));
  if(els.assistantLinkSettingsBtn){
    els.assistantLinkSettingsBtn.addEventListener('click', ()=>{
      fillAssistantLinkAdminForm();
      if(els.assistantLinkModal) openDialog(els.assistantLinkModal);
    });
  }
  if(els.cancelAssistantLinkBtn){
    els.cancelAssistantLinkBtn.addEventListener('click', ()=>{
      if(els.assistantLinkModal) closeDialog(els.assistantLinkModal);
    });
  }
  if(els.saveAssistantLinkBtn){
    els.saveAssistantLinkBtn.addEventListener('click', ()=>{
      const label = (els.assistantLinkLabelInput?.value || '').trim() || ASSISTANT_LINK_DEFAULTS.label;
      let url = (els.assistantLinkUrlInput?.value || '').trim();
      if(url && !/^https?:\/\//i.test(url)){
        url = 'https://' + url;
      }
      saveAssistantLinkSettings({ label, url });
      applyAssistantLinkButton();
      fillAssistantLinkAdminForm();
      if(els.assistantLinkModal) closeDialog(els.assistantLinkModal);
      try {
        downloadEventsJsonFile();
        toast(url
          ? 'הגדרות הכפתור נשמרו. הקובץ events.json הורד - העלה אותו לשרת'
          : 'נשמר השם ב-events.json. הוסיפו קישור כדי להפעיל את הכפתור',
          'success');
      } catch(err) {
        console.error('Failed to download events.json:', err);
        toast(url ? 'הגדרות הכפתור נשמרו' : 'נשמר השם. הוסיפו קישור כדי להפעיל את הכפתור', 'success');
      }
    });
  }
  if(els.themeAccentSettingsBtn){
    els.themeAccentSettingsBtn.addEventListener('click', ()=>{
      themeAccentPreviewBeforeOpen = loadThemeAccent();
      siteTitlePreviewBeforeOpen = loadSiteTitle();
      fillThemeAccentAdminForm();
      applyThemeAccent(themeAccentPreviewBeforeOpen);
      applySiteTitle(siteTitlePreviewBeforeOpen);
      if(els.themeAccentModal) openDialog(els.themeAccentModal);
    });
  }
  if(els.themeAccentInput){
    els.themeAccentInput.addEventListener('input', ()=>{
      setThemeAccentFormColor(els.themeAccentInput.value);
    });
  }
  if(els.themeAccentHexInput){
    els.themeAccentHexInput.addEventListener('input', ()=>{
      const hex = readThemeAccentFromHexField();
      if(hex){
        if(els.themeAccentInput) els.themeAccentInput.value = hex;
        applyThemeAccent(hex);
      }
    });
    els.themeAccentHexInput.addEventListener('blur', ()=>{
      const hex = readThemeAccentFromHexField() || loadThemeAccent();
      setThemeAccentFormColor(hex);
    });
  }
  if(els.siteTitleInput){
    els.siteTitleInput.addEventListener('input', ()=>{
      applySiteTitle(els.siteTitleInput.value);
    });
  }
  if(els.resetThemeAccentBtn){
    els.resetThemeAccentBtn.addEventListener('click', ()=>{
      setThemeAccentFormColor(THEME_ACCENT_DEFAULT);
    });
  }
  if(els.cancelThemeAccentBtn){
    els.cancelThemeAccentBtn.addEventListener('click', ()=>{
      applyThemeAccent(themeAccentPreviewBeforeOpen);
      applySiteTitle(siteTitlePreviewBeforeOpen);
      fillThemeAccentAdminForm();
      if(els.themeAccentModal) closeDialog(els.themeAccentModal);
    });
  }
  if(els.saveThemeAccentBtn){
    els.saveThemeAccentBtn.addEventListener('click', ()=>{
      const fromField = readThemeAccentFromHexField();
      const hex = normalizeHexColor(fromField || els.themeAccentInput?.value, THEME_ACCENT_DEFAULT);
      const title = normalizeSiteTitle(els.siteTitleInput?.value);
      saveThemeAccent(hex);
      saveSiteTitle(title);
      setThemeAccentFormColor(hex);
      applySiteTitle(title);
      themeAccentPreviewBeforeOpen = hex;
      siteTitlePreviewBeforeOpen = title;
      if(els.themeAccentModal) closeDialog(els.themeAccentModal);
      try {
        downloadEventsJsonFile();
        toast('עיצוב האתר נשמר. הקובץ events.json הורד - העלה אותו לשרת', 'success');
      } catch(err) {
        console.error('Failed to download events.json:', err);
        toast('עיצוב האתר נשמר', 'success');
      }
    });
  }
  if(els.themeAccentModal){
    els.themeAccentModal.addEventListener('cancel', ()=>{
      applyThemeAccent(themeAccentPreviewBeforeOpen);
      applySiteTitle(siteTitlePreviewBeforeOpen);
      fillThemeAccentAdminForm();
    });
  }
  if(els.assistantLinkBtn){
    els.assistantLinkBtn.addEventListener('click', (e)=>{
      const settings = loadAssistantLinkSettings();
      if(!settings.url){
        e.preventDefault();
        toast('הקישור עדיין לא הוגדר בממשק הניהול', 'error');
      }
    });
  }
  els.cancelForm.addEventListener('click', ()=>closeDialog(els.eventFormModal));
  els.modalCloseBtns.forEach(btn=>btn.addEventListener('click', (ev)=>{
    const dlg = ev.target.closest('dialog'); if(dlg) closeDialog(dlg);
  }));
  // Close eventsModal when clicking outside (on backdrop)
  if(els.eventsModal){
    els.eventsModal.addEventListener('click', (e)=>{
      // Check if the click is on the dialog backdrop itself (not on modal-content)
      if(e.target === els.eventsModal){
        closeDialog(els.eventsModal);
      }
    });
  }
  document.addEventListener('keydown',(e)=>{
    if(e.key==='Escape'){
      // Close modals first if they are open
      if(els.eventsModal && els.eventsModal.open) {
        closeDialog(els.eventsModal);
        return;
      }
      if(els.eventFormModal && els.eventFormModal.open) {
        closeDialog(els.eventFormModal);
        return;
      }
      if(els.passwordModal && els.passwordModal.open) {
        closeDialog(els.passwordModal);
        return;
      }
      // If month view is active and no modals are open, return to year view
      if(els.monthView && els.monthView.style.display === 'block'){
        backToYear();
      }
    }
  });
  // Password form handlers
  if(els.passwordForm){
    els.passwordForm.addEventListener('submit', (e)=>{
      e.preventDefault();
      checkPassword();
    });
  }
  if(els.cancelPassword){
    els.cancelPassword.addEventListener('click', ()=>{
      closeDialog(els.passwordModal);
    });
  }

  // Form submit
  els.eventForm.addEventListener('submit',(e)=>{
    e.preventDefault();
    // Validate before submit
    if(!validateEventForm()){
      toast('יש לתקן את השגיאות בטופס לפני שליחה','error');
      return;
    }
    // Normalize the date before creating payload
    const rawDate = document.getElementById('formDate').value.trim();
    const normalizedDate = AppUtils.normalizeDate(rawDate);
    if(!normalizedDate || normalizedDate === ''){
      toast('שגיאה: פורמט תאריך לא תקין','error');
      return;
    }
    if(!AppUtils.isValidDate(normalizedDate)){
      toast('שגיאה: פורמט תאריך לא תקין','error');
      return;
    }
    
    const payload = {
      id: document.getElementById('eventId').value || undefined,
      date: normalizedDate, // Use normalized date
      title: document.getElementById('formTitleInput').value.trim(),
      description: document.getElementById('formDescription').value.trim(),
      category: document.getElementById('formCategory').value.trim(),
      tags: document.getElementById('formTags').value.trim(),
      mediaSource: document.getElementById('formMediaSource').value.trim(),
      webSource: document.getElementById('formWebSource').value.trim(),
      officialDocument: document.getElementById('formOfficialDocument').value.trim(),
      status: currentFormMode === 'suggest' ? 'WaitingToApprove' : 'public'
    };
    
    // Add submitter info if in suggest mode
    if(currentFormMode === 'suggest'){
      const submitterName = document.getElementById('formSubmitterName');
      const submitterId = document.getElementById('formSubmitterId');
      if(submitterName) payload.submitterName = submitterName.value.trim();
      if(submitterId) payload.submitterId = submitterId.value.trim();
    } else if(currentFormMode === 'edit'){
      // When editing, preserve submitter info if it exists
      const originalEvent = eventsStore.toArray().find(e => e.id === payload.id);
      if(originalEvent && originalEvent.submitterName){
        payload.submitterName = originalEvent.submitterName;
        payload.submitterId = originalEvent.submitterId;
      }
      // Also allow updating submitter info when editing pending events
      const submitterName = document.getElementById('formSubmitterName');
      const submitterId = document.getElementById('formSubmitterId');
      if(submitterName && submitterName.value.trim()){
        payload.submitterName = submitterName.value.trim();
      }
      if(submitterId && submitterId.value.trim()){
        payload.submitterId = submitterId.value.trim();
      }
    }
    try {
      if(payload.id){
        // When editing, preserve the original status (to keep WaitingToApprove events as WaitingToApprove)
        const originalEvent = eventsStore.toArray().find(e => e.id === payload.id);
        if(originalEvent && originalEvent.status === 'WaitingToApprove'){
          payload.status = 'WaitingToApprove'; // Keep WaitingToApprove status when editing WaitingToApprove events
        }
        eventsStore.update(payload.id, payload);
        // Download updated events.json automatically when admin updates event
        try {
          downloadEventsJsonFile();
          toast('האירוע עודכן. הקובץ events.json הורד - העלה אותו לשרת','success');
        } catch(err) {
          console.error('Failed to download events.json:', err);
          toast('האירוע עודכן','success');
        } 
      } else { 
        if(currentFormMode === 'suggest'){
          payload.status = 'WaitingToApprove';
          eventsStore.add(payload);
          // Download updated events.json automatically for pending events
          try {
            downloadEventsJsonFile();
            toast('האירוע הוצע בהצלחה! הקובץ events.json הורד - העלה אותו לשרת. הוא ממתין לאישור מנהל','success');
          } catch(err) {
            console.error('Failed to download events.json:', err);
            toast('האירוע הוצע בהצלחה! הוא ממתין לאישור מנהל','success');
          }
        } else {
          payload.status = 'public';
          eventsStore.add(payload);
          // Download updated events.json automatically when admin adds event
          try {
            downloadEventsJsonFile();
            toast('האירוע נשמר. הקובץ events.json הורד - העלה אותו לשרת','success');
          } catch(err) {
            console.error('Failed to download events.json:', err);
            toast('האירוע נשמר','success');
          }
        }
      }
      closeDialog(els.eventFormModal);
      renderAdminTable(); 
      renderYearGrid(Number(els.yearInput.value));
      renderChronoList();
    } catch(err){ toast('שגיאה: פורמט תאריך לא תקין','error'); }
  });
  
  // Add real-time validation listeners
  const formDateInput = document.getElementById('formDate');
  const formTitleInput = document.getElementById('formTitleInput');
  const formSubmitterNameInput = document.getElementById('formSubmitterName');
  const formSubmitterIdInput = document.getElementById('formSubmitterId');
  if(formDateInput){
    formDateInput.addEventListener('input', validateEventForm);
    formDateInput.addEventListener('blur', validateEventForm);
  }
  if(formTitleInput){
    formTitleInput.addEventListener('input', validateEventForm);
    formTitleInput.addEventListener('blur', validateEventForm);
  }
  if(formSubmitterNameInput){
    formSubmitterNameInput.addEventListener('input', validateEventForm);
    formSubmitterNameInput.addEventListener('blur', validateEventForm);
  }
  if(formSubmitterIdInput){
    formSubmitterIdInput.addEventListener('input', validateEventForm);
    formSubmitterIdInput.addEventListener('blur', validateEventForm);
  }

  function updateExportButtons(){
    const has = eventsStore.toArray().length > 0;
    if(els.exportExcelBtn) els.exportExcelBtn.disabled = !has;
    if(els.downloadJsonBtn) els.downloadJsonBtn.disabled = !has;
  }

  if(els.excelFileInput){
    els.excelFileInput.addEventListener('change', async (e) => {
      const file = e.target.files && e.target.files[0];
      e.target.value = '';
      if(!file) return;
      try {
        const result = await ExcelIO.importExcel(file);
        if(result.cancelled) return;
        toast(`נטענו ${result.imported} אירועים בהצלחה. לשמירה קבועה - לחצו "הורד events.json"`, 'success');
        refreshAllViews();
      } catch(err) {
        console.error(err);
        toast(err.message || 'קובץ לא תקין', 'error');
      }
    });
  }
  if(els.exportExcelBtn){
    els.exportExcelBtn.addEventListener('click', () => {
      try {
        ExcelIO.exportExcel();
        toast('הקובץ events.xlsx הורד בהצלחה', 'success');
      } catch(err) {
        console.error(err);
        toast(err.message || 'שגיאה בייצוא', 'error');
      }
    });
  }
  if(els.downloadJsonBtn){
    els.downloadJsonBtn.addEventListener('click', downloadEventsJSON);
  }
  if(els.jsonFileInput){
    els.jsonFileInput.addEventListener('change', async (e) => {
      const file = e.target.files && e.target.files[0];
      e.target.value = '';
      if(!file) return;
      try {
        const result = await importEventsJSON(file);
        if(result.cancelled) return;
        toast(`נטענו ${result.imported} אירועים מ-JSON`, 'success');
        refreshAllViews();
      } catch(err) {
        console.error(err);
        toast(err.message || 'קובץ JSON לא תקין', 'error');
      }
    });
  }

  // DEMO DATA
  const DEMO = [
    { id: 'e1', date: '20-07-1969', title: 'נחיתת אפולו 11', description: 'אדם ראשון על הירח', category: 'מדע', tags: 'חלל,נאס"א', mediaSource: '', webSource: 'https://he.wikipedia.org/wiki/אפולו_11', officialDocument: '' },
    { id: 'e2', date: '09-11-1989', title: 'נפילת חומת ברלין', description: 'סמל סיום המלחמה הקרה', category: 'היסטוריה', tags: 'אירופה,גרמניה', mediaSource: '', webSource: '', officialDocument: '' }
  ];

  // Init
  async function init(){
    // load text size pref
    const savedYear = Number(localStorage.getItem('prefYear')) || new Date().getFullYear();
    els.yearInput.value = String(savedYear);
    els.yearInput.addEventListener('change', ()=>localStorage.setItem('prefYear', String(els.yearInput.value)));

    // Try loading from events.json first, then merge with localStorage, then demo
    if(eventsStore.toArray().length===0){
      let loaded = false;
      let eventsFromJson = [];
      let assistantLinkFromJson = null;
      let themeAccentFromJson = null;
      let siteTitleFromJson = null;
      let jsonFetchOk = false;
      try {
        const res = await fetch('events.json?t=' + Date.now(), { cache: 'no-store' });
        jsonFetchOk = res.ok;
        if(res.ok){
          const data = await res.json();
          const parsed = parseEventsFileData(data);
          if(parsed.events.length){
            eventsFromJson = parsed.events;
            assistantLinkFromJson = parsed.assistantLink;
            themeAccentFromJson = parsed.themeAccent;
            siteTitleFromJson = parsed.siteTitle;
            loaded = true;
          } else {
            if(parsed.assistantLink != null) assistantLinkFromJson = parsed.assistantLink;
            if(parsed.themeAccent != null) themeAccentFromJson = parsed.themeAccent;
            if(parsed.siteTitle != null) siteTitleFromJson = parsed.siteTitle;
          }
        }
      } catch(err) {
        console.warn('Failed to load events.json, trying localStorage...', err);
      }

      let eventsFromLocalStorage = [];
      const savedData = eventsStore.loadFromLocalStorage();
      if(savedData && savedData.length > 0){
        eventsFromLocalStorage = savedData;
      }

      const mergedEvents = [];
      if(eventsFromJson.length > 0){
        const jsonEventIds = new Set(eventsFromJson.map(e => e.id));
        mergedEvents.push(...eventsFromJson);
        for(const localEvent of eventsFromLocalStorage){
          if(localEvent.status === 'WaitingToApprove' && !jsonEventIds.has(localEvent.id)){
            mergedEvents.push(localEvent);
          }
        }
      } else if(eventsFromLocalStorage.length > 0){
        mergedEvents.push(...eventsFromLocalStorage);
      }

      // #region agent log
      fetch('http://127.0.0.1:7895/ingest/c3817e5e-4183-4b2d-8d53-00b23585245f',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'9c3b81'},body:JSON.stringify({sessionId:'9c3b81',location:'app.js:init',message:'init data load',data:{protocol:location.protocol,jsonFetchOk,jsonCount:eventsFromJson.length,localCount:eventsFromLocalStorage.length,mergedCount:mergedEvents.length},timestamp:Date.now(),hypothesisId:'A-B'})}).catch(()=>{});
      // #endregion

      if(mergedEvents.length > 0){
        eventsStore.loadFromArray(mergedEvents);
        eventsStore.saveToLocalStorage();
        loaded = true;
      }

      if(!loaded){
        eventsStore.loadFromArray(DEMO);
        eventsStore.saveToLocalStorage();
      }

      applyAssistantLinkFromFile(assistantLinkFromJson);
      applyThemeAccentFromFile(themeAccentFromJson);
      applySiteTitleFromFile(siteTitleFromJson);
    } else {
      // If events already loaded, still try to update from events.json in background
      try {
        const res = await fetch('events.json?t=' + Date.now(), { cache: 'no-store' });
        if(res.ok){
          const data = await res.json();
          const parsed = parseEventsFileData(data);
          applyAssistantLinkFromFile(parsed.assistantLink);
          applyThemeAccentFromFile(parsed.themeAccent);
          applySiteTitleFromFile(parsed.siteTitle);
          const arr = parsed.events;
          if(arr.length){ 
            // Check if the file has more events than what we have
            const currentCount = eventsStore.toArray().length;
            if(arr.length !== currentCount || arr.length > currentCount){
              eventsStore.loadFromArray(arr);
              eventsStore.saveToLocalStorage();
              // Refresh the view
              if(els.viewerSection.classList.contains('active')){
                renderYearGrid(Number(els.yearInput.value));
                if(els.monthView.style.display==='block'){
                  renderMonthView(currentViewYear, currentViewMonth);
                }
                renderChronoList();
              } else if(els.adminSection.classList.contains('active')){
                renderAdminTable();
              }
            }
          }
        }
      } catch(_) { /* ignore */ }
    }
    updateYearInputState();
    
    // Add sort listeners to table headers
    const sortableHeaders = document.querySelectorAll('#eventsTable thead th.sortable');
    sortableHeaders.forEach(th => {
      th.style.cursor = 'pointer';
      th.addEventListener('click', ()=>{
        const column = th.getAttribute('data-sort');
        if(column) handleSortClick(column);
      });
    });
    
    // Initialize sort indicators to show default sorting by date (descending)
    updateSortIndicators();
    
    applyAssistantLinkButton();
    fillAssistantLinkAdminForm();
    applyThemeAccent(loadThemeAccent());
    applySiteTitle(loadSiteTitle());
    fillThemeAccentAdminForm();

    if(!location.hash) location.hash = '#/viewer';
    updateExportButtons();
    navigate(location.hash);
  }

  document.addEventListener('DOMContentLoaded', init);
})();



