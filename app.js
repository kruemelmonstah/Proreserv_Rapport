const STORAGE_KEY = 'proreserv_rapport_v5';

const state = {
  entries: loadEntries(),
  selectedDate: new Date()
};

const el = {
  form: document.getElementById('entryForm'),
  formTitle: document.getElementById('formTitle'),
  editId: document.getElementById('editId'),
  date: document.getElementById('date'),
  timeRange: document.getElementById('timeRange'),
  breakMinutes: document.getElementById('breakMinutes'),
  minutes: document.getElementById('minutes'),
  supplier: document.getElementById('supplier'),
  module: document.getElementById('module'),
  market: document.getElementById('market'),
  kilometers: document.getElementById('kilometers'),
  breakfast: document.getElementById('breakfast'),
  lunch: document.getElementById('lunch'),
  dinner: document.getElementById('dinner'),
  parking: document.getElementById('parking'),
  taxi: document.getElementById('taxi'),
  hotel: document.getElementById('hotel'),
  other: document.getElementById('other'),
  routeText: document.getElementById('routeText'),
  infoText: document.getElementById('infoText'),
  saveBtn: document.getElementById('saveBtn'),
  resetBtn: document.getElementById('resetBtn'),
  prevWeekBtn: document.getElementById('prevWeekBtn'),
  currentWeekBtn: document.getElementById('currentWeekBtn'),
  nextWeekBtn: document.getElementById('nextWeekBtn'),
  exportBtn: document.getElementById('exportBtn'),
  backupBtn: document.getElementById('backupBtn'),
  restoreInput: document.getElementById('restoreInput'),
  weekLabel: document.getElementById('weekLabel'),
  sumEntries: document.getElementById('sumEntries'),
  sumMinutes: document.getElementById('sumMinutes'),
  sumHours: document.getElementById('sumHours'),
  sumKm: document.getElementById('sumKm'),
  sumMoney: document.getElementById('sumMoney'),
  weekContainer: document.getElementById('weekContainer'),
  supplierList: document.getElementById('supplierList'),
  marketList: document.getElementById('marketList')
};

init();

function init() {
  resetForm();
  bindEvents();
  render();
}

function bindEvents() {
  el.form.addEventListener('submit', onSave);
  el.resetBtn.addEventListener('click', resetForm);
  el.prevWeekBtn.addEventListener('click', () => { state.selectedDate = addDays(state.selectedDate, -7); render(); });
  el.currentWeekBtn.addEventListener('click', () => { state.selectedDate = new Date(); render(); });
  el.nextWeekBtn.addEventListener('click', () => { state.selectedDate = addDays(state.selectedDate, 7); render(); });
  el.exportBtn.addEventListener('click', exportCurrentWeekExcel);
  el.backupBtn.addEventListener('click', exportBackup);
  el.restoreInput.addEventListener('change', importBackup);
  el.timeRange.addEventListener('input', updateCalculatedMinutes);
  el.breakMinutes.addEventListener('input', updateCalculatedMinutes);
}

function updateCalculatedMinutes() {
  const minutes = calculateMinutesFromTimeRange(el.timeRange.value, el.breakMinutes.value);
  el.minutes.value = minutes === null ? '' : String(minutes);
}

function onSave(event) {
  event.preventDefault();

  const normalizedSupplier = normalizeSupplier(el.supplier.value.trim());
  const normalizedMarket = normalizeMarket(el.market.value.trim());
  const date = el.date.value;
  const minutes = calculateMinutesFromTimeRange(el.timeRange.value, el.breakMinutes.value);

  if (!date || !normalizedSupplier || !el.timeRange.value.trim()) {
    alert('Bitte Datum, Zeit und Lieferant ausfüllen.');
    return;
  }

  if (minutes === null) {
    alert('Bitte eine gültige Uhrzeit eingeben, z. B. 08:00-10:55.');
    return;
  }

  if (minutes > 720) {
    const ok = confirm('Warnung: Mehr als 12 Stunden erfasst. Trotzdem speichern?');
    if (!ok) return;
  }

  const entry = {
    id: el.editId.value || makeId(),
    date,
    day: dayCodeFromDate(date),
    timeRange: normalizeTimeRangeDisplay(el.timeRange.value.trim()),
    breakMinutes: numberValue(el.breakMinutes.value),
    minutes,
    kilometers: numberValue(el.kilometers.value),
    supplier: normalizedSupplier,
    module: el.module.value.trim(),
    market: normalizedMarket,
    breakfast: numberValue(el.breakfast.value),
    lunch: numberValue(el.lunch.value),
    dinner: numberValue(el.dinner.value),
    parking: numberValue(el.parking.value),
    taxi: numberValue(el.taxi.value),
    hotel: numberValue(el.hotel.value),
    other: numberValue(el.other.value),
    routeText: el.routeText.value.trim(),
    infoText: el.infoText.value.trim()
  };

  const confirmationText =
    `${entry.day}, ${formatDateCH(entry.date)} – ${entry.timeRange} (${formatMoney(entry.minutes)} Min, Pause ${formatMoney(entry.breakMinutes)}), ` +
    `${formatMoney(entry.kilometers)} km, ${entry.supplier}, ${entry.market || '-'}`;

  const confirmed = confirm(`Eintrag speichern?\n\n${confirmationText}`);
  if (!confirmed) return;

  const index = state.entries.findIndex(item => item.id === entry.id);
  if (index >= 0) {
    state.entries[index] = entry;
  } else {
    state.entries.push(entry);
  }

  saveEntries();
  state.selectedDate = new Date(`${entry.date}T12:00:00`);
  resetForm();
  render();
}

function resetForm() {
  el.formTitle.textContent = 'Eintrag erfassen';
  el.saveBtn.textContent = 'Eintrag speichern';
  el.editId.value = '';
  const today = todayLocalISO();
  el.date.value = today;
  el.timeRange.value = '';
  el.breakMinutes.value = '0';
  el.minutes.value = '';
  el.supplier.value = '';
  el.module.value = '';
  el.market.value = '';
  el.kilometers.value = '';
  el.breakfast.value = '0.00';
  el.lunch.value = '0.00';
  el.dinner.value = '0.00';
  el.parking.value = '0.00';
  el.taxi.value = '0.00';
  el.hotel.value = '0.00';
  el.other.value = '0.00';
  el.routeText.value = '';
  el.infoText.value = '';
}

function render() {
  const week = getIsoWeekInfo(state.selectedDate);
  el.weekLabel.textContent = `KW ${String(week.week).padStart(2, '0')} / ${week.year}`;

  renderSuggestions();

  const entries = getEntriesForWeek(week.year, week.week).sort(sortEntries);
  const grouped = groupByDate(entries);
  const weeklyTotals = calculateTotals(entries);

  el.sumEntries.textContent = String(entries.length);
  el.sumMinutes.textContent = formatMoney(weeklyTotals.minutes);
  el.sumHours.textContent = formatMoney(weeklyTotals.minutes / 60);
  el.sumKm.textContent = formatMoney(weeklyTotals.kilometers);
  el.sumMoney.textContent = formatMoney(weeklyTotals.totalMoney);

  renderWeek(grouped);
}

function renderSuggestions() {
  const suppliers = uniqueValues(state.entries.map(x => x.supplier));
  const markets = uniqueValues(state.entries.map(x => x.market));

  el.supplierList.innerHTML = suppliers.map(v => `<option value="${escapeAttr(v)}"></option>`).join('');
  el.marketList.innerHTML = markets.map(v => `<option value="${escapeAttr(v)}"></option>`).join('');
}

function renderWeek(grouped) {
  el.weekContainer.innerHTML = '';

  if (!grouped.length) {
    el.weekContainer.innerHTML = '<div class="empty">Keine Einträge in dieser Woche.</div>';
    return;
  }

  grouped.forEach(group => {
    const dayTotals = calculateTotals(group.entries);
    const missingKm = group.entries.length > 0 && numberValue(group.entries[0].kilometers) === 0;
    const warningText = missingKm ? ' – Achtung: Kilometer fehlen' : '';

    const card = document.createElement('div');
    card.className = 'day-card';

    card.innerHTML = `
      <div class="day-head">
        <div>
          <div class="day-title">${group.dayCode}, ${formatDateCH(group.date)}${warningText}</div>
          <div class="muted">${group.entries.length} Eintrag / Einträge</div>
        </div>
        <div class="muted">Tageszeit: ${formatMoney(dayTotals.minutes)} Min / ${formatMoney(dayTotals.minutes / 60)} h</div>
      </div>

      <div class="entries">
        <table>
          <thead>
            <tr>
              <th>Tag</th>
              <th>Datum</th>
              <th>Zeit</th>
              <th>Pause</th>
              <th>Minuten</th>
              <th>Kilometer</th>
              <th>Lieferant</th>
              <th>Modul</th>
              <th>Markt</th>
              <th>Spesen (Essen)</th>
              <th>Parking</th>
              <th>Taxen</th>
              <th>Hotel</th>
              <th>Sonstige</th>
              <th>Aktionen</th>
            </tr>
          </thead>
          <tbody>
            ${group.entries.map(entry => `
              <tr>
                <td>${escapeHtml(entry.day)}</td>
                <td>${escapeHtml(formatDateCH(entry.date))}</td>
                <td>${escapeHtml(entry.timeRange)}</td>
                <td style="text-align:center;">${formatMoney(entry.breakMinutes || 0)}</td>
                <td style="text-align:center;">${formatMoney(entry.minutes)}</td>
                <td style="text-align:center;">${formatMoney(entry.kilometers)}</td>
                <td>${escapeHtml(entry.supplier)}</td>
                <td>${escapeHtml(entry.module)}</td>
                <td>${escapeHtml(entry.market)}</td>
                <td>${formatMoney(entry.breakfast + entry.lunch + entry.dinner)}</td>
                <td>${formatMoney(entry.parking)}</td>
                <td>${formatMoney(entry.taxi)}</td>
                <td>${formatMoney(entry.hotel)}</td>
                <td>${formatMoney(entry.other)}</td>
                <td>
                  <div class="small-actions">
                    <button class="btn btn-secondary" data-edit="${entry.id}">Bearbeiten</button>
                    <button class="btn btn-danger" data-delete="${entry.id}">Löschen</button>
                  </div>
                </td>
              </tr>
              ${entry.routeText ? `<tr class="detail-row"><td colspan="15"><strong>Fahrtroute:</strong> ${escapeHtml(entry.routeText)}</td></tr>` : ''}
              ${entry.infoText ? `<tr class="detail-row"><td colspan="15"><strong>Info:</strong> ${escapeHtml(entry.infoText)}</td></tr>` : ''}
            `).join('')}

            <tr class="total-row">
              <td colspan="4">Total</td>
              <td style="text-align:center;">${formatMoney(dayTotals.minutes)}</td>
              <td style="text-align:center;">${formatMoney(dayTotals.kilometers)}</td>
              <td colspan="3"></td>
              <td>${formatMoney(dayTotals.mealMoney)}</td>
              <td>${formatMoney(dayTotals.parking)}</td>
              <td>${formatMoney(dayTotals.taxi)}</td>
              <td>${formatMoney(dayTotals.hotel)}</td>
              <td>${formatMoney(dayTotals.other)}</td>
              <td></td>
            </tr>
          </tbody>
        </table>
      </div>
    `;

    el.weekContainer.appendChild(card);
  });

  el.weekContainer.querySelectorAll('[data-edit]').forEach(btn => btn.addEventListener('click', () => editEntry(btn.dataset.edit)));
  el.weekContainer.querySelectorAll('[data-delete]').forEach(btn => btn.addEventListener('click', () => deleteEntry(btn.dataset.delete)));
}

function editEntry(id) {
  const entry = state.entries.find(x => x.id === id);
  if (!entry) return;

  el.formTitle.textContent = 'Eintrag bearbeiten';
  el.saveBtn.textContent = 'Änderungen speichern';

  el.editId.value = entry.id;
  el.date.value = entry.date;
  el.timeRange.value = entry.timeRange.replace('–', '-');
  el.breakMinutes.value = String(round2(entry.breakMinutes || 0));
  el.minutes.value = formatMoney(entry.minutes);
  el.supplier.value = entry.supplier;
  el.module.value = entry.module;
  el.market.value = entry.market;
  el.kilometers.value = formatMoney(entry.kilometers);
  el.breakfast.value = formatMoney(entry.breakfast);
  el.lunch.value = formatMoney(entry.lunch);
  el.dinner.value = formatMoney(entry.dinner);
  el.parking.value = formatMoney(entry.parking);
  el.taxi.value = formatMoney(entry.taxi);
  el.hotel.value = formatMoney(entry.hotel);
  el.other.value = formatMoney(entry.other);
  el.routeText.value = entry.routeText;
  el.infoText.value = entry.infoText;

  updateCalculatedMinutes();
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function deleteEntry(id) {
  const entry = state.entries.find(x => x.id === id);
  if (!entry) return;

  const ok = confirm(`Eintrag vom ${formatDateCH(entry.date)} ${entry.timeRange} löschen?`);
  if (!ok) return;

  state.entries = state.entries.filter(x => x.id !== id);
  saveEntries();
  render();
}

async function exportCurrentWeekExcel() {
  const week = getIsoWeekInfo(state.selectedDate);
  const entries = getEntriesForWeek(week.year, week.week).sort(sortEntries);

  if (!entries.length) {
    alert('Keine Einträge in dieser Woche.');
    return;
  }

  const grouped = groupByDate(entries);
  const weekTotals = calculateTotals(entries);
  const workbook = new ExcelJS.Workbook();
  const sheetName = `KW${String(week.week).padStart(2, '0')}`;
  const ws = workbook.addWorksheet(sheetName);

  ws.columns = [
    { header: 'Tag', key: 'day', width: 16 },
    { header: 'Datum', key: 'date', width: 12 },
    { header: 'Zeit', key: 'timeRange', width: 16 },
    { header: 'Pause', key: 'breakMinutes', width: 10 },
    { header: 'Minuten', key: 'minutes', width: 10 },
    { header: 'Kilometer', key: 'kilometers', width: 10 },
    { header: 'Lieferant', key: 'supplier', width: 18 },
    { header: 'Modul', key: 'module', width: 14 },
    { header: 'Markt', key: 'market', width: 24 },
    { header: 'Spesen (Essen)', key: 'meal', width: 14 },
    { header: 'Parking', key: 'parking', width: 10 },
    { header: 'Taxen', key: 'taxi', width: 10 },
    { header: 'Hotel', key: 'hotel', width: 10 },
    { header: 'Sonstige', key: 'other', width: 10 }
  ];

  ws.mergeCells('A1:N1');
  ws.getCell('A1').value = `Kalenderwoche ${String(week.week).padStart(2, '0')}`;
  ws.getCell('A1').font = { name: 'Arial', bold: true, size: 14 };
  ws.getCell('A1').alignment = { horizontal: 'left', vertical: 'middle' };

  const headerRow = ws.getRow(3);
  headerRow.values = [
    'Tag', 'Datum', 'Zeit', 'Pause', 'Minuten', 'Kilometer', 'Lieferant', 'Modul',
    'Markt', 'Spesen (Essen)', 'Parking', 'Taxen', 'Hotel', 'Sonstige'
  ];

  headerRow.eachCell((cell) => {
    cell.font = { name: 'Arial', bold: true };
    cell.alignment = { vertical: 'middle', horizontal: 'left' };
    cell.border = thinBorder('FFBFC9D4');
  });

  let rowPointer = 4;

  grouped.forEach((group) => {
    group.entries.forEach((entry, idx) => {
      const row = ws.getRow(rowPointer);
      row.values = [
        entry.day,
        formatDateCH(entry.date),
        entry.timeRange,
        round2(entry.breakMinutes || 0),
        round2(entry.minutes),
        idx === 0 ? round2(entry.kilometers) : '',
        entry.supplier,
        entry.module,
        entry.market,
        round2(entry.breakfast + entry.lunch + entry.dinner),
        round2(entry.parking),
        round2(entry.taxi),
        round2(entry.hotel),
        round2(entry.other)
      ];

      row.eachCell((cell, colNumber) => {
        cell.font = { name: 'Arial' };
        cell.border = thinBorder('FFD8E0EA');
        cell.alignment = {
          vertical: 'middle',
          horizontal: [4, 5, 6].includes(colNumber) ? 'center' : 'left'
        };
        if ([4, 5, 6, 10, 11, 12, 13, 14].includes(colNumber) && cell.value !== '') {
          cell.numFmt = '0.00';
        }
      });

      rowPointer++;
    });

    const totals = calculateTotals(group.entries);
    const totalRow = ws.getRow(rowPointer);
    totalRow.values = [
      'Total', '', '', '',
      round2(totals.minutes),
      round2(totals.kilometers),
      '', '', '',
      round2(totals.mealMoney),
      round2(totals.parking),
      round2(totals.taxi),
      round2(totals.hotel),
      round2(totals.other)
    ];

    styleTotalRow(totalRow);
    rowPointer++;

    const route = group.entries.find(x => x.routeText)?.routeText || '';
    if (route) {
      ws.mergeCells(`A${rowPointer}:N${rowPointer}`);
      const c = ws.getCell(`A${rowPointer}`);
      c.value = `Kilometer ${group.dayCode}: ${route}`;
      c.font = { name: 'Arial' };
      rowPointer++;
    }

const infoLines = group.entries.map(x => x.infoText).filter(Boolean);
if (infoLines.length) {
  ws.mergeCells(`A${rowPointer}:N${rowPointer}`);
  const c = ws.getCell(`A${rowPointer}`);

  c.value = infoLines.join(' | ');
  c.font = {
    name: 'Arial',
    size: 15,
    color: { argb: 'FF8B0000' } // dunkelrot
  };

  c.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFF8CACA' } // hellrot als Feldfarbe
  };

  c.alignment = {
    vertical: 'middle',
    horizontal: 'left',
    wrapText: true
  };

  c.border = thinBorder('FFBFC9D4');

  rowPointer++;
}

    rowPointer++;
  });

  const weeklyRow = ws.getRow(rowPointer);
  weeklyRow.values = [
    'Wochensumme', '', '', '',
    round2(weekTotals.minutes),
    round2(weekTotals.kilometers),
    '', '', '',
    round2(weekTotals.mealMoney),
    round2(weekTotals.parking),
    round2(weekTotals.taxi),
    round2(weekTotals.hotel),
    round2(weekTotals.other)
  ];

  styleTotalRow(weeklyRow);

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  downloadBlob(blob, `Wochenrapport_KW${String(week.week).padStart(2, '0')}.xlsx`);

  function styleTotalRow(row) {
    row.eachCell((cell, colNumber) => {
      cell.font = { name: 'Arial', bold: true };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9EAF7' } };
      cell.border = thinBorder('FFBFC9D4');
      cell.alignment = { vertical: 'middle', horizontal: [5, 6].includes(colNumber) ? 'center' : 'left' };
      if ([5, 6, 10, 11, 12, 13, 14].includes(colNumber) && cell.value !== '') {
        cell.numFmt = '0.00';
      }
    });
  }
}

function thinBorder(argb) {
  return {
    top: { style: 'thin', color: { argb } },
    left: { style: 'thin', color: { argb } },
    bottom: { style: 'thin', color: { argb } },
    right: { style: 'thin', color: { argb } }
  };
}

function exportBackup() {
  const payload = { version: 1, exportedAt: new Date().toISOString(), entries: state.entries };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
  downloadBlob(blob, 'rapport-backup.json');
}

async function importBackup(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  try {
    const text = await file.text();
    const payload = JSON.parse(text);
    if (!payload.entries || !Array.isArray(payload.entries)) throw new Error('Ungültiges Backup');
    state.entries = payload.entries;
    saveEntries();
    render();
    alert('Backup importiert.');
  } catch (error) {
    alert('Backup konnte nicht importiert werden.');
  } finally {
    event.target.value = '';
  }
}

function calculateMinutesFromTimeRange(rawTimeRange, rawBreakMinutes) {
  const text = String(rawTimeRange || '').trim();
  if (!text) return null;

  const normalized = text.replace(/[–—]/g, '-').replace(/\s+/g, '');
  const parts = normalized.split('-');
  if (parts.length !== 2) return null;

  const start = parseTimeToMinutes(parts[0]);
  const end = parseTimeToMinutes(parts[1]);
  if (start === null || end === null) return null;

  const breakMinutes = Math.max(0, numberValue(rawBreakMinutes));
  const diff = end - start - breakMinutes;
  if (diff < 0) return null;
  return diff;
}

function parseTimeToMinutes(value) {
  const match = String(value).match(/^(\d{1,2}):(\d{2})$/);
  if (!match) return null;
  const hours = Number(match[1]);
  const minutes = Number(match[2]);
  if (hours < 0 || hours > 23 || minutes < 0 || minutes > 59) return null;
  return hours * 60 + minutes;
}

function normalizeTimeRangeDisplay(raw) {
  return String(raw || '').trim().replace(/[–—]/g, '-').replace(/\s*-\s*/g, '–');
}

function getEntriesForWeek(year, week) {
  return state.entries.filter(entry => {
    const info = getIsoWeekInfo(new Date(`${entry.date}T12:00:00`));
    return info.year === year && info.week === week;
  });
}

function groupByDate(entries) {
  const map = new Map();
  entries.forEach(entry => {
    if (!map.has(entry.date)) {
      map.set(entry.date, { date: entry.date, dayCode: entry.day, entries: [] });
    }
    map.get(entry.date).entries.push(entry);
  });
  return [...map.values()].sort((a, b) => a.date.localeCompare(b.date));
}

function calculateTotals(entries) {
  return entries.reduce((acc, entry) => {
    acc.minutes += numberValue(entry.minutes);
    acc.kilometers += numberValue(entry.kilometers);
    acc.mealMoney += numberValue(entry.breakfast) + numberValue(entry.lunch) + numberValue(entry.dinner);
    acc.parking += numberValue(entry.parking);
    acc.taxi += numberValue(entry.taxi);
    acc.hotel += numberValue(entry.hotel);
    acc.other += numberValue(entry.other);
    return acc;
  }, {
    minutes: 0,
    kilometers: 0,
    mealMoney: 0,
    parking: 0,
    taxi: 0,
    hotel: 0,
    other: 0,
    get totalMoney() {
      return this.mealMoney + this.parking + this.taxi + this.hotel + this.other;
    }
  });
}

function loadEntries() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch (error) {
    return [];
  }
}

function saveEntries() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.entries));
}

function normalizeSupplier(value) {
  const lower = value.toLowerCase().trim();
  if (!lower) return '';
  if (lower.includes('bahar') || lower === 'bag') return 'Bahag';
  if (lower.includes('sido')) return 'Facido';
  if (lower === 'led' || lower.includes('levent')) return 'Ledvance';
  if (lower.includes('concord')) return 'Conacord';
  return titleCase(value.trim());
}

function normalizeMarket(value) {
  const lower = value.toLowerCase().trim();
  if (!lower) return '';
  if (lower.includes('jambo')) return 'Jumbo';
  if (lower.includes('irma')) return 'Ermatingen';
  if (lower.includes('sturz')) return 'Storz';
  if (lower.includes('wales')) return 'Rail';
  if (lower.includes('embrach') || lower.includes('keyboard')) return 'Kyburz Embrach';
  if (lower.includes('sirach')) return 'Sirnach';
  if (lower.includes('galgen')) return 'Galgenen';
  if (lower.includes('elch') || lower.includes('lq')) return 'Elkuch';
  return titleCase(value.trim());
}

function sortEntries(a, b) {
  return a.date.localeCompare(b.date) || a.timeRange.localeCompare(b.timeRange);
}

function dayCodeFromDate(isoDate) {
  const date = new Date(`${isoDate}T12:00:00`);
  return ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'][date.getDay()];
}

function todayLocalISO() {
  const now = new Date();
  const tz = now.getTimezoneOffset() * 60000;
  return new Date(now.getTime() - tz).toISOString().slice(0, 10);
}

function getIsoWeekInfo(dateInput) {
  const date = new Date(dateInput);
  const utc = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const day = utc.getUTCDay() || 7;
  utc.setUTCDate(utc.getUTCDate() + 4 - day);
  const yearStart = new Date(Date.UTC(utc.getUTCFullYear(), 0, 1));
  const week = Math.ceil((((utc - yearStart) / 86400000) + 1) / 7);
  return { year: utc.getUTCFullYear(), week };
}

function addDays(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

function formatDateCH(isoDate) {
  const [y, m, d] = isoDate.split('-');
  return `${d}.${m}.${y}`;
}

function numberValue(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function round2(value) {
  return Math.round(numberValue(value) * 100) / 100;
}

function formatMoney(value) {
  return round2(value).toFixed(2);
}

function titleCase(text) {
  return text.replace(/\w\S*/g, word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase());
}

function makeId() {
  return crypto.randomUUID ? crypto.randomUUID() : `id_${Date.now()}_${Math.random().toString(36).slice(2)}`;
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function uniqueValues(values) {
  return [...new Set(values.map(v => String(v || '').trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b));
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function escapeAttr(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}
