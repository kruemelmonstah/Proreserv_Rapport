const STORAGE_KEY='proreserv_rapport_v3';
const state={entries:loadEntries(),selectedDate:new Date()};
const el={
form:document.getElementById('entryForm'),
formTitle:document.getElementById('formTitle'),
editId:document.getElementById('editId'),
date:document.getElementById('date'),
timeRange:document.getElementById('timeRange'),
supplier:document.getElementById('supplier'),
module:document.getElementById('module'),
market:document.getElementById('market'),
minutes:document.getElementById('minutes'),
kilometers:document.getElementById('kilometers'),
breakfast:document.getElementById('breakfast'),
lunch:document.getElementById('lunch'),
dinner:document.getElementById('dinner'),
parking:document.getElementById('parking'),
taxi:document.getElementById('taxi'),
hotel:document.getElementById('hotel'),
other:document.getElementById('other'),
routeText:document.getElementById('routeText'),
infoText:document.getElementById('infoText'),
saveBtn:document.getElementById('saveBtn'),
resetBtn:document.getElementById('resetBtn'),
prevWeekBtn:document.getElementById('prevWeekBtn'),
currentWeekBtn:document.getElementById('currentWeekBtn'),
nextWeekBtn:document.getElementById('nextWeekBtn'),
exportBtn:document.getElementById('exportBtn'),
backupBtn:document.getElementById('backupBtn'),
restoreInput:document.getElementById('restoreInput'),
weekLabel:document.getElementById('weekLabel'),
sumEntries:document.getElementById('sumEntries'),
sumMinutes:document.getElementById('sumMinutes'),
sumHours:document.getElementById('sumHours'),
sumKm:document.getElementById('sumKm'),
sumMoney:document.getElementById('sumMoney'),
weekContainer:document.getElementById('weekContainer'),
supplierList:document.getElementById('supplierList'),
marketList:document.getElementById('marketList')
};
init();
function init(){resetForm();bindEvents();render()}
function bindEvents(){
  el.form.addEventListener('submit',onSave);
  el.resetBtn.addEventListener('click',resetForm);
  el.prevWeekBtn.addEventListener('click',()=>{state.selectedDate=addDays(state.selectedDate,-7);render()});
  el.currentWeekBtn.addEventListener('click',()=>{state.selectedDate=new Date();render()});
  el.nextWeekBtn.addEventListener('click',()=>{state.selectedDate=addDays(state.selectedDate,7);render()});
  el.exportBtn.addEventListener('click',exportCurrentWeekExcel);
  el.backupBtn.addEventListener('click',exportBackup);
  el.restoreInput.addEventListener('change',importBackup);
}
function onSave(event){
  event.preventDefault();
  const normalizedSupplier=normalizeSupplier(el.supplier.value.trim());
  const normalizedMarket=normalizeMarket(el.market.value.trim());
  const minutes=numberValue(el.minutes.value);
  const date=el.date.value;
  if(!date||!normalizedSupplier||!el.timeRange.value.trim()){alert('Bitte Datum, Zeit und Lieferant ausfüllen.');return}
  if(minutes>720){const ok=confirm('Warnung: Mehr als 12 Stunden erfasst. Trotzdem speichern?');if(!ok)return}
  const entry={
    id:el.editId.value||makeId(),
    date,
    day:dayCodeFromDate(date),
    timeRange:el.timeRange.value.trim(),
    minutes,
    kilometers:numberValue(el.kilometers.value),
    supplier:normalizedSupplier,
    module:el.module.value.trim(),
    market:normalizedMarket,
    breakfast:numberValue(el.breakfast.value),
    lunch:numberValue(el.lunch.value),
    dinner:numberValue(el.dinner.value),
    parking:numberValue(el.parking.value),
    taxi:numberValue(el.taxi.value),
    hotel:numberValue(el.hotel.value),
    other:numberValue(el.other.value),
    routeText:el.routeText.value.trim(),
    infoText:el.infoText.value.trim()
  };
  const confirmationText=`${entry.day}, ${formatDateCH(entry.date)} – ${entry.timeRange} (${formatMoney(entry.minutes)} Min), ${formatMoney(entry.kilometers)} km, ${entry.supplier}, ${entry.market||'-'}`;
  if(!confirm(`Eintrag speichern?\n\n${confirmationText}`))return;
  const index=state.entries.findIndex(item=>item.id===entry.id);
  if(index>=0){state.entries[index]=entry}else{state.entries.push(entry)}
  saveEntries();
  state.selectedDate=new Date(`${entry.date}T12:00:00`);
  resetForm();
  render();
}
function resetForm(){
  el.formTitle.textContent='Eintrag erfassen';
  el.saveBtn.textContent='Eintrag speichern';
  el.editId.value='';
  const today=todayLocalISO();
  el.date.value=today;el.timeRange.value='';el.supplier.value='';el.module.value='';el.market.value='';
  el.minutes.value='';el.kilometers.value='';el.breakfast.value='0.00';el.lunch.value='0.00';el.dinner.value='0.00';
  el.parking.value='0.00';el.taxi.value='0.00';el.hotel.value='0.00';el.other.value='0.00';
  el.routeText.value='';el.infoText.value='';
}
function render(){
  const week=getIsoWeekInfo(state.selectedDate);
  el.weekLabel.textContent=`KW ${String(week.week).padStart(2,'0')} / ${week.year}`;
  renderSuggestions();
  const entries=getEntriesForWeek(week.year,week.week).sort(sortEntries);
  const grouped=groupByDate(entries);
  const weeklyTotals=calculateTotals(entries);
  el.sumEntries.textContent=String(entries.length);
  el.sumMinutes.textContent=formatMoney(weeklyTotals.minutes);
  el.sumHours.textContent=formatMoney(weeklyTotals.minutes/60);
  el.sumKm.textContent=formatMoney(weeklyTotals.kilometers);
  el.sumMoney.textContent=formatMoney(weeklyTotals.totalMoney);
  renderWeek(grouped);
}
function renderSuggestions(){
  const suppliers=uniqueValues(state.entries.map(x=>x.supplier));
  const markets=uniqueValues(state.entries.map(x=>x.market));
  el.supplierList.innerHTML=suppliers.map(v=>`<option value="${escapeAttr(v)}"></option>`).join('');
  el.marketList.innerHTML=markets.map(v=>`<option value="${escapeAttr(v)}"></option>`).join('');
}
function renderWeek(grouped){
  el.weekContainer.innerHTML='';
  if(!grouped.length){el.weekContainer.innerHTML='<div class="empty">Keine Einträge in dieser Woche.</div>';return}
  grouped.forEach(group=>{
    const dayTotals=calculateTotals(group.entries);
    const missingKm=group.entries.length>0&&numberValue(group.entries[0].kilometers)===0;
    const warningText=missingKm?' – Achtung: Kilometer fehlen':'';
    const card=document.createElement('div');
    card.className='day-card';
    card.innerHTML=`
      <div class="day-head">
        <div>
          <div class="day-title">${group.dayCode}, ${formatDateCH(group.date)}${warningText}</div>
          <div class="muted">${group.entries.length} Eintrag / Einträge</div>
        </div>
        <div class="muted">Tageszeit: ${formatMoney(dayTotals.minutes)} Min / ${formatMoney(dayTotals.minutes/60)} h</div>
      </div>
      <div class="entries">
        <table>
          <thead>
            <tr>
              <th>Tag</th><th>Datum</th><th>Zeit</th><th>Minuten</th><th>Kilometer</th><th>Lieferant</th><th>Modul</th><th>Markt</th><th>Spesen (Essen)</th><th>Parking</th><th>Taxen</th><th>Hotel</th><th>Sonstige</th><th>Aktionen</th>
            </tr>
          </thead>
          <tbody>
            ${group.entries.map(entry=>`
              <tr>
                <td>${escapeHtml(entry.day)}</td>
                <td>${escapeHtml(formatDateCH(entry.date))}</td>
                <td>${escapeHtml(entry.timeRange)}</td>
                <td style="text-align:center;">${formatMoney(entry.minutes)}</td>
                <td style="text-align:center;">${formatMoney(entry.kilometers)}</td>
                <td>${escapeHtml(entry.supplier)}</td>
                <td>${escapeHtml(entry.module)}</td>
                <td>${escapeHtml(entry.market)}</td>
                <td>${formatMoney(entry.breakfast+entry.lunch+entry.dinner)}</td>
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
              ${entry.routeText?`<tr class="detail-row"><td colspan="14"><strong>Fahrtroute:</strong> ${escapeHtml(entry.routeText)}</td></tr>`:''}
              ${entry.infoText?`<tr class="detail-row"><td colspan="14"><strong>Info:</strong> ${escapeHtml(entry.infoText)}</td></tr>`:''}
            `).join('')}
            <tr class="total-row">
              <td colspan="3">Total</td>
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
      </div>`;
    el.weekContainer.appendChild(card);
  });
  el.weekContainer.querySelectorAll('[data-edit]').forEach(btn=>btn.addEventListener('click',()=>editEntry(btn.dataset.edit)));
  el.weekContainer.querySelectorAll('[data-delete]').forEach(btn=>btn.addEventListener('click',()=>deleteEntry(btn.dataset.delete)));
}
function editEntry(id){
  const entry=state.entries.find(x=>x.id===id);
  if(!entry)return;
  el.formTitle.textContent='Eintrag bearbeiten';
  el.saveBtn.textContent='Änderungen speichern';
  el.editId.value=entry.id;el.date.value=entry.date;el.timeRange.value=entry.timeRange;el.supplier.value=entry.supplier;
  el.module.value=entry.module;el.market.value=entry.market;el.minutes.value=formatMoney(entry.minutes);el.kilometers.value=formatMoney(entry.kilometers);
  el.breakfast.value=formatMoney(entry.breakfast);el.lunch.value=formatMoney(entry.lunch);el.dinner.value=formatMoney(entry.dinner);
  el.parking.value=formatMoney(entry.parking);el.taxi.value=formatMoney(entry.taxi);el.hotel.value=formatMoney(entry.hotel);el.other.value=formatMoney(entry.other);
  el.routeText.value=entry.routeText;el.infoText.value=entry.infoText;
  window.scrollTo({top:0,behavior:'smooth'});
}
function deleteEntry(id){
  const entry=state.entries.find(x=>x.id===id);
  if(!entry)return;
  if(!confirm(`Eintrag vom ${formatDateCH(entry.date)} ${entry.timeRange} löschen?`))return;
  state.entries=state.entries.filter(x=>x.id!==id);saveEntries();render();
}
function exportCurrentWeekExcel(){
  const week=getIsoWeekInfo(state.selectedDate);
  const entries=getEntriesForWeek(week.year,week.week).sort(sortEntries);
  if(!entries.length){alert('Keine Einträge in dieser Woche.');return}
  const grouped=groupByDate(entries);
  const rows=[];
  rows.push([`Kalenderwoche ${String(week.week).padStart(2,'0')}`]);
  rows.push([]);
  rows.push(['Tag','Datum','Zeit','Minuten','Kilometer','Lieferant','Modul','Markt','Spesen (Essen)','Parking','Taxen','Hotel','Sonstige']);
  grouped.forEach(group=>{
    group.entries.forEach((entry,idx)=>{
      rows.push([entry.day,formatDateCH(entry.date),entry.timeRange,round2(entry.minutes),idx===0?round2(entry.kilometers):'',entry.supplier,entry.module,entry.market,round2(entry.breakfast+entry.lunch+entry.dinner),round2(entry.parking),round2(entry.taxi),round2(entry.hotel),round2(entry.other)]);
    });
    const totals=calculateTotals(group.entries);
    rows.push(['Total','','',round2(totals.minutes),round2(totals.kilometers),'','','',round2(totals.mealMoney),round2(totals.parking),round2(totals.taxi),round2(totals.hotel),round2(totals.other)]);
    const route=group.entries.find(x=>x.routeText)?.routeText||'';
    if(route)rows.push([`Kilometer ${group.dayCode}: ${route}`]);
    const infoLines=group.entries.map(x=>x.infoText).filter(Boolean);
    if(infoLines.length)rows.push([infoLines.join(' | ')]);
    rows.push([]);
  });
  const weekTotals=calculateTotals(entries);
  rows.push(['Wochensumme','','',round2(weekTotals.minutes),round2(weekTotals.kilometers),'','','',round2(weekTotals.mealMoney),round2(weekTotals.parking),round2(weekTotals.taxi),round2(weekTotals.hotel),round2(weekTotals.other)]);
 const wb = XLSX.utils.book_new();
const ws = XLSX.utils.aoa_to_sheet(rows);

ws['!cols'] = [
  { wch: 6 }, { wch: 12 }, { wch: 16 }, { wch: 10 }, { wch: 10 },
  { wch: 18 }, { wch: 14 }, { wch: 24 }, { wch: 14 }, { wch: 10 },
  { wch: 10 }, { wch: 10 }, { wch: 10 }
];

const range = XLSX.utils.decode_range(ws['!ref']);

for (let R = 0; R <= range.e.r; R++) {
  const firstCellRef = XLSX.utils.encode_cell({ r: R, c: 0 });
  const firstValue = ws[firstCellRef] ? ws[firstCellRef].v : "";

  for (let C = 0; C <= range.e.c; C++) {
    const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
    if (!ws[cellRef]) continue;

    ws[cellRef].s = {
      font: {
        name: "Arial",
        bold: (
          R === 0 ||          // Titel
          R === 2 ||          // Kopfzeile
          firstValue === "Total" ||
          firstValue === "Wochensumme"
        )
      },
      alignment: {
        vertical: "center",
        horizontal: (C === 3 || C === 4) ? "center" : "left"
      }
    };

    if (firstValue === "Total" || firstValue === "Wochensumme") {
      ws[cellRef].s.fill = {
        patternType: "solid",
        fgColor: { rgb: "D9EAF7" }
      };
    }
  }
}

XLSX.utils.book_append_sheet(wb, ws, `KW${String(week.week).padStart(2, '0')}`);
XLSX.writeFile(wb, `Wochenrapport_KW${String(week.week).padStart(2, '0')}.xlsx`);

}
function exportBackup(){
  const payload={version:1,exportedAt:new Date().toISOString(),entries:state.entries};
  const blob=new Blob([JSON.stringify(payload,null,2)],{type:'application/json'});
  downloadBlob(blob,'rapport-backup.json');
}
async function importBackup(event){
  const file=event.target.files?.[0];
  if(!file)return;
  try{
    const text=await file.text();
    const payload=JSON.parse(text);
    if(!payload.entries||!Array.isArray(payload.entries))throw new Error('Ungültig');
    state.entries=payload.entries;saveEntries();render();alert('Backup importiert.');
  }catch(error){alert('Backup konnte nicht importiert werden.')}
  finally{event.target.value=''}
}
function getEntriesForWeek(year,week){return state.entries.filter(entry=>{const info=getIsoWeekInfo(new Date(`${entry.date}T12:00:00`));return info.year===year&&info.week===week})}
function groupByDate(entries){
  const map=new Map();
  entries.forEach(entry=>{if(!map.has(entry.date)){map.set(entry.date,{date:entry.date,dayCode:entry.day,entries:[]})}map.get(entry.date).entries.push(entry)});
  return [...map.values()].sort((a,b)=>a.date.localeCompare(b.date));
}
function calculateTotals(entries){
  return entries.reduce((acc,entry)=>{acc.minutes+=numberValue(entry.minutes);acc.kilometers+=numberValue(entry.kilometers);acc.mealMoney+=numberValue(entry.breakfast)+numberValue(entry.lunch)+numberValue(entry.dinner);acc.parking+=numberValue(entry.parking);acc.taxi+=numberValue(entry.taxi);acc.hotel+=numberValue(entry.hotel);acc.other+=numberValue(entry.other);return acc},{minutes:0,kilometers:0,mealMoney:0,parking:0,taxi:0,hotel:0,other:0,get totalMoney(){return this.mealMoney+this.parking+this.taxi+this.hotel+this.other}});
}
function loadEntries(){try{const raw=localStorage.getItem(STORAGE_KEY);return raw?JSON.parse(raw):[]}catch(error){return[]}}
function saveEntries(){localStorage.setItem(STORAGE_KEY,JSON.stringify(state.entries))}
function normalizeSupplier(value){
  const lower=value.toLowerCase().trim(); if(!lower)return'';
  if(lower.includes('bahar')||lower==='bag')return'Bahag';
  if(lower.includes('sido'))return'Facido';
  if(lower==='led'||lower.includes('levent'))return'Ledvance';
  if(lower.includes('concord'))return'Conacord';
  return titleCase(value.trim());
}
function normalizeMarket(value){
  const lower=value.toLowerCase().trim(); if(!lower)return'';
  if(lower.includes('jambo'))return'Jumbo';
  if(lower.includes('irma'))return'Ermatingen';
  if(lower.includes('sturz'))return'Storz';
  if(lower.includes('wales'))return'Rail';
  if(lower.includes('embrach')||lower.includes('keyboard'))return'Kyburz Embrach';
  if(lower.includes('sirach'))return'Sirnach';
  if(lower.includes('galgen'))return'Galgenen';
  if(lower.includes('elch')||lower.includes('lq'))return'Elkuch';
  return titleCase(value.trim());
}
function sortEntries(a,b){return a.date.localeCompare(b.date)||a.timeRange.localeCompare(b.timeRange)}
function dayCodeFromDate(isoDate){const date=new Date(`${isoDate}T12:00:00`);return['So','Mo','Di','Mi','Do','Fr','Sa'][date.getDay()]}
function todayLocalISO(){const now=new Date();const tz=now.getTimezoneOffset()*60000;return new Date(now.getTime()-tz).toISOString().slice(0,10)}
function getIsoWeekInfo(dateInput){const date=new Date(dateInput);const utc=new Date(Date.UTC(date.getFullYear(),date.getMonth(),date.getDate()));const day=utc.getUTCDay()||7;utc.setUTCDate(utc.getUTCDate()+4-day);const yearStart=new Date(Date.UTC(utc.getUTCFullYear(),0,1));const week=Math.ceil((((utc-yearStart)/86400000)+1)/7);return{year:utc.getUTCFullYear(),week}}
function addDays(date,days){const d=new Date(date);d.setDate(d.getDate()+days);return d}
function formatDateCH(isoDate){const [y,m,d]=isoDate.split('-');return `${d}.${m}.${y}`}
function numberValue(value){const n=Number(value);return Number.isFinite(n)?n:0}
function round2(value){return Math.round(numberValue(value)*100)/100}
function formatMoney(value){return round2(value).toFixed(2)}
function titleCase(text){return text.replace(/\w\S*/g,word=>word.charAt(0).toUpperCase()+word.slice(1).toLowerCase())}
function makeId(){return(crypto.randomUUID?crypto.randomUUID():`id_${Date.now()}_${Math.random().toString(36).slice(2)}`)}
function downloadBlob(blob,filename){const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=filename;a.click();setTimeout(()=>URL.revokeObjectURL(url),1000)}
function uniqueValues(values){return[...new Set(values.map(v=>String(v||'').trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b))}
function escapeHtml(value){return String(value??'').replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;').replaceAll("'",'&#039;')}
function escapeAttr(value){return String(value??'').replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;')}
