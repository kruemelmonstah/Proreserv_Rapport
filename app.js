let data = JSON.parse(localStorage.getItem("data") || "[]");

// ---------------- SAVE ----------------
function save(){

let entry = {
date: date.value,
supplier: fixSupplier(supplier.value),
module: module.value,
market: fixMarket(market.value),
time: time.value,
minutes: parseFloat(minutes.value || 0),
km: parseFloat(km.value || 0),
meal: parseFloat(meal.value || 0),
parking: parseFloat(parking.value || 0),
taxi: parseFloat(taxi.value || 0),
hotel: parseFloat(hotel.value || 0),
other: parseFloat(other.value || 0)
};

// Bestätigung
if(!confirm("Eintrag speichern?")) return;

data.push(entry);
localStorage.setItem("data", JSON.stringify(data));

render();
}

// ---------------- WEEK ----------------
function getWeek(d){
let date = new Date(d);
let onejan = new Date(date.getFullYear(),0,1);
return Math.ceil((((date-onejan)/86400000)+onejan.getDay()+1)/7);
}

// ---------------- RENDER ----------------
function render(){

list.innerHTML = "";

let now = new Date();
let w = getWeek(now);

let totalMin = 0;
let totalKm = 0;
let totalSpesen = 0;

data.forEach(e => {

if(getWeek(e.date)==w){

let li = document.createElement("li");

li.innerText =
`${e.date} | ${e.supplier} | ${e.minutes}min | ${e.km}km`;

list.appendChild(li);

totalMin += e.minutes;
totalKm += e.km;
totalSpesen += (e.meal+e.parking+e.taxi+e.hotel+e.other);

}
});

total.innerText =
`Min: ${totalMin} | KM: ${totalKm} | Spesen: ${totalSpesen.toFixed(2)}`;
}

// ---------------- EXCEL ----------------
function exportExcel(){

let now = new Date();
let w = getWeek(now);

let filtered = data.filter(e => getWeek(e.date)==w);

if(filtered.length==0){
alert("Keine Daten");
return;
}

let rows = [
["Kalenderwoche " + w],
[],
["Datum","Zeit","Minuten","KM","Lieferant","Modul","Markt","Essen","Parking","Taxen","Hotel","Sonstige"]
];

let totalMin=0;
let totalKm=0;

filtered.forEach(e=>{

rows.push([
e.date,
e.time,
e.minutes,
e.km,
e.supplier,
e.module,
e.market,
e.meal,
e.parking,
e.taxi,
e.hotel,
e.other
]);

totalMin += e.minutes;
totalKm += e.km;

});

rows.push(["TOTAL","",totalMin,totalKm]);

let wb = XLSX.utils.book_new();
let ws = XLSX.utils.aoa_to_sheet(rows);

ws['!cols'] = [
{wch:12},{wch:15},{wch:10},{wch:10},
{wch:20},{wch:15},{wch:20},
{wch:10},{wch:10},{wch:10},{wch:10},{wch:10}
];

XLSX.utils.book_append_sheet(wb, ws, "KW"+w);

XLSX.writeFile(wb, "Rapport_KW"+w+".xlsx");
}

// ---------------- AUTO FIX ----------------
function fixSupplier(s){

if(!s) return "";

s = s.toLowerCase();

if(s.includes("bahar")||s.includes("bag")) return "Bahag";
if(s.includes("sido")) return "Facido";
if(s.includes("led")||s.includes("levent")) return "Ledvance";
if(s.includes("concord")) return "Conacord";

return s;
}

function fixMarket(m){

if(!m) return "";

m = m.toLowerCase();

if(m.includes("jambo")) return "Jumbo";
if(m.includes("irma")) return "Ermatingen";
if(m.includes("sturz")) return "Storz";
if(m.includes("wales")) return "Rail";
if(m.includes("embrach")||m.includes("keyboard")) return "Kyburz Embrach";
if(m.includes("sirach")) return "Sirnach";
if(m.includes("galgen")) return "Galgenen";
if(m.includes("elch")||m.includes("lq")) return "Elkuch";

return m;
}

render();