function exportExcel() {

let now = new Date();
let w = getWeek(now);

let filtered = data.filter(e => getWeek(e.date)==w);

if(filtered.length==0){
alert("Keine Daten");
return;
}

// Daten vorbereiten
let rows = [
["Kalenderwoche " + w],
[],
["Tag","Datum","Zeit","Minuten","Kilometer","Lieferant","Modul","Markt","Essen","Parking","Taxen","Hotel","Sonstige"]
];

let totalMin=0;
let totalKm=0;
let totalSpesen=0;

filtered.forEach(e=>{
rows.push([
e.day || "",
e.date,
e.time || "",
e.minutes,
e.km,
fixSupplier(e.supplier),
e.module || "",
fixMarket(e.market),
e.meal || 0,
e.parking || 0,
e.taxi || 0,
e.hotel || 0,
e.other || 0
]);

totalMin+=e.minutes||0;
totalKm+=e.km||0;
totalSpesen+=(e.meal||0)+(e.parking||0)+(e.taxi||0)+(e.hotel||0)+(e.other||0);
});

// Total-Zeile
rows.push([
"TOTAL","","",totalMin,totalKm,"","","",
"", "", "", "", ""
]);

// Workbook erstellen
let wb = XLSX.utils.book_new();
let ws = XLSX.utils.aoa_to_sheet(rows);

// Spaltenbreite
ws['!cols'] = [
{wch:5},{wch:12},{wch:15},{wch:10},{wch:10},
{wch:20},{wch:15},{wch:20},
{wch:10},{wch:10},{wch:10},{wch:10},{wch:10}
];

// Sheet hinzufügen
XLSX.utils.book_append_sheet(wb, ws, "KW"+w);

// Datei speichern
XLSX.writeFile(wb, "Rapport_KW"+w+".xlsx");
}

function fixSupplier(s){
if(!s) return "";
s=s.toLowerCase();

if(s.includes("bahar")||s.includes("bag")) return "Bahag";
if(s.includes("sido")) return "Facido";
if(s.includes("led")||s.includes("levent")) return "Ledvance";
if(s.includes("concord")) return "Conacord";

return s;
}

function fixMarket(m){
if(!m) return "";
m=m.toLowerCase();

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