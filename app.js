let data = JSON.parse(localStorage.getItem("data")||"[]");

function save(){
let e={
date:date.value,
supplier:supplier.value,
minutes:parseFloat(minutes.value||0),
km:parseFloat(km.value||0),
meal:parseFloat(meal.value||0),
parking:parseFloat(parking.value||0),
taxi:parseFloat(taxi.value||0),
hotel:parseFloat(hotel.value||0)
};

data.push(e);
localStorage.setItem("data",JSON.stringify(data));
render();
}

function getWeek(d){
let date=new Date(d);
let onejan=new Date(date.getFullYear(),0,1);
return Math.ceil((((date-onejan)/86400000)+onejan.getDay()+1)/7);
}

function render(){
list.innerHTML="";
let now=new Date();
let w=getWeek(now);

let totalMin=0;
let totalKm=0;
let totalSpesen=0;

data.forEach(e=>{
if(getWeek(e.date)==w){
let li=document.createElement("li");
li.innerText=`${e.date} | ${e.supplier} | ${e.minutes}min | ${e.km}km`;
list.appendChild(li);

totalMin+=e.minutes;
totalKm+=e.km;
totalSpesen+=e.meal+e.parking+e.taxi+e.hotel;
}
});

total.innerText=`Min: ${totalMin} | KM: ${totalKm} | Spesen: ${totalSpesen.toFixed(2)}`;
}

render();