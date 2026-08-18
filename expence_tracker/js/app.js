let rawRows=[];
let columns={};
let dailyData={};
let currentFilteredRows=[];
let trendChart=null,donutChart=null,monthlyChart=null,campaignChart=null,nameChart=null;
let detectedDateMode="MDY";

const $=id=>document.getElementById(id);

$("fileInput").addEventListener("change",loadExcel);

document.querySelector(".upload-box").addEventListener("click",function(e){
 if(e.target.id!=="fileInput") $("fileInput").click();
});

async function loadExcel(event){
 const file=event.target.files[0];
 if(!file)return;

 showLoader(true);
 hideError();

 try{
  if(typeof XLSX==="undefined")
   throw new Error("Excel library load nahi hui. Internet connection check karein.");

  const buffer=await file.arrayBuffer();

  const workbook=XLSX.read(buffer,{
   type:"array",
   raw:true,
   cellDates:false
  });

  if(!workbook.SheetNames.length)
   throw new Error("Workbook mein koi sheet nahi mili.");

  const sheet=workbook.Sheets[workbook.SheetNames[0]];

  const matrix=XLSX.utils.sheet_to_json(sheet,{
   header:1,
   raw:true,
   defval:""
  });

  if(matrix.length<2)
   throw new Error("Excel mein header/data rows nahi mili.");

  const headerIndex=findHeaderRow(matrix);

  if(headerIndex===-1)
   throw new Error("Date / Type / Expenses headers detect nahi ho sake.");

  const headers=matrix[headerIndex].map(v=>String(v??"").trim());

  rawRows=matrix
   .slice(headerIndex+1)
   .filter(row=>row.some(v=>v!==""&&v!==null&&v!==undefined))
   .map(row=>{
    const obj={};
    headers.forEach((header,index)=>{
     if(header)obj[header]=row[index]??"";
    });
    return obj;
   });

  if(!rawRows.length)
   throw new Error("Excel mein usable transaction records nahi mile.");

  columns=detectColumns(rawRows);
  validateColumns();

  detectedDateMode=detectDateMode(rawRows);

  const before=rawRows.length;

  rawRows=rawRows.filter(row=>!!parseDate(row[columns.date]));

  if(!rawRows.length)
   throw new Error("Valid date wali transaction rows nahi mili.");

  populateFilters();

  $("empty").style.display="none";
  $("dashboard").style.display="block";

  $("statusText").textContent=
   file.name+" • "+workbook.SheetNames[0]+" • "+
   (before-rawRows.length)+" invalid date rows ignored";

  $("fileDot").classList.add("active");

  updateDashboard();

 }catch(error){
  console.error("Dashboard Error:",error);
  showError(error.message||"Excel process karte waqt error aaya.");
 }finally{
  showLoader(false);
 }
}

function findHeaderRow(matrix){
 const maxRows=Math.min(matrix.length,20);

 for(let i=0;i<maxRows;i++){
  const row=matrix[i].map(normalize);

  const hasDate=row.some(v=>[
   "date","transactiondate","createddate","datetime"
  ].includes(v));

  const hasType=row.some(v=>[
   "type","servicetype","transactiontype","communicationtype"
  ].includes(v));

  const hasExpense=row.some(v=>[
   "expense","expenses","expence","expences","totalexpense",
   "totalexpenses","amount"
  ].includes(v));

  if(hasDate&&(hasType||hasExpense))return i;
 }

 return -1;
}

function normalize(value){
 return String(value??"")
  .toLowerCase()
  .trim()
  .replace(/[\s_\-\/]+/g,"")
  .replace(/[()&]/g,"");
}

function findColumn(headers,names){
 const normalizedHeaders=headers.map(header=>({
  original:header,
  normalized:normalize(header)
 }));

 for(const name of names){
  const wanted=normalize(name);
  const exact=normalizedHeaders.find(x=>x.normalized===wanted);
  if(exact)return exact.original;
 }

 for(const name of names){
  const wanted=normalize(name);
  const partial=normalizedHeaders.find(x=>x.normalized.includes(wanted));
  if(partial)return partial.original;
 }

 return null;
}

function detectColumns(rows){
 const headers=Object.keys(rows[0]||{});

 return {
  date:findColumn(headers,[
   "date","transaction date","created date","date & time","datetime"
  ]),
  name:findColumn(headers,[
   "name","customer name","client name","customer"
  ]),
  campaign:findColumn(headers,[
   "campaign name","campaign","campaignname"
  ]),
  count:findColumn(headers,[
   "count data","countdata","count","data count"
  ]),
  delivered:findColumn(headers,[
   "delivered","delivery","delivered count"
  ]),
  time:findColumn(headers,[
   "time","schedule time"
  ]),
  schedule:findColumn(headers,["schedule"]),
  type:findColumn(headers,[
   "type","service type","transaction type","communication type"
  ]),
  expense:findColumn(headers,[
   "expenses","expense","expences","expence",
   "total expenses","total expense","amount"
  ]),
  pulse:findColumn(headers,["pulse"]),
  charge:findColumn(headers,[
   "charge","rate","per unit charge"
  ])
 };
}

function validateColumns(){
 const missing=[];

 if(!columns.date)missing.push("Date");
 if(!columns.type)missing.push("Type");
 if(!columns.expense)missing.push("Expenses");

 if(missing.length){
  throw new Error(
   "Required columns nahi mili: "+missing.join(", ")+
   ". Expected headers: date, type, expenses."
  );
 }
}

function detectDateMode(rows){
 let dmyEvidence=0,mdyEvidence=0;

 rows.slice(0,5000).forEach(row=>{
  const value=row[columns.date];

  if(typeof value!=="string")return;

  const match=value.trim().match(
   /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/
  );

  if(!match)return;

  const a=Number(match[1]);
  const b=Number(match[2]);

  if(a>12)dmyEvidence++;
  if(b>12)mdyEvidence++;
 });

 return dmyEvidence>mdyEvidence?"DMY":"MDY";
}

function parseDate(value){
 if(value===null||value===undefined||value==="")return null;

 if(value instanceof Date){
  if(!Number.isNaN(value.getTime())){
   return new Date(
    value.getFullYear(),
    value.getMonth(),
    value.getDate()
   );
  }
 }

 if(typeof value==="number"&&Number.isFinite(value)){
  if(XLSX.SSF&&XLSX.SSF.parse_date_code){
   const decoded=XLSX.SSF.parse_date_code(value);

   if(decoded&&decoded.y&&decoded.m&&decoded.d){
    return safeDate(decoded.y,decoded.m,decoded.d);
   }
  }
 }

 const text=String(value).trim();

 let match=text.match(
  /^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/
 );

 if(match){
  return safeDate(
   Number(match[1]),
   Number(match[2]),
   Number(match[3])
  );
 }

 match=text.match(
  /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/
 );

 if(match){
  const a=Number(match[1]);
  const b=Number(match[2]);
  const year=Number(match[3]);

  let day,month;

  if(a>12){
   day=a;month=b;
  }else if(b>12){
   month=a;day=b;
  }else if(detectedDateMode==="DMY"){
   day=a;month=b;
  }else{
   month=a;day=b;
  }

  return safeDate(year,month,day);
 }

 const native=new Date(text);

 if(!Number.isNaN(native.getTime())){
  return new Date(
   native.getFullYear(),
   native.getMonth(),
   native.getDate()
  );
 }

 return null;
}

function safeDate(year,month,day){
 const date=new Date(year,month-1,day);

 if(
  date.getFullYear()!==year||
  date.getMonth()!==month-1||
  date.getDate()!==day
 )return null;

 return date;
}

function dateKey(date){
 return [
  date.getFullYear(),
  String(date.getMonth()+1).padStart(2,"0"),
  String(date.getDate()).padStart(2,"0")
 ].join("-");
}

function monthKey(date){
 return [
  date.getFullYear(),
  String(date.getMonth()+1).padStart(2,"0")
 ].join("-");
}

function formatDate(date){
 if(!date)return "-";

 return String(date.getDate()).padStart(2,"0")+
  "-"+
  String(date.getMonth()+1).padStart(2,"0")+
  "-"+
  date.getFullYear();
}

function keyToDate(key){
 const parts=key.split("-").map(Number);
 return new Date(parts[0],parts[1]-1,parts[2]||1);
}

function getCategory(row){
 const rawType=String(row[columns.type]??"").trim();
 const type=normalize(rawType);

 if(
  type==="obdtrans"||
  type.includes("obdtrans")||
  type==="call"||
  type.includes("call")
 )return "call";

 if(
  type==="whatsapp"||
  type.includes("whatsapp")||
  type==="message"||
  type.includes("message")||
  type==="msg"||
  type.includes("msg")
 )return "msg";

 if(
  type==="voiceandsms"||
  type.includes("voiceandsms")||
  type==="sms"||
  type.includes("sms")
 )return "sms";

 const campaign=columns.campaign?
  normalize(row[columns.campaign]):"";

 if(campaign.includes("obdtrans")||campaign.includes("call"))
  return "call";

 if(campaign.includes("whatsapp")||campaign.includes("message"))
  return "msg";

 if(campaign.includes("voiceandsms")||campaign.includes("sms"))
  return "sms";

 return "other";
}

function numberValue(value){
 if(value===null||value===undefined||value==="")return 0;

 if(typeof value==="number")
  return Number.isFinite(value)?value:0;

 let text=String(value)
  .trim()
  .replace(/₹/g,"")
  .replace(/rs\.?/gi,"")
  .replace(/,/g,"")
  .replace(/\s/g,"");

 if(text.startsWith("(")&&text.endsWith(")"))
  text="-"+text.slice(1,-1);

 const result=Number(text);

 return Number.isFinite(result)?result:0;
}

function money(value){
 return "₹"+Number(value||0).toLocaleString("en-IN",{
  minimumFractionDigits:3,
  maximumFractionDigits:3
 });
}

function populateFilters(){
 const years=new Set(),months=new Set(),names=new Set();

 rawRows.forEach(row=>{
  const date=parseDate(row[columns.date]);

  if(date){
   years.add(String(date.getFullYear()));
   months.add(monthKey(date));
  }

  if(columns.name){
   const name=String(row[columns.name]??"").trim();
   if(name)names.add(name);
  }
 });

 $("yearFilter").innerHTML='<option value="all">All Years</option>';

 [...years].sort((a,b)=>Number(b)-Number(a)).forEach(year=>{
  const option=document.createElement("option");
  option.value=year;
  option.textContent=year;
  $("yearFilter").appendChild(option);
 });

 $("monthFilter").innerHTML='<option value="all">All Months</option>';

 [...months].sort().reverse().forEach(key=>{
  const option=document.createElement("option");
  option.value=key;
  option.textContent=monthLabel(key);
  $("monthFilter").appendChild(option);
 });

 $("nameFilter").innerHTML='<option value="all">All Names</option>';

 [...names].sort((a,b)=>a.localeCompare(b,undefined,{
  sensitivity:"base"
 })).forEach(name=>{
  const option=document.createElement("option");
  option.value=name;
  option.textContent=name;
  $("nameFilter").appendChild(option);
 });
}

function monthLabel(key){
 const parts=key.split("-").map(Number);

 return new Date(parts[0],parts[1]-1,1).toLocaleString(
  "en-IN",
  {month:"long",year:"numeric"}
 );
}

[
 "fromDate","toDate","yearFilter","monthFilter",
 "categoryFilter","nameFilter"
].forEach(id=>{
 $(id).addEventListener("change",updateDashboard);
});

$("search").addEventListener(
 "input",
 debounce(updateDashboard,120)
);

$("resetBtn").addEventListener("click",resetFilters);

function getFilteredRows(){
 let from=$("fromDate").value;
 let to=$("toDate").value;

 if(from&&to&&from>to){
  const temp=from;
  from=to;
  to=temp;
 }

 const year=$("yearFilter").value;
 const month=$("monthFilter").value;
 const category=$("categoryFilter").value;
 const name=$("nameFilter").value;
 const search=normalize($("search").value);

 return rawRows.filter(row=>{
  const date=parseDate(row[columns.date]);
  if(!date)return false;

  const key=dateKey(date);

  if(from&&key<from)return false;
  if(to&&key>to)return false;

  if(
   year!=="all"&&
   String(date.getFullYear())!==year
  )return false;

  if(month!=="all"&&monthKey(date)!==month)
   return false;

  if(category!=="all"&&getCategory(row)!==category)
   return false;

  if(name!=="all"&&columns.name){
   const rowName=String(row[columns.name]??"").trim();
   if(rowName!==name)return false;
  }

  if(search){
   const searchText=[
    columns.name?row[columns.name]:"",
    columns.campaign?row[columns.campaign]:"",
    row[columns.type],
    columns.delivered?row[columns.delivered]:"",
    columns.count?row[columns.count]:""
   ].map(normalize).join(" ");

   if(!searchText.includes(search))return false;
  }

  return true;
 });
}

function updateDashboard(){
 if(!rawRows.length)return;

 const data=getFilteredRows();
 currentFilteredRows=data;

 $("recordCount").textContent=
  data.length.toLocaleString("en-IN")+" records";

 const totals={call:0,msg:0,sms:0,other:0};
 const counts={call:0,msg:0,sms:0,other:0};

 data.forEach(row=>{
  const category=getCategory(row);
  const expense=numberValue(row[columns.expense]);

  totals[category]+=expense;
  counts[category]++;
 });

 const grandTotal=
  totals.call+totals.msg+totals.sms+totals.other;

 animateMoney($("callTotal"),totals.call);
 animateMoney($("msgTotal"),totals.msg);
 animateMoney($("smsTotal"),totals.sms);
 animateMoney($("grandTotal"),grandTotal);

 $("callRows").textContent=
  counts.call.toLocaleString("en-IN")+" transactions";

 $("msgRows").textContent=
  counts.msg.toLocaleString("en-IN")+" transactions";

 $("smsRows").textContent=
  counts.sms.toLocaleString("en-IN")+" transactions";

 $("totalRows").textContent=
  data.length.toLocaleString("en-IN")+" transactions";

 buildDaily(data);
 renderMiniStats(totals,grandTotal);
 renderDailyTable();
 renderCharts(totals);
 renderTransactions(data);
}

function buildDaily(data){
 dailyData={};

 data.forEach(row=>{
  const date=parseDate(row[columns.date]);
  if(!date)return;

  const key=dateKey(date);

  if(!dailyData[key]){
   dailyData[key]={
    call:0,msg:0,sms:0,other:0,total:0
   };
  }

  const category=getCategory(row);
  const amount=numberValue(row[columns.expense]);

  dailyData[key][category]+=amount;
  dailyData[key].total+=amount;
 });
}

function renderMiniStats(totals,grandTotal){
 const keys=Object.keys(dailyData).sort();

 $("activeDays").textContent=
  keys.length.toLocaleString("en-IN");

 if(!keys.length){
  $("avgDaily").textContent=money(0);
  $("highestDay").textContent="-";
  $("highestCategory").textContent="-";
  return;
 }

 let highest=-Infinity;
 let highestDate=null;

 keys.forEach(key=>{
  if(dailyData[key].total>highest){
   highest=dailyData[key].total;
   highestDate=key;
  }
 });

 $("avgDaily").textContent=money(grandTotal/keys.length);

 $("highestDay").textContent=
  formatDate(keyToDate(highestDate))+" • "+money(highest);

 const categories=[
  {label:"Call",value:totals.call},
  {label:"Message",value:totals.msg},
  {label:"SMS",value:totals.sms},
  {label:"Other",value:totals.other}
 ];

 categories.sort((a,b)=>b.value-a.value);

 $("highestCategory").textContent=
  categories[0].value>0?
  categories[0].label+" • "+money(categories[0].value):
  "-";
}

function renderDailyTable(){
 const tbody=$("dailyTable");
 tbody.innerHTML="";

 const keys=Object.keys(dailyData).sort();

 if(!keys.length){
  tbody.innerHTML=`
   <tr>
    <td colspan="5" class="no-data">No matching data found.</td>
   </tr>`;
  return;
 }

 let totalCall=0,totalMsg=0,totalSms=0,total=0;

 const fragment=document.createDocumentFragment();

 keys.forEach(key=>{
  const item=dailyData[key];

  totalCall+=item.call;
  totalMsg+=item.msg;
  totalSms+=item.sms;
  total+=item.total;

  const tr=document.createElement("tr");

  tr.innerHTML=`
   <td><strong>${formatDate(keyToDate(key))}</strong></td>
   <td>${money(item.call)}</td>
   <td>${money(item.msg)}</td>
   <td>${money(item.sms)}</td>
   <td><strong>${money(item.total)}</strong></td>
  `;

  fragment.appendChild(tr);
 });

 tbody.appendChild(fragment);

 const totalRow=document.createElement("tr");
 totalRow.className="total-row";

 totalRow.innerHTML=`
  <td>TOTAL</td>
  <td>${money(totalCall)}</td>
  <td>${money(totalMsg)}</td>
  <td>${money(totalSms)}</td>
  <td>${money(total)}</td>
 `;

 tbody.appendChild(totalRow);
}

function chartFont(){
 return{
  family:'Inter,"Segoe UI",Arial',
  size:10
 };
}

function renderCharts(totals){
 renderTrendChart();
 renderDonutChart(totals);
 renderMonthlyChart();
 renderCampaignChart();
 renderNameChart();
}

function renderTrendChart(){
 const keys=Object.keys(dailyData).sort();

 const labels=keys.map(key=>{
  const date=keyToDate(key);
  return String(date.getDate()).padStart(2,"0")+
   "-"+
   String(date.getMonth()+1).padStart(2,"0");
 });

 if(trendChart)trendChart.destroy();

 trendChart=new Chart($("trendChart"),{
  type:"line",
  data:{
   labels,
   datasets:[
    {
     label:"Call",
     data:keys.map(k=>dailyData[k].call),
     borderColor:"#1478c9",
     backgroundColor:"rgba(20,120,201,.07)",
     fill:true,tension:.35,borderWidth:2,
     pointRadius:2,pointHoverRadius:5
    },
    {
     label:"Message",
     data:keys.map(k=>dailyData[k].msg),
     borderColor:"#159447",
     backgroundColor:"rgba(21,148,71,.06)",
     fill:true,tension:.35,borderWidth:2,
     pointRadius:2,pointHoverRadius:5
    },
    {
     label:"SMS",
     data:keys.map(k=>dailyData[k].sms),
     borderColor:"#e89a16",
     backgroundColor:"rgba(232,154,22,.06)",
     fill:true,tension:.35,borderWidth:2,
     pointRadius:2,pointHoverRadius:5
    }
   ]
  },
  options:{
   responsive:true,
   maintainAspectRatio:false,
   interaction:{mode:"index",intersect:false},
   plugins:{
    legend:{
     position:"bottom",
     labels:{usePointStyle:true,padding:15,font:chartFont()}
    },
    tooltip:{
     callbacks:{
      label:context=>
       " "+context.dataset.label+": "+money(context.raw)
     }
    }
   },
   scales:{
    x:{grid:{display:false},ticks:{font:chartFont()}},
    y:{
     beginAtZero:true,
     ticks:{
      font:chartFont(),
      callback:value=>"₹"+Number(value).toLocaleString("en-IN")
     }
    }
   }
  }
 });
}

function renderDonutChart(totals){
 if(donutChart)donutChart.destroy();

 donutChart=new Chart($("donutChart"),{
  type:"doughnut",
  data:{
   labels:["Call","Message","SMS","Other"],
   datasets:[{
    data:[totals.call,totals.msg,totals.sms,totals.other],
    backgroundColor:["#1478c9","#159447","#e89a16","#94a3b8"],
    borderWidth:0,hoverOffset:8
   }]
  },
  options:{
   responsive:true,
   maintainAspectRatio:false,
   cutout:"68%",
   plugins:{
    legend:{
     position:"bottom",
     labels:{usePointStyle:true,padding:13,font:chartFont()}
    },
    tooltip:{
     callbacks:{
      label:context=>
       " "+context.label+": "+money(context.raw)
     }
    }
   }
  }
 });
}

function renderMonthlyChart(){
 const monthly={};

 currentFilteredRows.forEach(row=>{
  const date=parseDate(row[columns.date]);
  if(!date)return;

  const key=monthKey(date);

  if(!monthly[key])monthly[key]=0;

  monthly[key]+=numberValue(row[columns.expense]);
 });

 const keys=Object.keys(monthly).sort();

 if(monthlyChart)monthlyChart.destroy();

 monthlyChart=new Chart($("monthlyChart"),{
  type:"bar",
  data:{
   labels:keys.map(monthLabel),
   datasets:[{
    label:"Expense",
    data:keys.map(k=>monthly[k]),
    backgroundColor:"#1478c9",
    borderRadius:7,maxBarThickness:40
   }]
  },
  options:{
   responsive:true,
   maintainAspectRatio:false,
   plugins:{
    legend:{display:false},
    tooltip:{
     callbacks:{
      label:context=>" "+money(context.raw)
     }
    }
   },
   scales:{
    x:{grid:{display:false},ticks:{font:chartFont()}},
    y:{
     beginAtZero:true,
     ticks:{
      font:chartFont(),
      callback:value=>"₹"+Number(value).toLocaleString("en-IN")
     }
    }
   }
  }
 });
}

function renderCampaignChart(){
 const campaignTotals={};

 currentFilteredRows.forEach(row=>{
  const campaign=columns.campaign?
   String(row[columns.campaign]??"").trim():"";

  const label=campaign||"Unknown";

  if(!campaignTotals[label])campaignTotals[label]=0;

  campaignTotals[label]+=numberValue(row[columns.expense]);
 });

 const sorted=Object.entries(campaignTotals)
  .sort((a,b)=>b[1]-a[1])
  .slice(0,8);

 if(campaignChart)campaignChart.destroy();

 campaignChart=new Chart($("campaignChart"),{
  type:"bar",
  data:{
   labels:sorted.map(x=>truncate(x[0],20)),
   datasets:[{
    label:"Expense",
    data:sorted.map(x=>x[1]),
    backgroundColor:"#7650d1",
    borderRadius:6,maxBarThickness:30
   }]
  },
  options:{
   indexAxis:"y",
   responsive:true,
   maintainAspectRatio:false,
   plugins:{
    legend:{display:false},
    tooltip:{
     callbacks:{label:context=>" "+money(context.raw)}
    }
   },
   scales:{
    x:{
     beginAtZero:true,
     ticks:{
      font:chartFont(),
      callback:value=>"₹"+Number(value).toLocaleString("en-IN")
     }
    },
    y:{grid:{display:false},ticks:{font:chartFont()}}
   }
  }
 });
}

function renderNameChart(){
 const totals={};

 currentFilteredRows.forEach(row=>{
  const name=columns.name?
   String(row[columns.name]??"").trim():"";

  const label=name||"Unknown";

  if(!totals[label])totals[label]=0;

  totals[label]+=numberValue(row[columns.expense]);
 });

 const sorted=Object.entries(totals)
  .sort((a,b)=>b[1]-a[1])
  .slice(0,8);

 if(nameChart)nameChart.destroy();

 nameChart=new Chart($("nameChart"),{
  type:"bar",
  data:{
   labels:sorted.map(x=>truncate(x[0],18)),
   datasets:[{
    label:"Expense",
    data:sorted.map(x=>x[1]),
    backgroundColor:"#159447",
    borderRadius:6,maxBarThickness:30
   }]
  },
  options:{
   indexAxis:"y",
   responsive:true,
   maintainAspectRatio:false,
   plugins:{
    legend:{display:false},
    tooltip:{
     callbacks:{label:context=>" "+money(context.raw)}
    }
   },
   scales:{
    x:{
     beginAtZero:true,
     ticks:{
      font:chartFont(),
      callback:value=>"₹"+Number(value).toLocaleString("en-IN")
     }
    },
    y:{grid:{display:false},ticks:{font:chartFont()}}
   }
  }
 });
}

/* =========================================================
   TRANSACTION TABLE
   DATE FORMAT FIXED: DD-MM-YYYY
========================================================= */

function renderTransactions(data){

 const tbody=$("transactionTable");
 tbody.innerHTML="";

 const maxRows=1000;

 $("transactionInfo").textContent=
  "Showing "+
  Math.min(data.length,maxRows).toLocaleString("en-IN")+
  " of "+
  data.length.toLocaleString("en-IN")+
  " matching transactions";

 if(!data.length){
  tbody.innerHTML=`
   <tr>
    <td colspan="11" class="no-data">
     No matching transactions.
    </td>
   </tr>`;
  return;
 }

 const visible=data.slice()
  .sort((a,b)=>{
   const da=parseDate(a[columns.date]);
   const db=parseDate(b[columns.date]);
   return db-da;
  })
  .slice(0,maxRows);

 const fragment=document.createDocumentFragment();

 visible.forEach(row=>{

  const date=parseDate(row[columns.date]);

  const category=getCategory(row);

  let badge="badge-other";
  let label="Other";

  if(category==="call"){
   badge="badge-call";
   label="Call";
  }
  else if(category==="msg"){
   badge="badge-msg";
   label="Message";
  }
  else if(category==="sms"){
   badge="badge-sms";
   label="SMS";
  }

  /*
   * IMPORTANT:
   * Transaction Details date is DD-MM-YYYY
   */

  const transactionDate=date?
   String(date.getDate()).padStart(2,"0")+
   "-"+
   String(date.getMonth()+1).padStart(2,"0")+
   "-"+
   date.getFullYear():
   "-";

  const tr=document.createElement("tr");

  tr.innerHTML=`

   <td>
    ${transactionDate}
   </td>

   <td>
    ${safe(
     columns.name?row[columns.name]:""
    )}
   </td>

   <td>
    ${safe(
     columns.campaign?row[columns.campaign]:""
    )}
   </td>

   <td>
    ${safe(row[columns.type])}
   </td>

   <td>
    <span class="badge ${badge}">
     ${label}
    </span>
   </td>

   <td>
    ${safe(
     columns.count?row[columns.count]:""
    )}
   </td>

   <td>
    ${safe(
     columns.delivered?row[columns.delivered]:""
    )}
   </td>

   <td>
    ${safe(
     columns.time?row[columns.time]:""
    )}
   </td>

   <td>
    ${safe(
     columns.pulse?row[columns.pulse]:""
    )}
   </td>

   <td>
    ${safe(
     columns.charge?row[columns.charge]:""
    )}
   </td>

   <td>
    <strong>
     ${money(
      numberValue(row[columns.expense])
     )}
    </strong>
   </td>

  `;

  fragment.appendChild(tr);
 });

 tbody.appendChild(fragment);
}

function resetFilters(){
 $("fromDate").value="";
 $("toDate").value="";
 $("yearFilter").value="all";
 $("monthFilter").value="all";
 $("categoryFilter").value="all";
 $("nameFilter").value="all";
 $("search").value="";
 updateDashboard();
}

$("exportDailyBtn").addEventListener("click",exportDailyCSV);

function exportDailyCSV(){
 const keys=Object.keys(dailyData).sort();

 if(!keys.length){
  alert("Export ke liye data available nahi hai.");
  return;
 }

 const output=[[
  "Date",
  "Call Expense",
  "Message Expense",
  "SMS Expense",
  "Other Expense",
  "Total Expense"
 ]];

 keys.forEach(key=>{
  const item=dailyData[key];

  output.push([
   key,
   item.call.toFixed(3),
   item.msg.toFixed(3),
   item.sms.toFixed(3),
   item.other.toFixed(3),
   item.total.toFixed(3)
  ]);
 });

 downloadCSV(output,"communication-expense-summary.csv");
}

$("exportTransactionBtn")
 .addEventListener("click",exportTransactionsCSV);

function exportTransactionsCSV(){

 if(!currentFilteredRows.length){
  alert("Export ke liye matching transactions available nahi hain.");
  return;
 }

 const output=[[
  "Date",
  "Name",
  "Campaign",
  "Type",
  "Category",
  "Count Data",
  "Delivered",
  "Time",
  "Schedule",
  "Pulse",
  "Charge",
  "Expenses"
 ]];

 currentFilteredRows.forEach(row=>{
  output.push([
   dateKey(parseDate(row[columns.date])),
   columns.name?row[columns.name]:"",
   columns.campaign?row[columns.campaign]:"",
   row[columns.type]??"",
   getCategory(row),
   columns.count?row[columns.count]:"",
   columns.delivered?row[columns.delivered]:"",
   columns.time?row[columns.time]:"",
   columns.schedule?row[columns.schedule]:"",
   columns.pulse?row[columns.pulse]:"",
   columns.charge?row[columns.charge]:"",
   numberValue(row[columns.expense]).toFixed(3)
  ]);
 });

 downloadCSV(
  output,
  "filtered-communication-transactions.csv"
 );
}

function downloadCSV(output,filename){
 const csv=output.map(row=>
  row.map(csvEscape).join(",")
 ).join("\n");

 const blob=new Blob(
  ["\uFEFF"+csv],
  {type:"text/csv;charset=utf-8"}
 );

 const url=URL.createObjectURL(blob);

 const link=document.createElement("a");
 link.href=url;
 link.download=filename;

 document.body.appendChild(link);
 link.click();
 link.remove();

 setTimeout(()=>URL.revokeObjectURL(url),100);
}

function csvEscape(value){
 const text=String(value??"");

 if(/[",\n\r]/.test(text)){
  return '"'+text.replace(/"/g,'""')+'"';
 }

 return text;
}

const previousValues=new WeakMap();

function animateMoney(element,target){

 const previous=previousValues.get(element)??0;
 const start=performance.now();
 const duration=450;

 function animate(now){

  const progress=Math.min(
   (now-start)/duration,
   1
  );

  const eased=1-Math.pow(1-progress,3);

  const current=
   previous+(target-previous)*eased;

  element.textContent=money(current);

  if(progress<1){
   requestAnimationFrame(animate);
  }
  else{
   previousValues.set(element,target);
  }
 }

 requestAnimationFrame(animate);
}

function safe(value){
 if(value===null||value===undefined||value==="")
  return "-";

 return escapeHTML(String(value));
}

function escapeHTML(value){
 return String(value)
  .replace(/&/g,"&amp;")
  .replace(/</g,"&lt;")
  .replace(/>/g,"&gt;")
  .replace(/"/g,"&quot;")
  .replace(/'/g,"&#039;");
}

function truncate(value,length){
 const text=String(value??"");

 if(text.length<=length)return text;

 return text.slice(0,length-1)+"…";
}

function debounce(fn,delay){
 let timer;

 return function(){
  const context=this;
  const args=arguments;

  clearTimeout(timer);

  timer=setTimeout(
   ()=>fn.apply(context,args),
   delay
  );
 };
}

function showLoader(show){
 $("loader").style.display=show?"grid":"none";
}

function showError(message){
 $("error").style.display="block";

 $("error").innerHTML=
  "<strong>Error:</strong> "+
  escapeHTML(message);
}

function hideError(){
 $("error").style.display="none";
 $("error").textContent="";
}

console.log("Communication Expense Tracker loaded.");
