/* Табель: выгрузка отчётов в Excel (ExcelJS). Один и тот же код работает в приложении и при сборке образцов. */
(function(root){
"use strict";
const GREEN='FF17614F',SOFT='FFE1EEE8',ZEBRA='FFF6F7F3',GOLD='FFF7EDD3',GREY='FF5B685F',LINE='FFDCE1D8',WHITE='FFFFFFFF';
const L={
  ru:{summary:'Сводка',group:'Участок',num:'№',name:'Имя',hours:'Всего часов',h100:'Часы 100%',h125:'Часы 125%',h150:'Часы 150%',pay:'Зарплата за часы',travel:'Проезд',plus:'Доплаты',minus:'Вычеты',total:'К выплате',
      date:'Дата',dow:'День',in:'Вход',out:'Выход',br:'Перерыв, мин',h:'Часы',rate:'Ставка, ₪/ч',dayPay:'Сумма за день',note:'Примечание',totalMonth:'Итого за период',days:'Рабочих дней',
      grand:'ИТОГО',subtotal:'Итого по участку',company:'Компания',workers:'Работников',period:'Период',payPeriod:'Оплата',weekly:'понедельно',monthly:'помесячно',
      dows:['Вс','Пн','Вт','Ср','Чт','Пт','Сб'],fri:'Пятница',sat:'Суббота',hol:'Праздник',eve:'Канун праздника',hoursH:'Часы',rateH:'Ставка',payH:'Зарплата',totalPay:'ИТОГО ЗП',
      h100w:'100% (будни)',h125w:'125% (сверх нормы, пятница)',h150w:'150% (суббота, праздники)',travelDays:'Проезд ({n} дн. × {r} ₪)',companiesTitle:'Сводка по компаниям',byWeeks:'По неделям',monthTotal:'Итого за месяц',weeksNote:'Недели с воскресенья по субботу; первая и последняя неделя обрезаны по границам месяца.'},
  he:{summary:'סיכום',group:'אזור',num:'מס׳',name:'שם',hours:'סך שעות',h100:'שעות רגילות',h125:'שעות נוספות 125%',h150:'שעות נוספות 150%',pay:'שכר מחושב',travel:'נסיעות',plus:'תוספות',minus:'ניכויים',total:'סך שכר',
      date:'תאריך',dow:'יום',in:'כניסה',out:'יציאה',br:'הפסקה דק׳',h:'שעות',rate:'תעריף',dayPay:'שכר יומי',note:'הערה',totalMonth:'סה״כ לתקופה',days:'ימי עבודה',
      grand:'סה״כ',subtotal:'סה״כ אזור',company:'חברה',workers:'עובדים',period:'תקופה',payPeriod:'תשלום',weekly:'שבועי',monthly:'חודשי',
      dows:['א׳','ב׳','ג׳','ד׳','ה׳','ו׳','ש׳'],fri:'שישי',sat:'שבת',hol:'חג',eve:'ערב חג',hoursH:'שעות',rateH:'תעריף',payH:'שכר',totalPay:'סה״כ שכר',
      h100w:'100%',h125w:'125%',h150w:'150%',travelDays:'נסיעות ({n} ימים × {r} ₪)',companiesTitle:'סיכום לפי חברות',byWeeks:'לפי שבועות',monthTotal:'סה״כ לחודש',weeksNote:'שבוע מיום א׳ עד שבת; השבוע הראשון והאחרון חתוכים לפי גבולות החודש.'}
};
const pad=n=>String(n).padStart(2,'0');
const parseDay=k=>{const [y,m,d]=k.split('-').map(Number);return new Date(y,m-1,d);};
const addDays=(k,n)=>{const d=parseDay(k);d.setDate(d.getDate()+n);return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate());};
const xlDate=k=>{const [y,m,d]=k.split('-').map(Number);return new Date(Date.UTC(y,m-1,d));};
const r2=x=>Math.round(x*10000)/10000;
const fmtD=k=>{const [y,m,d]=k.split('-');return d+'.'+m+'.'+y;};
const fmtDM=k=>{const [,m,d]=k.split('-');return d+'.'+m;};
const font=(o)=>Object.assign({name:'Arial',size:10},o||{});
const border={top:{style:'thin',color:{argb:LINE}},left:{style:'thin',color:{argb:LINE}},bottom:{style:'thin',color:{argb:LINE}},right:{style:'thin',color:{argb:LINE}}};
const fill=c=>({type:'pattern',pattern:'solid',fgColor:{argb:c}});
const MONEY='#,##0.00 ₪;-#,##0.00 ₪;"–"',HRS='0.00;-0.00;"–"',TIME='hh:mm',TSUM='[h]:mm;-[h]:mm;"–"';
function safeName(s,used){s=String(s).replace(/[\[\]\*\?\/\\:]/g,' ').replace(/\s+/g,' ').trim().slice(0,31)||'Лист';let n=s,i=2;while(used.has(n.toLowerCase())){const suf=' ('+i+++')';n=s.slice(0,31-suf.length)+suf;}used.add(n.toLowerCase());return n;}
function head(ws,row,labels,opts){
  labels.forEach((t,i)=>{const c=ws.getCell(row,i+1);c.value=t;c.font=font({bold:true,color:{argb:WHITE},size:(opts&&opts.size)||10});c.fill=fill(GREEN);c.alignment={horizontal:'center',vertical:'middle',wrapText:true};c.border=border;});
  ws.getRow(row).height=(opts&&opts.height)||30;
}
const F=(formula,result)=>({formula,result});
const sum=(col,a,b,res)=>F(`SUM(${col}${a}:${col}${b})`,res);
function special(day,t){if(!day)return '';if(day.hol)return day.hol.type==='eve'?t.eve+(day.hol.name?' · '+day.hol.name:''):t.hol+(day.hol.name?' · '+day.hol.name:'');const w=parseDay(day.date).getDay();return w===5?t.fri:w===6?t.sat:'';}
const rateOf=(w,d)=>d&&d.rate!=null?d.rate:(w.rate||0);
function payOf(w){return w.days.reduce((s,d)=>s+rateOf(w,d)*(d.b100+1.25*d.b125+1.5*d.b150),0);}

/* ---------- Отчёт за месяц: сводка + лист на каждого работника ---------- */
function monthly(ExcelJS,o){
  const t=L[o.lang]||L.ru,rtl=o.lang==='he';const wb=new ExcelJS.Workbook();wb.creator='Табель';
  const used=new Set();const sumName=safeName(t.summary,used);
  const sv=wb.addWorksheet(sumName,{views:[{rightToLeft:rtl,state:'frozen',ySplit:3,showGridLines:false}]});
  const wkSheet=o.byWeeks?wb.addWorksheet(safeName(t.byWeeks,used),{views:[{rightToLeft:rtl,state:'frozen',xSplit:2,ySplit:4,showGridLines:false}]}):null;
  const anyTravel=o.workers.some(w=>w.travelSum),anyPlus=o.workers.some(w=>w.plus),anyMinus=o.workers.some(w=>w.minus);
  const grouped=o.workers.some(w=>w.group);
  // per-worker sheets first (to know totals rows)
  const info=[];
  o.workers.forEach(w=>{
    const name=safeName(w.name+(w.code?' '+w.code:''),used);
    const ws=wb.addWorksheet(name,{views:[{rightToLeft:rtl,state:'frozen',ySplit:3,showGridLines:false}],pageSetup:{orientation:'landscape',fitToPage:true,fitToWidth:1,fitToHeight:0}});
    const cols=[t.date,t.dow,t.in,t.out,t.br,t.h,t.rate,t.h100,t.h125,t.h150,t.dayPay,t.note];
    ws.columns=[12,6,9,9,9,9,10,10,10,10,13,20].map(width=>({width}));
    ws.mergeCells(1,1,1,cols.length);const tc=ws.getCell(1,1);tc.value=w.name+' — '+(w.group||o.company);tc.font=font({bold:true,size:14,color:{argb:WHITE}});tc.fill=fill(GREEN);tc.alignment={horizontal:'center',vertical:'middle'};ws.getRow(1).height=26;
    ws.mergeCells(2,1,2,cols.length);const sc=ws.getCell(2,1);sc.value=o.company+' · '+t.period+': '+fmtD(o.from)+' – '+fmtD(o.to)+(w.code?' · '+t.num+' '+w.code:'');sc.font=font({color:{argb:GREY}});sc.alignment={horizontal:'center'};
    head(ws,3,cols);
    const byDate={};w.days.forEach(d=>{byDate[d.date]=d;});
    let r=4,k=o.from;
    while(k<=o.to){
      const d=byDate[k];const row=ws.getRow(r);
      row.getCell(1).value=xlDate(k);row.getCell(1).numFmt='dd.mm.yyyy';
      row.getCell(2).value=t.dows[parseDay(k).getDay()];
      if(d){
        row.getCell(3).value=d.in/1440;row.getCell(4).value=(d.out%1440)/1440;row.getCell(3).numFmt=row.getCell(4).numFmt=TIME;
        row.getCell(5).value=d.brMin||null;
        row.getCell(6).value=d.H;row.getCell(7).value=rateOf(w,d);
        row.getCell(8).value=d.b100;row.getCell(9).value=d.b125;row.getCell(10).value=d.b150;
        row.getCell(11).value=F(`G${r}*(H${r}+I${r}*1.25+J${r}*1.5)`,r2(rateOf(w,d)*(d.b100+1.25*d.b125+1.5*d.b150)));
      }
      row.getCell(12).value=special(d||{date:k,hol:(o.holidays||{})[k]},t)||null;
      for(let c=1;c<=12;c++){const x=row.getCell(c);x.font=font();x.border=border;if(r%2===1)x.fill=fill(ZEBRA);}
      [6,8,9,10].forEach(c=>row.getCell(c).numFmt=HRS);[7,11].forEach(c=>row.getCell(c).numFmt=MONEY);
      const wd=parseDay(k).getDay();if(wd===5||wd===6||(o.holidays||{})[k])row.getCell(2).fill=row.getCell(12).fill=fill(GOLD);
      r++;k=addDays(k,1);
    }
    const last=r-1,tot=r;
    const tr=ws.getRow(tot);tr.getCell(1).value=t.totalMonth;ws.mergeCells(tot,1,tot,2);
    tr.getCell(3).value=F(`COUNT(C4:C${last})`,w.days.length);tr.getCell(3).numFmt='0';tr.getCell(4).value=t.days;
    const H=w.days.reduce((s,d)=>s+d.H,0),b1=w.days.reduce((s,d)=>s+d.b100,0),b2=w.days.reduce((s,d)=>s+d.b125,0),b3=w.days.reduce((s,d)=>s+d.b150,0),pay=payOf(w);
    tr.getCell(5).value=sum('E',4,last,w.days.reduce((s,d)=>s+(d.brMin||0),0));
    tr.getCell(6).value=sum('F',4,last,r2(H));tr.getCell(8).value=sum('H',4,last,r2(b1));tr.getCell(9).value=sum('I',4,last,r2(b2));tr.getCell(10).value=sum('J',4,last,r2(b3));tr.getCell(11).value=sum('K',4,last,r2(pay));
    for(let c=1;c<=12;c++){const x=tr.getCell(c);x.font=font({bold:true});x.fill=fill(SOFT);x.border={top:{style:'medium',color:{argb:GREEN}}};}
    [6,8,9,10].forEach(c=>tr.getCell(c).numFmt=HRS);tr.getCell(11).numFmt=MONEY;tr.getCell(11).font=font({bold:true,color:{argb:GREEN}});
    // extras
    let er=tot+1;const extra={};
    const addLine=(label,val,key)=>{const row=ws.getRow(er);ws.mergeCells(er,1,er,10);row.getCell(1).value=label;row.getCell(1).alignment={horizontal:rtl?'left':'right'};row.getCell(11).value=val;row.getCell(11).numFmt=MONEY;row.getCell(1).font=row.getCell(11).font=font();extra[key]=er;er++;};
    if(w.travelSum)addLine(t.travelDays.replace('{n}',w.travelDays).replace('{r}',w.travel),r2(w.travelSum),'travel');
    if(w.plus)addLine(t.plus+(w.plusNote?' · '+w.plusNote:''),r2(w.plus),'plus');
    if(w.minus)addLine(t.minus+(w.minusNote?' · '+w.minusNote:''),-r2(w.minus),'minus');
    let totalRow=tot;
    if(er>tot+1){const row=ws.getRow(er);ws.mergeCells(er,1,er,10);row.getCell(1).value=t.total;row.getCell(1).alignment={horizontal:rtl?'left':'right'};
      row.getCell(11).value=F(`SUM(K${tot}:K${er-1})`,r2(pay+(w.travelSum||0)+(w.plus||0)-(w.minus||0)));row.getCell(11).numFmt=MONEY;
      row.getCell(1).font=font({bold:true});row.getCell(11).font=font({bold:true,color:{argb:GREEN}});row.getCell(11).fill=fill(SOFT);totalRow=er;}
    info.push({w,name,tot,totalRow,extra,H,b1,b2,b3,pay});
  });
  // summary
  sv.mergeCells(1,1,1,4);const st=sv.getCell(1,1);st.value=o.company+' · '+fmtD(o.from)+' – '+fmtD(o.to);st.font=font({bold:true,size:14,color:{argb:GREEN}});
  sv.getCell(2,1).value=o.subtitle||'';sv.getCell(2,1).font=font({color:{argb:GREY}});
  const cols=[...(grouped?[t.group]:[]),t.num,t.name,t.hours,t.h100,t.h125,t.h150,t.pay,...(anyTravel?[t.travel]:[]),...(anyPlus?[t.plus]:[]),...(anyMinus?[t.minus]:[]),t.total];
  head(sv,3,cols);
  sv.columns=cols.map((c,i)=>({width:[t.name].includes(c)?24:(c===t.group?14:13)}));
  const off=grouped?1:0;const C=i=>String.fromCharCode(65+i);
  let r=4;const groups=[];const order=[];
  if(grouped){const g={};info.forEach(x=>{const k=x.w.group||'—';(g[k]=g[k]||[]).push(x);});Object.keys(g).forEach(k=>order.push([k,g[k]]));}else order.push([null,info]);
  const subRows=[];
  order.forEach(([gname,list])=>{
    const first=r;
    list.forEach(x=>{
      const row=sv.getRow(r);const q="'"+x.name.replace(/'/g,"''")+"'!";
      if(grouped)row.getCell(1).value=gname;
      row.getCell(1+off).value=x.w.code||'';row.getCell(2+off).value=x.w.name;
      row.getCell(3+off).value=F(q+'F'+x.tot,r2(x.H));row.getCell(4+off).value=F(q+'H'+x.tot,r2(x.b1));row.getCell(5+off).value=F(q+'I'+x.tot,r2(x.b2));row.getCell(6+off).value=F(q+'J'+x.tot,r2(x.b3));
      row.getCell(7+off).value=F(q+'K'+x.tot,r2(x.pay));
      let c=8+off;
      if(anyTravel){row.getCell(c).value=x.extra.travel?F(q+'K'+x.extra.travel,r2(x.w.travelSum)):null;c++;}
      if(anyPlus){row.getCell(c).value=x.extra.plus?F(q+'K'+x.extra.plus,r2(x.w.plus)):null;c++;}
      if(anyMinus){row.getCell(c).value=x.extra.minus?F('-'+q+'K'+x.extra.minus,r2(x.w.minus)):null;c++;}
      row.getCell(c).value=F(q+'K'+x.totalRow,r2(x.pay+(x.w.travelSum||0)+(x.w.plus||0)-(x.w.minus||0)));
      for(let k=1;k<=cols.length;k++){const z=row.getCell(k);z.font=font();z.border=border;if(r%2===1)z.fill=fill(ZEBRA);if(k>=3+off)z.numFmt=k<=6+off?HRS:MONEY;}
      row.getCell(cols.length).font=font({bold:true});
      r++;
    });
    if(grouped){
      const row=sv.getRow(r);row.getCell(1).value=t.subtotal+' · '+gname;sv.mergeCells(r,1,r,2+off);
      for(let k=3+off;k<=cols.length;k++){const col=C(k-1);let res=0;for(let rr=first;rr<r;rr++){const v=sv.getCell(rr,k).value;res+=v&&typeof v==='object'?Number(v.result)||0:Number(v)||0;}row.getCell(k).value=F(`SUM(${col}${first}:${col}${r-1})`,r2(res));row.getCell(k).numFmt=k<=6+off?HRS:MONEY;}
      for(let k=1;k<=cols.length;k++){const z=row.getCell(k);z.font=font({bold:true});z.fill=fill(GOLD);z.border=border;}
      subRows.push(r);r++;
    }
  });
  const gt=sv.getRow(r);gt.getCell(1).value=t.grand;sv.mergeCells(r,1,r,2+off);
  for(let k=3+off;k<=cols.length;k++){const col=C(k-1);const refs=grouped?subRows.map(x=>col+x):[`${col}4:${col}${r-1}`];let res=0;(grouped?subRows:Array.from({length:r-4},(_,i)=>i+4)).forEach(rr=>{const v=sv.getCell(rr,k).value;res+=v&&typeof v==='object'?Number(v.result)||0:Number(v)||0;});gt.getCell(k).value=F(`SUM(${refs.join(',')})`,r2(res));gt.getCell(k).numFmt=k<=6+off?HRS:MONEY;}
  for(let k=1;k<=cols.length;k++){const z=gt.getCell(k);z.font=font({bold:true,size:11,color:{argb:k===cols.length?GREEN:'FF000000'}});z.fill=fill(SOFT);z.border={top:{style:'medium',color:{argb:GREEN}}};}
  sv.pageSetup={orientation:'landscape',fitToPage:true,fitToWidth:1,fitToHeight:0};
  if(wkSheet)weeksSheet(wkSheet,o,t,info);
  return wb;
}
/* лист «По неделям»: для компаний с еженедельной оплатой — месяц, разбитый на недели (Вс–Сб) */
function weeksSheet(ws,o,t,info){
  const weeks=[];let k=o.from;while(k<=o.to){let e=k;while(parseDay(e).getDay()!==6&&e<o.to)e=addDays(e,1);weeks.push([k,e]);k=addDays(e,1);}
  const nC=2+weeks.length*2+2;
  ws.mergeCells(1,1,1,8);ws.getCell(1,1).value=o.company+' · '+t.byWeeks+' · '+fmtD(o.from)+' – '+fmtD(o.to);ws.getCell(1,1).font=font({bold:true,size:14,color:{argb:GREEN}});
  ws.getCell(2,1).value=t.weeksNote;ws.getCell(2,1).font=font({color:{argb:GREY},size:9});
  const set=(r,c,v)=>{const x=ws.getCell(r,c);x.value=v;x.font=font({bold:true,color:{argb:WHITE},size:9});x.fill=fill(GREEN);x.alignment={horizontal:'center',vertical:'middle',wrapText:true};x.border=border;};
  set(3,1,t.num);ws.mergeCells(3,1,4,1);set(3,2,t.name);ws.mergeCells(3,2,4,2);
  weeks.forEach(([a,b],i)=>{const c=3+i*2;set(3,c,fmtDM(a)+' – '+fmtDM(b));ws.mergeCells(3,c,3,c+1);set(4,c,t.h);set(4,c+1,t.payH);});
  const tc=3+weeks.length*2;set(3,tc,t.monthTotal);ws.mergeCells(3,tc,3,tc+1);set(4,tc,t.h);set(4,tc+1,t.payH);
  ws.getRow(3).height=22;ws.getRow(4).height=18;ws.getColumn(1).width=8;ws.getColumn(2).width=24;for(let c=3;c<=nC;c++)ws.getColumn(c).width=c%2?8:12;
  const CL=n=>{let s='';while(n>0){const m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s;};
  const dt=k=>{const [y,m,d]=k.split('-').map(Number);return `DATE(${y},${m},${d})`;};
  let r=5;
  info.forEach(x=>{
    const row=ws.getRow(r);const q="'"+x.name.replace(/'/g,"''")+"'!";const last=x.tot-1;
    row.getCell(1).value=x.w.code||null;row.getCell(2).value=x.w.name;
    let hT=0,pT=0;
    weeks.forEach(([a,b],i)=>{const c=3+i*2;const ds=x.w.days.filter(d=>d.date>=a&&d.date<=b);
      const h=ds.reduce((s,d)=>s+d.H,0),p=ds.reduce((s,d)=>s+rateOf(x.w,d)*(d.b100+1.25*d.b125+1.5*d.b150),0);hT+=h;pT+=p;
      const crit=`${q}$A$4:$A$${last},">="&${dt(a)},${q}$A$4:$A$${last},"<="&${dt(b)}`;
      row.getCell(c).value=F(`SUMIFS(${q}$F$4:$F$${last},${crit})`,r2(h));row.getCell(c+1).value=F(`SUMIFS(${q}$K$4:$K$${last},${crit})`,r2(p));});
    const hs=weeks.map((_,i)=>CL(3+i*2)+r).join('+'),ps=weeks.map((_,i)=>CL(4+i*2)+r).join('+');
    row.getCell(tc).value=F(hs,r2(hT));row.getCell(tc+1).value=F(ps,r2(pT));
    for(let c=1;c<=nC;c++){const z=row.getCell(c);z.font=font({size:9,bold:c>=tc});z.border=border;if(r%2===1)z.fill=fill(ZEBRA);if(c>=3)z.numFmt=c%2?HRS:MONEY;}
    row.getCell(2).font=font();r++;
  });
  const gt=ws.getRow(r);gt.getCell(1).value=t.grand;ws.mergeCells(r,1,r,2);
  for(let c=3;c<=nC;c++){const col=CL(c);let res=0;for(let rr=5;rr<r;rr++){const v=ws.getCell(rr,c).value;res+=v&&typeof v==='object'?Number(v.result)||0:0;}gt.getCell(c).value=F(`SUM(${col}5:${col}${r-1})`,r2(res));gt.getCell(c).numFmt=c%2?HRS:MONEY;}
  for(let c=1;c<=nC;c++){const z=gt.getCell(c);z.font=font({bold:true,color:{argb:c===nC?GREEN:'FF000000'}});z.fill=fill(SOFT);z.border={top:{style:'medium',color:{argb:GREEN}}};}
  ws.pageSetup={orientation:'landscape',fitToPage:true,fitToWidth:1,fitToHeight:0};
}

/* ---------- Отчёт за неделю: сетка, лист на каждую неделю ---------- */
function weekly(ExcelJS,o){
  const t=L[o.lang]||L.ru,rtl=o.lang==='he';const wb=new ExcelJS.Workbook();wb.creator='Табель';const used=new Set();
  o.weeks.forEach(wk=>{
    const ws=wb.addWorksheet(safeName(fmtDM(wk.dates[0])+'–'+fmtDM(wk.dates[6]),used),{views:[{rightToLeft:rtl,state:'frozen',xSplit:2,ySplit:2,showGridLines:false}],pageSetup:{orientation:'landscape',fitToPage:true,fitToWidth:1,fitToHeight:0}});
    const anyTravel=wk.workers.some(w=>w.travelSum),anyPlus=wk.workers.some(w=>w.plus),anyMinus=wk.workers.some(w=>w.minus);
    // header rows
    const set=(r,c,v)=>{const x=ws.getCell(r,c);x.value=v;x.font=font({bold:true,color:{argb:WHITE},size:9});x.fill=fill(GREEN);x.alignment={horizontal:'center',vertical:'middle',wrapText:true};x.border=border;return x;};
    set(1,1,t.name);ws.mergeCells(1,1,2,1);set(1,2,t.num);ws.mergeCells(1,2,2,2);
    wk.dates.forEach((k,i)=>{const c=3+i*3;set(1,c,t.dows[parseDay(k).getDay()]+' '+fmtD(k));ws.mergeCells(1,c,1,c+2);[t.in,t.out,t.h].forEach((l,j)=>{const x=set(2,c+j,l);const wd=parseDay(k).getDay();if(wd===5||wd===6){x.fill=fill('FFB08A2A');}});});
    let c=24;set(1,c,t.hoursH);ws.mergeCells(1,c,1,c+2);[t.h100w,t.h125w,t.h150w].forEach((l,j)=>set(2,c+j,l));
    c=27;set(1,c,t.rateH);ws.mergeCells(1,c,2,c);
    c=28;set(1,c,t.payH);ws.mergeCells(1,c,1,c+2);['100%','125%','150%'].forEach((l,j)=>set(2,c+j,l));
    c=31;const extraCols=[];if(anyTravel)extraCols.push(['travel',t.travel]);if(anyPlus)extraCols.push(['plus',t.plus]);if(anyMinus)extraCols.push(['minus',t.minus]);
    extraCols.forEach(([,l],j)=>{set(1,c+j,l);ws.mergeCells(1,c+j,2,c+j);});
    const totCol=c+extraCols.length;set(1,totCol,t.totalPay);ws.mergeCells(1,totCol,2,totCol);
    ws.getRow(1).height=22;ws.getRow(2).height=34;
    ws.getColumn(1).width=22;ws.getColumn(2).width=11;for(let i=3;i<=23;i++)ws.getColumn(i).width=6.5;for(let i=24;i<=26;i++)ws.getColumn(i).width=9;ws.getColumn(27).width=8;for(let i=28;i<=totCol;i++)ws.getColumn(i).width=11;
    const CL=n=>{let s='';while(n>0){const m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s;};
    let r=3;
    wk.workers.forEach(w=>{
      const row=ws.getRow(r);row.getCell(1).value=w.name;row.getCell(2).value=w.code||null;
      const byDate={};w.days.forEach(d=>{byDate[d.date]=d;});
      wk.dates.forEach((k,i)=>{const d=byDate[k];const cc=3+i*3;if(d){row.getCell(cc).value=d.in/1440;row.getCell(cc+1).value=(d.out%1440)/1440;row.getCell(cc+2).value=d.H/24;}[cc,cc+1].forEach(z=>row.getCell(z).numFmt=TIME);row.getCell(cc+2).numFmt=TSUM;});
      const b1=w.days.reduce((s,d)=>s+d.b100,0),b2=w.days.reduce((s,d)=>s+d.b125,0),b3=w.days.reduce((s,d)=>s+d.b150,0);
      row.getCell(24).value=b1/24;row.getCell(25).value=b2/24;row.getCell(26).value=b3/24;[24,25,26].forEach(z=>row.getCell(z).numFmt=TSUM);
      row.getCell(27).value=w.rate||0;row.getCell(27).numFmt='#,##0.##';
      row.getCell(28).value=F(`X${r}*24*AA${r}`,r2(b1*(w.rate||0)));row.getCell(29).value=F(`Y${r}*24*AA${r}*1.25`,r2(b2*(w.rate||0)*1.25));row.getCell(30).value=F(`Z${r}*24*AA${r}*1.5`,r2(b3*(w.rate||0)*1.5));
      let cc=31;const vals={travel:w.travelSum||0,plus:w.plus||0,minus:-(w.minus||0)};extraCols.forEach(([key])=>{row.getCell(cc).value=vals[key]||null;row.getCell(cc).numFmt=MONEY;cc++;});
      const pay=b1*(w.rate||0)+b2*(w.rate||0)*1.25+b3*(w.rate||0)*1.5+(w.travelSum||0)+(w.plus||0)-(w.minus||0);
      row.getCell(totCol).value=F(`SUM(${CL(28)}${r}:${CL(totCol-1)}${r})`,r2(pay));
      for(let z=1;z<=totCol;z++){const x=row.getCell(z);x.font=font({size:9});x.border=border;if(r%2===0)x.fill=fill(ZEBRA);if(z>=28&&z!==27)x.numFmt=MONEY;}
      row.getCell(1).font=font({size:10});row.getCell(totCol).font=font({bold:true});
      r++;
    });
    const last=r-1;const tr=ws.getRow(r);tr.getCell(1).value=t.grand;
    for(let z=24;z<=totCol;z++){if(z===27)continue;const col=CL(z);let res=0;for(let rr=3;rr<=last;rr++){const v=ws.getCell(rr,z).value;res+=v&&typeof v==='object'?Number(v.result)||0:Number(v)||0;}tr.getCell(z).value=F(`SUM(${col}3:${col}${last})`,z<=26?res:r2(res));tr.getCell(z).numFmt=z<=26?TSUM:MONEY;}
    for(let z=1;z<=totCol;z++){const x=tr.getCell(z);x.font=font({bold:true});x.fill=fill(SOFT);x.border={top:{style:'medium',color:{argb:GREEN}}};}
    tr.getCell(totCol).font=font({bold:true,color:{argb:GREEN}});
  });
  return wb;
}

/* ---------- Сводка по компаниям (для закрытия месяца) ---------- */
function companies(ExcelJS,o){
  const t=L[o.lang]||L.ru;const wb=new ExcelJS.Workbook();wb.creator='Табель';
  const ws=wb.addWorksheet(t.companiesTitle.slice(0,31),{views:[{state:'frozen',ySplit:3,showGridLines:false}]});
  ws.mergeCells(1,1,1,6);ws.getCell(1,1).value=t.companiesTitle+' · '+fmtD(o.from)+' – '+fmtD(o.to);ws.getCell(1,1).font=font({bold:true,size:14,color:{argb:GREEN}});
  const cols=[t.company,t.payPeriod,t.workers,t.hours,t.h125,t.h150,t.pay,t.travel,t.plus,t.minus,t.total];
  head(ws,3,cols);ws.columns=[20,12,11,12,12,12,15,12,12,12,15].map(width=>({width}));
  let r=4;o.rows.forEach(x=>{const row=ws.getRow(r);
    [x.company,x.payPeriod==='week'?t.weekly:t.monthly,x.workers,r2(x.H),r2(x.h125),r2(x.h150),r2(x.pay),r2(x.travel||0),r2(x.plus||0),-r2(x.minus||0)].forEach((v,i)=>{row.getCell(i+1).value=v;});
    row.getCell(11).value=F(`SUM(G${r}:J${r})`,r2(x.pay+(x.travel||0)+(x.plus||0)-(x.minus||0)));
    for(let k=1;k<=11;k++){const z=row.getCell(k);z.font=font({size:11,bold:k===1||k===11});z.border=border;if(r%2===1)z.fill=fill(ZEBRA);if(k>=4)z.numFmt=k<=6?HRS:MONEY;}
    r++;});
  const tr=ws.getRow(r);tr.getCell(1).value=t.grand;
  for(let k=3;k<=11;k++){const col=String.fromCharCode(64+k);let res=0;for(let rr=4;rr<r;rr++){const v=ws.getCell(rr,k).value;res+=v&&typeof v==='object'?Number(v.result)||0:Number(v)||0;}tr.getCell(k).value=F(`SUM(${col}4:${col}${r-1})`,r2(res));if(k>=4)tr.getCell(k).numFmt=k<=6?HRS:MONEY;}
  for(let k=1;k<=11;k++){const z=tr.getCell(k);z.font=font({bold:true,size:11,color:{argb:k===11?GREEN:'FF000000'}});z.fill=fill(SOFT);z.border={top:{style:'medium',color:{argb:GREEN}}};}
  ws.pageSetup={orientation:'landscape',fitToPage:true,fitToWidth:1,fitToHeight:0};
  return wb;
}

/* ---------- расчёт дня по правилам компании (как в приложении) ---------- */
function splitDay(H,dateKey,r,hol){
  const wd=parseDay(dateKey).getDay();let dr=r.day[wd];if(hol)dr=Math.max(dr,hol.type==='eve'?r.eve:r.holiday);
  const normal=Math.min(H,r.norm),o1=Math.min(Math.max(H-r.norm,0),r.ot1Hours),o2=Math.max(H-r.norm-r.ot1Hours,0);
  const b={};[[normal,dr],[o1,Math.max(dr,r.ot1)],[o2,Math.max(dr,r.ot2)]].forEach(([h,p])=>{if(h>0)b[p]=(b[p]||0)+h;});
  const out={b100:0,b125:0,b150:0};Object.keys(b).forEach(p=>{p=Number(p);const h=b[p];if(p<=100)out.b100+=h*p/100;else if(p<=125)out.b125+=h*p/125;else out.b150+=h*p/150;});
  return out;
}
root.TabelReports={monthly,weekly,companies,splitDay,L};
})(typeof window!=='undefined'?window:globalThis);
