import React, { useEffect, useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { Building2, Check, Download, FileImage, FileSpreadsheet, Plus, Save, ScanText, Settings2, Trash2, UploadCloud, X } from 'lucide-react';
import './styles.css';
import { parseGuestFolioOcr, parseInvoiceOcr } from './ocr-parser.js';
import { validateInvoice } from './validation.js';

const seedCompanies = [
  { id: crypto.randomUUID(), code: 'EXPRO', name: 'EXPRO EGYPT LLC', tax: '204-829-542', invoiceTax: '205938604', po: '4300228298' },
  { id: crypto.randomUUID(), code: 'WEATHERFORD', name: 'WeatherFord', tax: '706-010-205', invoiceTax: '706010205', po: '' }
];
const emptyInvoice = { invoiceNo: '', invoiceDate: '', roomNo: '', guest: '', checkIn: '', checkOut: '', nights: 1, single: 1, double: 0, triple: 0, roomRate: '', accommodation: '', restaurant: '', laundry: '', total: '', amountWords: '' };
const money = (v) => Number(String(v || 0).replace(/,/g, '')) || 0;
const money2 = (v) => Math.trunc((money(v)+Number.EPSILON)*100)/100;
const dateForFile = () => new Intl.DateTimeFormat('en-GB', { day:'2-digit', month:'2-digit', year:'numeric' }).format(new Date()).replaceAll('/', '-');
const digitDistance=(a,b)=>{a=String(a||'').replace(/\D/g,'');b=String(b||'').replace(/\D/g,'');if(a.length!==b.length)return 99;return [...a].reduce((n,ch,i)=>n+(ch!==b[i]),0);};
const addDays=(value,days)=>{const m=String(value).match(/(\d{1,2})-(\d{1,2})-(\d{4})/);if(!m)return '';const d=new Date(Date.UTC(+m[3],+m[2]-1,+m[1]));d.setUTCDate(d.getUTCDate()+Number(days||0));return `${String(d.getUTCDate()).padStart(2,'0')}-${String(d.getUTCMonth()+1).padStart(2,'0')}-${d.getUTCFullYear()}`;};

async function preprocessForOcr(file){
  const bitmap=await createImageBitmap(file);const target=Math.max(1800,Math.min(3000,bitmap.width*2));const scale=target/bitmap.width;
  const canvas=document.createElement('canvas');canvas.width=Math.round(bitmap.width*scale);canvas.height=Math.round(bitmap.height*scale);
  const ctx=canvas.getContext('2d',{willReadFrequently:true});ctx.drawImage(bitmap,0,0,canvas.width,canvas.height);
  const img=ctx.getImageData(0,0,canvas.width,canvas.height);for(let i=0;i<img.data.length;i+=4){const g=.299*img.data[i]+.587*img.data[i+1]+.114*img.data[i+2];const v=g>205?255:g<95?0:Math.max(0,Math.min(255,(g-128)*1.65+128));img.data[i]=img.data[i+1]=img.data[i+2]=v;}ctx.putImageData(img,0,0);return canvas;
}
function cropCanvas(source,left,top,right,bottom){const x=Math.round(source.width*left),y=Math.round(source.height*top),w=Math.round(source.width*(right-left)),h=Math.round(source.height*(bottom-top));const canvas=document.createElement('canvas');canvas.width=w;canvas.height=h;canvas.getContext('2d').drawImage(source,x,y,w,h,0,0,w,h);return canvas;}

function amountInWords(value){
  const exact=money2(value),n=Math.trunc(exact); if(!exact)return '';
  const ones=['','One','Two','Three','Four','Five','Six','Seven','Eight','Nine','Ten','Eleven','Twelve','Thirteen','Fourteen','Fifteen','Sixteen','Seventeen','Eighteen','Nineteen'];
  const tens=['','','Twenty','Thirty','Forty','Fifty','Sixty','Seventy','Eighty','Ninety'];
  const under1000=(x)=>{const parts=[];if(x>=100){parts.push(ones[Math.floor(x/100)]+' Hundred');x%=100;}if(x>=20){parts.push(tens[Math.floor(x/10)]);if(x%10)parts.push(ones[x%10]);}else if(x)parts.push(ones[x]);return parts.join(' ')};
  const parts=[];let x=n;if(x>=1000000){parts.push(under1000(Math.floor(x/1000000))+' Million');x%=1000000;}if(x>=1000){parts.push(under1000(Math.floor(x/1000))+' Thousand');x%=1000;}if(x)parts.push(under1000(x));const piastres=Math.trunc((exact+1e-7)*100)%100;return `${parts.join(' ')||'Zero'} EGP${piastres?` #${String(piastres).padStart(2,'0')}#`:''}`;
}

async function createExcel(company, inv) {
  const ExcelJS = (await import('exceljs')).default;
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Invoice Sheet Maker';
  const ws = wb.addWorksheet('ملخص الفاتورة', {
    pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 1, horizontalCentered: true, verticalCentered: false, margins: { left:.22, right:.22, top:.22, bottom:.22, header:.1, footer:.1 } },
    views: [{ showGridLines: false }]
  });
  ws.pageSetup.printArea = 'A1:I31';
  ws.properties.defaultRowHeight = 17;
  const widths = [25,13,13,7,7,7,7,13,15]; widths.forEach((w,i)=> ws.getColumn(i+1).width=w);
  for (let r=1;r<=31;r++) for(let c=1;c<=9;c++) { const cell=ws.getCell(r,c); cell.font={name:'Arial',size:9.5,bold:true}; cell.alignment={vertical:'middle',horizontal:'center',wrapText:true}; }
  ws.getRow(1).height=22; ws.getRow(2).height=18;
  [['H4','I4'],['H5','I5'],['H6','I6'],['H9','I9'],['H10','I10'],['H11','I11']].forEach(([a,b])=>ws.mergeCells(`${a}:${b}`));
  ws.mergeCells('F8:I8');
  if(company.po){ws.mergeCells('A7:B7');ws.mergeCells('C7:E7');ws.getCell('A7').value='PO';ws.getCell('C7').value=company.po;for(let c=1;c<=5;c++){ws.getCell(7,c).border=thinBorder();ws.getCell(7,c).font={name:'Arial',size:10,bold:true};}ws.getRow(7).height=22;}
  ws.getCell('F4').value='DATE :'; ws.getCell('F5').value='INV.NO :'; ws.getCell('F6').value='ROOM NO :';
  ws.getCell('H4').value=inv.invoiceDate; ws.getCell('H5').value=inv.invoiceNo; ws.getCell('H6').value=inv.roomNo;
  ws.getCell('F8').value='BILL TO'; ws.getCell('F8').fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF111111'}}; ws.getCell('F8').font={name:'Arial',size:11,bold:true,color:{argb:'FFFFFFFF'}};
  ws.getCell('F9').value='GUEST Name :'; ws.getCell('F10').value='Company Name :'; ws.getCell('F11').value='Tax Registration :';
  ws.getCell('H9').value=inv.guest; ws.getCell('H10').value=company.name; ws.getCell('H11').value=company.tax;
  for(let r=4;r<=6;r++) for(let c=6;c<=9;c++) ws.getCell(r,c).border=thinBorder();
  for(let r=9;r<=11;r++) for(let c=6;c<=9;c++) ws.getCell(r,c).border=thinBorder();
  const heads=['DESCRIPTION','C/I','C/O','NGT','DBL','TPL','SGL','RATE E.P','AMOUNT'];
  heads.forEach((h,i)=>{const cell=ws.getCell(13,i+1);cell.value=h;cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF111111'}};cell.font={name:'Arial',size:10,bold:true,color:{argb:'FFFFFFFF'}};cell.border=thinBorder('FFFFFFFF');});
  ws.getRow(13).height=21;
  const accommodation = money2(inv.accommodation) || money2(inv.roomRate);
  const rows=[['Accommodation',inv.checkIn,inv.checkOut,inv.nights,money(inv.double)||'',money(inv.triple)||'',money(inv.single)||'',money2(inv.roomRate)||accommodation,accommodation],['Restaurant','','','','','','','',money2(inv.restaurant)],['Laundry','','','','','','','',money2(inv.laundry)]];
  rows.forEach((row,idx)=>row.forEach((v,i)=>ws.getCell(14+idx,i+1).value=v));
  for(let r=14;r<=21;r++){ws.getRow(r).height=21;for(let c=1;c<=9;c++)ws.getCell(r,c).border=thinBorder('FF888888');}
  ws.mergeCells('A23:F23'); ws.getCell('A23').value='OTHER COMMENTS'; ws.getCell('A23').fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF111111'}}; ws.getCell('A23').font={name:'Arial',size:11,bold:true,color:{argb:'FFFFFFFF'}};
  ws.mergeCells('G23:H23'); ws.getCell('G23').value='Sub Total'; ws.getCell('I23').value={formula:'SUM(I14:I21)'};
  ws.mergeCells('G25:H25'); ws.getCell('G25').value='TOTAL'; ws.getCell('I25').value={formula:'I23'};
  ws.mergeCells('C27:G27'); ws.getCell('C27').value=inv.amountWords;
  ws.mergeCells('A31:C31'); ws.getCell('A31').value='Credit Manager'; ws.mergeCells('G31:I31'); ws.getCell('G31').value='Make all Checks Payable To Pearl Hotel';
  ['H14','I14','I15','I16','I23','I25'].forEach(a=>ws.getCell(a).numFmt='#,##0.00');
  const buffer=await wb.xlsx.writeBuffer();
  const blob=new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const safeInvoice=String(inv.invoiceNo||dateForFile()).replace(/[\\/:*?"<>|]+/g,'-');
  return {blob,filename:`${company.code} - ${safeInvoice}.xlsx`};
}
function thinBorder(color='FF111111'){const b={style:'thin',color:{argb:color}};return{top:b,left:b,bottom:b,right:b};}

function App(){
  const [companies,setCompanies]=useState(()=>{const saved=JSON.parse(localStorage.getItem('invoice-companies')||'null')||[];const merged=[...saved];for(const base of seedCompanies)if(!merged.some(c=>c.code===base.code))merged.push(base);return merged.map(c=>({...c,invoiceTax:c.invoiceTax||String(c.tax||'').replace(/\D/g,'')}));});
  const [tab,setTab]=useState('create'); const [companyId,setCompanyId]=useState(companies[0]?.id||'');
  const [invoice,setInvoice]=useState(emptyInvoice); const [file,setFile]=useState(null); const [preview,setPreview]=useState('');
  const [folioFile,setFolioFile]=useState(null); const [folioPreview,setFolioPreview]=useState(''); const [folioChecks,setFolioChecks]=useState(null);
  const [progress,setProgress]=useState(0); const [reading,setReading]=useState(false); const [notice,setNotice]=useState('');
  const [history,setHistory]=useState(()=>{try{return JSON.parse(localStorage.getItem('invoice-export-history')||'[]')}catch{return []}});
  const [installPrompt,setInstallPrompt]=useState(null);
  const [draftRestored,setDraftRestored]=useState(false);
  const company=useMemo(()=>companies.find(c=>c.id===companyId)||companies[0],[companies,companyId]);
  const enteredTotal=money2(money2(invoice.accommodation)+money2(invoice.restaurant)+money2(invoice.laundry));
  const issues=useMemo(()=>validateInvoice(invoice,company,history),[invoice,company,history]);
  const badFields=new Set(issues.flatMap(item=>item.fields));
  const verified=issues.length===0;
  useEffect(()=>localStorage.setItem('invoice-companies',JSON.stringify(companies)),[companies]);
  useEffect(()=>localStorage.setItem('invoice-export-history',JSON.stringify(history)),[history]);
  useEffect(()=>{try{const saved=JSON.parse(localStorage.getItem('invoice-current-draft')||'null');if(saved?.invoice){setInvoice({...emptyInvoice,...saved.invoice});if(companies.some(c=>c.id===saved.companyId))setCompanyId(saved.companyId);}}catch{}setDraftRestored(true);},[]);
  useEffect(()=>{if(draftRestored)localStorage.setItem('invoice-current-draft',JSON.stringify({invoice,companyId}));},[invoice,companyId,draftRestored]);
  useEffect(()=>{const install=e=>{e.preventDefault();setInstallPrompt(e)};window.addEventListener('beforeinstallprompt',install);return()=>window.removeEventListener('beforeinstallprompt',install)},[]);
  useEffect(()=>{if('serviceWorker' in navigator)navigator.serviceWorker.register('/sw.js').catch(()=>{});},[]);
  const setInv=(k,v)=>setInvoice(p=>{const next={...p,[k]:v};if(k==='checkIn')next.invoiceDate=v;if(k==='nights'||k==='roomRate'){const n=money(next.nights),rate=money2(next.roomRate);if(n&&rate)next.accommodation=money2(n*rate);}return next;});
  const pickFile=(f)=>{if(!f)return;setFile(f);setPreview(old=>{if(old)URL.revokeObjectURL(old);return URL.createObjectURL(f)});setNotice('');};
  const pickFolio=(f)=>{if(!f)return;setFolioFile(f);setFolioPreview(old=>{if(old)URL.revokeObjectURL(old);return URL.createObjectURL(f)});setFolioChecks(null);setNotice('');};
  async function readImage(){if(!file)return;setReading(true);setProgress(3);try{
    const canvas=await preprocessForOcr(file);const {createWorker}=await import('tesseract.js');const worker=await createWorker('eng',1,{logger:m=>m.status==='recognizing text'&&setProgress(8+Math.round(m.progress*82))});
    const texts=[];for(const mode of ['4','6','11']){await worker.setParameters({tessedit_pageseg_mode:mode,preserve_interword_spaces:'1',tessedit_char_whitelist:''});texts.push((await worker.recognize(canvas)).data.text);}
    await worker.setParameters({tessedit_pageseg_mode:'6',preserve_interword_spaces:'1'});const description=(await worker.recognize(cropCanvas(canvas,.03,.50,.40,.87))).data.text;
    const parsed=parseInvoiceOcr(texts,description);let folio=null;
    if(folioFile){const folioCanvas=await preprocessForOcr(folioFile);const folioTexts=[];for(const mode of ['4','6','11']){await worker.setParameters({tessedit_pageseg_mode:mode,preserve_interword_spaces:'1'});folioTexts.push((await worker.recognize(folioCanvas)).data.text);}folio=parseGuestFolioOcr(folioTexts);}
    await worker.terminate();
    const matched=companies.find(c=>digitDistance(c.invoiceTax||c.tax,parsed.companyTax)<=1);if(matched)setCompanyId(matched.id);
    const calculatedOut=folio?.checkIn?addDays(folio.checkIn,parsed.nights):'';const shownFolioOut=folio?.actualCheckOut||folio?.printedCheckOut||'';
    const dateMatch=Boolean(calculatedOut&&shownFolioOut&&calculatedOut===shownFolioOut);const rateMatch=Boolean(folio?.accommodation&&parsed.accommodation&&Math.abs(folio.accommodation-parsed.accommodation)<=1);
    setFolioChecks(folio?{dateMatch,rateMatch,calculatedOut,shownFolioOut,folioAccommodation:folio.accommodation}:null);
    setInvoice(p=>{const next={...p,...Object.fromEntries(Object.entries(parsed).filter(([k,v])=>k!=='verified'&&v!==null))};if(folio){next.roomNo=folio.roomNo||next.roomNo;next.guest=folio.guest||next.guest;next.checkIn=folio.checkIn||next.checkIn;next.checkOut=calculatedOut||shownFolioOut||next.checkOut;if(folio.single||folio.double||folio.triple){next.single=folio.single;next.double=folio.double;next.triple=folio.triple;}}if(next.checkIn)next.invoiceDate=next.checkIn;next.amountWords=amountInWords(parsed.total);return next});
    setNotice(parsed.verified?`✓ تمت مطابقة الأرقام${matched?` واختيار شركة ${matched.code} تلقائيًا`:''}. الشيت جاهز.`:'⚠ لم تتطابق الأرقام تلقائيًا. راجع البنود حتى يساوي مجموعها إجمالي الفاتورة.');
  }catch(e){setNotice('تعذرت القراءة. جرّب صورة أوضح ومستقيمة، أو اكتب البيانات يدويًا.');}finally{setReading(false);setProgress(0);}}
  const finalInvoice=()=>({...invoice,invoiceDate:invoice.checkIn,amountWords:amountInWords(enteredTotal)});
  function recordExport(){setHistory(old=>[...old,{code:company.code,invoiceNo:String(invoice.invoiceNo).trim(),date:new Date().toISOString()}]);}
  function saveBlob(blob,filename){const url=URL.createObjectURL(blob);const link=document.createElement('a');link.href=url;link.download=filename;document.body.append(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);}
  async function exportExcel(share=false){
    if(issues.length){setNotice('راجع الأخطاء باللون الأحمر قبل التصدير.');return;}
    try{const {blob,filename}=await createExcel(company,finalInvoice());
      if(share&&navigator.canShare&&navigator.share){const sharedFile=new File([blob],filename,{type:blob.type});if(navigator.canShare({files:[sharedFile]})){await navigator.share({files:[sharedFile],title:filename});recordExport();setNotice('تم فتح مشاركة الملف.');return;}}
      saveBlob(blob,filename);recordExport();setNotice(share?'المشاركة غير مدعومة هنا، تم تنزيل Excel لتشاركه يدويًا.':`تم تنزيل ${filename}`);
    }catch(error){if(error?.name!=='AbortError')setNotice('تعذر تجهيز الملف أو مشاركته. جرّب التنزيل مباشرة.');}
  }
  function newInvoice(){setInvoice({...emptyInvoice});setFile(null);setFolioFile(null);setPreview('');setFolioPreview('');setFolioChecks(null);setNotice('تم فتح فاتورة جديدة.');}
  function printInvoice(){if(issues.length){setNotice('صحّح أخطاء المراجعة قبل الطباعة.');return;}window.print();}
  async function installApp(){if(!installPrompt)return;installPrompt.prompt();await installPrompt.userChoice;setInstallPrompt(null);}

  return <div className="app-shell">
    <header><div className="brand"><div className="brand-icon"><FileSpreadsheet/></div><div><h1>مُجهّز فواتير الشركات</h1><p>حوّل الفاتورة الضريبية إلى شيت جاهز للطباعة</p></div></div><nav>{installPrompt&&<button onClick={installApp}><Download/>تثبيت التطبيق</button>}<button className={tab==='create'?'active':''} onClick={()=>setTab('create')}><ScanText/>إنشاء شيت</button><button className={tab==='companies'?'active':''} onClick={()=>setTab('companies')}><Building2/>الشركات</button></nav></header>
    <main>{tab==='create'?<>
      <section className="step-card company-strip"><div><span className="step-no">1</span><div><h2>اختار الشركة</h2><p>البيانات الثابتة هتتحط تلقائيًا في الشيت</p></div></div><select value={companyId} onChange={e=>setCompanyId(e.target.value)}>{companies.map(c=><option key={c.id} value={c.id}>{c.code} — {c.name}</option>)}</select></section>
      <div className="workspace"><section className="step-card upload-panel"><div className="section-title"><span className="step-no">2</span><div><h2>ارفع صور الفاتورة</h2><p>الضريبية أساسية وGuest Folio اختياري</p></div></div><div className="upload-label">1 — الفاتورة الضريبية</div><label className="camera-button"><ScanText/> صوّر الفاتورة <input type="file" accept="image/*" capture="environment" onChange={e=>pickFile(e.target.files[0])}/></label><small className="camera-note">تفتح كاميرا الموبايل إن كانت مدعومة. تحسين التباين والإضاءة يتم تلقائيًا قبل OCR؛ قص حدود الورقة وتصحيح المنظور غير متاحين تلقائيًا بعد.</small><label className={`dropzone ${preview?'has-file':''}`}><input type="file" accept="image/*" onChange={e=>pickFile(e.target.files[0])}/>{preview?<img src={preview} alt="معاينة الفاتورة الضريبية"/>:<><UploadCloud/><strong>اضغط لاختيار الصورة الأساسية</strong><span>JPG أو PNG</span></>}</label><div className="upload-label optional">2 — Guest Folio <span>اختياري</span></div><label className={`dropzone folio-zone ${folioPreview?'has-file':''}`}><input type="file" accept="image/*" onChange={e=>pickFolio(e.target.files[0])}/>{folioPreview?<img src={folioPreview} alt="معاينة Guest Folio"/>:<><FileImage/><strong>أضف صورة تكملة الفندق</strong><span>للغرفة والتواريخ ونوع الإشغال</span></>}</label><button className="primary wide" disabled={!file||reading} onClick={readImage}>{reading?<><span className="spinner"/>جاري القراءة {progress}%</>:<><ScanText/>قراءة ومطابقة البيانات مجانًا</>}</button><div className="privacy"><Check/>القراءة تتم على جهازك ولا تحتاج API أو اشتراك</div></section>
      <section className="step-card form-panel"><div className="section-title"><span className="step-no">3</span><div><h2>راجع البيانات</h2><p>عدّل أي خانة قبل إنشاء الشيت</p></div></div><div className="form-grid">
        <Field invalid={badFields.has('invoiceNo')} label="رقم الفاتورة" value={invoice.invoiceNo} onChange={v=>setInv('invoiceNo',v)}/><Field invalid={badFields.has('invoiceDate')} label="تاريخ الفاتورة" value={invoice.invoiceDate} onChange={v=>setInv('invoiceDate',v)} placeholder="13-08-2026"/><Field invalid={badFields.has('roomNo')} label="رقم الغرفة" value={invoice.roomNo} onChange={v=>setInv('roomNo',v)}/><Field invalid={badFields.has('guest')} label="اسم النزيل" value={invoice.guest} onChange={v=>setInv('guest',v)} wide/>
        <Field invalid={badFields.has('checkIn')} label="Check in" value={invoice.checkIn} onChange={v=>setInv('checkIn',v)} placeholder="11-08-2026"/><Field invalid={badFields.has('checkOut')} label="Check out" value={invoice.checkOut} onChange={v=>setInv('checkOut',v)} placeholder="12-08-2026"/><Field invalid={badFields.has('nights')} label="عدد الليالي NGT" type="number" value={invoice.nights} onChange={v=>setInv('nights',v)}/><Field label="سنجل SGL" type="number" value={invoice.single} onChange={v=>setInv('single',v)}/><Field label="دابل DBL" type="number" value={invoice.double} onChange={v=>setInv('double',v)}/><Field label="ترابل TPL" type="number" value={invoice.triple} onChange={v=>setInv('triple',v)}/><Field invalid={badFields.has('roomRate')} label="سعر الليلة" type="number" value={invoice.roomRate} onChange={v=>setInv('roomRate',v)}/><Field invalid={badFields.has('accommodation')} label="إجمالي الإقامة" type="number" value={invoice.accommodation} onChange={v=>setInv('accommodation',v)}/><Field invalid={badFields.has('restaurant')} label="المطعم / Food & Drink" type="number" value={invoice.restaurant} onChange={v=>setInv('restaurant',v)}/><Field invalid={badFields.has('laundry')} label="Laundry" type="number" value={invoice.laundry} onChange={v=>setInv('laundry',v)}/><Field label="الإجمالي المكتوب بالفاتورة" type="number" invalid={badFields.has('total')} value={invoice.total} onChange={v=>setInv('total',v)}/><Field label="التفقيط بالإنجليزي" value={invoice.amountWords} onChange={v=>setInv('amountWords',v)} placeholder="Two Thousand One Hundred EGP" wide/>
      </div><section className={`audit-box ${issues.length?'audit-error':'audit-good'}`}><strong>{issues.length?`⚠ ${issues.length} تنبيه لازم تراجعه قبل التصدير`:'✓ اكتملت فحوصات الفاتورة'}</strong>{issues.length>0&&<ul>{issues.map((issue,i)=><li key={i}>{issue.message}</li>)}</ul>}</section>{folioChecks&&<div className="folio-checks"><div className={folioChecks.dateMatch?'check-ok':'check-warn'}><strong>{folioChecks.dateMatch?'✓ التواريخ مطابقة':'⚠ تم تصحيح تاريخ المغادرة'}</strong><small>المحسوب: {folioChecks.calculatedOut||'—'} · المكتوب: {folioChecks.shownFolioOut||'—'}</small></div><div className={folioChecks.rateMatch?'check-ok':'check-warn'}><strong>{folioChecks.rateMatch?'✓ قيمة الإقامة مطابقة':'⚠ قيمة الإقامة غير مطابقة'}</strong><small>المعتمد من الضريبية: {money2(invoice.accommodation).toFixed(2)} · بالفوليو: {money2(folioChecks.folioAccommodation).toFixed(2)}</small></div></div>}<div className="total-row"><span>الإجمالي المحسوب</span><strong>{enteredTotal.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2})} EGP</strong></div><div className={`match-status ${verified?'ok':'bad'}`}><span>{verified?'✓ الأرقام متطابقة':'⚠ الأرقام غير متطابقة'}</span><small>إجمالي الفاتورة: {money2(invoice.total).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2})} EGP</small></div><button className="download" disabled={!verified} onClick={()=>exportExcel(false)}><Download/>تنزيل Excel <small>{verified?`${company?.code} - ${String(invoice.invoiceNo||'رقم الفاتورة').replaceAll('/','-')}.xlsx`:'صحّح التنبيهات أولًا'}</small></button><div className="export-actions"><button onClick={newInvoice}>فاتورة جديدة</button><button disabled={!verified} onClick={printInvoice}>طباعة A4 / حفظ PDF</button><button disabled={!verified} onClick={()=>exportExcel(true)}>مشاركة Excel</button><button onClick={()=>{localStorage.setItem('invoice-current-draft',JSON.stringify({invoice,companyId}));setNotice('تم حفظ المسودة على الجهاز.')}}>حفظ المسودة</button></div><small className="camera-note">PDF: اختار «حفظ كـ PDF» من نافذة الطباعة. المسودة وسجل الأرقام محفوظين على هذا الجهاز فقط.</small>{notice&&<div className="notice">{notice}</div>}</section></div>
      <div className="print-only" dir="ltr"><h2>INVOICE</h2><p>{company?.name}</p><p>Invoice: {invoice.invoiceNo} | Date: {invoice.checkIn} | Room: {invoice.roomNo}</p><p>Guest: {invoice.guest} {company?.po?`| PO: ${company.po}`:''}</p><table><thead><tr><th>Description</th><th>Check-in</th><th>Check-out</th><th>Nights</th><th>Rate</th><th>Amount EGP</th></tr></thead><tbody><tr><td>Accommodation</td><td>{invoice.checkIn}</td><td>{invoice.checkOut}</td><td>{invoice.nights}</td><td>{invoice.roomRate}</td><td>{invoice.accommodation}</td></tr><tr><td>Restaurant</td><td colSpan="4"></td><td>{invoice.restaurant}</td></tr><tr><td>Laundry</td><td colSpan="4"></td><td>{invoice.laundry}</td></tr><tr><th colSpan="5">TOTAL</th><th>{enteredTotal.toFixed(2)}</th></tr></tbody></table><p>{amountInWords(enteredTotal)}</p></div>
    </>:<Companies companies={companies} setCompanies={setCompanies} setCompanyId={setCompanyId}/>}</main>
  </div>
}

function Field({label,value,onChange,type='text',placeholder='',wide=false,invalid=false}){return <label className={`${wide?'field wide':'field'} ${invalid?'field-invalid':''}`}><span>{label}</span><input type={type} value={value} placeholder={placeholder} onChange={e=>onChange(e.target.value)}/></label>}
function Companies({companies,setCompanies,setCompanyId}){const [draft,setDraft]=useState({code:'',name:'',tax:'',invoiceTax:'',po:''});const add=()=>{if(!draft.code||!draft.name)return;const item={...draft,invoiceTax:draft.invoiceTax||String(draft.tax).replace(/\D/g,''),id:crypto.randomUUID()};setCompanies(p=>[...p,item]);setCompanyId(item.id);setDraft({code:'',name:'',tax:'',invoiceTax:'',po:''});};return <section className="companies-page"><div className="page-head"><div><h2>بيانات الشركات</h2><p>سجّل البيانات الثابتة مرة واحدة فقط</p></div><Building2/></div><div className="company-form"><Field label="الكود المختصر" value={draft.code} onChange={v=>setDraft(p=>({...p,code:v.toUpperCase()}))} placeholder="EXPRO"/><Field label="اسم الشركة" value={draft.name} onChange={v=>setDraft(p=>({...p,name:v}))}/><Field label="Tax Registration في الشيت" value={draft.tax} onChange={v=>setDraft(p=>({...p,tax:v}))}/><Field label="رقم الشركة داخل الفاتورة" value={draft.invoiceTax} onChange={v=>setDraft(p=>({...p,invoiceTax:v}))} placeholder="للتعرف التلقائي"/><Field label="PO" value={draft.po} onChange={v=>setDraft(p=>({...p,po:v}))}/><button className="primary" onClick={add}><Plus/>إضافة الشركة</button></div><div className="company-list">{companies.map(c=><div className="company-item" key={c.id}><div className="company-badge">{c.code.slice(0,2)}</div><div className="company-data"><strong>{c.code}</strong><span>{c.name}</span><small>Tax: {c.tax||'—'} · تعريف الفاتورة: {c.invoiceTax||'—'} · PO: {c.po||'—'}</small></div><button className="danger" aria-label="حذف" onClick={()=>setCompanies(p=>p.filter(x=>x.id!==c.id))}><Trash2/></button></div>)}</div></section>}

createRoot(document.getElementById('root')).render(<App/>);
