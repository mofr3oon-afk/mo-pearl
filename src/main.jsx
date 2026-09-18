import React, { useEffect, useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { Building2, Check, Download, FileImage, FileSpreadsheet, Plus, Save, ScanText, Settings2, Trash2, UploadCloud, X, Search, History, RefreshCw } from 'lucide-react';
import './styles.css';
import { parseGuestFolioOcr, parseInvoiceOcr } from './ocr-parser.js';
import { validateInvoice } from './validation.js';
import { roundMoney } from './money.js';
import { readDraftMedia, writeDraftMedia } from './draft-media.js';
import { comparePo, extractPo } from './po-match.js';

const seedCompanies = [
  { id: crypto.randomUUID(), code: 'EXPRO', name: 'EXPRO EGYPT LLC', tax: '204-829-542', invoiceTax: '205938604', po: '4300228298' },
  { id: crypto.randomUUID(), code: 'WEATHERFORD', name: 'WeatherFord', tax: '706-010-205', invoiceTax: '706010205', po: '' }
];
const emptyInvoice = { invoiceNo: '', invoiceDate: '', roomNo: '', guest: '', checkIn: '', checkOut: '', nights: 1, single: 1, double: 0, triple: 0, roomRate: '', accommodation: '', restaurant: '', laundry: '', total: '', amountWords: '' };
const money = (v) => Number(String(v || 0).replace(/,/g, '')) || 0;
const money2 = (v) => roundMoney(v);
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
  const parts=[];let x=n;if(x>=1000000){parts.push(under1000(Math.floor(x/1000000))+' Million');x%=1000000;}if(x>=1000){parts.push(under1000(Math.floor(x/1000))+' Thousand');x%=1000;}if(x)parts.push(under1000(x));const piastres=Math.round(exact*100)%100;return `${parts.join(' ')||'Zero'} EGP${piastres?` #${String(piastres).padStart(2,'0')}#`:''}`;
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
  const [folioFiles,setFolioFiles]=useState([]); const [folioChecks,setFolioChecks]=useState(null);
  const [progress,setProgress]=useState(0); const [reading,setReading]=useState(false); const [notice,setNotice]=useState('');
  const [history,setHistory]=useState(()=>{try{return JSON.parse(localStorage.getItem('invoice-export-history')||'[]')}catch{return []}});
  const [poFile,setPoFile]=useState(null); const [poPreview,setPoPreview]=useState(''); const [poText,setPoText]=useState(''); const [poReading,setPoReading]=useState(false);
  const [invoicePo,setInvoicePo]=useState(companies[0]?.po||''); const [imageSide,setImageSide]=useState('invoice');
  const [historyQuery,setHistoryQuery]=useState(''); const [restoredId,setRestoredId]=useState(''); const [previewHistory,setPreviewHistory]=useState(null);
  const [installPrompt,setInstallPrompt]=useState(null);
  const [draftRestored,setDraftRestored]=useState(false);
  const [mediaRestored,setMediaRestored]=useState(false);
  const [extraPages,setExtraPages]=useState([]);
  const [draftStatus,setDraftStatus]=useState('');
  const company=useMemo(()=>companies.find(c=>c.id===companyId)||companies[0],[companies,companyId]);
  const enteredTotal=money2(money2(invoice.accommodation)+money2(invoice.restaurant)+money2(invoice.laundry));
  const issues=useMemo(()=>validateInvoice(invoice,company,history.filter(h=>!restoredId||h.id!==restoredId)),[invoice,company,history,restoredId]);
  const poChecks=useMemo(()=>comparePo(poText,company,invoice,invoicePo),[poText,company,invoice,invoicePo]);
  const poErrors=[];
  const filteredHistory=useMemo(()=>history.filter(h=>[h.code,h.companyName,h.invoiceNo,h.guest,h.roomNo,h.invoiceDate].some(v=>String(v||'').toLowerCase().includes(historyQuery.toLowerCase()))).slice().reverse(),[history,historyQuery]);
  const badFields=new Set(issues.flatMap(item=>item.fields));
  const verified=issues.length===0 && poErrors.length===0;
  useEffect(()=>localStorage.setItem('invoice-companies',JSON.stringify(companies)),[companies]);
  useEffect(()=>localStorage.setItem('invoice-export-history',JSON.stringify(history)),[history]);
  useEffect(()=>{try{const saved=JSON.parse(localStorage.getItem('invoice-current-draft')||'null');if(saved?.invoice){setInvoice({...emptyInvoice,...saved.invoice});if(companies.some(c=>c.id===saved.companyId))setCompanyId(saved.companyId);if(saved.invoicePo!==undefined)setInvoicePo(saved.invoicePo);}}catch{}setDraftRestored(true);},[]);
  useEffect(()=>{if(!draftRestored)return;try{localStorage.setItem('invoice-current-draft',JSON.stringify({invoice,companyId,invoicePo}));setDraftStatus('البيانات محفوظة تلقائيًا');}catch{setDraftStatus('تعذر حفظ البيانات: مساحة الجهاز ممتلئة أو التخزين محظور');}},[invoice,companyId,invoicePo,draftRestored]);
  useEffect(()=>{let active=true;readDraftMedia().then(saved=>{if(!active||!saved)return;
    if(saved.invoice){setFile(saved.invoice);setPreview(URL.createObjectURL(saved.invoice));}
    if(saved.folio){setFolioFiles([{file:saved.folio,url:URL.createObjectURL(saved.folio)}]);}
    if(saved.po){setPoFile(saved.po);setPoPreview(URL.createObjectURL(saved.po));}
    if(Array.isArray(saved.pages))setExtraPages(saved.pages.map(f=>({file:f,url:URL.createObjectURL(f)})));
  }).catch(()=>{}).finally(()=>{if(active)setMediaRestored(true)});return()=>{active=false};},[]);
  useEffect(()=>{if(!mediaRestored)return;writeDraftMedia({invoice:file,folio:folioFiles[0]?.file||null,po:poFile,pages:extraPages.map(p=>p.file)}).catch(()=>setDraftStatus('تم حفظ البيانات، لكن صور المسودة لم تُحفظ (راجع مساحة الجهاز)'));},[file,folioFiles,poFile,extraPages,mediaRestored]);
  useEffect(()=>{const install=e=>{e.preventDefault();setInstallPrompt(e)};window.addEventListener('beforeinstallprompt',install);return()=>window.removeEventListener('beforeinstallprompt',install)},[]);
  useEffect(()=>{if('serviceWorker' in navigator)navigator.serviceWorker.register('/sw.js').catch(()=>{});},[]);
  const setInv=(k,v)=>setInvoice(p=>{const next={...p,[k]:v};if(k==='checkIn')next.invoiceDate=v;if(k==='nights'||k==='roomRate'){const n=money(next.nights),rate=money2(next.roomRate);if(n&&rate)next.accommodation=money2(n*rate);}return next;});
  const pickFile=(f)=>{if(!f)return;setFile(f);setPreview(old=>{if(old)URL.revokeObjectURL(old);return URL.createObjectURL(f)});setNotice('');};
  const appendPages=(files)=>{const additions=Array.from(files||[]).filter(f=>f.type.startsWith('image/'));
    if(!additions.length)return;
    setExtraPages(old=>[...old,...additions.map(f=>({file:f,url:URL.createObjectURL(f)}))]);
    setNotice(`تمت إضافة ${additions.length} صفحة. اضغط قراءة لقراءة جميع الصفحات معًا.`);
  };
  const removePage=(index)=>setExtraPages(old=>old.filter((page,i)=>{if(i===index){URL.revokeObjectURL(page.url);return false;}return true;}));
  const pickPo=(f)=>{if(!f)return;setPoFile(f);setPoText('');setPoPreview(old=>{if(old)URL.revokeObjectURL(old);return URL.createObjectURL(f)});setNotice('اضغط قراءة PO لمراجعة الرقم واسم الشركة.');};
  async function readPo(){if(!poFile)return;setPoReading(true);try{const worker=(await import('tesseract.js')).createWorker;const instance=await worker('eng');const image=await preprocessForOcr(poFile);const result=await instance.recognize(image);await instance.terminate();setPoText(result.data.text);const extracted=extractPo(result.data.text);if(extracted.number&&!invoicePo)setInvoicePo(extracted.number);setNotice('تمت قراءة أمر الشراء. راجع نتائج المطابقة يدويًا.');}catch(e){setNotice('تعذرت قراءة PO. جرّب صورة أوضح أو اكتب الرقم يدويًا.');}finally{setPoReading(false);}}
  const pickFolios=async (files)=>{
    const incoming=Array.from(files||[]); if(!incoming.length)return;
    const images=[]; const rejected=[];
    for(const f of incoming){
      if(f.type.startsWith('image/'))images.push({file:f,url:URL.createObjectURL(f)});
      else if(/\.xlsx?$/i.test(f.name)){
        try{const ExcelJS=(await import('exceljs')).default;const wb=new ExcelJS.Workbook();await wb.xlsx.load(await f.arrayBuffer());
          const lines=[];wb.eachSheet(ws=>ws.eachRow(row=>lines.push(row.values.slice(1).map(cell=>String(cell?.text??cell?.result??cell??'')).join('  '))));
          images.push({file:f,url:'',text:lines.join('\n'),kind:'excel'});
        }catch{rejected.push(`${f.name}: تعذرت قراءة Excel`);}
      }else if(f.type==='application/pdf'||/\.pdf$/i.test(f.name)){
        images.push({file:f,url:URL.createObjectURL(f),kind:'pdf'});
      }else rejected.push(`${f.name}: نوع غير مدعوم`);
    }
    setFolioFiles(old=>[...old,...images]);setFolioChecks(null);
    setNotice(`${images.length} ملف مضاف. ${rejected.join('، ')}${images.some(f=>f.kind==='pdf')?' تنبيه: PDF متاح للعرض فقط في النسخة دي؛ حوّله لصور لاستخراج البيانات.':''}`);
  };
  const removeFolio=(i)=>setFolioFiles(old=>old.filter((f,j)=>{if(i===j){if(f.url)URL.revokeObjectURL(f.url);return false;}return true;}));
  async function readImage(){if(!file)return;setReading(true);setProgress(3);try{
    const canvas=await preprocessForOcr(file);const {createWorker}=await import('tesseract.js');const worker=await createWorker('eng',1,{logger:m=>m.status==='recognizing text'&&setProgress(8+Math.round(m.progress*82))});
    const texts=[];for(const mode of ['4','6','11']){await worker.setParameters({tessedit_pageseg_mode:mode,preserve_interword_spaces:'1',tessedit_char_whitelist:''});texts.push((await worker.recognize(canvas)).data.text);}
    await worker.setParameters({tessedit_pageseg_mode:'6',preserve_interword_spaces:'1'});const description=(await worker.recognize(cropCanvas(canvas,.03,.50,.40,.87))).data.text;
    for(const page of extraPages){const extraCanvas=await preprocessForOcr(page.file);await worker.setParameters({tessedit_pageseg_mode:'6',preserve_interword_spaces:'1'});
      const recognized=(await worker.recognize(extraCanvas)).data.text;
      for(let i=0;i<texts.length;i++)texts[i]+='\n'+recognized;
    }
    const parsed=parseInvoiceOcr(texts,description);let folio=null;
    if(folioFiles.length){const folioTexts=[];for(const doc of folioFiles){
      if(doc.kind==='pdf')continue;
      if(doc.kind==='excel'){folioTexts.push(doc.text);continue;}
      const folioCanvas=await preprocessForOcr(doc.file);
      for(const mode of ['4','6','11']){await worker.setParameters({tessedit_pageseg_mode:mode,preserve_interword_spaces:'1'});folioTexts.push((await worker.recognize(folioCanvas)).data.text);}
    }if(folioTexts.length)folio=parseGuestFolioOcr(folioTexts);}
    await worker.terminate();
    const detectedPo=extractPo(texts.join('\n')).number;const matched=companies.find(c=>digitDistance(c.invoiceTax||c.tax,parsed.companyTax)<=1);if(matched){setCompanyId(matched.id);setInvoicePo(detectedPo||matched.po||'');}else if(detectedPo)setInvoicePo(detectedPo);
    else setNotice('الشركة لم تتحدد تلقائيًا؛ اخترها من القائمة وسجّل بياناتها عند الحاجة.');
    const calculatedOut=folio?.checkIn?addDays(folio.checkIn,parsed.nights):'';const shownFolioOut=folio?.actualCheckOut||folio?.printedCheckOut||'';
    const dateMatch=Boolean(calculatedOut&&shownFolioOut&&calculatedOut===shownFolioOut);const rateMatch=Boolean(folio?.accommodation&&parsed.accommodation&&Math.abs(folio.accommodation-parsed.accommodation)<=1);
    setFolioChecks(folio?{dateMatch,rateMatch,calculatedOut,shownFolioOut,folioAccommodation:folio.accommodation}:null);
    setInvoice(p=>{const next={...p,...Object.fromEntries(Object.entries(parsed).filter(([k,v])=>k!=='verified'&&v!==null))};if(folio){next.roomNo=folio.roomNo||next.roomNo;next.guest=folio.guest||next.guest;next.checkIn=folio.checkIn||next.checkIn;next.checkOut=calculatedOut||shownFolioOut||next.checkOut;if(folio.single||folio.double||folio.triple){next.single=folio.single;next.double=folio.double;next.triple=folio.triple;}}if(next.checkIn)next.invoiceDate=next.checkIn;next.amountWords=amountInWords(parsed.total);return next});
    setNotice(folioFiles.some(f=>f.kind==='pdf')?'تنبيه: لم تتم قراءة PDF تلقائيًا؛ حوّل صفحاته لصور وأعد القراءة.':parsed.verified?`✓ تمت مطابقة الأرقام${matched?` واختيار شركة ${matched.code} تلقائيًا`:''}. الشيت جاهز.`:'⚠ لم تتطابق الأرقام تلقائيًا. راجع البنود حتى يساوي مجموعها إجمالي الفاتورة.');
  }catch(e){setNotice('تعذرت القراءة. جرّب صورة أوضح ومستقيمة، أو اكتب البيانات يدويًا.');}finally{setReading(false);setProgress(0);}}
  const finalInvoice=()=>({...invoice,invoiceDate:invoice.checkIn,amountWords:amountInWords(enteredTotal)});
  function recordExport(){setRestoredId('');setHistory(old=>{const record={id:crypto.randomUUID(),code:company.code,companyName:company.name,companyId:company.id,invoiceNo:String(invoice.invoiceNo).trim(),invoiceDate:invoice.invoiceDate,guest:invoice.guest,roomNo:invoice.roomNo,po:invoicePo,invoice:{...finalInvoice()},date:new Date().toISOString(),saved:true};return [...old.filter(h=>!(h.code===record.code&&h.invoiceNo===record.invoiceNo)),record];});}
  function restoreHistory(item){if(!item.invoice){setNotice('السجل القديم يحتوي على رقم الفاتورة فقط ولا يمكن استعادة باقي بياناته.');setTab('create');return;}setInvoice({...emptyInvoice,...item.invoice});setRestoredId(item.id||'');const matched=companies.find(c=>c.id===item.companyId||c.code===item.code);if(matched)setCompanyId(matched.id);setInvoicePo(item.po||matched?.po||'');setFile(null);setPreview('');setFolioFiles([]);setPoFile(null);setPoPreview('');setPoText('');setExtraPages([]);setTab('create');setNotice('تم استرجاع بيانات الفاتورة. الصورة الأصلية ليست مخزنة؛ ارفعها مرة أخرى للمراجعة.');}
  function deleteHistory(id){if(window.confirm('حذف الفاتورة من السجل على هذا الجهاز؟'))setHistory(old=>old.filter(x=>x.id!==id));}
  function saveBlob(blob,filename){const url=URL.createObjectURL(blob);const link=document.createElement('a');link.href=url;link.download=filename;document.body.append(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);}
  async function exportExcel(share=false){
    if(!verified){setNotice('راجع الأخطاء والمطابقة مع PO قبل التصدير.');return;}
    try{const {blob,filename}=await createExcel({...company,po:invoicePo},finalInvoice());
      if(share&&navigator.share){const sharedFile=new File([blob],filename,{type:blob.type});if(navigator.canShare?.({files:[sharedFile]})){await navigator.share({files:[sharedFile],title:filename});recordExport();setNotice('تمت مشاركة الملف أو إرساله للتطبيق المحدد.');return;}}
      saveBlob(blob,filename);recordExport();setNotice(share?'المشاركة غير مدعومة هنا، تم تنزيل Excel لتشاركه يدويًا.':`تم تنزيل ${filename}`);
    }catch(error){if(error?.name==='AbortError'){setNotice('تم إلغاء المشاركة بدون تنزيل.');return;}setNotice('تعذر تجهيز الملف أو مشاركته. جرّب التنزيل مباشرة.');}
  }
  function newInvoice(){setRestoredId('');setInvoice({...emptyInvoice});setFile(null);setFolioFiles([]);setPoFile(null);setPoText('');setExtraPages(old=>{old.forEach(p=>URL.revokeObjectURL(p.url));return [];});setInvoicePo(company?.po||'');setPreview('');setPoPreview('');setFolioChecks(null);setNotice('تم فتح فاتورة جديدة.');}
  function printInvoice(){if(!verified){setNotice('صحّح أخطاء المراجعة قبل الطباعة.');return;}window.requestAnimationFrame(()=>window.print());}
  async function installApp(){if(!installPrompt)return;installPrompt.prompt();await installPrompt.userChoice;setInstallPrompt(null);}

  return <div className="app-shell">
    <header><div className="brand"><div className="brand-icon"><FileSpreadsheet/></div><div><h1>مُجهّز فواتير الشركات</h1><p>حوّل الفاتورة الضريبية إلى شيت جاهز للطباعة</p></div></div><nav>{installPrompt&&<button onClick={installApp}><Download/>تثبيت التطبيق</button>}<button className={tab==='create'?'active':''} onClick={()=>setTab('create')}><ScanText/>إنشاء شيت</button><button className={tab==='companies'?'active':''} onClick={()=>setTab('companies')}><Building2/>الشركات</button><button className={tab==='history'?'active':''} onClick={()=>setTab('history')}><History/>السجل</button></nav></header>
    <main>{tab==='create'?<>
      <div className="mobile-hero"><strong>فاتورة جديدة</strong><span>صوّر المستند، راجع البيانات، وصدّر الشيت</span><label className="mobile-main-camera"><ScanText/> صوّر الفاتورة الآن <input type="file" accept="image/*" capture="environment" onChange={e=>{pickFile(e.target.files[0]);e.target.value='';}}/></label><small>{draftStatus||'البيانات تحفظ تلقائيًا على هذا الجهاز'}</small></div>
      <section className="step-card company-strip"><div><span className="step-no">1</span><div><h2>اختار الشركة</h2><p>البيانات الثابتة هتتحط تلقائيًا في الشيت</p></div></div><select value={companyId} onChange={e=>{setCompanyId(e.target.value);setInvoicePo(companies.find(c=>c.id===e.target.value)?.po||'');}}>{companies.map(c=><option key={c.id} value={c.id}>{c.code} — {c.name}</option>)}</select></section>
      <div className="workspace"><section className="step-card upload-panel"><div className="section-title"><span className="step-no">2</span><div><h2>ارفع صور الفاتورة</h2><p>الضريبية أساسية وGuest Folio اختياري</p></div></div><div className="upload-label">1 — الفاتورة الضريبية</div><label className="camera-button"><ScanText/> صوّر الفاتورة <input type="file" accept="image/*" capture="environment" onChange={e=>pickFile(e.target.files[0])}/></label><small className="camera-note">تفتح كاميرا الموبايل إن كانت مدعومة. تحسين التباين والإضاءة يتم تلقائيًا قبل OCR؛ قص حدود الورقة وتصحيح المنظور غير متاحين تلقائيًا بعد.</small><label className={`dropzone ${preview?'has-file':''}`}><input type="file" accept="image/*" onChange={e=>pickFile(e.target.files[0])}/>{preview?<img src={preview} alt="معاينة الفاتورة الضريبية"/>:<><UploadCloud/><strong>اضغط لاختيار الصورة الأساسية</strong><span>JPG أو PNG</span></>}</label><div className="upload-label optional">2 — Guest Folio <span>اختياري</span></div>
      <label className="camera-button"><Plus/> تصوير ورقة Guest Folio <input type="file" accept="image/*" capture="environment" onChange={e=>{pickFolios(e.target.files);e.target.value='';}}/></label>
      <label className="camera-button secondary-camera"><FileImage/> إضافة صور أو Excel أو PDF <input type="file" accept="image/*,.xlsx,.pdf,application/pdf" multiple onChange={e=>{pickFolios(e.target.files);e.target.value='';}}/></label>
      {!!folioFiles.length&&<div className="page-thumbnails">{folioFiles.map((doc,i)=><div className="page-thumb" key={doc.url||doc.file.name+i}>{doc.kind==='excel'?<FileSpreadsheet/>:doc.kind==='pdf'?<FileImage/>:<img src={doc.url} alt={`Guest Folio ${i+1}`}/>}<span>{doc.file.name}</span><button type="button" aria-label="حذف الملف" onClick={()=>removeFolio(i)}><X/></button></div>)}</div>}
      <small className="camera-note">الصور وExcel يدخلوا في القراءة. PDF يمكن اختياره وعرضه، لكن استخراج بياناته يتطلب تحويل صفحاته لصور حاليًا.</small>
      <button className="primary wide" disabled={!file||reading} onClick={readImage}>{reading?<><span className="spinner"/>جاري القراءة {progress}%</>:<><ScanText/>قراءة ومطابقة البيانات مجانًا</>}</button><div className="privacy"><Check/>القراءة تتم على جهازك ولا تحتاج API أو اشتراك</div></section>
      <section className="step-card form-panel"><div className="section-title"><span className="step-no">3</span><div><h2>راجع البيانات</h2><p>عدّل أي خانة قبل إنشاء الشيت</p></div></div><div className="review-layout"><div className="review-original"><div className="review-switch"><button className={imageSide==='invoice'?'selected':''} onClick={()=>setImageSide('invoice')}>الضريبية</button>{folioFiles.map((doc,i)=><button key={i} className={imageSide===`folio-${i}`?'selected':''} onClick={()=>setImageSide(`folio-${i}`)}>Folio {i+1}</button>)}</div>{imageSide==='invoice'&&preview?<a href={preview} target="_blank" rel="noreferrer"><img src={preview} alt="الفاتورة الأصلية"/></a>:imageSide.startsWith('folio-')&&folioFiles[Number(imageSide.split('-')[1])]?<div>{folioFiles[Number(imageSide.split('-')[1])].kind==='pdf'?<iframe title="Guest Folio PDF" src={folioFiles[Number(imageSide.split('-')[1])].url} width="100%" height="470"/>:folioFiles[Number(imageSide.split('-')[1])].kind==='excel'?<pre className="folio-text">{folioFiles[Number(imageSide.split('-')[1])].text}</pre>:<a href={folioFiles[Number(imageSide.split('-')[1])].url} target="_blank" rel="noreferrer"><img src={folioFiles[Number(imageSide.split('-')[1])].url} alt="Guest Folio"/></a>}</div>:<p>ارفع الفاتورة أو Guest Folio لعرضهما هنا.</p>}</div><div className="form-grid">
        <Field invalid={badFields.has('invoiceNo')} label="رقم الفاتورة" value={invoice.invoiceNo} onChange={v=>setInv('invoiceNo',v)}/><Field invalid={badFields.has('invoiceDate')} label="تاريخ الفاتورة" value={invoice.invoiceDate} onChange={v=>setInv('invoiceDate',v)} placeholder="13-08-2026"/><Field invalid={badFields.has('roomNo')} label="رقم الغرفة" value={invoice.roomNo} onChange={v=>setInv('roomNo',v)}/><Field invalid={badFields.has('guest')} label="اسم النزيل" value={invoice.guest} onChange={v=>setInv('guest',v)} wide/>
        <Field invalid={badFields.has('checkIn')} label="Check in" value={invoice.checkIn} onChange={v=>setInv('checkIn',v)} placeholder="11-08-2026"/><Field invalid={badFields.has('checkOut')} label="Check out" value={invoice.checkOut} onChange={v=>setInv('checkOut',v)} placeholder="12-08-2026"/><Field invalid={badFields.has('nights')} label="عدد الليالي NGT" type="number" value={invoice.nights} onChange={v=>setInv('nights',v)}/><Field label="سنجل SGL" type="number" value={invoice.single} onChange={v=>setInv('single',v)}/><Field label="دابل DBL" type="number" value={invoice.double} onChange={v=>setInv('double',v)}/><Field label="ترابل TPL" type="number" value={invoice.triple} onChange={v=>setInv('triple',v)}/><Field invalid={badFields.has('roomRate')} label="سعر الليلة" type="number" value={invoice.roomRate} onChange={v=>setInv('roomRate',v)}/><Field invalid={badFields.has('accommodation')} label="إجمالي الإقامة" type="number" value={invoice.accommodation} onChange={v=>setInv('accommodation',v)}/><Field invalid={badFields.has('restaurant')} label="المطعم / Food & Drink" type="number" value={invoice.restaurant} onChange={v=>setInv('restaurant',v)}/><Field invalid={badFields.has('laundry')} label="Laundry" type="number" value={invoice.laundry} onChange={v=>setInv('laundry',v)}/><Field label="الإجمالي المكتوب بالفاتورة" type="number" invalid={badFields.has('total')} value={invoice.total} onChange={v=>setInv('total',v)}/><Field label="التفقيط بالإنجليزي" value={invoice.amountWords} onChange={v=>setInv('amountWords',v)} placeholder="Two Thousand One Hundred EGP" wide/>
      </div></div><section className={`audit-box ${issues.length||poErrors.length?'audit-error':'audit-good'}`}><strong>{issues.length||poErrors.length?`⚠ ${issues.length+poErrors.length} تنبيه لازم تراجعه قبل التصدير`:'✓ اكتملت فحوصات الفاتورة'}</strong>{(issues.length>0||poErrors.length>0)&&<ul>{issues.map((issue,i)=><li key={`inv-${i}`}>{issue.message}</li>)}{poErrors.map((issue,i)=><li key={`po-${i}`}>{issue.message}</li>)}</ul>}</section>{folioChecks&&<div className="folio-checks"><div className={folioChecks.dateMatch?'check-ok':'check-warn'}><strong>{folioChecks.dateMatch?'✓ التواريخ مطابقة':'⚠ تم تصحيح تاريخ المغادرة'}</strong><small>المحسوب: {folioChecks.calculatedOut||'—'} · المكتوب: {folioChecks.shownFolioOut||'—'}</small></div><div className={folioChecks.rateMatch?'check-ok':'check-warn'}><strong>{folioChecks.rateMatch?'✓ قيمة الإقامة مطابقة':'⚠ قيمة الإقامة غير مطابقة'}</strong><small>المعتمد من الضريبية: {money2(invoice.accommodation).toFixed(2)} · بالفوليو: {money2(folioChecks.folioAccommodation).toFixed(2)}</small></div></div>}<div className="total-row"><span>الإجمالي المحسوب</span><strong>{enteredTotal.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2})} EGP</strong></div><div className={`match-status ${verified?'ok':'bad'}`}><span>{verified?'✓ الأرقام متطابقة':'⚠ الأرقام غير متطابقة'}</span><small>إجمالي الفاتورة: {money2(invoice.total).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2})} EGP</small></div><button className="download" disabled={!verified} onClick={()=>exportExcel(false)}><Download/>تنزيل Excel <small>{verified?`${company?.code} - ${String(invoice.invoiceNo||'رقم الفاتورة').replaceAll('/','-')}.xlsx`:'صحّح التنبيهات أولًا'}</small></button><div className="export-actions"><button onClick={newInvoice}>فاتورة جديدة</button><button disabled={!verified} onClick={printInvoice}>طباعة A4 / حفظ PDF</button><button disabled={!verified} onClick={()=>exportExcel(true)}>مشاركة Excel</button><button onClick={()=>{try{localStorage.setItem('invoice-current-draft',JSON.stringify({invoice,companyId,invoicePo}));setNotice('تم حفظ البيانات؛ الصور تحفظ تلقائيًا إذا سمحت مساحة الجهاز.');}catch{setNotice('تعذر الحفظ على الجهاز.')}}}>حفظ المسودة</button></div><small className="camera-note">السجل يحفظ البيانات فقط على المتصفح الحالي وليس صور المستندات. PDF: اختار «حفظ كـ PDF» من نافذة الطباعة. المسودة وسجل الأرقام محفوظين على هذا الجهاز فقط.</small><small className="draft-status" role="status">{draftStatus}</small>{notice&&<div className="notice">{notice}</div>}</section></div>
      <div className="print-only"><InvoiceSheet invoice={finalInvoice()} company={company} po={invoicePo}/></div>
    </>:tab==='companies'?<Companies companies={companies} setCompanies={setCompanies} setCompanyId={setCompanyId}/>:<section className="companies-page"><div className="page-head"><div><h2>سجل الفواتير</h2><p>بحث واستعادة بيانات الفواتير على الجهاز الحالي فقط. لا يتم حفظ صور المستندات.</p></div><History/></div><label className="history-search"><Search/><input value={historyQuery} onChange={e=>setHistoryQuery(e.target.value)} placeholder="ابحث بالشركة أو رقم الفاتورة أو النزيل أو التاريخ"/></label><div className="company-list">{filteredHistory.length?filteredHistory.map((item,i)=><div className="company-item history-invoice" key={item.id||`${item.code}-${item.invoiceNo}-${i}`} role="button" tabIndex={0} onClick={()=>item.invoice?setPreviewHistory(item):restoreHistory(item)} onKeyDown={e=>{if(e.key==='Enter')item.invoice?setPreviewHistory(item):restoreHistory(item);}}><div className="company-badge">{item.code?.slice(0,2)}</div><div className="company-data"><strong>{item.code} · {item.invoiceNo}</strong><span>{item.guest||'بيانات قديمة'} · {item.invoiceDate||item.date?.slice(0,10)||''}</span><small>{item.companyName||''} {item.roomNo?`· غرفة ${item.roomNo}`:''}</small></div><button className="sub-action" onClick={e=>{e.stopPropagation();item.invoice?setPreviewHistory(item):restoreHistory(item);}}><FileSpreadsheet/> عرض</button><button className="danger" onClick={e=>{e.stopPropagation();deleteHistory(item.id);}} disabled={!item.id} title="حذف"><Trash2/></button></div>):<p>لا توجد فواتير تطابق البحث.</p>}</div></section>}</main>
    {previewHistory&&<div className="preview-overlay" role="dialog" aria-modal="true" aria-label="معاينة الشيت" onClick={()=>setPreviewHistory(null)}><div className="preview-dialog" onClick={e=>e.stopPropagation()}><div className="preview-toolbar"><strong>معاينة الشيت قبل التحميل</strong><button type="button" onClick={()=>setPreviewHistory(null)}>إغلاق ✕</button><button type="button" onClick={()=>{restoreHistory(previewHistory);setPreviewHistory(null);}}>فتح للتعديل / التنزيل</button></div><InvoiceSheet invoice={previewHistory.invoice} company={companies.find(c=>c.id===previewHistory.companyId)||{name:previewHistory.companyName,code:previewHistory.code,tax:''}} po={previewHistory.po}/></div></div>}
    <div className="mobile-bottom-nav" aria-label="تنقل الموبايل"><button className={tab==='create'?'active':''} onClick={()=>setTab('create')}><ScanText/>فاتورة</button><button onClick={()=>{setTab('create');document.querySelector('.upload-panel')?.scrollIntoView({behavior:'smooth'});}}><FileImage/>تصوير</button><button className={tab==='history'?'active':''} onClick={()=>setTab('history')}><History/>السجل</button><button className={tab==='companies'?'active':''} onClick={()=>setTab('companies')}><Settings2/>الشركات</button></div>
  </div>
}

function InvoiceSheet({invoice:inv,company,po}){
  if(!inv)return <p>الفاتورة القديمة لا تحتوي على بيانات للعرض.</p>;
  const fmt=v=>money2(v).toFixed(2);
  const total=money2(money2(inv.accommodation)+money2(inv.restaurant)+money2(inv.laundry));
  return <article className="invoice-sheet" dir="ltr"><div className="sheet-meta"><div>{po&&<div className="sheet-po"><b>PO</b><b>{po}</b></div>}</div><div><div>DATE : <b>{inv.checkIn||inv.invoiceDate}</b></div><div>INV.NO : <b>{inv.invoiceNo}</b></div><div>ROOM NO : <b>{inv.roomNo}</b></div></div></div><div className="sheet-bill"><strong>BILL TO</strong><div>GUEST Name : <b>{inv.guest}</b></div><div>Company Name : <b>{company?.name}</b></div><div>Tax Registration : <b>{company?.tax}</b></div></div><table><colgroup>{[25,13,13,7,7,7,7,13,15].map((v,i)=><col key={i} style={{width:`${v/107*100}%`}}/>)}</colgroup><thead><tr>{['DESCRIPTION','C/I','C/O','NGT','DBL','TPL','SGL','RATE E.P','AMOUNT'].map(t=><th key={t}>{t}</th>)}</tr></thead><tbody><tr><td>Accommodation</td><td>{inv.checkIn}</td><td>{inv.checkOut}</td><td>{inv.nights}</td><td>{inv.double||''}</td><td>{inv.triple||''}</td><td>{inv.single||''}</td><td>{fmt(inv.roomRate||inv.accommodation)}</td><td>{fmt(inv.accommodation)}</td></tr><tr><td>Restaurant</td>{Array.from({length:7},(_,i)=><td key={i}/>)}<td>{fmt(inv.restaurant)}</td></tr><tr><td>Laundry</td>{Array.from({length:7},(_,i)=><td key={i}/>)}<td>{fmt(inv.laundry)}</td></tr>{Array.from({length:5},(_,i)=><tr key={i}><td colSpan="9">&nbsp;</td></tr>)}</tbody></table><div className="sheet-total"><strong>OTHER COMMENTS</strong><div><span>Sub Total</span><b>{fmt(total)}</b><span>TOTAL</span><b>{fmt(total)}</b></div></div><div className="sheet-words">{inv.amountWords||amountInWords(total)}</div><footer><b>Credit Manager</b><b>Make all Checks Payable To Pearl Hotel</b></footer></article>;
}

function Field({label,value,onChange,type='text',placeholder='',wide=false,invalid=false}){return <label className={`${wide?'field wide':'field'} ${invalid?'field-invalid':''}`}><span>{label}</span><input type={type} inputMode={type==='number'?'decimal':/date|تاريخ|Check in|Check out/i.test(label)?'numeric':'text'} step={type==='number'?'any':undefined} enterKeyHint="next" autoComplete="off" value={value} placeholder={placeholder} onChange={e=>onChange(e.target.value)}/></label>}
function Companies({companies,setCompanies,setCompanyId}){
  const blank={code:'',name:'',tax:'',invoiceTax:'',po:''};
  const [draft,setDraft]=useState(blank);const [editing,setEditing]=useState('');
  const save=()=>{if(!draft.code.trim()||!draft.name.trim())return;const item={...draft,code:draft.code.trim().toUpperCase(),invoiceTax:draft.invoiceTax||String(draft.tax).replace(/\D/g,''),id:editing||crypto.randomUUID()};setCompanies(p=>[...p.filter(c=>c.id!==editing),item]);setCompanyId(item.id);setDraft(blank);setEditing('');};
  return <section className="companies-page"><div className="page-head"><div><h2>إعدادات الشركات التلقائية</h2><p>كل شركة لها كود واسم ورقم ضريبي وPO خاص بها. عدّل الإعدادات عند تغيرها.</p></div><Building2/></div><div className="company-form"><Field label="الكود المختصر" value={draft.code} onChange={v=>setDraft(p=>({...p,code:v.toUpperCase()}))} placeholder="EXPRO"/><Field label="اسم الشركة" value={draft.name} onChange={v=>setDraft(p=>({...p,name:v}))}/><Field label="Tax Registration في الشيت" value={draft.tax} onChange={v=>setDraft(p=>({...p,tax:v}))}/><Field label="رقم الشركة داخل الفاتورة" value={draft.invoiceTax} onChange={v=>setDraft(p=>({...p,invoiceTax:v}))} placeholder="للتعرف التلقائي"/><Field label="PO الافتراضي (اختياري)" value={draft.po} onChange={v=>setDraft(p=>({...p,po:v}))}/><button className="primary" onClick={save}><Save/>{editing?'حفظ التعديل':'إضافة الشركة'}</button>{editing&&<button className="sub-action" onClick={()=>{setEditing('');setDraft(blank);}}>إلغاء التعديل</button>}</div><div className="company-list">{companies.map(c=><div className="company-item" key={c.id}><div className="company-badge">{c.code.slice(0,2)}</div><div className="company-data"><strong>{c.code}</strong><span>{c.name}</span><small>Tax: {c.tax||'—'} · تعريف الفاتورة: {c.invoiceTax||'—'} · PO: {c.po||'بدون PO'}</small></div><button className="sub-action" onClick={()=>{setEditing(c.id);setDraft({code:c.code,name:c.name,tax:c.tax||'',invoiceTax:c.invoiceTax||'',po:c.po||''});}}>تعديل</button><button className="danger" aria-label="حذف" onClick={()=>{if(window.confirm('حذف إعدادات هذه الشركة؟'))setCompanies(p=>p.filter(x=>x.id!==c.id));}}><Trash2/></button></div>)}</div></section>;
}

createRoot(document.getElementById('root')).render(<App/>);
