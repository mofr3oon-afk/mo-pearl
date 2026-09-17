export function parsedNumber(raw) {
  let s=String(raw||'').replace(/\s/g,'').replace(/[^\d.,]/g,'');
  if(!s)return 0;
  const lastDot=s.lastIndexOf('.'),lastComma=s.lastIndexOf(','),last=Math.max(lastDot,lastComma);
  const separators=(s.match(/[.,]/g)||[]).length;
  if(lastDot>=0&&lastComma>=0){
    s=s.slice(0,last).replace(/[.,]/g,'')+'.'+s.slice(last+1);
  }else if(separators>1){
    const tail=s.slice(last+1);
    if(tail.length>=3&&tail.length<=5)s=s.slice(0,last).replace(/[.,]/g,'')+'.'+tail;
    else s=s.replace(/[.,]/g,'');
  }else if(last>=0){
    const tail=s.slice(last+1);
    if(tail.length===3&&s.slice(0,last).length<=2)s=s.replace(/[.,]/g,'');
    else s=s.slice(0,last).replace(/[.,]/g,'')+'.'+tail;
  }
  return Number(s)||0;
}

function section(text){
  const start=text.search(/Line\s*Items/i);if(start<0)return text;
  const part=text.slice(start);const end=part.search(/Total\s*Sales\s*\(?(?:EGP|£6)/i);
  return end>0?part.slice(0,end):part;
}

function candidatesFrom(text,total){
  const body=section(text);
  return [...body.matchAll(/\b\d[\d.,]*\d\b/g)]
    .map(x=>Math.round(parsedNumber(x[0])))
    .filter(n=>n>=20&&n<=total+1);
}

function hintValues(text,total){
  const lines=text.split(/\n/),hints=[];
  for(let i=0;i<lines.length;i++){
    if(!/(Accommodation|Food\s*(?:and|&)\s*Drink|Laundry)/i.test(lines[i]))continue;
    const nearby=lines.slice(i,i+4).join(' ');
    for(const m of nearby.matchAll(/\b\d[\d.,]*\d\b/g)){
      const n=Math.round(parsedNumber(m[0]));if(n>=20&&n<=total)hints.push(n);
    }
  }
  return hints;
}

function findCombination(values,count,total,hints){
  const nums=[...new Set(values)].sort((a,b)=>a-b);
  const frequency=new Map();for(const n of values)frequency.set(n,(frequency.get(n)||0)+1);
  const hintSet=new Set(hints);let best=null,bestDiff=Infinity,bestScore=-Infinity;
  function walk(start,picked,sum){
    if(picked.length===count){const diff=Math.abs(sum-total);const score=picked.reduce((s,n)=>s+(hintSet.has(n)?100:0)+(frequency.get(n)||0),0);if(diff<bestDiff||(diff===bestDiff&&score>bestScore)){bestDiff=diff;bestScore=score;best=[...picked];}return;}
    for(let i=start;i<nums.length;i++){
      if(sum+nums[i]>total+2)break;
      picked.push(nums[i]);walk(i,picked,sum+nums[i]);picked.pop();
    }
  }
  walk(0,[],0);return bestDiff<=1?best:null;
}

function findTotal(texts){
  const found=[];
  for(const text of texts){
    for(const m of text.matchAll(/Total\s*Amount\s*\(?(?:EGP|£6)\)?[^\d]{0,12}([\d.,]+)/gi)){
      const n=Math.round(parsedNumber(m[1]));if(n>0)found.push(n);
    }
  }
  if(!found.length)return 0;
  const frequency=new Map();for(const n of found)frequency.set(n,(frequency.get(n)||0)+1);
  return [...frequency].sort((a,b)=>b[1]-a[1]||b[0]-a[0])[0][0];
}

function guestFromDescription(text){
  const block=text.match(/(?:Accommodation|ecommodation)[\s\S]*/i)?.[0]||'';
  const banned=/^(Accommodation|ecommodation|Services?|EGS|EG|C|C62|yi|yt|Vi|Wil|Ce|oo|cs|egg)$/i;
  const guests=[[]];
  for(const line of block.split(/\n+/)){
    if(/Total\s+T(?:a[xk]|ox)\s+Amount/i.test(line))break;
    const startsNewGuest=/^\s*[_—–-]{1,}\s*[A-Za-z]/.test(line);
    const safe=line.split(/\bEGS\b/i)[0].replace(/EG-?\d[\d-]*/gi,' ').replace(/[\d.,/%]+/g,' ');
    const words=(safe.match(/[A-Za-z]+/g)||[]).filter(token=>!banned.test(token)&&(token.length>=3||/^(Mr|HB)$/i.test(token)));
    if(words.length){if(startsNewGuest&&guests[0].length)guests.push([]);guests.at(-1).push(...words);}
  }
  const names=guests.map(words=>words.slice(0,5).join(' ').trim()).filter(Boolean);
  return {name:names.join(' / '),count:Math.min(3,Math.max(1,names.length))};
}
function guestFrom(text){
  const area=text.match(/(?:Accommodation|ecommodation)[\s\S]{0,420}/i)?.[0]||'';
  const banned=/(Accommodation|Services|Total|Amount|Value added|general product|Unit Price|Tax|Food|Laundry|EGS|C62)/i;
  return area.split(/\n+/).map(x=>x.replace(/[^A-Za-z ]/g,' ').replace(/\s+/g,' ').trim())
    .find(x=>x.length>=7&&x.length<=45&&/^[A-Za-z]+(?:\s+[A-Za-z]+)+$/.test(x)&&!banned.test(x))||'';
}

export function parseInvoiceOcr(texts,descriptionText=''){
  const all=texts.join('\n');
  const internal=all.match(/Internal\s*(?:ID|IO|1D|10)\s*[:;\-]?\s*([\d]+\s*\/\s*[\d]+)/i)?.[1]?.replace(/\s/g,'')||'';
  const issuance=all.match(/Issuance\s*Date\s*[:\-]?\s*(\d{1,2}[/-]\d{1,2}[/-]\d{4})/i)?.[1]||'';
  const companyTax=all.match(/Registration\s*Number\s*[:;\-]?\s*([\d\s-]{7,20})/i)?.[1]?.replace(/\D/g,'')||'';
  const accommodationDescription=descriptionText.match(/(?:Accommodation|ecommodation)[\s\S]*?(?=Total\s+T(?:a[xk]|ox)\s+Amount|$)/i)?.[0]||'';
  const nightCandidates=texts.flatMap(t=>{const block=t.match(/(?:Accommodation|ecommodation)[\s\S]{0,650}/i)?.[0]||'';return [...block.matchAll(/\b(\d{1,2})[.,]0{3,5}\s*\//g)].map(m=>Number(m[1]));});
  const nights=Math.max(1,...nightCandidates,Number(accommodationDescription.match(/\b(\d{1,2})[.,]0{3,5}\s*\//)?.[1]||1));
  const hasRestaurant=/(Food\s*(?:and|&)\s*Drink|Restaurant)/i.test(all);
  const hasLaundry=/Laundry/i.test(all);
  const count=1+Number(hasRestaurant)+Number(hasLaundry);
  const total=findTotal(texts);
  const values=texts.flatMap(t=>candidatesFrom(t,total));
  const hints=texts.flatMap(t=>hintValues(t,total));
  const combo=total?findCombination(values,count,total,hints):null;
  let accommodation='',restaurant='',laundry='';
  if(combo){
    const sorted=[...combo].sort((a,b)=>a-b);
    if(count===1)accommodation=sorted[0];
    else if(hasRestaurant&&hasLaundry){laundry=sorted[0];restaurant=sorted[1];accommodation=sorted[2];}
    else if(hasLaundry){laundry=sorted[0];accommodation=sorted[1];}
    else {restaurant=sorted[0];accommodation=sorted[1];}
  }
  const sum=Number(accommodation||0)+Number(restaurant||0)+Number(laundry||0);
  const guestInfo=guestFromDescription(descriptionText);const guestCount=guestInfo.name?guestInfo.count:1;
  return {invoiceNo:internal,invoiceDate:issuance.replaceAll('/','-'),companyTax,guest:guestInfo.name||guestFrom(texts.find(t=>/Accommodation/i.test(t))||all),single:guestCount===1?1:0,double:guestCount===2?1:0,triple:guestCount>=3?1:0,nights,roomRate:accommodation&&nights?Math.round((accommodation/nights)*100)/100:'',accommodation,restaurant,laundry,total,verified:Boolean(total&&combo&&Math.abs(sum-total)<=1)};
}

function cleanGuest(raw=''){
  let value=raw.replace(/\b(?:User|Arrival|Departure|Company)\b.*$/i,'').replace(/^\s*\d+\s*/,'').trim();
  if(value.includes(',')){const [last,first]=value.split(',').map(x=>x.trim());value=`${first} ${last}`;}
  return value.replace(/[^A-Za-z .'-]/g,' ').replace(/\s+/g,' ').trim();
}
function normalizeDate(raw=''){
  const m=raw.match(/(\d{1,2})[-\/]([01]?\d)(?:[-\/](\d{2,4}))?/);if(!m)return '';
  let year=m[3]||'';if(year.length===2)year=`20${year}`;
  return `${m[1].padStart(2,'0')}-${m[2].padStart(2,'0')}${year?`-${year}`:''}`;
}
function lastMoneyBefore(text,marker){
  const end=text.search(marker);if(end<0)return 0;
  const values=[...text.slice(0,end).matchAll(/\b\d{1,6}[.,]\d{2}\b/g)].map(m=>parsedNumber(m[0]));
  return values.at(-1)||0;
}
export function parseGuestFolioOcr(texts){
  const all=texts.join('\n');
  const checkIn=normalizeDate(all.match(/Arrival\s*[:\-]?\s*(\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4})/i)?.[1]||'');
  let actualCheckOut=normalizeDate(all.match(/checked\s*out\s*on\s*(\d{1,2}[-\/]\d{1,2}(?:[-\/]\d{2,4})?)/i)?.[1]||'');
  if(actualCheckOut&&actualCheckOut.split('-').length===2&&checkIn)actualCheckOut+=`-${checkIn.split('-')[2]}`;
  const printedCheckOut=normalizeDate(all.match(/Departure\s*[:\-]?\s*(\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4})/i)?.[1]||'');
  const roomNo=all.match(/Room\s*:\s*(\d{1,5})/i)?.[1]||'';
  const guest=cleanGuest(all.match(/Guest\s*:\s*([^\n]+)/i)?.[1]||'');
  const rooming=all.match(/Rooming\s*:\s*([^\n]+)/i)?.[1]||'';
  const single=Number(rooming.match(/(\d+)\s*Sgl/i)?.[1]||0),double=Number(rooming.match(/(\d+)\s*(?:Dbl|Double)/i)?.[1]||0),triple=Number(rooming.match(/(\d+)\s*(?:Tpl|Triple)/i)?.[1]||0);
  const accommodation=Math.round(lastMoneyBefore(all,/Main\s+(?:Restaurant|Rest\.)/i)*100)/100;
  const totalMatches=[...all.matchAll(/(?:Total\s+)?(\d{1,6}[.,]\d{2})/gi)].map(m=>parsedNumber(m[1]));
  const total=Math.max(0,...totalMatches.filter(n=>n<100000));
  return {roomNo,guest,checkIn,actualCheckOut,printedCheckOut,single,double,triple,accommodation,total,billNo:all.match(/Bill\s*No\.?\s*:\s*(\d+)/i)?.[1]||''};
}
