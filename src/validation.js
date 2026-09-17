// Pure, testable checks used both by the review screen and export actions.
export const amount = value => {
  const text=String(value??'').replace(/,/g,'').trim();
  if(!text)return 0;
  const match=text.match(/^(-?\d+)(?:\.(\d+))?$/);
  return match ? Number(match[1])+Math.sign(Number(match[1])||1)*Number((match[2]||'').slice(0,2).padEnd(2,'0'))/100 : NaN;
};
export const cents = value => Math.trunc((Number(value||0)+Number.EPSILON)*100);
export const dateValue = value => {
  const m=String(value||'').match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
  if(!m)return null;
  const date=new Date(Date.UTC(+m[3],+m[2]-1,+m[1]));
  return date.getUTCFullYear()===+m[3]&&date.getUTCMonth()===+m[2]-1&&date.getUTCDate()===+m[1]?date:null;
};
export function validateInvoice(invoice, company, history=[]){
  const issues=[];
  const add=(fields,message)=>issues.push({fields,message});
  if(!company)add(['company'],'اختار الشركة.');
  if(!String(invoice.invoiceNo||'').trim())add(['invoiceNo'],'رقم الفاتورة مطلوب.');
  if(!String(invoice.roomNo||'').trim())add(['roomNo'],'رقم الغرفة ناقص.');
  if(!String(invoice.guest||'').trim())add(['guest'],'اسم النزيل ناقص.');
  const start=dateValue(invoice.checkIn),end=dateValue(invoice.checkOut),nights=Number(invoice.nights);
  if(!start)add(['checkIn'],'تاريخ Check-in غير موجود أو غير صحيح.');
  if(!end)add(['checkOut'],'تاريخ Check-out غير موجود أو غير صحيح.');
  if(!Number.isInteger(nights)||nights<1)add(['nights'],'عدد الليالي لازم يكون عدد صحيح أكبر من صفر.');
  if(start&&end&&Number.isInteger(nights)&&nights>0&&Math.round((end-start)/86400000)!==nights)add(['checkIn','checkOut','nights'],'الفرق بين تاريخ الدخول والخروج لا يساوي عدد الليالي.');
  if(start&&invoice.invoiceDate!==invoice.checkIn)add(['invoiceDate','checkIn'],'تاريخ الفاتورة لازم يساوي Check-in.');
  const values=['accommodation','restaurant','laundry','total','roomRate'];
  for(const key of values)if(!Number.isFinite(amount(invoice[key]))||amount(invoice[key])<0)add([key],`قيمة ${key} غير صحيحة.`);
  const subtotal=cents(amount(invoice.accommodation))+cents(amount(invoice.restaurant))+cents(amount(invoice.laundry));
  if(Number.isFinite(subtotal)&&subtotal!==cents(amount(invoice.total)))add(['accommodation','restaurant','laundry','total'],'مجموع بنود الإقامة والمطعم والمغسلة لا يساوي إجمالي الفاتورة.');
  if(cents(amount(invoice.total))<=0)add(['total'],'إجمالي الفاتورة لازم يكون أكبر من صفر.');
  if(cents(amount(invoice.accommodation))>0&&cents(amount(invoice.roomRate))<=0)add(['roomRate'],'سعر الليلة ناقص.');
  if(cents(amount(invoice.roomRate))>0&&Number.isInteger(nights)&&nights>0&&cents(amount(invoice.roomRate))*nights!==cents(amount(invoice.accommodation)))add(['roomRate','nights','accommodation'],'سعر الليلة × الليالي لا يساوي إجمالي الإقامة (راجع إذا كان سعر الغرفة يشمل أكثر من وحدة).');
  if(company&&invoice.invoiceNo&&history.some(item=>item.code===company.code&&item.invoiceNo.trim().toLowerCase()===String(invoice.invoiceNo).trim().toLowerCase()))add(['invoiceNo'],'رقم الفاتورة ده اتصدّر قبل كده لنفس الشركة.');
  return issues;
}
