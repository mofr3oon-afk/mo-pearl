// Purchase orders can be partially invoiced: a different amount is advisory, never a hard error.
export const normalize = value => String(value ?? '').toUpperCase().replace(/[^A-Z0-9]/g, '');
export function extractPo(text) {
  const source=String(text||'');
  const number=source.match(/(?:PURCHASE\s*ORDER|P\s*\.?\s*O\s*\.?|ORDER\s*(?:NO|NUMBER))\s*(?:NUMBER|NO\.?|#|:|-)*\s*[:#-]?\s*([A-Z0-9][A-Z0-9/\-]{3,25})/i)?.[1]||'';
  return { number, text:source };
}
export function comparePo(poText, company, invoice, invoicePo='') {
  if(!String(poText||'').trim()) return {status:'empty',checks:[]};
  const found=extractPo(poText);
  const expected=String(invoicePo||company?.po||'').trim();
  const checks=[];
  if(found.number && expected) checks.push({kind:normalize(found.number)===normalize(expected)?'ok':'error',message:`رقم PO: المستند ${found.number} / المسجل ${expected}`});
  else checks.push({kind:'warning',message:'رقم PO مش واضح في المستند أو غير مسجل. راجعه يدويًا.'});
  const words=String(company?.name||'').toUpperCase().split(/[^A-Z0-9]+/).filter(w=>w.length>=4);
  if(words.length) checks.push({kind:words.some(w=>String(poText).toUpperCase().includes(w))?'ok':'warning',message:'مطابقة اسم الشركة في أمر الشراء (راجع الاختلاف إن وجد).'});
  const totalMatch=String(poText).match(/(?:GRAND\s*TOTAL|ORDER\s*TOTAL|TOTAL\s*AMOUNT)\s*[:=]?\s*(?:EGP|LE|USD|\$)?\s*([\d,]+(?:\.\d{1,2})?)/i);
  if(totalMatch){const orderAmount=Number(totalMatch[1].replace(/,/g,''));const invoiceAmount=Number(invoice?.total||0);checks.push({kind:'warning',message:`إجمالي PO المقروء ${orderAmount.toFixed(2)} مقابل إجمالي الفاتورة ${invoiceAmount.toFixed(2)}. الاختلاف مش بالضرورة خطأ لأن أمر الشراء ممكن يغطي أكثر من فاتورة أو عملة مختلفة.`});}
  else checks.push({kind:'warning',message:'لم نتمكن من استخراج إجمالي PO بشكل موثوق؛ راجع الأسعار والإجمالي يدويًا.'});
  checks.push({kind:'warning',message:'مطابقة الأصناف والأسعار التفصيلية تحتاج مراجعة بشرية؛ لا توجد مطابقة آلية مؤكدة لها في هذه النسخة.'});
  return {status:checks.some(c=>c.kind==='error')?'error':'review',foundNumber:found.number,checks};
}
