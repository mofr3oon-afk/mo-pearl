# مُجهّز فواتير الشركات

موقع عربي بسيط يحوّل صورة الفاتورة الضريبية إلى ملف Excel منسّق على صفحة A4 واحدة وجاهز للطباعة على ورق الفندق المطبوع.

## V13 — إضافات الموبايل

- واجهة موبايل مخصصة مع زر تصوير كبير وشريط تنقل سفلي.
- تصوير صفحات إضافية واحدة وراء الأخرى أو اختيار عدة صور وقراءتها مع الفاتورة.
- حفظ تلقائي للبيانات في Local Storage ولصور المسودة في IndexedDB (على الجهاز فقط).
- إدخال أسرع على الموبايل بلوحة أرقام مناسبة للقيم والتواريخ وأزرار لمس أكبر.
- تقريب مالي صحيح لخانتين عشريتين: `934.6199 → 934.62` ويطبق على OCR والحسابات وExcel والتفقيط.

## المميزات

- حفظ بيانات كل شركة: الكود، الاسم، Tax Registration وPO.
- اكتشاف الشركة تلقائيًا من رقم التسجيل الظاهر داخل الفاتورة.
- قراءة مجانية داخل المتصفح، بدون API أو اشتراك مدفوع.
- استخراج Accommodation وRestaurant / Food & Drink وLaundry كل بند منفصلًا.
- حساب سعر الليلة من إجمالي الإقامة وعدد الليالي، مع خانات SGL وDBL وTPL.
- مراجعة وتعديل كل البيانات قبل إنشاء الملف.
- إنشاء ملف باسم الشركة ورقم الفاتورة، مثل `EXPRO - 122529-71622.xlsx`.
- معالجة الصورة داخل المتصفح وعدم رفعها إلى قاعدة بيانات.
- متوافق مع الموبايل والكمبيوتر.

## التشغيل محليًا

```bash
npm install
npm run dev
```

## النشر على Vercel

ارفع المجلد إلى GitHub، ثم من Vercel اختر **Add New Project** وحدد المستودع. إعدادات Vite تُكتشف تلقائيًا:

- Build Command: `npm run build`
- Output Directory: `dist`

لا تحتاج إلى إضافة أي Environment Variables أو مفاتيح API.

## ملاحظة مهمة

حتى مع القراءة الذكية، الموقع يعرض شاشة مراجعة قبل تنزيل ملف Excel لأن الفاتورة مستند مالي.

## V12 — الميزات الجديدة
- إعدادات شركات قابلة للتعديل، مع اكتشاف الشركة برقم التسجيل الضريبي وتعبئة PO الافتراضي أو المستخرج من الفاتورة.
- عرض صور الفاتورة وGuest Folio وPO بجوار الخانات؛ انقر الصورة للتكبير. لا توجد بعد خاصية تحديد موضع الخانة تلقائيًا داخل الصورة.
- سجل قابل للبحث برقم الفاتورة والشركة والنزيل والتاريخ. يستعيد البيانات المحفوظة عند التصدير، **ولا يحفظ صور المستندات**. السجل مخزن محليًا في المتصفح، فلا ينتقل بين الأجهزة، وقد يُحذف عند محو بيانات الموقع.
- قراءة صورة PO (JPG/PNG) ومقارنة رقمه بالرقم المسجل، مع فحص إرشادي لاسم الشركة والإجمالي. قد يكون أمر الشراء لعدة فواتير، فلا نرفض اختلاف الإجمالي تلقائيًا. مطابقة الأصناف والأسعار التفصيلية تحتاج مراجعة يدوية. PDF كمدخل للـPO غير مدعوم.
- التصدير والطباعة يتوقفان عند تضارب رقم PO الذي تم التعرف عليه حتى تُراجع الرقم.

شغّل `npm ci && npm run build` في بيئة يتوفر فيها الإنترنت والمكتبات قبل النشر على Vercel. الاختبار الآلي لا يحل محل اختبار صور فواتير واقعية.

### V14 review and limits
- Guest Folio: add multiple images or Excel (.xlsx); Excel cell text is extracted locally. PDF files can be selected and previewed but **are not automatically extracted**; convert PDF pages to images for OCR.
- Previous invoice history now offers a spreadsheet-style preview, with a button to open for editing/export; original invoice photos are not stored in history.
- Print uses the same invoice-layout preview on A4 (browser print / Save as PDF); Web Share API is attempted on devices that support sharing .xlsx, otherwise the file downloads.
- Browser print dialogs and share targets require real-device testing: unavailable here. Previously stored drafts still only retain the first Guest Folio image.
