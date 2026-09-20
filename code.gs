/**
 * Hamsa - Leads + Offers + Admin backend for Google Sheets
 * ==========================================================
 * ازاي تستخدمه:
 * 1) افتح Extensions > Apps Script في نفس شيت جوجل بتاعك.
 * 2) امسح الكود القديم كله والصق هذا الكود بدلاً منه.
 * 3) غيّر قيمة ADMIN_KEY تحت لكلمة سر سرية من عندك (حروف وأرقام).
 * 4) شغّل الدالة setup() مرة واحدة يدويًا من القائمة أعلى المحرر.
 * 5) اعمل Deploy > New deployment > Web app
 *    (Execute as: Me, Who has access: Anyone)
 *    وخد رابط /exec ده، هو نفسه اللي هتحطه في index.html وفي admin.html.
 * 6) لنقل العروض الحالية للشيت مرة واحدة: شغّل seedExistingOffers() (موجودة في آخر الملف).
 *
 * ملحوظة أمان: هذا المفتاح بسيط ومناسب لموقع صغير، مش نظام حماية بنكي.
 * لو حسّيت إن حد عرف المفتاح، غيّره فورًا من هنا واعمل Deploy جديد.
 */

const SHEET_NAME = 'Leads';
const HEADERS = [
  'Timestamp','Name','Phone','Governorate','Service','Program',
  'Persons','TravelDate','PaymentMethod','DownPayment','InstallmentMonths',
  'Notes','Status'
];

const OFFERS_SHEET_NAME = 'Offers';
const OFFERS_HEADERS = [
  'ID','Section','Title','Subtitle','Items','Price','PriceLabel',
  'Badge','Image','Order','Active'
];

/* غيّر هذا المفتاح لكلمة سر خاصة بيك قبل ما تنشر الموقع */
const ADMIN_KEY = 'CHANGE_ME_SECRET_KEY';

/*
  شرح أعمدة شيت Offers:
  - Section : umrah = العمرة | hajj = الحج | cash = كاش اقتصادي | cash5 = كاش 5 نجوم
  - Items   : كل نقطة اتفصلت عن التانية بعلامة | (شرطة رأسية)
  - Order   : رقم ترتيب الظهور (الأصغر يظهر أولاً)
  - Active  : "نعم" يعني يظهر، "لا" يعني مخفي
*/

function setup(){
  setupLeads();
  setupOffers();
}

function setupLeads(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if(!sh) sh = ss.insertSheet(SHEET_NAME);
  if(sh.getLastRow() === 0) sh.appendRow(HEADERS);
  else if(sh.getRange(1,1,1,HEADERS.length).getValues()[0].every(v=>v==='')) sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
  sh.setFrozenRows(1);
}

function setupOffers(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(OFFERS_SHEET_NAME);
  if(!sh) sh = ss.insertSheet(OFFERS_SHEET_NAME);
  if(sh.getLastRow() === 0){
    sh.appendRow(OFFERS_HEADERS);
    sh.appendRow([String(Date.now()),'umrah','عمرة التيسير','ابدأ من 1,000 جنيه',
      '1000 جنيه × 24 شهر|1250 جنيه × 24 شهر|1500 جنيه × 24 شهر|2000 جنيه × 12 شهر',
      'ابدأ من 1,000 جنيه','', '', '', 1, 'نعم']);
    sh.appendRow([String(Date.now()+1),'cash','ريع بخش - اقتصادي بالمواصلات','11 ليلة مكة + 3 ليالي المدينة',
      'رباعي 40,500|ثلاثي 43,500|ثنائي 48,900|سينجل 64,500',
      '40,500 جنيه','', '', '', 1, 'نعم']);
  } else if(sh.getRange(1,1,1,OFFERS_HEADERS.length).getValues()[0].every(v=>v==='')){
    sh.getRange(1,1,1,OFFERS_HEADERS.length).setValues([OFFERS_HEADERS]);
  }
  sh.setFrozenRows(1);
}

/* ================== GET: عرض العروض ================== */

function doGet(e){
  setup();
  const p = e && e.parameter ? e.parameter : {};

  if(p.action === 'offers'){
    return getOffersJson(p.section || '', false, '');
  }
  if(p.action === 'offersAdmin'){
    return getOffersJson('', true, p.key || '');
  }
  return json({ok:true,service:'hamsa-leads'});
}

function getOffersJson(sectionFilter, adminMode, key){
  if(adminMode && key !== ADMIN_KEY){
    return json({ok:false,error:'مفتاح الدخول غير صحيح'});
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(OFFERS_SHEET_NAME);
  if(!sh) return json({ok:false,error:'لا يوجد شيت عروض بعد',offers:[]});

  const lastRow = sh.getLastRow();
  if(lastRow < 2) return json({ok:true,offers:[]});

  const values = sh.getRange(2,1,lastRow-1,OFFERS_HEADERS.length).getValues();
  let offers = values.map(function(row){
    return {
      id: String(row[0]),
      section: String(row[1]||'').trim(),
      title: String(row[2]||'').trim(),
      subtitle: String(row[3]||'').trim(),
      items: String(row[4]||'').split('|').map(function(s){return s.trim();}).filter(Boolean),
      itemsRaw: String(row[4]||''),
      price: String(row[5]||'').trim(),
      priceLabel: String(row[6]||'').trim(),
      badge: String(row[7]||'').trim(),
      image: String(row[8]||'').trim(),
      order: Number(row[9]) || 999,
      active: String(row[10]||'').trim()
    };
  });

  if(!adminMode){
    offers = offers.filter(function(o){
      return o.title && (o.active === 'نعم' || o.active.toLowerCase() === 'yes' || o.active === 'true' || o.active === '1');
    });
    if(sectionFilter){
      offers = offers.filter(function(o){ return o.section === sectionFilter; });
    }
  }

  offers.sort(function(a,b){ return a.order - b.order; });
  return json({ok:true,offers:offers});
}

/* ================== POST: طلبات الحجز + إدارة العروض ================== */

function doPost(e){
  try{
    setup();
    const p = e && e.parameter ? e.parameter : {};
    const action = String(p.action || 'lead').trim();

    if(action === 'lead')       return handleLead(p);
    if(action === 'addOffer')   return requireAdmin(p, addOffer);
    if(action === 'updateOffer')return requireAdmin(p, updateOffer);
    if(action === 'deleteOffer')return requireAdmin(p, deleteOffer);

    return json({ok:false,error:'إجراء غير معروف'});
  }catch(err){
    return json({ok:false,error:String(err)});
  }
}

function requireAdmin(p, fn){
  if(String(p.key||'') !== ADMIN_KEY){
    return json({ok:false,error:'مفتاح الدخول غير صحيح'});
  }
  return fn(p);
}

function handleLead(p){
  // Honeypot: بوتات تملأ هذا الحقل المخفي بيتم تجاهلها
  if((p.website || '').trim() !== '') return json({ok:true});
  const name = String(p.name || '').trim();
  const phone = String(p.phone || '').trim();
  const service = String(p.service || '').trim();
  if(!name || !phone || !service) return json({ok:false,error:'البيانات الأساسية ناقصة'});
  if(!/^01\d{9}$/.test(phone.replace(/[\s-]/g,''))) return json({ok:false,error:'رقم الموبايل غير صحيح'});
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  sh.appendRow([
    new Date(), name, phone,
    String(p.governorate || '').trim(),
    service,
    String(p.program || '').trim(),
    String(p.persons || '1').trim(),
    String(p.travelDate || '').trim(),
    String(p.paymentMethod || '').trim(),
    String(p.downPayment || '').trim(),
    String(p.installmentMonths || '').trim(),
    String(p.notes || '').trim(),
    'جديد'
  ]);
  return json({ok:true,message:'تم تسجيل الطلب'});
}

function addOffer(p){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(OFFERS_SHEET_NAME);
  const id = String(Date.now());
  sh.appendRow([
    id,
    String(p.section||'').trim(),
    String(p.title||'').trim(),
    String(p.subtitle||'').trim(),
    String(p.items||'').trim(),
    String(p.price||'').trim(),
    String(p.priceLabel||'').trim(),
    String(p.badge||'').trim(),
    String(p.image||'').trim(),
    Number(p.order)||999,
    String(p.active||'نعم').trim()
  ]);
  return json({ok:true,id:id});
}

function updateOffer(p){
  const id = String(p.id||'').trim();
  if(!id) return json({ok:false,error:'رقم العرض مفقود'});
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(OFFERS_SHEET_NAME);
  const lastRow = sh.getLastRow();
  if(lastRow < 2) return json({ok:false,error:'العرض غير موجود'});
  const ids = sh.getRange(2,1,lastRow-1,1).getValues();
  for(let i=0;i<ids.length;i++){
    if(String(ids[i][0]) === id){
      const row = i+2;
      sh.getRange(row,2,1,10).setValues([[
        String(p.section||'').trim(),
        String(p.title||'').trim(),
        String(p.subtitle||'').trim(),
        String(p.items||'').trim(),
        String(p.price||'').trim(),
        String(p.priceLabel||'').trim(),
        String(p.badge||'').trim(),
        String(p.image||'').trim(),
        Number(p.order)||999,
        String(p.active||'نعم').trim()
      ]]);
      return json({ok:true});
    }
  }
  return json({ok:false,error:'العرض غير موجود'});
}

function deleteOffer(p){
  const id = String(p.id||'').trim();
  if(!id) return json({ok:false,error:'رقم العرض مفقود'});
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(OFFERS_SHEET_NAME);
  const lastRow = sh.getLastRow();
  if(lastRow < 2) return json({ok:false,error:'العرض غير موجود'});
  const ids = sh.getRange(2,1,lastRow-1,1).getValues();
  for(let i=0;i<ids.length;i++){
    if(String(ids[i][0]) === id){
      sh.deleteRow(i+2);
      return json({ok:true});
    }
  }
  return json({ok:false,error:'العرض غير موجود'});
}

function json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ================== نقل العروض الحالية للشيت (شغّلها مرة واحدة فقط) ==================
 * تحذير: الدالة دي بتمسح كل الصفوف الموجودة في شيت Offers وتكتب العروض الحالية من جديد.
 */
function seedExistingOffers(){
  setupOffers();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OFFERS_SHEET_NAME);
  const last = sh.getLastRow();
  if(last > 1) sh.getRange(2,1,last-1,OFFERS_HEADERS.length).clearContent();

  const RIYAL = 'الطيران: سعودي — القاهرة / المدينة / جدة / القاهرة';

  // [القسم, الاسم, نص الزرار/سطر إضافي, [النقاط], السعر, لون الكارت, الصورة]
  const offers = [
    // ---------- العمرة ----------
    ['umrah','عمرة التيسير','',
      ['1,000 أو 1,250 أو 1,500 جنيه شهريًا','على مدار 24 شهر','2,000 أو 2,500 أو 3,000 جنيه شهريًا','على مدار 12 شهر'],
      'ابدأ من 1,000 جنيه','','kaaba.jpg'],
    ['umrah','تيسير رمضان','',
      ['3,000 جنيه على 16 شهر','بدأ من شهر 6','2,500 جنيه على 20 شهر','بدأ من شهر 5'],
      'حجز مرن','ramadan','ramadan.jpg'],
    ['umrah','العمرة بالقسط','',
      ['اختار البرنامج','ادفع 50% أو 60%','قسط الباقي من 3 إلى 12 شهر','بدون تعقيد'],
      '3 - 12 شهر','purple','installment.jpg'],
    ['umrah','عمرة التيسير البلاتينوم','',
      ['5,000 جنيه','على 12 شهر','برنامج مميز'],
      '5,000 × 12 شهر','gold-card','hotel.jpg'],
    ['umrah','عمرة التيسير الذهبية','',
      ['4,000 جنيه','على 12 شهر','خيار مميز للتيسير'],
      '4,000 × 12 شهر','gold-card','medina.jpg'],

    // ---------- الحج ----------
    ['hajj','الحج السياحي','استفسر عن التفاصيل',
      ['برامج متنوعة ومميزة','خدمات مميزة للحجاج','إشراف ومتابعة'],'','',''],
    ['hajj','الحج الميسر','استفسر عن التفاصيل',
      ['برنامج حج ميسر','خدمات متكاملة','أسعار وخيارات متعددة'],'','',''],
    ['hajj','الحج الميسر بالتقسيط','اطلب التفاصيل',
      ['حتى 120,000 جنيه','تقسيط حتى 24 شهر','خطة سداد مرنة'],'','',''],

    // ---------- كاش اقتصادي ----------
    ['cash','ريع بخش — اقتصادي بالمواصلات','',
      ['المدينة: نسك المدينة / درة الإيمان / طيبة هيلز — 3 ليالي','مكة: قصر العليان / النخبة 1 / أبراج القصواء — 11 ليلة'],
      'رباعي : 40,500|ثلاثي : 43,500|ثنائي : 48,900|سينجل : 64,500|طفل : 31,500|رضيع : 17,500','',''],
    ['cash','اقتصادي مشي — بير بليلة','',
      ['المدينة: نسك المدينة / درة الإيمان / طيبة هيلز — 3 ليالي','مكة: إعمار أفاق — 11 ليلة'],
      'رباعي : 42,000|ثلاثي : 46,500|ثنائي : 52,900|سينجل : 72,900|طفل : 31,500|رضيع : 17,500','',''],
    ['cash','اقتصادي مشي — أجياد السد','',
      ['المدينة: نسك المدينة / درة الإيمان / طيبة هيلز — 3 ليالي','مكة: واحة الضيافة / العليان أجياد — 11 ليلة'],
      'رباعي : 42,500|ثلاثي : 46,900|ثنائي : 53,500|سينجل : 73,900|طفل : 31,500|رضيع : 17,500','',''],

    // ---------- كاش 5 نجوم ----------
    ['cash5','5 نجوم أ — 3 ليالي المدينة + 4 ليالي مكة','',
      [RIYAL,'المدينة: فندق الحرم — 3 ليالي بالأفطار','مكة: الصفوه البرج الثالث — 4 ليالي بالأفطار، أول مطل على الحرم'],
      'رباعي : 66,900|ثلاثي : 70,900|ثنائي : 79,900','',''],
    ['cash5','5 نجوم ب — 3 ليالي المدينة + 4 ليالي مكة','',
      [RIYAL,'المدينة: دره الايمان — 3 ليالي بالأفطار','مكة: الشهداء — 4 ليالي بالأفطار، أول مطل على الحرم'],
      'رباعي : 55,500|ثلاثي : 58,900|ثنائي : 64,500','',''],
    ['cash5','اقتصادي مميز — 4 + 5 ليالي','',
      [RIYAL,'المدينة: كونكورد دار الخير — 4 ليالي','مكة: الماسة جراند — 5 ليالي، 400 متر من الحرم'],
      'رباعي : 47,000|ثلاثي : 50,900|ثنائي : 57,900','','']
  ];

  const counters = {};
  const rows = offers.map(function(o,i){
    counters[o[0]] = (counters[o[0]] || 0) + 1;
    return [String(Date.now()+i), o[0], o[1], o[2], o[3].join('|'), o[4], '', o[5], o[6], counters[o[0]], 'نعم'];
  });
  sh.getRange(2,1,rows.length,OFFERS_HEADERS.length).setValues(rows);
}
