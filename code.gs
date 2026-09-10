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
  - Section : umrah = العمرة | hajj = الحج | cash = عمرة الكاش
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
