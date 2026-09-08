/**
 * Hamsa - Leads + Offers backend for Google Sheets
 * =================================================
 * Public:
 *   GET  ?action=offers&section=umrah|hajj|cash
 *
 * Admin:
 *   POST action=adminLogin
 *   POST action=adminListOffers
 *   POST action=adminSaveOffer
 *   POST action=adminDeleteOffer
 *
 * IMPORTANT:
 * Put the admin password in Apps Script:
 * Project Settings > Script properties
 * Property: ADMIN_PASSWORD
 *
 * Do NOT put the password inside admin.html.
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

const ADMIN_PASSWORD_PROPERTY = 'ADMIN_PASSWORD';
const ADMIN_TOKEN_PREFIX = 'HAMSA_ADMIN_';
const ADMIN_TOKEN_TTL = 21600; // 6 hours

function setup(){
  setupLeads();
  setupOffers();
}

function setupLeads(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if(!sh) sh = ss.insertSheet(SHEET_NAME);
  if(sh.getLastRow() === 0) sh.appendRow(HEADERS);
  else if(sh.getRange(1,1,1,HEADERS.length).getValues()[0].every(v=>v==='')){
    sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
  }
  sh.setFrozenRows(1);
}

function setupOffers(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(OFFERS_SHEET_NAME);
  if(!sh) sh = ss.insertSheet(OFFERS_SHEET_NAME);

  if(sh.getLastRow() === 0){
    sh.appendRow(OFFERS_HEADERS);
    sh.appendRow([
      1,'umrah','عمرة التيسير','ابدأ من 1,000 جنيه',
      '1000 جنيه × 24 شهر|1250 جنيه × 24 شهر|1500 جنيه × 24 شهر|2000 جنيه × 12 شهر',
      'ابدأ من 1,000 جنيه','', '', '', 1, 'نعم'
    ]);
    sh.appendRow([
      2,'cash','ريع بخش - اقتصادي بالمواصلات','11 ليلة مكة + 3 ليالي المدينة',
      'رباعي 40,500|ثلاثي 43,500|ثنائي 48,900|سينجل 64,500',
      '40,500 جنيه','', '', '', 1, 'نعم'
    ]);
  } else if(sh.getRange(1,1,1,OFFERS_HEADERS.length).getValues()[0].every(v=>v==='')){
    sh.getRange(1,1,1,OFFERS_HEADERS.length).setValues([OFFERS_HEADERS]);
  }
  sh.setFrozenRows(1);
}

/* =========================
   PUBLIC API
   ========================= */

function doGet(e){
  setup();
  const p = e && e.parameter ? e.parameter : {};

  if(p.action === 'offers'){
    return getOffersJson(p.section || '');
  }

  return json({
    ok:true,
    service:'hamsa-leads',
    offersApi:'?action=offers&section=umrah|hajj|cash'
  });
}

function getOffersJson(sectionFilter){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(OFFERS_SHEET_NAME);
  if(!sh) return json({ok:false,error:'لا يوجد شيت عروض بعد',offers:[]});

  const lastRow = sh.getLastRow();
  if(lastRow < 2) return json({ok:true,offers:[]});

  const values = sh.getRange(2,1,lastRow-1,OFFERS_HEADERS.length).getValues();

  let offers = values.map(function(row){
    return {
      id: String(row[0] || ''),
      section: String(row[1] || '').trim(),
      title: String(row[2] || '').trim(),
      subtitle: String(row[3] || '').trim(),
      items: String(row[4] || '')
        .split('|')
        .map(function(s){ return s.trim(); })
        .filter(Boolean),
      price: String(row[5] || '').trim(),
      priceLabel: String(row[6] || '').trim(),
      badge: String(row[7] || '').trim(),
      image: String(row[8] || '').trim(),
      order: Number(row[9]) || 999,
      active: String(row[10] || '').trim()
    };
  });

  offers = offers.filter(function(o){
    const active = o.active.toLowerCase();
    return o.title &&
      (o.active === 'نعم' ||
       active === 'yes' ||
       active === 'true' ||
       active === '1');
  });

  if(sectionFilter){
    offers = offers.filter(function(o){
      return o.section === sectionFilter;
    });
  }

  offers.sort(function(a,b){
    return a.order - b.order;
  });

  return json({ok:true,offers:offers});
}

/* =========================
   POST ROUTER
   ========================= */

function doPost(e){
  try{
    setup();
    const p = e && e.parameter ? e.parameter : {};
    const action = String(p.action || '').trim();

    // Public lead submission - keeps the original behavior.
    if(!action || action === 'lead'){
      return saveLead(p);
    }

    // Admin API.
    if(action === 'adminLogin'){
      return adminLogin(p);
    }

    if(action === 'adminListOffers'){
      requireAdmin(p);
      return adminListOffers();
    }

    if(action === 'adminSaveOffer'){
      requireAdmin(p);
      return adminSaveOffer(p);
    }

    if(action === 'adminDeleteOffer'){
      requireAdmin(p);
      return adminDeleteOffer(p);
    }

    if(action === 'adminLogout'){
      adminLogout(p);
      return json({ok:true});
    }

    return json({ok:false,error:'الإجراء غير معروف'});
  }catch(err){
    return json({
      ok:false,
      error:String(err && err.message ? err.message : err)
    });
  }
}

/* =========================
   LEADS
   ========================= */

function saveLead(p){
  // Honeypot: bots that fill this hidden field are ignored.
  if((p.website || '').trim() !== '') return json({ok:true});

  const name = String(p.name || '').trim();
  const phoneRaw = String(p.phone || '').trim();
  const phone = phoneRaw.replace(/[\s-]/g,'');
  const service = String(p.service || '').trim();

  if(!name || !phone || !service){
    return json({ok:false,error:'البيانات الأساسية ناقصة'});
  }

  if(!/^01\d{9}$/.test(phone)){
    return json({ok:false,error:'رقم الموبايل غير صحيح'});
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);

  sh.appendRow([
    new Date(),
    name,
    phone,
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

/* =========================
   ADMIN AUTH
   ========================= */

function getAdminPassword(){
  const password = PropertiesService.getScriptProperties()
    .getProperty(ADMIN_PASSWORD_PROPERTY);

  if(!password){
    throw new Error(
      'لم يتم ضبط ADMIN_PASSWORD في Script Properties'
    );
  }

  return String(password);
}

function adminLogin(p){
  const password = String(p.password || '');

  if(!password){
    return json({ok:false,error:'اكتب كلمة المرور'});
  }

  if(password !== getAdminPassword()){
    return json({ok:false,error:'كلمة المرور غير صحيحة'});
  }

  const token = Utilities.getUuid().replace(/-/g,'');
  CacheService.getScriptCache().put(
    ADMIN_TOKEN_PREFIX + token,
    '1',
    ADMIN_TOKEN_TTL
  );

  return json({
    ok:true,
    token:token,
    expiresIn:ADMIN_TOKEN_TTL
  });
}

function requireAdmin(p){
  const token = String(p.token || '').trim();

  if(!token){
    throw new Error('جلسة الإدارة غير موجودة');
  }

  const valid = CacheService.getScriptCache()
    .get(ADMIN_TOKEN_PREFIX + token);

  if(valid !== '1'){
    throw new Error('جلسة الإدارة منتهية، سجل الدخول مرة أخرى');
  }

  // Refresh the session while the admin is active.
  CacheService.getScriptCache().put(
    ADMIN_TOKEN_PREFIX + token,
    '1',
    ADMIN_TOKEN_TTL
  );
}

function adminLogout(p){
  const token = String(p.token || '').trim();
  if(token){
    CacheService.getScriptCache()
      .remove(ADMIN_TOKEN_PREFIX + token);
  }
}

/* =========================
   ADMIN OFFERS CRUD
   ========================= */

function adminListOffers(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(OFFERS_SHEET_NAME);

  if(!sh || sh.getLastRow() < 2){
    return json({ok:true,offers:[]});
  }

  const values = sh.getRange(
    2,1,sh.getLastRow()-1,OFFERS_HEADERS.length
  ).getValues();

  const offers = values.map(function(row, index){
    return {
      row:index + 2,
      id:String(row[0] || ''),
      section:String(row[1] || '').trim(),
      title:String(row[2] || '').trim(),
      subtitle:String(row[3] || '').trim(),
      items:String(row[4] || '')
        .split('|')
        .map(function(s){return s.trim();})
        .filter(Boolean),
      price:String(row[5] || '').trim(),
      priceLabel:String(row[6] || '').trim(),
      badge:String(row[7] || '').trim(),
      image:String(row[8] || '').trim(),
      order:Number(row[9]) || 999,
      active:String(row[10] || '').trim()
    };
  });

  offers.sort(function(a,b){
    return (a.order - b.order) || (a.row - b.row);
  });

  return json({ok:true,offers:offers});
}

function adminSaveOffer(p){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(OFFERS_SHEET_NAME);

  if(!sh){
    throw new Error('شيت Offers غير موجود');
  }

  const id = String(p.id || '').trim();
  const section = String(p.section || '').trim();
  const title = String(p.title || '').trim();
  const subtitle = String(p.subtitle || '').trim();
  const items = String(p.items || '').trim();
  const price = String(p.price || '').trim();
  const priceLabel = String(p.priceLabel || '').trim();
  const badge = String(p.badge || '').trim();
  const image = String(p.image || '').trim();
  const order = Number(p.order);
  const active = normalizeActive(p.active);

  if(!section || !title){
    throw new Error('القسم والعنوان مطلوبان');
  }

  if(['umrah','hajj','cash'].indexOf(section) === -1){
    throw new Error('القسم يجب أن يكون umrah أو hajj أو cash');
  }

  const safeOrder = isFinite(order) ? order : 999;

  // Existing ID => update. Empty ID => create.
  if(id){
    const rowNumber = findOfferRowById_(sh,id);

    if(rowNumber){
      sh.getRange(rowNumber,1,1,OFFERS_HEADERS.length).setValues([[
        id, section, title, subtitle, items, price,
        priceLabel, badge, image, safeOrder, active
      ]]);

      return json({
        ok:true,
        message:'تم تعديل العرض',
        id:id
      });
    }
  }

  const newId = id || nextOfferId_();

  sh.appendRow([
    newId, section, title, subtitle, items, price,
    priceLabel, badge, image, safeOrder, active
  ]);

  return json({
    ok:true,
    message:'تم إضافة العرض',
    id:String(newId)
  });
}

function adminDeleteOffer(p){
  const id = String(p.id || '').trim();

  if(!id){
    throw new Error('معرف العرض مطلوب');
  }

  const sh = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(OFFERS_SHEET_NAME);

  const rowNumber = findOfferRowById_(sh,id);

  if(!rowNumber){
    throw new Error('العرض غير موجود');
  }

  sh.deleteRow(rowNumber);

  return json({
    ok:true,
    message:'تم حذف العرض',
    id:id
  });
}

function findOfferRowById_(sh,id){
  if(!sh || sh.getLastRow() < 2) return 0;

  const ids = sh.getRange(
    2,1,sh.getLastRow()-1,1
  ).getDisplayValues();

  for(let i=0;i<ids.length;i++){
    if(String(ids[i][0]).trim() === String(id).trim()){
      return i + 2;
    }
  }

  return 0;
}

function nextOfferId_(){
  const sh = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(OFFERS_SHEET_NAME);

  if(!sh || sh.getLastRow() < 2) return '1';

  const ids = sh.getRange(
    2,1,sh.getLastRow()-1,1
  ).getValues();

  let max = 0;

  ids.forEach(function(r){
    const n = Number(r[0]);
    if(isFinite(n) && n > max) max = n;
  });

  return String(max + 1);
}

function normalizeActive(value){
  const v = String(value || '').trim().toLowerCase();

  if(
    v === 'لا' ||
    v === 'no' ||
    v === 'false' ||
    v === '0' ||
    v === 'inactive'
  ){
    return 'لا';
  }

  return 'نعم';
}

/* =========================
   JSON RESPONSE
   ========================= */

function json(obj){
  return ContentService.createTextOutput(
    JSON.stringify(obj)
  ).setMimeType(ContentService.MimeType.JSON);
}
