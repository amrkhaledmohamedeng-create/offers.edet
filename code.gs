/**
 * Hamsa - Leads + Offers + Admin Backend
 * =======================================
 *
 * نفس Google Sheet
 *
 * Sheets:
 * 1) Leads  -> طلبات العملاء
 * 2) Offers -> عروض الموقع
 *
 * Public API:
 * GET:
 *   ?action=offers
 *   ?action=offers&section=umrah
 *   ?action=offers&section=hajj
 *   ?action=offers&section=cash
 *
 * POST:
 *   طلب عميل عادي
 *
 * Admin API:
 *   action=adminLogin
 *   action=adminListOffers
 *   action=adminSaveOffer
 *   action=adminDeleteOffer
 *   action=adminLogout
 *
 * =====================================================
 */


/* =====================================================
   إعدادات Leads
   ===================================================== */

const SHEET_NAME = 'Leads';

const HEADERS = [
  'Timestamp',
  'Name',
  'Phone',
  'Governorate',
  'Service',
  'Program',
  'Persons',
  'TravelDate',
  'PaymentMethod',
  'DownPayment',
  'InstallmentMonths',
  'Notes',
  'Status'
];


/* =====================================================
   إعدادات Offers
   ===================================================== */

const OFFERS_SHEET_NAME = 'Offers';

const OFFERS_HEADERS = [
  'ID',
  'Section',
  'Title',
  'Subtitle',
  'Items',
  'Price',
  'PriceLabel',
  'Badge',
  'Image',
  'Order',
  'Active'
];


/*
  شرح شيت Offers:

  ID
  رقم مميز للعرض

  Section
  umrah = العمرة
  hajj  = الحج
  cash  = عمرة الكاش

  Title
  اسم العرض

  Subtitle
  الوصف المختصر

  Items
  كل نقطة وبينها |

  مثال:
  1000 جنيه × 24 شهر|1250 جنيه × 24 شهر|1500 جنيه × 24 شهر

  Price
  السعر الرئيسي

  PriceLabel
  وصف السعر

  Badge
  مثلا:
  الأكثر طلبًا

  Image
  رابط صورة العرض

  Order
  ترتيب العرض

  Active
  نعم = يظهر
  لا = لا يظهر
*/


/* =====================================================
   إعدادات Admin
   ===================================================== */

/*
  مهم جدًا:

  لا تكتب كلمة السر هنا.

  كلمة السر يتم تخزينها في:
  Apps Script
  > Project Settings
  > Script Properties

  Key:
  ADMIN_PASSWORD

  Value:
  كلمة السر التي تختارها
*/

const ADMIN_PASSWORD_KEY = 'ADMIN_PASSWORD';


/*
  مدة صلاحية تسجيل الدخول:

  21600 ثانية = 6 ساعات
*/

const ADMIN_TOKEN_TTL = 21600;


/* =====================================================
   Setup
   ===================================================== */

function setup() {

  setupLeads();

  setupOffers();

}


/* =====================================================
   Setup Leads
   ===================================================== */

function setupLeads() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sh = ss.getSheetByName(SHEET_NAME);

  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
  }


  if (sh.getLastRow() === 0) {

    sh.appendRow(HEADERS);

  } else {

    const firstRow = sh
      .getRange(1, 1, 1, HEADERS.length)
      .getValues()[0];

    const empty = firstRow.every(function(v) {
      return v === '';
    });

    if (empty) {

      sh
        .getRange(1, 1, 1, HEADERS.length)
        .setValues([HEADERS]);

    }

  }

  sh.setFrozenRows(1);

}


/* =====================================================
   Setup Offers
   ===================================================== */

function setupOffers() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sh = ss.getSheetByName(OFFERS_SHEET_NAME);

  if (!sh) {

    sh = ss.insertSheet(OFFERS_SHEET_NAME);

  }


  /*
    لو الشيت جديد تمامًا
  */

  if (sh.getLastRow() === 0) {

    sh.appendRow(OFFERS_HEADERS);


    /*
      عرض تجريبي 1
    */

    sh.appendRow([
      1,
      'umrah',
      'عمرة التيسير',
      'ابدأ من 1,000 جنيه',
      '1000 جنيه × 24 شهر|1250 جنيه × 24 شهر|1500 جنيه × 24 شهر|2000 جنيه × 12 شهر',
      'ابدأ من 1,000 جنيه',
      '',
      '',
      '',
      1,
      'نعم'
    ]);


    /*
      عرض تجريبي 2
    */

    sh.appendRow([
      2,
      'cash',
      'ريع بخش - اقتصادي بالمواصلات',
      '11 ليلة مكة + 3 ليالي المدينة',
      'رباعي 40,500|ثلاثي 43,500|ثنائي 48,900|سينجل 64,500',
      '40,500 جنيه',
      '',
      '',
      '',
      1,
      'نعم'
    ]);

  }


  /*
    لو الشيت موجود لكن الصف الأول فاضي
  */

  else {

    const firstRow = sh
      .getRange(1, 1, 1, OFFERS_HEADERS.length)
      .getValues()[0];

    const empty = firstRow.every(function(v) {
      return v === '';
    });

    if (empty) {

      sh
        .getRange(1, 1, 1, OFFERS_HEADERS.length)
        .setValues([OFFERS_HEADERS]);

    }

  }


  sh.setFrozenRows(1);

}


/* =====================================================
   GET
   ===================================================== */

function doGet(e) {

  try {

    setup();


    const p = e && e.parameter
      ? e.parameter
      : {};


    /*
      جلب العروض
    */

    if (p.action === 'offers') {

      return getOffersJson(
        p.section || ''
      );

    }


    /*
      اختبار الـ API
    */

    return json({

      ok: true,

      service: 'hamsa-leads',

      message: 'Hamsa API is running'

    });

  }

  catch (err) {

    return json({

      ok: false,

      error: String(err)

    });

  }

}


/* =====================================================
   GET OFFERS
   ===================================================== */

function getOffersJson(sectionFilter) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sh = ss.getSheetByName(
    OFFERS_SHEET_NAME
  );


  if (!sh) {

    return json({

      ok: false,

      error: 'لا يوجد شيت عروض بعد',

      offers: []

    });

  }


  const lastRow = sh.getLastRow();


  if (lastRow < 2) {

    return json({

      ok: true,

      offers: []

    });

  }


  const values = sh
    .getRange(
      2,
      1,
      lastRow - 1,
      OFFERS_HEADERS.length
    )
    .getValues();


  let offers = values.map(function(row) {

    return {

      id: row[0],

      section: String(
        row[1] || ''
      ).trim(),

      title: String(
        row[2] || ''
      ).trim(),

      subtitle: String(
        row[3] || ''
      ).trim(),

      items: String(
        row[4] || ''
      )
      .split('|')
      .map(function(s) {
        return s.trim();
      })
      .filter(Boolean),

      price: String(
        row[5] || ''
      ).trim(),

      priceLabel: String(
        row[6] || ''
      ).trim(),

      badge: String(
        row[7] || ''
      ).trim(),

      image: String(
        row[8] || ''
      ).trim(),

      order: Number(row[9]) || 999,

      active: String(
        row[10] || ''
      ).trim()

    };

  });


  /*
    إظهار العروض المفعلة فقط
  */

  offers = offers.filter(function(o) {

    return (

      o.title &&

      (
        o.active === 'نعم' ||

        o.active.toLowerCase() === 'yes' ||

        o.active.toLowerCase() === 'true' ||

        o.active === '1'

      )

    );

  });


  /*
    فلترة القسم
  */

  if (sectionFilter) {

    offers = offers.filter(function(o) {

      return o.section === sectionFilter;

    });

  }


  /*
    ترتيب العروض
  */

  offers.sort(function(a, b) {

    return a.order - b.order;

  });


  return json({

    ok: true,

    offers: offers

  });

}


/* =====================================================
   POST
   ===================================================== */

function doPost(e) {

  try {

    setup();


    const p = e && e.parameter
      ? e.parameter
      : {};


    /*
      =========================================
      Admin Login
      =========================================
    */

    if (p.action === 'adminLogin') {

      return adminLogin(
        p.password || ''
      );

    }


    /*
      =========================================
      Admin List Offers
      =========================================
    */

    if (p.action === 'adminListOffers') {

      requireAdmin(
        p.token || ''
      );

      return adminListOffers();

    }


    /*
      =========================================
      Admin Save Offer
      =========================================
    */

    if (p.action === 'adminSaveOffer') {

      requireAdmin(
        p.token || ''
      );

      return adminSaveOffer(p);

    }


    /*
      =========================================
      Admin Delete Offer
      =========================================
    */

    if (p.action === 'adminDeleteOffer') {

      requireAdmin(
        p.token || ''
      );

      return adminDeleteOffer(
        p.id
      );

    }


    /*
      =========================================
      Admin Logout
      =========================================
    */

    if (p.action === 'adminLogout') {

      return adminLogout(
        p.token || ''
      );

    }


    /*
      =========================================
      طلب عميل عادي
      =========================================
    */


    /*
      Honeypot
      منع بعض الـ bots
    */

    if (
      (p.website || '').trim() !== ''
    ) {

      return json({
        ok: true
      });

    }


    const name = String(
      p.name || ''
    ).trim();

    const phone = String(
      p.phone || ''
    ).trim();

    const service = String(
      p.service || ''
    ).trim();


    /*
      التحقق من البيانات الأساسية
    */

    if (
      !name ||
      !phone ||
      !service
    ) {

      return json({

        ok: false,

        error: 'البيانات الأساسية ناقصة'

      });

    }


    /*
      تنظيف رقم الهاتف
    */

    const cleanPhone = phone.replace(
      /[\s-]/g,
      ''
    );


    /*
      التحقق من رقم الهاتف المصري
    */

    if (
      !/^01\d{9}$/.test(cleanPhone)
    ) {

      return json({

        ok: false,

        error: 'رقم الموبايل غير صحيح'

      });

    }


    const ss =
      SpreadsheetApp.getActiveSpreadsheet();


    const sh =
      ss.getSheetByName(
        SHEET_NAME
      );


    /*
      حفظ الطلب
    */

    sh.appendRow([

      new Date(),

      name,

      cleanPhone,

      String(
        p.governorate || ''
      ).trim(),

      service,

      String(
        p.program || ''
      ).trim(),

      String(
        p.persons || '1'
      ).trim(),

      String(
        p.travelDate || ''
      ).trim(),

      String(
        p.paymentMethod || ''
      ).trim(),

      String(
        p.downPayment || ''
      ).trim(),

      String(
        p.installmentMonths || ''
      ).trim(),

      String(
        p.notes || ''
      ).trim(),

      'جديد'

    ]);


    return json({

      ok: true,

      message: 'تم تسجيل الطلب'

    });

  }

  catch (err) {

    return json({

      ok: false,

      error: String(err)

    });

  }

}


/* =====================================================
   ADMIN LOGIN
   ===================================================== */

function adminLogin(password) {

  const savedPassword =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        ADMIN_PASSWORD_KEY
      );


  /*
    لو كلمة السر غير موجودة
  */

  if (!savedPassword) {

    return json({

      ok: false,

      error:
        'لم يتم إعداد كلمة مرور المدير. افتح Project Settings ثم Script Properties وأضف ADMIN_PASSWORD'

    });

  }


  /*
    التحقق من كلمة السر
  */

  if (
    String(password || '') !==
    String(savedPassword)
  ) {

    return json({

      ok: false,

      error: 'كلمة المرور غير صحيحة'

    });

  }


  /*
    إنشاء Token
  */

  const token =
    Utilities.getUuid();


  /*
    حفظ الـ Token في Cache
  */

  CacheService
    .getScriptCache()
    .put(
      'ADMIN_' + token,
      '1',
      ADMIN_TOKEN_TTL
    );


  return json({

    ok: true,

    token: token,

    expiresIn: ADMIN_TOKEN_TTL

  });

}


/* =====================================================
   CHECK ADMIN TOKEN
   ===================================================== */

function requireAdmin(token) {

  if (!token) {

    throw new Error(
      'غير مصرح'
    );

  }


  const cache =
    CacheService.getScriptCache();


  const key =
    'ADMIN_' + token;


  const valid =
    cache.get(key);


  if (valid !== '1') {

    throw new Error(
      'انتهت جلسة المدير، سجل الدخول مرة أخرى'
    );

  }


  /*
    تجديد الجلسة
  */

  cache.put(
    key,
    '1',
    ADMIN_TOKEN_TTL
  );

}


/* =====================================================
   ADMIN LIST OFFERS
   ===================================================== */

function adminListOffers() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const sh =
    ss.getSheetByName(
      OFFERS_SHEET_NAME
    );


  if (!sh) {

    return json({

      ok: false,

      error: 'شيت Offers غير موجود',

      offers: []

    });

  }


  const lastRow =
    sh.getLastRow();


  if (lastRow < 2) {

    return json({

      ok: true,

      offers: []

    });

  }


  const values =
    sh.getRange(
      2,
      1,
      lastRow - 1,
      OFFERS_HEADERS.length
    ).getValues();


  const offers =
    values.map(function(row) {

      return {

        id: row[0],

        section: String(
          row[1] || ''
        ).trim(),

        title: String(
          row[2] || ''
        ).trim(),

        subtitle: String(
          row[3] || ''
        ).trim(),

        items: String(
          row[4] || ''
        )
        .split('|')
        .map(function(item) {
          return item.trim();
        })
        .filter(Boolean),

        price: String(
          row[5] || ''
        ).trim(),

        priceLabel: String(
          row[6] || ''
        ).trim(),

        badge: String(
          row[7] || ''
        ).trim(),

        image: String(
          row[8] || ''
        ).trim(),

        order: Number(
          row[9]
        ) || 999,

        active:
          normalizeActive(
            row[10]
          )

      };

    });


  offers.sort(function(a, b) {

    return a.order - b.order;

  });


  return json({

    ok: true,

    offers: offers

  });

}


/* =====================================================
   ADMIN SAVE OFFER
   ===================================================== */

function adminSaveOffer(p) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const sh =
    ss.getSheetByName(
      OFFERS_SHEET_NAME
    );


  if (!sh) {

    throw new Error(
      'شيت Offers غير موجود'
    );

  }


  /*
    البيانات
  */

  const id =
    String(p.id || '').trim();


  const section =
    String(p.section || '')
      .trim()
      .toLowerCase();


  const title =
    String(p.title || '')
      .trim();


  const subtitle =
    String(p.subtitle || '')
      .trim();


  const items =
    String(p.items || '')
      .trim()
      .split(/\r?\n/)
      .map(function(item) {
        return item.trim();
      })
      .filter(Boolean)
      .join('|');


  const price =
    String(p.price || '')
      .trim();


  const priceLabel =
    String(p.priceLabel || '')
      .trim();


  const badge =
    String(p.badge || '')
      .trim();


  const image =
    String(p.image || '')
      .trim();


  const order =
    Number(p.order) || 999;


  const active =
    normalizeActive(
      p.active
    );


  /*
    التحقق من القسم
  */

  const allowedSections = [
    'umrah',
    'hajj',
    'cash'
  ];


  if (
    allowedSections.indexOf(section) === -1
  ) {

    throw new Error(
      'القسم غير صحيح. استخدم umrah أو hajj أو cash'
    );

  }


  /*
    العنوان مطلوب
  */

  if (!title) {

    throw new Error(
      'اسم العرض مطلوب'
    );

  }


  /*
    البحث عن العرض لو ID موجود
  */

  const lastRow =
    sh.getLastRow();


  let targetRow = -1;


  if (
    id &&
    lastRow >= 2
  ) {

    const ids =
      sh
        .getRange(
          2,
          1,
          lastRow - 1,
          1
        )
        .getValues();


    for (
      let i = 0;
      i < ids.length;
      i++
    ) {

      if (
        String(ids[i][0]) === id
      ) {

        targetRow = i + 2;

        break;

      }

    }

  }


  /*
    لو العرض موجود:
    تعديل
  */

  if (targetRow !== -1) {

    const currentId =
      sh
        .getRange(
          targetRow,
          1
        )
        .getValue();


    sh
      .getRange(
        targetRow,
        1,
        1,
        OFFERS_HEADERS.length
      )
      .setValues([[
        currentId,
        section,
        title,
        subtitle,
        items,
        price,
        priceLabel,
        badge,
        image,
        order,
        active
      ]]);


    return json({

      ok: true,

      message: 'تم تعديل العرض',

      id: currentId

    });

  }


  /*
    لو عرض جديد:
    إنشاء ID جديد
  */

  const newId =
    getNextOfferId();


  sh.appendRow([

    newId,

    section,

    title,

    subtitle,

    items,

    price,

    priceLabel,

    badge,

    image,

    order,

    active

  ]);


  return json({

    ok: true,

    message: 'تم إضافة العرض',

    id: newId

  });

}


/* =====================================================
   GET NEXT OFFER ID
   ===================================================== */

function getNextOfferId() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const sh =
    ss.getSheetByName(
      OFFERS_SHEET_NAME
    );


  const lastRow =
    sh.getLastRow();


  if (lastRow < 2) {

    return 1;

  }


  const ids =
    sh
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getValues();


  let maxId = 0;


  ids.forEach(function(row) {

    const number =
      Number(row[0]);


    if (
      !isNaN(number) &&
      number > maxId
    ) {

      maxId = number;

    }

  });


  return maxId + 1;

}


/* =====================================================
   ADMIN DELETE OFFER
   ===================================================== */

function adminDeleteOffer(id) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const sh =
    ss.getSheetByName(
      OFFERS_SHEET_NAME
    );


  if (!sh) {

    throw new Error(
      'شيت Offers غير موجود'
    );

  }


  const cleanId =
    String(id || '').trim();


  if (!cleanId) {

    throw new Error(
      'ID العرض غير موجود'
    );

  }


  const lastRow =
    sh.getLastRow();


  if (lastRow < 2) {

    return json({

      ok: false,

      error: 'لا توجد عروض'

    });

  }


  const ids =
    sh
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getValues();


  for (
    let i = 0;
    i < ids.length;
    i++
  ) {

    if (
      String(ids[i][0]) === cleanId
    ) {

      sh.deleteRow(i + 2);


      return json({

        ok: true,

        message: 'تم حذف العرض',

        id: cleanId

      });

    }

  }


  return json({

    ok: false,

    error: 'العرض غير موجود'

  });

}


/* =====================================================
   ADMIN LOGOUT
   ===================================================== */

function adminLogout(token) {

  if (token) {

    CacheService
      .getScriptCache()
      .remove(
        'ADMIN_' + token
      );

  }


  return json({

    ok: true,

    message: 'تم تسجيل الخروج'

  });

}


/* =====================================================
   ACTIVE VALUE
   ===================================================== */

function normalizeActive(value) {

  const v =
    String(value || '')
      .trim()
      .toLowerCase();


  if (
    v === 'نعم' ||
    v === 'yes' ||
    v === 'true' ||
    v === '1'
  ) {

    return 'نعم';

  }


  return 'لا';

}


/* =====================================================
   JSON RESPONSE
   ===================================================== */

function json(obj) {

  return ContentService
    .createTextOutput(
      JSON.stringify(obj)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );

}
