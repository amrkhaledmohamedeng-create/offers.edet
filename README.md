# Hamsa Offers Admin

## 1) Apps Script

1. افتح Google Sheet المستخدم حاليًا للموقع.
2. Extensions → Apps Script.
3. استبدل `Code.gs` بالكود الموجود في `code.gs`.
4. احفظ المشروع.
5. من Project Settings → Script properties أضف:
   - Property: `ADMIN_PASSWORD`
   - Value: كلمة مرور قوية خاصة بالإدارة.
6. شغّل `setup()` مرة واحدة يدويًا واسمح بالصلاحيات.

## 2) Web App

Deploy → New deployment → Web app

- Execute as: Me
- Who has access: Anyone

انسخ رابط `/exec`.

الرابط نفسه يظل مستخدمًا في نموذج الموقع لإرسال الـ Leads.

## 3) admin.html

افتح `admin.html` وابحث عن:

`const API_URL = 'PUT_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE';`

واستبدلها برابط Web App الذي نسخته، ثم ارفع `admin.html` إلى الريبو الجديد.

## 4) مهم: الموقع الرئيسي

لو الموقع الرئيسي ما زال يعرض العروض مكتوبة يدويًا داخل HTML، فلن تتغير العروض بمجرد تعديل Google Sheet.

التعديل المطلوب مرة واحدة فقط هو جعل قسم العروض يقرأ من:

`YOUR_WEB_APP_URL?action=offers&section=umrah`

وبنفس الطريقة:

`section=hajj`

`section=cash`

بعد هذا التعديل، أي إضافة/تعديل/حذف من لوحة الإدارة ينعكس على الموقع بدون تعديل HTML مرة أخرى.

## 5) الحقول

- Section: `umrah` أو `hajj` أو `cash`
- Title: اسم العرض
- Subtitle: وصف مختصر
- Items: داخل لوحة الإدارة كل نقطة في سطر مستقل
- Price: السعر الظاهر
- Badge: شارة اختيارية
- Image: رابط الصورة
- Order: ترتيب العرض
- Active: نعم/لا

## 6) الأمان

لا تضع `ADMIN_PASSWORD` داخل `admin.html`.
كلمة المرور محفوظة في Script Properties، وتسجيل الدخول يعطي Session Token مؤقتًا.
