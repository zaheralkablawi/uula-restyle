# UULA Restyle

إضافة PowerPoint: تحويل شرائح A4 إلى ستايل مذكرات علا بضغطة زر، مع قوالب المذكرة الجاهزة (نص، سؤال، جواب، جدول، ملاحظة…).

- `taskpane.html` يفتح دائمًا أحدث نسخة من `app.html`.
- `app.html` يُبنى من مجلد `src` بالأمر `python3 build.py`.
- `src/engine.js` محرك التحويل (يعمل على XML الشريحة)، و`src/templates.js` القوالب.
- قواعد التصميم مأخوذة من مهارة `uula-restyle` (مرفقة في `docs/SKILL.md`).
- التثبيت على Mac: انسخ `manifest.xml` إلى
  `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/`

By Zaher Alkablawi
