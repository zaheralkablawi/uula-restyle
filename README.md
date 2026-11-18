# UULA Restyle

إضافة PowerPoint وWord:
- في الباوربوينت: تحويل شرائح A4 إلى ستايل مذكرات علا بضغطة زر، مع قوالب المذكرة الجاهزة، وإدراج ملف وورد مميَّز (كل جزء في مربع مستقل).
- في الوورد: تمييز أجزاء المذكرة (سؤال، حل، تعريف، ملاحظة، جدول…) بإطار ملوّن (Content Control بوسم `UULA:<type>`).

- `taskpane.html` يفتح دائمًا أحدث نسخة من `app.html`.
- `app.html` يُبنى من مجلد `src` بالأمر `python3 build.py`.
- `src/engine.js` محرك التحويل (يعمل على XML الشريحة)، و`src/templates.js` القوالب، و`src/word.js` أنواع التمييز وقراءة ملف الوورد.
- قواعد التصميم مأخوذة من مهارة `uula-restyle` (مرفقة في `docs/SKILL.md`).
- التثبيت على Mac: انسخ `UULA-Restyle.xml` إلى
  `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/`
  و`~/Library/Containers/com.microsoft.Word/Data/Documents/wef/`
- ويندوز: `windows/UULA-Restyle-Windows.zip` (يثبّت في البرنامجين).

By Zaher Alkablawi
