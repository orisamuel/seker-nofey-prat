# סקר התושבים נופי פרת

סקר תושבים שנתי ליישוב נופי פרת, במקום Google Forms. מובייל פירסט, RTL, אנונימי (שם רשות), כל נושא נשמר בנפרד, הגרלה למסיימים, דשבורד לצוות ודוח ציבורי מצרפי.

## איפה הוא חי
- אתר: **https://sekernofey.online** (GitHub Pages מהענף `main`; הכתובת הישנה orisamuel.github.io/seker-nofey-prat מפנה אליו)
- ריפו: https://github.com/orisamuel/seker-nofey-prat (ציבורי)
- דומיין: נרכש ב-Cloudflare ב-04/10/2026. ה-DNS שם: 4 רשומות A ו-4 AAAA לשרתי GitHub Pages, ו-CNAME ל-www, כולן DNS only (ענן אפור). המפתח: `CLOUDFLARE_API_TOKEN` (zone id `2ba142fd76358f5b1bbb088d38a5547e`).
- שרת: Apps Script בחשבון ori@42creative.co.il, צמוד לגיליון.
  - גיליון: https://docs.google.com/spreadsheets/d/1Ya4TerGGhxIcl2K5ziu6xjDssB9bAn2J2oz9Ka6xm9c/edit
  - סקריפט: https://script.google.com/d/15D01lxC92irPoJzy5Kl56qorQqtcKBK_f0nce0Si_RruRLxLs5Oht_gI/edit
  - Web app (ב-`config.js` וב-`deploy.sh`): deployment `AKfycbyBz0e1WFXrhBSw_slxQ58ElikngdGcP9BgrC1dXRmCumXiQ1NkOVk3eGvSo03egJowfA`

## הקבצים
| קובץ | תפקיד |
|---|---|
| `index.html` | הסקר עצמו |
| `survey-data.js` | השאלון (מקור לזריעת הגיליון + תצוגה מיידית עד שהשרת עונה) |
| `appscript.gs` | השרת. נדחף עם clasp |
| `styles.css` | מערכת העיצוב (v3: "פוסטר" רגוע בצבעי מעיין, קרנטינה + רוביק) |
| `icons.js` | אייקוני Lucide מוטמעים (בלי CDN). שם לא מוכר מוצג כטקסט |
| `dashboard.html` | דשבורד צוות (סיסמה מטאב "הגדרות") |
| `report.html` | דוח ציבורי, נפתח רק כש-`publicReport` = כן |
| `setup.html` | זריעת השאלות מ-`survey-data.js` לגיליון |
| `og.png` | תמונת השיתוף בוואטסאפ (1200x630) |
| `tests/` | בדיקות: `node tests/validate-survey.js`, `node tests/backend.test.js` (825 בדיקות מול גיליון מדומה) |
| `קובץ שאלות 270926.xlsx` | המקור מהוועדות. לא בגיט (ריפו ציבורי + הערות פנימיות) |

## השאלון
מקור: `קובץ שאלות 270926.xlsx`. **גיליון2 הוא השאלון הסופי** (עמודה לכל נושא: ראשי, משני, שאלה N, תשובות אפשריות N). גיליון1 הוא טיוטת העבודה עם הערות ראשי הוועדות, ומשמש רק להבהרת סוגי שאלות. איך פורש כל דבר: `DECISIONS.md`, סעיף "שאלון 2026".
- כשמחליפים שאלון: מעלים את `meta.version` ב-`survey-data.js`. דפדפן עם טיוטות מגרסה קודמת מתחיל נקי.
- אחרי הפריסה, עריכה שוטפת נעשית בטאב "שאלות" בגיליון. `setup.html` דורס את הטאב, רק לזריעה מחדש מכוונת.

## פריסה (clasp, בחשבון ori@42creative.co.il)
כבר נעשה ב-04/10/2026 (הצעדים למטה לתיעוד, או להקמה מחדש). מתיקיית הפרויקט:
1. `npx @google/clasp@2.4.2 create --type sheets --title "סקר תושבים נופי פרת" --rootDir .`
2. להעתיק מ-`~/.claude/skills/sheets-platform/templates/` את `appsscript.json`, `.claspignore`, `deploy.sh` (אחרי create, כי create כותב manifest בלי webapp). `.clasp.json` לא נכנס לגיט.
3. `SHEET_ID` ב-`appscript.gs` = ה-ID של הגיליון שנוצר.
4. `npx @google/clasp@2.4.2 push -f` ואז `deploy -d "v1"`. כתובת ה-`/exec` נכנסת ל-`config.js` (`SCRIPT_URL`, ו-`SHEETS_URL`), ה-deployment ID נכנס ל-`deploy.sh`.
5. **צעד ידני אחד של אורי:** `clasp open`, להריץ את `setup()` ולאשר הרשאות. setup יוצר את הטאבים, סיסמת צוות אקראית (בלוג ובטאב "הגדרות") וטריגר חימום.
6. לפתוח `setup.html` ולסנכרן (בפעם הראשונה בלי סיסמה).
7. לבדוק round-trip: `?action=ping` מחזיר `v2`, ומילוי נושא אחד מופיע בטאב "תוצאות".

כל שינוי בשרת אחר כך: `./deploy.sh "מה השתנה"` (גרסה חדשה על אותה כתובת). `deploy` בלי `-i` יוצר כתובת חדשה, לא לעשות.

## הדומיין
מחובר מ-04/10/2026. אם צריך לשחזר: הרשומות למעלה ב-Cloudflare, ואז `gh api -X PUT repos/orisamuel/seker-nofey-prat/pages -f cname=sekernofey.online` ואחרי שהתעודה מוכנה `-F https_enforced=true`. סדר חשוב: להגדיר את הדומיין ב-GitHub לפני שה-DNS עונה = האתר מפנה לכתובת שלא עונה. תמונת השיתוף (`og.png`) מוגדרת ב-`index.html` על הדומיין; המקור שלה ב-`.design/og.html` (לא בגיט).

## החלטות שכדאי לזכור
- **הסטאק נשאר** (אתר סטטי + Apps Script + Sheets). הוועדות עורכות שאלות ורואות תוצאות בגיליון, 0 ₪, וזה מספיק בגדול ל-300 משקי בית. הנימוק המלא ב-`DECISIONS.md`.
- **המשך ממכשיר אחר: טלפון + קוד 4 ספרות.** טלפון לבד אפשר לכל מי שיודע מספר של שכן לטעון את התשובות שלו. הטלפון נשמר כ-HMAC עם מפתח ב-Script Properties, 5 ניסיונות לשעה.
- **הגרלה בלי זמן ובלי קוד עונה, ממוינת לפי שם**, כדי שאי אפשר יהיה לקשר הרשמה לתשובות לפי שעה.
- **אין סיסמת ברירת מחדל בקוד** (הריפו ציבורי).
- כל מידע אישי וסיסמה נשלחים ב-POST, לא בכתובת.
- בטקסט שהתושבים רואים: בלי מקף ארוך ובלי מקף עברי. `tests/validate-survey.js` בודק את זה בשאלון.

## בדיקה מקומית
- `index.html` נפתח ישירות בדפדפן, במצב תצוגה מקדימה (שמירה במכשיר בלבד).
- סקר מלא מול שרת מדומה: `preview_start` עם `e2e-mock` מ-`.claude/launch.json` (מריץ את `appscript.gs` האמיתי על גיליון בזיכרון, `config.js` מוחלף אוטומטית). `/__sheet?name=תוצאות` מראה מה נשמר.
- צילומי מובייל: Chrome headless לא יורד מתחת ל-512px, לכן מצלמים דף עם iframe ברוחב 375.
