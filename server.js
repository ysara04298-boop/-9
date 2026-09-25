// server.js
// سيرفر بسيط بيستقبل نتايج الطلاب من صفحة الامتحان، بيخزنها في ملف JSON،
// وبيوفر endpoint لتصدير كل النتائج كملف Excel مباشرة من السيرفر.

const express = require('express');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const app = express();
app.use(express.json());

// ---------- تخزين البيانات ----------
// ملحوظة مهمة عن Railway: الفولدر ده بيفضل موجود طول ما السيرفر شغال
// وحتى بعد أي "restart"، لكن ممكن يتمسح لو عملتي "redeploy" جديد من غير
// ما تربطي Volume دائم. لو عايزة الداتا تفضل موجودة على مدى طويل
// (مش بس امتحان النهارده)، اعملي Railway Volume واربطيه بمسار DATA_DIR.
const DATA_DIR = path.join(__dirname, 'data');
const DATA_FILE = path.join(DATA_DIR, 'results.json');

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
if (!fs.existsSync(DATA_FILE)) fs.writeFileSync(DATA_FILE, '{}', 'utf8');

function readData() {
  try {
    return JSON.parse(fs.readFileSync(DATA_FILE, 'utf8') || '{}');
  } catch (e) {
    return {};
  }
}

function writeData(obj) {
  fs.writeFileSync(DATA_FILE, JSON.stringify(obj, null, 2), 'utf8');
}

// ---------- API ----------

// فحص إن السيرفر شغال (بيستخدمه الفرونت إند عشان يعرف لو متصل)
app.get('/api/ping', (req, res) => {
  res.json({ ok: true, time: Date.now() });
});

// هل الكود ده اتستخدم قبل كده؟
app.get('/api/players/:id', (req, res) => {
  const data = readData();
  const rec = data[req.params.id];
  res.json({ exists: !!rec, data: rec || null });
});

// حفظ/تحديث نتيجة طالب
app.post('/api/players/:id', (req, res) => {
  const id = req.params.id;
  const body = req.body || {};
  const data = readData();
  data[id] = Object.assign({ id }, body);
  writeData(data);
  res.json({ ok: true });
});

// كل النتائج (للوحة الصدارة وللتصدير)
app.get('/api/players', (req, res) => {
  const data = readData();
  res.json(Object.values(data));
});

// تصدير Excel مباشرة من السيرفر (بديل/دعم لزرار "تحميل Excel" في الصفحة)
app.get('/api/export/excel', (req, res) => {
  const data = readData();
  const rows = Object.values(data)
    .sort((a, b) => (b.score || 0) - (a.score || 0) || (a.totalTimeMs || 0) - (b.totalTimeMs || 0))
    .map((p, i) => ({
      'الترتيب': i + 1,
      'الكود': p.name || p.id,
      'رقم الموبايل': p.phone || '',
      'النقاط': p.score || 0,
      'الإجابات الصحيحة': p.correctCount || 0,
      'عدد الأسئلة': p.totalQuestions || 0,
      'أطول سلسلة صح': p.bestStreak || 0,
      'الوقت الإجمالي (ثانية)': p.totalTimeMs ? Math.round(p.totalTimeMs / 1000) : 0
    }));

  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'نتائج الامتحان');
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  res.setHeader('Content-Disposition', 'attachment; filename="نتائج_الامتحان.xlsx"');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buf);
});

// حذف كل النتائج (لو عايزة تبدئي امتحان جديد من الصفر)
app.delete('/api/players', (req, res) => {
  writeData({});
  res.json({ ok: true });
});

// ---------- تقديم صفحة الامتحان نفسها ----------
app.use(express.static(path.join(__dirname, 'public')));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
