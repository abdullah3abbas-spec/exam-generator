# 🎓 مولّد الامتحانات الذكي

موقع لتوليد امتحانات كيمياء للصف الثاني عشر - المنهج القطري

---

## 📁 هيكل المشروع

```
exam-project/
├── index.html          ← الموقع الرئيسي
├── api/
│   └── generate.js     ← Backend (Vercel Serverless)
├── vercel.json         ← إعدادات Vercel
└── README.md
```

---

## 🚀 خطوات الرفع على Vercel

### ١. رفع على GitHub
```bash
git init
git add .
git commit -m "first commit"
git remote add origin https://github.com/USERNAME/exam-generator.git
git push -u origin main
```

### ٢. ربط Vercel
1. روح vercel.com → New Project
2. Import الـ GitHub Repo
3. اضغط Deploy

### ٣. إضافة API Key
في Vercel Dashboard:
- Settings → Environment Variables
- أضف: `ANTHROPIC_API_KEY` = مفتاحك من console.anthropic.com
- اضغط Save ثم Redeploy

### ٤. ربط الدومين
في Vercel Dashboard:
- Settings → Domains
- أضف دومينك
- في إعدادات DNS بتاعتك أضف:
  ```
  Type: CNAME
  Name: @
  Value: cname.vercel-dns.com
  ```

---

## 🔑 الحصول على API Key
1. روح: https://console.anthropic.com
2. API Keys → Create Key
3. انسخه وحطه في Vercel

---

## 💰 تكلفة التشغيل
- كل امتحان ≈ $0.01 - $0.03
- Vercel Hosting: مجاني
