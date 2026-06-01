# 📱 SellScan Web App — Deploy చేయడం ఎలా?
## Render.com లో Free గా Host చేయడం

---

### Step 1 — GitHub లో పెట్టడం

1. https://github.com లో account create చేయండి (free)
2. "New Repository" → name: `sellscan-app`
3. ఈ files అన్నీ upload చేయండి:
   - `app.py`
   - `screener.py`
   - `requirements.txt`
   - `Procfile`
   - `templates/index.html`  ← (templates folder తో సహా)

---

### Step 2 — Render.com లో Deploy చేయడం

1. https://render.com లో free account create చేయండి
2. "New → Web Service" click చేయండి
3. GitHub repo connect చేయండి
4. Settings:
   - **Name**: sellscan-app
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 300`
5. **Deploy** click చేయండి

Deploy అవ్వడానికి 3-5 నిమిషాలు పడుతుంది.

---

### Step 3 — Mobile లో Use చేయడం

Render మీకు ఇలాంటి URL ఇస్తుంది:
```
https://sellscan-app.onrender.com
```

- Mobile లో Chrome/Safari లో ఆ URL open చేయండి
- Excel upload చేయండి → Scan → Results!
- Laptop లో కూడా same URL — same data ✅

### Mobile Home Screen కి Add చేయడం (Optional — APK లా feel)

**Android (Chrome):**
1. Browser లో URL open చేయండి
2. Menu (3 dots) → "Add to Home screen"
3. Install → done! App icon వస్తుంది

**iPhone (Safari):**
1. Safari లో URL open చేయండి
2. Share button → "Add to Home Screen"
3. Done!

---

### ⚠️ Important Note

Render free plan లో:
- 15 నిమిషాలు use చేయకపోతే "sleep" అవుతుంది
- First request కి 30-60 seconds పడవచ్చు (wakeup time)
- తర్వాత fast గా పని చేస్తుంది

Monthly 750 hours free — daily use కి సరిపోతుంది.

---

### Timeout గురించి

Stock scan చాలా time తీసుకుంటుంది (stock ఒక్కింటికి ~2-3 sec).
- 50 stocks → ~2-3 నిమిషాలు
- 100 stocks → ~4-5 నిమిషాలు
- 200+ stocks → parts గా scan చేయండి

---
