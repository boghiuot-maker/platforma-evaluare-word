Platforma evaluare Word - versiunea completa (V4)

Ce conține proiectul:
- app.py - aplicație Flask cu evaluare automată
- templates/ - pagini HTML (index, admin)
- uploads/ - fișiere încărcate de elevi
- data/results.csv - fișiere rezultate
- requirements.txt - dependințe
- Procfile, Dockerfile - pentru deploy

================== Deploy pe Render (pas cu pas) ==================
1. Creează un repo pe GitHub și urcă tot conținutul acestui folder.
2. Conectează GitHub la Render -> New -> Web Service -> alege repo-ul.
3. Setări recomandate:
   - Build command: pip install -r requirements.txt
   - Start command: gunicorn app:app
   - Environment variable (envs):
     - ADMIN_PASS = parola_ta (ex: profesor)
     - FLASK_SECRET = orice_text_secret
   - Region: Frankfurt (EU)
4. Apasă Deploy. După build, aplicația va fi live.

================== Utilizare ==================
- Elevii descarcă /download-template, completează și încarcă pe prima pagină.
- Profesorul intră pe /admin, introduce parola (ADMIN_PASS) și poate vedea toți elevii.
- Profesorul poate descărca CSV/XLSX cu toate rezultatele și PDF pentru fiecare elev.
