Deploy instructions (Render)
1. Upload the project to a GitHub repo (include all files).
2. On Render.com create a new Web Service:
   - Connect your repo
   - Build command: pip install -r app/requirements.txt
   - Start command: python app/server.py
3. Ensure PORT is 5000 (Render will map automatically).
4. Put your support files comp_p1..comp_p6.pptx into app/files/ in the repo.

Local run:
1. python3 -m venv venv
2. source venv/bin/activate
3. pip install -r app/requirements.txt
4. python app/server.py
