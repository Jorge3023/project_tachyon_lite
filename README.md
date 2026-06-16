Herramienta web para análisis de reportes de producción.
Sube el Excel crudo → descarga el análisis. Sin base de datos, sin login complejo.

---

## Correr localmente

```bash
pip install -r requirements.txt
python app.py
```
Abre: http://localhost:5000


## Deploy gratuito en Render

1. Sube a GitHub
2. Entra a https://render.com → "New Web Service"
3. Conecta el repo
4. Build command: `pip install -r requirements.txt`
5. Start command: `gunicorn app:app --bind 0.0.0.0:$PORT --timeout 120`
6. En "Environment Variables": `ACCESS_KEY` = tu clave
7. Deploy

## Estructura

```
smt-analyzer/
├── app.py           ← Servidor Flask
├── excel_parser.py  ← Lógica de análisis
├── index.html       ← UI completa (un solo archivo)
├── requirements.txt
├── Procfile         ← Para Railway/Render
└── .gitignore
```
