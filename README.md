dash>=2.9.0
pandas>=1.5.0
plotly>=5.9.0
openpyxl
xlrd


python -m venv venv
source venv/bin/activate        # Linux / Mac
venv\Scripts\activate.bat       # Windows


pip install -r requirements.txt

python app.py

gunicorn app:app.server --bind 0.0.0.0:8050 --workers 2
