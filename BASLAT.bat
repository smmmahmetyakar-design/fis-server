@echo off
chcp 65001 >nul
echo === Fis Aktarim Araci baslatiliyor ===
where docker >nul 2>nul
if %errorlevel%==0 (
  echo Docker bulundu, container olarak baslatiliyor...
  docker build -t fis-otomasyon .
  docker run -d --name fis-otomasyon --restart unless-stopped -p 8091:8091 -v "%cd%\data:/data" fis-otomasyon
  echo Tarayicidan http://localhost:8091 adresini acin.
) else (
  echo Docker yok, Python ile baslatiliyor...
  pip install -r requirements.txt
  python -m uvicorn app.main:app --host 0.0.0.0 --port 8091
)
pause
