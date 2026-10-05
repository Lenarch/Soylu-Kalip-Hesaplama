@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [1/2] Gerekli paketler kuruluyor...
python -m pip install -r requirements.txt
if errorlevel 1 goto hata

echo [2/2] .exe olusturuluyor (birkac dakika surebilir)...
python -m PyInstaller --noconfirm --onefile --windowed ^
  --name "KapasiteHesaplama" ^
  --collect-data customtkinter ^
  app.py
if errorlevel 1 goto hata

echo.
echo TAMAM: dist\KapasiteHesaplama.exe hazir.
pause
exit /b 0

:hata
echo.
echo HATA olustu, yukaridaki mesaji kontrol edin.
pause
exit /b 1
