@echo off
chcp 65001 >nul
echo [UÇAR HUKUK] Sistem ve Kütüphane Onarımı Başlatılıyor...
python -m pip install --upgrade PyQt6 pymupdf python-docx openpyxl pandas requests reportlab pytesseract pillow
if not exist "Ucar_Hukuk_Arsiv" mkdir "Ucar_Hukuk_Arsiv"
echo [UÇAR HUKUK] Tüm kütüphaneler doğrulandı. Çıkmak için bir tuşa basınız.
pause