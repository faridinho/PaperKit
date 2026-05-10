@echo off
python -m PyInstaller --clean --noconfirm --onedir --windowed ^
  --name "PaperKit" ^
  --icon "assets\paperkit_icon.ico" ^
  --add-data "assets;assets" ^
  --collect-all tkinterdnd2 ^
  app.py
pause
