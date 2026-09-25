@echo off
rem exe のビルド（--onedir 形式）。dist\LogMerger\LogMerger.exe ができる。
pip install -r requirements-dev.txt
pyinstaller --noconfirm --clean --onedir --windowed --collect-data sv_ttk --name LogMerger main.py
echo.
echo ビルド完了: dist\LogMerger\LogMerger.exe
