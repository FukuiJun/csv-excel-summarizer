@echo off
rem exe のビルド（--onedir 形式）。dist\LogMerger\LogMerger.exe ができる。
pip install -r requirements-dev.txt
rem バージョン：最新のタグ（例 v1.2.0。タグより後のコミットがあると v1.2.0-3-gabc1234）
set VER=
for /f "delims=" %%v in ('git describe --tags 2^>nul') do set VER=%%v
if "%VER%"=="" set VER=dev-local
python tools\write_version.py %VER%
pyinstaller --noconfirm --clean --onedir --windowed --collect-data sv_ttk --collect-all tkinterdnd2 --icon assets\LogMerger.ico --add-data "assets;assets" --hidden-import _version --name LogMerger main.py
python tools\write_version.py %VER% dist\LogMerger
echo.
echo ビルド完了: dist\LogMerger\LogMerger.exe（バージョン %VER%）
