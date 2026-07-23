@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

echo ============================================================
echo   CAI DAT TU DONG VA DONG GOI CONG CU THANH FILE .EXE
echo   (Lan dau chay co the mat 5-15 phut tuy toc do mang)
echo ============================================================
echo.

net session >nul 2>nul
if errorlevel 1 (
    echo [CANH BAO] Nen chay file nay bang quyen Administrator
    echo            ^(chuot phai vao file --^> "Run as administrator"^)
    echo            de viec tu cai Python / LibreOffice khong bi loi quyen.
    echo.
    timeout /t 4 >nul
)

call :CheckPython
if errorlevel 1 goto :Fail

call :CheckLibreOffice

call :InstallPipDeps
if errorlevel 1 goto :Fail

call :BuildExe
if errorlevel 1 goto :Fail

echo.
echo ============================================================
echo   XONG! Ket qua nam trong thu muc "dist":
echo     dist\CongCuSoanBaiGiang.exe   ^<- bam dup de chay
echo     dist\sample\                 ^<- file mau de thu ngay
echo.
echo   Tu lan sau, chi can bam dup vao file .exe nay,
echo   khong can chay lai file .bat nua.
echo   Co the copy ca thu muc "dist" sang may khac de dung
echo   ^(may do van can co san LibreOffice^).
echo ============================================================
pause
exit /b 0

:Fail
echo.
echo Qua trinh dung lai do gap loi o buoc phia tren.
pause
exit /b 1


REM ================================================================
REM  1) PYTHON
REM ================================================================
:CheckPython
echo [1/4] Kiem tra Python...
where python >nul 2>nul
if not errorlevel 1 (
    echo       Da co Python san.
    python --version
    exit /b 0
)

echo       Chua co Python. Dang cai tu dong...
where winget >nul 2>nul
if not errorlevel 1 (
    winget install -e --id Python.Python.3.12 --silent --accept-source-agreements --accept-package-agreements
    goto :RefreshPathAfterPython
)

echo       May khong co winget, dang tai bo cai Python truc tiep tu python.org...
powershell -NoProfile -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.4/python-3.12.4-amd64.exe' -OutFile '%TEMP%\python_installer.exe'"
if not exist "%TEMP%\python_installer.exe" (
    echo [LOI] Khong tai duoc bo cai Python. Kiem tra ket noi mang roi chay lai file nay.
    exit /b 1
)
"%TEMP%\python_installer.exe" /quiet InstallAllUsers=1 PrependPath=1 Include_test=0

:RefreshPathAfterPython
REM Lam moi bien PATH ngay trong phien lam viec nay, khong can mo cua so moi
for /f "usebackq tokens=2,*" %%A in (`reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v Path 2^>nul`) do set "SYS_PATH=%%B"
for /f "usebackq tokens=2,*" %%A in (`reg query "HKCU\Environment" /v Path 2^>nul`) do set "USR_PATH=%%B"
set "PATH=%SYS_PATH%;%USR_PATH%;%PATH%"

where python >nul 2>nul
if errorlevel 1 (
    echo [LOI] Da cai Python nhung cua so nay chua nhan duoc thay doi.
    echo       Hay DONG cua so nay va bam dup lai file .bat mot lan nua.
    exit /b 1
)
echo       Da cai xong Python.
python --version
exit /b 0


REM ================================================================
REM  2) LIBREOFFICE (dung de tao ban xem truoc PDF)
REM ================================================================
:CheckLibreOffice
echo.
echo [2/4] Kiem tra LibreOffice...

if exist "C:\Program Files\LibreOffice\program\soffice.exe" goto :LO_Found
if exist "C:\Program Files (x86)\LibreOffice\program\soffice.exe" goto :LO_Found
where soffice >nul 2>nul
if not errorlevel 1 goto :LO_Found
goto :LO_Install

:LO_Found
echo       Da co LibreOffice san.
exit /b 0

:LO_Install
echo       Chua co LibreOffice. Dang cai tu dong ^(file khoang 300MB, co the mat vai phut^)...
where winget >nul 2>nul
if not errorlevel 1 (
    winget install -e --id TheDocumentFoundation.LibreOffice --silent --accept-source-agreements --accept-package-agreements
    echo       Da yeu cau cai LibreOffice qua winget.
    exit /b 0
)

echo       May khong co winget, dang tai bo cai LibreOffice truc tiep...
echo       ^(Neu duong dan tai ben duoi bi cu, hay tai thu cong tai libreoffice.org^)
powershell -NoProfile -Command "Invoke-WebRequest -Uri 'https://download.documentfoundation.org/libreoffice/stable/24.8.4/win/x86_64/LibreOffice_24.8.4_Win_x86-64.msi' -OutFile '%TEMP%\libreoffice_installer.msi'"
if not exist "%TEMP%\libreoffice_installer.msi" (
    echo [CANH BAO] Khong tai duoc bo cai LibreOffice tu dong.
    echo             Hay tai va cai thu cong tai: https://www.libreoffice.org/download/
    echo             ^(Chuc nang Xem truoc se khong dung duoc cho toi khi cai xong^)
    exit /b 0
)
msiexec /i "%TEMP%\libreoffice_installer.msi" /quiet /norestart
echo       Da cai xong LibreOffice.
exit /b 0


REM ================================================================
REM  3) THU VIEN PYTHON
REM ================================================================
:InstallPipDeps
echo.
echo [3/4] Dang cai thu vien Python can thiet...
python -m pip install --upgrade pip --quiet
python -m pip install -r requirements.txt pyinstaller --quiet
if errorlevel 1 (
    echo [LOI] Cai thu vien khong thanh cong. Kiem tra ket noi mang roi chay lai file nay.
    exit /b 1
)
echo       Da cai xong thu vien.
exit /b 0


REM ================================================================
REM  4) DONG GOI THANH FILE .EXE
REM ================================================================
:BuildExe
echo.
echo [4/4] Dang dong goi thanh file .exe...
python -m PyInstaller --noconfirm --onefile --console ^
    --name "CongCuSoanBaiGiang" ^
    --add-data "templates;templates" ^
    --add-data "static;static" ^
    --add-data "sample/MAU_BaiGiang.docx;sample" ^
    app.py
if errorlevel 1 (
    echo [LOI] Dong goi khong thanh cong. Xem thong bao loi phia tren.
    exit /b 1
)

echo       Dang sao chep du lieu mau...
xcopy /E /I /Y sample "dist\sample" >nul
rd /s /q build >nul 2>nul
del /q "CongCuSoanBaiGiang.spec" >nul 2>nul
exit /b 0
