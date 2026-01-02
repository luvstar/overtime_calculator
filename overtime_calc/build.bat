@echo off
echo =======================================================
echo == Overtime App Builder v4.0 (holidays included)
echo == Python libraries will be installed first.
echo =======================================================
echo.

echo [1/8] Installing pandas...
pip install pandas

echo [2/8] Installing selenium...
pip install selenium

echo [3/8] Installing selenium-wire...
pip install selenium-wire

echo [4/8] Installing webdriver-manager...
pip install webdriver-manager

echo [5/8] Installing holidays (For Public Holiday Logic)...
pip install holidays

echo [6/8] Installing pyinstaller (for .exe)...
pip install pyinstaller

echo [7/8] Installing blinker (for compatibility)...
pip install blinker==1.7.0

echo [8/8] All libraries installed successfully.
echo.

rem =======================================================
rem == (!!!) selenium-wire 인증서 경로 찾기 (!!!)
rem =======================================================
echo Locating selenium-wire package path...
FOR /F "delims=" %%i IN ('python -c "import seleniumwire, os; print(os.path.dirname(seleniumwire.__file__))"') DO SET "SELENIUMWIRE_PATH=%%i"

IF NOT DEFINED SELENIUMWIRE_PATH (
    echo ERROR: Could not find selenium-wire package path.
    echo Please make sure 'python' command works.
    pause
    exit /b
)

echo selenium-wire path found: %SELENIUMWIRE_PATH%
echo.

echo =======================================================
echo == Building the .exe file...
echo == Adding selenium-wire certificates (crt + key)...
echo =======================================================
echo.

:: (!!!) 파이썬 스크립트 파일명이 이전 버전(overtime_v2.py) 기준입니다. 
:: 만약 파일명을 바꾸셨다면 아래 'overtime_v2.py' 부분을 실제 파일명으로 수정하세요.

rem python -m PyInstaller --onefile --windowed ^
rem --add-data "%SELENIUMWIRE_PATH%\ca.crt;seleniumwire" ^
rem --add-data "%SELENIUMWIRE_PATH%\ca.key;seleniumwire" ^
rem overtime_v2.py

python -m PyInstaller --onefile --windowed ^
 --collect-submodules holidays ^
 --add-data "%SELENIUMWIRE_PATH%\ca.crt;seleniumwire" ^
 --add-data "%SELENIUMWIRE_PATH%\ca.key;seleniumwire" ^
 overtime_v2.py

echo.
echo =======================================================
echo == BUILD COMPLETE!
echo.
echo == Your .exe file is located in the 'dist' folder.
echo == This should be the final version.
echo =======================================================
echo.
pause