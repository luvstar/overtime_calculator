@echo off
echo =======================================================
echo == Overtime App Builder v3.0 (with CRT + KEY Fix)
echo == Python libraries will be installed first.
echo =======================================================
echo.

echo [1/7] Installing pandas...
pip install pandas

echo [2/7] Installing selenium...
pip install selenium

echo [3/7] Installing selenium-wire...
pip install selenium-wire

echo [4/7] Installing webdriver-manager...
pip install webdriver-manager

echo [5/7] Installing pyinstaller (for .exe)...
pip install pyinstaller

echo [6/7] Installing blinker (for compatibility)...
pip install blinker==1.7.0

echo [7/7] All libraries installed successfully.
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

:: (!!!) --add-data 플래그를 두 번 사용하여 crt와 key 파일을 모두 포함
python -m PyInstaller --onefile --windowed ^
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