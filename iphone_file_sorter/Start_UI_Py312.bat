@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo ============================================================
echo  iPhone File Sorter - launch with Python 3.12 conda env
echo ============================================================
echo.

where conda >nul 2>nul
if errorlevel 1 (
  echo conda was not found. Open "Anaconda Prompt" and run this file from there,
  echo or install Anaconda/Miniconda first.
  pause
  exit /b 1
)

echo Ensuring conda env "iphone_copy" with Python 3.12 exists...
call conda create -n iphone_copy python=3.12 -y
if errorlevel 1 (
  echo Failed to create conda env.
  pause
  exit /b 1
)

call conda activate iphone_copy
if errorlevel 1 (
  echo Failed to activate conda env iphone_copy.
  pause
  exit /b 1
)

echo.
echo Using Python:
python -c "import sys; print(sys.executable); print(sys.version)"

echo.
echo Installing / updating dependencies into iphone_copy...
python -m pip install -U pip setuptools wheel
python -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo Dependency install failed.
  echo If lzfse/build errors appear, confirm Python is 3.12:
  echo   python -c "import sys; print(sys.version)"
  pause
  exit /b 1
)

echo.
echo Starting UI at http://localhost:8501
echo Keep this window open.
echo.
python -m streamlit run app.py
echo.
pause
