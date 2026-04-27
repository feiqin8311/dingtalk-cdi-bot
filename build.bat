@echo off
setlocal

cd /d "%~dp0"

set "ENV_NAME=%~1"
if "%ENV_NAME%"=="" set "ENV_NAME=dingtalk-bot"

echo [1/3] Activating conda environment: %ENV_NAME%

set "CONDA_BAT="

if defined CONDA_EXE (
    for %%I in ("%CONDA_EXE%") do (
        if /I "%%~nxI"=="conda.bat" set "CONDA_BAT=%%~fI"
        if /I "%%~nxI"=="conda.exe" if exist "%%~dpI..\condabin\conda.bat" set "CONDA_BAT=%%~dpI..\condabin\conda.bat"
    )
) 

if not defined CONDA_BAT (
    if exist "D:\miniconda3\condabin\conda.bat" (
        set "CONDA_BAT=D:\miniconda3\condabin\conda.bat"
    ) else if exist "%USERPROFILE%\miniconda3\condabin\conda.bat" (
        set "CONDA_BAT=%USERPROFILE%\miniconda3\condabin\conda.bat"
    ) else if exist "%USERPROFILE%\anaconda3\condabin\conda.bat" (
        set "CONDA_BAT=%USERPROFILE%\anaconda3\condabin\conda.bat"
    ) else (
        echo Unable to find conda.bat. Please start this from an Anaconda Prompt, or edit build.bat with your conda path.
        pause
        exit /b 1
    )
)

call "%CONDA_BAT%" activate "%ENV_NAME%"

if errorlevel 1 (
    echo Failed to activate conda environment: %ENV_NAME%
    pause
    exit /b 1
)

echo [2/3] Running PyInstaller...
pyinstaller --noconfirm customs_declaration_info.spec
if errorlevel 1 (
    echo PyInstaller build failed.
    pause
    exit /b 1
)

echo [3/3] Build completed.
pause

endlocal
