@echo off
chcp 65001 >nul
title Installation de SchoolBot Mail

set "SRC=%~dp0SchoolBot Mail"
set "DEST=C:\SchoolBotMail"

echo.
echo  ============================================================
echo    Installation de SchoolBot Mail
echo  ============================================================
echo.

if not exist "%SRC%" (
    echo  [!] Le dossier "SchoolBot Mail" est introuvable a cote de ce fichier.
    echo.
    echo      1. Cliquez sur "Extraire tout" dans le ZIP
    echo      2. Ouvrez le dossier extrait
    echo      3. Double-cliquez a nouveau sur "Installer SchoolBot Mail"
    echo.
    pause
    exit /b 1
)

echo  Installation dans %DEST% ...
echo  (cela ne prend que quelques secondes)
echo.

if not exist "%DEST%" mkdir "%DEST%"

robocopy "%SRC%" "%DEST%" /E /NFL /NDL /NJH /NJS /NP >nul
if errorlevel 8 (
    echo  [!] Erreur pendant la copie.
    echo.
    echo      Verifiez que SchoolBot Mail est fermé, puis réessayez.
    echo.
    pause
    exit /b 1
)

:: Creation du raccourci sur le Bureau
powershell -NoProfile -Command "$s = (New-Object -ComObject WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Desktop') + '\SchoolBot Mail.lnk'); $s.TargetPath = '%DEST%\SchoolBot Mail.exe'; $s.WorkingDirectory = '%DEST%'; $s.Save()" >nul 2>&1

echo  ------------------------------------------------------------
echo   Installation terminee !
echo.
echo   - Programme installe dans : %DEST%
echo   - Raccourci "SchoolBot Mail" ajoute sur le Bureau
echo.
echo   La premiere installation du moteur IA (environ 1,5 Go) peut
echo   prendre quelques minutes : laissez la fenetre ouverte.
echo  ------------------------------------------------------------
echo.

echo  Lancement de SchoolBot Mail...
start "" "%DEST%\SchoolBot Mail.exe"

timeout /t 3 >nul
exit /b 0
