@echo off
setlocal

set "ZIP_NAME=package.zip"

powershell -NoProfile -Command "Compress-Archive -Path 'manifest.json','popup.css','popup.html','popup.js','icons' -DestinationPath '%ZIP_NAME%' -Force"

if %ERRORLEVEL% neq 0 (
  echo Failed to create %ZIP_NAME%.
  exit /b 1
)

echo Created %ZIP_NAME%.
endlocal
