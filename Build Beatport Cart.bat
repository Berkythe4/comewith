@echo off
rem One-click Beatport cart: runs the /beatport-cart skill headlessly (no Claude
rem Code window). Double-click = current working station. From a terminal you
rem can pass an episode number:  "Build Beatport Cart.bat" 3
cd /d "%~dp0"
echo Building your Beatport cart from the station tracklist...
echo (First run will ask you to paste your Beatport token - follow the steps.)
echo.
claude -p "/beatport-cart %*" --allowedTools "Bash,Read,Write,Edit,Glob,Grep,WebFetch,WebSearch" --output-format text
echo.
pause
