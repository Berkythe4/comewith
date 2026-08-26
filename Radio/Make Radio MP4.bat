@echo off
setlocal enabledelayedexpansion
rem ===========================================================================
rem  Make Radio MP4  --  the whole Come With NYC Radio video, start to finish.
rem
rem  Double-click it. It asks three things -- which episode, when the next one
rem  drops, and whether to update the website's running order -- then does
rem  everything else:
rem
rem      typed tracklist + database  ->  cues  ->  render  ->  verify  ->  frames
rem
rem  If something is missing it tells you what to go and get, and stops. It
rem  never renders a half-finished video.
rem
rem  What it needs in Radio\Week N\ first:
rem      the mix        any .wav / .mp3 / .m4a  -- the recorded set
rem      the tracklist  "Track List ....txt"    -- typed by hand, with times
rem      (artwork)      square image with "artwork" in the name; optional
rem
rem  Full instructions: Radio\HOW_TO_MAKE_THE_MP4.md
rem ===========================================================================
cd /d "%~dp0"
rem The repo root is whichever folder contains Radio\render. This file works
rem sitting in the Comewith folder OR inside Radio\ -- find the root instead of
rem assuming one, because `cd /d "%~dp0"` silently makes every relative path
rem below resolve against the wrong place when the file gets moved.
if not exist "Radio\render\make_episode.py" if exist "render\make_episode.py" cd ..
if not exist "Radio\render\make_episode.py" (
  echo.
  echo   ** Cannot find Radio\render\ from here.
  echo      Keep this file in the Comewith folder, or inside Radio\.
  echo.
  pause
  exit /b 1
)
title Make Radio MP4

echo.
echo   COME WITH NYC RADIO  --  make the YouTube MP4
echo   ============================================
echo.

set "EP=%~1"
if "%EP%"=="" (
  echo   Which episode? Type the number the AUDIENCE knows -- Ep 3, not show 7.
  echo.
  set /p "EP=  Episode number: "
)
if "%EP%"=="" goto :nothing

rem Folders are "Episode N". "Week N" is the old name and is still accepted
rem so a folder that never got renamed keeps working.
set "WK=Radio\Episode %EP%"
if not exist "%WK%\" if exist "Radio\Week %EP%\" set "WK=Radio\Week %EP%"
if not exist "%WK%\" (
  echo.
  echo   ** There is no folder "%WK%".
  echo      Make "Radio\Episode %EP%", then put the mix and tracklist inside.
  goto :done
)

echo.
echo   Episode %EP%   folder: %WK%
echo   ---------------------------------------------------------------

rem ---- 1. the mix ----------------------------------------------------------
set "MIX="
for %%F in ("%WK%\*.wav" "%WK%\*.WAV" "%WK%\*.aiff" "%WK%\*.flac" "%WK%\*.mp3" "%WK%\*.m4a") do (
  if exist "%%~fF" set "MIX=%%~nxF"
)
if "!MIX!"=="" (
  echo   ** No mix audio in %WK%.
  echo      Drop the recorded set in there.
  goto :done
)
echo   mix        : !MIX!

rem ---- 2. the typed tracklist ----------------------------------------------
set "TXT="
for %%F in ("%WK%\Track List*.txt" "%WK%\*racklist*.txt") do set "TXT=%%~nxF"
if "!TXT!"=="" (
  echo   tracklist  : none yet
  echo.
  echo   ** No typed tracklist in %WK%.
  echo.
  echo      Write one -- a plain .txt named "Track List ... .txt", one line
  echo      per track, in the order you played them:
  echo.
  echo          1 ALL THE TIME - John Summit - 0:00
  echo          2 Pop Pop - Channel Tres - 1:39
  echo          3 Evergreen Kings - Adriatic - 3:47
  echo.
  echo      Number, then the track, then the start time. That file is the
  echo      only place the running order and the times exist.
  goto :done
)
echo   tracklist  : !TXT!

rem ---- 3. artwork ----------------------------------------------------------
rem Find the episode cover WITHOUT insisting on a filename. This matched only
rem *artwork*, so "CWR_EP.3 COVER.JPG" was invisible and every card would have
rem silently rendered with the generic station art instead -- the kind of wrong
rem that only shows up after a 15-minute render.
rem Prefer a file that names itself; fall back to any image in the folder.
set "ART="
for %%P in (cover artwork art) do (
  for %%E in (jpg jpeg png webp) do (
    if not defined ART for %%F in ("%WK%\*%%P*.%%E") do if exist "%%~fF" set "ART=%%~fF"
  )
)
if not defined ART for %%E in (jpg jpeg png webp) do (
  if not defined ART for %%F in ("%WK%\*.%%E") do if exist "%%~fF" set "ART=%%~fF"
)
if "!ART!"=="" (
  set "ART=%CD%\Radio\Artwork\Radio_Thumbnail.jpg"
  echo   artwork    : station default ^(no episode cover found^)
) else (
  for %%F in ("!ART!") do echo   artwork    : %%~nxF
  python Radio\render\_check_cover.py "!ART!"
)

rem ---- 4. the next drop date ------------------------------------------------
echo.
echo   The closing slide says "WE PLUG BACK IN THURSDAY" and then a date.
echo   When does the NEXT episode drop?  Format: 2026-09-10
echo   ^(press Enter to use whatever the dashboard has scheduled^)
echo.
set "NEXT="
set /p "NEXT=  Next drop date: "
set "NEXTARG="
if not "!NEXT!"=="" set "NEXTARG=--next-date !NEXT!"

rem ---- 5. build the cues from the tracklist + the database -------------------
echo.
echo   Matching your tracklist against the dashboard
echo   ---------------------------------------------------------------
python Radio\render\tracklist_from_txt.py --week %EP% --out-cues
if errorlevel 1 (
  echo.
  echo   ** Could not build the tracklist. See above.
  goto :done
)

echo.
echo   Does the list above look right, in the order you played it?
set "GO="
set /p "GO=  Type Y to render, anything else to stop: "
if /i not "!GO!"=="Y" (
  echo.
  echo   Stopped. Nothing rendered. Fix the tracklist and run this again.
  goto :done
)

rem ---- 6. optionally push that order to the website --------------------------
echo.
set "SITE="
set /p "SITE=  Update the WEBSITE's tracklist order to match too? (y/N): "
if /i "!SITE!"=="Y" (
  python Radio\render\tracklist_from_txt.py --week %EP% --write-order
  if errorlevel 1 echo   ** Site order not updated -- see above. Carrying on with the render.
)

rem ---- 7. render -------------------------------------------------------------
echo.
echo   Rendering. An hour-long mix takes about 15 minutes -- leave it alone.
echo   ---------------------------------------------------------------
echo.

python Radio\render\make_episode.py --week %EP% --cover "!ART!" --cues "%WK%\EP%EP%_cues.csv" !NEXTARG!
if errorlevel 1 (
  echo.
  echo   ** The render stopped. The message above says why.
  goto :done
)

echo.
echo   Checking what came out
echo   ---------------------------------------------------------------
python Radio\render\verify_episode.py --week %EP%
if errorlevel 1 (
  echo.
  echo   ** Verification FAILED -- do not upload this yet.
  goto :done
)

rem ---- 8. the tags to paste when uploading -----------------------------------
echo.
echo   Tags for YouTube + SoundCloud
echo   ---------------------------------------------------------------
python Radio\render\make_tags.py --episode %EP%

echo.
echo   ===============================================================
echo    Done.  %WK%\CWR_Ep%EP%_YouTube.mp4
echo.
echo    Tags to paste:  %WK%\EP%EP%_tags.txt
echo.
echo    Last step is yours: open the three frames in %WK%\_preview\
echo    and look at them. A card can draw the wrong text and still
echo    pass every automatic check.
echo   ===============================================================
start "" "%CD%\%WK%\_preview"
goto :done

:nothing
echo   No episode number given -- nothing to do.

:done
echo.
pause
endlocal
