@echo off
REM ---------------------------------------------------------------------------
REM  Rebuild the proctor site from the Excel sheet and publish it to GitHub.
REM  Double-click this file after you change "Proctor Schedule by Date.xlsx".
REM ---------------------------------------------------------------------------
setlocal
cd /d "%~dp0"

REM ---------------------------------------------------------------------------
REM  The site was closed on 25 September 2026, after the September exam period.
REM  index.html is now a hand-written "this period has ended" notice, and a
REM  rebuild would overwrite it with the old schedule.
REM
REM  To reopen for the next exam period:
REM    1. set WINDOW_START in build.py to the first day of that period
REM    2. delete the four lines below, down to and including the "exit /b 0"
REM    3. re-enable publishing: in .github\workflows\pages.yml, put back the
REM       "push: branches: [main]" trigger
REM ---------------------------------------------------------------------------
echo.
echo  The schedule site is CLOSED for the September 2026 exam period.
echo  Nothing was rebuilt or published. See the notes inside this file to reopen it.
pause
exit /b 0

echo.
echo  Rebuilding from the Excel sheet...
echo  ---------------------------------------------------------------
py -3 build.py
if errorlevel 1 (
  echo.
  echo  BUILD FAILED - nothing was published. See the message above.
  pause
  exit /b 1
)

echo.
echo  Publishing to GitHub...
echo  ---------------------------------------------------------------
git add -A
git diff --cached --quiet
if not errorlevel 1 (
  echo  No changes since the last sync - nothing to publish.
  pause
  exit /b 0
)

for /f "tokens=*" %%d in ('powershell -NoProfile -Command "Get-Date -Format \"yyyy-MM-dd HH:mm\""') do set STAMP=%%d
git commit -m "Update proctor schedule - %STAMP%"
if errorlevel 1 (
  echo.
  echo  COMMIT FAILED. If this is the first run, set your identity once:
  echo     git config --global user.name  "Your Name"
  echo     git config --global user.email "you@htu.edu.jo"
  pause
  exit /b 1
)

git push
if errorlevel 1 (
  echo.
  echo  PUSH FAILED. Check your network and that the 'origin' remote is set.
  pause
  exit /b 1
)

echo.
echo  ---------------------------------------------------------------
echo  Done. The site updates in about a minute.
echo.
pause
