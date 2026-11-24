@echo off
setlocal
cd /d "%~dp0"

echo [1/5] Installing locked dependencies...
call npm ci
if errorlevel 1 goto :failed

echo [2/5] Running typecheck, unit tests, and production build...
call npm run check
if errorlevel 1 goto :failed

echo [3/5] Installing Playwright browsers...
call npx playwright install chromium firefox webkit
if errorlevel 1 goto :failed

echo [4/5] Running the production E2E suite...
call npm run test:e2e
if errorlevel 1 goto :failed

echo [5/5] Running the worksheet browser suite...
call npm run test:worksheet
if errorlevel 1 goto :failed

echo All checks passed.
exit /b 0

:failed
set "EXIT_CODE=%ERRORLEVEL%"
echo Verification failed with exit code %EXIT_CODE%.
exit /b %EXIT_CODE%
