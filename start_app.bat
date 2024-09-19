@echo off
setlocal enabledelayedexpansion

:: Check if Node.js is already installed
echo Checking for Node.js installation...
node -v >nul 2>&1
if %errorlevel% equ 0 (
  echo Node.js is already installed. Version: 
  node -v
) else (
  :: Download Node.js installer
  echo Node.js is not installed. Downloading Node.js installer...
  bitsadmin /transfer "NodeJS" https://nodejs.org/dist/v18.9.1/node-v18.9.1-x64.msi %temp%\nodejs.msi

  :: Install Node.js
  echo Installing Node.js...
  msiexec /i %temp%\nodejs.msi /quiet /norestart

  :: Verify installation
  node -v >nul 2>&1
  if %errorlevel% equ 0 (
    echo Node.js installation completed successfully. Version: 
    node -v
  ) else (
    echo Node.js installation failed. Exiting...
    pause
    exit /b 1
  )
)

:: Check Node.js version
for /f "tokens=1,2 delims=v." %%a in ('node -v') do (
  set major=%%a
  set minor=%%b
)

if %major% lss 16 (
  echo Node.js version 18 or higher is required. Please update Node.js.
  pause
  exit /b 1
)

:: Check if node_modules folder exists
if not exist "node_modules" (
  echo node_modules folder not found. Running npm install...
  npm install

  :: Verify npm install success
  if %errorlevel% neq 0 (
    echo npm install failed. Exiting...
    pause
    exit /b 1
  )
) else (
  echo node_modules folder found. Skipping npm install.
)

:: Run the script
echo Starting the application...
npm start

:: Verify application start success
if %errorlevel% neq 0 (
  echo Application failed to start.
) else (
  echo Application started successfully.
)

:: Keep the window open after execution
echo Press any key to exit...
pause