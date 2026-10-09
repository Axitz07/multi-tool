@echo off
setlocal
cd /d "%~dp0"
where node >nul 2>nul || (echo Node.js belum terpasang. Install Node.js 20+ dulu. & pause & exit /b 1)
if not exist node_modules (echo Menginstall dependency ToolkitIni... & npm install || (echo npm install gagal. Cek koneksi internet dan versi Node.js. & pause & exit /b 1))
npm run dev
