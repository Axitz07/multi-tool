@echo off
setlocal
cd /d "%~dp0"
where node >nul 2>nul || (echo Node.js 20+ belum terpasang. & pause & exit /b 1)
if not exist node_modules (echo Menginstall dependency ToolkitIni... & npm install || (echo npm install gagal. Cek koneksi internet. & pause & exit /b 1))
npm run build || (echo Build frontend gagal. & pause & exit /b 1)
if not exist android (echo Menyiapkan proyek Android untuk pertama kali... & npx cap add android || (echo cap add android gagal. Periksa Android SDK dan dependency. & pause & exit /b 1))
npx cap sync android || (echo Sinkronisasi Android gagal. & pause & exit /b 1)
npx cap open android
