@echo off
setlocal
cd /d "%~dp0"
where node >nul 2>nul || (echo Node.js 20+ belum terpasang. & pause & exit /b 1)
where cargo >nul 2>nul || (echo Rust/Cargo belum terpasang. Install Rust dari https://rustup.rs, lalu buka terminal baru. & pause & exit /b 1)
if not exist node_modules (echo Menginstall dependency ToolkitIni... & npm install || (echo npm install gagal. Cek koneksi internet. & pause & exit /b 1))
echo Membuat installer Windows NSIS/MSI...
npm run tauri:build
if errorlevel 1 (echo Build gagal. Periksa error di atas; Windows C++ Build Tools dan WebView2 diperlukan. & pause & exit /b 1)
echo Selesai. Cek src-tauri\target\release\bundle\
pause
