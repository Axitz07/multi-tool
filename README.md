# ToolkitIni Next — Student Productivity Toolkit

A local-first toolkit for student documents, images, academic tasks, and focus sessions. Built with Vue 3 + Vite, `pdf-lib`, PDF.js, Tesseract.js, JSZip, and `docx`.

## Features included

### Documents
- Image to PDF
- OCR for photos and scanned PDFs (Indonesian + English)
- PDF compression by rasterizing pages (reduces editability/searchability)
- Merge PDF files
- Extract selected page ranges
- Reorder/rotate selected PDF pages and omit pages
- PDF text extraction to TXT or plain-text DOCX
- Export PDF pages to PNG/JPG in ZIP
- PDF annotations (simple text/highlight overlay)
- Page numbers and text watermark

### Images
- JPG/PNG/WebP conversion
- Resize and quality adjustment; batch results are downloaded as one ZIP so mobile browsers do not silently block multiple downloads
- Scanner-style grayscale, threshold, contrast, and crop controls

### Academic productivity
- IPS/GPA calculator using A=4, AB=3.5, B=3, BC=2.5, C=2, D=1, E=0 (verify the scale used by your campus)
- Assignment deadline tracker using local browser/device storage
- Basic APA 7 / IEEE citation formatter
- Pomodoro focus timer
- Extractive study notes and quiz prompts (heuristic-based; not generative AI)

## Requirements
- Node.js 20+ and npm
- Android builds: Android Studio + Android SDK + a JDK supported by that Android toolchain
- Windows EXE installer: Windows 10/11 + Rust toolchain + Microsoft C++ Build Tools + WebView2

## Run locally

```bash
npm install
npm run dev
```

## Tests and production build

```bash
npm test
npm run build
npm run preview
```

## Android APK

```bash
npm run build
npm run cap:add:android   # first time only
npm run cap:sync
npm run cap:open:android
```

Then in Android Studio, build an APK from **Build > Build Bundle(s) / APK(s) > Build APK(s)**. For a release build, configure signing in Android Studio. Keep the generated `android/` directory out of git if you do not want to commit the platform project.

## Windows EXE

```bash
npm install
npm run tauri:dev
npm run tauri:build
```

Tauri build output is produced under `src-tauri/target/release/bundle/` (NSIS/MSI when the target toolchain is available).

## Privacy and limitations

- File conversion is performed in the browser/webview locally; source files are not intentionally uploaded to a ToolkitIni server.
- OCR downloads its language model on first use. Network access may be required for that model download.
- PDF compression converts each page to a JPEG-backed page. It can lower image quality and removes selectable/searchable text from the output. Always review the exported file.
- PDF-to-DOCX exports extracted text as paragraphs; it is not a faithful layout conversion.
- PDF annotation currently adds a simple highlight/text note at a fixed location, not freehand drawing at arbitrary coordinates.
- The study notes tool is a local heuristic prototype, not an AI summary. Review all output before using it academically.
- `localStorage` data is device/browser specific and is not synced across browsers or app installations.
- No cloud sync, sign-in, push notifications, or automatic calendar integration is implemented.

## Project structure

```text
src/
  App.vue            Main UI and workflows
  style.css          Responsive styling
  lib/pdfTools.js    PDF processing functions
  lib/imageTools.js  Image processing functions
  lib/utils.js       Page parsing, GPA, citations, study notes
src-tauri/           Tauri v2 desktop app config and Rust shell
capacitor.config.ts  Capacitor Android config
```

## Recommended next steps

1. Review the calculations/grade scale against your campus' academic rules.
2. Add browser-level tests around PDF imports, downloads, OCR, and device permissions.
3. Test on a low-memory Android device and large multi-page PDFs; memory use can spike during rendering.
4. Configure app icons, splash screens, and release signing before distributing APK/EXE.

## Windows quick-start scripts

- `run-local.bat`: installs npm packages (when needed) and starts the web dev server.
- `build-exe.bat`: runs the Tauri Windows bundler after checking for Cargo.
- `build-apk.bat`: builds the web frontend, creates the Android platform project on first run, syncs it, then opens Android Studio.

These scripts do not install Node.js, Rust, Android Studio, the Android SDK, or Microsoft C++ Build Tools for you. Install those prerequisites first. Dependency installation requires a working npm registry connection.
