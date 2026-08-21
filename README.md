# Noah Cross-Platform Add-on

## 📖 Overview

The **Noah Cross-Platform Add-on** is a unified React-based frontend designed to run seamlessly as both a **Microsoft Office Add-in** and a **Google Workspace Add-on**. 

By utilizing a single codebase for the frontend, we minimize duplication and ensure a consistent user experience across platforms. 

### How it Works

- **Abstracted Host APIs**: The React application does not call Office or Google APIs directly. Instead, it interacts with an abstracted `DocumentHost` interface.
- **Environment Detection**: At startup, the app detects whether it is running inside MS Office (`window.Office`) or Google Workspace (`google.script`) and instantiates the appropriate host wrapper (`OfficeHost` or `GoogleHost`).
- **Single-file Compilation for Google**: Google Apps Script's `HtmlService` requires all HTML, CSS, and JS to be bundled into a single file. We use Vite with a single-file plugin (`vite-plugin-singlefile`) to inline all assets into one `index.html` file when building for Google.
- **Native Web Frames for Office**: MS Office natively supports modern web technologies, so the standard Vite build is served directly via a web server (or hosted statically) and referenced by the Office `manifest.xml`.

---

## 🚀 Deployment Steps

### Prerequisites

Ensure you have the following installed on your system:
- [Node.js](https://nodejs.org/) (v16+)
- [npm](https://npmjs.com/) or yarn
- Google Apps Script CLI (`clasp`): Installed globally via `npm install -g @google/clasp`

---

### 1. Microsoft Office Add-in

The Office Add-in relies on a manifest file (`manifest.xml`) which points to the hosted web application.

#### Local Development & Sideloading
1. Navigate to the cross-plugin directory:
   ```bash
   cd noah-cross-plugin
   ```
2. Install dependencies:
   ```bash
   npm install
   ```
3. Start the local dev server and sideload the add-in into Office (Excel/Word):
   ```bash
   npm run dev
   # OR for desktop specific sideloading:
   npm run sideload
   ```
4. To validate your manifest:
   ```bash
   npm run validate
   ```

#### Production Deployment (Office)
1. Build the production assets:
   ```bash
   npm run build
   ```
2. Host the contents of the `dist/` directory on a secure web server (must be HTTPS).
3. Update the URLs in your `manifest.xml` to point to your hosted environment.
4. Distribute the `manifest.xml` to users (via Microsoft 365 Admin Center for organizational deployment, or AppSource for public release).

---

### 2. Google Workspace Add-on

Google requires the frontend code to be pushed to a Google Apps Script project using `clasp`.

#### Initial Setup
1. Navigate to the cross-plugin directory:
   ```bash
   cd noah-cross-plugin
   ```
2. Login to `clasp` with your Google Account (only needed once):
   ```bash
   clasp login
   ```
3. If you haven't already, create a new Apps Script project or link an existing one in the `google/` directory:
   ```bash
   cd google
   clasp create --type standalone
   # OR if you have a script ID already:
   clasp clone <YOUR_SCRIPT_ID>
   ```

#### Build and Deploy (Google)
1. From the `noah-cross-plugin` directory, run the Google deployment script:
   ```bash
   npm run deploy:google
   ```
   **What this does:**
   - Runs `npm run build:google` (uses Vite to compile and inline all assets into a single `index.html`).
   - Copies the generated `index.html` into the `google/` directory.
   - Navigates into the `google/` directory and runs `clasp push` to upload the code to Google Apps Script.

2. **Test the Add-on:**
   - Open your Google Apps Script project URL (you can find this via `clasp open`).
   - Run the `onOpen` function or deploy it as a test add-on in Google Sheets/Docs to view the sidebar.

---

## 🛠️ Project Scripts Reference

| Command | Platform | Description |
|---|---|---|
| `npm run dev` | Office | Starts Vite dev server and sideloads the add-in via `manifest.xml`. |
| `npm run build` | Office | Compiles the React app into standard static assets in `dist/`. |
| `npm run sideload` | Office | Sideloads the manifest into the desktop version of Office. |
| `npm run validate` | Office | Validates the `manifest.xml` against Office schemas. |
| `npm run build:google` | Google | Compiles the React app into a single `index.html` file suitable for Apps Script. |
| `npm run deploy:google`| Google | Builds the single file, copies it, and pushes to Apps Script using `clasp`. |
