# FSIS Logger

## Run with backend PDF generation (recommended)

This starts a local server that can generate **Inspection Order PDFs** by rendering the existing `io_fsis.html` in headless Chrome (Puppeteer), then returning it as a download.

1) Install dependencies:

```bash
npm install
```

2) Start the server:

```bash
npm start
```

3) Open the app in your browser:

- `http://127.0.0.1:3000/home.html`

Now **Download IO (PDF)** will call the backend at `POST /api/io/pdf` and download a 1‑page PDF that matches the IO HTML.

## Offline / file mode

If you open `home.html` directly (e.g. `file:///.../home.html`), the backend is not available.
In that case, **Download IO (PDF)** falls back to client-side export from `io_fsis.html?download=pdf`.

