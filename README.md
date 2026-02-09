# PST Extractor (GitHub Pages)

Pure static HTML/JS — no build, no npm. Runs entirely in the browser. Host on GitHub Pages.

## Deploy

1. Push this folder to a GitHub repo.
2. Settings → Pages → Source: deploy from branch.
3. Branch: `main`, folder: `/ (root)` or `/docs`.

**Required files** (must be in the deployed folder):
- `index.html`
- `app.js`
- `extraction.js`

Libraries (pst-parser, JSZip) load from [esm.sh](https://esm.sh) CDN — no npm or build step.

**Note:** Must be served over HTTP/HTTPS (e.g. GitHub Pages). Opening `index.html` directly (`file://`) may fail due to module/CORS restrictions.

## How it works

- **pst-parser** + **JSZip**: loaded from esm.sh CDN
- All extraction runs client-side
- No data uploaded — files stay in your browser

## Limitations

- Large PST files (>100MB) may cause the browser to slow or crash
- pst-parser may not handle every PST structure (see [pst-parser FAQ](https://github.com/IJMacD/pst-parser))
- Attachment extraction depends on PST format; some attachments may not extract

## Output structure (in ZIP)

```
extraction_YYYYMMDD/
├── pst_name/
│   ├── emails_jsonl/
│   ├── attachments/
│   └── teams_messages/
└── summary.json
```

© 2025 Victor Lü
