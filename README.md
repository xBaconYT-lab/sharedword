# Shared Word

A shared Word-style document hosted on GitHub Pages. Anyone who opens the link can
edit it, and every change shows up for everyone else instantly (like Google Docs).

- **Upload .docx** – open an existing Word file in the shared editor (or drag the file onto the page).
- **Download .docx** – save the current document as a real Word file.
- **TH Sarabun PSK** is the default font. It is bundled with the site (`fonts/`) so it
  renders for everyone, and it is set as the font inside the exported Word file (16 pt).
- Each person's cursor and name are shown while they type.
- The document is also saved in your browser, so it opens even when you are offline
  and syncs again when you are back.

## How it works

Everything is static (no server of your own). The page is built with:

| Piece | Library |
|---|---|
| Real-time sync (CRDT) | [Yjs](https://github.com/yjs/yjs) |
| Relay between people | public Yjs websocket relay `wss://demos.yjs.dev/ws` + WebRTC peer-to-peer (`y-webrtc`) |
| Local copy | `y-indexeddb` |
| Editor | [Quill 2](https://quilljs.com) with `y-quill` and `quill-cursors` |
| Word import | [Mammoth](https://github.com/mwilliamson/mammoth.js) (.docx → HTML) |
| Word export | [docx](https://github.com/dolanmiu/docx) |

Everyone who opens the site joins the same "room" (the `room` value in `index.html`),
so they all see one document. Add `#some-name` to the URL, or click
*open a new blank document*, to get a separate document on the same site.

## Deploy / update on GitHub Pages

The built files (`index.html`, `app.js`, `app.css`, `fonts/`) are committed, so GitHub
Pages serves the repo root directly: **Settings → Pages → Source: Deploy from a branch → `main` / `/ (root)`**.

To change the editor code:

```bash
npm install
npm run build      # rebuilds app.js and app.css from src/
git commit -am "update" && git push
```

`npm run dev` serves the site on http://localhost:8765 and rebuilds on every reload.

## Settings (no rebuild needed)

Open `index.html` and edit `window.SHAREDWORD_CONFIG`:

- `room` – the document id shared by everyone. Change it to start a fresh document for
  everybody (old content stays in people's browsers but is no longer shown).
- `websocket` – the Yjs websocket relay. If the public demo relay ever goes away, run
  your own with `npx y-websocket` (see the y-websocket README) and put its `wss://` URL here.
- `signaling` – WebRTC signaling servers for direct peer-to-peer sync.

## Good to know

- The relay is a free public demo server run by the Yjs project. It does not promise to
  keep documents forever: while nobody has the page open, the document lives in each
  participant's browser (IndexedDB) and is re-shared when any of them returns. Use
  **Download .docx** to keep a permanent copy of important documents.
- Anyone with the link can edit. There are no accounts or permissions.
- Tables in uploaded Word files are flattened into text (the editor has no table support).
- Fonts: TH Sarabun PSK is © Department of Intellectual Property (DIP) and SIPA, Thailand,
  distributed under the DIP & SIPA Font License (see `fonts/LICENSE-DIP-SIPA.txt`).
