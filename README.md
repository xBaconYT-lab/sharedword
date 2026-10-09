# Shared Word

A shared Word-style document hosted on GitHub Pages. Anyone who opens the link can
edit it, and every change shows up for everyone else instantly (like Google Docs).

**Live:** https://xbaconyt-lab.github.io/sharedword/

## What it does

- **Open .docx** – load an existing Word file into the shared editor (or drag it onto the
  page). Paragraph layout is kept like in Word: first-line indents (ย่อหน้า), left/right
  and hanging indents, alignment (including Thai distributed), line spacing, space
  before/after, fonts, sizes, colours, highlights, bold/italic/underline, headings,
  numbered and bulleted lists, tabs, images and tables. Style-based formatting
  (e.g. an indent defined in the Normal style) is resolved too.
- **Export** – Word (.docx), PDF (print dialog → Save as PDF) or plain text. A "file ready"
  panel offers Download and, on phones, Share (save to Files / send in LINE). LINE's in-app
  browser cannot save files, so the panel offers to reopen the page in Safari/Chrome. The Word
  file is A4 with 2.54 cm margins, TH Sarabun PSK 16 pt by default, real tab stops,
  `firstLine`/`hanging` indents, Thai-distributed justification for Thai text and
  bordered tables.
- **ย่อหน้า** – press **Tab** at the start of a paragraph (a real tab, 1.27 cm stops like
  Word) or click the ย่อหน้า toolbar button for a 2.5 cm first-line indent.
- **Tables** – the table button inserts a table; right-click a cell (tap on phones) to add or
  delete rows/columns, merge/split cells and set cell colours. Cells can hold several lines,
  and merged cells, column widths, shading and header rows come through from Word and back.
- **Page setup** – a Word file keeps its page size and margins (e.g. Thai official letters
  with a 3 cm left margin); the editor page and the exported file use them, so lines wrap like Word.
- **Fonts** – TH Sarabun PSK (bundled, default), 31 Thai fonts from Google Fonts
  (Sarabun, Kanit, Prompt, Mitr, Chakra Petch, Niramit, Krub, Mali, …), the Thai fonts
  that ship with Windows/Office (Angsana New, Cordia New, Browallia New, the UPC family,
  Leelawadee, Tahoma, …) and common Latin fonts (Arial, Calibri, Times New Roman, …).
  Fonts that are not bundled render only if installed, but the exported Word file
  always carries the exact font name.
- Each person's cursor and name are shown while they type.
- The document is also saved in your browser, so it opens offline and syncs when you are back.

## How it works

Everything is static (no server of your own). Backups: `node scripts/room-backup.mjs` saves the
shared document to `backups/` (gitignored); `node scripts/room-restore.mjs backups/<file>.yjs` merges it back. The page is built with:

| Piece | Library |
|---|---|
| Real-time sync (CRDT) | [Yjs](https://github.com/yjs/yjs) |
| Relay between people | public Yjs websocket relay `wss://demos.yjs.dev/ws` + WebRTC peer-to-peer (`y-webrtc`) |
| Local copy | `y-indexeddb` |
| Editor | [Quill 2](https://quilljs.com) with `y-quill`, `quill-cursors` and [quill-table-up](https://github.com/zzxming/quill-table-up) |
| Word import | own parser in `src/docx-import.js` (JSZip + DOMParser); [Mammoth](https://github.com/mwilliamson/mammoth.js) as fallback |
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
Tip: test on `http://localhost:8765/#some-test-room` so you do not edit the real shared document.

## Settings (no rebuild needed)

Open `index.html` and edit `window.SHAREDWORD_CONFIG`:

- `room` – the document id shared by everyone. Change it to start a fresh document for
  everybody (old content stays in people's browsers but is no longer shown).
- `websocket` – the Yjs websocket relay. If the public demo relay ever goes away, run
  your own with `npx y-websocket` (see the y-websocket README) and put its `wss://` URL here.
- `signaling` – WebRTC signaling servers for direct peer-to-peer sync.

The font list lives in `src/fonts.js` (rebuild after editing it).

## Good to know

- The relay is a free public demo server run by the Yjs project. It does not promise to
  keep documents forever: while nobody has the page open, the document lives in each
  participant's browser (IndexedDB) and is re-shared when any of them returns. Use
  **Export → Word** to keep a permanent copy of important documents.
- Anyone with the link can edit. There are no accounts or permissions.
- Headers/footers, footnotes, text boxes and tracked changes in Word files are not imported.
- Fonts: TH Sarabun PSK is © Department of Intellectual Property (DIP) and SIPA, Thailand,
  distributed under the DIP & SIPA Font License (see `fonts/LICENSE-DIP-SIPA.txt`).
  Google Fonts families are under the SIL Open Font License.
