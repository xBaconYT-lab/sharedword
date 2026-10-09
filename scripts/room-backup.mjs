// Save a full copy of a shared room's document (Yjs state + plain text) to ./backups.
// Usage: node scripts/room-backup.mjs [room] [relay]
import * as Y from 'yjs'
import { WebsocketProvider } from 'y-websocket'
import WebSocket from 'ws'
import fs from 'node:fs'
import path from 'node:path'

const cfg = fs.readFileSync(new URL('../index.html', import.meta.url), 'utf8')
const room = process.argv[2] || /room:\s*"([^"]+)"/.exec(cfg)[1]
const relay = process.argv[3] || /websocket:\s*"([^"]+)"/.exec(cfg)[1]
const doc = new Y.Doc()
const p = new WebsocketProvider(relay, room, doc, { WebSocketPolyfill: WebSocket })
p.awareness.setLocalState(null) // do not show up as a person in the document
const timer = setTimeout(() => { console.error('timeout: relay did not sync'); process.exit(2) }, 20000)
p.on('sync', (synced) => {
  if (!synced) return
  clearTimeout(timer)
  setTimeout(() => {
    const text = doc.getText('quill').toString()
    const stamp = new Date().toISOString().replace(/[:.]/g, '-')
    const dir = path.join(path.dirname(new URL(import.meta.url).pathname), '..', 'backups')
    fs.mkdirSync(dir, { recursive: true })
    const base = path.join(dir, `${room}-${stamp}`)
    fs.writeFileSync(base + '.yjs', Buffer.from(Y.encodeStateAsUpdate(doc)))
    fs.writeFileSync(base + '.txt', text)
    console.log(JSON.stringify({ room, chars: text.length, file: base + '.yjs' }))
    p.destroy(); process.exit(0)
  }, 1500)
})
