// Merge a saved backup back into a shared room. Yjs merges are idempotent:
// content that is already in the room is not duplicated.
// Usage: node scripts/room-restore.mjs backups/<file>.yjs [room] [relay]
import * as Y from 'yjs'
import { WebsocketProvider } from 'y-websocket'
import WebSocket from 'ws'
import fs from 'node:fs'

const file = process.argv[2]
if (!file) { console.error('usage: node scripts/room-restore.mjs backups/<file>.yjs [room] [relay]'); process.exit(1) }
const cfg = fs.readFileSync(new URL('../index.html', import.meta.url), 'utf8')
const room = process.argv[3] || /room:\s*"([^"]+)"/.exec(cfg)[1]
const relay = process.argv[4] || /websocket:\s*"([^"]+)"/.exec(cfg)[1]
const doc = new Y.Doc()
const p = new WebsocketProvider(relay, room, doc, { WebSocketPolyfill: WebSocket })
p.awareness.setLocalState(null)
const timer = setTimeout(() => { console.error('timeout: relay did not sync'); process.exit(2) }, 20000)
p.on('sync', (synced) => {
  if (!synced) return
  clearTimeout(timer)
  const before = doc.getText('quill').length
  Y.applyUpdate(doc, new Uint8Array(fs.readFileSync(file)))
  const after = doc.getText('quill').length
  setTimeout(() => {
    console.log(JSON.stringify({ room, charsBefore: before, charsAfter: after }))
    p.destroy(); process.exit(0)
  }, 2500)
})
