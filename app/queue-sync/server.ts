import { neon } from "@neondatabase/serverless";
import { applyDelta, type Delta, type Snapshot } from "./protocol";
type Document = Snapshot & { receipts: string[]; aliases?: Record<string, string> };
const empty = (): Document => ({ revision: 0, queue: { version: 1, people: [], call: null }, receipts: [] });
let initialized: Promise<void> | undefined;
function db() { if (!process.env.DATABASE_URL) throw new Error("동기화 저장소가 설정되지 않았습니다."); return neon(process.env.DATABASE_URL); }
const room = () => process.env.QUEUE_SYNC_ROOM || "main";
async function setup() {
  if (!initialized) initialized = (async () => {
    const sql = db();
    await sql`CREATE TABLE IF NOT EXISTS saju_queue_sync (room TEXT PRIMARY KEY, revision BIGINT NOT NULL DEFAULT 0, document JSONB NOT NULL, updated_at TIMESTAMPTZ NOT NULL DEFAULT now())`;
    await sql`CREATE TABLE IF NOT EXISTS saju_queue_sync_ops (room TEXT NOT NULL REFERENCES saju_queue_sync(room) ON DELETE CASCADE, operation_id TEXT NOT NULL, created_at TIMESTAMPTZ NOT NULL DEFAULT now(), PRIMARY KEY (room, operation_id))`;
    await sql`INSERT INTO saju_queue_sync (room, document) VALUES (${room()}, ${JSON.stringify(empty())}::jsonb) ON CONFLICT (room) DO NOTHING`;
  })().catch(e => { initialized = undefined; throw e; });
  await initialized;
}
export async function getDocument(): Promise<Document> {
  await setup();
  let rows = await db()`SELECT revision, document FROM saju_queue_sync WHERE room = ${room()}`;
  if (!rows[0]) {
    await db()`INSERT INTO saju_queue_sync (room, document) VALUES (${room()}, ${JSON.stringify(empty())}::jsonb) ON CONFLICT (room) DO NOTHING`;
    rows = await db()`SELECT revision, document FROM saju_queue_sync WHERE room = ${room()}`;
  }
  if (!rows[0]) throw new Error("동기화 저장소를 찾을 수 없습니다.");
  return { ...rows[0].document, revision: Number(rows[0].revision) } as Document;
}
export async function commitDelta(delta: Delta): Promise<Document> {
  for (let attempt = 0; attempt < 8; attempt++) {
    const previous = await getDocument();
    if (previous.receipts.includes(delta.operationId)) return previous;
    const receipt = await db()`SELECT operation_id FROM saju_queue_sync_ops WHERE room = ${room()} AND operation_id = ${delta.operationId}`;
    if (receipt.length) return previous;
    const aliases = { ...previous.aliases };
    const queue = applyDelta(previous.queue, delta, aliases);
    if (queue.people.length > 5000) throw new Error("대기열 보관 한도를 초과했습니다.");
    const next: Document = { queue, aliases, revision: previous.revision + 1, receipts: [...previous.receipts.slice(-999), delta.operationId] };
    const rows = await db()`WITH changed AS (UPDATE saju_queue_sync SET revision = ${next.revision}, document = ${JSON.stringify(next)}::jsonb, updated_at = now() WHERE room = ${room()} AND revision = ${previous.revision} AND NOT EXISTS (SELECT 1 FROM saju_queue_sync_ops WHERE room = ${room()} AND operation_id = ${delta.operationId}) RETURNING room) INSERT INTO saju_queue_sync_ops (room, operation_id) SELECT room, ${delta.operationId} FROM changed RETURNING operation_id`;
    if (rows.length) return next;
  }
  throw new Error("동시에 변경 중입니다. 잠시 후 다시 시도해 주세요.");
}
