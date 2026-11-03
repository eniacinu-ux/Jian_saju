import type { NextRequest } from "next/server";
import { authenticated, configured, sameOrigin } from "../../queue-sync/auth";
import { commitDelta, getDocument } from "../../queue-sync/server";
import { validateDelta } from "../../queue-sync/protocol";
export const runtime = "nodejs";
export const dynamic = "force-dynamic";
const reply = (body: unknown, status = 200) => Response.json(body, { status, headers: { "Cache-Control": "no-store" } });
function denied(request: NextRequest) {
  if (!configured()) return reply({ error: "동기화 서버가 준비되지 않았습니다." }, 503);
  if (!authenticated(request)) return reply({ error: "연결 코드를 확인해 주세요." }, 401);
}
export async function GET(request: NextRequest) {
  const denial = denied(request); if (denial) return denial;
  try { const data = await getDocument(); return reply({ revision: data.revision, queue: data.queue }); }
  catch { return reply({ error: "서버에 연결하지 못했습니다. 변경 내용은 기기에 보관됩니다." }, 503); }
}
export async function POST(request: NextRequest) {
  const denial = denied(request); if (denial) return denial;
  if (!sameOrigin(request)) return reply({ error: "허용되지 않은 요청입니다." }, 403);
  let delta;
  try {
    const text = await request.text();
    if (text.length > 300000) return reply({ error: "요청 크기를 초과했습니다." }, 413);
    delta = validateDelta(JSON.parse(text));
  } catch (e) { return reply({ error: e instanceof Error ? e.message : "잘못된 요청입니다." }, 400); }
  try { const data = await commitDelta(delta); return reply({ revision: data.revision, queue: data.queue }); }
  catch { return reply({ error: "변경 내용을 동기화하지 못했습니다. 기기에 보관하고 다시 시도합니다." }, 503); }
}
