import { NextResponse, type NextRequest } from "next/server";
import { authenticated, configured, cookieName, sameOrigin, validKey } from "../../../queue-sync/auth";
export const dynamic = "force-dynamic";
export async function GET(request: NextRequest) { return NextResponse.json({ configured: configured(), authenticated: authenticated(request) }, { headers: { "Cache-Control": "no-store" } }); }
export async function POST(request: NextRequest) {
  if (!sameOrigin(request)) return NextResponse.json({ error: "허용되지 않은 요청입니다." }, { status: 403 });
  if (!configured()) return NextResponse.json({ error: "동기화 서버가 준비되지 않았습니다." }, { status: 503 });
  let code = "";
  try { const text = await request.text(); if (text.length > 512) throw new Error(); code = JSON.parse(text).code; } catch { return NextResponse.json({ error: "연결 코드를 입력해 주세요." }, { status: 400 }); }
  if (typeof code !== "string" || !validKey(code.trim())) return NextResponse.json({ error: "연결 코드가 올바르지 않습니다." }, { status: 401 });
  const response = NextResponse.json({ ok: true });
  response.cookies.set(cookieName, code.trim(), { httpOnly: true, secure: request.nextUrl.protocol === "https:", sameSite: "strict", path: "/api/queue-sync", maxAge: 60 * 60 * 24 * 30 });
  return response;
}
export async function DELETE(request: NextRequest) {
  if (!sameOrigin(request)) return NextResponse.json({ error: "허용되지 않은 요청입니다." }, { status: 403 });
  const response = NextResponse.json({ ok: true }); response.cookies.set(cookieName, "", { path: "/api/queue-sync", maxAge: 0 }); return response;
}
