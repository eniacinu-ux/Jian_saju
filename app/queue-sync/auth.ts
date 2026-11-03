import { createHash, timingSafeEqual } from "node:crypto";
import type { NextRequest } from "next/server";
export const cookieName = "saju_queue_session";
export function configured() { return !!(process.env.QUEUE_SYNC_KEY && process.env.DATABASE_URL); }
export function validKey(key: string) {
  const expected = process.env.QUEUE_SYNC_KEY;
  return !!expected && timingSafeEqual(createHash("sha256").update(key).digest(), createHash("sha256").update(expected).digest());
}
export function authenticated(request: NextRequest) {
  const authorization = request.headers.get("authorization");
  return validKey(authorization?.startsWith("Bearer ") ? authorization.slice(7) : request.cookies.get(cookieName)?.value ?? "");
}
export function sameOrigin(request: NextRequest) {
  if (request.headers.get("sec-fetch-site") === "cross-site") return false;
  const origin = request.headers.get("origin");
  if (!origin) return true; // Native clients authenticate with their bearer key.
  try { const url = new URL(origin); return ["http:", "https:"].includes(url.protocol) && url.host === (request.headers.get("host") || request.nextUrl.host); }
  catch { return false; }
}
