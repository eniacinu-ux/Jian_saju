"use client";
import { useEffect, useSyncExternalStore } from "react";
import { applyDelta, diffQueue, hasChanges, type Delta, type Snapshot } from "../queue-sync/protocol";
import { cleanName, type Person, type Queue } from "./rules";
import { receiveAlerts } from "./audio-alerts";
export type { Person, Queue } from "./rules";
const KEY = "saju-queue-v1";
const STATE_KEY = "saju-queue-sync-v1";
const empty: Queue = { version: 1, people: [], call: null };
type Local = { queue: Queue; pending: Delta[]; enabled: boolean; revision: number };
type View = Local & { ready: boolean; error: string; syncStatus: string };
const initial: View = { queue: empty, pending: [], enabled: false, revision: -1, ready: false, error: "", syncStatus: "연결 안 됨" };
let state: View = initial;
const listeners = new Set<() => void>();
let started = false;
let busy = false;
let timer: ReturnType<typeof setTimeout> | undefined;
function notify() { for (const listener of listeners) listener(); }
function setView(patch: Partial<View>) { state = { ...state, ...patch }; notify(); }
function persist(next: Local) {
  localStorage.setItem(STATE_KEY, JSON.stringify(next));
  localStorage.setItem(KEY, JSON.stringify(next.queue));
  state = { ...state, ...next }; notify();
  window.dispatchEvent(new Event("saju-queue-change"));
}
function load(): Local {
  const saved = localStorage.getItem(STATE_KEY);
  if (saved) {
    const data = JSON.parse(saved);
    if (data?.queue?.version !== 1 || !Array.isArray(data.queue.people) || !Array.isArray(data.pending)) throw new Error("대기열 저장소를 확인해 주세요.");
    return data;
  }
  const raw = localStorage.getItem(KEY);
  const q = raw ? JSON.parse(raw) : empty;
  if (q.version !== 1 || !Array.isArray(q.people)) throw new Error("대기열 저장소를 확인해 주세요.");
  return { queue: { ...q, people: q.people.map((p: Person) => ({ ...p, name: cleanName(p.name), paid: !!p.paid, freeOrder: !!p.freeOrder })) }, pending: [], enabled: false, revision: -1 };
}
export function readQueue(): Queue { return load().queue; }
function receive(data: Snapshot, acknowledged?: string) {
  const latest = load();
  if (data.revision < latest.revision) return;
  receiveAlerts(data.alerts ?? []);
  const pending = latest.pending.filter(d => d.operationId !== acknowledged);
  const aliases: Record<string, string> = {};
  const queue = pending.reduce((q, d) => applyDelta(q, d, aliases), data.queue);
  persist({ ...latest, queue, pending, revision: data.revision });
  setView({ error: "", syncStatus: pending.length ? "전송 대기" : "동기화됨" });
}
async function request(path: string, options: RequestInit = {}) {
  const response = await fetch(path, { ...options, cache: "no-store", signal: AbortSignal.timeout(12000), headers: { "Content-Type": "application/json", ...options.headers } });
  if (response.status === 304) return null;
  const data = await response.json();
  if (!response.ok) throw new Error(response.status === 401 ? "연결 코드를 다시 입력해 주세요." : data.error || "연결을 확인해 주세요.");
  return data;
}
async function tick() {
  if (busy || !state.enabled || document.visibilityState !== "visible") return;
  busy = true;
  try {
    const local = load();
    const pending = local.pending[0];
    if (pending) receive(await request("/api/queue-sync", { method: "POST", body: JSON.stringify(pending) }), pending.operationId);
    else {
      const data = await request(`/api/queue-sync${local.revision >= 0 ? `?since=${local.revision}` : ""}`);
      if (data) receive(data);
      else setView({ error: "", syncStatus: load().pending.length ? "전송 대기" : "동기화됨" });
    }
  } catch (e) { setView({ syncStatus: "연결 대기", error: e instanceof Error ? e.message : "연결을 확인해 주세요." }); }
  finally { busy = false; schedule(state.pending.length && !state.error ? 100 : 10000); }
}
function schedule(delay = 10000) { if (timer) clearTimeout(timer); timer = setTimeout(() => { void tick(); }, delay); }
function start() {
  if (started) return;
  started = true;
  try { setView({ ...load(), ready: true, syncStatus: "연결 안 됨" }); } catch (e) { setView({ error: e instanceof Error ? e.message : "저장소 오류" }); }
  const wake = () => { schedule(0); };
  document.addEventListener("visibilitychange", wake);
  window.addEventListener("online", wake);
  window.addEventListener("storage", event => { if (event.key === STATE_KEY) { try { setView(load()); wake(); } catch { setView({ error: "대기열 저장소 오류" }); } } });
  if (state.enabled) schedule(0);
}
export function useQueue() {
  const current = useSyncExternalStore(listener => { listeners.add(listener); return () => { listeners.delete(listener); }; }, () => state, () => initial);
  useEffect(start, []);
  function update(change: (q: Queue) => Queue) {
    try {
      const local = load();
      const queue = change(local.queue);
      const delta = diffQueue(local.queue, queue, crypto.randomUUID());
      persist({ ...local, queue, pending: (local.enabled || local.revision >= 0 || local.pending.length > 0) && hasChanges(delta) ? [...local.pending, delta] : local.pending });
      setView({ error: "" }); if (local.enabled) schedule(0);
      return true;
    } catch (e) { setView({ error: e instanceof Error ? e.message : "저장하지 못했습니다." }); return false; }
  }
  async function connect(code: string) {
    try {
      await request("/api/queue-sync/session", { method: "POST", body: JSON.stringify({ code }) });
      const local = load();
      if (!local.enabled) {
        localStorage.setItem(`saju-queue-before-sync-${Date.now()}`, JSON.stringify(local.queue));
        const seed = diffQueue(empty, local.queue, crypto.randomUUID());
        persist({ ...local, enabled: true, pending: [...(local.revision < 0 && local.pending.length === 0 && hasChanges(seed) ? [seed] : []), ...local.pending] });
      }
      setView({ error: "", syncStatus: "연결 중" }); schedule(0); return true;
    } catch (e) { setView({ error: e instanceof Error ? e.message : "연결하지 못했습니다." }); return false; }
  }
  async function disconnect() {
    try { await request("/api/queue-sync/session", { method: "DELETE" }); const local = load(); persist({ ...local, enabled: false }); setView({ syncStatus: "연결 안 됨", error: "" }); } catch { setView({ error: "연결 해제에 실패했습니다." }); }
  }
  return { ...current, update, connect, disconnect };
}
