"use client";
import { useEffect, useSyncExternalStore } from "react";
import { type Queue } from "./rules";
import { EMPTY_QUEUE, FILES, parseQueueTxt, formatQueueTxt, type CallKind } from "./local-text";
import { receiveAlerts } from "./audio-alerts";
export type { Person, Queue } from "./rules";

const CACHE_KEY = "saju-queue-v1";
const DB_NAME = "saju-local-file-handles";
const DB_STORE = "handles";
const DIR_KEY = "queue-directory";
const POLL_MS = 500;

type Permission = "granted" | "denied" | "prompt";
type DirectoryHandle = FileSystemDirectoryHandle & {
  queryPermission(options?: { mode: "read" | "readwrite" }): Promise<Permission>;
  requestPermission(options?: { mode: "read" | "readwrite" }): Promise<Permission>;
};
type BrowserFileSystemWindow = Window & {
  showDirectoryPicker(options?: { id?: string; mode?: "read" | "readwrite" }): Promise<DirectoryHandle>;
};
type View = { queue: Queue; ready: boolean; enabled: boolean; error: string; syncStatus: string; folderName: string; writable: boolean };
const initial: View = { queue: EMPTY_QUEUE, ready: false, enabled: false, error: "", syncStatus: "폴더 연결 필요", folderName: "", writable: false };
let state: View = initial;
const listeners = new Set<() => void>();
let started = false;
let directory: DirectoryHandle | null = null;
let pollTimer: ReturnType<typeof setTimeout> | undefined;
let busy = false;
let running = false;
let queueRaw: string | null = null;
const calls: Record<CallKind, string | null> = { chat: null, payment: null };
let queueTicks = 0;

function notify() { listeners.forEach(listener => listener()); }
function view(patch: Partial<View>) { state = { ...state, ...patch }; notify(); }
function persistQueue(queue: Queue) {
  try { localStorage.setItem(CACHE_KEY, JSON.stringify(queue)); } catch { /* noncritical */ }
  view({ queue, ready: true });
  window.dispatchEvent(new Event("saju-queue-change"));
}
function readableError(error: unknown): string {
  if (error instanceof DOMException && error.name === "NotAllowedError") return "폴더 접근 권한이 없습니다. '폴더 다시 연결'을 눌러 허용해 주세요.";
  if (error instanceof DOMException && error.name === "NotFoundError") return "TXT 파일을 찾지 못했습니다. Dropbox 폴더와 파일 이름을 확인해 주세요.";
  return error instanceof Error ? error.message : "로컬 파일을 읽지 못했습니다.";
}

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = () => { if (!request.result.objectStoreNames.contains(DB_STORE)) request.result.createObjectStore(DB_STORE); };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}
async function cachedDirectory(): Promise<DirectoryHandle | null> {
  const db = await openDb();
  try {
    return await new Promise((resolve, reject) => {
      const request = db.transaction(DB_STORE, "readonly").objectStore(DB_STORE).get(DIR_KEY);
      request.onsuccess = () => resolve((request.result ?? null) as DirectoryHandle | null);
      request.onerror = () => reject(request.error);
    });
  } finally { db.close(); }
}
async function saveDirectory(handle: DirectoryHandle | null) {
  const db = await openDb();
  try {
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction(DB_STORE, "readwrite");
      if (handle) tx.objectStore(DB_STORE).put(handle, DIR_KEY);
      else tx.objectStore(DB_STORE).delete(DIR_KEY);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error);
    });
  } finally { db.close(); }
}
async function readFile(name: string): Promise<string> {
  if (!directory) throw new Error("먼저 Dropbox 폴더를 연결해 주세요.");
  // Resolve a fresh file handle every time: Dropbox may replace the file while syncing.
  const handle = await directory.getFileHandle(name);
  const file = await handle.getFile();
  if (file.size > 512_000) throw new Error(`${name} 파일이 500KB를 초과했습니다.`);
  return file.text();
}
async function ensureCallFiles() {
  if (!directory) return;
  for (const file of [FILES.chat, FILES.payment]) {
    try { await directory.getFileHandle(file); }
    catch (error) {
      if (error instanceof DOMException && error.name === "NotFoundError") await directory.getFileHandle(file, { create: true });
      else throw error;
    }
  }
  try { await directory.getFileHandle(FILES.queue); }
  catch (error) {
    if (error instanceof DOMException && error.name === "NotFoundError") await directory.getFileHandle(FILES.queue, { create: true });
    else throw error;
  }
}
function clearSignals() { calls.chat = null; calls.payment = null; queueRaw = null; queueTicks = 0; }
async function refreshQueue() {
  const raw = await readFile(FILES.queue);
  if (raw === queueRaw) return;
  const parsed = parseQueueTxt(raw);
  queueRaw = raw;
  persistQueue(parsed);
}
async function refreshCall(kind: CallKind, file: string) {
  const raw = (await readFile(file)).trim().replace(/^\uFEFF/, "");
  const last = calls[kind];
  if (last === null) { calls[kind] = raw; return; } // Startup baseline; never replay stale calls.
  if (last === raw) return;
  calls[kind] = raw;
  if (!raw) return;
  // Unique CALL_ID on every Android tap guarantees a new event without duplicate playback.
  receiveAlerts([{ id: `${kind}:${raw}`, kind, createdAt: Date.now() }]);
}
async function poll() {
  if (busy || !directory || !state.enabled) return;
  busy = true;
  try {
    if (queueTicks++ % 2 === 0) await refreshQueue(); // 1000 ms
    await refreshCall("chat", FILES.chat); // 500 ms
    await refreshCall("payment", FILES.payment); // 500 ms
    if (state.error) view({ error: "", syncStatus: "로컬 TXT 연결됨" });
  } catch (error) {
    view({ error: readableError(error), syncStatus: "파일 확인 필요" });
  } finally { busy = false; }
}
function schedule() {
  if (pollTimer) clearTimeout(pollTimer);
  if (!directory || !state.enabled) return;
  const loop = async () => {
    if (!running || !directory || !state.enabled) return;
    await poll();
    if (running && directory && state.enabled) pollTimer = setTimeout(loop, POLL_MS);
  };
  running = true;
  pollTimer = setTimeout(loop, POLL_MS);
}
async function activate(handle: DirectoryHandle) {
  // No network requests and no Vercel Function calls are made here.
  directory = handle;
  clearSignals();
  const permission = await handle.queryPermission({ mode: "readwrite" });
  if (permission !== "granted") {
    running = false;
    view({ enabled: false, ready: false, folderName: handle.name, writable: false, syncStatus: "파일 권한 승인 필요", error: "폴더 다시 연결을 눌러 접근 권한을 허용해 주세요." });
    return false;
  }
  await ensureCallFiles();
  // Establish a clean baseline before starting polls.
  await refreshQueue();
  await refreshCall("chat", FILES.chat);
  await refreshCall("payment", FILES.payment);
  view({ enabled: true, writable: true, folderName: handle.name, error: "", syncStatus: "로컬 TXT 연결됨" });
  schedule();
  return true;
}
async function connect(): Promise<boolean> {
  try {
    if (!("showDirectoryPicker" in window)) throw new Error("Chrome 또는 Edge 최신 버전에서 사용해 주세요.");
    const picked = await (window as BrowserFileSystemWindow).showDirectoryPicker({ id: "saju-queue", mode: "readwrite" });
    // The picker itself is user initiated. Request only after that user gesture.
    await saveDirectory(picked);
    return await activate(picked);
  } catch (error) { view({ error: readableError(error) }); return false; }
}
async function reconnect(): Promise<boolean> {
  try {
    const handle = directory ?? await cachedDirectory();
    if (!handle) return connect();
    if (await handle.requestPermission({ mode: "readwrite" }) !== "granted") {
      view({ error: "접근 권한이 허용되지 않았습니다." }); return false;
    }
    await saveDirectory(handle);
    return await activate(handle);
  } catch (error) { view({ error: readableError(error) }); return false; }
}
async function disconnect() {
  running = false;
  if (pollTimer) clearTimeout(pollTimer);
  directory = null;
  clearSignals();
  await saveDirectory(null);
  view({ ...initial, syncStatus: "폴더 연결 해제됨" });
}
function start() {
  if (started) return;
  started = true;
  if (!("showDirectoryPicker" in window)) {
    view({ error: "Chrome 또는 Edge에서만 로컬 TXT를 자동으로 읽을 수 있습니다." }); return;
  }
  void cachedDirectory().then(async handle => {
    if (!handle) { view({ syncStatus: "Dropbox 폴더를 선택해 주세요" }); return; }
    try { await activate(handle); }
    catch (error) { view({ error: readableError(error), syncStatus: "파일 확인 필요" }); }
  }).catch(error => view({ error: readableError(error) }));
  window.addEventListener("focus", () => { if (state.enabled) void poll(); });
  document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible" && state.enabled) void poll(); });
}

async function updateQueue(change: (queue: Queue) => Queue): Promise<boolean> {
  if (!directory || !state.enabled || !state.writable) { view({ error: "먼저 Dropbox TXT 폴더를 연결해 주세요." }); return false; }
  if (busy) { view({ error: "TXT 파일을 읽고 있습니다. 잠시 후 다시 시도해 주세요." }); return false; }
  busy = true;
  try {
    const currentRaw = await readFile(FILES.queue);
    // Prevent silent overwrites if Android/Dropbox changed queue.txt since last poll.
    if (queueRaw !== currentRaw) {
      queueRaw = currentRaw;
      persistQueue(parseQueueTxt(currentRaw));
      throw new Error("다른 기기에서 대기열이 변경되었습니다. 새 명단을 확인한 후 다시 시도해 주세요.");
    }
    const next = change(parseQueueTxt(currentRaw));
    const text = formatQueueTxt(next);
    const file = await directory.getFileHandle(FILES.queue);
    const writable = await file.createWritable();
    try { await writable.write(text); await writable.close(); }
    catch (error) { await writable.abort().catch(() => {}); throw error; }
    queueRaw = text;
    persistQueue(parseQueueTxt(text));
    view({ error: "", syncStatus: "로컬 TXT 저장됨" });
    return true;
  } catch (error) { view({ error: readableError(error) }); return false; }
  finally { busy = false; }
}

export function readQueue() { return state.queue; }
export function useQueue() {
  const current = useSyncExternalStore(
    listener => { listeners.add(listener); return () => { listeners.delete(listener); }; },
    () => state, () => initial,
  );
  useEffect(start, []);
  return { ...current, update: updateQueue, connect, reconnect, disconnect };
}
