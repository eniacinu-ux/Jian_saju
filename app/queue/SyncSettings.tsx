"use client";
import { useState } from "react";
import { useQueue } from "./store";
import styles from "./queue.module.css";
export default function SyncSettings() {
  const { enabled, syncStatus, pending, connect, disconnect } = useQueue();
  const [code, setCode] = useState("");
  const [busy, setBusy] = useState(false);
  return <details className={styles.syncSettings}><summary>기기 연결 · {enabled ? syncStatus : "연결 안 됨"}</summary>
    <p>같은 연결 코드를 입력한 PC와 휴대폰이 대기 명단을 공유합니다.</p>
    <form onSubmit={async e => { e.preventDefault(); setBusy(true); try { if (await connect(code)) setCode(""); } finally { setBusy(false); } }}><label htmlFor="queue-sync-code">연결 코드</label><input id="queue-sync-code" type="password" value={code} onChange={e => setCode(e.target.value)} autoComplete="off" required /><button disabled={busy || !code.trim()}>{busy ? "연결 중…" : "연결"}</button></form>
    {enabled && <><p>전송 대기 {pending.length}건</p><button onClick={() => void disconnect()}>연결 해제</button></>}
  </details>;
}
