"use client";
import { useState } from "react";
import { useQueue } from "./store";
import styles from "./queue.module.css";
import { enableAlerts, disableAlerts } from "./audio-alerts";
export default function SyncSettings() {
  const { enabled, syncStatus, pending, connect, disconnect } = useQueue();
  const [code, setCode] = useState("");
  const [busy, setBusy] = useState(false);
  const [audio, setAudio] = useState(false);
  const [audioError, setAudioError] = useState("");
  return <details className={styles.syncSettings}><summary>기기 연결 · {enabled ? syncStatus : "연결 안 됨"}</summary>
    <p>같은 연결 코드를 입력한 PC와 휴대폰이 대기 명단을 공유합니다.</p>
    <form onSubmit={async e => { e.preventDefault(); setBusy(true); try { if (await connect(code)) setCode(""); } finally { setBusy(false); } }}><label htmlFor="queue-sync-code">연결 코드</label><input id="queue-sync-code" type="password" value={code} onChange={e => setCode(e.target.value)} autoComplete="off" required /><button disabled={busy || !code.trim()}>{busy ? "연결 중…" : "연결"}</button></form>
    {enabled && <><p>전송 대기 {pending.length}건</p><button type="button" onClick={() => { try { if (audio) disableAlerts(); else enableAlerts(); setAudio(!audio); setAudioError(""); } catch (e) { setAudioError(e instanceof Error ? e.message : "음성 알림을 켜지 못했습니다."); } }}>{audio ? "알림 소리 끄기" : "알림 소리 켜기"}</button><p>이 화면을 열어 두면 휴대폰 알림이 최대 약 10초 후 재생됩니다. 새로고침 후에는 소리를 다시 켜 주세요.</p>{audioError && <p role="alert">{audioError}</p>}<button onClick={() => { disableAlerts(); setAudio(false); void disconnect(); }}>연결 해제</button></>}
  </details>;
}
