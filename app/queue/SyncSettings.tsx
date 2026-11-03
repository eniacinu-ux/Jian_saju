"use client";
import { useState } from "react";
import { useQueue } from "./store";
import styles from "./queue.module.css";
import { enableAlerts } from "./audio-alerts";

export default function SyncSettings() {
  const { enabled, syncStatus, folderName, connect, reconnect, disconnect } = useQueue();
  const [busy, setBusy] = useState(false);
  const [audioError, setAudioError] = useState("");
  async function invoke(task: () => Promise<unknown>) {
    setBusy(true);
    try { await task(); } finally { setBusy(false); }
  }
  function playAudioConfirmation() {
    try {
      enableAlerts();
      setAudioError("");
    } catch (error) { setAudioError(error instanceof Error ? error.message : "음성 알림을 켜지 못했습니다."); }
  }
  return <details className={styles.syncSettings}>
    <summary>로컬 TXT · {syncStatus}</summary>
    <p>PC Dropbox 폴더를 선택하면 대기열은 1초, 채팅·입금 호출은 0.5초 간격으로 PC 파일에서 직접 읽습니다. Vercel 서버를 호출하지 않습니다.</p>
    {folderName && <p>선택 폴더: <strong>{folderName}</strong></p>}
    <p>필요한 파일: <code>queue.txt</code>, <code>call_chat.txt</code>, <code>call_payment.txt</code></p>
    {!enabled && <>
      <button type="button" disabled={busy} onClick={() => void invoke(connect)}>Dropbox 폴더 선택</button>{" "}
      {folderName && <button type="button" disabled={busy} onClick={() => void invoke(reconnect)}>폴더 다시 연결·권한 허용</button>}
    </>}
    {enabled && <button type="button" disabled={busy} onClick={() => void invoke(disconnect)}>폴더 연결 해제</button>}
    <button type="button" onClick={playAudioConfirmation}>알림 소리 켜기</button>
    {audioError && <p role="alert">{audioError}</p>}
    <p>음성 알림은 기본으로 켜져 있습니다. 브라우저에서 재생이 차단되면 알림 소리 켜기 버튼을 눌러 주세요. 재실행 시 기존 호출은 다시 재생하지 않습니다. Chrome 백그라운드 절전 상태에서는 확인이 지연될 수 있습니다.</p>
  </details>;
}
