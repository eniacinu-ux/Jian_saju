"use client";
import { useState } from "react";
import QueueManager from "./QueueManager";
import QueueDisplay from "./QueueDisplay";
import styles from "./queue.module.css";

export default function QueueWindows() {
  const [open, setOpen] = useState(true);
  return <>
    {open ? <aside className={styles.queueDock} aria-label="대기열 관리">
      <div className={styles.dockTitle}>대기 명단<button aria-label="대기 명단 닫기" onClick={() => setOpen(false)}>×</button></div>
      <div className={styles.dockBody}><QueueManager embedded /></div>
    </aside> : <button className={styles.openQueueDock} onClick={() => setOpen(true)}>대기 명단 열기</button>}
    <QueueDisplay compact />
  </>;
}
