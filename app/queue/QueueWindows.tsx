"use client";
import { useRef, useState } from "react";
import QueueManager from "./QueueManager";
import QueueDisplay from "./QueueDisplay";
import styles from "./queue.module.css";

export default function QueueWindows() {
  const [open, setOpen] = useState(true);
  const [position, setPosition] = useState<{ left: number; top: number }>();
  const panel = useRef<HTMLElement>(null);
  const drag = useRef<{ x: number; y: number } | null>(null);

  return <>
    {open && <aside id="queue-dock" ref={panel} className={styles.queueDock} aria-label="대기열 관리"
      style={position && { ...position, right: "auto", bottom: "auto" }}>
      <div className={styles.dockTitle}
        style={{ cursor: "move", touchAction: "none", userSelect: "none" }}
        onPointerDown={event => {
          if (event.button !== 0 || !panel.current || (event.target as HTMLElement).closest("button")) return;
          const bounds = panel.current.getBoundingClientRect();
          drag.current = { x: event.clientX - bounds.left, y: event.clientY - bounds.top };
          event.currentTarget.setPointerCapture(event.pointerId);
        }}
        onPointerMove={event => {
          if (!drag.current || !panel.current || !event.currentTarget.hasPointerCapture(event.pointerId)) return;
          const bounds = panel.current.getBoundingClientRect();
          setPosition({
            left: Math.max(0, Math.min(event.clientX - drag.current.x, window.innerWidth - bounds.width)),
            top: Math.max(0, Math.min(event.clientY - drag.current.y, window.innerHeight - bounds.height)),
          });
        }}
        onLostPointerCapture={() => { drag.current = null; }}>
        대기 명단<button aria-label="대기 명단 닫기" onClick={() => setOpen(false)}>×</button>
      </div>
      <div className={styles.dockBody}><QueueManager embedded /></div>
    </aside>}
    <button type="button" className={styles.openQueueDock} aria-expanded={open}
      aria-controls={open ? "queue-dock" : undefined} onClick={() => setOpen(previous => !previous)}>
      {open ? "대기 명단 닫기" : "대기 명단 열기"}
    </button>
    <QueueDisplay compact />
  </>;
}