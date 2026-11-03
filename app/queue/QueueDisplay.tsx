"use client";
import { useEffect, useRef, useState, type CSSProperties } from "react";
import { useQueue } from "./store";
import { formatQueue } from "./rules";
import styles from "./queue.module.css";

export default function QueueDisplay({ compact = false }: { compact?: boolean }) {
  const { queue, ready, error } = useQueue();
  const text = ready ? formatQueue(queue) : "[상담 대기열]불러오는 중…";
  const viewport = useRef<HTMLDivElement>(null);
  const content = useRef<HTMLSpanElement>(null);
  const [width, setWidth] = useState(0);
  const [overflow, setOverflow] = useState(false);
  useEffect(() => {
    const frame = viewport.current;
    const label = content.current;
    if (!frame || !label) return;
    let disposed = false;
    function measure() {
      if (disposed || !frame || !label) return;
      const measured = label.getBoundingClientRect().width;
      setWidth(measured);
      setOverflow(measured > frame.clientWidth + 1);
    }
    measure();
    const observer = new ResizeObserver(measure);
    observer.observe(frame);
    observer.observe(label);
    document.fonts.ready.then(measure);
    return () => { disposed = true; observer.disconnect(); };
  }, [text]);
  const motion = { "--queue-distance": `${width + 48}px`, "--queue-duration": `${Math.max(7, (width + 48) / 80)}s` } as CSSProperties;
  return <section className={compact ? `${styles.bottomDisplay} ${styles.queueTextDisplay}` : styles.display} aria-label="전광판" aria-live="polite">
    <div ref={viewport} className={styles.marqueeViewport} tabIndex={overflow ? 0 : undefined}>
      <div key={text} className={`${styles.marqueeTrack} ${overflow ? styles.marqueeMoving : ""}`} style={motion}>
        <span ref={content} className={styles.marqueeText}>{text}</span>
        {overflow && <span className={styles.marqueeText} aria-hidden="true">{text}</span>}
      </div>
    </div>
    {error && <span role="alert">{error}</span>}
  </section>;
}
