"use client";
import { useState } from "react";
import SyncSettings from "./SyncSettings";
import { useQueue } from "./store";
import { activePeople, complete, movePerson, priority, register, setFreeOrder, setKind } from "./rules";
import styles from "./queue.module.css";

export default function QueueManager({ embedded = false }: { embedded?: boolean }) {
  const { queue, ready, error, update } = useQueue();
  const [name, setName] = useState("");
  const [paid, setPaid] = useState(false);
  const [manage, setManage] = useState(false);
  const people = activePeople(queue);

  return <main className={styles.simpleManager}>
    {!embedded && <h1>대기열 관리</h1>}
    <form className={styles.simpleForm} onSubmit={async event => { event.preventDefault(); if (!ready) return; if (await update(q => register(q, name, paid, crypto.randomUUID()))) setName(""); }}>
      <label className={styles.simpleNameLabel} htmlFor="queue-name">이름</label>
      <input id="queue-name" aria-label="이름" value={name} onChange={e => setName(e.target.value)} placeholder="이름 입력" maxLength={60} required autoComplete="off" />
      <div className={styles.simpleFormBottom}><div className={styles.feeOptions} role="group" aria-label="상담 구분"><button type="button" aria-pressed={paid} onClick={() => setPaid(true)}>유료</button><button type="button" aria-pressed={!paid} onClick={() => setPaid(false)}>무료</button></div><button className={styles.addGuest} disabled={!ready || !name.trim()}>추가</button></div>
    </form>
    {error && <p role="alert" className={styles.error}>{error}</p>}
    <div className={styles.simpleCount}>대기 명단 <strong>{people.length}명</strong></div>
    <button className={styles.manageQueue} aria-expanded={manage} onClick={() => setManage(value => !value)}>순서 관리 {manage ? "접기" : "열기"}</button>
    {!ready ? <p className={styles.empty}>Dropbox 폴더 연결 필요</p> : !people.length ? <p className={styles.empty}>대기자가 없습니다.</p> : <ol className={styles.simpleList}>{people.map((person, index) => <li key={person.id}>
      <span className={styles.simpleOrder}>{index + 1}</span><strong className={styles.simpleGuest}>{person.name}</strong>
      <button className={`${styles.feeBadge} ${person.paid ? styles.paidBadge : ""}`} aria-label={`${person.name} ${person.paid ? "유료" : "무료"}, 누르면 변경`} onClick={() => { void update(q => setKind(q, person.id, !person.paid)); }}>{person.paid ? "유료" : "무료"}</button>
      <button className={styles.simpleRemove} aria-label={`${person.name} 완료`} title="상담 완료" onClick={() => { void update(q => complete(q, person.id)); }}>×</button>
      {manage && <div className={styles.orderControls}>{person.paid && <label><input type="checkbox" checked={!!person.freeOrder} onChange={e => { void update(q => setFreeOrder(q, person.id, e.target.checked)); }} /> 무료 순서</label>}<div><button aria-label={`${person.name} 위로`} disabled={!people[index - 1] || priority(people[index - 1]) !== priority(person)} onClick={() => { void update(q => movePerson(q, person.id, -1)); }}>↑</button><button aria-label={`${person.name} 아래로`} disabled={!people[index + 1] || priority(people[index + 1]) !== priority(person)} onClick={() => { void update(q => movePerson(q, person.id, 1)); }}>↓</button></div></div>}
    </li>)}</ol>}
    <SyncSettings />
  </main>;
}
