import { activePeople, cleanName, priority, type Person, type Queue } from "../queue/rules";
export type SyncPerson = Person & { channelId?: string; note?: string; createdAt?: number };
export type Change = { id: string; fields: Partial<Omit<SyncPerson, "id">>; matchName?: string; create?: boolean };
export type Alert = { id: string; kind: "chat" | "payment"; createdAt: number };
export type Delta = { operationId: string; changes: Change[]; removedIds: string[]; order?: string[]; signal?: Alert["kind"] };
export type Snapshot = { revision: number; queue: Queue; alerts?: Alert[] };
const fields = ["name", "paid", "freeOrder", "status", "channelId", "note", "createdAt"] as const;
export function diffQueue(before: Queue, after: Queue, operationId: string): Delta {
  const old = new Map(before.people.map(p => [p.id, p as SyncPerson]));
  const changes: Change[] = [];
  for (const p of after.people as SyncPerson[]) {
    const previous = old.get(p.id);
    const changed: Change["fields"] = {};
    for (const key of fields) if ((!previous || p[key] !== previous[key]) && p[key] !== undefined) Object.assign(changed, { [key]: p[key] });
    if (Object.keys(changed).length) changes.push({ id: p.id, fields: changed, matchName: p.name, create: !previous });
  }
  const nextIds = new Set(after.people.map(p => p.id));
  const order = activePeople(after).map(p => p.id);
  return { operationId, changes, removedIds: before.people.filter(p => !nextIds.has(p.id)).map(p => p.id), ...(JSON.stringify(order) !== JSON.stringify(activePeople(before).map(p => p.id)) ? { order } : {}) };
}
export function hasChanges(delta: Delta) { return !!(delta.changes.length || delta.removedIds.length || delta.order); }
const key = (p: SyncPerson) => cleanName(p.name).toLocaleLowerCase();
export function applyDelta(queue: Queue, delta: Delta, identityAliases: Record<string, string> = {}): Queue {
  let people: SyncPerson[] = queue.people.map(p => ({ ...p }));
  const aliases = new Map<string, string>(Object.entries(identityAliases));
  for (const change of delta.changes) {
    const canonical = aliases.get(change.id) ?? change.id;
    let index = people.findIndex(p => p.id === canonical);
    if (index < 0 && !change.create && change.matchName) index = people.findIndex(p => p.status !== "done" && key(p) === cleanName(change.matchName!).toLocaleLowerCase());
    if (index < 0) {
      if (!change.fields.name) continue;
      const row = { paid: false, freeOrder: false, status: "waiting", ...change.fields, id: change.id, name: cleanName(change.fields.name) } as SyncPerson;
      const duplicate = people.find(p => (p.status !== "done" && row.status !== "done" && key(p) === key(row)) || (row.channelId && p.channelId === row.channelId));
      if (duplicate) {
        aliases.set(row.id, duplicate.id); identityAliases[row.id] = duplicate.id;
        if (duplicate.status === "done" && row.status !== "done") { people = people.filter(p => p.id !== duplicate.id); people.push({ ...duplicate, ...row, id: duplicate.id }); }
        continue;
      }
      people.push(row);
    } else {
      const previous = people[index];
      if (change.id !== previous.id) { aliases.set(change.id, previous.id); identityAliases[change.id] = previous.id; }
      // Initial uploads of an already-known ID must not overwrite newer server edits.
      if (change.create) continue;
      const row = { ...previous, ...change.fields, ...(change.fields.name ? { name: cleanName(change.fields.name) } : {}) } as SyncPerson;
      if (row.status !== "done" && people.some(p => p.id !== row.id && p.status !== "done" && key(p) === key(row))) throw new Error("같은 이름이 이미 대기 중입니다.");
      if (priority(previous) !== priority(row) || (previous.status === "done" && row.status !== "done")) { people.splice(index, 1); people.push(row); } else people[index] = row;
    }
  }
  const removed = new Set(delta.removedIds.map(id => aliases.get(id) ?? id));
  people = people.map(p => removed.has(p.id) ? { ...p, status: "done" } : p);
  let active = activePeople({ ...queue, people }) as SyncPerson[];
  if (delta.order) {
    const order = [...new Set(delta.order.map(id => aliases.get(id) ?? id))];
    for (const group of [true, false]) {
      const requested = order.map(id => active.find(p => p.id === id)).filter((p): p is SyncPerson => !!p && priority(p) === group);
      const ids = new Set(requested.map(p => p.id));
      let at = 0;
      active = active.map(p => ids.has(p.id) ? requested[at++] : p);
    }
  }
  return { version: 1, call: null, people: [...active, ...people.filter(p => p.status === "done")] };
}
export function validateDelta(value: unknown): Delta {
  if (!value || typeof value !== "object") throw new Error("잘못된 동기화 요청입니다.");
  const d = value as Delta;
  if (d.signal !== undefined && d.signal !== "chat" && d.signal !== "payment") throw new Error("잘못된 알림입니다.");
  const id = (s: unknown) => typeof s === "string" && s.length > 0 && s.length <= 160;
  if (!id(d.operationId) || !Array.isArray(d.changes) || d.changes.length > 1000 || !Array.isArray(d.removedIds) || d.removedIds.length > 1000 || !d.removedIds.every(id) || (d.order !== undefined && (!Array.isArray(d.order) || d.order.length > 1000 || !d.order.every(id)))) throw new Error("잘못된 동기화 요청입니다.");
  for (const c of d.changes) {
    if (!c || !id(c.id) || !c.fields || typeof c.fields !== "object" || Array.isArray(c.fields)) throw new Error("잘못된 대기자 정보입니다.");
    if ((c.matchName !== undefined && (typeof c.matchName !== "string" || c.matchName.length > 256)) || (c.create !== undefined && typeof c.create !== "boolean")) throw new Error("잘못된 대기자 정보입니다.");
    for (const [k, v] of Object.entries(c.fields)) {
      if (!fields.includes(k as typeof fields[number])) throw new Error("지원하지 않는 대기자 정보입니다.");
      if ((k === "paid" || k === "freeOrder") && typeof v !== "boolean") throw new Error("잘못된 상담 구분입니다.");
      if (k === "status" && !["waiting", "called", "done"].includes(v as string)) throw new Error("잘못된 대기 상태입니다.");
      if (["name", "channelId", "note"].includes(k) && (typeof v !== "string" || v.length > (k === "note" ? 2000 : 256) || (k === "name" && !cleanName(v)))) throw new Error("잘못된 이름 또는 메모입니다.");
      if (k === "createdAt" && (typeof v !== "number" || !Number.isSafeInteger(v) || v < 0)) throw new Error("잘못된 등록 시간입니다.");
    }
  }
  return d;
}
