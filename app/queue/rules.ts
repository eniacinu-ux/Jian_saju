export type Person = { id: string; name: string; paid?: boolean; freeOrder?: boolean; status: "waiting" | "called" | "done"; kindLabel?: string };
export type Queue = { version: 1; people: Person[]; call: { id: string; personId: string; name: string } | null };

export function cleanName(name: string) { return name.trim().replace(/^@+/, "").trim().normalize("NFC").replace(/-(?=[a-z0-9]*[a-z])(?=[a-z0-9]*\d)[a-z0-9]{2,12}$/i, ""); }
const nameKey = (name: string) => cleanName(name).toLocaleLowerCase();
export function priority(person: Person) { return !!person.paid && !person.freeOrder; }
export function activePeople(queue: Queue) {
  return queue.people.filter(p => p.status !== "done").sort((a, b) => Number(priority(b)) - Number(priority(a)));
}
function putLast(queue: Queue, person: Person): Queue {
  const rest = queue.people.filter(p => p.id !== person.id);
  const active = activePeople({ ...queue, people: rest });
  const first = active.filter(priority);
  const general = active.filter(p => !priority(p));
  if (priority(person)) first.push(person); else general.push(person);
  return { ...queue, people: [...first, ...general, ...rest.filter(p => p.status === "done")] };
}
export function register(queue: Queue, raw: string, paid: boolean, id: string): Queue {
  const name = cleanName(raw);
  if (!name) throw new Error("이름을 입력해 주세요.");
  if (activePeople(queue).some(p => nameKey(p.name) === nameKey(name))) throw new Error("이미 대기 명단에 있는 이름입니다.");
  const previous = [...queue.people].reverse().find(p => p.status === "done" && nameKey(p.name) === nameKey(name));
  return putLast(queue, { id: previous?.id ?? id, name, paid, freeOrder: paid ? false : previous?.freeOrder ?? false, status: "waiting" });
}
export function setKind(queue: Queue, id: string, paid: boolean) {
  const person = queue.people.find(p => p.id === id && p.status !== "done");
  return person ? putLast(queue, { ...person, paid, freeOrder: false, kindLabel: undefined }) : queue;
}
export function setFreeOrder(queue: Queue, id: string, freeOrder: boolean) {
  const person = queue.people.find(p => p.id === id && p.status !== "done");
  return person?.paid ? putLast(queue, { ...person, freeOrder }) : queue;
}
export function movePerson(queue: Queue, id: string, direction: -1 | 1): Queue {
  const people = activePeople(queue);
  const from = people.findIndex(p => p.id === id);
  const to = from + direction;
  if (from < 0 || !people[to] || priority(people[from]) !== priority(people[to])) return queue;
  [people[from], people[to]] = [people[to], people[from]];
  return { ...queue, people: [...people, ...queue.people.filter(p => p.status === "done")] };
}
export function complete(queue: Queue, id: string): Queue {
  return { ...queue, call: queue.call?.personId === id ? null : queue.call, people: queue.people.map(p => p.id === id ? { ...p, status: "done" } : p) };
}
export function restore(queue: Queue, id: string): Queue {
  const person = queue.people.find(p => p.id === id && p.status === "done");
  if (!person) return queue;
  if (activePeople(queue).some(p => nameKey(p.name) === nameKey(person.name))) throw new Error("같은 이름이 이미 대기 중입니다.");
  return putLast(queue, { ...person, status: "waiting" });
}
export function queueEntries(queue: Queue) { return activePeople(queue).map((p, i) => `${i + 1}. ${cleanName(p.name)}[${p.paid ? "유" : "무"}]`); }
export function formatQueue(queue: Queue) { return `[상담 대기열]${queueEntries(queue).join("-") || "대기자가 없습니다."}`; }

// For a future chat-send action: split only between complete entries, never truncate.
export function chatMessages(queue: Queue) {
  const prefix = "[상담 대기열]";
  const messages: string[] = [];
  let current = prefix;
  for (const entry of queueEntries(queue)) {
    const next = `${current}${current === prefix ? "" : "-"}${entry}`;
    if (Array.from(next).length <= 200) { current = next; continue; }
    if (current === prefix) throw new Error("대기 명단이 200자 메시지 2개에 담기지 않습니다.");
    messages.push(current); current = prefix + entry;
    if (messages.length >= 2 || Array.from(current).length > 200) throw new Error("대기 명단이 200자 메시지 2개에 담기지 않습니다.");
  }
  messages.push(current === prefix ? formatQueue(queue) : current);
  return messages;
}
