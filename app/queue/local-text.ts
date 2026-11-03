import { cleanName, type Person, type Queue } from "./rules";

/** Dropbox text format (UTF-8): one numbered guest per line, or a single '[대기열] ...' line. */
const sectionPrefix = /^\s*\[\s*(?:상담\s*)?대기열\s*\]\s*/i;
const numbered = /^\d+\s*[.)]\s*/;
const guestCategory = /\s*(?:\((유료|무료|소개팅|멤버십|맴버쉽|회원)\)|\[(유|무)\])\s*$/;
const categoryPaid: Record<string, boolean> = { 유료: true, 무료: false, 소개팅: false, 멤버십: true, 맴버쉽: true, 회원: true, 유: true, 무: false };

export const EMPTY_QUEUE: Queue = { version: 1, people: [], call: null };

/** Stable ID across file updates, even after a browser refresh. */
function stableId(name: string): string {
  // IDs are only used within this UI; collisions are additionally disambiguated by name.
  return `txt:${encodeURIComponent(name.toLocaleLowerCase("ko-KR"))}`;
}

export function parseQueueTxt(input: string): Queue {
  let text = input.replace(/^\uFEFF/, "").replace(/^\u200B/, "").trim();
  text = text.replace(sectionPrefix, "").trim();
  text = text.replace(/\s*\(유료는\s*우선순위\)\s*$/, "").trim();
  if (!text || /^대기자(?:가)?\s*없습니다\.?$/.test(text)) return EMPTY_QUEUE;

  const entries = text.split(/\r?\n|\s+-\s+(?=\d+\s*[.)])/g).map(s => s.trim()).filter(Boolean);
  if (entries.length > 5000) throw new Error("TXT 대기 인원은 최대 5,000명까지 읽을 수 있습니다.");
  const seen = new Set<string>();
  const people: Person[] = [];
  let passedGeneral = false;
  for (let entry of entries) {
    if (entry.startsWith("#")) continue;
    entry = entry.replace(numbered, "").trim();
    if (!entry) continue;
    const match = entry.match(guestCategory);
    const kindLabel = match ? (match[1] || match[2]) : "무료";
    const paid = categoryPaid[kindLabel] ?? false;
    const name = cleanName(match ? entry.slice(0, -match[0].length) : entry);
    const key = name.toLocaleLowerCase("ko-KR");
    if (!name || name.length > 60 || !key) throw new Error("대기열 TXT에 잘못된 이름이 있습니다.");
    if (seen.has(key)) throw new Error(`대기열 TXT에 중복된 이름이 있습니다: ${name}`);
    seen.add(key);
    if (!paid) passedGeneral = true;
    // Preserve a manually lowered paid guest rather than moving them above all free guests.
    people.push({ id: stableId(name), name, paid, freeOrder: paid && passedGeneral, status: "waiting", kindLabel });
  }
  return { version: 1, people, call: null };
}

export function formatQueueTxt(queue: Queue): string {
  return queue.people.filter(p => p.status !== "done").map((person, index) => {
    const lastKind = person.kindLabel;
    const compatible = lastKind && categoryPaid[lastKind] === !!person.paid;
    const label = compatible ? lastKind : (person.paid ? "유료" : "무료");
    return `${index + 1}. ${cleanName(person.name)} (${label})`;
  }).join("\n") + (queue.people.some(p => p.status !== "done") ? "\n" : "");
}

export const FILES = { queue: "queue.txt", chat: "call_chat.txt", payment: "call_payment.txt" } as const;
export type CallKind = "chat" | "payment";
