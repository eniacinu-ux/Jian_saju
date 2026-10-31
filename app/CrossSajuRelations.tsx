"use client";
import { useState, type ReactNode } from "react";
import type { GanZhi } from "./lib/sajuCalculator";
import type { InsightPillar } from "./SajuInsights";

interface Group { branches: string[]; label: string }
interface Rules {
  yukhap: string[][]; amhap: string[][]; chung: string[][];
  sanghyeong: string[][]; pa: string[][]; hae: string[][];
  jahyeong: string[]; samhap: Group[]; samhapHalf: Group[];
  banghap: Group[]; samhyeong: Group[];
}
interface Member { side: 0 | 1; label: string; data: GanZhi }
interface Relation { kind: "stem" | "branch"; name: string; members: Member[] }
interface Props {
  left: InsightPillar[]; right: InsightPillar[];
  leftName: string; rightName: string;
  rules: Rules; stemHap: string[][]; stemChung: string[][];
  renderHanja: (data: GanZhi, type: "stem" | "branch") => ReactNode;
}
const pair = (a: string, b: string) => [a, b].sort().join("-");
const stems = ["갑", "을", "병", "정", "무", "기", "경", "신", "임", "계"];
const branches = ["자", "축", "인", "묘", "진", "사", "오", "미", "신", "유", "술", "해"];
const branch = (member: Member) => branches[member.data.branchIndex];

export function crossRelations(left: InsightPillar[], right: InsightPillar[], rules: Rules, stemHap: string[][], stemChung: string[][]): Relation[] {
  const members = (items: InsightPillar[], side: 0 | 1): Member[] => items.flatMap(item => item.data ? [{side, label: item.label, data: item.data}] : []);
  const a = members(left, 0), b = members(right, 1), all = [...a, ...b];
  const result: Relation[] = [];
  const add = (kind: Relation["kind"], name: string, matched: Member[]) => result.push({kind, name, members: matched});
  for (const x of a) for (const y of b) {
    const stemKey = pair(stems[x.data.stemIndex], stems[y.data.stemIndex]);
    if (stemHap.some(([s,t]) => pair(s,t) === stemKey)) add("stem", "천간합", [x,y]);
    if (stemChung.some(([s,t]) => pair(s,t) === stemKey)) add("stem", "천간충", [x,y]);
    const bx = branch(x), by = branch(y), key = pair(bx, by);
    for (const [name, pairs] of [["육합",rules.yukhap],["암합",rules.amhap],["충",rules.chung],["형",rules.sanghyeong],["파",rules.pa],["해",rules.hae]] as const) {
      if (pairs.some(([s,t]) => pair(s,t) === key)) add("branch",name,[x,y]);
    }
    if (bx === by && rules.jahyeong.includes(bx)) add("branch","자형",[x,y]);
    if (bx !== by) {
      for (const group of rules.samhapHalf) if (pair(...group.branches as [string,string]) === key) add("branch",group.label.replace("삼합", "반합"),[x,y]);
      for (const group of rules.banghap) if (group.branches.includes(bx) && group.branches.includes(by)) add("branch",group.label.replace("방합", "방합 2자"),[x,y]);
      for (const group of rules.samhyeong) if (group.branches.includes(bx) && group.branches.includes(by)) add("branch","형(2자)",[x,y]);
    }
  }
  for (const group of [...rules.samhap,...rules.banghap,...rules.samhyeong]) {
    const matches = all.filter(item => group.branches.includes(branch(item)));
    // Three distinct branches, with a contribution from each person.
    for (let i=0;i<matches.length;i++) for(let j=i+1;j<matches.length;j++) for(let k=j+1;k<matches.length;k++) {
      const triple=[matches[i],matches[j],matches[k]];
      if (new Set(triple.map(branch)).size === 3 && new Set(triple.map(item=>item.side)).size === 2) add("branch",`${group.label} 3자 구성`,triple);
    }
  }
  return result;
}

export default function CrossSajuRelations({left,right,leftName,rightName,rules,stemHap,stemChung,renderHanja}: Props) {
  const [filter, setFilter] = useState<"all" | "stem" | "branch">("branch");
  const [expanded, setExpanded] = useState<Set<string>>(() => new Set());
  const relations=crossRelations(left,right,rules,stemHap,stemChung);
  const groups = new Map<string, Relation[]>();
  relations.filter(row => filter === "all" || row.kind === filter).forEach(row => {
    const key = `${row.kind}:${row.name}`;
    const group = groups.get(key) ?? [];
    group.push(row);
    groups.set(key, group);
  });
  const toggle = (key: string) => setExpanded(current => {
    const next = new Set(current);
    if (next.has(key)) next.delete(key); else next.add(key);
    return next;
  });
  const membersCell = (relation: Relation, side: 0 | 1) => <td className="cross-person-cell">{relation.members.filter(member => member.side === side).map(member => <span key={member.label} className="cross-member"><span className="insight-label text-zinc-500">{member.label}</span><span className="insight-ganji inline-flex">{renderHanja(member.data,relation.kind)}</span></span>)}</td>;
  return <section aria-label="두 사람 사이 합충" className="saju-insights cross-relations mt-4 overflow-hidden rounded-xl border border-[#e9e3da] bg-white">
    <div className="flex flex-wrap items-center justify-between gap-3 px-4 py-3"><h3 className="insight-title font-bold text-[#513b2a]">두 사람 사이 합충</h3><div className="flex gap-1" aria-label="합충 종류">{([['all','전체'],['stem','천간'],['branch','지지']] as const).map(([value,label]) => <button key={value} type="button" aria-pressed={filter === value} onClick={() => setFilter(value)} className={`insight-label rounded-md px-3 py-1 font-bold focus-visible:outline-2 ${filter === value ? 'bg-[#eae1d4] text-[#513b2a]' : 'text-zinc-400 hover:bg-zinc-100'}`}>{label}</button>)}</div></div>
    <div className="cross-polarities grid">{([true, false] as const).map(positive => {
      const sectionGroups = Array.from(groups).filter(([,rows]) => rows[0].name.includes("합") === positive);
      return <div key={String(positive)} className={`cross-polarity min-w-0 ${positive ? "is-positive" : "is-negative"}`}><h4 className="cross-polarity-title insight-title px-4 py-2 font-bold">{positive ? "긍정 · 합" : "부정 · 충·형·파·해"}</h4><div className="overflow-x-auto"><table className="cross-table w-full table-fixed border-collapse text-left"><colgroup><col className="w-[23%]" /><col className="w-[38.5%]" /><col className="w-[38.5%]" /></colgroup><thead><tr className="bg-[#f8f5ef]"><th scope="col" className="insight-label px-4 py-2 font-bold text-[#513b2a]">관계</th><th scope="col" className="insight-label px-4 py-2 font-bold text-[#513b2a]">{leftName}</th><th scope="col" className="insight-label px-4 py-2 font-bold text-[#513b2a]">{rightName}</th></tr></thead>
      {sectionGroups.map(([key,rows]) => {
        const isOpen=expanded.has(key);
        const visible=isOpen ? rows : rows.slice(0,3);
        return <tbody key={key} className="cross-group">{visible.map((relation,index) => <tr key={index}>{index === 0 ? <th scope="rowgroup" rowSpan={visible.length} className="cross-group-name insight-body"><span>{relation.name}</span><span className="insight-note ml-2 text-zinc-400">{rows.length}</span></th> : null}{membersCell(relation,0)}{membersCell(relation,1)}</tr>)}{rows.length > 3 ? <tr><td colSpan={3} className="px-4 py-1.5"><button type="button" aria-expanded={isOpen} onClick={()=>toggle(key)} className="insight-note font-bold text-zinc-500 hover:text-[#513b2a]">{rows[0].name} {isOpen ? '접기' : `${rows.length-3}개 더 보기`}</button></td></tr> : null}</tbody>;
      })}
      {!sectionGroups.length ? <tbody><tr><td colSpan={3} className="insight-body px-4 py-4 text-zinc-400">해당 관계 없음</td></tr></tbody> : null}
    </table></div></div>;
    })}</div>
    <details className="border-t border-[#e9e3da] bg-[#fdfcf9] px-4 py-2 text-zinc-500"><summary className="insight-note w-fit cursor-pointer font-bold">표시 기준</summary><p className="insight-note mt-2">두 사람 사이의 조합만 표시합니다. 긍정은 합, 부정은 충·형·파·해를 분류한 것이며 궁합 전체의 좋고 나쁨을 단정하지 않습니다. 같은 기둥에 합과 파 등이 함께 있으면 양쪽 모두 표시합니다. 방합 2자와 3자 구성은 구분합니다. 합화 성립을 판정하는 표는 아닙니다.{!left.find(item=>item.label==='시주')?.data || !right.find(item=>item.label==='시주')?.data ? ' 시간 미상인 시주는 제외합니다.' : ''}</p></details>
  </section>;
}