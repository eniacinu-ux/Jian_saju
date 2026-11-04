"use client";

import { useState, type ReactNode } from "react";
import type { GanZhi } from "./lib/sajuCalculator";
import { HONGYEOM_RULE_NOTE } from "./lib/shinsalRules";

export interface InsightPillar {
  label: string;
  data?: GanZhi | null;
  branchRelations?: string[];
  shinsals?: string[];
}
interface Props {
  items: InsightPillar[];
  luckyStars: { label: string; value: string }[];
  dayElement: string;
  renderHanja: (data: GanZhi, type: "stem" | "branch") => ReactNode;
}
const cycle = ["목", "화", "토", "금", "수"];


export function strengthSummary(items: InsightPillar[], dayElement: string) {
  const index = cycle.indexOf(dayElement);
  if (index < 0) return null;
  const mother = cycle[(index + 4) % 5];
  const rows = items.flatMap(item => {
    if (!item.data) return [];
    const branchWeight = item.label === "월주" ? 30 : 15;
    const stemWeight = item.label === "시주" ? 5 : 10;
    return [
      { label: `${item.label} 지지`, element: item.data.branchElement, weight: branchWeight },
      ...(item.label === "일주" ? [] : [{ label: `${item.label} 천간`, element: item.data.stemElement, weight: stemWeight }]),
    ];
  });
  const total = rows.reduce((sum, row) => sum + row.weight, 0);
  if (!total) return null;
  const support = rows.filter(row => row.element === dayElement || row.element === mother);
  const score = Math.round(support.reduce((sum, row) => sum + row.weight, 0) / total * 100);
  return { score, label: score < 40 ? "신약" : score <= 60 ? "중화" : "신강", support, total };
}

export default function SajuInsights({ items, luckyStars, dayElement, renderHanja }: Props) {
  const [relationType, setRelationType] = useState<"stem" | "branch">("branch");
  const strength = strengthSummary(items, dayElement);
  const unknownHour = !items.find(item => item.label === "시주")?.data;
  const relationships: { indices: number[]; name: string }[] = [];
  if (relationType === "branch") {
    const seen = new Set<string>();
    items.forEach((item, i) => (item.branchRelations ?? []).forEach(text => {
      const other = items.findIndex(target => text.startsWith(`${target.label} `));
      const indices = other >= 0 ? [i, other].sort((a, b) => a - b) : items.flatMap((target, j) => target.branchRelations?.includes(text) ? [j] : []);
      const name = other >= 0 ? text.slice(items[other].label.length + 1) : text;
      const key = `${indices.join("-")}:${name}`;
      if (indices.length < 2 || seen.has(key)) return;
      seen.add(key);
      relationships.push({ indices, name });
    }));
  } else {
    items.forEach((left, i) => items.slice(i + 1).forEach((right, offset) => {
      if (!left.data || !right.data) return;
      const a = left.data.stemIndex, b = right.data.stemIndex;
      const pair = [a, b].sort((x, y) => x - y).join("-");
      const name = Math.abs(a - b) === 5 ? "천간합" : ["0-6", "1-7", "2-8", "3-9"].includes(pair) ? "천간충" : "";
      if (name) relationships.push({ indices: [i, i + offset + 1], name });
    }));
  }

  const badge = (item: InsightPillar, type: "stem" | "branch") =>
    item.data ? renderHanja(item.data, type) : null;
  const shortLabel = (index: number) => items[index].label.replace("주", "");
  const starRows = items.map(item => {
    if (!item.data) return { item, stars: [] as string[] };
    const hits = luckyStars.filter(star => {
      const targets = star.value.split("·");
      return targets.includes(item.data!.stem) || targets.includes(item.data!.branch);
    }).map(star => star.label === "금여록" ? star.label : `${star.label}귀인`);
    return { item, stars: [...new Set([...(item.shinsals ?? []), ...hits])] };
  });
  const helpfulElements = strength ? [
    { label: "비겁", element: dayElement },
    { label: "인성", element: cycle[(cycle.indexOf(dayElement) + 4) % 5] },
  ].map(row => ({ ...row, percent: Math.round(strength.support.filter(item => item.element === row.element).reduce((sum, item) => sum + item.weight, 0) / strength.total * 100) })) : [];
  const heading = "flex min-h-12 flex-wrap items-center justify-between gap-x-3 gap-y-1 py-2 border-b border-[#c5b39f] insight-title font-bold text-[#513b2a]";

  return <div className="saju-insights mt-4 overflow-hidden rounded-xl border border-[#c5b39f] bg-[#fdfcf9]">
    <div className="insight-columns grid">
      <section aria-label="사주관계" className="insight-section min-w-0 px-4 pb-3">
        <div className={heading}><h4>사주관계</h4><div className="flex gap-1">{(["stem", "branch"] as const).map(type => <button key={type} type="button" aria-pressed={relationType === type} onClick={() => setRelationType(type)} className={`rounded px-2 py-1 insight-label transition focus-visible:outline-2 ${relationType === type ? "bg-[#eae1d4] text-[#513b2a]" : "font-normal text-zinc-400 hover:bg-[#f2eee7]"}`}>{type === "stem" ? "천간" : "지지"}</button>)}</div></div>
        {relationships.length ? <>
          <ul className="divide-y divide-[#d8ccbd]">{relationships.slice(0, 5).map((relation, i) => <li key={i} className="flex min-h-10 items-center gap-2 py-2 insight-body"><span className="shrink-0 whitespace-nowrap insight-label text-zinc-600">{relation.indices.map(shortLabel).join("·")}</span><span className="leading-snug text-[#513b2a]">{relation.name}</span></li>)}</ul>
          {relationships.length > 5 ? <details className="mt-2 insight-label text-zinc-600"><summary className="cursor-pointer">관계 {relationships.length - 5}개 더 보기</summary><ul className="mt-1 divide-y divide-[#d8ccbd]">{relationships.slice(5).map((relation, i) => <li key={i} className="py-1.5">{relation.indices.map(shortLabel).join("·")}　{relation.name}</li>)}</ul></details> : null}
        </> : <p className="py-3 insight-score text-zinc-400">해당하는 {relationType === "stem" ? "천간 합·충" : "지지 관계"} 없음</p>}
      </section>

      <section aria-label="신살·길성" className="insight-section min-w-0 px-4 pb-3">
        <div className={heading}><h4>신살·길성</h4><span className="insight-note font-normal text-zinc-600">원국 기준</span></div>
        <div className="divide-y-[3px] divide-[#89664a]">{starRows.map(({ item, stars }) => <div key={item.label} className="grid grid-cols-[max-content_6.5rem_minmax(0,1fr)] items-center gap-2 py-2">
          <span className="insight-label font-semibold text-zinc-600">{item.label}</span>
          <div className="insight-ganji flex items-center justify-center whitespace-nowrap">{badge(item, "stem")}{badge(item, "branch")}</div>
          <p className="insight-score leading-relaxed text-[#513b2a]">{!item.data ? <span className="text-zinc-400">시간 미상</span> : stars.length ? stars.map((star, i) => <span key={star} className="mr-1 inline-block whitespace-nowrap"><span className={star === "도화살" || star === "홍염살" ? "font-semibold text-red-700" : star === "역마살" ? "font-semibold text-blue-700" : ""}>{star}</span>{i < stars.length - 1 && <span aria-hidden="true" className="ml-1 text-zinc-300">·</span>}</span>) : <span className="text-zinc-400">해당 없음</span>}</p>
        </div>)}</div>
      </section>

      <section aria-label="신강·신약" className="insight-section min-w-0 px-4 pb-3">
        <div className={heading}><h4>신강·신약</h4><span className="insight-note font-normal text-zinc-600">간이 지표</span></div>
        {strength ? <>
          <div className="flex items-baseline justify-between gap-2 py-3"><strong className="insight-score text-[#513b2a]">{strength.label}</strong><span className="insight-score font-semibold tabular-nums text-[#513b2a]">{strength.score}<span className="ml-1 insight-note font-normal text-zinc-600">/ 100</span></span></div>
          <div role="meter" aria-label="일간을 돕는 오행의 가중 비율" aria-valuemin={0} aria-valuemax={100} aria-valuenow={strength.score} className="relative h-2 overflow-hidden rounded-full bg-[#e9e4db]"><div className="h-full rounded-full bg-[#a88461]" style={{ width: `${strength.score}%` }} /><span className="absolute inset-y-0 left-[40%] w-px bg-white" /><span className="absolute inset-y-0 left-[60%] w-px bg-white" /></div>
          <div className="mt-1 flex justify-between insight-note text-zinc-600"><span>신약</span><span>중화</span><span>신강</span></div>
          <dl className="mt-3 divide-y divide-[#d8ccbd] insight-body">{helpfulElements.map(row => <div key={row.label} className="flex justify-between py-2"><dt className="text-zinc-500">{row.label} <span className="text-zinc-400">({row.element})</span></dt><dd className="tabular-nums text-[#513b2a]">{row.percent}%</dd></div>)}<div className="flex justify-between py-2"><dt className="text-zinc-500">도움이 되는 오행</dt><dd className="font-semibold tabular-nums text-[#513b2a]">{strength.score}%</dd></div></dl>
        </> : <p className="py-3 insight-score text-zinc-400">일간 정보 없음</p>}
      </section>
    </div>
    <div className="flex flex-wrap items-start gap-x-5 gap-y-1 border-t border-[#e9e3da] bg-[#f8f5ef] px-4 py-2 insight-label text-zinc-600">
      <details className="min-w-0 flex-1 basis-64"><summary className="w-fit cursor-pointer">신살 판정 기준</summary><p className="mt-2 leading-relaxed">도화는 년지·일지 중 하나라도 성립하면 표시합니다. 그 외 12신살은 일지 기준입니다. {HONGYEOM_RULE_NOTE}. 홍염의 대응표는 학파에 따라 다를 수 있습니다. 길성은 대상 글자가 원국에 있을 때 표시합니다.</p></details>
      <details className="min-w-0 flex-1 basis-64"><summary className="w-fit cursor-pointer">신강·신약 계산 기준</summary><p className="mt-2 leading-relaxed">일간을 돕는 비겁·인성의 가중 비율입니다. 월지 30, 나머지 지지 각 15, 년간·월간 각 10, 시간 5이며 일간 자체는 제외합니다. 시주 미상은 남은 가중치로 환산합니다. 신약 0~39, 중화 40~60, 신강 61~100. 지지 본기만 사용하며 조후·지장간·합화·종격은 반영하지 않습니다.</p></details>
      {unknownHour ? <span>시간 미상 · 시주 제외</span> : null}
    </div>
  </div>;
}