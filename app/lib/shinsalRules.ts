// Adopted table: 甲乙午 丙寅 丁未 戊己辰 庚戌 辛酉 壬子 癸申.
// Alternative traditions differ, notably for 乙 and 壬.
export const HONGYEOM_RULE_NOTE = "홍염: 일간 기준 · 갑·을→오, 병→인, 정→미, 무·기→진, 경→술, 신→유, 임→자, 계→신";
const stems = ["갑", "을", "병", "정", "무", "기", "경", "신", "임", "계"];
const stemHanja = ["甲", "乙", "丙", "丁", "戊", "己", "庚", "辛", "壬", "癸"];
const branches = ["자", "축", "인", "묘", "진", "사", "오", "미", "신", "유", "술", "해"];
const branchHanja = ["子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥"];
const targets = ["오", "오", "인", "미", "진", "진", "술", "유", "자", "신"];

export function isHongyeom(dayStem: string, targetBranch: string): boolean {
  const stemIndex = stems.indexOf(dayStem.trim());
  const index = stemIndex >= 0 ? stemIndex : stemHanja.indexOf(dayStem.trim());
  const branchIndex = branchHanja.indexOf(targetBranch.trim());
  const branch = branchIndex >= 0 ? branches[branchIndex] : targetBranch.trim();
  return index >= 0 && targets[index] === branch;
}
