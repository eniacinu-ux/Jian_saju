// sajuCalculator.ts

export type Gender = "male" | "female";
export type CalendarType = "solar" | "lunar";

export interface TwelveStages {
  year: string;
  month: string;
  day: string;
  hour: string;
}

export interface SajuInput {
  birthDate: string;
  birthTime: string;
  calendarType?: CalendarType;
  gender?: Gender;
  timezone?: "Asia/Seoul";
  lateZiMode?: boolean;
}

export interface ElementCount {
  wood: number;
  fire: number;
  earth: number;
  metal: number;
  water: number;
}

export interface TenGods {
  yearStem: string;
  monthStem: string;
  dayStem: string;
  hourStem: string;

  yearBranch: string;
  monthBranch: string;
  dayBranch: string;
  hourBranch: string;
}

export interface GanZhi {
  stem: string;
  branch: string;

  stemKor: string;
  branchKor: string;

  ganji: string;
  ganjiKor: string;

  stemIndex: number;
  branchIndex: number;

  stemElement: string;
  branchElement: string;
}
export interface DaewoonItem {
  index: number;
  startAge: number;
  startAgeText: string;
  ganji: GanZhi;
  stemTenGod: string;
  branchTenGod: string;
}

export interface SajuResult {
  year: GanZhi;
  month: GanZhi;
  day: GanZhi;
  hour: GanZhi;

  elementCount: ElementCount;
  tenGods: TenGods;
  twelveStages: TwelveStages;

  daewoon: DaewoonItem[];
}

const STEMS = ["甲", "乙", "丙", "丁", "戊", "己", "庚", "辛", "壬", "癸"];

const TWELVE_STAGE_NAMES = [
  "장생",
  "목욕",
  "관대",
  "건록",
  "제왕",
  "쇠",
  "병",
  "사",
  "묘",
  "절",
  "태",
  "양",
];

const TWELVE_STAGE_START_BRANCH = [
  11, // 甲: 亥 장생
  6,  // 乙: 午 장생
  2,  // 丙: 寅 장생
  9,  // 丁: 酉 장생
  2,  // 戊: 寅 장생
  9,  // 己: 酉 장생
  5,  // 庚: 巳 장생
  0,  // 辛: 子 장생
  8,  // 壬: 申 장생
  3,  // 癸: 卯 장생
];

function getTwelveStage(
  dayStemIndex: number,
  targetBranchIndex: number
): string {
  const startBranchIndex = TWELVE_STAGE_START_BRANCH[dayStemIndex];

  const isYangStem = dayStemIndex % 2 === 0;

  const offset = isYangStem
    ? mod(targetBranchIndex - startBranchIndex, 12)
    : mod(startBranchIndex - targetBranchIndex, 12);

  return TWELVE_STAGE_NAMES[offset];
}

function getTwelveStages(
  year: GanZhi,
  month: GanZhi,
  day: GanZhi,
  hour: GanZhi
): TwelveStages {
  const dayStemIndex = day.stemIndex;

  return {
    year: getTwelveStage(dayStemIndex, year.branchIndex),
    month: getTwelveStage(dayStemIndex, month.branchIndex),
    day: getTwelveStage(dayStemIndex, day.branchIndex),
    hour: getTwelveStage(dayStemIndex, hour.branchIndex),
  };
}

const BRANCHES = [
  "子",
  "丑",
  "寅",
  "卯",
  "辰",
  "巳",
  "午",
  "未",
  "申",
  "酉",
  "戌",
  "亥",
];

const STEMS_KOR = [
  "갑",
  "을",
  "병",
  "정",
  "무",
  "기",
  "경",
  "신",
  "임",
  "계",
];

const BRANCHES_KOR = [
  "자",
  "축",
  "인",
  "묘",
  "진",
  "사",
  "오",
  "미",
  "신",
  "유",
  "술",
  "해",
];

const STEM_ELEMENTS = [
  "목",
  "목",
  "화",
  "화",
  "토",
  "토",
  "금",
  "금",
  "수",
  "수",
];

const BRANCH_ELEMENTS = [
  "수",
  "토",
  "목",
  "목",
  "토",
  "화",
  "화",
  "토",
  "금",
  "금",
  "토",
  "수",
];

const BRANCH_MAIN_STEM_INDEX = [
  9, // 子 癸
  5, // 丑 己
  0, // 寅 甲
  1, // 卯 乙
  4, // 辰 戊
  2, // 巳 丙
  3, // 午 丁
  5, // 未 己
  6, // 申 庚
  7, // 酉 辛
  4, // 戌 戊
  8, // 亥 壬
];

function mod(n: number, m: number): number {
  return ((n % m) + m) % m;
}

function parseKoreanDateTime(dateStr: string, timeStr: string): Date {
  const [y, m, d] = dateStr.split("-").map(Number);
  const [hhRaw, mmRaw] = timeStr.split(":").map(Number);

  const hh = Number.isFinite(hhRaw) ? hhRaw : 0;
  const mm = Number.isFinite(mmRaw) ? mmRaw : 0;

  return new Date(y, m - 1, d, hh, mm, 0, 0);
}

function applySolarTimeCorrection(
  date: Date,
  longitude = 127.0
): Date {
  const standardLongitude = 135.0;

  const correctionMinutes =
    (longitude - standardLongitude) * 4;

  const corrected = new Date(date);

  corrected.setMinutes(
    corrected.getMinutes() + correctionMinutes
  );

  return corrected;
}

function makeGanZhi(
  stemIndex: number,
  branchIndex: number
): GanZhi {
  stemIndex = mod(stemIndex, 10);
  branchIndex = mod(branchIndex, 12);

  return {
    stem: STEMS[stemIndex],
    branch: BRANCHES[branchIndex],

    stemKor: STEMS_KOR[stemIndex],
    branchKor: BRANCHES_KOR[branchIndex],

    ganji: `${STEMS[stemIndex]}${BRANCHES[branchIndex]}`,
    ganjiKor: `${STEMS_KOR[stemIndex]}${BRANCHES_KOR[branchIndex]}`,

    stemIndex,
    branchIndex,

    stemElement: STEM_ELEMENTS[stemIndex],
    branchElement: BRANCH_ELEMENTS[branchIndex],
  };
}

function getYearGanZhi(date: Date): GanZhi {
  let year = date.getFullYear();

  const lichun = new Date(year, 1, 4, 0, 0, 0);

  if (date < lichun) {
    year -= 1;
  }

  const stemIndex = mod(year - 4, 10);
  const branchIndex = mod(year - 4, 12);

  return makeGanZhi(stemIndex, branchIndex);
}

function getMonthBranchIndex(date: Date): number {
  const y = date.getFullYear();

  const terms = [
    { date: new Date(y, 0, 6), branchIndex: 1 },
    { date: new Date(y, 1, 4), branchIndex: 2 },
    { date: new Date(y, 2, 6), branchIndex: 3 },
    { date: new Date(y, 3, 5), branchIndex: 4 },
    { date: new Date(y, 4, 6), branchIndex: 5 },
    { date: new Date(y, 5, 6), branchIndex: 6 },
    { date: new Date(y, 6, 7), branchIndex: 7 },
    { date: new Date(y, 7, 8), branchIndex: 8 },
    { date: new Date(y, 8, 8), branchIndex: 9 },
    { date: new Date(y, 9, 8), branchIndex: 10 },
    { date: new Date(y, 10, 7), branchIndex: 11 },
    { date: new Date(y, 11, 7), branchIndex: 0 },
  ];

  let branchIndex = 1;

  for (const term of terms) {
    if (date >= term.date) {
      branchIndex = term.branchIndex;
    }
  }

  return branchIndex;
}

function getMonthGanZhi(
  date: Date,
  yearStemIndex: number
): GanZhi {
  const monthBranchIndex = getMonthBranchIndex(date);

  const tigerStemStartMap = [
    2,
    4,
    6,
    8,
    0,
    2,
    4,
    6,
    8,
    0,
  ];

  const tigerStemIndex =
    tigerStemStartMap[yearStemIndex];

  const offsetFromTiger = mod(
    monthBranchIndex - 2,
    12
  );

  const monthStemIndex = mod(
    tigerStemIndex + offsetFromTiger,
    10
  );

  return makeGanZhi(
    monthStemIndex,
    monthBranchIndex
  );
}

function getDayGanZhi(date: Date): GanZhi {
  const baseDate = new Date(1993, 7, 4);

  const baseStemIndex = 3;
  const baseBranchIndex = 5;

  const targetDate = new Date(
    date.getFullYear(),
    date.getMonth(),
    date.getDate()
  );

  const diffDays = Math.floor(
    (targetDate.getTime() - baseDate.getTime()) /
      86400000
  );

  return makeGanZhi(
    baseStemIndex + diffDays,
    baseBranchIndex + diffDays
  );
}

function getHourBranchIndex(hour: number): number {
  if (hour === 23 || hour === 0) return 0;
  if (hour >= 1 && hour <= 2) return 1;
  if (hour >= 3 && hour <= 4) return 2;
  if (hour >= 5 && hour <= 6) return 3;
  if (hour >= 7 && hour <= 8) return 4;
  if (hour >= 9 && hour <= 10) return 5;
  if (hour >= 11 && hour <= 12) return 6;
  if (hour >= 13 && hour <= 14) return 7;
  if (hour >= 15 && hour <= 16) return 8;
  if (hour >= 17 && hour <= 18) return 9;
  if (hour >= 19 && hour <= 20) return 10;

  return 11;
}

function getHourGanZhi(
  date: Date,
  dayStemIndex: number
): GanZhi {
  const hourBranchIndex = getHourBranchIndex(
    date.getHours()
  );

  const ziStemStartMap = [
    0,
    2,
    4,
    6,
    8,
    0,
    2,
    4,
    6,
    8,
  ];

  const ziStemIndex =
    ziStemStartMap[dayStemIndex];

  const hourStemIndex = mod(
    ziStemIndex + hourBranchIndex,
    10
  );

  return makeGanZhi(
    hourStemIndex,
    hourBranchIndex
  );
}

function countElements(
  items: GanZhi[]
): ElementCount {
  const count: ElementCount = {
    wood: 0,
    fire: 0,
    earth: 0,
    metal: 0,
    water: 0,
  };

  for (const item of items) {
    const elements = [
      item.stemElement,
      item.branchElement,
    ];

    for (const element of elements) {
      if (element === "목") count.wood += 1;
      if (element === "화") count.fire += 1;
      if (element === "토") count.earth += 1;
      if (element === "금") count.metal += 1;
      if (element === "수") count.water += 1;
    }
  }

  return count;
}

function getStemYinYang(
  stemIndex: number
): "yang" | "yin" {
  return stemIndex % 2 === 0
    ? "yang"
    : "yin";
}

function getStemElementByIndex(
  stemIndex: number
): string {
  return STEM_ELEMENTS[mod(stemIndex, 10)];
}

function getTenGod(
  dayStemIndex: number,
  targetStemIndex: number
): string {
  if (dayStemIndex === targetStemIndex) {
    return "일간";
  }

  const dayElement =
    getStemElementByIndex(dayStemIndex);

  const targetElement =
    getStemElementByIndex(targetStemIndex);

  const dayYinYang =
    getStemYinYang(dayStemIndex);

  const targetYinYang =
    getStemYinYang(targetStemIndex);

  const sameYinYang =
    dayYinYang === targetYinYang;

  // 비겁
  if (dayElement === targetElement) {
    return sameYinYang
      ? "비견"
      : "겁재";
  }

  // 식상
  if (
    (dayElement === "목" &&
      targetElement === "화") ||
    (dayElement === "화" &&
      targetElement === "토") ||
    (dayElement === "토" &&
      targetElement === "금") ||
    (dayElement === "금" &&
      targetElement === "수") ||
    (dayElement === "수" &&
      targetElement === "목")
  ) {
    return sameYinYang
      ? "식신"
      : "상관";
  }

  // 재성
  if (
    (dayElement === "목" &&
      targetElement === "토") ||
    (dayElement === "화" &&
      targetElement === "금") ||
    (dayElement === "토" &&
      targetElement === "수") ||
    (dayElement === "금" &&
      targetElement === "목") ||
    (dayElement === "수" &&
      targetElement === "화")
  ) {
    return sameYinYang
      ? "편재"
      : "정재";
  }

  // 관성
  if (
    (dayElement === "목" &&
      targetElement === "금") ||
    (dayElement === "화" &&
      targetElement === "수") ||
    (dayElement === "토" &&
      targetElement === "목") ||
    (dayElement === "금" &&
      targetElement === "화") ||
    (dayElement === "수" &&
      targetElement === "토")
  ) {
    return sameYinYang
      ? "편관"
      : "정관";
  }

  // 인성
  return sameYinYang
    ? "편인"
    : "정인";
}

function getTenGods(
  year: GanZhi,
  month: GanZhi,
  day: GanZhi,
  hour: GanZhi
): TenGods {
  const dayStemIndex = day.stemIndex;

  return {
    yearStem: getTenGod(
      dayStemIndex,
      year.stemIndex
    ),

    monthStem: getTenGod(
      dayStemIndex,
      month.stemIndex
    ),

    dayStem: "일간",

    hourStem: getTenGod(
      dayStemIndex,
      hour.stemIndex
    ),

    yearBranch: getTenGod(
      dayStemIndex,
      BRANCH_MAIN_STEM_INDEX[
        year.branchIndex
      ]
    ),

    monthBranch: getTenGod(
      dayStemIndex,
      BRANCH_MAIN_STEM_INDEX[
        month.branchIndex
      ]
    ),

    dayBranch: getTenGod(
      dayStemIndex,
      BRANCH_MAIN_STEM_INDEX[
        day.branchIndex
      ]
    ),

    hourBranch: getTenGod(
      dayStemIndex,
      BRANCH_MAIN_STEM_INDEX[
        hour.branchIndex
      ]
    ),
  };
}
function getSolarTermDates(year: number): Date[] {
  return [
    new Date(year, 0, 6),
    new Date(year, 1, 4),
    new Date(year, 2, 6),
    new Date(year, 3, 5),
    new Date(year, 4, 6),
    new Date(year, 5, 6),
    new Date(year, 6, 7),
    new Date(year, 7, 8),
    new Date(year, 8, 8),
    new Date(year, 9, 8),
    new Date(year, 10, 7),
    new Date(year, 11, 7),
  ];
}

function isForwardDaewoon(
  gender: Gender | undefined,
  yearStemIndex: number
): boolean {
  const isYangYear = yearStemIndex % 2 === 0;

  if (gender === "male") return isYangYear;
  if (gender === "female") return !isYangYear;

  return true;
}

function getDaewoonStartAge(
  birthDate: Date,
  forward: boolean
): {
  startAge: number;
  startAgeText: string;
} {
  if (Number.isNaN(birthDate.getTime())) {
    return {
      startAge: 1,
      startAgeText: "1대운",
    };
  }

  const years = [
    birthDate.getFullYear() - 1,
    birthDate.getFullYear(),
    birthDate.getFullYear() + 1,
  ];

  const terms = years
    .flatMap(getSolarTermDates)
    .sort((a, b) => a.getTime() - b.getTime());

  let targetTerm: Date | undefined;

  if (forward) {
    targetTerm = terms.find(
      (term) => term.getTime() > birthDate.getTime()
    );
  } else {
    targetTerm = [...terms]
      .reverse()
      .find(
        (term) => term.getTime() < birthDate.getTime()
      );
  }

  if (!targetTerm) {
    targetTerm = forward
      ? terms[terms.length - 1]
      : terms[0];
  }

 const diffDays =
  Math.abs(targetTerm.getTime() - birthDate.getTime()) /
  (1000 * 60 * 60 * 24);

const startAge = Math.max(
  1,
  Math.floor(diffDays / 3)
);
  return {
    startAge,
    startAgeText: `${startAge}대운`,
  };
}

function getDaewoon(
  birthDate: Date,
  gender: Gender | undefined,
  year: GanZhi,
  month: GanZhi,
  day: GanZhi
): DaewoonItem[] {
  const forward = isForwardDaewoon(
    gender,
    year.stemIndex
  );

  const startInfo = getDaewoonStartAge(
    birthDate,
    forward
  );

  const direction = forward ? 1 : -1;

  return Array.from({ length: 10 }, (_, i) => {
    const step = i + 1;

    const ganji = makeGanZhi(
      month.stemIndex + direction * step,
      month.branchIndex + direction * step
    );

    return {
      index: step,
      startAge: startInfo.startAge + i * 10,
startAgeText: `${startInfo.startAge + i * 10}대운`,
      ganji,
      stemTenGod: getTenGod(
        day.stemIndex,
        ganji.stemIndex
      ),
      branchTenGod: getTenGod(
        day.stemIndex,
        BRANCH_MAIN_STEM_INDEX[ganji.branchIndex]
      ),
    };
  });
}
export function calculateSaju(
  input: SajuInput
): SajuResult {
  if (input.calendarType === "lunar") {
    throw new Error(
      "음력 변환은 아직 지원하지 않습니다."
    );
  }

  const date = parseKoreanDateTime(
    input.birthDate,
    input.birthTime
  );

  const correctedTimeDate =
    applySolarTimeCorrection(date, 127.0);

  const dayDate = new Date(date);

  if (
    input.lateZiMode &&
    correctedTimeDate.getHours() === 23
  ) {
    dayDate.setDate(dayDate.getDate() + 1);
  }

  const year = getYearGanZhi(date);

  const month = getMonthGanZhi(
    date,
    year.stemIndex
  );

  const day = getDayGanZhi(dayDate);

  const hour = getHourGanZhi(
    correctedTimeDate,
    day.stemIndex
  );

  const elementCount = countElements([
    year,
    month,
    day,
    hour,
  ]);

  const tenGods = getTenGods(
    year,
    month,
    day,
    hour
  );
  const twelveStages = getTwelveStages(year, month, day, hour);
  const daewoon = getDaewoon(
  date,
  input.gender,
  year,
  month,
  day
);

  return {
  year,
  month,
  day,
  hour,
  elementCount,
  tenGods,
  twelveStages,
  daewoon,
};
}