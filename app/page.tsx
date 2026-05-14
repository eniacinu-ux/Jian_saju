"use client";

import KoreanLunarCalendar from "korean-lunar-calendar";
import { useRef, useState } from "react";
import { saveAs } from "file-saver";
import {
  Document,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from "docx";

import { calculateSaju } from "./lib/sajuCalculator";

export default function Home() {
  const [lunarToSolarDate, setLunarToSolarDate] = useState("");
  const [lunarToSolarIsLeapMonth, setLunarToSolarIsLeapMonth] = useState(false);
  const [lunarToSolarResult, setLunarToSolarResult] = useState("");


  const [solarToLunarDate, setSolarToLunarDate] = useState("");
  const resultRef = useRef<HTMLDivElement>(null);

  const [mode, setMode] = useState<"saju" | "compatibility" | "zodiac">("saju");

  const [form, setForm] = useState({
    name: "",
    gender: "남성",
    birthDate: "",
    birthTime: "23:00",
    birthTimeUnknown: false,
    birthLocation: "",
    calendarType: "solar",
    isLeapMonth: false,
  });

  const [compatibilityForm, setCompatibilityForm] = useState({
    left: {
      name: "",
      gender: "남성",
      birthDate: "",
      birthTime: "23:00",
      birthTimeUnknown: false,
      calendarType: "solar",
      isLeapMonth: false,
    },
    right: {
      name: "",
      gender: "여성",
      birthDate: "",
      birthTime: "23:00",
      birthTimeUnknown: false,
      calendarType: "solar",
      isLeapMonth: false,
    },
  });
  const [selectedDaewoonKey, setSelectedDaewoonKey] = useState<Record<string, string | null>>({});

  const [selectedYearLuckKey, setSelectedYearLuckKey] = useState<Record<string, string | null>>({});
  const [showCompatibilityRelations, setShowCompatibilityRelations] =
    useState(false);
  const [result, setResult] = useState("");
  const [loading, setLoading] = useState(false);
  const [showSaju, setShowSaju] = useState(false);
  const [sajuResult, setSajuResult] = useState<any>(null);
  const [compatibilityResult, setCompatibilityResult] = useState<any>({
    left: null,
    right: null,
  });

  const formatDateInput = (value: string) => {
    const onlyNumber = value.replace(/\D/g, "").slice(0, 8);

    if (onlyNumber.length > 6) {
      return `${onlyNumber.slice(0, 4)}-${onlyNumber.slice(
        4,
        6,
      )}-${onlyNumber.slice(6)}`;
    }

    if (onlyNumber.length > 4) {
      return `${onlyNumber.slice(0, 4)}-${onlyNumber.slice(4)}`;
    }

    return onlyNumber;
  };

  const formatTimeInput = (value: string) => {
    const onlyNumber = value.replace(/\D/g, "").slice(0, 4);

    if (onlyNumber.length >= 3) {
      return `${onlyNumber.slice(0, 2)}:${onlyNumber.slice(2)}`;
    }

    return onlyNumber;
  };

  const convertSolarToLunar = (solarDate: string) => {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(solarDate)) return "";

    const [year, month, day] = solarDate.split("-").map(Number);

    const calendar = new KoreanLunarCalendar();
    const success = calendar.setSolarDate(year, month, day);

    if (!success) return "변환 불가";

    const lunar = calendar.getLunarCalendar();

    return `${lunar.year}-${String(lunar.month).padStart(2, "0")}-${String(
      lunar.day,
    ).padStart(2, "0")}${lunar.intercalation ? " 윤달" : ""}`;
  };
  const convertLunarToSolar = () => {
    const calendar: any = new KoreanLunarCalendar();

    if (!/^\d{4}-\d{2}-\d{2}$/.test(lunarToSolarDate)) {
      setLunarToSolarResult("");
      return;
    }

    const [year, month, day] = lunarToSolarDate.split("-").map(Number);

    const success = calendar.setLunarDate(
      year,
      month,
      day,
      lunarToSolarIsLeapMonth,

    );

    if (!success) {
      setLunarToSolarResult("변환 불가");
      return;
    }

    const solar = calendar.getSolarCalendar();

    setLunarToSolarResult(
      `${solar.year}-${String(solar.month).padStart(2, "0")}-${String(
        solar.day,
      ).padStart(2, "0")}`,
    );
  };
  const countElementsWithoutHour = (calculated: any) => {
    const count = {
      wood: 0,
      fire: 0,
      earth: 0,
      metal: 0,
      water: 0,
    };

    const addElement = (element: string) => {
      if (element === "목") count.wood += 1;
      if (element === "화") count.fire += 1;
      if (element === "토") count.earth += 1;
      if (element === "금") count.metal += 1;
      if (element === "수") count.water += 1;
    };

    [calculated.year, calculated.month, calculated.day].forEach((pillar) => {
      addElement(pillar.stemElement);
      addElement(pillar.branchElement);
    });

    return count;
  };

  const makeTimeUnknownSaju = (calculated: any) => {
    if (!calculated) return null;

    return {
      ...calculated,
      hour: null,
      elementCount: countElementsWithoutHour(calculated),
      tenGods: {
        ...calculated.tenGods,
        hourStem: "",
        hourBranch: "",
      },
      twelveStages: {
        ...calculated.twelveStages,
        hour: "",
      },
      birthTimeUnknown: true,
    };
  };
  const calculateOneSaju = (targetForm: any) => {
    if (!targetForm.birthDate) return null;

    if (!targetForm.birthTimeUnknown && !targetForm.birthTime) return null;

    let solarBirthDate = targetForm.birthDate;

    if (targetForm.calendarType === "lunar") {
      const calendar: any = new KoreanLunarCalendar();

      const [year, month, day] = targetForm.birthDate.split("-").map(Number);

      const success = calendar.setLunarDate(
        year,
        month,
        day,
        targetForm.isLeapMonth || false,
      );

      if (!success) return null;

      const solar = calendar.getSolarCalendar();

      solarBirthDate = `${solar.year}-${String(solar.month).padStart(
        2,
        "0",
      )}-${String(solar.day).padStart(2, "0")}`;
    }

    const calculated = calculateSaju({
      birthDate: solarBirthDate,
      birthTime: targetForm.birthTimeUnknown ? "12:00" : targetForm.birthTime,
      calendarType: "solar",
      timezone: "Asia/Seoul",
      lateZiMode: false,
      gender: targetForm.gender === "남성" ? "male" : "female",
    });

    if (targetForm.birthTimeUnknown) {
      return makeTimeUnknownSaju(calculated);
    }

    return calculated;
  };

  const handleCalculateSaju = () => {
    const calculated = calculateOneSaju(form);
    if (!calculated) return;

    setSajuResult(calculated);
  };

  const handleCalculateCompatibility = () => {
    const left = calculateOneSaju(compatibilityForm.left);
    const right = calculateOneSaju(compatibilityForm.right);

    setCompatibilityResult({
      left,
      right,
    });
  };

  async function handleSubmit() {
    setLoading(true);
    setResult("요청 보내는 중...");

    try {
      let endpoint = "/api/saju";
      let bodyData: any = {
        ...form,
        birthTime: form.birthTimeUnknown ? null : form.birthTime,
        mode,
        saju: sajuResult,
      };

      if (mode === "zodiac") {
        endpoint = "/api/zodiac";
      }

      if (mode === "compatibility") {
        endpoint = "/api/compatibility";

        const left =
          compatibilityResult.left || calculateOneSaju(compatibilityForm.left);

        const right =
          compatibilityResult.right ||
          calculateOneSaju(compatibilityForm.right);

        setCompatibilityResult({
          left,
          right,
        });

        bodyData = {
          mode,
          left: {
            ...compatibilityForm.left,
            birthTime: compatibilityForm.left.birthTimeUnknown
              ? null
              : compatibilityForm.left.birthTime,
          },
          right: {
            ...compatibilityForm.right,
            birthTime: compatibilityForm.right.birthTimeUnknown
              ? null
              : compatibilityForm.right.birthTime,
          },
          leftSaju: left,
          rightSaju: right,
        };
      }

      const res = await fetch(endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(bodyData),
      });

      const data = await res.json();

      if (mode === "zodiac") {
        setResult(JSON.stringify(data, null, 2));
      } else {
        setResult(data.result || "결과가 비어 있습니다.");
      }
    } catch (error) {
      console.error(error);
      setResult("오류가 발생했습니다. 콘솔을 확인하세요.");
    } finally {
      setLoading(false);
    }
  }

  const downloadWord = async () => {
    if (!sajuResult) {
      alert("먼저 만세력을 계산해 주세요.");
      return;
    }

    try {
      const doc = new Document({
        sections: [
          {
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "사주 분석 결과",
                    bold: true,
                    size: 36,
                  }),
                ],
              }),

              new Paragraph({ text: "" }),

              new Paragraph({
                children: [
                  new TextRun({
                    text: `이름: ${form.name}`,
                    size: 24,
                  }),
                ],
              }),

              new Paragraph({
                children: [
                  new TextRun({
                    text: `생년월일: ${form.birthDate} ${form.birthTimeUnknown ? "시간 미상" : form.birthTime
                      }`,
                    size: 24,
                  }),
                ],
              }),

              new Paragraph({ text: "" }),

              new Paragraph({
                children: [
                  new TextRun({
                    text: "만세력",
                    bold: true,
                    size: 28,
                  }),
                ],
              }),

              new Paragraph({ text: "" }),

              new Table({
                width: {
                  size: 100,
                  type: WidthType.PERCENTAGE,
                },
                rows: [
                  new TableRow({
                    children: [
                      new TableCell({
                        children: [new Paragraph("시주")],
                      }),
                      new TableCell({
                        children: [new Paragraph("일주")],
                      }),
                      new TableCell({
                        children: [new Paragraph("월주")],
                      }),
                      new TableCell({
                        children: [new Paragraph("년주")],
                      }),
                    ],
                  }),

                  new TableRow({
                    children: [
                      new TableCell({
                        children: [
                          new Paragraph(
                            sajuResult.hour
                              ? `${sajuResult.hour.ganji} (${sajuResult.hour.ganjiKor})`
                              : "시간 미상",
                          ),
                        ],
                      }),

                      new TableCell({
                        children: [
                          new Paragraph(
                            `${sajuResult.day.ganji} (${sajuResult.day.ganjiKor})`,
                          ),
                        ],
                      }),

                      new TableCell({
                        children: [
                          new Paragraph(
                            `${sajuResult.month.ganji} (${sajuResult.month.ganjiKor})`,
                          ),
                        ],
                      }),

                      new TableCell({
                        children: [
                          new Paragraph(
                            `${sajuResult.year.ganji} (${sajuResult.year.ganjiKor})`,
                          ),
                        ],
                      }),
                    ],
                  }),
                ],
              }),

              new Paragraph({ text: "" }),

              new Paragraph({
                children: [
                  new TextRun({
                    text: "오행 분포",
                    bold: true,
                    size: 28,
                  }),
                ],
              }),

              new Paragraph({
                text:
                  `목 ${sajuResult.elementCount.wood} / ` +
                  `화 ${sajuResult.elementCount.fire} / ` +
                  `토 ${sajuResult.elementCount.earth} / ` +
                  `금 ${sajuResult.elementCount.metal} / ` +
                  `수 ${sajuResult.elementCount.water}`,
              }),

              new Paragraph({
                text: `공망(일주 기준): ${getDayGongmang(sajuResult)}`,
              }),

              ...getGwiyinList(sajuResult).map(
                (gwiyin) =>
                  new Paragraph({
                    text: `${gwiyin.label}: ${gwiyin.value}`,
                  }),
              ),

              new Paragraph({ text: "" }),

              new Paragraph({
                children: [
                  new TextRun({
                    text: "사주 풀이",
                    bold: true,
                    size: 28,
                  }),
                ],
              }),

              ...(result
                ? result.split("\n").map(
                  (line) =>
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: line,
                          size: 24,
                        }),
                      ],
                      spacing: {
                        after: 120,
                      },
                    }),
                )
                : [
                  new Paragraph({
                    text: "사주 풀이 결과가 없습니다.",
                  }),
                ]),
            ],
          },
        ],
      });

      const blob = await Packer.toBlob(doc);

      const safeName = form.name || "이름없음";
      const safeBirth = form.birthDate.replaceAll("-", "");

      saveAs(blob, `${safeName}_${safeBirth}.docx`);
    } catch (error) {
      console.error(error);
      alert("워드 파일 생성 중 오류가 발생했습니다.");
    }
  };

  const getElementColor = (element: string) => {
    switch (element) {
      case "목":
        return "#16a34a";

      case "화":
        return "#dc2626";

      case "토":
        return "#eab308";

      case "금":
        return "#ffffff";

      case "수":
        return "#000000";

      default:
        return "#ffffff";
    }
  };

  const STEMS = ["갑", "을", "병", "정", "무", "기", "경", "신", "임", "계"];
  const BRANCHES = [
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
  const STEM_HANJA: any = {
    갑: "甲",
    을: "乙",
    병: "丙",
    정: "丁",
    무: "戊",
    기: "己",
    경: "庚",
    신: "辛",
    임: "壬",
    계: "癸",
  };

  const BRANCH_HANJA: any = {
    자: "子",
    축: "丑",
    인: "寅",
    묘: "卯",
    진: "辰",
    사: "巳",
    오: "午",
    미: "未",
    신: "申",
    유: "酉",
    술: "戌",
    해: "亥",
  };

  const HANJA_TO_STEM: any = {
    甲: "갑",
    乙: "을",
    丙: "병",
    丁: "정",
    戊: "무",
    己: "기",
    庚: "경",
    辛: "신",
    壬: "임",
    癸: "계",
  };

  const HANJA_TO_BRANCH: any = {
    子: "자",
    丑: "축",
    寅: "인",
    卯: "묘",
    辰: "진",
    巳: "사",
    午: "오",
    未: "미",
    申: "신",
    酉: "유",
    戌: "술",
    亥: "해",
  };

  const normalizeStem = (stem: string) => {
    const value = String(stem || "").trim();
    return HANJA_TO_STEM[value] || value;
  };

  const normalizeBranch = (branch: string) => {
    const value = String(branch || "").trim();
    return HANJA_TO_BRANCH[value] || value;
  };

  const getItemBranch = (item: any) => {
    const directBranch = normalizeBranch(item?.data?.branch);
    if (directBranch) return directBranch;

    const ganji = String(item?.data?.ganji || "").trim();
    return normalizeBranch(ganji.slice(1, 2));
  };

  const getDayGanjiParts = (targetSaju: any) => {
    const dayGanji = String(targetSaju?.day?.ganji || "");

    return {
      stem: normalizeStem(targetSaju?.day?.stem || dayGanji.slice(0, 1)),
      branch: normalizeBranch(targetSaju?.day?.branch || dayGanji.slice(1, 2)),
    };
  };

  const getDayGongmang = (targetSaju: any) => {
    const { stem, branch } = getDayGanjiParts(targetSaju);
    const stemIndex = STEMS.indexOf(stem);
    const branchIndex = BRANCHES.indexOf(branch);

    if (stemIndex < 0 || branchIndex < 0) return "-";

    const ganjiIndex = Array.from({ length: 60 }).findIndex((_, index) => {
      return index % 10 === stemIndex && index % 12 === branchIndex;
    });

    if (ganjiIndex < 0) return "-";

    const xunStartIndex = Math.floor(ganjiIndex / 10) * 10;
    const firstEmptyBranch = BRANCHES[(xunStartIndex + 10) % 12];
    const secondEmptyBranch = BRANCHES[(xunStartIndex + 11) % 12];

    return `${BRANCH_HANJA[firstEmptyBranch]}·${BRANCH_HANJA[secondEmptyBranch]}`;
  };

  const formatBranchTargets = (targets: string[]) => {
    if (!targets?.length) return "-";

    return targets
      .map((target: string) => {
        const normalizedBranch = normalizeBranch(target);

        return BRANCH_HANJA[normalizedBranch] || target;
      })
      .join("·");
  };

  const formatStemTargets = (targets: string[]) => {
    if (!targets?.length) return "-";

    return targets
      .map((target: string) => {
        const normalizedStem = normalizeStem(target);

        return STEM_HANJA[normalizedStem] || target;
      })
      .join("·");
  };

  const formatTypedGwiyinTargets = (
    targets: { type: "stem" | "branch"; value: string }[],
  ) => {
    if (!targets?.length) return "-";

    return targets
      .map((target) => {
        if (target.type === "stem") {
          const normalizedStem = normalizeStem(target.value);

          return STEM_HANJA[normalizedStem] || target.value;
        }

        const normalizedBranch = normalizeBranch(target.value);

        return BRANCH_HANJA[normalizedBranch] || target.value;
      })
      .join("·");
  };

  const getMonthBranch = (targetSaju: any) => {
    const monthGanji = String(targetSaju?.month?.ganji || "");

    return normalizeBranch(targetSaju?.month?.branch || monthGanji.slice(1, 2));
  };

  const getCheoneulGwiyin = (targetSaju: any) => {
    const { stem } = getDayGanjiParts(targetSaju);

    const cheoneulMap: any = {
      갑: ["축", "미"],
      무: ["축", "미"],
      경: ["축", "미"],
      을: ["자", "신"],
      기: ["자", "신"],
      병: ["해", "유"],
      정: ["해", "유"],
      임: ["사", "묘"],
      계: ["사", "묘"],
      신: ["오", "인"],
    };

    return formatBranchTargets(cheoneulMap[stem] || []);
  };

  const getHakdangGwiyin = (targetSaju: any) => {
    const { stem } = getDayGanjiParts(targetSaju);

    const hakdangMap: any = {
      갑: ["해"],
      을: ["오"],
      병: ["인"],
      정: ["유"],
      무: ["인"],
      기: ["유"],
      경: ["사"],
      신: ["자"],
      임: ["신"],
      계: ["묘"],
    };

    return formatBranchTargets(hakdangMap[stem] || []);
  };

  const getMungoukGwiyin = (targetSaju: any) => {
    const { stem } = getDayGanjiParts(targetSaju);

    const mungokMap: any = {
      갑: ["해"],
      을: ["해"],
      병: ["인"],
      정: ["유"],
      무: ["인"],
      기: ["유"],
      경: ["사"],
      신: ["자"],
      임: ["신"],
      계: ["묘"],
    };

    return formatBranchTargets(mungokMap[stem] || []);
  };

  const getTaegukGwiyin = (targetSaju: any) => {
    const { stem } = getDayGanjiParts(targetSaju);

    const taegukMap: any = {
      갑: ["자", "오"],
      을: ["자", "오"],
      병: ["묘", "유"],
      정: ["묘", "유"],
      무: ["진", "술", "축", "미"],
      기: ["진", "술", "축", "미"],
      경: ["인", "해"],
      신: ["인", "해"],
      임: ["사", "신"],
      계: ["사", "신"],
    };

    return formatBranchTargets(taegukMap[stem] || []);
  };

  const getCheondeokGwiyin = (targetSaju: any) => {
    const monthBranch = getMonthBranch(targetSaju);

    const cheondeokMap: Record<
      string,
      { type: "stem" | "branch"; value: string }[]
    > = {
      인: [{ type: "stem", value: "정" }],
      묘: [{ type: "stem", value: "신" }],
      진: [{ type: "stem", value: "계" }],
      사: [{ type: "stem", value: "임" }],
      오: [{ type: "stem", value: "신" }],
      미: [{ type: "stem", value: "갑" }],
      신: [{ type: "stem", value: "계" }],
      유: [{ type: "branch", value: "인" }],
      술: [{ type: "stem", value: "병" }],
      해: [{ type: "stem", value: "을" }],
      자: [{ type: "branch", value: "사" }],
      축: [{ type: "stem", value: "경" }],
    };

    return formatTypedGwiyinTargets(cheondeokMap[monthBranch] || []);
  };

  const getWoldeokGwiyin = (targetSaju: any) => {
    const monthBranch = getMonthBranch(targetSaju);

    const woldeokMap: any = {
      인: ["병"],
      묘: ["갑"],
      진: ["임"],
      사: ["경"],
      오: ["병"],
      미: ["갑"],
      신: ["임"],
      유: ["경"],
      술: ["병"],
      해: ["갑"],
      자: ["임"],
      축: ["경"],
    };

    return formatStemTargets(woldeokMap[monthBranch] || []);
  };

  const getCheonjuGwiyin = (targetSaju: any) => {
    const { stem } = getDayGanjiParts(targetSaju);

    const cheonjuMap: any = {
      갑: ["사"],
      을: ["오"],
      병: ["사"],
      정: ["오"],
      무: ["신"],
      기: ["유"],
      경: ["해"],
      신: ["자"],
      임: ["인"],
      계: ["묘"],
    };

    return formatBranchTargets(cheonjuMap[stem] || []);
  };

  const getGeumnyeorok = (targetSaju: any) => {
    const { stem } = getDayGanjiParts(targetSaju);

    const geumnyeoMap: any = {
      갑: ["진"],
      을: ["사"],
      병: ["미"],
      정: ["신"],
      무: ["미"],
      기: ["신"],
      경: ["술"],
      신: ["해"],
      임: ["축"],
      계: ["인"],
    };

    return formatBranchTargets(geumnyeoMap[stem] || []);
  };

  const getBokseongGwiyin = (targetSaju: any) => {
    const { stem } = getDayGanjiParts(targetSaju);

    const bokseongMap: any = {
      갑: ["인"],
      을: ["묘"],
      병: ["술"],
      정: ["해"],
      무: ["신"],
      기: ["미"],
      경: ["오"],
      신: ["사"],
      임: ["진"],
      계: ["축"],
    };

    return formatBranchTargets(bokseongMap[stem] || []);
  };

  const getGukinGwiyin = (targetSaju: any) => {
    const { stem } = getDayGanjiParts(targetSaju);

    const gukinMap: any = {
      갑: ["술"],
      을: ["해"],
      병: ["축"],
      정: ["인"],
      무: ["축"],
      기: ["인"],
      경: ["진"],
      신: ["사"],
      임: ["미"],
      계: ["신"],
    };

    return formatBranchTargets(gukinMap[stem] || []);
  };

  const getGwiyinList = (targetSaju: any) => [
    { label: "천을", value: getCheoneulGwiyin(targetSaju) },
    { label: "학당", value: getHakdangGwiyin(targetSaju) },
    { label: "문곡", value: getMungoukGwiyin(targetSaju) },
    { label: "태극", value: getTaegukGwiyin(targetSaju) },
    { label: "천덕", value: getCheondeokGwiyin(targetSaju) },
    { label: "월덕", value: getWoldeokGwiyin(targetSaju) },
    { label: "천주", value: getCheonjuGwiyin(targetSaju) },
    { label: "금여록", value: getGeumnyeorok(targetSaju) },
    { label: "복성", value: getBokseongGwiyin(targetSaju) },
    { label: "국인", value: getGukinGwiyin(targetSaju) },
  ];

  const renderSpecialInfo = (targetSaju: any) => {
    if (!targetSaju) return null;

    const gwiyinList = getGwiyinList(targetSaju);

    return (
      <div className="mt-5 grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-6">
        <div className="rounded-2xl border border-[#ead8c4] bg-[#fffaf3] p-4 text-center">
          <div className="text-smm font-bold text-zinc-700">
            공망 / <br /> 일주
          </div>

          <div className="mt-2 text-2xl font-bold text-[#6b3f24]">
            {getDayGongmang(targetSaju)}
          </div>
        </div>

        {gwiyinList.map((gwiyin) => (
          <div
            key={gwiyin.label}
            className="rounded-2xl border border-[#ead8c4] bg-[#fffaf3] p-4 text-center"
          >
            <div className="text-smm font-bold text-zinc-700">
              {gwiyin.label}
            </div>

            <div className="mt-2 text-2xl font-bold text-[#6b3f24]">
              {gwiyin.value}
            </div>
          </div>
        ))}
      </div>
    );
  };


  const STEM_INFO: any = {
    갑: { element: "목", yinYang: "양" },
    을: { element: "목", yinYang: "음" },
    병: { element: "화", yinYang: "양" },
    정: { element: "화", yinYang: "음" },
    무: { element: "토", yinYang: "양" },
    기: { element: "토", yinYang: "음" },
    경: { element: "금", yinYang: "양" },
    신: { element: "금", yinYang: "음" },
    임: { element: "수", yinYang: "양" },
    계: { element: "수", yinYang: "음" },
  };

  const BRANCH_MAIN_STEM: any = {
    자: "계",
    축: "기",
    인: "갑",
    묘: "을",
    진: "무",
    사: "병",
    오: "정",
    미: "기",
    신: "경",
    유: "신",
    술: "무",
    해: "임",
  };

  const BRANCH_HIDDEN_STEMS: any = {
    자: ["임", "계"],
    축: ["계", "신", "기"],
    인: ["무", "병", "갑"],
    묘: ["갑", "을"],
    진: ["을", "계", "무"],
    사: ["무", "경", "병"],
    오: ["병", "기", "정"],
    미: ["정", "을", "기"],
    신: ["무", "임", "경"],
    유: ["경", "신"],
    술: ["신", "정", "무"],
    해: ["무", "갑", "임"],
  };

  const getHiddenStemsText = (branch: string) => {
    const normalizedBranch = normalizeBranch(branch);
    const hiddenStems = BRANCH_HIDDEN_STEMS[normalizedBranch] || [];

    if (!hiddenStems.length) return "";

    return hiddenStems.map((stem: string) => STEM_HANJA[stem] ?? "").join("");
  };

  const BRANCH_RELATION_RULES: any = {
    yukhap: [
      ["자", "축"],
      ["인", "해"],
      ["묘", "술"],
      ["진", "유"],
      ["사", "신"],
      ["오", "미"],
    ],
    samhap: [
      { branches: ["해", "묘", "미"], label: "삼합(木)" },
      { branches: ["인", "오", "술"], label: "삼합(火)" },
      { branches: ["사", "유", "축"], label: "삼합(金)" },
      { branches: ["신", "자", "진"], label: "삼합(水)" },
    ],
    banghap: [
      { branches: ["인", "묘", "진"], label: "방합(東)" },
      { branches: ["사", "오", "미"], label: "방합(南)" },
      { branches: ["신", "유", "술"], label: "방합(西)" },
      { branches: ["해", "자", "축"], label: "방합(北)" },
    ],
    amhap: [
      ["자", "사"],
      ["자", "진"],
      ["자", "술"],
      ["인", "축"],
      ["인", "오"],
      ["인", "미"],
      ["묘", "신"],
      ["사", "축"],
    ],
    chung: [
      ["자", "오"],
      ["축", "미"],
      ["인", "신"],
      ["묘", "유"],
      ["진", "술"],
      ["사", "해"],
    ],
    samhyeong: [
      { branches: ["인", "사", "신"], label: "삼형" },
      { branches: ["축", "술", "미"], label: "삼형" },
    ],
    sanghyeong: [["자", "묘"]],
    jahyeong: ["진", "오", "유", "해"],
    pa: [
      ["자", "유"],
      ["축", "진"],
      ["인", "해"],
      ["묘", "오"],
      ["사", "신"],
      ["미", "술"],
    ],
    hae: [
      ["자", "미"],
      ["축", "오"],
      ["인", "사"],
      ["묘", "진"],
      ["신", "해"],
      ["유", "술"],
    ],
    gwimun: [
      ["자", "미"],
      ["축", "오"],
      ["인", "미"],
      ["묘", "신"],
      ["진", "해"],
      ["사", "술"],
    ],
    wonjin: [
      ["자", "미"],
      ["축", "오"],
      ["인", "유"],
      ["묘", "신"],
      ["진", "해"],
      ["사", "술"],
    ],
  };

  const makeBranchPairKey = (a: string, b: string) =>
    [normalizeBranch(a), normalizeBranch(b)].sort().join("-");

  const formatBranchSet = (branches: string[]) =>
    branches.map((branch) => BRANCH_HANJA[branch] || branch).join("");

  const addBranchRelationsToItems = (items: any[]) => {
    const activeItems = items.filter((item) => getItemBranch(item));
    const relationMap = new Map<string, string[]>();

    activeItems.forEach((item) => relationMap.set(item.label, []));

    const addRelation = (labels: string[], text: string) => {
      labels.forEach((label) => {
        const current = relationMap.get(label) || [];
        if (!current.includes(text)) current.push(text);
        relationMap.set(label, current);
      });
    };

    const addPairRelation = (left: any, right: any, name: string) => {
      const leftBranch = getItemBranch(left);
      const rightBranch = getItemBranch(right);
      const pairText = formatBranchSet([leftBranch, rightBranch]);

      addRelation([left.label], `${right.label} ${name}(${pairText})`);
      addRelation([right.label], `${left.label} ${name}(${pairText})`);
    };

    const pairRules = [
      { name: "육합", rules: BRANCH_RELATION_RULES.yukhap },
      { name: "암합", rules: BRANCH_RELATION_RULES.amhap },
      { name: "충", rules: BRANCH_RELATION_RULES.chung },
      { name: "상형", rules: BRANCH_RELATION_RULES.sanghyeong },
      { name: "파", rules: BRANCH_RELATION_RULES.pa },
      { name: "해", rules: BRANCH_RELATION_RULES.hae },
      { name: "귀문", rules: BRANCH_RELATION_RULES.gwimun },
      { name: "원진", rules: BRANCH_RELATION_RULES.wonjin },
    ];

    for (let i = 0; i < activeItems.length; i += 1) {
      for (let j = i + 1; j < activeItems.length; j += 1) {
        const left = activeItems[i];
        const right = activeItems[j];
        const leftBranch = getItemBranch(left);
        const rightBranch = getItemBranch(right);
        const pairKey = makeBranchPairKey(leftBranch, rightBranch);

        pairRules.forEach(({ name, rules }) => {
          const matched = rules.some(
            ([a, b]: string[]) => makeBranchPairKey(a, b) === pairKey,
          );

          if (matched) addPairRelation(left, right, name);
        });

        BRANCH_RELATION_RULES.samhyeong.forEach((rule: any) => {
          if (
            leftBranch !== rightBranch &&
            rule.branches.includes(leftBranch) &&
            rule.branches.includes(rightBranch)
          ) {
            addPairRelation(left, right, "형");
          }
        });

        BRANCH_RELATION_RULES.samhap.forEach((rule: any) => {
          if (
            rule.branches.includes(leftBranch) &&
            rule.branches.includes(rightBranch)
          ) {
            addPairRelation(left, right, rule.label);
          }
        });

        BRANCH_RELATION_RULES.banghap.forEach((rule: any) => {
          if (
            rule.branches.includes(leftBranch) &&
            rule.branches.includes(rightBranch)
          ) {
            addPairRelation(left, right, rule.label);
          }
        });
      }
    }

    BRANCH_RELATION_RULES.samhap.forEach((rule: any) => {
      const matched = activeItems.filter((item) =>
        rule.branches.includes(getItemBranch(item)),
      );
      const matchedBranchCount = new Set(
        matched.map((item) => getItemBranch(item)),
      ).size;

      if (matchedBranchCount === 3) {
        addRelation(
          matched.map((item) => item.label),
          `${rule.label} 완성(${formatBranchSet(rule.branches)})`,
        );
      }
    });

    BRANCH_RELATION_RULES.banghap.forEach((rule: any) => {
      const matched = activeItems.filter((item) =>
        rule.branches.includes(getItemBranch(item)),
      );
      const matchedBranchCount = new Set(
        matched.map((item) => getItemBranch(item)),
      ).size;

      if (matchedBranchCount === 3) {
        addRelation(
          matched.map((item) => item.label),
          `${rule.label} 완성(${formatBranchSet(rule.branches)})`,
        );
      }
    });

    BRANCH_RELATION_RULES.samhyeong.forEach((rule: any) => {
      const matched = activeItems.filter((item) =>
        rule.branches.includes(getItemBranch(item)),
      );
      const matchedBranchCount = new Set(
        matched.map((item) => getItemBranch(item)),
      ).size;

      if (matchedBranchCount === 3) {
        addRelation(
          matched.map((item) => item.label),
          `${rule.label} 완성(${formatBranchSet(rule.branches)})`,
        );
      }
    });

    BRANCH_RELATION_RULES.jahyeong.forEach((branch: string) => {
      const matched = activeItems.filter(
        (item) => getItemBranch(item) === branch,
      );

      if (matched.length >= 2) {
        matched.forEach((item) => {
          const others = matched
            .filter((target: any) => target.label !== item.label)
            .map((target: any) => target.label)
            .join("·");
          addRelation(
            [item.label],
            `${others} 자형(${formatBranchSet([branch, branch])})`,
          );
        });
      }
    });

    return items.map((item) => ({
      ...item,
      branchRelations: relationMap.get(item.label) || [],
    }));
  };

  const SHINSAL_LABELS: any = {
    geopsal: "겁살",
    jaesal: "재살",
    cheonsal: "천살",
    jisal: "지살",
    dosal: "도화살",
    wolsal: "월살",
    mangsinsal: "망신살",
    jangseongsal: "장성살",
    banansal: "반안살",
    yeokmasal: "역마살",
    yukhaesal: "육해살",
    hwagaesal: "화개살",
  };

  const TWELVE_SHINSAL_RULES: any = {
    water: {
      baseBranches: ["신", "자", "진"],
      byBranch: {
        사: "geopsal",
        오: "jaesal",
        미: "cheonsal",
        신: "jisal",
        유: "dosal",
        술: "wolsal",
        해: "mangsinsal",
        자: "jangseongsal",
        축: "banansal",
        인: "yeokmasal",
        묘: "yukhaesal",
        진: "hwagaesal",
      },
    },
    fire: {
      baseBranches: ["인", "오", "술"],
      byBranch: {
        해: "geopsal",
        자: "jaesal",
        축: "cheonsal",
        인: "jisal",
        묘: "dosal",
        진: "wolsal",
        사: "mangsinsal",
        오: "jangseongsal",
        미: "banansal",
        신: "yeokmasal",
        유: "yukhaesal",
        술: "hwagaesal",
      },
    },
    metal: {
      baseBranches: ["사", "유", "축"],
      byBranch: {
        인: "geopsal",
        묘: "jaesal",
        진: "cheonsal",
        사: "jisal",
        오: "dosal",
        미: "wolsal",
        신: "mangsinsal",
        유: "jangseongsal",
        술: "banansal",
        해: "yeokmasal",
        자: "yukhaesal",
        축: "hwagaesal",
      },
    },
    wood: {
      baseBranches: ["해", "묘", "미"],
      byBranch: {
        신: "geopsal",
        유: "jaesal",
        술: "cheonsal",
        해: "jisal",
        자: "dosal",
        축: "wolsal",
        인: "mangsinsal",
        묘: "jangseongsal",
        진: "banansal",
        사: "yeokmasal",
        오: "yukhaesal",
        미: "hwagaesal",
      },
    },
  };

  const HYUNCHIMSAL_STEMS = ["갑", "신"];
  const HYUNCHIMSAL_BRANCHES = ["묘", "오"];

  const getTwelveShinsalBaseRule = (baseBranch: string) => {
    const normalizedBaseBranch = normalizeBranch(baseBranch);

    return Object.values(TWELVE_SHINSAL_RULES).find((rule: any) =>
      rule.baseBranches.includes(normalizedBaseBranch),
    ) as any;
  };

  const getItemStem = (item: any) => {
    const directStem = normalizeStem(item?.data?.stem);
    if (directStem) return directStem;

    const ganji = String(item?.data?.ganji || "").trim();
    return normalizeStem(ganji.slice(0, 1));
  };

  const getBaseBranchForShinsal = (targetSaju: any) => {
    const dayBranch = normalizeBranch(
      targetSaju?.day?.branch ||
      String(targetSaju?.day?.ganji || "").slice(1, 2),
    );

    if (dayBranch) return dayBranch;

    return normalizeBranch(
      targetSaju?.year?.branch ||
      String(targetSaju?.year?.ganji || "").slice(1, 2),
    );
  };

  const addShinsalToItems = (items: any[], targetSaju: any) => {
    const activeItems = items.filter((item) => getItemBranch(item));
    const shinsalMap = new Map<string, string[]>();
    const baseBranch = getBaseBranchForShinsal(targetSaju);
    const baseRule = getTwelveShinsalBaseRule(baseBranch);

    activeItems.forEach((item) => shinsalMap.set(item.label, []));

    const addShinsal = (label: string, text: string) => {
      const current = shinsalMap.get(label) || [];
      if (!current.includes(text)) current.push(text);
      shinsalMap.set(label, current);
    };

    activeItems.forEach((item) => {
      const branch = getItemBranch(item);
      const stem = getItemStem(item);
      const twelveShinsalKey = baseRule?.byBranch?.[branch];

      if (twelveShinsalKey) {
        addShinsal(item.label, SHINSAL_LABELS[twelveShinsalKey]);
      }

      if (
        HYUNCHIMSAL_STEMS.includes(stem) ||
        HYUNCHIMSAL_BRANCHES.includes(branch)
      ) {
        addShinsal(item.label, "현침살");
      }
    });

    for (let i = 0; i < activeItems.length; i += 1) {
      for (let j = i + 1; j < activeItems.length; j += 1) {
        const left = activeItems[i];
        const right = activeItems[j];
        const leftBranch = getItemBranch(left);
        const rightBranch = getItemBranch(right);
        const pairKey = makeBranchPairKey(leftBranch, rightBranch);

        const hasWonjin = BRANCH_RELATION_RULES.wonjin.some(
          ([a, b]: string[]) => makeBranchPairKey(a, b) === pairKey,
        );

        if (hasWonjin) {
          addShinsal(left.label, `${right.label} 원진살`);
          addShinsal(right.label, `${left.label} 원진살`);
        }

        const hasGwimun = BRANCH_RELATION_RULES.gwimun.some(
          ([a, b]: string[]) => makeBranchPairKey(a, b) === pairKey,
        );

        if (hasGwimun) {
          addShinsal(left.label, `${right.label} 귀문관살`);
          addShinsal(right.label, `${left.label} 귀문관살`);
        }
      }
    }

    return items.map((item) => ({
      ...item,
      shinsals: shinsalMap.get(item.label) || [],
    }));
  };

  const ELEMENT_GENERATES: any = {
    목: "화",
    화: "토",
    토: "금",
    금: "수",
    수: "목",
  };

  const ELEMENT_CONTROLS: any = {
    목: "토",
    화: "금",
    토: "수",
    금: "목",
    수: "화",
  };

  const getTenGod = (dayStem: string, targetStem: string) => {
    const normalizedDayStem = normalizeStem(dayStem);
    const normalizedTargetStem = normalizeStem(targetStem);

    const day = STEM_INFO[normalizedDayStem];
    const target = STEM_INFO[normalizedTargetStem];

    if (!day || !target) return "";

    const sameYinYang = day.yinYang === target.yinYang;

    if (day.element === target.element) {
      return sameYinYang ? "비견" : "겁재";
    }

    if (ELEMENT_GENERATES[day.element] === target.element) {
      return sameYinYang ? "식신" : "상관";
    }

    if (ELEMENT_GENERATES[target.element] === day.element) {
      return sameYinYang ? "편인" : "정인";
    }

    if (ELEMENT_CONTROLS[day.element] === target.element) {
      return sameYinYang ? "편재" : "정재";
    }

    if (ELEMENT_CONTROLS[target.element] === day.element) {
      return sameYinYang ? "편관" : "정관";
    }

    return "";
  };

  const getGanjiByYear = (year: number) => {
    const stem = STEMS[(((year - 4) % 10) + 10) % 10];
    const branch = BRANCHES[(((year - 4) % 12) + 12) % 12];

    return {
      stem,
      branch,
      stemElement: STEM_INFO[stem].element,
      branchElement: STEM_INFO[BRANCH_MAIN_STEM[branch]].element,
    };
  };

  const getDaewoonStartAge = (item: any) => {
    if (typeof item.startAge === "number") return item.startAge;

    const matched = String(item.startAgeText || "").match(/\d+/);
    return matched ? Number(matched[0]) : 0;
  };

  const buildYearLuckList = (
    targetSaju: any,
    daewoonItem: any,
    birthDate: string,
  ) => {
    const birthYear = Number(birthDate.slice(0, 4));
    const startAge = getDaewoonStartAge(daewoonItem);
    const dayStem = normalizeStem(
      targetSaju?.day?.stem || String(targetSaju?.day?.ganji || "").slice(0, 1),
    );

    return Array.from({ length: 10 }).map((_, index) => {
      const age = startAge + index;
      const year = birthYear + age - 1;
      const ganji = getGanjiByYear(year);
      const branchMainStem = BRANCH_MAIN_STEM[ganji.branch];

      return {
        year,
        age,
        ganji,
        stemTenGod: getTenGod(dayStem, ganji.stem),
        branchTenGod: getTenGod(dayStem, branchMainStem),
      };
    });
  };

  const MONTH_BRANCHES = [
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
    "자",
    "축",
  ];

  const YEAR_STEM_TO_FIRST_MONTH_STEM: any = {
    갑: "병",
    기: "병",
    을: "무",
    경: "무",
    병: "경",
    신: "경",
    정: "임",
    임: "임",
    무: "갑",
    계: "갑",
  };

  const getGanjiByMonth = (yearGanji: any, monthIndex: number) => {
    const yearStem = normalizeStem(yearGanji?.stem || "");
    const firstMonthStem = YEAR_STEM_TO_FIRST_MONTH_STEM[yearStem] || "병";
    const firstStemIndex = STEMS.indexOf(firstMonthStem);
    const stem = STEMS[(firstStemIndex + monthIndex) % 10];
    const branch = MONTH_BRANCHES[monthIndex];

    return {
      stem,
      branch,
      stemElement: STEM_INFO[stem].element,
      branchElement: STEM_INFO[BRANCH_MAIN_STEM[branch]].element,
    };
  };

  const buildMonthLuckList = (targetSaju: any, yearLuck: any) => {
    const dayStem = normalizeStem(
      targetSaju?.day?.stem || String(targetSaju?.day?.ganji || "").slice(0, 1),
    );

    const monthOrder = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 1];

    return Array.from({ length: 12 }).map((_, index) => {
      const month = monthOrder[index];

      const ganji = getGanjiByMonth(yearLuck?.ganji, index);
      const branchMainStem = BRANCH_MAIN_STEM[ganji.branch];

      return {
        year: yearLuck.year,
        month,
        ganji,
        stemTenGod: getTenGod(dayStem, ganji.stem),
        branchTenGod: getTenGod(dayStem, branchMainStem),
      };
    });
  };
  const buildSajuItems = (targetSaju: any) => {
    if (!targetSaju) return [];

    return addShinsalToItems(
      addBranchRelationsToItems([
        {
          label: "시주",
          data: targetSaju.hour,
          tenGodStem: targetSaju.tenGods?.hourStem ?? "",
          tenGodBranch: targetSaju.tenGods?.hourBranch ?? "",
          twelveStage: targetSaju.twelveStages?.hour ?? "",
        },
        {
          label: "일주",
          data: targetSaju.day,
          tenGodStem: targetSaju.tenGods.dayStem,
          tenGodBranch: targetSaju.tenGods.dayBranch,
          twelveStage: targetSaju.twelveStages.day,
        },
        {
          label: "월주",
          data: targetSaju.month,
          tenGodStem: targetSaju.tenGods.monthStem,
          tenGodBranch: targetSaju.tenGods.monthBranch,
          twelveStage: targetSaju.twelveStages.month,
        },
        {
          label: "년주",
          data: targetSaju.year,
          tenGodStem: targetSaju.tenGods.yearStem,
          tenGodBranch: targetSaju.tenGods.yearBranch,
          twelveStage: targetSaju.twelveStages.year,
        },
      ]),
      targetSaju,
    );
  };

  const sajuItems = buildSajuItems(sajuResult);

  const renderPillarCard = (item: any, _cardKey = "main") => {
    const isCompatibilityMode = mode === "compatibility";

    if (!item.data) {
      return (
        <div key={item.label} className="rounded-xl bg-zinc-100 p-3">
          <div className="text-smm font-bold text-black">{item.label}</div>

          <div className="mt-8 flex min-h-[150px] items-center justify-center rounded-xl border border-dashed border-zinc-300 bg-white/60 text-smm font-bold text-zinc-400">
            시간 미상
          </div>
        </div>
      );
    }

    return (
      <div key={item.label} className="rounded-xl bg-zinc-100 p-3">
        <div className="text-smm font-bold text-black">
          {item.label}</div>

        <div className="mt-2 flex flex-col items-center">
          <div
            className="text-3xl font-bold leading-none"
            style={{
              color: getElementColor(item.data.stemElement),
              WebkitTextStroke: "1.5px black",
              textShadow: "0 0 2px black, 0 0 4px black",
            }}
          >
            {item.data.stem}
          </div>

          <div
            className="mt-2 text-3xl font-bold leading-none"
            style={{
              color: getElementColor(item.data.branchElement),
              WebkitTextStroke: "1.5px black",
              textShadow: "0 0 2px black, 0 0 4px black",
            }}
          >
            {item.data.branch}
          </div>

          <div className="mt-3 text-smm font-semibold text-zinc-600">
            {item.data.stemKor}
            {item.data.branchKor}
          </div>

          <div className="mt-1 rounded-lg bg-white/70 px-2 py-1 text-[15px] font-semibold text-zinc-600">
            {getHiddenStemsText(item.data.branch)}
          </div>
        </div>

        <div className="mt-3 text-sm font-bold text-black">
          {item.tenGodStem}
        </div>

        <div className="text-sm font-bold text-black">
          {item.tenGodBranch}
        </div>
        <div className="mt-2 text-sm font-bold text-[#6b3f24]">
          {item.twelveStage}
        </div>

        {item.branchRelations?.length > 0 &&
          (!isCompatibilityMode || showCompatibilityRelations) && (
            <div className="mt-2 flex flex-wrap justify-center gap-1">
              {item.branchRelations.map((relation: string) => (
                <span
                  key={relation}
                  className="rounded-full border border-[#ead8c4] bg-[#fffaf3] px-2 py-0.5 text-sm font-bold text-[#6b3f24]"
                >
                  {relation}
                </span>
              ))}
            </div>
          )}

        {item.shinsals?.length > 0 &&
          (!isCompatibilityMode || showCompatibilityRelations) && (
            <div className="mt-2 border-t border-[#ead8c4] pt-2">
              <div className="mb-1 text-center text-sm font-bold text-zinc-500">
                신살
              </div>

              <div className="flex flex-wrap justify-center gap-1">
                {item.shinsals.map((shinsal: string) => (
                  <span
                    key={shinsal}
                    className="rounded-full border border-[#d7c4ad] bg-white px-2 py-0.5 text-sm font-bold text-[#5f3a20]"
                  >
                    {shinsal}
                  </span>
                ))}
              </div>
            </div>
          )}
      </div>
    );
  };

  const renderLuckPanel = (
    targetSaju: any,
    cardKey: string,
    birthDate: string,
  ) => {
    if (!targetSaju?.daewoon || !birthDate) return null;

    const selectedDaewoon = targetSaju.daewoon.find(
      (item: any) =>
        selectedDaewoonKey[cardKey] === `${cardKey}-daewoon-${item.index}`,
    );

    const yearLuckList = selectedDaewoon
      ? buildYearLuckList(targetSaju, selectedDaewoon, birthDate)
      : [];

    const selectedYearLuck = yearLuckList.find(
      (yearLuck: any) =>
        selectedYearLuckKey[cardKey] === `${cardKey}-year-${yearLuck.year}`,
    );

    const monthLuckList = selectedYearLuck
      ? buildMonthLuckList(targetSaju, selectedYearLuck)
      : [];

    return (
      <div className="mt-6 rounded-2xl bg-white p-4 shadow-sm">
        <h3 className="text-lg font-bold text-black">대운</h3>

        <div className="mt-4 grid grid-cols-2 gap-3 sm:grid-cols-5">
          {targetSaju.daewoon.map((item: any) => {
            const daewoonKey = `${cardKey}-daewoon-${item.index}`;
            const selected = selectedDaewoonKey[cardKey] === daewoonKey;

            return (
              <button
                type="button"
                key={item.index}
                onClick={() => {
                  setSelectedDaewoonKey((prev) => ({
                    ...prev,
                    [cardKey]: selected ? null : daewoonKey,
                  }));
                  setSelectedYearLuckKey((prev) => ({
                    ...prev,
                    [cardKey]: null,
                  }));
                }}
                className={`rounded-xl p-3 text-center transition ${selected
                  ? "bg-[#6b3f24] text-white shadow-md"
                  : "bg-zinc-100 text-black hover:bg-zinc-200"
                  }`}
              >
                <div
                  className={
                    selected ? "text-sm text-white/80 font-bold " : "text-sm text-black font-bold "
                  }
                >
                  {item.startAgeText}
                </div>

                <div className="mt-2 flex flex-col items-center">
                  <div
                    className="text-3xl font-bold"
                    style={{
                      color: getElementColor(item.ganji.stemElement),
                      WebkitTextStroke: "1px black",
                    }}
                  >
                    {item.ganji.stem}
                  </div>

                  <div
                    className="text-3xl font-bold "
                    style={{
                      color: getElementColor(item.ganji.branchElement),
                      WebkitTextStroke: "1px black",
                    }}
                  >
                    {item.ganji.branch}
                  </div>
                </div>

                <div
                  className={
                    selected
                      ? "mt-2 text-sm font-bold text-white"
                      : "mt-2 text-sm font-bold text-black"
                  }
                >
                  {item.stemTenGod}
                </div>

                <div
                  className={
                    selected
                      ? "text-sm font-bold text-white"
                      : "text-sm font-bold text-black"
                  }
                >
                  {item.branchTenGod}
                </div>
              </button>
            );
          })}
        </div>

        {selectedDaewoon && (
          <div className="mt-5 rounded-2xl bg-zinc-100 p-4">
            <h4 className="font-bold text-black">선택한 대운의 년운</h4>

            <div className="mt-4 grid grid-cols-2 gap-3 sm:grid-cols-5">
              {yearLuckList.map((yearLuck: any) => {
                const yearLuckKey = `${cardKey}-year-${yearLuck.year}`;
                const selected = selectedYearLuckKey[cardKey] === yearLuckKey;

                return (
                  <button
                    type="button"
                    key={yearLuck.year}
                    onClick={() =>
                      setSelectedYearLuckKey((prev) => ({
                        ...prev,
                        [cardKey]: selected ? null : yearLuckKey,
                      }))
                    }
                    className={`rounded-xl p-3 text-center shadow-sm transition ${selected
                      ? "bg-[#6b3f24] text-white shadow-md"
                      : "bg-white text-black hover:bg-zinc-50"
                      }`}
                  >
                    <div
                      className={
                        selected
                          ? "text-sm font-bold text-white/80"
                          : "text-sm font-bold text-black"
                      }
                    >
                      {yearLuck.year} / <br /> {yearLuck.age}세
                    </div>

                    <div className="mt-2 flex flex-col items-center leading-tight">
                      <div
                        className="text-3xl font-bold leading-none"
                        style={{
                          color: getElementColor(yearLuck.ganji.stemElement),
                          WebkitTextStroke: "1.5px black",
                        }}
                      >
                        {STEM_HANJA[yearLuck.ganji.stem]}
                      </div>

                      <div
                        className="mt-1 text-3xl font-bold leading-none"
                        style={{
                          color: getElementColor(yearLuck.ganji.branchElement),
                          WebkitTextStroke: "1.5px black",
                        }}
                      >
                        {BRANCH_HANJA[yearLuck.ganji.branch]}
                      </div>

                      <div
                        className={
                          selected
                            ? "text-smm font-semibold font-bold text-white"
                            : "text-smm font-semibold font-bold text-black"
                        }
                      >
                        {yearLuck.stemTenGod}
                      </div>

                      <div
                        className={
                          selected
                            ? "text-smm font-semibold text-white"
                            : "text-smm font-semibold text-black"
                        }
                      >
                        {yearLuck.branchTenGod}
                      </div>
                    </div>
                  </button>
                );
              })}
            </div>

            {selectedYearLuck && (
              <div className="mt-5 rounded-2xl bg-white p-4">
                <h4 className="font-bold text-black">
                  {selectedYearLuck.year}년 월운
                </h4>

                <div className="mt-4 grid grid-cols-2 gap-3 sm:grid-cols-4">
                  {monthLuckList.map((monthLuck: any) => (
                    <div
                      key={`${monthLuck.year}-${monthLuck.month}`}
                      className="rounded-xl bg-zinc-100 p-3 text-center shadow-sm"
                    >
                      <div className="text-sm font-bold text-black">
                        {monthLuck.month}월
                      </div>

                      <div className="mt-2 flex flex-col items-center leading-tight">
                        <div
                          className="text-3xl font-bold leading-none"
                          style={{
                            color: getElementColor(monthLuck.ganji.stemElement),
                            WebkitTextStroke: "1.5px black",
                          }}
                        >
                          {STEM_HANJA[monthLuck.ganji.stem]}
                        </div>

                        <div
                          className="mt-1 text-3xl font-bold leading-none"
                          style={{
                            color: getElementColor(
                              monthLuck.ganji.branchElement,
                            ),
                            WebkitTextStroke: "1.5px black",
                          }}
                        >
                          {BRANCH_HANJA[monthLuck.ganji.branch]}
                        </div>

                        <div className="text-sm font-semibold font-bold  text-black">
                          {monthLuck.stemTenGod}
                        </div>

                        <div className="text-sm font-semibold font-bold  text-black">
                          {monthLuck.branchTenGod}
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        )}
      </div>
    );
  };

  const renderSajuCard = (
    targetSaju: any,
    cardKey: string,
    birthDate: string,
  ) => {
    const items = buildSajuItems(targetSaju);

    if (!targetSaju) return null;

    return (
      <div className="mt-4 rounded-2xl bg-[#ffffff] p-4 text-[#000000] shadow-sm">
        <h3 className="text-lg font-bold">사주팔자</h3>

        <div className="mt-4 grid grid-cols-2 gap-3 text-center sm:grid-cols-4">
          {items.map((item) => renderPillarCard(item, cardKey))}
        </div>

        {renderSpecialInfo(targetSaju)}

        <div className="mt-5 rounded-2xl bg-zinc-100 p-4">
          <h4 className="font-bold">오행 분포</h4>

          <div className="mt-3 grid grid-cols-5 gap-2 text-center text-sm">
            <div>
              <div className="font-bold text-xl">木</div>
              <div className="font-bold text-xl"> {targetSaju?.elementCount?.wood ?? 0}</div>
            </div>

            <div>
              <div className="font-bold text-xl">火</div>
              <div className="font-bold text-xl"> {targetSaju?.elementCount?.fire ?? 0}</div>
            </div>

            <div>
              <div className="font-bold text-xl">土</div>
              <div className="font-bold text-xl"> {targetSaju?.elementCount?.earth ?? 0}</div>
            </div>

            <div>
              <div className="font-bold text-xl">金</div>
              <div className="font-bold text-xl"> {targetSaju?.elementCount?.metal ?? 0}</div>
            </div>

            <div>
              <div className="font-bold text-xl">水</div>
              <div className="font-bold text-xl"> {targetSaju?.elementCount?.water ?? 0}</div>
            </div>
          </div>
        </div>

        {renderLuckPanel(targetSaju, cardKey, birthDate)}
      </div>
    );
  };

  return (
    <main className="min-h-screen bg-[#f7efe3] px-5 py-10 text-[#2b1d12]">
      <div className="mx-auto w-[1400px] min-w-[1400px] rounded-3xl bg-white p-6 shadow-xl">
        <h1 className="text-center text-3xl font-bold">
          {mode === "saju"
            ? "사주 분석"
            : mode === "compatibility"
              ? "궁합 분석"
              : "점성술 분석"}
        </h1>

        <div className="mt-6 rounded-2xl border border-[#ead8c4] bg-[#fffaf3] p-4 shadow-inner">
          <h2 className="text-lg font-bold">양력 → 음력 변환</h2>

          <div className="mt-3 grid grid-cols-1 gap-3 md:grid-cols-2">
            <input
              type="text"
              inputMode="numeric"
              placeholder="예: 1993-08-04"
              className="w-full rounded-xl border p-3"
              value={solarToLunarDate}
              onChange={(e) =>
                setSolarToLunarDate(formatDateInput(e.target.value))
              }
            />

            <div className="rounded-xl bg-white p-3 text-smm font-bold text-[#6b3f24]">
              음력:{" "}
              {solarToLunarDate
                ? convertSolarToLunar(solarToLunarDate)
                : "양력을 입력하세요"}
            </div>
          </div>
        </div>
        <div className="mt-4 rounded-2xl border border-[#ead8c4] bg-[#fffaf3] p-4 shadow-inner">
          <h2 className="text-lg font-bold">음력 → 양력 변환</h2>

          <div className="mt-3 grid grid-cols-1 gap-3 md:grid-cols-2">
            <input
              type="text"
              inputMode="numeric"
              placeholder="예: 1993-08-04"
              className="w-full rounded-xl border p-3"
              value={lunarToSolarDate}
              onChange={(e) => {
                const formatted = formatDateInput(e.target.value);

                setLunarToSolarDate(formatted);

                if (/^\d{4}-\d{2}-\d{2}$/.test(formatted)) {
                  const calendar: any = new KoreanLunarCalendar();

                  const [year, month, day] = formatted.split("-").map(Number);

                  const success = calendar.setLunarDate(
                    year,
                    month,
                    day,
                    lunarToSolarIsLeapMonth,
                  );

                  if (!success) {
                    setLunarToSolarResult("변환 불가");
                    return;
                  }

                  const solar = calendar.getSolarCalendar();

                  setLunarToSolarResult(
                    `${solar.year}-${String(solar.month).padStart(2, "0")}-${String(
                      solar.day,
                    ).padStart(2, "0")}`,
                  );
                } else {
                  setLunarToSolarResult("");
                }
              }}
            />

            <div className="rounded-xl bg-white p-3 text-smm font-bold text-[#6b3f24]">
              양력: {lunarToSolarResult || "음력을 입력하세요"}
            </div>
          </div>

          <label className="mt-3 flex items-center gap-2 rounded-xl border bg-white px-4 py-3 text-sm font-bold text-[#6b3f24]">
            <input
              type="checkbox"
              checked={lunarToSolarIsLeapMonth}
              onChange={(e) => {
                const checked = e.target.checked;

                setLunarToSolarIsLeapMonth(checked);

                if (!/^\d{4}-\d{2}-\d{2}$/.test(lunarToSolarDate)) {
                  setLunarToSolarResult("");
                  return;
                }

                const calendar: any = new KoreanLunarCalendar();

                const [year, month, day] = lunarToSolarDate.split("-").map(Number);

                const success = calendar.setLunarDate(
                  year,
                  month,
                  day,
                  checked,
                );

                if (!success) {
                  setLunarToSolarResult("변환 불가");
                  return;
                }

                const solar = calendar.getSolarCalendar();

                setLunarToSolarResult(
                  `${solar.year}-${String(solar.month).padStart(2, "0")}-${String(
                    solar.day,
                  ).padStart(2, "0")}`,
                );
              }}
            />
            윤달
          </label>


        </div>
        <div className="mt-6 grid grid-cols-3 gap-2 rounded-2xl bg-[#f7efe3] p-1">
          <button
            type="button"
            onClick={() => {
              setMode("saju");
              setResult("");
              setShowSaju(false);
            }}
            className={`rounded-xl py-3 text-smm font-bold transition ${mode === "saju"
              ? "bg-[#6b3f24] text-white shadow"
              : "text-[#6b3f24]"
              }`}
          >
            사주 모드
          </button>

          <button
            type="button"
            onClick={() => {
              setMode("compatibility");
              setResult("");
              setShowSaju(false);
            }}
            className={`rounded-xl py-3 text-smm font-bold transition ${mode === "compatibility"
              ? "bg-[#6b3f24] text-white shadow"
              : "text-[#6b3f24]"
              }`}
          >
            궁합 모드
          </button>

          <button
            type="button"
            onClick={() => {
              setMode("zodiac");
              setResult("");
              setShowSaju(false);
            }}
            className={`rounded-xl py-3 text-smm font-bold transition ${mode === "zodiac"
              ? "bg-[#6b3f24] text-white shadow"
              : "text-[#6b3f24]"
              }`}
          >
            별자리 모드
          </button>
        </div>

        <div className="mt-8 space-y-4">
          {mode !== "compatibility" && (
            <>
              <input
                className="w-full rounded-xl border p-3"
                placeholder="이름"
                value={form.name}
                onChange={(e) =>
                  setForm({
                    ...form,
                    name: e.target.value,
                  })
                }
              />

              <select
                className="w-full rounded-xl border p-3"
                value={form.gender}
                onChange={(e) =>
                  setForm({
                    ...form,
                    gender: e.target.value,
                  })
                }
              >
                <option value="남성">남성</option>
                <option value="여성">여성</option>
              </select>

              <div className="flex items-center gap-3">
                <input
                  type="text"
                  inputMode="numeric"
                  placeholder="1993-08-04"
                  className="flex-1 rounded-xl border p-3"
                  value={form.birthDate}
                  onChange={(e) => {
                    setForm({
                      ...form,
                      birthDate: formatDateInput(e.target.value),
                    });

                    setSajuResult(null);
                    setResult("");
                  }}
                />

                <label className="flex items-center gap-1 text-sm font-bold">
                  <input
                    type="checkbox"
                    checked={form.calendarType === "solar"}
                    onChange={() =>
                      setForm({
                        ...form,
                        calendarType: "solar",
                        isLeapMonth: false,
                      })
                    }
                  />
                  양력
                </label>

                <label className="flex items-center gap-1 text-sm font-bold">
                  <input
                    type="checkbox"
                    checked={form.calendarType === "lunar"}
                    onChange={() =>
                      setForm({
                        ...form,
                        calendarType: "lunar",
                      })
                    }
                  />
                  음력
                </label>

                {form.calendarType === "lunar" && (
                  <label className="flex items-center gap-1 text-sm font-bold text-[#6b3f24]">
                    <input
                      type="checkbox"
                      checked={form.isLeapMonth}
                      onChange={(e) =>
                        setForm({
                          ...form,
                          isLeapMonth: e.target.checked,
                        })
                      }
                    />
                    윤달
                  </label>
                )}
              </div>

              <input
                type="text"
                inputMode="numeric"
                placeholder="23:00"
                className="w-full rounded-xl border p-3 disabled:bg-zinc-100 disabled:text-zinc-400"
                value={form.birthTime}
                disabled={form.birthTimeUnknown}
                onChange={(e) => {
                  setForm({
                    ...form,
                    birthTime: formatTimeInput(e.target.value),
                  });

                  setSajuResult(null);
                  setResult("");
                }}
              />

              <label className="flex items-center gap-2 rounded-xl border border-[#ead8c4] bg-[#fffaf3] px-4 py-3 text-smm font-bold text-[#6b3f24]">
                <input
                  type="checkbox"
                  checked={form.birthTimeUnknown}
                  onChange={(e) => {
                    setForm({
                      ...form,
                      birthTimeUnknown: e.target.checked,
                    });

                    setSajuResult(null);
                    setResult("");
                  }}
                />
                출생 시간을 모릅니다
              </label>

              {mode === "zodiac" && (
                <input
                  type="text"
                  placeholder="태어난 위치 예: 서울, 대한민국"
                  className="w-full rounded-xl border p-3"
                  value={form.birthLocation}
                  onChange={(e) =>
                    setForm({
                      ...form,
                      birthLocation: e.target.value,
                    })
                  }
                />
              )}
            </>
          )}

          {mode === "compatibility" && (
            <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
              {(["left", "right"] as const).map((key) => (
                <div
                  key={key}
                  className="rounded-3xl border border-[#ead8c4] bg-[#fffaf3] p-5 shadow-inner"
                >
                  <h2 className="mb-4 text-center text-xl font-bold">
                    {key === "left" ? "본인" : "상대"}
                  </h2>

                  <div className="space-y-4">
                    <input
                      className="w-full rounded-xl border p-3"
                      placeholder="이름"
                      value={compatibilityForm[key].name}
                      onChange={(e) => {
                        setCompatibilityForm({
                          ...compatibilityForm,
                          [key]: {
                            ...compatibilityForm[key],
                            name: e.target.value,
                          },
                        });

                        setResult("");
                      }}
                    />

                    <select
                      className="w-full rounded-xl border p-3"
                      value={compatibilityForm[key].gender}
                      onChange={(e) => {
                        setCompatibilityForm({
                          ...compatibilityForm,
                          [key]: {
                            ...compatibilityForm[key],
                            gender: e.target.value,
                          },
                        });

                        setCompatibilityResult({
                          left: null,
                          right: null,
                        });

                        setResult("");
                      }}
                    >
                      <option value="남성">남성</option>
                      <option value="여성">여성</option>
                    </select>

                    <div className="flex items-center gap-3">
                      <input
                        type="text"
                        inputMode="numeric"
                        placeholder="1993-08-04"
                        className="flex-1 rounded-xl border p-3"
                        value={compatibilityForm[key].birthDate}
                        onChange={(e) => {
                          setCompatibilityForm({
                            ...compatibilityForm,
                            [key]: {
                              ...compatibilityForm[key],
                              birthDate: formatDateInput(e.target.value),
                            },
                          });

                          setCompatibilityResult({
                            left: null,
                            right: null,
                          });

                          setResult("");
                        }}
                      />

                      <label className="flex items-center gap-1 text-sm font-bold">
                        <input
                          type="checkbox"
                          checked={compatibilityForm[key].calendarType === "solar"}
                          onChange={() =>
                            setCompatibilityForm({
                              ...compatibilityForm,
                              [key]: {
                                ...compatibilityForm[key],
                                calendarType: "solar",
                                isLeapMonth: false,
                              },
                            })
                          }
                        />
                        양력
                      </label>

                      <label className="flex items-center gap-1 text-sm font-bold">
                        <input
                          type="checkbox"
                          checked={compatibilityForm[key].calendarType === "lunar"}
                          onChange={() =>
                            setCompatibilityForm({
                              ...compatibilityForm,
                              [key]: {
                                ...compatibilityForm[key],
                                calendarType: "lunar",
                              },
                            })
                          }
                        />
                        음력
                      </label>

                      {compatibilityForm[key].calendarType === "lunar" && (
                        <label className="flex items-center gap-1 text-sm font-bold text-[#6b3f24]">
                          <input
                            type="checkbox"
                            checked={compatibilityForm[key].isLeapMonth}
                            onChange={(e) =>
                              setCompatibilityForm({
                                ...compatibilityForm,
                                [key]: {
                                  ...compatibilityForm[key],
                                  isLeapMonth: e.target.checked,
                                },
                              })
                            }
                          />
                          윤달
                        </label>
                      )}
                    </div>

                    <input
                      type="text"
                      inputMode="numeric"
                      placeholder="23:00"
                      className="w-full rounded-xl border p-3 disabled:bg-zinc-100 disabled:text-zinc-400"
                      value={compatibilityForm[key].birthTime}
                      disabled={compatibilityForm[key].birthTimeUnknown}
                      onChange={(e) => {
                        setCompatibilityForm({
                          ...compatibilityForm,
                          [key]: {
                            ...compatibilityForm[key],
                            birthTime: formatTimeInput(e.target.value),
                          },
                        });

                        setCompatibilityResult({
                          left: null,
                          right: null,
                        });

                        setResult("");
                      }}
                    />

                    <label className="flex items-center gap-2 rounded-xl border border-[#ead8c4] bg-white px-4 py-3 text-smm font-bold text-[#6b3f24]">
                      <input
                        type="checkbox"
                        checked={compatibilityForm[key].birthTimeUnknown}
                        onChange={(e) => {
                          setCompatibilityForm({
                            ...compatibilityForm,
                            [key]: {
                              ...compatibilityForm[key],
                              birthTimeUnknown: e.target.checked,
                            },
                          });

                          setCompatibilityResult({
                            left: null,
                            right: null,
                          });

                          setResult("");
                        }}
                      />
                      출생 시간을 모릅니다
                    </label>
                  </div>
                </div>
              ))}
            </div>
          )}

          {mode === "saju" && (
            <>
              <button
                type="button"
                onClick={() => setShowSaju((prev) => !prev)}
                className="w-full rounded-xl border border-[#6b3f24]/40 bg-[#fff7ed] py-4 font-bold text-[#6b3f24] shadow-sm transition hover:bg-[#f3e1cf]"
              >
                {showSaju ? "만세력 닫기" : "만세력 보기"}
              </button>

              <button
                type="button"
                onClick={handleSubmit}
                disabled={loading}
                className="w-full cursor-pointer rounded-xl bg-[#6b3f24] py-4 font-bold text-white disabled:opacity-50"
              >
                {loading ? "분석 중..." : "사주 분석하기"}
              </button>
            </>
          )}

          {mode === "compatibility" && (
            <>
              <button
                type="button"
                onClick={handleCalculateCompatibility}
                disabled={
                  !compatibilityForm.left.birthDate ||
                  (!compatibilityForm.left.birthTime &&
                    !compatibilityForm.left.birthTimeUnknown) ||
                  !compatibilityForm.right.birthDate ||
                  (!compatibilityForm.right.birthTime &&
                    !compatibilityForm.right.birthTimeUnknown)
                }
                className="w-full rounded-xl border border-[#6b3f24]/40 bg-[#fff7ed] py-4 font-bold text-[#6b3f24] shadow-sm transition hover:bg-[#f3e1cf] disabled:opacity-40"
              >
                두 사람 만세력 보기
              </button>

              <button
                type="button"
                onClick={handleSubmit}
                disabled={loading}
                className="w-full cursor-pointer rounded-xl bg-[#6b3f24] py-4 font-bold text-white disabled:opacity-50"
              >
                {loading ? "분석 중..." : "궁합 분석하기"}
              </button>
            </>
          )}

          {mode === "zodiac" && (
            <button
              type="button"
              onClick={handleSubmit}
              disabled={loading}
              className="w-full cursor-pointer rounded-xl bg-[#6b3f24] py-4 font-bold text-white disabled:opacity-50"
            >
              {loading ? "분석 중..." : "점성술 분석하기"}
            </button>
          )}

          {mode === "saju" && showSaju && (
            <section className="mt-6 rounded-3xl border border-[#ead8c4] bg-[#fffaf3] p-5 shadow-inner">
              <h2 className="text-xl font-bold">만세력 계산</h2>

              <p className="mt-2 text-smm text-zinc-600">
                위에 입력한 생년월일시를 기준으로 사주팔자와 오행 분포를
                계산합니다.
              </p>

              <button
                type="button"
                onClick={handleCalculateSaju}
                disabled={
                  !form.birthDate || (!form.birthTime && !form.birthTimeUnknown)
                }
                className="mt-5 w-full rounded-xl bg-[#2b1d12] px-5 py-3 font-bold text-white transition hover:bg-[#4a2f1c] disabled:opacity-40"
              >
                계산하기
              </button>

              {sajuResult && (
                <div
                  ref={resultRef}
                  data-pdf-target
                  className="mt-6 rounded-2xl bg-[#ffffff] p-4 text-[#000000] shadow-sm"
                >
                  <button
                    type="button"
                    onClick={downloadWord}
                    className="mt-4 w-full rounded-xl bg-black px-5 py-3 font-bold text-white shadow-md"
                  >
                    저장하기
                  </button>

                  <h3 className="mt-4 text-lg font-bold">사주팔자</h3>

                  <div className="mt-4 grid grid-cols-2 gap-3 text-center sm:grid-cols-4 font-bold text-black">
                    {sajuItems.map((item) => renderPillarCard(item, "main"))}
                  </div>

                  {renderSpecialInfo(sajuResult)}

                  <div className="mt-5 rounded-2xl bg-zinc-100 p-4">
                    <h4 className="font-bold">오행 분포</h4>

                    <div className="mt-3 grid grid-cols-5 gap-2 text-center text-smm">
                      <div>
                        <div className="font-bold text-2xl">木</div>
                        <div className="font-bold text-lg">{sajuResult?.elementCount?.wood ?? 0}</div>
                      </div>

                      <div>
                        <div className="font-bold text-2xl">火</div>
                        <div className="font-bold text-lg">{sajuResult?.elementCount?.fire ?? 0}</div>
                      </div>

                      <div>
                        <div className="font-bold text-2xl">土</div>
                        <div className="font-bold text-lg">{sajuResult?.elementCount?.earth ?? 0}</div>
                      </div>

                      <div>
                        <div className="font-bold text-2xl">金</div>
                        <div className="font-bold text-lg"> {sajuResult?.elementCount?.metal ?? 0}</div>
                      </div>

                      <div>
                        <div className="font-bold text-2xl">水</div>
                        <div className="font-bold text-lg"> {sajuResult?.elementCount?.water ?? 0}</div>
                      </div>
                    </div>
                  </div>

                  {renderLuckPanel(sajuResult, "main", form.birthDate)}

                  {result && (
                    <div className="mt-5 rounded-2xl bg-zinc-100 p-4">
                      <h4 className="font-bold">사주 해석</h4>

                      <div className="mt-3 whitespace-pre-wrap leading-7 text-zinc-700">
                        {result}
                      </div>
                    </div>
                  )}
                </div>
              )}
            </section>
          )}

          {mode === "compatibility" &&
            (compatibilityResult.left || compatibilityResult.right) && (
              <section className="mt-6 rounded-3xl border border-[#ead8c4] bg-[#fffaf3] p-5 shadow-inner">
                <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                  <h2 className="text-xl font-bold">두 사람 만세력</h2>

                  <button
                    type="button"
                    onClick={() =>
                      setShowCompatibilityRelations((prev) => !prev)
                    }
                    className="rounded-full border border-[#6b3f24]/40 bg-white px-4 py-2 text-smm font-bold text-[#6b3f24] shadow-sm transition hover:bg-[#f3e1cf]"
                  >
                    {showCompatibilityRelations
                      ? "지지 관계·신살 전체 접기 ▲"
                      : "지지 관계·신살 전체 열기 ▼"}
                  </button>
                </div>

                <div className="mt-4 grid grid-cols-1 gap-4 lg:grid-cols-2">
                  <div>
                    <h3 className="text-center text-lg font-bold">
                      {compatibilityForm.left.name || "본인"}
                    </h3>
                    {renderSajuCard(
                      compatibilityResult.left,
                      "compat-left",
                      compatibilityForm.left.birthDate,
                    )}
                  </div>

                  <div>
                    <h3 className="text-center text-lg font-bold">
                      {compatibilityForm.right.name || "상대"}
                    </h3>
                    {renderSajuCard(
                      compatibilityResult.right,
                      "compat-right",
                      compatibilityForm.right.birthDate,
                    )}
                  </div>
                </div>
              </section>
            )}

          {mode === "compatibility" && result && (
            <section className="mt-6 rounded-3xl border border-[#ead8c4] bg-[#fffaf3] p-5 shadow-inner">
              <h2 className="text-xl font-bold">궁합 해석</h2>

              <div className="mt-3 whitespace-pre-wrap leading-7 text-zinc-700">
                {result}
              </div>
            </section>
          )}

          {mode === "zodiac" && result && (
            <section className="mt-6 rounded-3xl border border-[#ead8c4] bg-[#fffaf3] p-5 shadow-inner">
              <h2 className="text-xl font-bold">점성술 해석</h2>

              <div className="mt-3 whitespace-pre-wrap leading-7 text-zinc-700">
                {result}
              </div>
            </section>
          )}
        </div>
      </div>
    </main>
  );
}
