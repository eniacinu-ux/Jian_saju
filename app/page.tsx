"use client";

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
  const resultRef = useRef<HTMLDivElement>(null);

  const [mode, setMode] = useState<"saju" | "compatibility" | "zodiac">(
    "saju"
  );

  const [form, setForm] = useState({
    name: "",
    gender: "남성",
    birthDate: "",
    birthTime: "23:00",
    birthLocation: "",
  });

  const [compatibilityForm, setCompatibilityForm] = useState({
    left: {
      name: "",
      gender: "남성",
      birthDate: "",
      birthTime: "23:00",
    },
    right: {
      name: "",
      gender: "여성",
      birthDate: "",
      birthTime: "23:00",
    },
  });

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
        6
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

  const calculateOneSaju = (targetForm: {
    birthDate: string;
    birthTime: string;
    gender: string;
  }) => {
    if (!targetForm.birthDate || !targetForm.birthTime) return null;

    return calculateSaju({
      birthDate: targetForm.birthDate,
      birthTime: targetForm.birthTime,
      calendarType: "solar",
      timezone: "Asia/Seoul",
      lateZiMode: false,
      gender: targetForm.gender === "남성" ? "male" : "female",
    });
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
          compatibilityResult.right || calculateOneSaju(compatibilityForm.right);

        setCompatibilityResult({
          left,
          right,
        });

        bodyData = {
          mode,
          left: compatibilityForm.left,
          right: compatibilityForm.right,
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
                    text: `생년월일: ${form.birthDate} ${form.birthTime}`,
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
                            `${sajuResult.hour.ganji} (${sajuResult.hour.ganjiKor})`
                          ),
                        ],
                      }),

                      new TableCell({
                        children: [
                          new Paragraph(
                            `${sajuResult.day.ganji} (${sajuResult.day.ganjiKor})`
                          ),
                        ],
                      }),

                      new TableCell({
                        children: [
                          new Paragraph(
                            `${sajuResult.month.ganji} (${sajuResult.month.ganjiKor})`
                          ),
                        ],
                      }),

                      new TableCell({
                        children: [
                          new Paragraph(
                            `${sajuResult.year.ganji} (${sajuResult.year.ganjiKor})`
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
                      })
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

  const buildSajuItems = (targetSaju: any) =>
    targetSaju
      ? [
          {
            label: "시주",
            data: targetSaju.hour,
            tenGodStem: targetSaju.tenGods.hourStem,
            tenGodBranch: targetSaju.tenGods.hourBranch,
            twelveStage: targetSaju.twelveStages.hour,
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
        ]
      : [];

  const sajuItems = buildSajuItems(sajuResult);

  const renderSajuCard = (targetSaju: any) => {
    const items = buildSajuItems(targetSaju);

    if (!targetSaju) return null;

    return (
      <div className="mt-4 rounded-2xl bg-[#ffffff] p-4 text-[#000000] shadow-sm">
        <h3 className="text-lg font-bold">사주팔자</h3>

        <div className="mt-4 grid grid-cols-2 gap-3 text-center sm:grid-cols-4">
          {items.map((item) => (
            <div key={item.label} className="rounded-xl bg-zinc-100 p-3">
              <div className="text-xs text-zinc-500">{item.label}</div>

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

                <div className="mt-3 text-sm text-zinc-600">
                  {item.data.stemKor}
                </div>

                <div className="text-sm text-zinc-600">
                  {item.data.branchKor}
                </div>
              </div>

              <div className="mt-3 text-xs text-zinc-500">
                {item.tenGodStem}
              </div>

              <div className="text-xs text-zinc-500">
                {item.tenGodBranch}
              </div>

              <div className="mt-2 text-xs font-bold text-[#6b3f24]">
                {item.twelveStage}
              </div>
            </div>
          ))}
        </div>

        <div className="mt-5 rounded-2xl bg-zinc-100 p-4">
          
          <h4 className="font-bold">오행 분포</h4>

          <div className="mt-3 grid grid-cols-5 gap-2 text-center text-sm">
            <div>
              <div className="font-bold">木</div>
              <div>{targetSaju?.elementCount?.wood ?? 0}</div>
            </div>

            <div>
              <div className="font-bold">火</div>
              <div>{targetSaju?.elementCount?.fire ?? 0}</div>
            </div>

            <div>
              <div className="font-bold">土</div>
              <div>{targetSaju?.elementCount?.earth ?? 0}</div>
            </div>

            <div>
              <div className="font-bold">金</div>
              <div>{targetSaju?.elementCount?.metal ?? 0}</div>
            </div>

            <div>
              <div className="font-bold">水</div>
              <div>{targetSaju?.elementCount?.water ?? 0}</div>
            </div>
            
          </div>
        
        </div>
        {targetSaju?.daewoon && (
  <div className="mt-5 rounded-2xl bg-white p-4 shadow-sm">
    <h4 className="font-bold text-black">대운</h4>

    <div className="mt-4 grid grid-cols-2 gap-3 sm:grid-cols-5">
      {targetSaju.daewoon.map((item: any) => (
        <div
          key={item.index}
          className="rounded-xl bg-zinc-100 p-3 text-center"
        >
          <div className="text-xs text-zinc-500">
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
              className="text-3xl font-bold"
              style={{
                color: getElementColor(item.ganji.branchElement),
                WebkitTextStroke: "1px black",
              }}
            >
              {item.ganji.branch}
            </div>
          </div>

          <div className="mt-2 text-xs text-zinc-500">
            {item.stemTenGod}
          </div>

          <div className="text-xs text-zinc-500">
            {item.branchTenGod}
          </div>
        </div>
      ))}
    </div>
  </div>
)}
      </div>
      
    );
  };

  return (
    <main className="min-h-screen bg-[#f7efe3] px-5 py-10 text-[#2b1d12]">
      <div className="mx-auto max-w-5xl rounded-3xl bg-white p-6 shadow-xl">
        <h1 className="text-center text-3xl font-bold">
          {mode === "saju"
            ? "사주 분석"
            : mode === "compatibility"
            ? "궁합 분석"
            : "점성술 분석"}
        </h1>

        <div className="mt-6 grid grid-cols-3 gap-2 rounded-2xl bg-[#f7efe3] p-1">
          <button
            type="button"
            onClick={() => {
              setMode("saju");
              setResult("");
              setShowSaju(false);
            }}
            className={`rounded-xl py-3 text-sm font-bold transition ${
              mode === "saju"
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
            className={`rounded-xl py-3 text-sm font-bold transition ${
              mode === "compatibility"
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
            className={`rounded-xl py-3 text-sm font-bold transition ${
              mode === "zodiac"
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

              <input
                type="text"
                inputMode="numeric"
                placeholder="1993-08-04"
                className="w-full rounded-xl border p-3"
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

              <input
                type="text"
                inputMode="numeric"
                placeholder="23:00"
                className="w-full rounded-xl border p-3"
                value={form.birthTime}
                onChange={(e) => {
                  setForm({
                    ...form,
                    birthTime: formatTimeInput(e.target.value),
                  });

                  setSajuResult(null);
                  setResult("");
                }}
              />

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

                    <input
                      type="text"
                      inputMode="numeric"
                      placeholder="1993-08-04"
                      className="w-full rounded-xl border p-3"
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

                    <input
                      type="text"
                      inputMode="numeric"
                      placeholder="23:00"
                      className="w-full rounded-xl border p-3"
                      value={compatibilityForm[key].birthTime}
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
                  !compatibilityForm.left.birthTime ||
                  !compatibilityForm.right.birthDate ||
                  !compatibilityForm.right.birthTime
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

              <p className="mt-2 text-sm text-zinc-600">
                위에 입력한 생년월일시를 기준으로 사주팔자와 오행 분포를
                계산합니다.
              </p>

              <button
                type="button"
                onClick={handleCalculateSaju}
                disabled={!form.birthDate || !form.birthTime}
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

                  <div className="mt-4 grid grid-cols-2 gap-3 text-center sm:grid-cols-4">
                    {sajuItems.map((item) => (
                      <div
                        key={item.label}
                        className="rounded-xl bg-zinc-100 p-3"
                      >
                        <div className="text-xs text-zinc-500">
                          {item.label}
                        </div>

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

                          <div className="mt-3 text-sm text-zinc-600">
                            {item.data.stemKor}
                          </div>

                          <div className="text-sm text-zinc-600">
                            {item.data.branchKor}
                          </div>
                        </div>

                        <div className="mt-3 text-xs text-zinc-500">
                          {item.tenGodStem}
                        </div>

                        <div className="text-xs text-zinc-500">
                          {item.tenGodBranch}
                        </div>

                        <div className="mt-2 text-xs font-bold text-[#6b3f24]">
                          {item.twelveStage}
                        </div>
                      </div>
                    ))}
                  </div>

                  <div className="mt-5 rounded-2xl bg-zinc-100 p-4">
                    <h4 className="font-bold">오행 분포</h4>

                    <div className="mt-3 grid grid-cols-5 gap-2 text-center text-sm">
                      <div>
                        <div className="font-bold">木</div>
                        <div>{sajuResult?.elementCount?.wood ?? 0}</div>
                      </div>

                      <div>
                        <div className="font-bold">火</div>
                        <div>{sajuResult?.elementCount?.fire ?? 0}</div>
                      </div>

                      <div>
                        <div className="font-bold">土</div>
                        <div>{sajuResult?.elementCount?.earth ?? 0}</div>
                      </div>

                      <div>
                        <div className="font-bold">金</div>
                        <div>{sajuResult?.elementCount?.metal ?? 0}</div>
                      </div>

                      <div>
                        <div className="font-bold">水</div>
                        <div>{sajuResult?.elementCount?.water ?? 0}</div>
                      </div>
                    </div>
                  </div>

                  {sajuResult?.daewoon && (
                    <div className="mt-6 rounded-2xl bg-white p-4 shadow-sm">
                      <h3 className="text-lg font-bold text-black">대운</h3>

                      <div className="mt-4 grid grid-cols-2 gap-3 sm:grid-cols-5">
                        {sajuResult.daewoon.map((item: any) => (
                          <div
                            key={item.index}
                            className="rounded-xl bg-zinc-100 p-3 text-center"
                          >
                            <div className="text-xs text-zinc-500">
                              {item.startAgeText}
                            </div>

                            <div className="mt-2 flex flex-col items-center">
                              <div
                                className="text-3xl font-bold"
                                style={{
                                  color: getElementColor(
                                    item.ganji.stemElement
                                  ),
                                  WebkitTextStroke: "1px black",
                                }}
                              >
                                {item.ganji.stem}
                              </div>

                              <div
                                className="text-3xl font-bold"
                                style={{
                                  color: getElementColor(
                                    item.ganji.branchElement
                                  ),
                                  WebkitTextStroke: "1px black",
                                }}
                              >
                                {item.ganji.branch}
                              </div>
                            </div>

                            <div className="mt-2 text-xs text-zinc-500">
                              {item.stemTenGod}
                            </div>

                            <div className="text-xs text-zinc-500">
                              {item.branchTenGod}
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

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
                <h2 className="text-xl font-bold">두 사람 만세력</h2>

                <div className="mt-4 grid grid-cols-1 gap-4 lg:grid-cols-2">
                  <div>
                    <h3 className="text-center text-lg font-bold">
                      {compatibilityForm.left.name || "본인"}
                    </h3>
                    {renderSajuCard(compatibilityResult.left)}
                  </div>

                  <div>
                    <h3 className="text-center text-lg font-bold">
                      {compatibilityForm.right.name || "상대"}
                    </h3>
                    {renderSajuCard(compatibilityResult.right)}
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