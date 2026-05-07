"use client";

import { useState, useRef } from "react";
import { saveAs } from "file-saver";

import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
} from "docx";

import { calculateSaju } from "./lib/sajuCalculator";

export default function Home() {
  const resultRef = useRef<HTMLDivElement>(null);

  const [form, setForm] = useState({
    name: "",
    gender: "남성",
    birthDate: "",
    birthTime: "23:00",
  });

  const [result, setResult] = useState("");
  const [loading, setLoading] = useState(false);
  const [showSaju, setShowSaju] = useState(false);
  const [sajuResult, setSajuResult] = useState<any>(null);

  const handleCalculateSaju = () => {
    if (!form.birthDate || !form.birthTime) return;

    const calculated = calculateSaju({
      birthDate: form.birthDate,
      birthTime: form.birthTime,
      calendarType: "solar",
      timezone: "Asia/Seoul",
      lateZiMode: false,
    });

    setSajuResult(calculated);
  };

  async function handleSubmit() {
    setLoading(true);
    setResult("요청 보내는 중...");

    try {
      const res = await fetch("/api/saju", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(form),
      });

      const data = await res.json();

      setResult(data.result || "결과가 비어 있습니다.");
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
                      children: [new Paragraph("년주")],
                    }),
                    new TableCell({
                      children: [new Paragraph("월주")],
                    }),
                    new TableCell({
                      children: [new Paragraph("일주")],
                    }),
                    new TableCell({
                      children: [new Paragraph("시주")],
                    }),
                  ],
                }),

                new TableRow({
                  children: [
                    new TableCell({
                      children: [
                        new Paragraph(
                          `${sajuResult.year.ganji} (${sajuResult.year.ganjiKor})`
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
                          `${sajuResult.day.ganji} (${sajuResult.day.ganjiKor})`
                        ),
                      ],
                    }),

                    new TableCell({
                      children: [
                        new Paragraph(
                          `${sajuResult.hour.ganji} (${sajuResult.hour.ganjiKor})`
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

            new Paragraph({
              text: result || "사주 풀이 결과가 없습니다.",
            }),
          ],
        },
      ],
    });

    const blob = await Packer.toBlob(doc);

    const safeBirth = form.birthDate.replaceAll("-", "");

saveAs(
  blob,
  `${form.name}_${safeBirth}.docx`
);
  } catch (error) {
    console.error(error);
    alert("워드 파일 생성 중 오류가 발생했습니다.");
  }
};
  

  const sajuItems = sajuResult
    ? [
        { label: "년주", data: sajuResult.year },
        { label: "월주", data: sajuResult.month },
        { label: "일주", data: sajuResult.day },
        { label: "시주", data: sajuResult.hour },
      ]
    : [];

  return (
    <main className="min-h-screen bg-[#f7efe3] px-5 py-10 text-[#2b1d12]">
      <div className="mx-auto max-w-md rounded-3xl bg-white p-6 shadow-xl">
        <h1 className="text-center text-3xl font-bold">사주 분석</h1>

        <div className="mt-8 space-y-4">
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
              const onlyNumber = e.target.value
                .replace(/\D/g, "")
                .slice(0, 8);

              let formatted = onlyNumber;

              if (
                onlyNumber.length > 4 &&
                onlyNumber.length <= 6
              ) {
                formatted = `${onlyNumber.slice(
                  0,
                  4
                )}-${onlyNumber.slice(4)}`;
              }

              if (onlyNumber.length > 6) {
                formatted = `${onlyNumber.slice(
                  0,
                  4
                )}-${onlyNumber.slice(
                  4,
                  6
                )}-${onlyNumber.slice(6)}`;
              }

              setForm({
                ...form,
                birthDate: formatted,
              });

              setSajuResult(null);
            }}
          />

          <input
            type="text"
            inputMode="numeric"
            placeholder="23:00"
            className="w-full rounded-xl border p-3"
            value={form.birthTime}
            onChange={(e) => {
              const onlyNumber = e.target.value
                .replace(/\D/g, "")
                .slice(0, 4);

              let formatted = onlyNumber;

              if (onlyNumber.length >= 3) {
                formatted = `${onlyNumber.slice(
                  0,
                  2
                )}:${onlyNumber.slice(2)}`;
              }

              setForm({
                ...form,
                birthTime: formatted,
              });

              setSajuResult(null);
            }}
          />
          <button
            type="button"
            onClick={() =>
              setShowSaju((prev) => !prev)
            }
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

          

        
          {showSaju && (
            <section className="mt-6 rounded-3xl border border-[#ead8c4] bg-[#fffaf3] p-5 shadow-inner">
              <h2 className="text-xl font-bold">
                만세력 계산
              </h2>

              <p className="mt-2 text-sm text-zinc-600">
                위에 입력한 생년월일시를 기준으로
                사주팔자와 오행 분포를 계산합니다.
              </p>

              <button
                type="button"
                onClick={handleCalculateSaju}
                disabled={
                  !form.birthDate || !form.birthTime
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
                      onClick={downloadWord}
                      className="mt-4 w-full rounded-xl bg-black px-5 py-3 font-bold text-white shadow-md"
                    >
                      저장하기
                    </button>
                  <h3 className="text-lg font-bold">
                    사주팔자
                  </h3>

                  <div className="mt-4 grid grid-cols-2 gap-3 text-center sm:grid-cols-4">
                    {sajuItems.map((item) => (
                      <div
                        key={item.label}
                        className="rounded-xl bg-zinc-100 p-3"
                      >
                        <div className="text-xs text-zinc-500">
                          {item.label}
                        </div>

                        <div className="mt-1 text-2xl font-bold">
                          {item.data.ganji}
                        </div>

                        <div className="text-sm">
                          {item.data.ganjiKor}
                        </div>

                        <div className="mt-2 text-xs text-zinc-500">
                          천간 {item.data.stemElement}
                        </div>

                        <div className="text-xs text-zinc-500">
                          지지 {item.data.branchElement}
                        </div>
                      </div>
                    ))}
                  </div>

                  <div className="mt-5 rounded-2xl bg-zinc-100 p-4">
                    <h4 className="font-bold">
                      오행 분포
                    </h4>

                    <div className="mt-3 grid grid-cols-5 gap-2 text-center text-sm">
                      <div>
                        <div className="font-bold">
                          木
                        </div>
                        <div>
                          {sajuResult?.elementCount
                            ?.wood ?? 0}
                        </div>
                      </div>

                      <div>
                        <div className="font-bold">
                          火
                        </div>
                        <div>
                          {sajuResult?.elementCount
                            ?.fire ?? 0}
                        </div>
                      </div>

                      <div>
                        <div className="font-bold">
                          土
                        </div>
                        <div>
                          {sajuResult?.elementCount
                            ?.earth ?? 0}
                        </div>
                      </div>

                      <div>
                        <div className="font-bold">
                          金
                        </div>
                        <div>
                          {sajuResult?.elementCount
                            ?.metal ?? 0}
                        </div>
                      </div>

                      <div>
                        <div className="font-bold">
                          水
                        </div>
                        <div>
                          {sajuResult?.elementCount
                            ?.water ?? 0}
                        </div>
                      </div>
                    </div>
                  </div>

                  <div className="mt-5 rounded-2xl bg-zinc-100 p-4">
                    <h4 className="font-bold">
                      십성 (천간)
                    </h4>

                    <div className="mt-3 grid grid-cols-4 gap-2 text-center text-sm">
                      <div>
                        <div className="font-bold">
                          년간
                        </div>
                        <div>
                          {
                            sajuResult?.tenGods
                              ?.yearStem
                          }
                        </div>
                      </div>

                      <div>
                        <div className="font-bold">
                          월간
                        </div>
                        <div>
                          {
                            sajuResult?.tenGods
                              ?.monthStem
                          }
                        </div>
                      </div>

                      <div>
                        <div className="font-bold">
                          일간
                        </div>
                        <div>
                          {
                            sajuResult?.tenGods
                              ?.dayStem
                          }
                        </div>
                      </div>

                      <div>
                        <div className="font-bold">
                          시간
                        </div>
                        <div>
                          {
                            sajuResult?.tenGods
                              ?.hourStem
                          }
                        </div>
                      </div>
                    </div>
                  </div>

                  <div className="mt-5 rounded-2xl bg-zinc-100 p-4">
                    <h4 className="font-bold">
                      십성 (지지)
                    </h4>

                    <div className="mt-3 grid grid-cols-4 gap-2 text-center text-sm">
                      <div>
                        <div className="font-bold">
                          년지
                        </div>
                        <div>
                          {
                            sajuResult?.tenGods
                              ?.yearBranch
                          }
                        </div>
                      </div>

                      <div>
                        <div className="font-bold">
                          월지
                        </div>
                        <div>
                          {
                            sajuResult?.tenGods
                              ?.monthBranch
                          }
                        </div>
                      </div>

                      <div>
                        <div className="font-bold">
                          일지
                        </div>
                        <div>
                          {
                            sajuResult?.tenGods
                              ?.dayBranch
                          }
                        </div>
                      </div>

                      <div>
                        <div className="font-bold">
                          시지
                        </div>
                        <div>
                          {
                            sajuResult?.tenGods
                              ?.hourBranch
                          }
                        </div>
                      </div>
                    </div>
                  </div>

                                    <div className="mt-5 rounded-2xl bg-zinc-100 p-4">
                    <h4 className="font-bold">
                      십이운성
                    </h4>

                    <div className="mt-3 grid grid-cols-4 gap-2 text-center text-sm">
                      <div>
                        <div className="font-bold">
                          년지
                        </div>
                        <div>
                          {
                            sajuResult
                              ?.twelveStages?.year
                          }
                        </div>
                      </div>

                      <div>
                        <div className="font-bold">
                          월지
                        </div>
                        <div>
                          {
                            sajuResult
                              ?.twelveStages?.month
                          }
                        </div>
                      </div>

                      <div>
                        <div className="font-bold">
                          일지
                        </div>
                        <div>
                          {
                            sajuResult
                              ?.twelveStages?.day
                          }
                        </div>
                      </div>

                      <div>
                        <div className="font-bold">
                          시지
                        </div>
                        <div>
                          {
                            sajuResult
                              ?.twelveStages?.hour
                          }
                        </div>
                      </div>
                    </div>
                  </div>

                  {result && (
                    <div className="mt-5 rounded-2xl bg-zinc-100 p-4">
                      <h4 className="font-bold">
                        사주 해석
                      </h4>

                      <div className="mt-3 whitespace-pre-wrap leading-7 text-zinc-700">
                        {result}
                      </div>
                    </div>
                  )}
                </div>
              )}
            </section>
          )}
        </div>
      </div>
    </main>
  );
}