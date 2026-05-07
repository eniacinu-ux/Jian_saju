import OpenAI from "openai";
import { calculateSaju } from "@/app/lib/sajuCalculator";

const client = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

function buildSajuPrompt(form: any, saju: any) {
  return `
[시스템 가이드: 만세력 천을귀인]
1. 본 분석은 '만세력 천을귀인'의 정밀한 로직으로 산출된 데이터를 바탕으로 합니다.
2. 제공된 사주 정보는 검증된 값이므로 다시 계산하지 말고, 이 데이터를 절대적 기준으로 해석하십시오.
3. 답변 시작 시 '만세력 천을귀인'앱의 데이터를 바탕으로 해석함을 가볍게 언급하며, 전문가의 품격에 맞는 존댓말로 답변해 주십시오.

[사주 정보]
-성별 : ${form.gender}
-성함 : ${form.name}
-생년월일시 : (양력)${form.birthDate} ${form.birthTime}

-사주팔자 :
년주(${saju.year.ganjiKor}),
월주(${saju.month.ganjiKor}),
일주(${saju.day.ganjiKor}),
시주(${saju.hour.ganjiKor})

-십성(천간) :
${saju.tenGods.yearStem}(년),
${saju.tenGods.monthStem}(월),
${saju.tenGods.dayStem}(일),
${saju.tenGods.hourStem}(시)

-십성(지지) :
${saju.tenGods.yearBranch}(년),
${saju.tenGods.monthBranch}(월),
${saju.tenGods.dayBranch}(일),
${saju.tenGods.hourBranch}(시)

-십이운성 :
${saju.twelveStages.year}(년),
${saju.twelveStages.month}(월),
${saju.twelveStages.day}(일),
${saju.twelveStages.hour}(시)

-오행 분포 :
木 ${saju.elementCount.wood},
火 ${saju.elementCount.fire},
土 ${saju.elementCount.earth},
金 ${saju.elementCount.metal},
水 ${saju.elementCount.water}

[질문 사항]
위 데이터를 바탕으로 명리학 전문가의 관점에서 다음 사항을 상세히 분석해 주십시오.

1. 일간과 일주를 중심으로 본연의 기질과 중심 성격을 설명해 주십시오.
2. 월지에 배정된 기운과 전체적인 십성의 흐름을 바탕으로, 이 사주가 사회에서 어떤 환경에 놓이기 쉬우며 어떤 방식으로 역량을 발휘하는지 분석해 주십시오.
3. 주어진 십성 구성에서 나타나는 특징적인 장단점과 그에 따른 인생 흐름의 특성을 분석해 주십시오.
4. 제공된 오행 분포 수치를 절대적 기준으로 삼아, 부족하거나 과한 기운을 조절할 수 있는 실생활의 보완책(색상, 습관 등)을 제안해 주십시오.
5. 재물운, 연애·결혼운, 직업 적성, 건강운 등 주요 영역을 주어진 데이터를 근거로 종합 해석해 주십시오.
6. 전체적인 사주 구성의 균형을 맞추기 위해 이 사주가 지향해야 할 삶의 태도와 핵심적인 조언을 들려주십시오.
`;
}

export async function POST(req: Request) {
  try {
    const body = await req.json();

    const saju = calculateSaju({
      birthDate: body.birthDate,
      birthTime: body.birthTime,
      calendarType: "solar",
      timezone: "Asia/Seoul",
      lateZiMode: false,
    });

    const prompt = buildSajuPrompt(body, saju);

    const completion = await client.chat.completions.create({
      model: "gpt-4.1-mini",
      messages: [
        {
          role: "system",
          content:
            "당신은 전문 명리학자이며 깊이 있고 품격 있는 사주 해석을 제공합니다.",
        },
        {
          role: "user",
          content: prompt,
        },
      ],
      temperature: 0.8,
      max_tokens: 3000,
    });

    return Response.json({
      result:
        completion.choices[0].message.content ??
        "결과가 없습니다.",
    });
  } catch (error) {
    console.error(error);

    return Response.json(
      {
        result: "분석 중 오류가 발생했습니다.",
      },
      { status: 500 }
    );
  }
}