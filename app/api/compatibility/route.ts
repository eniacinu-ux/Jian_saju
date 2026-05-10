import { NextResponse } from "next/server";
import OpenAI from "openai";

export const runtime = "nodejs";

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

function buildCompatibilityPrompt(data: any) {
  const { left, right, leftSaju, rightSaju } = data;

  return `
[궁합 분석 요청]

아래는 두 사람의 기본 정보와 사주 정보입니다.
이 정보를 바탕으로 궁합을 분석하세요.

====================
[본인 정보]
====================

- 이름: ${left?.name || "미입력"}
- 성별: ${left?.gender || "미입력"}
- 생년월일: ${left?.birthDate || "미입력"}
- 출생시간: ${left?.birthTime || "미입력"}

[본인 정보]

- 년주: ${leftSaju?.year?.ganji || ""} (${leftSaju?.year?.ganjiKor || ""})
- 월주: ${leftSaju?.month?.ganji || ""} (${leftSaju?.month?.ganjiKor || ""})
- 일주: ${leftSaju?.day?.ganji || ""} (${leftSaju?.day?.ganjiKor || ""})
- 시주: ${leftSaju?.hour?.ganji || ""} (${leftSaju?.hour?.ganjiKor || ""})

- 일간: ${leftSaju?.day?.stem || ""}
- 일지: ${leftSaju?.day?.branch || ""}

- 오행 분포:
  목 ${leftSaju?.elementCount?.wood ?? 0}
  화 ${leftSaju?.elementCount?.fire ?? 0}
  토 ${leftSaju?.elementCount?.earth ?? 0}
  금 ${leftSaju?.elementCount?.metal ?? 0}
  수 ${leftSaju?.elementCount?.water ?? 0}

- 십성:
  년간: ${leftSaju?.tenGods?.yearStem || ""}
  년지: ${leftSaju?.tenGods?.yearBranch || ""}
  월간: ${leftSaju?.tenGods?.monthStem || ""}
  월지: ${leftSaju?.tenGods?.monthBranch || ""}
  일간: ${leftSaju?.tenGods?.dayStem || ""}
  일지: ${leftSaju?.tenGods?.dayBranch || ""}
  시간: ${leftSaju?.tenGods?.hourStem || ""}
  시지: ${leftSaju?.tenGods?.hourBranch || ""}

- 십이운성:
  년주: ${leftSaju?.twelveStages?.year || ""}
  월주: ${leftSaju?.twelveStages?.month || ""}
  일주: ${leftSaju?.twelveStages?.day || ""}
  시주: ${leftSaju?.twelveStages?.hour || ""}

- 대운 정보:
${Array.isArray(leftSaju?.daewoon)
  ? leftSaju.daewoon
      .map(
        (item: any) =>
          `${item.startAgeText || ""} : ${item.ganji?.ganji || ""} (${
            item.ganji?.ganjiKor || ""
          }) / 천간십성 ${item.stemTenGod || ""} / 지지십성 ${
            item.branchTenGod || ""
          }`
      )
      .join("\n")
  : "대운 정보 없음"}

====================
[상대 정보]
====================

- 이름: ${right?.name || "미입력"}
- 성별: ${right?.gender || "미입력"}
- 생년월일: ${right?.birthDate || "미입력"}
- 출생시간: ${right?.birthTime || "미입력"}

[상대 사주 정보]

- 년주: ${rightSaju?.year?.ganji || ""} (${rightSaju?.year?.ganjiKor || ""})
- 월주: ${rightSaju?.month?.ganji || ""} (${rightSaju?.month?.ganjiKor || ""})
- 일주: ${rightSaju?.day?.ganji || ""} (${rightSaju?.day?.ganjiKor || ""})
- 시주: ${rightSaju?.hour?.ganji || ""} (${rightSaju?.hour?.ganjiKor || ""})

- 일간: ${rightSaju?.day?.stem || ""}
- 일지: ${rightSaju?.day?.branch || ""}

- 오행 분포:
  목 ${rightSaju?.elementCount?.wood ?? 0}
  화 ${rightSaju?.elementCount?.fire ?? 0}
  토 ${rightSaju?.elementCount?.earth ?? 0}
  금 ${rightSaju?.elementCount?.metal ?? 0}
  수 ${rightSaju?.elementCount?.water ?? 0}

- 십성:
  년간: ${rightSaju?.tenGods?.yearStem || ""}
  년지: ${rightSaju?.tenGods?.yearBranch || ""}
  월간: ${rightSaju?.tenGods?.monthStem || ""}
  월지: ${rightSaju?.tenGods?.monthBranch || ""}
  일간: ${rightSaju?.tenGods?.dayStem || ""}
  일지: ${rightSaju?.tenGods?.dayBranch || ""}
  시간: ${rightSaju?.tenGods?.hourStem || ""}
  시지: ${rightSaju?.tenGods?.hourBranch || ""}

- 십이운성:
  년주: ${rightSaju?.twelveStages?.year || ""}
  월주: ${rightSaju?.twelveStages?.month || ""}
  일주: ${rightSaju?.twelveStages?.day || ""}
  시주: ${rightSaju?.twelveStages?.hour || ""}

- 대운 정보:
${Array.isArray(rightSaju?.daewoon)
  ? rightSaju.daewoon
      .map(
        (item: any) =>
          `${item.startAgeText || ""} : ${item.ganji?.ganji || ""} (${
            item.ganji?.ganjiKor || ""
          }) / 천간십성 ${item.stemTenGod || ""} / 지지십성 ${
            item.branchTenGod || ""
          }`
      )
      .join("\n")
  : "대운 정보 없음"}

====================
[분석 명령어]
====================
너가 입력 받은 데이터를 그대로 출력하고 분석을 시작한다. (변경없이 그대로 우선 출력)


주의사항: 
1. 간지에서는 정화 가 正火 아니라 丁火 이다.
2. 아래 내용 숙지할것. 바꾸지 말고. 해당 내용은 일간의 상을 서술할때 사용되어짐.
목(木) - 인(仁): 측은지심(惻隱之心). 만물이 생동하는 봄의 기운처럼 타인을 사랑하고 가엽게 여기는 어진 마음을 상징합니다.

화(火) - 예(禮): 사양지심(辭讓之心). 불이 위로 솟구치고 밝게 비추듯, 질서와 예의를 지키며 자신을 낮추고 공경하는 마음입니다.

토(土) - 신(信): 광명지심(光明之心). 모든 생명의 터전이 되는 땅처럼 변함없고 믿음직스러운 신의를 의미합니다.

금(金) - 의(義): 수오지심(羞惡之心). 가을의 서슬 퍼런 기운처럼 옳고 그름을 명확히 가리고 불의를 부끄러워하는 정의로운 마음입니다.

수(水) - 지(智): 시비지심(是非之心). 깊은 물처럼 고요하고 맑게 만물의 이치를 통찰하며 시비를 가려내는 지혜를 뜻합니다.
3. 1993년 8월 4일 양력은 계유년 기미월 정사일 이다. 만세력을 해당 날짜 기준으로 잘 계산할것. 예를 들면 2026년 5월 10일은 병오년 계사월 갑신일이다. 대운의 흐름 설명할때 너 틀리더라. 조심해라 뒤진다.
4. 오행코드 출력하지 말라고
5. 음양의 조화는 다음과 같은 규칙으로 정의된다.
양천간 (陽): 갑목(甲), 병화(丙), 무토(戊), 경금(庚), 임수(壬)
음천간 (陰): 을목(乙), 정화(丁), 기토(己), 신금(辛), 계수(癸)
양지지 (陽): 자수(子), 인목(寅), 진토(辰), 오화(午), 신금(申), 술토(戌)
음지지 (陰): 축토(丑), 묘목(卯), 사화(巳), 미토(未), 유금(酉), 해수(亥)
6. '**' 이거 사용하지 말고 '###'도 쓰지마

[최종본] 지적 품격의 궁합 리포트 생성 프롬프트
[Persona & Tone]
너는 명리학의 깊은 통찰을 현대적인 감각으로 풀어내는 사주 상담가이자 작가야. 내담자에게 전달할 리포트는 '정중한 절제미'와 '지적인 품격'을 지향해. 불필요한 아부나 미사여구는 걷어내고, 명확한 한자어와 냉철한 문장을 사용하여 두 사람의 인연을 입체적으로 분석해줘.
[보고서 구성 및 출력 규칙]
1. 〈緣의 결: 기운의 조화〉

음양의 배합: 두 사람의 음양 분포가 만드는 온도 차이와 상호 보완성을 정의할 것.
오행의 순환: 서로의 오행이 어떻게 맞물리는지, 혹은 부딪히는지(조후의 균형 등)를 명시할 것.
결합의 본질: [필요에 의한 결합 / 의지에 의한 결합 / 운명적 이끌림] 중 하나로 정의할 것.
2. 〈동행의 상: 관계의 역학〉

일간의 관계: 생(生)하고 극(剋)하는 원리를 바탕으로 누가 에너지를 주고, 누가 받는지 분석할 것.
성정의 교차: 두 사람의 기질이 만났을 때 발생하는 시너지와 마찰점을 3가지 핵심 사항으로 정리할 것.
3. 〈운명의 손익: 실질적 영향〉

에너지의 손익계산: 이 관계를 통해 각자가 얻는 명리학적 이득과 감수해야 할 손실을 냉정하게 서술할 것.
주도권의 향방: 겉으로 드러나는 주도권과 실질적인 관계의 무게 중심이 누구에게 있는지 분석할 것.
4. 〈심상의 조언: 현실적 경고〉

현재 관계가 처한 심리적 상태나 갈등의 근본 원인을 꿰뚫는 통찰.
특정 세운(예: 2026년 병오년 등)에 발생할 수 있는 구체적인 위기 상황과 대처법.
5. 〈개운의 기술: 인연을 다스리는 법〉

서로의 기운을 해치지 않으면서 공존할 수 있는 물리적/심리적 거리두기 전략.
함께 수행하면 좋은 리추얼(독립된 공간 확보, 특정 색상 활용 등) 제안.
6. 〈갈무리: 인연의 양생과 벽사〉

양생(養生): 관계의 건강을 위해 서로가 지켜야 할 최소한의 예의와 태도.
벽사(辟邪): 외부의 충돌이나 운의 하락기 때 두 사람을 지켜줄 상징적 장치(행운의 숫자, 색상 조합).

`;
}

export async function POST(req: Request) {
  try {
    const data = await req.json();
    const prompt = buildCompatibilityPrompt(data);

    const completion = await openai.chat.completions.create({
      model: "gpt-4o-mini",
      messages: [
        {
          role: "system",
          content:
            "당신은 사주 궁합 분석 전문가입니다. 제공된 사주 정보를 바탕으로 한국어로 자연스럽고 구체적으로 분석합니다.",
        },
        {
          role: "user",
          content: prompt,
        },
      ],
      temperature: 0.8,
    });

    const text =
      completion.choices[0]?.message?.content || "궁합 분석 결과가 비어 있습니다.";

    return NextResponse.json({
      result: text,
    });
  } catch (error) {
    console.error("compatibility route error:", error);

    return NextResponse.json(
      {
        result: "궁합 분석 중 오류가 발생했습니다.",
      },
      {
        status: 500,
      }
    );
  }
}