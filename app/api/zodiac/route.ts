import { NextResponse } from "next/server";
import { calculateBirthChart } from "../../lib/astrology/calculateBirthChart";

export async function POST(req: Request) {
  try {
    const body = await req.json();

    const { name, gender, birthDate, birthTime, birthLocation } = body;

    if (!birthDate || !birthTime || !birthLocation) {
      return NextResponse.json(
        {
          result: "생년월일, 출생 시간, 태어난 위치를 모두 입력해 주세요.",
        },
        { status: 400 }
      );
    }

    const birthChart = await calculateBirthChart({
      name,
      gender,
      birthDate,
      birthTime,
      birthLocation,
    });

    return NextResponse.json({
      success: true,
      input: {
        name,
        gender,
        birthDate,
        birthTime,
        birthLocation,
      },
      birthChart,
    });
  } catch (error) {
    console.error(error);

    return NextResponse.json(
      {
        success: false,
        result: "점성술 계산 중 오류가 발생했습니다.",
      },
      { status: 500 }
    );
  }
}