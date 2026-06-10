"use client";

import KoreanLunarCalendar from "korean-lunar-calendar";
import { useEffect, useRef, useState } from "react";
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
  const FONT = {
    // 제목
    pageTitle: "text-5xl",
    sectionTitle: "text-5xl",
    cardTitle: "text-5xl",
    panelTitle: "text-4xl",

    // 입력창 / 라벨 / 일반 설명
    inputText: "text-4xl",
    formLabel: "text-4xl",
    body: "text-4xl",
    analysisBody: "text-4xl",

    // 사주 카드
    pillarLabel: "text-3xl",
    pillarMainHanja: "text-6xl",
    pillarKor: "text-3xl",
    hiddenStem: "text-3xl",
    juGwonShin: "text-3xl",
    jeolgiLabel: "text-3xl",
    jeolgiValue: "text-3xl",
    timeUnknown: "text-3xl",

    // 십성 / 십이운성
    tenGod: "text-3xl",
    twelveState: "text-3xl",

    // 공망 / 귀인
    specialLabel: "text-3xl",
    specialValue: "text-4xl",

    // 지지 관계 / 신살
    relation: "text-3xl",
    shinsalTitle: "text-3xl",
    shinsal: "text-3xl",

    // 대운
    daewoonTitle: "text-3xl",
    daewoonAge: "text-2xl",
    daewoonHanja: "text-5xl",
    daewoonTenGod: "text-3xl",

    // 년운
    yearLuckTitle: "text-3xl",
    yearLuckAge: "text-3xl",
    yearLuckHanja: "text-5xl",
    yearLuckTenGod: "text-3xl",

    // 월운
    monthLuckTitle: "text-3xl",
    monthLuckMonth: "text-3xl",
    monthLuckHanja: "text-5xl",
    monthLuckTenGod: "text-3xl",

    // 오행
    elementTitle: "text-4xl",
    elementValue: "text-4xl",

    // 버튼
    modeButtonText: "text-3xl",
    buttonText: "text-3xl",
  };

  const COLOR = {
    pageTitle: "text-[#2b1d12]",
    sectionTitle: "text-[#2b1d12]",
    cardTitle: "text-black",
    panelTitle: "text-black",

    inputText: "text-black",
    formLabel: "text-[#6b3f24]",
    body: "text-zinc-600",
    analysisBody: "text-zinc-700",

    pillarLabel: "text-black",
    pillarKor: "text-zinc-600",
    hiddenStem: "text-zinc-600",
    juGwonShin: "text-[#6b3f24]",
    jeolgiLabel: "text-[#6b3f24]",
    jeolgiValue: "text-[#6b3f24]",
    timeUnknown: "text-zinc-400",

    tenGod: "text-zinc-700",
    twelveState: "text-[#6b3f24]",

    specialLabel: "text-zinc-700",
    specialValue: "text-[#6b3f24]",

    relation: "text-[#6b3f24]",
    shinsalTitle: "text-zinc-500",
    shinsal: "text-[#5f3a20]",

    daewoonTitle: "text-black",
    daewoonAge: "text-zinc-700",
    daewoonAgeSelected: "text-white/80",
    daewoonTenGod: "text-black",
    daewoonTenGodSelected: "text-white",

    yearLuckTitle: "text-black",
    yearLuckAge: "text-zinc-700",
    yearLuckAgeSelected: "text-white/80",
    yearLuckTenGod: "text-black",
    yearLuckTenGodSelected: "text-white",

    monthLuckTitle: "text-black",
    monthLuckMonth: "text-zinc-700",
    monthLuckTenGod: "text-black",

    elementTitle: "text-black",
    elementValue: "text-black",

    modeButtonText: "text-[#6b3f24]",
    buttonText: "text-white",
  };

  const WEIGHT = {
    pageTitle: "font-bold",
    sectionTitle: "font-bold",
    cardTitle: "font-bold",
    panelTitle: "font-bold",

    inputText: "font-normal",
    formLabel: "font-bold",
    body: "font-normal",
    analysisBody: "font-normal",

    pillarLabel: "font-bold",
    pillarMainHanja: "font-bold",
    pillarKor: "font-semibold",
    hiddenStem: "font-semibold",
    juGwonShin: "font-bold",
    jeolgiLabel: "font-bold",
    jeolgiValue: "font-bold",
    timeUnknown: "font-bold",

    tenGod: "font-bold",
    twelveState: "font-bold",

    specialLabel: "font-bold",
    specialValue: "font-bold",

    relation: "font-bold",
    shinsalTitle: "font-bold",
    shinsal: "font-bold",

    daewoonTitle: "font-bold",
    daewoonAge: "font-bold",
    daewoonHanja: "font-bold",
    daewoonTenGod: "font-bold",

    yearLuckTitle: "font-bold",
    yearLuckAge: "font-bold",
    yearLuckHanja: "font-bold",
    yearLuckTenGod: "font-bold",

    monthLuckTitle: "font-bold",
    monthLuckMonth: "font-bold",
    monthLuckHanja: "font-bold",
    monthLuckTenGod: "font-bold",

    elementTitle: "font-bold",
    elementValue: "font-bold",

    modeButtonText: "font-bold",
    buttonText: "font-bold",
  };

  const WORD_FONT = {
    title: 36,
    sectionTitle: 28,
    body: 24,
  };
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
    birthTime: "00:10",
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
      birthTime: "00:10",
      birthTimeUnknown: false,
      calendarType: "solar",
      isLeapMonth: false,
    },
    right: {
      name: "",
      gender: "여성",
      birthDate: "",
      birthTime: "00:10",
      birthTimeUnknown: false,
      calendarType: "solar",
      isLeapMonth: false,
    },
  });
  const [selectedDaewoonKey, setSelectedDaewoonKey] = useState<
    Record<string, string | null>
  >({});

  const [selectedYearLuckKey, setSelectedYearLuckKey] = useState<
    Record<string, string | null>
  >({});
  const [showCompatibilityRelations, setShowCompatibilityRelations] =
    useState(false);
  const [result, setResult] = useState("");
  const [loading, setLoading] = useState(false);
  const [showSaju, setShowSaju] = useState(false);
  const [showDailyCalendar, setShowDailyCalendar] = useState(false);
  const [calendarDate, setCalendarDate] = useState(new Date());
  const [sajuResult, setSajuResult] = useState<any>(null);
  const [compatibilityResult, setCompatibilityResult] = useState<any>({
    left: null,
    right: null,
  });
  const [recentPeople, setRecentPeople] = useState<any[]>([]);
  const [favoritePeople, setFavoritePeople] = useState<any[]>([]);
  const [peopleStorageOpen, setPeopleStorageOpen] = useState(false);
  const [favoritePeopleOpen, setFavoritePeopleOpen] = useState(true);
  const [recentPeopleOpen, setRecentPeopleOpen] = useState(true);
  const [storageLoaded, setStorageLoaded] = useState(false);
  const [memoOpen, setMemoOpen] = useState(false);
  const [memoText, setMemoText] = useState("");
  const [memoLoaded, setMemoLoaded] = useState(false);
  const [selectedHanja, setSelectedHanja] = useState<{
    type: "stem" | "branch";
    value: string;
  } | null>(null);


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
  const normalizeDateOnBlur = (value: string) => {
    return formatDateInput(value);
  };
  const normalizeDateForCalc = (value: string) => {
    const onlyNumber = value.replace(/\D/g, "");

    if (onlyNumber.length < 6) return value;

    const parts = value.split("-");

    if (parts.length === 3) {
      const year = parts[0].padStart(4, "0");
      const month = parts[1].padStart(2, "0");
      const day = parts[2].padStart(2, "0");

      return `${year}-${month}-${day}`;
    }

    return formatDateInput(value);
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

    let solarBirthDate = normalizeDateForCalc(targetForm.birthDate);

    if (targetForm.calendarType === "lunar") {
      const calendar: any = new KoreanLunarCalendar();

      const [year, month, day] = normalizeDateForCalc(targetForm.birthDate)
        .split("-")
        .map(Number);

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

    const calculatedWithBirthInfo = {
      ...calculated,
      solarBirthDate,
      birthTimeForJuGwonShin: targetForm.birthTimeUnknown
        ? "12:00"
        : targetForm.birthTime,
    };

    if (targetForm.birthTimeUnknown) {
      return makeTimeUnknownSaju(calculatedWithBirthInfo);
    }

    return calculatedWithBirthInfo;
  };

  const makePersonKey = (person: any) => {
    return [
      person.name || "이름없음",
      person.gender || "남성",
      normalizeDateForCalc(person.birthDate || ""),
      person.birthTimeUnknown ? "시간미상" : person.birthTime || "00:10",
      person.calendarType || "solar",
      person.isLeapMonth ? "윤달" : "평달",
    ].join("|");
  };

  const normalizeRecentPerson = (person: any) => {
    return {
      name: person.name || "",
      gender: person.gender || "남성",
      birthDate: normalizeDateForCalc(person.birthDate || ""),
      birthTime: person.birthTime || "00:10",
      birthTimeUnknown: person.birthTimeUnknown || false,
      calendarType: person.calendarType || "solar",
      isLeapMonth: person.isLeapMonth || false,
    };
  };

  const normalizePeopleList = (people: any[]) => {
    const seen = new Set<string>();

    return (people || [])
      .map((person) => normalizeRecentPerson(person))
      .filter((person) => person.birthDate)
      .filter((person) => {
        const key = makePersonKey(person);
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
  };

  const DEFAULT_FAVORITE_PEOPLE = [
    { name: "에스크", gender: "여성", birthDate: "1997-12-31", birthTime: "09:15", birthTimeUnknown: false, calendarType: "solar", isLeapMonth: false },
    { name: "Mgk", gender: "남성", birthDate: "1991-05-23", birthTime: "08:30", birthTimeUnknown: false, calendarType: "solar", isLeapMonth: false },
    { name: "희망", gender: "남성", birthDate: "1999-02-23", birthTime: "10:00", birthTimeUnknown: false, calendarType: "solar", isLeapMonth: false },
    { name: "유니콘", gender: "남성", birthDate: "1987-05-07", birthTime: "16:45", birthTimeUnknown: false, calendarType: "solar", isLeapMonth: false },
    { name: "우지안", gender: "여성", birthDate: "1995-11-30", birthTime: "01:29", birthTimeUnknown: false, calendarType: "solar", isLeapMonth: false },
    { name: "남태식", gender: "남성", birthDate: "1981-01-31", birthTime: "13:30", birthTimeUnknown: false, calendarType: "solar", isLeapMonth: false },
    { name: "스파크", gender: "남성", birthDate: "1986-12-25", birthTime: "11:53", birthTimeUnknown: false, calendarType: "solar", isLeapMonth: false },
    { name: "장민혁(쥬토피앙)", gender: "남성", birthDate: "1991-10-22", birthTime: "01:00", birthTimeUnknown: false, calendarType: "solar", isLeapMonth: false },
    { name: "임딩", gender: "여성", birthDate: "2002-08-13", birthTime: "11:59", birthTimeUnknown: false, calendarType: "solar", isLeapMonth: false },
    { name: "성민(민2)", gender: "남성", birthDate: "1993-03-28", birthTime: "04:30", birthTimeUnknown: false, calendarType: "solar", isLeapMonth: false },
    { name: "킹스맨", gender: "남성", birthDate: "1986-03-23", birthTime: "20:30", birthTimeUnknown: false, calendarType: "solar", isLeapMonth: false },
    { name: "뉴탄즈", gender: "남성", birthDate: "1993-08-04", birthTime: "23:00", birthTimeUnknown: false, calendarType: "solar", isLeapMonth: false },
    { name: "간지남", gender: "남성", birthDate: "1988-02-02", birthTime: "00:10", birthTimeUnknown: false, calendarType: "solar", isLeapMonth: false },
  ];

  useEffect(() => {
    if (typeof window === "undefined") return;

    try {
      const savedRecentPeople = window.localStorage.getItem("sajuRecentPeople");
      const savedFavoritePeople =
        window.localStorage.getItem("sajuFavoritePeople");
      const savedMemoText = window.localStorage.getItem("sajuMemoText");

      if (savedMemoText !== null) {
        setMemoText(savedMemoText);
      }

      setMemoLoaded(true);

      if (savedRecentPeople) {
        setRecentPeople(normalizePeopleList(JSON.parse(savedRecentPeople)));
      }

      const savedFavoriteList = savedFavoritePeople
        ? JSON.parse(savedFavoritePeople)
        : [];

      setFavoritePeople(
        normalizePeopleList([
          ...DEFAULT_FAVORITE_PEOPLE,
          ...savedFavoriteList,
        ]),
      );
    } catch (error) {
      console.error("저장된 사람 목록을 불러오지 못했습니다.", error);
    } finally {
      setMemoLoaded(true);
      setStorageLoaded(true);
    }
  }, []);

  useEffect(() => {
    if (!storageLoaded || typeof window === "undefined") return;

    window.localStorage.setItem(
      "sajuRecentPeople",
      JSON.stringify(recentPeople),
    );
  }, [recentPeople, storageLoaded]);

  useEffect(() => {
    if (!storageLoaded || typeof window === "undefined") return;

    window.localStorage.setItem(
      "sajuFavoritePeople",
      JSON.stringify(favoritePeople),
    );
  }, [favoritePeople, storageLoaded]);

  useEffect(() => {
    if (!memoLoaded || typeof window === "undefined") return;

    window.localStorage.setItem("sajuMemoText", memoText);
  }, [memoText, memoLoaded]);

  const saveRecentPerson = (person: any) => {
    if (!person?.birthDate) return;

    const normalizedPerson = normalizeRecentPerson(person);

    setRecentPeople((prev) => {
      const key = makePersonKey(normalizedPerson);

      return [
        normalizedPerson,
        ...prev.filter((item) => makePersonKey(item) !== key),
      ].slice(0, 50);
    });
  };

  const isFavoritePerson = (person: any) => {
    const key = makePersonKey(normalizeRecentPerson(person));

    return favoritePeople.some((item) => makePersonKey(item) === key);
  };

  const addFavoritePerson = (person: any) => {
    if (!person?.birthDate) return;

    const normalizedPerson = normalizeRecentPerson(person);
    const key = makePersonKey(normalizedPerson);

    setFavoritePeople((prev) =>
      [
        normalizedPerson,
        ...prev.filter((item) => makePersonKey(item) !== key),
      ].slice(0, 100),
    );
  };

  const removeFavoritePerson = (person: any) => {
    const key = makePersonKey(normalizeRecentPerson(person));

    setFavoritePeople((prev) =>
      prev.filter((item) => makePersonKey(item) !== key),
    );
  };

  const removeRecentPerson = (person: any) => {
    const key = makePersonKey(normalizeRecentPerson(person));

    setRecentPeople((prev) =>
      prev.filter((item) => makePersonKey(item) !== key),
    );
  };

  const loadRecentPersonToSaju = (person: any) => {
    const normalizedPerson = normalizeRecentPerson(person);

    setForm({
      ...form,
      ...normalizedPerson,
    });
    setSajuResult(null);
    setResult("");
  };

  const loadRecentPersonToCompatibility = (
    side: "left" | "right",
    person: any,
  ) => {
    const normalizedPerson = normalizeRecentPerson(person);

    setCompatibilityForm({
      ...compatibilityForm,
      [side]: {
        ...compatibilityForm[side],
        ...normalizedPerson,
      },
    });
    setCompatibilityResult({
      left: null,
      right: null,
    });
    setResult("");
  };

  const renderPeopleButton = (
    person: any,
    onSelect: (person: any) => void,
    options?: {
      favoriteButton?: boolean;
      removeFavoriteButton?: boolean;
      removeRecentButton?: boolean;
    },
  ) => {
    const favorite = isFavoritePerson(person);

    return (
      <div
        key={makePersonKey(person)}
        className="flex items-center gap-1 rounded-xl bg-white p-1 shadow-sm"
      >
        <button
          type="button"
          onClick={() => onSelect(person)}
          className="rounded-lg px-4 py-2 text-2xl font-bold text-[#6b3f24] transition hover:bg-[#f3e1cf]"
        >
          {person.name || "이름없음"} / {person.birthDate}
          {person.birthTimeUnknown ? " / 시간미상" : ` / ${person.birthTime}`}
        </button>

        {options?.favoriteButton && (
          <button
            type="button"
            onClick={() => addFavoritePerson(person)}
            disabled={favorite}
            className="rounded-lg px-3 py-2 text-2xl font-bold text-[#6b3f24] transition hover:bg-[#f3e1cf] disabled:opacity-40"
            title={favorite ? "이미 즐겨찾기에 저장됨" : "즐겨찾기 추가"}
          >
            {favorite ? "★" : "☆"}
          </button>
        )}

        {options?.removeFavoriteButton && (
          <button
            type="button"
            onClick={() => removeFavoritePerson(person)}
            className="rounded-lg px-3 py-2 text-2xl font-bold text-red-700 transition hover:bg-red-50"
            title="즐겨찾기 삭제"
          >
            삭제
          </button>
        )}

        {options?.removeRecentButton && (
          <button
            type="button"
            onClick={() => removeRecentPerson(person)}
            className="rounded-lg px-3 py-2 text-2xl font-bold text-red-700 transition hover:bg-red-50"
            title="최근 본 사람 삭제"
          >
            삭제
          </button>
        )}
      </div>
    );
  };

  const renderPeopleStoragePanel = (onSelect: (person: any) => void) => {
    return (
      <div className="rounded-2xl border border-[#ead8c4] bg-[#fffaf3] p-4">
        <button
          type="button"
          onClick={() => setPeopleStorageOpen((prev) => !prev)}
          className="flex w-full items-center justify-between rounded-xl bg-white px-4 py-3 text-left text-[#6b3f24] shadow-sm transition hover:bg-[#f3e1cf]"
        >
          <span className={`${FONT.formLabel} font-bold`}>
            저장한 사람 불러오기
          </span>
          <span className="text-2xl font-bold">
            {peopleStorageOpen ? "닫기 ▲" : "열기 ▼"}
          </span>
        </button>

        {peopleStorageOpen && (
          <>
            <div className="mt-4 rounded-2xl bg-white/70 p-3">
              <button
                type="button"
                onClick={() => setFavoritePeopleOpen((prev) => !prev)}
                className="flex w-full items-center justify-between text-left text-2xl font-bold text-[#6b3f24]"
              >
                <span>즐겨찾기</span>
                <span>{favoritePeopleOpen ? "▲" : "▼"}</span>
              </button>

              {favoritePeopleOpen && (
                <div className="mt-3 flex flex-wrap gap-2">
                  {favoritePeople.length > 0 ? (
                    favoritePeople.map((person) =>
                      renderPeopleButton(person, onSelect, {
                        removeFavoriteButton: true,
                      }),
                    )
                  ) : (
                    <div className="text-2xl font-bold text-zinc-400">
                      즐겨찾기한 사람이 없습니다.
                    </div>
                  )}
                </div>
              )}
            </div>

            <div className="mt-4 rounded-2xl bg-white/70 p-3">
              <div className="flex items-center justify-between">
                <button
                  type="button"
                  onClick={() => setRecentPeopleOpen((prev) => !prev)}
                  className="flex flex-1 items-center justify-between text-left text-2xl font-bold text-[#6b3f24]"
                >
                  <span>최근 본 사람</span>
                  <span>{recentPeopleOpen ? "▲" : "▼"}</span>
                </button>

                <button
                  type="button"
                  onClick={() => {
                    if (confirm("최근 본 사람 목록을 모두 삭제하시겠습니까?")) {
                      setRecentPeople([]);
                    }
                  }}
                  className="ml-3 rounded-lg px-3 py-2 text-xl font-bold text-red-700 transition hover:bg-red-50"
                >
                  전체삭제
                </button>
              </div>

              {recentPeopleOpen && (
                <div className="mt-3 flex flex-wrap gap-2">
                  {recentPeople.length > 0 ? (
                    recentPeople.map((person) =>
                      renderPeopleButton(person, onSelect, {
                        favoriteButton: true,
                        removeRecentButton: true,
                      }),
                    )
                  ) : (
                    <div className="text-2xl font-bold text-zinc-400">
                      최근 본 사람이 없습니다.
                    </div>
                  )}
                </div>
              )}
            </div>
          </>
        )}
      </div>
    );
  };

  const handleCalculateSaju = () => {
    const calculated = calculateOneSaju(form);
    if (!calculated) return;

    saveRecentPerson(form);
    setSajuResult(calculated);
  };

  const handleCalculateCompatibility = () => {
    const left = calculateOneSaju(compatibilityForm.left);
    const right = calculateOneSaju(compatibilityForm.right);

    if (left) saveRecentPerson(compatibilityForm.left);
    if (right) saveRecentPerson(compatibilityForm.right);

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
                    size: WORD_FONT.title,
                  }),
                ],
              }),

              new Paragraph({ text: "" }),

              new Paragraph({
                children: [
                  new TextRun({
                    text: `이름: ${form.name}`,
                    size: WORD_FONT.body,
                  }),
                ],
              }),

              new Paragraph({
                children: [
                  new TextRun({
                    text: `생년월일: ${form.birthDate} ${
                      form.birthTimeUnknown ? "시간 미상" : form.birthTime
                    }`,
                    size: WORD_FONT.body,
                  }),
                ],
              }),

              new Paragraph({ text: "" }),

              new Paragraph({
                children: [
                  new TextRun({
                    text: "만세력",
                    bold: true,
                    size: WORD_FONT.sectionTitle,
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
                    size: WORD_FONT.sectionTitle,
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
                    size: WORD_FONT.sectionTitle,
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
                            size: WORD_FONT.body,
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
const HANJA_OUTLINE_SIZE = 2.0;

const HANJA_STYLE = (color: string) => {
  if (color === "#000000") {
    return {
      color: "#000000",
    };
  }

  const s = HANJA_OUTLINE_SIZE;

  return {
    color,
    textShadow: `
      ${-s}px ${-s}px 0 #000,
       ${s}px ${-s}px 0 #000,
      ${-s}px  ${s}px 0 #000,
       ${s}px  ${s}px 0 #000
    `,
  };
};

const ELEMENT_HANJA_STYLE = (color: string) => {
  const s = HANJA_OUTLINE_SIZE;
  const outlineColor = color === "#000000" ? "#ffffff" : "#000000";

  return {
    color,
    textShadow: `
      ${-s}px ${-s}px 0 ${outlineColor},
       ${s}px ${-s}px 0 ${outlineColor},
      ${-s}px  ${s}px 0 ${outlineColor},
       ${s}px  ${s}px 0 ${outlineColor}
    `,
  };
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
    //{ label: "국인", value: getGukinGwiyin(targetSaju) },
  ];

  const renderSpecialInfo = (targetSaju: any) => {
    if (!targetSaju) return null;

    const gwiyinList = getGwiyinList(targetSaju);

    return (
      <div className="mt-5 grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-6">
        <div className="rounded-2xl border border-[#ead8c4] bg-[#fffaf3] p-4 text-center">
          <div
            className={`${FONT.specialLabel} ${WEIGHT.specialLabel} ${COLOR.specialLabel}`}
          >
            공망
          </div>

          <div
            className={`mt-2 ${FONT.specialValue} ${WEIGHT.specialValue} ${COLOR.specialValue}`}
          >
            {getDayGongmang(targetSaju)}
          </div>
        </div>

        {gwiyinList.map((gwiyin) => (
          <div
            key={gwiyin.label}
            className="rounded-2xl border border-[#ead8c4] bg-[#fffaf3] p-4 text-center"
          >
            <div
              className={`${FONT.specialLabel} ${WEIGHT.specialLabel} ${COLOR.specialLabel}`}
            >
              {gwiyin.label}
            </div>

            <div
              className={`mt-2 ${FONT.specialValue} ${WEIGHT.specialValue} ${COLOR.specialValue}`}
            >
              {gwiyin.value}
            </div>
          </div>
        ))}
      </div>
    );
  };

  const renderElementInfo = (targetSaju: any) => {
    if (!targetSaju) return null;

    return (
      <div className="rounded-2xl bg-zinc-100 p-4">
        <h4
          className={`${FONT.elementTitle} ${WEIGHT.elementTitle} ${COLOR.elementTitle}`}
        >

        </h4>

        <div
          className={`mt-3 grid grid-cols-5 gap-2 text-center ${FONT.elementValue}`}
        >
          <div>
            <div
              className={`${FONT.elementTitle} ${WEIGHT.elementTitle}`}
              style={ELEMENT_HANJA_STYLE(getElementColor("목"))}
            >
              木
            </div>
            <div
              className={`${FONT.elementValue} ${WEIGHT.elementValue}`}
              style={ELEMENT_HANJA_STYLE(getElementColor("목"))}
            >
              {targetSaju?.elementCount?.wood ?? 0}
            </div>
          </div>

          <div>
            <div
              className={`${FONT.elementTitle} ${WEIGHT.elementTitle}`}
              style={ELEMENT_HANJA_STYLE(getElementColor("화"))}
            >
              火
            </div>
            <div
              className={`${FONT.elementValue} ${WEIGHT.elementValue}`}
              style={ELEMENT_HANJA_STYLE(getElementColor("화"))}
            >
              {targetSaju?.elementCount?.fire ?? 0}
            </div>
          </div>

          <div>
            <div
              className={`${FONT.elementTitle} ${WEIGHT.elementTitle}`}
              style={ELEMENT_HANJA_STYLE(getElementColor("토"))}
            >
              土
            </div>
            <div
              className={`${FONT.elementValue} ${WEIGHT.elementValue}`}
              style={ELEMENT_HANJA_STYLE(getElementColor("토"))}
            >
              {targetSaju?.elementCount?.earth ?? 0}
            </div>
          </div>

          <div>
            <div
              className={`${FONT.elementTitle} ${WEIGHT.elementTitle}`}
              style={ELEMENT_HANJA_STYLE(getElementColor("금"))}
            >
              金
            </div>
            <div
              className={`${FONT.elementValue} ${WEIGHT.elementValue}`}
              style={ELEMENT_HANJA_STYLE(getElementColor("금"))}
            >
              {targetSaju?.elementCount?.metal ?? 0}
            </div>
          </div>

          <div>
            <div
              className={`${FONT.elementTitle} ${WEIGHT.elementTitle}`}
              style={ELEMENT_HANJA_STYLE(getElementColor("수"))}
            >
              水
            </div>
            <div
              className={`${FONT.elementValue} ${WEIGHT.elementValue}`}
              style={ELEMENT_HANJA_STYLE(getElementColor("수"))}
            >
              {targetSaju?.elementCount?.water ?? 0}
            </div>
          </div>
        </div>
      </div>
    );
  };

  const renderSajuOverview = (
    targetSaju: any,
    items: any[],
    cardKey: string,
  ) => {
    if (!targetSaju) return null;

    const isCompatibilityMode = mode === "compatibility";

    if (isCompatibilityMode) {
      return (
        <div className="mt-4 flex flex-col gap-4">
          <div className="grid grid-cols-4 gap-3 text-center font-bold text-black">
            {items.map((item) => renderPillarCard(item, cardKey))}
          </div>

          {renderElementInfo(targetSaju)}

          <div className="rounded-2xl border border-[#ead8c4] bg-[#fffaf3] p-4">
            <div className="grid grid-cols-5 gap-3 text-center">
              {[
                { label: "공망", value: getDayGongmang(targetSaju) },
                ...getGwiyinList(targetSaju),
              ].map((item) => (
                <div key={item.label}>
                  <div
                    className={`${FONT.specialLabel} ${WEIGHT.specialLabel} ${COLOR.specialLabel}`}
                  >
                    {item.label}
                  </div>

                  <div
                    className={`mt-2 ${FONT.specialValue} ${WEIGHT.specialValue} ${COLOR.specialValue}`}
                  >
                    {item.value}
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      );
    }

    return (
      <div className="mt-4 grid grid-cols-12 gap-4">
        <div className="col-span-9 flex flex-col gap-4">
          <div className="grid grid-cols-4 gap-3 text-center font-bold text-black">
            {items.map((item) => renderPillarCard(item, cardKey))}
          </div>

          {renderElementInfo(targetSaju)}
        </div>

        <div className="col-span-3">
          <div className="rounded-2xl border border-[#ead8c4] bg-[#fffaf3] p-4">
            <div
              className={`${FONT.specialLabel} ${WEIGHT.specialLabel} ${COLOR.specialLabel} space-y-2`}
            >
              <div className="flex items-center justify-between gap-3 border-b border-[#ead8c4] px-1 pb-2">
                <span>공망</span>
                <span
                  className={`${FONT.specialValue} ${WEIGHT.specialValue} ${COLOR.specialValue}`}
                >
                  {getDayGongmang(targetSaju)}
                </span>
              </div>

              {getGwiyinList(targetSaju).map((gwiyin) => (
                <div
                  key={gwiyin.label}
                  className="flex items-center justify-between gap-3 border-b border-[#ead8c4] px-1 pb-2"
                >
                  <span>{gwiyin.label}</span>
                  <span
                    className={`${FONT.specialValue} ${WEIGHT.specialValue} ${COLOR.specialValue}`}
                  >
                    {gwiyin.value}
                  </span>
                </div>
              ))}
            </div>
          </div>
        </div>
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

  const getBirthDateTimeForJuGwonShin = (targetSaju: any) => {
    const birthDate = String(targetSaju?.solarBirthDate || "");
    const birthTime = String(targetSaju?.birthTimeForJuGwonShin || "12:00");

    if (!/^\d{4}-\d{2}-\d{2}$/.test(birthDate)) return null;

    const [year, month, day] = birthDate.split("-").map(Number);
    const [hour = 12, minute = 0] = birthTime.split(":").map(Number);

    return new Date(year, month - 1, day, hour, minute, 0, 0);
  };

  const getLatestJeolipDate = (birthDateTime: Date, monthBranch: string) => {
    const normalizedMonthBranch = normalizeBranch(monthBranch);
    const birthYear = birthDateTime.getFullYear();

    const candidates = [birthYear - 1, birthYear, birthYear + 1]
      .flatMap((year) => getSolarTermsOfYear(year))
      .filter((term) => {
        return MONTH_BRANCH_BY_SOLAR_TERM[term.name] === normalizedMonthBranch;
      })
      .map((term) => term.date)
      .filter((date) => date.getTime() <= birthDateTime.getTime())
      .sort((a, b) => b.getTime() - a.getTime());

    return candidates[0] || null;
  };

  const getDaysAfterJeolip = (targetSaju: any) => {
    const birthDateTime = getBirthDateTimeForJuGwonShin(targetSaju);
    const monthBranch = getMonthBranch(targetSaju);

    if (!birthDateTime || !monthBranch) return null;

    const jeolipDate = getLatestJeolipDate(birthDateTime, monthBranch);

    if (!jeolipDate) return null;

    return Math.floor(
      (birthDateTime.getTime() - jeolipDate.getTime()) / (1000 * 60 * 60 * 24),
    );
  };

  const getJuGwonShin = (branch: string, daysAfterJeolip: number | null) => {
    const normalizedBranch = normalizeBranch(branch);

    const saengji: Record<string, string> = {
      인: "갑",
      신: "경",
      사: "병",
      해: "임",
    };

    if (saengji[normalizedBranch]) return saengji[normalizedBranch];

    if (daysAfterJeolip === null) return "";

    const wangjiEarly: Record<string, string> = {
      자: "임",
      오: "병",
      묘: "갑",
      유: "경",
    };
    const wangjiLate: Record<string, string> = {
      자: "계",
      오: "정",
      묘: "을",
      유: "신",
    };

    if (wangjiEarly[normalizedBranch]) {
      return daysAfterJeolip <= 6
        ? wangjiEarly[normalizedBranch]
        : wangjiLate[normalizedBranch];
    }

    const gojiEarly: Record<string, string> = {
      진: "을",
      술: "신",
      축: "계",
      미: "정",
    };
    const gojiLate: Record<string, string> = {
      진: "무",
      술: "무",
      축: "기",
      미: "기",
    };

    if (gojiEarly[normalizedBranch]) {
      return daysAfterJeolip <= 12
        ? gojiEarly[normalizedBranch]
        : gojiLate[normalizedBranch];
    }

    return "";
  };

  const getJuGwonShinForItem = (item: any) => {
    if (item?.label !== "월주") return "";

    const branch = getItemBranch(item);
    const daysAfterJeolip = getDaysAfterJeolip(item?.targetSaju);
    const juGwonShin = getJuGwonShin(branch, daysAfterJeolip);

    return juGwonShin ? STEM_HANJA[juGwonShin] || juGwonShin : "";
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

  const STEM_HAP_PAIRS = [
    ["갑", "기"],
    ["을", "경"],
    ["병", "신"],
    ["정", "임"],
    ["무", "계"],
  ];

  const STEM_CHUNG_PAIRS = [
    ["갑", "경"],
    ["을", "신"],
    ["병", "임"],
    ["정", "계"],
  ];

  const makeHanjaPairKey = (a: string, b: string) =>
    [a, b].sort().join("-");

  const handleHanjaClick = (type: "stem" | "branch", value: string) => {
    const normalizedValue =
      type === "stem" ? normalizeStem(value) : normalizeBranch(value);

    setSelectedHanja((prev) => {
      if (prev?.type === type && prev.value === normalizedValue) return null;

      return {
        type,
        value: normalizedValue,
      };
    });
  };

  const isHapWithSelectedHanja = (
    type: "stem" | "branch",
    value: string,
  ) => {
    if (!selectedHanja || selectedHanja.type !== type) return false;

    const normalizedValue =
      type === "stem" ? normalizeStem(value) : normalizeBranch(value);

    if (normalizedValue === selectedHanja.value) return false;

    const pairKey = makeHanjaPairKey(selectedHanja.value, normalizedValue);

    if (type === "stem") {
      return STEM_HAP_PAIRS.some(
        ([a, b]) => makeHanjaPairKey(a, b) === pairKey,
      );
    }

    const isYukhap = BRANCH_RELATION_RULES.yukhap.some(
      ([a, b]: string[]) => makeHanjaPairKey(a, b) === pairKey,
    );

    const isAmhap = BRANCH_RELATION_RULES.amhap.some(
      ([a, b]: string[]) => makeHanjaPairKey(a, b) === pairKey,
    );

    const isSamhap = BRANCH_RELATION_RULES.samhap.some((rule: any) => {
      return (
        rule.branches.includes(selectedHanja.value) &&
        rule.branches.includes(normalizedValue)
      );
    });

    const isBanghap = BRANCH_RELATION_RULES.banghap.some((rule: any) => {
      return (
        rule.branches.includes(selectedHanja.value) &&
        rule.branches.includes(normalizedValue)
      );
    });

    return isYukhap || isAmhap || isSamhap || isBanghap;
  };

  const isChungWithSelectedHanja = (
    type: "stem" | "branch",
    value: string,
  ) => {
    if (!selectedHanja || selectedHanja.type !== type) return false;

    const normalizedValue =
      type === "stem" ? normalizeStem(value) : normalizeBranch(value);

    if (normalizedValue === selectedHanja.value) return false;

    const pairKey = makeHanjaPairKey(selectedHanja.value, normalizedValue);

    if (type === "stem") {
      return STEM_CHUNG_PAIRS.some(
        ([a, b]) => makeHanjaPairKey(a, b) === pairKey,
      );
    }

    return BRANCH_RELATION_RULES.chung.some(
      ([a, b]: string[]) => makeHanjaPairKey(a, b) === pairKey,
    );
  };

  const getHanjaRelationClass = (
    type: "stem" | "branch",
    value: string,
  ) => {
    if (!selectedHanja) return "";

    const normalizedValue =
      type === "stem" ? normalizeStem(value) : normalizeBranch(value);

    if (selectedHanja.type === type && selectedHanja.value === normalizedValue) {
      return "ring-4 ring-[#2b1d12] ring-offset-2 ring-offset-white shadow-md";
    }

    if (isChungWithSelectedHanja(type, normalizedValue)) {
      return "bg-red-200 ring-4 ring-red-500/40 shadow-md";
    }

    if (isHapWithSelectedHanja(type, normalizedValue)) {
      return "bg-sky-200 ring-4 ring-sky-500/40 shadow-md";
    }

    return "";
  };

  const renderHanjaButton = ({
    type,
    value,
    element,
    className,
    children,
  }: any) => {
    const normalizedValue =
      type === "stem" ? normalizeStem(value) : normalizeBranch(value);

    return (
      <span
        role="button"
        tabIndex={0}
        onClick={(event) => {
          event.stopPropagation();
          handleHanjaClick(type, normalizedValue);
        }}
        onKeyDown={(event) => {
          if (event.key === "Enter" || event.key === " ") {
            event.preventDefault();
            event.stopPropagation();
            handleHanjaClick(type, normalizedValue);
          }
        }}
        className={`inline-block cursor-pointer rounded-xl px-3 py-1 transition ${getHanjaRelationClass(
          type,
          normalizedValue,
        )}`}
        title="클릭하면 충/합 관계를 표시합니다"
      >
        <span className={className} style={HANJA_STYLE(getElementColor(element))}>
          {children}
        </span>
      </span>
    );
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
  const DAY_STEMS = [
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
  const DAY_BRANCHES = [
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

  const getDayGanji = (date: Date) => {
    const baseDate = new Date(1936, 1, 12);

    const diffDays = Math.floor(
      (date.getTime() - baseDate.getTime()) / (1000 * 60 * 60 * 24),
    );

    const stem = DAY_STEMS[((diffDays % 10) + 10) % 10];
    const branch = DAY_BRANCHES[((diffDays % 12) + 12) % 12];

    return {
      stem,
      branch,
    };
  };
  const SOLAR_TERM_NAMES = [
    "소한",
    "대한",
    "입춘",
    "우수",
    "경칩",
    "춘분",
    "청명",
    "곡우",
    "입하",
    "소만",
    "망종",
    "하지",
    "소서",
    "대서",
    "입추",
    "처서",
    "백로",
    "추분",
    "한로",
    "상강",
    "입동",
    "소설",
    "대설",
    "동지",
  ];
  const MONTH_BRANCH_BY_SOLAR_TERM: Record<string, string> = {
    소한: "축",
    입춘: "인",
    경칩: "묘",
    청명: "진",
    입하: "사",
    망종: "오",
    소서: "미",
    입추: "신",
    백로: "유",
    한로: "술",
    입동: "해",
    대설: "자",
  };
  const degToRad = (deg: number) => (deg * Math.PI) / 180;

  const normalizeDegree = (deg: number) => {
    return ((deg % 360) + 360) % 360;
  };

  const normalizeDegree180 = (deg: number) => {
    const normalized = normalizeDegree(deg);
    return normalized > 180 ? normalized - 360 : normalized;
  };

  const dateToJulianDay = (date: Date) => {
    return date.getTime() / 86400000 + 2440587.5;
  };

  const getSunApparentLongitude = (date: Date) => {
    const jd = dateToJulianDay(date);
    const t = (jd - 2451545.0) / 36525;

    const meanLongitude = normalizeDegree(
      280.46646 + 36000.76983 * t + 0.0003032 * t * t,
    );

    const meanAnomaly = normalizeDegree(
      357.52911 + 35999.05029 * t - 0.0001537 * t * t,
    );

    const equationOfCenter =
      (1.914602 - 0.004817 * t - 0.000014 * t * t) *
        Math.sin(degToRad(meanAnomaly)) +
      (0.019993 - 0.000101 * t) * Math.sin(degToRad(2 * meanAnomaly)) +
      0.000289 * Math.sin(degToRad(3 * meanAnomaly));

    const trueLongitude = meanLongitude + equationOfCenter;
    const omega = 125.04 - 1934.136 * t;
    const apparentLongitude =
      trueLongitude - 0.00569 - 0.00478 * Math.sin(degToRad(omega));

    return normalizeDegree(apparentLongitude);
  };

  const getApproxSolarTermDate = (year: number, termIndex: number) => {
    const y = year - 1900;

    const minutes =
      525948.76 * y +
      6.2 +
      15.2184 * 24 * 60 * termIndex -
      1.9 * Math.sin(degToRad(0.262 * y));

    const base = new Date(1900, 0, 6, 2, 5);
    return new Date(base.getTime() + minutes * 60 * 1000);
  };

  const getSolarTermTargetLongitude = (termIndex: number) => {
    return normalizeDegree(285 + 15 * termIndex);
  };

  const getSolarTermDate = (year: number, termIndex: number) => {
    const targetLongitude = getSolarTermTargetLongitude(termIndex);
    const approxDate = getApproxSolarTermDate(year, termIndex);

    let left = new Date(approxDate.getTime() - 10 * 24 * 60 * 60 * 1000);
    let right = new Date(approxDate.getTime() + 10 * 24 * 60 * 60 * 1000);

    for (let index = 0; index < 80; index += 1) {
      const mid = new Date((left.getTime() + right.getTime()) / 2);
      const diff = normalizeDegree180(
        getSunApparentLongitude(mid) - targetLongitude,
      );

      if (diff < 0) {
        left = mid;
      } else {
        right = mid;
      }
    }

    return right;
  };

  const getSolarTermsOfYear = (year: number) => {
    return SOLAR_TERM_NAMES.map((name, index) => ({
      name,
      date: getSolarTermDate(year, index),
    }));
  };

  const formatTermTime = (date: Date) => {
    return `${String(date.getHours()).padStart(2, "0")}:${String(
      date.getMinutes(),
    ).padStart(2, "0")}`;
  };

  const getSolarTermsForDate = (date: Date) => {
    const year = date.getFullYear();

    return getSolarTermsOfYear(year).filter((term) => {
      return (
        term.date.getFullYear() === date.getFullYear() &&
        term.date.getMonth() === date.getMonth() &&
        term.date.getDate() === date.getDate()
      );
    });
  };

  const buildDailyCalendar = (date: Date) => {
    const year = date.getFullYear();
    const month = date.getMonth();

    const firstDay = new Date(year, month, 1);
    const lastDay = new Date(year, month + 1, 0);

    const startWeekDay = firstDay.getDay();
    const totalDays = lastDay.getDate();

    const blanks = Array.from({ length: startWeekDay }, () => null);

    const days = Array.from({ length: totalDays }, (_, index) => {
      const currentDate = new Date(year, month, index + 1);

      return {
        day: index + 1,
        ganji: getDayGanji(currentDate),
        solarTerms: getSolarTermsForDate(currentDate),
      };
    });

    return [...blanks, ...days];
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
          targetSaju,
          tenGodStem: targetSaju.tenGods?.hourStem ?? "",
          tenGodBranch: targetSaju.tenGods?.hourBranch ?? "",
          twelveStage: targetSaju.twelveStages?.hour ?? "",
        },
        {
          label: "일주",
          data: targetSaju.day,
          targetSaju,
          tenGodStem: targetSaju.tenGods.dayStem,
          tenGodBranch: targetSaju.tenGods.dayBranch,
          twelveStage: targetSaju.twelveStages.day,
        },
        {
          label: "월주",
          data: targetSaju.month,
          targetSaju,
          tenGodStem: targetSaju.tenGods.monthStem,
          tenGodBranch: targetSaju.tenGods.monthBranch,
          twelveStage: targetSaju.twelveStages.month,
        },
        {
          label: "년주",
          data: targetSaju.year,
          targetSaju,
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
          <div
            className={`${FONT.pillarLabel} ${WEIGHT.pillarLabel} ${COLOR.pillarLabel}`}
          >
            {item.label}
          </div>

          <div
            className={`mt-8 flex min-h-[150px] items-center justify-center rounded-xl border border-dashed border-zinc-300 bg-white/60 ${FONT.timeUnknown} ${WEIGHT.timeUnknown} ${COLOR.timeUnknown}`}
          >
            시간 미상
          </div>
        </div>
      );
    }

    return (
      <div key={item.label} className="rounded-xl bg-zinc-100 p-3">
        <div
          className={`${FONT.pillarLabel} ${WEIGHT.pillarLabel} ${COLOR.pillarLabel}`}
        >
          {item.label}
        </div>

        <div className="mt-2 flex flex-col items-center">
          {renderHanjaButton({
            type: "stem",
            value: item.data.stem,
            element: item.data.stemElement,
            className: `${FONT.pillarMainHanja} ${WEIGHT.pillarMainHanja} leading-none`,
            children: item.data.stem,
          })}

          {renderHanjaButton({
            type: "branch",
            value: item.data.branch,
            element: item.data.branchElement,
            className: `mt-2 ${FONT.pillarMainHanja} ${WEIGHT.pillarMainHanja} leading-none`,
            children: item.data.branch,
          })}

          {getJuGwonShinForItem(item) && (
            <div className="mt-1">
              <span
                className={`${FONT.jeolgiLabel} ${WEIGHT.jeolgiLabel} ${COLOR.jeolgiLabel}`}
              >
                主
              </span>{" "}
              <span
                className={`${FONT.jeolgiValue} ${WEIGHT.jeolgiValue} ${COLOR.jeolgiValue}`}
              >
                {getJuGwonShinForItem(item)}
              </span>
            </div>
          )}

          <div
            className={`mt-3 ${FONT.pillarKor} ${WEIGHT.pillarKor} ${COLOR.pillarKor}`}
          >
            {item.data.stemKor}
            {item.data.branchKor}
          </div>

          <div
            className={`mt-1 rounded-lg bg-white/70 px-2 py-1 ${FONT.hiddenStem} ${WEIGHT.hiddenStem} ${COLOR.hiddenStem}`}
          >
            {getHiddenStemsText(item.data.branch)}
          </div>
        </div>

        <div className={`mt-3 ${FONT.tenGod} ${WEIGHT.tenGod} ${COLOR.tenGod}`}>
          {item.tenGodStem}
        </div>

        <div className={`${FONT.tenGod} ${WEIGHT.tenGod} ${COLOR.tenGod}`}>
          {item.tenGodBranch}
        </div>
        <div
          className={`mt-2 ${FONT.twelveState} ${WEIGHT.twelveState} ${COLOR.twelveState}`}
        >
          {item.twelveStage}
        </div>

        {item.branchRelations?.length > 0 && showCompatibilityRelations && (
          <div className="mt-2 flex flex-wrap justify-center gap-1">
            {item.branchRelations.map((relation: string) => (
              <span
                key={relation}
                className={`rounded-full border border-[#ead8c4] bg-[#fffaf3] px-2 py-0.5 ${FONT.relation} ${WEIGHT.relation} ${COLOR.relation}`}
              >
                {relation}
              </span>
            ))}
          </div>
        )}

        {item.shinsals?.length > 0 && showCompatibilityRelations && (
          <div className="mt-2 border-t border-[#ead8c4] pt-2">
            <div
              className={`mb-1 text-center ${FONT.shinsalTitle} ${WEIGHT.shinsalTitle} ${COLOR.shinsalTitle}`}
            >
              신살
            </div>

            <div className="flex flex-wrap justify-center gap-1">
              {item.shinsals.map((shinsal: string) => (
                <span
                  key={shinsal}
                  className={`rounded-full border border-[#d7c4ad] bg-white px-2 py-0.5 ${FONT.shinsal} ${WEIGHT.shinsal} ${COLOR.shinsal}`}
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
        <h3
          className={`${FONT.daewoonTitle} ${WEIGHT.daewoonTitle} ${COLOR.daewoonTitle}`}
        >
          대운
        </h3>

        <div className="mt-4 grid grid-cols-10 gap-3">
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
                className={`rounded-xl p-3 text-center transition ${
                  selected
                    ? "bg-[#6b3f24] text-white shadow-md"
                    : "bg-zinc-100 text-black hover:bg-zinc-200"
                }`}
              >
                <div
                  className={
                    selected
                      ? `${FONT.daewoonAge} ${WEIGHT.daewoonAge} ${COLOR.daewoonAgeSelected}`
                      : `${FONT.daewoonAge} ${WEIGHT.daewoonAge} ${COLOR.daewoonAge}`
                  }
                >
                  {item.startAgeText}
                </div>

                <div className="mt-2 flex flex-col items-center">
                  {renderHanjaButton({
                    type: "stem",
                    value: item.ganji.stem,
                    element: item.ganji.stemElement,
                    className: `${FONT.daewoonHanja} font-bold`,
                    children: item.ganji.stem,
                  })}

                  {renderHanjaButton({
                    type: "branch",
                    value: item.ganji.branch,
                    element: item.ganji.branchElement,
                    className: `${FONT.daewoonHanja} font-bold`,
                    children: item.ganji.branch,
                  })}
                </div>

                <div
                  className={
                    selected
                      ? `mt-2 ${FONT.daewoonTenGod} ${WEIGHT.daewoonTenGod} ${COLOR.daewoonTenGodSelected}`
                      : `mt-2 ${FONT.daewoonTenGod} ${WEIGHT.daewoonTenGod} ${COLOR.daewoonTenGod}`
                  }
                >
                  {item.stemTenGod}
                </div>

                <div
                  className={
                    selected
                      ? `${FONT.daewoonTenGod} ${WEIGHT.daewoonTenGod} ${COLOR.daewoonTenGodSelected}`
                      : `${FONT.daewoonTenGod} ${WEIGHT.daewoonTenGod} ${COLOR.daewoonTenGod}`
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
            <h4
              className={`${FONT.yearLuckTitle} ${WEIGHT.yearLuckTitle} ${COLOR.yearLuckTitle}`}
            >
              선택한 대운의 년운
            </h4>

            <div
              className={
                mode === "compatibility"
                  ? "mt-4 flex gap-3 overflow-x-auto pb-3"
                  : "mt-4 grid grid-cols-2 gap-3 sm:grid-cols-10"
              }
            >
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
                    className={`${
                      mode === "compatibility" ? "min-w-[100px]" : ""
                    } rounded-xl p-3 text-center shadow-sm transition ${
                      selected
                        ? "bg-[#6b3f24] text-white shadow-md"
                        : "bg-white text-black hover:bg-zinc-50"
                    }`}
                  >
                    <div
                      className={
                        selected
                          ? `${FONT.yearLuckAge} ${WEIGHT.yearLuckAge} ${COLOR.yearLuckAgeSelected}`
                          : `${FONT.yearLuckAge} ${WEIGHT.yearLuckAge} ${COLOR.yearLuckAge}`
                      }
                    >
                      {yearLuck.year} <br /> {yearLuck.age}세
                    </div>

                    <div className="mt-2 flex flex-col items-center leading-tight">
                      {renderHanjaButton({
                        type: "stem",
                        value: yearLuck.ganji.stem,
                        element: yearLuck.ganji.stemElement,
                        className: `${FONT.yearLuckHanja} font-bold leading-none`,
                        children: STEM_HANJA[yearLuck.ganji.stem],
                      })}

                      {renderHanjaButton({
                        type: "branch",
                        value: yearLuck.ganji.branch,
                        element: yearLuck.ganji.branchElement,
                        className: `mt-1 ${FONT.yearLuckHanja} font-bold leading-none`,
                        children: BRANCH_HANJA[yearLuck.ganji.branch],
                      })}

                      <div
                        className={
                          selected
                            ? `${FONT.yearLuckTenGod} ${WEIGHT.yearLuckTenGod} ${COLOR.yearLuckTenGodSelected}`
                            : `${FONT.yearLuckTenGod} ${WEIGHT.yearLuckTenGod} ${COLOR.yearLuckTenGod}`
                        }
                      >
                        {yearLuck.stemTenGod}
                      </div>

                      <div
                        className={
                          selected
                            ? `${FONT.yearLuckTenGod} ${WEIGHT.yearLuckTenGod} ${COLOR.yearLuckTenGodSelected}`
                            : `${FONT.yearLuckTenGod} ${WEIGHT.yearLuckTenGod} ${COLOR.yearLuckTenGod}`
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
                <h4
                  className={`${FONT.monthLuckTitle} ${WEIGHT.monthLuckTitle} ${COLOR.monthLuckTitle}`}
                >
                  {selectedYearLuck.year}년 월운
                </h4>

                <div
                  className={
                    mode === "compatibility"
                      ? "mt-4 flex gap-3 overflow-x-auto pb-3"
                      : "mt-4 grid grid-cols-2 gap-3 sm:grid-cols-12"
                  }
                >
                  {monthLuckList.map((monthLuck: any) => (
                    <div
                      key={`${monthLuck.year}-${monthLuck.month}`}
                      className={
                        mode === "compatibility"
                          ? "min-w-[100px] rounded-xl bg-zinc-100 p-3 text-center shadow-sm"
                          : "rounded-xl bg-zinc-100 p-3 text-center shadow-sm"
                      }
                    >
                      <div
                        className={`${FONT.monthLuckMonth} ${WEIGHT.monthLuckMonth} ${COLOR.monthLuckMonth}`}
                      >
                        {monthLuck.month}월
                      </div>

                      <div className="mt-2 flex flex-col items-center leading-tight">
                        {renderHanjaButton({
                          type: "stem",
                          value: monthLuck.ganji.stem,
                          element: monthLuck.ganji.stemElement,
                          className: `${FONT.monthLuckHanja} font-bold leading-none`,
                          children: STEM_HANJA[monthLuck.ganji.stem],
                        })}

                        {renderHanjaButton({
                          type: "branch",
                          value: monthLuck.ganji.branch,
                          element: monthLuck.ganji.branchElement,
                          className: `mt-1 ${FONT.monthLuckHanja} font-bold leading-none`,
                          children: BRANCH_HANJA[monthLuck.ganji.branch],
                        })}

                        <div
                          className={`${FONT.monthLuckTenGod} ${WEIGHT.monthLuckTenGod} ${COLOR.monthLuckTenGod}`}
                        >
                          {monthLuck.stemTenGod}
                        </div>

                        <div
                          className={`${FONT.monthLuckTenGod} ${WEIGHT.monthLuckTenGod} ${COLOR.monthLuckTenGod}`}
                        >
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
        <h3
          className={`${FONT.cardTitle} ${WEIGHT.cardTitle} ${COLOR.cardTitle}`}
        >
          사주팔자
        </h3>

        {renderSajuOverview(targetSaju, items, cardKey)}

        {renderLuckPanel(targetSaju, cardKey, birthDate)}
      </div>
    );
  };

  return (
    <main className="min-h-screen bg-[#f7efe3] px-5 py-10 text-[#2b1d12]">
      <div className="mx-auto w-[1400px] min-w-[1400px] rounded-3xl bg-white p-6 shadow-xl">
        <h1
          className={`text-center ${FONT.pageTitle} ${WEIGHT.pageTitle} ${COLOR.pageTitle}`}
        >
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
            className={`rounded-xl py-3 ${FONT.modeButtonText} ${WEIGHT.modeButtonText} ${COLOR.modeButtonText} transition ${
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
            className={`rounded-xl py-3 ${FONT.modeButtonText} ${WEIGHT.modeButtonText} ${COLOR.modeButtonText} transition ${
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
            className={`rounded-xl py-3 ${FONT.modeButtonText} ${WEIGHT.modeButtonText} ${COLOR.modeButtonText} transition ${
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
                className={`w-full rounded-xl border p-3 ${FONT.inputText}`}
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
                className={`w-full rounded-xl border p-3 ${FONT.inputText}`}
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
                  placeholder="1988-02-02"
                  className={`flex-1 rounded-xl border p-3 ${FONT.inputText}`}
                  value={form.birthDate}
                  onChange={(e) => {
                    setForm({
                      ...form,
                      birthDate: e.target.value,
                    });

                    setSajuResult(null);
                    setResult("");
                  }}
                  onBlur={(e) => {
                    setForm({
                      ...form,
                      birthDate: normalizeDateOnBlur(e.target.value),
                    });
                  }}
                />

                <label
                  className={`flex items-center gap-1 ${FONT.formLabel} font-bold`}
                >
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

                <label
                  className={`flex items-center gap-1 ${FONT.formLabel} font-bold`}
                >
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
                  <label
                    className={`flex items-center gap-1 ${FONT.formLabel} font-bold text-[#6b3f24]`}
                  >
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
                placeholder="00:10"
                className={`w-full rounded-xl border p-3 ${FONT.inputText} disabled:bg-zinc-100 disabled:text-zinc-400`}
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

              <label
                className={`flex items-center gap-2 rounded-xl border border-[#ead8c4] bg-[#fffaf3] px-4 py-3 ${FONT.formLabel} font-bold text-[#6b3f24]`}
              >
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

              {mode === "saju" &&
                renderPeopleStoragePanel(loadRecentPersonToSaju)}

              {mode === "zodiac" && (
                <input
                  type="text"
                  placeholder="태어난 위치 예: 서울, 대한민국"
                  className={`w-full rounded-xl border p-3 ${FONT.inputText}`}
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
                  <h2
                    className={`mb-4 text-center ${FONT.sectionTitle} ${WEIGHT.sectionTitle} ${COLOR.sectionTitle}`}
                  >
                    {key === "left" ? "본인" : "상대"}
                  </h2>

                  <div className="space-y-4">
                    <input
                      className={`w-full rounded-xl border p-3 ${FONT.inputText}`}
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
                      className={`w-full rounded-xl border p-3 ${FONT.inputText}`}
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
                        placeholder="1988-02-02"
                        className={`flex-1 rounded-xl border p-3 ${FONT.inputText}`}
                        value={compatibilityForm[key].birthDate}
                        onChange={(e) => {
                          setCompatibilityForm({
                            ...compatibilityForm,
                            [key]: {
                              ...compatibilityForm[key],
                              birthDate: e.target.value,
                            },
                          });

                          setCompatibilityResult({
                            left: null,
                            right: null,
                          });

                          setResult("");
                        }}
                        onBlur={(e) => {
                          setCompatibilityForm({
                            ...compatibilityForm,
                            [key]: {
                              ...compatibilityForm[key],
                              birthDate: normalizeDateOnBlur(e.target.value),
                            },
                          });
                        }}
                      />

                      <label
                        className={`flex items-center gap-1 ${FONT.formLabel} font-bold`}
                      >
                        <input
                          type="checkbox"
                          checked={
                            compatibilityForm[key].calendarType === "solar"
                          }
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

                      <label
                        className={`flex items-center gap-1 ${FONT.formLabel} font-bold`}
                      >
                        <input
                          type="checkbox"
                          checked={
                            compatibilityForm[key].calendarType === "lunar"
                          }
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
                        <label
                          className={`flex items-center gap-1 ${FONT.formLabel} font-bold text-[#6b3f24]`}
                        >
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
                      placeholder="00:10"
                      className={`w-full rounded-xl border p-3 ${FONT.inputText} disabled:bg-zinc-100 disabled:text-zinc-400`}
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

                    <label
                      className={`flex items-center gap-2 rounded-xl border border-[#ead8c4] bg-white px-4 py-3 ${FONT.formLabel} font-bold text-[#6b3f24]`}
                    >
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

                    {renderPeopleStoragePanel((person) =>
                      loadRecentPersonToCompatibility(key, person),
                    )}
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
                className={`w-full rounded-xl border border-[#6b3f24]/40 bg-[#fff7ed] py-4 ${FONT.buttonText} font-bold text-[#6b3f24] shadow-sm transition hover:bg-[#f3e1cf]`}
              >
                {showSaju ? "만세력 닫기" : "만세력 보기"}
              </button>
              <button
                type="button"
                onClick={() => setShowDailyCalendar(true)}
                className={`w-full rounded-xl border border-[#6b3f24]/40 bg-[#fff7ed] py-4 ${FONT.buttonText} font-bold text-[#6b3f24] shadow-sm transition hover:bg-[#f3e1cf]`}
              >
                일진달력 보기
              </button>

              {/*<button
                type="button"
                onClick={handleSubmit}
                disabled={loading}
                className={`w-full cursor-pointer rounded-xl bg-[#6b3f24] py-4 ${FONT.buttonText} ${WEIGHT.buttonText} ${COLOR.buttonText} disabled:opacity-50`}
              >
                {loading ? "분석 중..." : "사주 분석하기"}
              </button>*/}
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
                className={`w-full rounded-xl border border-[#6b3f24]/40 bg-[#fff7ed] py-4 ${FONT.buttonText} font-bold text-[#6b3f24] shadow-sm transition hover:bg-[#f3e1cf] disabled:opacity-40`}
              >
                두 사람 만세력 보기
              </button>

              {/*<button
                type="button"
                onClick={handleSubmit}
                disabled={loading}
                className={`w-full cursor-pointer rounded-xl bg-[#6b3f24] py-4 ${FONT.buttonText} ${WEIGHT.buttonText} ${COLOR.buttonText} disabled:opacity-50`}
              >
                {loading ? "분석 중..." : "궁합 분석하기"}
              </button>*/}
            </>
          )}

          {mode === "zodiac" && (
            <button
              type="button"
              onClick={handleSubmit}
              disabled={loading}
              className={`w-full cursor-pointer rounded-xl bg-[#6b3f24] py-4 ${FONT.buttonText} ${WEIGHT.buttonText} ${COLOR.buttonText} disabled:opacity-50`}
            >
              {loading ? "분석 중..." : "점성술 분석하기"}
            </button>
          )}

          {mode === "saju" && showSaju && (
            <section className="mt-6 rounded-3xl border border-[#ead8c4] bg-[#fffaf3] p-5 shadow-inner">
              <h2
                className={`${FONT.sectionTitle} ${WEIGHT.sectionTitle} ${COLOR.sectionTitle}`}
              >
                만세력 계산
              </h2>

              <p className={`mt-2 ${FONT.body} ${WEIGHT.body} ${COLOR.body}`}>
                위에 입력한 생년월일시를 기준으로 사주팔자와 오행 분포를
                계산합니다.
              </p>

              <button
                type="button"
                onClick={handleCalculateSaju}
                disabled={
                  !form.birthDate || (!form.birthTime && !form.birthTimeUnknown)
                }
                className={`mt-5 w-full rounded-xl bg-[#2b1d12] px-5 py-3 ${FONT.buttonText} ${WEIGHT.buttonText} ${COLOR.buttonText} transition hover:bg-[#4a2f1c] disabled:opacity-40`}
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
                    className={`mt-4 w-full rounded-xl bg-black px-5 py-3 ${FONT.buttonText} ${WEIGHT.buttonText} ${COLOR.buttonText} shadow-md`}
                  >
                    저장하기
                  </button>

                  <div className="mt-4 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                    <h3
                      className={`${FONT.cardTitle} ${WEIGHT.cardTitle} ${COLOR.cardTitle}`}
                    >
                      사주팔자
                    </h3>

                    <button
                      type="button"
                      onClick={() =>
                        setShowCompatibilityRelations((prev) => !prev)
                      }
                      className={`rounded-full border border-[#6b3f24]/40 bg-white px-4 py-2 ${FONT.buttonText} ${WEIGHT.buttonText} text-[#6b3f24] shadow-sm transition hover:bg-[#f3e1cf]`}
                    >
                      {showCompatibilityRelations
                        ? "지지 관계·신살 전체 접기 ▲"
                        : "지지 관계·신살 전체 열기 ▼"}
                    </button>
                  </div>

                  {renderSajuOverview(sajuResult, sajuItems, "main")}

                  {renderLuckPanel(
                    sajuResult,
                    "main",
                    normalizeDateForCalc(form.birthDate),
                  )}

                  {result && (
                    <div className="mt-5 rounded-2xl bg-zinc-100 p-4">
                      <h4
                        className={`${FONT.panelTitle} ${WEIGHT.panelTitle} ${COLOR.panelTitle}`}
                      >
                        사주 해석
                      </h4>

                      <div
                        className={`mt-3 whitespace-pre-wrap leading-7 ${FONT.analysisBody} ${WEIGHT.analysisBody} ${COLOR.analysisBody}`}
                      >
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
                  <h2
                    className={`${FONT.sectionTitle} ${WEIGHT.sectionTitle} ${COLOR.sectionTitle}`}
                  >
                    두 사람 만세력
                  </h2>

                  <button
                    type="button"
                    onClick={() =>
                      setShowCompatibilityRelations((prev) => !prev)
                    }
                    className={`rounded-full border border-[#6b3f24]/40 bg-white px-4 py-2 ${FONT.buttonText} font-bold text-[#6b3f24] shadow-sm transition hover:bg-[#f3e1cf]`}
                  >
                    {showCompatibilityRelations
                      ? "지지 관계·신살 전체 접기 ▲"
                      : "지지 관계·신살 전체 열기 ▼"}
                  </button>
                </div>

                <div className="mt-4 grid grid-cols-1 gap-4 lg:grid-cols-2">
                  <div>
                    <h3
                      className={`text-center ${FONT.cardTitle} ${WEIGHT.cardTitle} ${COLOR.cardTitle}`}
                    >
                      {compatibilityForm.left.name || "본인"}
                    </h3>
                    {renderSajuCard(
                      compatibilityResult.left,
                      "compat-left",
                      normalizeDateForCalc(compatibilityForm.left.birthDate),
                    )}
                  </div>

                  <div>
                    <h3
                      className={`text-center ${FONT.cardTitle} ${WEIGHT.cardTitle} ${COLOR.cardTitle}`}
                    >
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
              <h2
                className={`${FONT.sectionTitle} ${WEIGHT.sectionTitle} ${COLOR.sectionTitle}`}
              >
                궁합 해석
              </h2>

              <div
                className={`mt-3 whitespace-pre-wrap leading-7 ${FONT.analysisBody} ${WEIGHT.analysisBody} ${COLOR.analysisBody}`}
              >
                {result}
              </div>
            </section>
          )}

          {mode === "zodiac" && result && (
            <section className="mt-6 rounded-3xl border border-[#ead8c4] bg-[#fffaf3] p-5 shadow-inner">
              <h2
                className={`${FONT.sectionTitle} ${WEIGHT.sectionTitle} ${COLOR.sectionTitle}`}
              >
                점성술 해석
              </h2>

              <div
                className={`mt-3 whitespace-pre-wrap leading-7 ${FONT.analysisBody} ${WEIGHT.analysisBody} ${COLOR.analysisBody}`}
              >
                {result}
              </div>
            </section>
          )}
        </div>
      </div>
      {showDailyCalendar && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-6">
          <div className="max-h-[90vh] w-[1200px] overflow-y-auto rounded-3xl bg-white p-6 shadow-2xl">
            <div className="mb-6 flex items-center justify-between">
              <button
                type="button"
                onClick={() =>
                  setCalendarDate(
                    new Date(
                      calendarDate.getFullYear(),
                      calendarDate.getMonth() - 1,
                      1,
                    ),
                  )
                }
                className="rounded-xl bg-zinc-100 px-4 py-2 text-2xl font-bold"
              >
                이전달
              </button>

              <div className="text-4xl font-bold text-[#6b3f24]">
                {calendarDate.getFullYear()}년 {calendarDate.getMonth() + 1}월
                일진달력
              </div>

              <button
                type="button"
                onClick={() =>
                  setCalendarDate(
                    new Date(
                      calendarDate.getFullYear(),
                      calendarDate.getMonth() + 1,
                      1,
                    ),
                  )
                }
                className="rounded-xl bg-zinc-100 px-4 py-2 text-2xl font-bold"
              >
                다음달
              </button>
            </div>

            <div className="mb-2 grid grid-cols-7 gap-2 text-center text-3xl font-bold">
              {["일", "월", "화", "수", "목", "금", "토"].map((day) => (
                <div key={day}>{day}</div>
              ))}
            </div>

            <div className="grid grid-cols-7 gap-2">
              {buildDailyCalendar(calendarDate).map((item, index) =>
                item === null ? (
                  <div key={index} className="h-[120px]" />
                ) : (
                  <div
                    key={index}
                    className="rounded-2xl border bg-[#fffaf3] p-2 text-center"
                  >
                    <div className="text-2xl font-bold">{item.day}</div>
                    {item.solarTerms?.map((term: any) => (
                      <div
                        key={term.name}
                        className="mt-1 rounded-full bg-[#6b3f24] px-2 py-1 text-lg font-bold text-white"
                      >
                        {term.name} {formatTermTime(term.date)}
                      </div>
                    ))}

                    {renderHanjaButton({
                      type: "stem",
                      value: item.ganji.stem,
                      element: STEM_INFO[item.ganji.stem].element,
                      className: "mt-2 text-5xl font-bold",
                      children: STEM_HANJA[item.ganji.stem],
                    })}

                    {renderHanjaButton({
                      type: "branch",
                      value: item.ganji.branch,
                      element: STEM_INFO[BRANCH_MAIN_STEM[item.ganji.branch]].element,
                      className: "text-5xl font-bold",
                      children: BRANCH_HANJA[item.ganji.branch],
                    })}
                  </div>
                ),
              )}
            </div>

            <button
              type="button"
              onClick={() => setShowDailyCalendar(false)}
              className="mt-6 w-full rounded-2xl bg-[#6b3f24] py-4 text-3xl font-bold text-white"
            >
              닫기
            </button>
          </div>
        </div>
      )}

      {!memoOpen && (
        <button
          type="button"
          onClick={() => setMemoOpen(true)}
          className="fixed right-6 top-1/2 z-[60] -translate-y-1/2 rounded-l-2xl rounded-r-md bg-[#6b3f24] px-4 py-6 text-2xl font-bold text-white shadow-2xl transition hover:bg-[#4a2f1c]"
        >
          메모 열기
        </button>
      )}

      {memoOpen && (
        <div className="fixed right-24 top-24 z-[60] flex h-[70vh] w-[420px] flex-col rounded-3xl border border-[#ead8c4] bg-[#fffaf3] p-4 shadow-2xl">
          <div className="mb-3 flex items-center justify-between">
            <h2 className="text-3xl font-bold text-[#6b3f24]">상담 메모장</h2>

            <button
              type="button"
              onClick={() => setMemoOpen(false)}
              className="rounded-xl bg-white px-3 py-2 text-xl font-bold text-[#6b3f24] shadow-sm"
            >
              닫기
            </button>
          </div>

          <textarea
            value={memoText}
            onChange={(e) => setMemoText(e.target.value)}
            placeholder="상담 중 메모를 입력하세요. 새로고침 후에도 유지됩니다."
            className="min-h-0 flex-1 resize-none rounded-2xl border border-[#ead8c4] bg-white p-4 text-2xl leading-relaxed text-black outline-none"
          />

          <div className="mt-3 text-right text-xl font-bold text-[#6b3f24]">
            자동 저장됨
          </div>
        </div>
      )}
    </main>
  );
}
