"use client";

import KoreanLunarCalendar from "korean-lunar-calendar";
import { useEffect, useRef, useState } from "react";
import { flushSync } from "react-dom";
import { saveAs } from "file-saver";
import { toPng } from "html-to-image";
import JSZip from "jszip";
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

    // 스크롤 고정 정보바
    floatingInfo: "text-3xl",

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
  const overviewCaptureRef = useRef<HTMLDivElement>(null);
  const luckCaptureRef = useRef<HTMLDivElement>(null);

  const [mode, setMode] = useState<"saju" | "compatibility" | "tarot">("saju");

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
  const [recentBirthDateNotice, setRecentBirthDateNotice] = useState(false);
  const [recentBirthDateNoticeKey, setRecentBirthDateNoticeKey] = useState("");
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
  const [recentPeopleSearch, setRecentPeopleSearch] = useState("");
  const [storageLoaded, setStorageLoaded] = useState(false);
  const [memoOpen, setMemoOpen] = useState(false);
  const [memoText, setMemoText] = useState("");
  const [memoLoaded, setMemoLoaded] = useState(false);
  const [memoPosition, setMemoPosition] = useState({ x: 1460, y: 240 });
  const [draggingMemo, setDraggingMemo] = useState(false);
  const [drawingBoardOpen, setDrawingBoardOpen] = useState(false);
  const [timerOpen, setTimerOpen] = useState(true);
  const [timerInputMinutes, setTimerInputMinutes] = useState("10");
  const [timerRemainingSeconds, setTimerRemainingSeconds] = useState(10 * 60);
  const [timerRunning, setTimerRunning] = useState(false);
  const [timerPosition, setTimerPosition] = useState({ x: 1020, y: 24 });
  const [draggingTimer, setDraggingTimer] = useState(false);
  const [timerFinished, setTimerFinished] = useState(false);
  const [timerBlink, setTimerBlink] = useState(false);
  const [selectedHanja, setSelectedHanja] = useState<{
    type: "stem" | "branch";
    value: string;
  } | null>(null);
  const [selectedTarotCard, setSelectedTarotCard] = useState<string | null>(null);


  type TarotCard = {
    id: string;
    group: "메이저" | "지팡이" | "컵" | "검" | "동전";
    number: string;
    name: string;
    korean: string;
    upright: string;
    reversed: string;
  };

  const TAROT_CARDS: TarotCard[] = [
    { id: "major-0", group: "메이저", number: "0", name: "The Fool", korean: "바보", upright: "시작, 모든 일·사건·관계의 시작, 새로운 곳으로의 여행이나 이동, 순수한 마음과 생각, 긍정적이고 어린아이 같은 태도, 낙천적이고 자유로운 행동.", reversed: "끝이나 정지, 흐름의 정체와 지체, 엉뚱하거나 돌발적인 행동, 생각 없이 행동하는 어리석음, 행동이 따르지 않는 공상, 낭비와 탕진." },
    { id: "major-1", group: "메이저", number: "I", name: "The Magician", korean: "마법사", upright: "창조, 수완, 능력, 재능의 발휘, 자신감, 기회를 현실로 만드는 힘.", reversed: "겁이 많음, 기만, 능력의 오용, 과신, 말뿐인 계획, 속임수." },
    { id: "major-2", group: "메이저", number: "II", name: "The High Priestess", korean: "여사제", upright: "지식, 총명, 순결, 직관, 신중함, 내면의 지혜.", reversed: "잔혹, 무례함, 냉담함, 비밀의 오용, 직관을 무시함." },
    { id: "major-3", group: "메이저", number: "III", name: "The Empress", korean: "여황제", upright: "풍성함, 모성, 풍요, 돌봄, 성장, 생산성과 결실.", reversed: "과잉, 허영, 지나친 의존이나 집착, 낭비, 성장이 막힘." },
    { id: "major-4", group: "메이저", number: "IV", name: "The Emperor", korean: "황제", upright: "책임, 부성, 카리스마, 질서, 통제력, 안정된 기반.", reversed: "오만, 강한 자만심, 독선, 지나친 통제, 권위의 남용." },
    { id: "major-5", group: "메이저", number: "V", name: "The Hierophant", korean: "교황", upright: "가르침, 지성, 관대함, 전통, 조언, 제도와 규범.", reversed: "꽉 막힘, 나태함, 고정관념, 형식주의, 잘못된 권위." },
    { id: "major-6", group: "메이저", number: "VI", name: "The Lovers", korean: "연인", upright: "연애, 성적 끌림, 쾌락, 선택, 조화로운 관계, 결합.", reversed: "사랑하기 힘듦, 배신감, 관계의 불균형, 잘못된 선택, 갈등." },
    { id: "major-7", group: "메이저", number: "VII", name: "The Chariot", korean: "전차", upright: "전진, 승리, 조화, 강한 추진력, 목표를 향한 돌파.", reversed: "폭주, 좌절, 브레이크 없는 상태, 방향 상실, 무리한 추진." },
    { id: "major-8", group: "메이저", number: "VIII", name: "Strength", korean: "힘", upright: "힘, 용기, 능력, 인내, 본능을 다스리는 내적 강인함.", reversed: "본성에 휘둘림, 자만, 자신감 저하, 감정 조절의 어려움." },
    { id: "major-9", group: "메이저", number: "IX", name: "The Hermit", korean: "은둔자", upright: "탐색, 사려 깊음, 차분함, 성찰, 혼자만의 시간, 지혜.", reversed: "음습함, 탐욕, 고립, 지나친 폐쇄성, 외로움." },
    { id: "major-10", group: "메이저", number: "X", name: "Wheel of Fortune", korean: "운명의 수레바퀴", upright: "기회, 일시적인 행운, 운명적인 변화, 흐름의 전환.", reversed: "오산, 불행의 연속, 시기 불일치, 반복되는 악순환." },
    { id: "major-11", group: "메이저", number: "XI", name: "Justice", korean: "정의", upright: "균형, 정당함, 깔끔함, 공정한 판단, 원인과 결과.", reversed: "편견, 부정부패, 불공정, 책임 회피, 잘못된 판단." },
    { id: "major-12", group: "메이저", number: "XII", name: "The Hanged Man", korean: "매달린 사람", upright: "자기희생, 인내, 멈춤을 통한 통찰, 관점의 전환.", reversed: "엉뚱하고 무의미한 희생, 맹목적 태도, 정체, 보상 없는 기다림." },
    { id: "major-13", group: "메이저", number: "XIII", name: "Death", korean: "죽음", upright: "격변, 이별, 끝을 통한 새로운 희망, 큰 변화와 전환.", reversed: "변화의 유보, 끝내지 못함, 과거에 매임, 정체가 길어짐." },
    { id: "major-14", group: "메이저", number: "XIV", name: "Temperance", korean: "절제", upright: "조화, 중용, 견실함, 균형, 적절한 조절과 타협.", reversed: "낭비, 불안정, 과도함, 조절 실패, 균형의 붕괴." },
    { id: "major-15", group: "메이저", number: "XV", name: "The Devil", korean: "악마", upright: "사심, 쾌락, 속박, 타락, 집착과 중독, 욕망에 묶임.", reversed: "악순환으로부터의 각성, 속박에서 벗어남, 집착을 끊으려는 움직임." },
    { id: "major-16", group: "메이저", number: "XVI", name: "The Tower", korean: "탑", upright: "파괴, 파멸, 급격한 변화, 충격적인 깨달음, 기존 구조의 붕괴.", reversed: "파괴되지 않음, 변화의 지연, 충격을 피하려 함, 붕괴를 간신히 막음." },
    { id: "major-17", group: "메이저", number: "XVII", name: "The Star", korean: "별", upright: "희망, 연예인 같은 매력, 동경, 회복, 미래에 대한 기대.", reversed: "환멸, 비애, 추락, 희망 상실, 자신감 저하." },
    { id: "major-18", group: "메이저", number: "XVIII", name: "The Moon", korean: "달", upright: "불안, 애매함, 보이지 않는 것, 착각, 감정의 흔들림.", reversed: "명료함, 혼돈의 끝, 진실이 드러남, 불안의 완화." },
    { id: "major-19", group: "메이저", number: "XIX", name: "The Sun", korean: "태양", upright: "밝은 미래, 만족, 기대, 성공, 활력과 긍정적인 결과.", reversed: "우울함, 실패, 기대 이하, 자신감 저하, 밝음이 가려짐." },
    { id: "major-20", group: "메이저", number: "XX", name: "Judgement", korean: "심판", upright: "부활, 개선, 소식, 재기, 결단, 과거를 정리하고 다시 시작함.", reversed: "재기불능, 후회막급, 잘못된 판단, 기회를 놓침, 과거에 붙잡힘." },
    { id: "major-21", group: "메이저", number: "XXI", name: "The World", korean: "세계", upright: "완성, 완전, 윤회, 성취, 한 사이클의 성공적인 마무리.", reversed: "미완성, 어중간함, 마무리 부족, 목표 직전의 지연." },

    { id: "wands-ace", group: "지팡이", number: "Ace", name: "Ace of Wands", korean: "지팡이 에이스", upright: "열정, 창조력, 출발, 새로운 기회, 행동의 시작.", reversed: "의욕 저하, 시작 지연, 방향 없는 열정, 기회를 놓침." },
    { id: "wands-2", group: "지팡이", number: "2", name: "Two of Wands", korean: "지팡이 2", upright: "야망과 신념, 이동, 계획, 더 넓은 세계를 바라봄.", reversed: "계획 부족, 두려움, 우유부단, 좁은 시야." },
    { id: "wands-3", group: "지팡이", number: "3", name: "Three of Wands", korean: "지팡이 3", upright: "교역, 사업상의 협력, 확장, 전망, 기다리던 성과.", reversed: "협력의 지연, 계획 차질, 기대 이하의 성과." },
    { id: "wands-4", group: "지팡이", number: "4", name: "Four of Wands", korean: "지팡이 4", upright: "안정과 번영, 휴식, 평화, 축하, 기반의 완성.", reversed: "불안정한 기반, 가족·조직 내 갈등, 축하의 지연." },
    { id: "wands-5", group: "지팡이", number: "5", name: "Five of Wands", korean: "지팡이 5", upright: "치열한 경쟁, 싸움, 의견 충돌, 서로 겨루는 상황.", reversed: "갈등 회피, 경쟁의 완화, 내부 갈등, 소모적인 다툼." },
    { id: "wands-6", group: "지팡이", number: "6", name: "Six of Wands", korean: "지팡이 6", upright: "승리자, 정복과 성공, 인정, 자신감, 좋은 소식.", reversed: "인정받지 못함, 자신감 저하, 승리의 지연, 평판 문제." },
    { id: "wands-7", group: "지팡이", number: "7", name: "Seven of Wands", korean: "지팡이 7", upright: "용기, 자기방어, 자신감 있게 저항함, 장애물을 극복함, 타협하지 않는 정신력.", reversed: "근심과 걱정, 난처한 상황, 끝없는 장애, 이길 수 없음, 용기를 내기 어려움." },
    { id: "wands-8", group: "지팡이", number: "8", name: "Eight of Wands", korean: "지팡이 8", upright: "활동성, 재빠름, 빠른 전개, 이동, 소식과 추진력.", reversed: "지연, 엇갈린 소통, 조급함, 방향 없는 움직임." },
    { id: "wands-9", group: "지팡이", number: "9", name: "Nine of Wands", korean: "지팡이 9", upright: "싸움에 지친 상태, 마지막 방어, 끈기, 경계심.", reversed: "탈진, 방어 포기, 지나친 의심, 더 버티기 어려움." },
    { id: "wands-10", group: "지팡이", number: "10", name: "Ten of Wands", korean: "지팡이 10", upright: "억압, 부담, 책임 과중, 일을 혼자 짊어짐.", reversed: "부담을 내려놓음, 책임 회피, 과로의 한계, 짐을 나눔." },
    { id: "wands-page", group: "지팡이", number: "Page", name: "Page of Wands", korean: "지팡이 시종", upright: "젊은 남성, 사회초년병, 호기심, 새로운 소식, 모험의 시작.", reversed: "미숙함, 충동적 행동, 계획 없는 시작, 좋지 않은 소식." },
    { id: "wands-knight", group: "지팡이", number: "Knight", name: "Knight of Wands", korean: "지팡이 기사", upright: "출발, 정열적임, 행동력, 모험, 빠른 이동.", reversed: "성급함, 무모함, 쉽게 식는 열정, 충돌과 지연." },
    { id: "wands-queen", group: "지팡이", number: "Queen", name: "Queen of Wands", korean: "지팡이 여왕", upright: "커리어 우먼 같은 모습, 자신감, 독립성, 따뜻한 카리스마.", reversed: "질투, 예민함, 독선, 자신감의 흔들림." },
    { id: "wands-king", group: "지팡이", number: "King", name: "King of Wands", korean: "지팡이 왕", upright: "지도자 같은 사람, 비전, 추진력, 창조적 리더십.", reversed: "독단, 성급한 리더십, 고집, 과도한 자신감." },

    { id: "cups-ace", group: "컵", number: "Ace", name: "Ace of Cups", korean: "컵 에이스", upright: "사랑의 기쁨, 시작, 감정의 충만, 새로운 관계.", reversed: "감정 억압, 사랑의 지연, 공허함, 관계 시작의 어려움." },
    { id: "cups-2", group: "컵", number: "2", name: "Two of Cups", korean: "컵 2", upright: "사랑, 우정, 결혼, 상호 교감, 좋은 파트너십.", reversed: "관계 불균형, 오해, 이별 가능성, 감정의 엇갈림." },
    { id: "cups-3", group: "컵", number: "3", name: "Three of Cups", korean: "컵 3", upright: "풍족함, 행복, 성취, 우정, 축하와 모임.", reversed: "과도한 즐거움, 삼각관계, 인간관계의 피로, 소문." },
    { id: "cups-4", group: "컵", number: "4", name: "Four of Cups", korean: "컵 4", upright: "권태, 낙담, 무관심, 제안을 외면함, 감정적 정체.", reversed: "새로운 관심, 권태에서 벗어남, 기회를 다시 봄." },
    { id: "cups-5", group: "컵", number: "5", name: "Five of Cups", korean: "컵 5", upright: "손실, 상심, 후회, 잃은 것에 집중함.", reversed: "회복, 과거를 놓음, 관계 회복, 희망을 다시 봄." },
    { id: "cups-6", group: "컵", number: "6", name: "Six of Cups", korean: "컵 6", upright: "과거를 돌이켜봄, 회상, 추억, 순수한 호의.", reversed: "과거에 매임, 미련, 현실 회피, 오래된 문제의 반복." },
    { id: "cups-7", group: "컵", number: "7", name: "Seven of Cups", korean: "컵 7", upright: "환상, 허황됨, 여러 선택지, 꿈과 욕망.", reversed: "현실적인 선택, 환상에서 깨어남, 우선순위 정리." },
    { id: "cups-8", group: "컵", number: "8", name: "Eight of Cups", korean: "컵 8", upright: "성공의 방치, 포기, 떠남, 더 나은 것을 찾아감.", reversed: "떠나지 못함, 미련, 반복되는 관계, 결단 지연." },
    { id: "cups-9", group: "컵", number: "9", name: "Nine of Cups", korean: "컵 9", upright: "물질적 안녕, 만족, 소원 성취, 즐거움.", reversed: "과욕, 만족하지 못함, 겉보기와 다른 공허함." },
    { id: "cups-10", group: "컵", number: "10", name: "Ten of Cups", korean: "컵 10", upright: "만족, 가족적인 행복, 정서적 완성, 화목.", reversed: "가족 갈등, 관계의 불안정, 기대했던 행복의 흔들림." },
    { id: "cups-page", group: "컵", number: "Page", name: "Page of Cups", korean: "컵 시종", upright: "공부에 힘쓰는 젊은이, 호기심, 감성적 소식, 새로운 감정.", reversed: "감정 미숙, 현실성 부족, 예민함, 소식의 지연." },
    { id: "cups-knight", group: "컵", number: "Knight", name: "Knight of Cups", korean: "컵 기사", upright: "도착, 발전, 청혼, 제안, 로맨틱한 접근.", reversed: "비현실적 약속, 감정 기복, 거짓된 호의, 관계 지연." },
    { id: "cups-queen", group: "컵", number: "Queen", name: "Queen of Cups", korean: "컵 여왕", upright: "선량하고 공정한 여성, 강한 모성, 공감과 배려.", reversed: "감정 과잉, 의존, 예민함, 타인에게 휘둘림." },
    { id: "cups-king", group: "컵", number: "King", name: "King of Cups", korean: "컵 왕", upright: "공정한 남성, 창조적 지성, 감정적 성숙과 안정.", reversed: "감정 조종, 냉정함, 속마음을 숨김, 감정 통제 실패." },

    { id: "swords-ace", group: "검", number: "Ace", name: "Ace of Swords", korean: "검 에이스", upright: "힘의 승리, 기회와 출발, 명확한 판단, 결단.", reversed: "혼란, 잘못된 판단, 말로 인한 상처, 기회의 오용." },
    { id: "swords-2", group: "검", number: "2", name: "Two of Swords", korean: "검 2", upright: "균형, 우유부단, 결정을 미룸, 갈등을 잠시 막음.", reversed: "결정 압박, 혼란 심화, 감춰진 사실이 드러남." },
    { id: "swords-3", group: "검", number: "3", name: "Three of Swords", korean: "검 3", upright: "후퇴, 단절, 슬픔, 상처, 이별.", reversed: "상처 회복, 슬픔을 놓음, 관계 회복의 가능성." },
    { id: "swords-4", group: "검", number: "4", name: "Four of Swords", korean: "검 4", upright: "은둔, 회복, 자기치유, 휴식, 재정비.", reversed: "휴식 부족, 불안, 다시 움직여야 함, 회복 지연." },
    { id: "swords-5", group: "검", number: "5", name: "Five of Swords", korean: "검 5", upright: "타락, 패배, 손실, 상처뿐인 승리, 갈등.", reversed: "화해 시도, 갈등을 끝냄, 패배를 인정함." },
    { id: "swords-6", group: "검", number: "6", name: "Six of Swords", korean: "검 6", upright: "작업을 끝마침, 중개자, 여행, 문제에서 벗어나는 이동.", reversed: "벗어나지 못함, 이동 지연, 과거 문제의 반복." },
    { id: "swords-7", group: "검", number: "7", name: "Seven of Swords", korean: "검 7", upright: "서두름을 경고, 계획, 전략, 남몰래 움직임.", reversed: "계획 노출, 속임수가 드러남, 전략 수정, 죄책감." },
    { id: "swords-8", group: "검", number: "8", name: "Eight of Swords", korean: "검 8", upright: "구속된 힘, 비난, 고립, 스스로 만든 제약.", reversed: "제약에서 벗어남, 해결책 발견, 통제력을 되찾음." },
    { id: "swords-9", group: "검", number: "9", name: "Nine of Swords", korean: "검 9", upright: "실망, 환멸, 불안, 걱정, 잠 못 이루는 고민.", reversed: "불안 완화, 최악에서 벗어남, 두려움과 대면함." },
    { id: "swords-10", group: "검", number: "10", name: "Ten of Swords", korean: "검 10", upright: "황폐, 고통, 종결, 피할 수 없는 끝.", reversed: "회복의 시작, 최악이 지나감, 다시 일어설 준비." },
    { id: "swords-page", group: "검", number: "Page", name: "Page of Swords", korean: "검 시종", upright: "경솔, 감시, 경계, 관찰, 정보를 모음.", reversed: "헛소문, 지나친 경계, 말실수, 정보의 왜곡." },
    { id: "swords-knight", group: "검", number: "Knight", name: "Knight of Swords", korean: "검 기사", upright: "용감함, 성급함, 직진, 빠른 결정과 행동.", reversed: "무모함, 공격성, 성급한 판단, 충돌." },
    { id: "swords-queen", group: "검", number: "Queen", name: "Queen of Swords", korean: "검 여왕", upright: "슬픔이 많은 여성, 미망인, 냉철함, 독립성과 판단력.", reversed: "냉혹함, 비판적 태도, 원한, 지나친 방어." },
    { id: "swords-king", group: "검", number: "King", name: "King of Swords", korean: "검 왕", upright: "가부장적, 권위, 명령, 이성적 판단과 통솔.", reversed: "권위 남용, 독단, 냉정한 통제, 판단의 왜곡." },

    { id: "pentacles-ace", group: "동전", number: "Ace", name: "Ace of Pentacles", korean: "동전 에이스", upright: "물질적 안정, 돈을 투자함, 재정적 기회, 현실적인 시작.", reversed: "금전 기회 손실, 투자 실패, 불안정, 준비 부족." },
    { id: "pentacles-2", group: "동전", number: "2", name: "Two of Pentacles", korean: "동전 2", upright: "명랑함, 현실과 갈등, 여러 일을 조율함, 균형 유지.", reversed: "균형 붕괴, 일정·재정 관리 실패, 우선순위 혼란." },
    { id: "pentacles-3", group: "동전", number: "3", name: "Three of Pentacles", korean: "동전 3", upright: "기예, 숙련공, 협업, 전문성, 실력을 인정받음.", reversed: "협업 실패, 실력 부족, 낮은 완성도, 인정받지 못함." },
    { id: "pentacles-4", group: "동전", number: "4", name: "Four of Pentacles", korean: "동전 4", upright: "소유물에 대한 집착, 선물, 재산을 지키려 함.", reversed: "집착 완화, 재정 손실, 지나친 소비, 소유를 놓음." },
    { id: "pentacles-5", group: "동전", number: "5", name: "Five of Pentacles", korean: "동전 5", upright: "물질적인 고통, 빈곤, 소외, 어려운 재정 상황.", reversed: "재정 회복, 도움을 받음, 어려움에서 벗어나기 시작함." },
    { id: "pentacles-6", group: "동전", number: "6", name: "Six of Pentacles", korean: "동전 6", upright: "성공, 자선, 주고받음, 도움, 금전적 균형.", reversed: "불공평한 거래, 조건부 도움, 빚, 금전 불균형." },
    { id: "pentacles-7", group: "동전", number: "7", name: "Seven of Pentacles", korean: "동전 7", upright: "금전적 망설임, 교역, 기다림, 투자 결과를 지켜봄.", reversed: "성과 부족, 조급함, 투자 재검토, 노력 대비 보상 부족." },
    { id: "pentacles-8", group: "동전", number: "8", name: "Eight of Pentacles", korean: "동전 8", upright: "장인 기질, 준비함, 숙련, 반복을 통한 실력 향상.", reversed: "대충함, 기술 부족, 반복되는 실수, 일에 대한 권태." },
    { id: "pentacles-9", group: "동전", number: "9", name: "Nine of Pentacles", korean: "동전 9", upright: "물질적인 풍요, 달성, 독립, 노력의 결실.", reversed: "과소비, 겉치레, 재정 의존, 성취의 불안정." },
    { id: "pentacles-10", group: "동전", number: "10", name: "Ten of Pentacles", korean: "동전 10", upright: "이익, 재산, 가족, 유산, 장기적인 안정.", reversed: "가족 재산 문제, 금전 갈등, 기반 흔들림, 장기 계획 차질." },
    { id: "pentacles-page", group: "동전", number: "Page", name: "Page of Pentacles", korean: "동전 시종", upright: "성실근면, 학생 같은 태도, 공부, 실무적 기회의 시작.", reversed: "게으름, 계획만 있음, 학업·업무 집중 저하, 기회 활용 실패." },
    { id: "pentacles-knight", group: "동전", number: "Knight", name: "Knight of Pentacles", korean: "동전 기사", upright: "유용, 재산, 신뢰감, 꾸준함, 성실하게 진행함.", reversed: "정체, 지나친 보수성, 게으름, 변화 없는 반복." },
    { id: "pentacles-queen", group: "동전", number: "Queen", name: "Queen of Pentacles", korean: "동전 여왕", upright: "부, 현모양처, 안정, 현실적 돌봄, 풍요로운 생활.", reversed: "금전 불안, 자기 돌봄 부족, 소유 집착, 가정과 일의 불균형." },
    { id: "pentacles-king", group: "동전", number: "King", name: "King of Pentacles", korean: "동전 왕", upright: "사업에 성공, 자선, 재정적 안정, 현실적인 리더십.", reversed: "물질 집착, 탐욕, 돈을 이용한 통제, 사업적 완고함." },
  ];

  const TAROT_GROUPS = ["메이저", "지팡이", "컵", "검", "동전"] as const;

  const timerDragOffsetRef = useRef({ x: 0, y: 0 });
  const memoDragOffsetRef = useRef({ x: 0, y: 0 });

  const getDefaultMemoPosition = () => {
    if (typeof window === "undefined") return { x: 1460, y: 240 };

    return {
      x: Math.max(16, window.innerWidth - 444),
      y: 240,
    };
  };

  const restoreMemoPosition = () => {
    setMemoOpen(true);
    setDraggingMemo(false);
    setMemoPosition(getDefaultMemoPosition());
  };

  const getDefaultTimerPosition = () => {
    if (typeof window === "undefined") return { x: 1020, y: 24 };

    return {
      x: Math.max(16, window.innerWidth - 350),
      y: 24,
    };
  };

  const restoreTimerPosition = () => {
    setTimerOpen(true);
    setDraggingTimer(false);
    setTimerPosition(getDefaultTimerPosition());
  };


  type PenPoint = { x: number; y: number };
  type PenStroke = {
    mode: "draw" | "erase";
    points: PenPoint[];
  };

  const penCanvasRef = useRef<HTMLCanvasElement>(null);
  const penCanvasSizeRef = useRef({ width: 0, height: 0, dpr: 1 });
  const penDrawingRef = useRef(false);
  const penErasingRef = useRef(false);
  const penEraseHasContactRef = useRef(false);
  const penCurrentPointerIdRef = useRef<number | null>(null);
  const penCurrentStrokeRef = useRef<PenStroke | null>(null);
  const penStrokesRef = useRef<PenStroke[]>([]);
  const pagePenStrokesRef = useRef<PenStroke[]>([]);
  const drawingBoardStrokesRef = useRef<PenStroke[]>([]);
  const penModeRef = useRef(true);
  const [penMode, setPenMode] = useState(true);
  const [penCanvasMounted, setPenCanvasMounted] = useState(false);

  // 와콤 펜 접촉 직후 브라우저가 mouse/click 이벤트를 추가로 발생시키는 것을 막기 위한 값
  const penClickSuppressUntilRef = useRef(0);
  const penLastScreenPointRef = useRef<{ x: number; y: number } | null>(null);

  const PEN_COLOR = "#dc2626";
  const PEN_ALPHA = 0.55;
  const PEN_WIDTH = 8;
  const ERASER_WIDTH = PEN_WIDTH * 4;

  const isEraserButtonPressed = (event: PointerEvent) => {
    return (event.buttons & 2) === 2 || (event.buttons & 32) === 32;
  };

  const isPenTipContact = (event: PointerEvent) => {
    // Hover 상태에서 사이드 버튼만 누르면 보통 buttons가 2 또는 32로만 들어온다.
    // 실제 펜촉 접촉은 primary bit(1)가 같이 들어오거나 pressure가 0보다 커지는 경우만 인정한다.
    return (event.buttons & 1) === 1 || event.pressure > 0;
  };

  const getPenCanvasContext = () => {
    const canvas = penCanvasRef.current;
    if (!canvas) return null;

    const context = canvas.getContext("2d");
    if (!context) return null;

    context.lineCap = "round";
    context.lineJoin = "round";
    context.strokeStyle = PEN_COLOR;
    context.lineWidth = PEN_WIDTH;
    context.globalAlpha = 1;
    context.globalCompositeOperation = "source-over";

    return context;
  };

  const drawPenStroke = (context: CanvasRenderingContext2D, stroke: PenStroke) => {
    if (stroke.points.length < 2) return;

    context.save();

    if (stroke.mode === "erase") {
      context.globalCompositeOperation = "destination-out";
      context.globalAlpha = 1;
      context.lineWidth = ERASER_WIDTH;
    } else {
      context.globalCompositeOperation = "source-over";
      context.globalAlpha = PEN_ALPHA;
      context.strokeStyle = PEN_COLOR;
      context.lineWidth = PEN_WIDTH;
    }

    context.lineCap = "round";
    context.lineJoin = "round";
    context.beginPath();
    context.moveTo(stroke.points[0].x, stroke.points[0].y);

    for (let index = 1; index < stroke.points.length; index += 1) {
      context.lineTo(stroke.points[index].x, stroke.points[index].y);
    }

    context.stroke();
    context.restore();
  };

  const redrawPenCanvas = (includeCurrentStroke = false) => {
    const canvas = penCanvasRef.current;
    const context = getPenCanvasContext();
    if (!canvas || !context) return;

    context.clearRect(0, 0, canvas.width, canvas.height);
    penStrokesRef.current.forEach((stroke) => drawPenStroke(context, stroke));

    if (includeCurrentStroke && penCurrentStrokeRef.current) {
      drawPenStroke(context, penCurrentStrokeRef.current);
    }
  };

  const resizePenCanvas = () => {
    const canvas = penCanvasRef.current;
    if (!canvas || typeof window === "undefined") return;

    // DPR을 1로 제한해야 대형 페이지에서 캔버스 메모리 폭증을 막을 수 있음
    const dpr = 1;
    const width = Math.max(
      document.documentElement.scrollWidth,
      document.body.scrollWidth,
      window.innerWidth,
    );
    const height = Math.max(
      document.documentElement.scrollHeight,
      document.body.scrollHeight,
      window.innerHeight,
    );

    const previousSize = penCanvasSizeRef.current;
    const sameSize =
      previousSize.width === width &&
      previousSize.height === height &&
      previousSize.dpr === dpr;

    if (!sameSize) {
      canvas.style.width = `${width}px`;
      canvas.style.height = `${height}px`;
      canvas.width = Math.ceil(width * dpr);
      canvas.height = Math.ceil(height * dpr);
      penCanvasSizeRef.current = { width, height, dpr };
    }

    const context = canvas.getContext("2d");
    if (!context) return;

    context.setTransform(dpr, 0, 0, dpr, 0, 0);
    context.lineCap = "round";
    context.lineJoin = "round";
    context.strokeStyle = PEN_COLOR;
    context.lineWidth = PEN_WIDTH;
    context.globalAlpha = 1;
    context.globalCompositeOperation = "source-over";

    context.clearRect(0, 0, width, height);
    penStrokesRef.current.forEach((stroke) => drawPenStroke(context, stroke));
    if (penDrawingRef.current && penCurrentStrokeRef.current) {
      drawPenStroke(context, penCurrentStrokeRef.current);
    }
  };

  const getPenPoint = (event: PointerEvent) => {
    return {
      x: event.clientX + window.scrollX,
      y: event.clientY + window.scrollY,
    };
  };

  const stopPenEvent = (event: PointerEvent) => {
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation();
  };

  const markPenClickSuppress = (event: PointerEvent) => {
    penClickSuppressUntilRef.current = performance.now() + 900;
    penLastScreenPointRef.current = {
      x: event.clientX,
      y: event.clientY,
    };
  };

  const shouldSuppressMouseEventAfterPen = (event: MouseEvent) => {
    if (performance.now() > penClickSuppressUntilRef.current) return false;

    const lastPoint = penLastScreenPointRef.current;
    if (!lastPoint) return true;

    const dx = event.clientX - lastPoint.x;
    const dy = event.clientY - lastPoint.y;

    return dx * dx + dy * dy <= 80 * 80;
  };

  const stopMouseEventAfterPen = (event: MouseEvent) => {
    if (!shouldSuppressMouseEventAfterPen(event)) return;

    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation();
  };

  const distancePointToSegment = (
    px: number,
    py: number,
    x1: number,
    y1: number,
    x2: number,
    y2: number,
  ) => {
    const dx = x2 - x1;
    const dy = y2 - y1;

    if (dx === 0 && dy === 0) {
      return Math.hypot(px - x1, py - y1);
    }

    const t = ((px - x1) * dx + (py - y1) * dy) / (dx * dx + dy * dy);
    const clampedT = Math.max(0, Math.min(1, t));
    const closestX = x1 + clampedT * dx;
    const closestY = y1 + clampedT * dy;

    return Math.hypot(px - closestX, py - closestY);
  };

  const isStrokeHitByPoint = (stroke: PenStroke, point: PenPoint) => {
    const hitRadius = ERASER_WIDTH / 2;

    if (stroke.points.length === 1) {
      const onlyPoint = stroke.points[0];
      return Math.hypot(point.x - onlyPoint.x, point.y - onlyPoint.y) <= hitRadius;
    }

    for (let index = 1; index < stroke.points.length; index += 1) {
      const previousPoint = stroke.points[index - 1];
      const currentPoint = stroke.points[index];
      const distance = distancePointToSegment(
        point.x,
        point.y,
        previousPoint.x,
        previousPoint.y,
        currentPoint.x,
        currentPoint.y,
      );

      if (distance <= hitRadius) return true;
    }

    return false;
  };

  const eraseStrokeAtPoint = (point: PenPoint) => {
    for (let strokeIndex = penStrokesRef.current.length - 1; strokeIndex >= 0; strokeIndex -= 1) {
      const stroke = penStrokesRef.current[strokeIndex];

      if (!isStrokeHitByPoint(stroke, point)) continue;

      penStrokesRef.current = penStrokesRef.current.filter((_, index) => index !== strokeIndex);
      requestAnimationFrame(() => redrawPenCanvas(false));
      return true;
    }

    return false;
  };

  const clearPenCanvas = () => {
    penStrokesRef.current = [];
    penCurrentStrokeRef.current = null;
    penDrawingRef.current = false;
    penErasingRef.current = false;
    penEraseHasContactRef.current = false;
    penCurrentPointerIdRef.current = null;

    requestAnimationFrame(() => redrawPenCanvas(false));
  };

  const undoPenStroke = () => {
    penStrokesRef.current = penStrokesRef.current.slice(0, -1);
    penCurrentStrokeRef.current = null;
    penDrawingRef.current = false;
    penErasingRef.current = false;
    penEraseHasContactRef.current = false;
    penCurrentPointerIdRef.current = null;

    requestAnimationFrame(() => redrawPenCanvas(false));
  };

  const openDrawingBoard = () => {
    pagePenStrokesRef.current = penStrokesRef.current;
    penStrokesRef.current = drawingBoardStrokesRef.current;
    penCurrentStrokeRef.current = null;
    penDrawingRef.current = false;
    penErasingRef.current = false;
    penCurrentPointerIdRef.current = null;
    setPenMode(true);
    setPenCanvasMounted(true);
    setDrawingBoardOpen(true);

    requestAnimationFrame(() => {
      resizePenCanvas();
      redrawPenCanvas(false);
    });
  };

  const closeDrawingBoard = () => {
    drawingBoardStrokesRef.current = penStrokesRef.current;
    penStrokesRef.current = pagePenStrokesRef.current;
    penCurrentStrokeRef.current = null;
    penDrawingRef.current = false;
    penErasingRef.current = false;
    penCurrentPointerIdRef.current = null;
    setDrawingBoardOpen(false);

    requestAnimationFrame(() => {
      resizePenCanvas();
      redrawPenCanvas(false);
    });
  };

  useEffect(() => {
    penModeRef.current = penMode;
  }, [penMode]);

  useEffect(() => {
    if (typeof window === "undefined") return;

    const ensureCanvasReady = () => {
      setPenCanvasMounted(true);
      requestAnimationFrame(() => {
        resizePenCanvas();
      });
    };

    const capturePenPointer = (pointerId: number) => {
      const canvas = penCanvasRef.current;
      if (!canvas) return;

      try {
        canvas.setPointerCapture(pointerId);
      } catch {
        // 일부 브라우저/와콤 조합에서는 캡처가 실패할 수 있음
      }
    };

    const releasePenPointer = (pointerId: number | null) => {
      const canvas = penCanvasRef.current;
      if (!canvas || pointerId === null) return;

      try {
        if (canvas.hasPointerCapture(pointerId)) {
          canvas.releasePointerCapture(pointerId);
        }
      } catch {
        // 무시
      }
    };

    const handlePointerDown = (event: PointerEvent) => {
      // 시작할 때만 펜인지 확인한다.
      // 이동 중 pointerType/pointerId가 흔들리는 와콤 드라이버가 있어서 move에서는 검사하지 않는다.
      if (!penModeRef.current || event.pointerType !== "pen") return;

      markPenClickSuppress(event);
      stopPenEvent(event);
      ensureCanvasReady();

      const nextPoint = getPenPoint(event);
      const isErasing = isEraserButtonPressed(event);

      penDrawingRef.current = true;
      penErasingRef.current = isErasing;
      penEraseHasContactRef.current = !isErasing;
      penCurrentPointerIdRef.current = event.pointerId;

      if (isErasing) {
        // 와콤 사이드 버튼은 hover 상태에서도 pointerdown/right-click 이벤트가 발생할 수 있다.
        // 그래서 pointerdown 순간에는 삭제하지 않고, 실제 드래그(pointermove)가 들어올 때만 획을 삭제한다.
        penCurrentStrokeRef.current = null;
      } else {
        penCurrentStrokeRef.current = {
          mode: "draw",
          points: [nextPoint],
        };
      }

      requestAnimationFrame(() => capturePenPointer(event.pointerId));
    };

    const handlePointerMove = (event: PointerEvent) => {
      if (!penModeRef.current || !penDrawingRef.current) return;

      markPenClickSuppress(event);
      stopPenEvent(event);

      const nextPoint = getPenPoint(event);

      if (penErasingRef.current) {
        if (!isPenTipContact(event)) return;
        penEraseHasContactRef.current = true;
        eraseStrokeAtPoint(nextPoint);
        return;
      }

      const stroke = penCurrentStrokeRef.current;
      if (!stroke) return;

      const prevPoint = stroke.points[stroke.points.length - 1];

      if (!prevPoint) {
        stroke.points.push(nextPoint);
        requestAnimationFrame(() => redrawPenCanvas(true));
        return;
      }

      // 좌표가 완전히 같은 이벤트는 건너뜀
      if (prevPoint.x === nextPoint.x && prevPoint.y === nextPoint.y) return;

      stroke.points.push(nextPoint);

      // 반투명 선이 이동 중에 겹쳐서 진해지는 문제를 막기 위해
      // 매 move마다 누적해서 그리지 않고, 저장된 stroke를 한 번씩만 다시 그림
      requestAnimationFrame(() => redrawPenCanvas(true));
    };

    const finishStroke = (event?: PointerEvent) => {
      if (!penDrawingRef.current) return;

      if (event) {
        markPenClickSuppress(event);
        stopPenEvent(event);
      }

      const stroke = penCurrentStrokeRef.current;
      if (!penErasingRef.current && stroke && stroke.points.length > 1) {
        penStrokesRef.current = [...penStrokesRef.current, stroke];
      }

      releasePenPointer(penCurrentPointerIdRef.current);

      penCurrentStrokeRef.current = null;
      penDrawingRef.current = false;
      penErasingRef.current = false;
      penEraseHasContactRef.current = false;
      penCurrentPointerIdRef.current = null;
      requestAnimationFrame(() => redrawPenCanvas(false));
    };

    const handlePointerCancel = (event: PointerEvent) => {
      // 와콤에서 pointercancel이 잦게 발생하면 선이 중간에 끊긴다.
      // 그래서 cancel은 선 종료로 보지 않고, 브라우저 기본 동작만 막는다.
      if (!penDrawingRef.current) return;
      stopPenEvent(event);
    };

    const handleBlur = () => {
      finishStroke();
    };

    window.addEventListener("pointerdown", handlePointerDown, true);
    window.addEventListener("pointermove", handlePointerMove, true);
    window.addEventListener("pointerup", finishStroke, true);
    window.addEventListener("pointercancel", handlePointerCancel, true);
    window.addEventListener("mousedown", stopMouseEventAfterPen, true);
    window.addEventListener("mouseup", stopMouseEventAfterPen, true);
    window.addEventListener("click", stopMouseEventAfterPen, true);
    window.addEventListener("dblclick", stopMouseEventAfterPen, true);
    window.addEventListener("contextmenu", stopMouseEventAfterPen, true);
    window.addEventListener("blur", handleBlur);

    return () => {
      window.removeEventListener("pointerdown", handlePointerDown, true);
      window.removeEventListener("pointermove", handlePointerMove, true);
      window.removeEventListener("pointerup", finishStroke, true);
      window.removeEventListener("pointercancel", handlePointerCancel, true);
      window.removeEventListener("mousedown", stopMouseEventAfterPen, true);
      window.removeEventListener("mouseup", stopMouseEventAfterPen, true);
      window.removeEventListener("click", stopMouseEventAfterPen, true);
      window.removeEventListener("dblclick", stopMouseEventAfterPen, true);
      window.removeEventListener("contextmenu", stopMouseEventAfterPen, true);
      window.removeEventListener("blur", handleBlur);
    };
  }, []);

  useEffect(() => {
    if (typeof window === "undefined") return;

    const html = document.documentElement;
    const body = document.body;

    const previousHtmlTouchAction = html.style.touchAction;
    const previousBodyTouchAction = body.style.touchAction;
    const previousHtmlUserSelect = html.style.userSelect;
    const previousBodyUserSelect = body.style.userSelect;

    if (penMode) {
      html.style.touchAction = "none";
      body.style.touchAction = "none";
      html.style.userSelect = "none";
      body.style.userSelect = "none";
    }

    return () => {
      html.style.touchAction = previousHtmlTouchAction;
      body.style.touchAction = previousBodyTouchAction;
      html.style.userSelect = previousHtmlUserSelect;
      body.style.userSelect = previousBodyUserSelect;
    };
  }, [penMode]);

  useEffect(() => {
    if (typeof window === "undefined" || !drawingBoardOpen) return;

    const previousOverflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";

    return () => {
      document.body.style.overflow = previousOverflow;
    };
  }, [drawingBoardOpen]);

  useEffect(() => {
    if (typeof window === "undefined") return;

    const handleKeyDown = (event: KeyboardEvent) => {
      const key = event.key.toLowerCase();

      if (event.ctrlKey && event.shiftKey && key === "z") {
        event.preventDefault();
        undoPenStroke();
        return;
      }

      if (event.ctrlKey && event.shiftKey && key === "x") {
        event.preventDefault();
        clearPenCanvas();
        return;
      }
    };

    window.addEventListener("keydown", handleKeyDown);

    return () => window.removeEventListener("keydown", handleKeyDown);
  }, []);

  useEffect(() => {
    if (typeof window === "undefined" || !penCanvasMounted) return;

    const handleResize = () => resizePenCanvas();

    requestAnimationFrame(resizePenCanvas);
    window.addEventListener("resize", handleResize);

    return () => window.removeEventListener("resize", handleResize);
  }, [penCanvasMounted]);

  useEffect(() => {
    if (typeof window === "undefined" || !penCanvasMounted) return;

    const timer = window.setTimeout(() => {
      resizePenCanvas();
    }, 150);

    return () => window.clearTimeout(timer);
  }, [
    penCanvasMounted,
    mode,
    showSaju,
    showDailyCalendar,
    memoOpen,
    drawingBoardOpen,
    sajuResult,
    compatibilityResult,
    result,
    selectedDaewoonKey,
    selectedYearLuckKey,
    showCompatibilityRelations,
  ]);



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

  const formatTimerTime = (seconds: number) => {
    const hours = Math.floor(seconds / 3600);
    const minutes = Math.floor((seconds % 3600) / 60);
    const secs = seconds % 60;

    if (hours > 0) {
      return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(
        2,
        "0",
      )}:${String(secs).padStart(2, "0")}`;
    }

    return `${String(minutes).padStart(2, "0")}:${String(secs).padStart(
      2,
      "0",
    )}`;
  };

  const getTimerInputSeconds = () => {
    const minutes = Number(timerInputMinutes);

    if (!Number.isFinite(minutes) || minutes <= 0) return 0;

    return Math.round(minutes * 60);
  };

  const applyTimerPreset = (minutes: number) => {
    setTimerFinished(false);
    setTimerBlink(false);
    setTimerRunning(false);
    setTimerInputMinutes(String(minutes));
    setTimerRemainingSeconds(minutes * 60);
  };

  const startCountdownTimer = () => {
    const inputSeconds = getTimerInputSeconds();

    if (timerRemainingSeconds <= 0) {
      if (inputSeconds <= 0) return;

      setTimerRemainingSeconds(inputSeconds);
    }

    setTimerFinished(false);
    setTimerBlink(false);
    setTimerRunning(true);
  };

  const resetCountdownTimer = () => {
    setTimerFinished(false);
    setTimerBlink(false);
    setTimerRunning(false);
    setTimerRemainingSeconds(getTimerInputSeconds());
  };

  const playTimerAlarm = () => {
    if (typeof window === "undefined") return;

    const AudioContextClass =
      window.AudioContext || (window as any).webkitAudioContext;

    if (!AudioContextClass) {
      alert("타이머 시간이 끝났습니다.");
      return;
    }

    const audioContext = new AudioContextClass();
    const now = audioContext.currentTime;

    // 알람음 3번 울리기
    [0, 1, 2, 3, 4].forEach((offset) => {
      const oscillator = audioContext.createOscillator();
      const gain = audioContext.createGain();

      oscillator.type = "sine";
      oscillator.frequency.setValueAtTime(880, now + offset);

      gain.gain.setValueAtTime(0.001, now + offset);
      gain.gain.exponentialRampToValueAtTime(0.35, now + offset + 0.04);
      gain.gain.exponentialRampToValueAtTime(0.001, now + offset + 0.45);

      oscillator.connect(gain);
      gain.connect(audioContext.destination);

      oscillator.start(now + offset);
      oscillator.stop(now + offset + 0.5);
    });

    window.setTimeout(() => {
      audioContext.close().catch(() => {});
    }, 5200);
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
      person.birthTimeUnknown ? "시간미상" : person.birthTime || "23:00",
      person.calendarType || "solar",
      person.isLeapMonth ? "윤달" : "평달",
    ].join("|");
  };

  const normalizeRecentPerson = (person: any) => {
    return {
      name: person.name || "",
      gender: person.gender || "남성",
      birthDate: normalizeDateForCalc(person.birthDate || ""),
      birthTime: person.birthTime || "23:00",
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
    { name: "우지안", gender: "여성", birthDate: "1995-11-30", birthTime: "01:29", birthTimeUnknown: false, calendarType: "solar", isLeapMonth: false },
    { name: "남태식", gender: "남성", birthDate: "1981-01-31", birthTime: "13:30", birthTimeUnknown: false, calendarType: "solar", isLeapMonth: false },   
    { name: "잉딩", gender: "여성", birthDate: "2002-08-13", birthTime: "11:59", birthTimeUnknown: false, calendarType: "solar", isLeapMonth: false },
    { name: "성민(민2)", gender: "남성", birthDate: "1993-03-28", birthTime: "04:30", birthTimeUnknown: false, calendarType: "solar", isLeapMonth: false },
    { name: "킹스맨", gender: "남성", birthDate: "1986-03-23", birthTime: "20:30", birthTimeUnknown: false, calendarType: "solar", isLeapMonth: false },
    { name: "탄게", gender: "남성", birthDate: "1993-08-04", birthTime: "23:00", birthTimeUnknown: false, calendarType: "solar", isLeapMonth: false },
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

  useEffect(() => {
    if (typeof window === "undefined") return;

    try {
      const savedMemoPosition = window.localStorage.getItem("sajuMemoPosition");

      if (savedMemoPosition) {
        const parsed = JSON.parse(savedMemoPosition);

        if (typeof parsed?.x === "number" && typeof parsed?.y === "number") {
          setMemoPosition(parsed);
          return;
        }
      }

      setMemoPosition(getDefaultMemoPosition());
    } catch {
      setMemoPosition(getDefaultMemoPosition());
    }
  }, []);

  useEffect(() => {
    if (typeof window === "undefined") return;

    window.localStorage.setItem(
      "sajuMemoPosition",
      JSON.stringify(memoPosition),
    );
  }, [memoPosition]);

  useEffect(() => {
    if (typeof window === "undefined") return;

    try {
      const savedTimerPosition = window.localStorage.getItem("sajuTimerPosition");

      if (savedTimerPosition) {
        const parsed = JSON.parse(savedTimerPosition);

        if (
          typeof parsed?.x === "number" &&
          typeof parsed?.y === "number"
        ) {
          setTimerPosition(parsed);
          return;
        }
      }

      setTimerPosition(getDefaultTimerPosition());
    } catch {
      setTimerPosition(getDefaultTimerPosition());
    }
  }, []);

  useEffect(() => {
    if (typeof window === "undefined") return;

    window.localStorage.setItem(
      "sajuTimerPosition",
      JSON.stringify(timerPosition),
    );
  }, [timerPosition]);

  useEffect(() => {
    if (typeof window === "undefined") return;

    const handleMouseMove = (event: MouseEvent) => {
      if (draggingTimer) {
        setTimerPosition({
          x: event.clientX - timerDragOffsetRef.current.x,
          y: event.clientY - timerDragOffsetRef.current.y,
        });
      }

      if (draggingMemo) {
        setMemoPosition({
          x: event.clientX - memoDragOffsetRef.current.x,
          y: event.clientY - memoDragOffsetRef.current.y,
        });
      }
    };

    const handleMouseUp = () => {
      setDraggingTimer(false);
      setDraggingMemo(false);
    };

    window.addEventListener("mousemove", handleMouseMove);
    window.addEventListener("mouseup", handleMouseUp);

    return () => {
      window.removeEventListener("mousemove", handleMouseMove);
      window.removeEventListener("mouseup", handleMouseUp);
    };
  }, [draggingTimer, draggingMemo]);

  useEffect(() => {
    if (typeof window === "undefined") return;

    const handleTimerRestoreShortcut = (event: KeyboardEvent) => {
      const key = event.key.toLowerCase();

      if (event.ctrlKey && event.altKey && key === "t") {
        event.preventDefault();
        restoreTimerPosition();
      }

      if (event.ctrlKey && event.altKey && key === "m") {
        event.preventDefault();
        restoreMemoPosition();
      }

      if (event.key === "Escape" && drawingBoardOpen) {
        event.preventDefault();
        closeDrawingBoard();
      }
    };

    window.addEventListener("keydown", handleTimerRestoreShortcut);

    return () => {
      window.removeEventListener("keydown", handleTimerRestoreShortcut);
    };
  }, [drawingBoardOpen]);

  useEffect(() => {
    if (!timerRunning) return;

    const interval = window.setInterval(() => {
      setTimerRemainingSeconds((prev) => {
        if (prev <= 1) {
          window.clearInterval(interval);
          setTimerRunning(false);
          setTimerFinished(true);
          playTimerAlarm();

          return 0;
        }

        return prev - 1;
      });
    }, 1000);

    return () => window.clearInterval(interval);
  }, [timerRunning]);

  useEffect(() => {
    if (!timerFinished) {
      setTimerBlink(false);
      return;
    }

    const interval = window.setInterval(() => {
      setTimerBlink((prev) => !prev);
    }, 500);

    return () => window.clearInterval(interval);
  }, [timerFinished]);

  const saveRecentPerson = (person: any) => {
    if (!person?.birthDate) return;

    const normalizedPerson = normalizeRecentPerson(person);

    setRecentPeople((prev) => {
      const key = makePersonKey(normalizedPerson);

      return [
        normalizedPerson,
        ...prev.filter((item) => makePersonKey(item) !== key),
      ].slice(0, 500);
    });
  };

  const hasRecentBirthDate = (person: any) => {
    const birthDate = normalizeDateForCalc(person?.birthDate || "");

    if (!birthDate) return false;

    return recentPeople.some(
      (item) => normalizeDateForCalc(item?.birthDate || "") === birthDate,
    );
  };

  const shouldShowRecentBirthDateNotice =
    recentBirthDateNotice &&
    recentBirthDateNoticeKey === normalizeDateForCalc(form.birthDate || "");

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

  const getFilteredRecentPeople = () => {
    const keyword = recentPeopleSearch.trim().toLowerCase();

    if (!keyword) return recentPeople.slice(0, 20);

    const normalizedKeyword = keyword.replace(/\s/g, "");

    return recentPeople.filter((person) => {
      const searchableValues = [
        person.name || "",
        person.birthDate || "",
        String(person.birthDate || "").replace(/-/g, ""),
        person.birthTimeUnknown ? "시간미상" : person.birthTime || "",
        person.gender || "",
        person.calendarType === "lunar" ? "음력" : "양력",
      ];

      return searchableValues.some((value) =>
        String(value)
          .toLowerCase()
          .replace(/\s/g, "")
          .includes(normalizedKeyword),
      );
    });
  };

  const renderPeopleStoragePanel = (onSelect: (person: any) => void) => {
    const searchKeyword = recentPeopleSearch.trim();
    const searchResults = searchKeyword ? getFilteredRecentPeople() : [];
    const visibleRecentPeople = recentPeople.slice(0, 20);

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
              <input
                type="text"
                value={recentPeopleSearch}
                onChange={(event) =>
                  setRecentPeopleSearch(event.target.value)
                }
                placeholder="이름, 생년월일, 출생시간 검색"
                className="w-full rounded-xl border border-[#ead8c4] bg-white px-4 py-3 text-2xl font-bold text-black outline-none placeholder:text-zinc-400"
              />
              <div className="mt-2 text-xl font-bold text-[#6b3f24]">
                {searchKeyword
                  ? `검색 결과 ${searchResults.length}명`
                  : `최근 본 사람 전체 ${recentPeople.length}명`}
              </div>

              {searchKeyword && (
                <div className="mt-3 flex flex-wrap gap-2">
                  {searchResults.length > 0 ? (
                    searchResults.map((person) =>
                      renderPeopleButton(person, onSelect, {
                        favoriteButton: true,
                        removeRecentButton: true,
                      }),
                    )
                  ) : (
                    <div className="text-2xl font-bold text-zinc-400">
                      검색 결과가 없습니다.
                    </div>
                  )}
                </div>
              )}
            </div>

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
                <>
                  <div className="mt-3 flex flex-wrap gap-2">
                    {recentPeople.length === 0 ? (
                      <div className="text-2xl font-bold text-zinc-400">
                        최근 본 사람이 없습니다.
                      </div>
                    ) : (
                      visibleRecentPeople.map((person) =>
                        renderPeopleButton(person, onSelect, {
                          favoriteButton: true,
                          removeRecentButton: true,
                        }),
                      )
                    )}
                  </div>
                </>
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

    const shouldOpenTimer = window.confirm("만세력 계산 전에 타이머를 세팅하시겠습니까?");

    if (shouldOpenTimer) {
      setTimerOpen(true);
      setTimerFinished(false);
      setTimerBlink(false);
    }

    const alreadyViewedBirthDate = hasRecentBirthDate(form);
    setRecentBirthDateNotice(alreadyViewedBirthDate);
    setRecentBirthDateNoticeKey(normalizeDateForCalc(form.birthDate || ""));

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

      setResult(data.result || "결과가 비어 있습니다.");
    } catch (error) {
      console.error(error);
      setResult("오류가 발생했습니다. 콘솔을 확인하세요.");
    } finally {
      setLoading(false);
    }
  }

  const waitForRender = () =>
    new Promise<void>((resolve) => {
      requestAnimationFrame(() => {
        requestAnimationFrame(() => {
          setTimeout(() => resolve(), 100);
        });
      });
    });

  const safeFileName = (value: string) =>
    String(value || "사주")
      .replace(/[\\/:*?"<>|]/g, "_")
      .trim();

  const createWordBlob = async () => {
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
                  text: `이름: ${form.name || "이름없음"}`,
                  size: WORD_FONT.body,
                }),
              ],
            }),

            new Paragraph({
              children: [
                new TextRun({
                  text: `생년월일: ${normalizeDateForCalc(form.birthDate)} ${
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
                    new TableCell({ children: [new Paragraph("시주")] }),
                    new TableCell({ children: [new Paragraph("일주")] }),
                    new TableCell({ children: [new Paragraph("월주")] }),
                    new TableCell({ children: [new Paragraph("년주")] }),
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

    return Packer.toBlob(doc);
  };

  const downloadCaptureZip = async () => {
    if (!sajuResult || !overviewCaptureRef.current) {
      alert("먼저 만세력을 계산해 주세요.");
      return;
    }

    const style = document.createElement("style");
    style.innerHTML = `
      [data-capture-target],
      [data-capture-target] * {
        transition: none !important;
        animation: none !important;
      }
    `;
    document.head.appendChild(style);

    const previousDaewoonKey = selectedDaewoonKey.main ?? null;
    const previousYearLuckKey = selectedYearLuckKey.main ?? null;

    try {
      const zip = new JSZip();
      const name = safeFileName(form.name || "이름없음");
      const birth = safeFileName(
        normalizeDateForCalc(form.birthDate).replaceAll("-", ""),
      );

      const wordBlob = await createWordBlob();
      zip.file(`${name}_${birth}_사주분석.docx`, wordBlob);

      flushSync(() => {
        setSelectedDaewoonKey((prev) => ({
          ...prev,
          main: null,
        }));

        setSelectedYearLuckKey((prev) => ({
          ...prev,
          main: null,
        }));
      });

      await waitForRender();

      const baseDataUrl = await toPng(overviewCaptureRef.current, {
        cacheBust: true,
        pixelRatio: 2,
        backgroundColor: "#ffffff",
      });

      zip.file(
        `${name}_${birth}_00_기본_사주팔자_오행_귀인.png`,
        baseDataUrl.split(",")[1],
        { base64: true },
      );

      for (const item of sajuResult.daewoon) {
        const daewoonKey = `main-daewoon-${item.index}`;

        flushSync(() => {
          setSelectedDaewoonKey((prev) => ({
            ...prev,
            main: daewoonKey,
          }));

          setSelectedYearLuckKey((prev) => ({
            ...prev,
            main: null,
          }));
        });

        await waitForRender();

        if (!luckCaptureRef.current) continue;

        const dataUrl = await toPng(luckCaptureRef.current, {
          cacheBust: true,
          pixelRatio: 2,
          backgroundColor: "#ffffff",
        });

        zip.file(
          `${name}_${birth}_${String(item.index + 1).padStart(
            2,
            "0",
          )}_대운_${safeFileName(item.startAgeText)}_년운.png`,
          dataUrl.split(",")[1],
          { base64: true },
        );
      }

      const blob = await zip.generateAsync({ type: "blob" });

      saveAs(blob, `${name}_${birth}_사주_대운년운_워드포함.zip`);
    } catch (error) {
      console.error(error);
      alert("캡쳐 ZIP 저장 중 오류가 발생했습니다.");
    } finally {
      style.remove();

      flushSync(() => {
        setSelectedDaewoonKey((prev) => ({
          ...prev,
          main: previousDaewoonKey,
        }));

        setSelectedYearLuckKey((prev) => ({
          ...prev,
          main: previousYearLuckKey,
        }));
      });
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
  const formatLuckTenGod = (tenGod: string) => {
    return tenGod === "일간" ? "비견" : tenGod;
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

  const TWELVE_STAGE_BY_DAY_STEM: Record<string, Record<string, string>> = {
    갑: { 해: "장생", 자: "목욕", 축: "관대", 인: "건록", 묘: "제왕", 진: "쇠", 사: "병", 오: "사", 미: "묘", 신: "절", 유: "태", 술: "양" },
    을: { 오: "장생", 사: "목욕", 진: "관대", 묘: "건록", 인: "제왕", 축: "쇠", 자: "병", 해: "사", 술: "묘", 유: "절", 신: "태", 미: "양" },
    병: { 인: "장생", 묘: "목욕", 진: "관대", 사: "건록", 오: "제왕", 미: "쇠", 신: "병", 유: "사", 술: "묘", 해: "절", 자: "태", 축: "양" },
    정: { 유: "장생", 신: "목욕", 미: "관대", 오: "건록", 사: "제왕", 진: "쇠", 묘: "병", 인: "사", 축: "묘", 자: "절", 해: "태", 술: "양" },
    무: { 인: "장생", 묘: "목욕", 진: "관대", 사: "건록", 오: "제왕", 미: "쇠", 신: "병", 유: "사", 술: "묘", 해: "절", 자: "태", 축: "양" },
    기: { 유: "장생", 신: "목욕", 미: "관대", 오: "건록", 사: "제왕", 진: "쇠", 묘: "병", 인: "사", 축: "묘", 자: "절", 해: "태", 술: "양" },
    경: { 사: "장생", 오: "목욕", 미: "관대", 신: "건록", 유: "제왕", 술: "쇠", 해: "병", 자: "사", 축: "묘", 인: "절", 묘: "태", 진: "양" },
    신: { 자: "장생", 해: "목욕", 술: "관대", 유: "건록", 신: "제왕", 미: "쇠", 오: "병", 사: "사", 진: "묘", 묘: "절", 인: "태", 축: "양" },
    임: { 신: "장생", 유: "목욕", 술: "관대", 해: "건록", 자: "제왕", 축: "쇠", 인: "병", 묘: "사", 진: "묘", 사: "절", 오: "태", 미: "양" },
    계: { 묘: "장생", 인: "목욕", 축: "관대", 자: "건록", 해: "제왕", 술: "쇠", 유: "병", 신: "사", 미: "묘", 오: "절", 사: "태", 진: "양" },
  };

  const getTwelveStage = (dayStem: string, branch: string) => {
    const normalizedDayStem = normalizeStem(dayStem);
    const normalizedBranch = normalizeBranch(branch);

    return TWELVE_STAGE_BY_DAY_STEM[normalizedDayStem]?.[normalizedBranch] || "";
  };

  const getLuckTwelveStage = (targetSaju: any, branch: string) => {
    const dayStem = normalizeStem(
      targetSaju?.day?.stem || String(targetSaju?.day?.ganji || "").slice(0, 1),
    );

    return getTwelveStage(dayStem, branch);
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
        twelveStage: getTwelveStage(dayStem, ganji.branch),
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
        twelveStage: getTwelveStage(dayStem, ganji.branch),
      };
    });
  };
  const formatOriginalStemTenGod = (label: string, tenGod: string) => {
    if (label === "일주") return "일간";

    return tenGod === "일간" ? "비견" : tenGod;
  };

  const buildSajuItems = (targetSaju: any) => {
    if (!targetSaju) return [];

    return addShinsalToItems(
      addBranchRelationsToItems([
        {
          label: "시주",
          data: targetSaju.hour,
          targetSaju,
          tenGodStem: formatOriginalStemTenGod(
            "시주",
            targetSaju.tenGods?.hourStem ?? "",
          ),
          tenGodBranch: targetSaju.tenGods?.hourBranch ?? "",
          twelveStage: targetSaju.twelveStages?.hour ?? "",
        },
        {
          label: "일주",
          data: targetSaju.day,
          targetSaju,
          tenGodStem: formatOriginalStemTenGod(
            "일주",
            targetSaju.tenGods.dayStem,
          ),
          tenGodBranch: targetSaju.tenGods.dayBranch,
          twelveStage: targetSaju.twelveStages.day,
        },
        {
          label: "월주",
          data: targetSaju.month,
          targetSaju,
          tenGodStem: formatOriginalStemTenGod(
            "월주",
            targetSaju.tenGods.monthStem,
          ),
          tenGodBranch: targetSaju.tenGods.monthBranch,
          twelveStage: targetSaju.twelveStages.month,
        },
        {
          label: "년주",
          data: targetSaju.year,
          targetSaju,
          tenGodStem: formatOriginalStemTenGod(
            "년주",
            targetSaju.tenGods.yearStem,
          ),
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
                  {formatLuckTenGod(item.stemTenGod)}
                </div>

                <div
                  className={
                    selected
                      ? `${FONT.daewoonTenGod} ${WEIGHT.daewoonTenGod} ${COLOR.daewoonTenGodSelected}`
                      : `${FONT.daewoonTenGod} ${WEIGHT.daewoonTenGod} ${COLOR.daewoonTenGod}`
                  }
                >
                  {formatLuckTenGod(item.branchTenGod)}
                </div>

                <div
                  className={
                    selected
                      ? `mt-1 ${FONT.twelveState} ${WEIGHT.twelveState} text-white`
                      : `mt-1 ${FONT.twelveState} ${WEIGHT.twelveState} ${COLOR.twelveState}`
                  }
                >
                  {getLuckTwelveStage(targetSaju, item.ganji.branch)}
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
                        {formatLuckTenGod(yearLuck.stemTenGod)}
                      </div>

                      <div
                        className={
                          selected
                            ? `${FONT.yearLuckTenGod} ${WEIGHT.yearLuckTenGod} ${COLOR.yearLuckTenGodSelected}`
                            : `${FONT.yearLuckTenGod} ${WEIGHT.yearLuckTenGod} ${COLOR.yearLuckTenGod}`
                        }
                      >
                        {formatLuckTenGod(yearLuck.branchTenGod)}
                      </div>

                      <div
                        className={
                          selected
                            ? `mt-1 ${FONT.twelveState} ${WEIGHT.twelveState} text-white`
                            : `mt-1 ${FONT.twelveState} ${WEIGHT.twelveState} ${COLOR.twelveState}`
                        }
                      >
                        {yearLuck.twelveStage}
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
                          {formatLuckTenGod(monthLuck.stemTenGod)}
                        </div>

                        <div
                          className={`${FONT.monthLuckTenGod} ${WEIGHT.monthLuckTenGod} ${COLOR.monthLuckTenGod}`}
                        >
                          {formatLuckTenGod(monthLuck.branchTenGod)}
                        </div>

                        <div
                          className={`mt-1 ${FONT.twelveState} ${WEIGHT.twelveState} ${COLOR.twelveState}`}
                        >
                          {monthLuck.twelveStage}
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


  const renderTarotMode = () => {
    const selectedCard = TAROT_CARDS.find((card) => card.id === selectedTarotCard) || null;

    return (
      <section className="rounded-3xl border border-[#ead8c4] bg-[#fffaf3] p-5 shadow-inner">
        <div className="mb-5">
          <h2 className={`${FONT.sectionTitle} ${WEIGHT.sectionTitle} ${COLOR.sectionTitle}`}>
            타로 카드 해석
          </h2>
          <p className={`mt-2 ${FONT.body} ${WEIGHT.body} ${COLOR.body}`}>
            카드를 클릭하면 정방향과 역방향 해석이 표시됩니다.
          </p>
        </div>

        {selectedCard && (
          <div className="sticky top-6 z-[54] mb-6 rounded-3xl border-2 border-[#6b3f24] bg-white p-5 shadow-xl">
            <div className="flex items-start justify-between gap-4">
              <div>
                <div className="text-2xl font-bold text-[#9a7657]">
                  {selectedCard.group} · {selectedCard.number}
                </div>
                <h3 className="mt-1 text-5xl font-bold text-[#2b1d12]">
                  {selectedCard.name}
                </h3>
                <div className="mt-1 text-3xl font-bold text-[#6b3f24]">
                  {selectedCard.korean}
                </div>
              </div>

              <button
                type="button"
                onClick={() => setSelectedTarotCard(null)}
                className="rounded-xl bg-zinc-100 px-4 py-2 text-2xl font-bold text-zinc-700 hover:bg-zinc-200"
              >
                닫기
              </button>
            </div>

            <div className="mt-5 grid grid-cols-2 gap-4">
              <div className="rounded-2xl border border-emerald-200 bg-emerald-50 p-5">
                <div className="text-3xl font-bold text-emerald-800">정방향</div>
                <div className="mt-3 text-3xl font-semibold leading-relaxed text-black">
                  {selectedCard.upright}
                </div>
              </div>

              <div className="rounded-2xl border border-red-200 bg-red-50 p-5">
                <div className="text-3xl font-bold text-red-800">역방향</div>
                <div className="mt-3 text-3xl font-semibold leading-relaxed text-black">
                  {selectedCard.reversed}
                </div>
              </div>
            </div>
          </div>
        )}

        <div className="space-y-7">
          {TAROT_GROUPS.map((group) => {
            const cards = TAROT_CARDS.filter((card) => card.group === group);

            return (
              <div key={group}>
                <h3 className="mb-3 text-4xl font-bold text-[#6b3f24]">{group}</h3>
                <div className={group === "메이저" ? "grid grid-cols-6 gap-3" : "grid grid-cols-7 gap-3"}>
                  {cards.map((card) => {
                    const selected = selectedTarotCard === card.id;

                    return (
                      <button
                        key={card.id}
                        type="button"
                        onClick={() => setSelectedTarotCard(card.id)}
                        className={`min-h-[130px] rounded-2xl border p-3 text-center shadow-sm transition ${
                          selected
                            ? "border-[#6b3f24] bg-[#6b3f24] text-white shadow-md"
                            : "border-[#ead8c4] bg-white text-[#2b1d12] hover:bg-[#f3e1cf]"
                        }`}
                      >
                        <div className={`text-xl font-bold ${selected ? "text-white/80" : "text-[#9a7657]"}`}>
                          {card.number}
                        </div>
                        <div className="mt-1 text-2xl font-bold leading-tight">{card.name}</div>
                        <div className={`mt-2 text-xl font-bold ${selected ? "text-white/90" : "text-[#6b3f24]"}`}>
                          {card.korean}
                        </div>
                      </button>
                    );
                  })}
                </div>
              </div>
            );
          })}
        </div>
      </section>
    );
  };

  return (
    <>
      {drawingBoardOpen && (
        <div className="fixed inset-0 z-[9997] bg-white">
          <div className="fixed left-1/2 top-4 z-[10000] flex -translate-x-1/2 items-center gap-3 rounded-2xl border border-zinc-300 bg-white/95 px-4 py-3 shadow-2xl">
            <div className="mr-2 text-2xl font-bold text-black">와콤 그림판</div>

            <button
              type="button"
              onClick={undoPenStroke}
              className="rounded-xl bg-zinc-100 px-4 py-2 text-xl font-bold text-black hover:bg-zinc-200"
            >
              실행취소
            </button>

            <button
              type="button"
              onClick={() => {
                if (confirm("그림판을 모두 지우시겠습니까?")) clearPenCanvas();
              }}
              className="rounded-xl bg-red-50 px-4 py-2 text-xl font-bold text-red-700 hover:bg-red-100"
            >
              전체 지우기
            </button>

            <button
              type="button"
              onClick={closeDrawingBoard}
              className="rounded-xl bg-black px-4 py-2 text-xl font-bold text-white hover:bg-zinc-800"
            >
              닫기
            </button>
          </div>

          <div className="fixed bottom-5 left-1/2 z-[10000] -translate-x-1/2 rounded-full bg-black/75 px-5 py-2 text-lg font-bold text-white">
            펜촉: 그리기 · 사이드 버튼: 지우기 · Esc: 닫기
          </div>
        </div>
      )}

      {penCanvasMounted && (
        <canvas
          ref={penCanvasRef}
          className="absolute left-0 top-0 z-[9998]"
          style={{
            pointerEvents: "none",
            touchAction: "none",
          }}
        />
      )}

      {/* {penMode && (
        <div className="fixed right-6 top-6 z-[10000] rounded-full bg-red-600 px-5 py-3 text-2xl font-bold text-white shadow-2xl">
          🔴 PEN MODE
        </div>
      )} */}

    <main className="min-h-screen bg-[#f7efe3] px-5 py-10 text-[#2b1d12]">
      <div className="mx-auto w-[1400px] min-w-[1400px] rounded-3xl bg-white p-6 shadow-xl">
        <h1
          className={`text-center ${FONT.pageTitle} ${WEIGHT.pageTitle} ${COLOR.pageTitle}`}
        >
          {mode === "saju"
            ? "사주 분석기"
            : mode === "compatibility"
              ? "궁합 분석"
              : "타로 카드 해석"}
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
              setMode("tarot");
              setResult("");
              setShowSaju(false);
            }}
            className={`rounded-xl py-3 ${FONT.modeButtonText} ${WEIGHT.modeButtonText} ${COLOR.modeButtonText} transition ${
              mode === "tarot"
                ? "bg-[#6b3f24] text-white shadow"
                : "text-[#6b3f24]"
            }`}
          >
            타로 모드
          </button>
        </div>

        {mode !== "tarot" && (
        <div className="sticky top-6 z-[55] mt-4 overflow-hidden rounded-2xl border border-[#d7c4ad] bg-white/95 px-4 py-2 shadow-lg backdrop-blur">
          {mode !== "compatibility" ? (
            <div className={`flex items-center justify-center gap-4 whitespace-nowrap ${FONT.floatingInfo} font-bold text-[#2b1d12]`}>
              <span>{form.name || "이름 미입력"}</span>
              <span className="text-[#b59474]">|</span>
              <span>{form.gender}</span>
              <span className="text-[#b59474]">|</span>
              <span>{form.calendarType === "solar" ? "양력" : form.isLeapMonth ? "음력 윤달" : "음력"}</span>
              <span>{form.birthDate || "생년월일 미입력"}</span>
              <span>{form.birthTimeUnknown ? "시간 미상" : form.birthTime || "시간 미입력"}</span>
            </div>
          ) : (
            <div className={`flex items-center justify-center gap-5 whitespace-nowrap ${FONT.floatingInfo} font-bold text-[#2b1d12]`}>
              <span className="rounded-lg bg-[#fff4e8] px-3 py-1">
                {compatibilityForm.left.name || "이름 미입력"} · {compatibilityForm.left.gender} · {compatibilityForm.left.calendarType === "solar" ? "양력" : compatibilityForm.left.isLeapMonth ? "음력 윤달" : "음력"} {compatibilityForm.left.birthDate || "생년월일 미입력"} {compatibilityForm.left.birthTimeUnknown ? "시간 미상" : compatibilityForm.left.birthTime || "시간 미입력"}
              </span>
              <span className="rounded-lg bg-[#f4efe9] px-3 py-1">
                {compatibilityForm.right.name || "이름 미입력"} · {compatibilityForm.right.gender} · {compatibilityForm.right.calendarType === "solar" ? "양력" : compatibilityForm.right.isLeapMonth ? "음력 윤달" : "음력"} {compatibilityForm.right.birthDate || "생년월일 미입력"} {compatibilityForm.right.birthTimeUnknown ? "시간 미상" : compatibilityForm.right.birthTime || "시간 미입력"}
              </span>
            </div>
          )}
        </div>
        )}

        <div className="mt-8 space-y-4">
          {mode === "tarot" && renderTarotMode()}

          {mode === "saju" && (
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
                  placeholder="1993-08-04"
                  className={`w-[360px] shrink-0 rounded-xl border p-3 ${FONT.inputText}`}
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
                        placeholder="1993-08-04"
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
                      placeholder="23:00"
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



          {mode === "saju" && showSaju && (
            <section className="mt-6 rounded-3xl border border-[#ead8c4] bg-[#fffaf3] p-5 shadow-inner">
              {shouldShowRecentBirthDateNotice && (
                <div className="mb-4 rounded-xl bg-red-50 px-4 py-3 text-3xl font-bold text-red-600">
                  최근에 본 이력이 있는 생년월일 입니다
                </div>
              )}

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
                  <div>
                    <div ref={overviewCaptureRef} data-capture-target>
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
                    </div>

                    <div ref={luckCaptureRef} data-capture-target>
                      {renderLuckPanel(
                        sajuResult,
                        "main",
                        normalizeDateForCalc(form.birthDate),
                      )}
                    </div>
                  </div>

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

                 {/*<button
                    type="button"
                    onClick={downloadCaptureZip}
                    className={`mt-6 w-full rounded-xl bg-black px-5 py-3 ${FONT.buttonText} ${WEIGHT.buttonText} ${COLOR.buttonText} shadow-md`}
                  >
                    저장하기
                  </button>*/}
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

      {!timerOpen && (
        <button
          type="button"
          onClick={() => {
            setTimerOpen(true);
            setTimerFinished(false);
            setTimerBlink(false);
          }}
          className={`fixed right-6 top-24 z-[60] rounded-l-2xl rounded-r-md px-4 py-4 text-2xl font-bold text-white shadow-2xl transition ${
            timerFinished
              ? timerBlink
                ? "bg-red-600"
                : "bg-black"
              : "bg-black hover:bg-zinc-800"
          }`}
        >
          <div>타이머 열기</div>
          <div
            className={`mt-2 text-4xl ${
              timerRemainingSeconds <= 60 ? "text-red-400" : "text-white"
            }`}
          >
            {formatTimerTime(timerRemainingSeconds)}
          </div>
        </button>
      )}

      {timerOpen && (
        <button
          type="button"
          onClick={restoreTimerPosition}
          className="fixed right-6 top-24 z-[61] rounded-l-2xl rounded-r-md bg-[#6b3f24] px-4 py-3 text-xl font-bold text-white shadow-2xl transition hover:bg-[#4a2f1c]"
          title="단축키: Ctrl + Alt + T"
        >
          타이머 위치복구
        </button>
      )}

      {timerOpen && (
        <div
          className="fixed z-[60] w-[340px] rounded-3xl border border-[#ead8c4] bg-[#fffaf3] p-4 shadow-2xl"
          style={{
            left: timerPosition.x,
            top: timerPosition.y,
          }}
        >
          <div
            className="mb-3 flex cursor-move items-center justify-between rounded-xl bg-[#f3e1cf] p-2"
            onMouseDown={(event) => {
              setDraggingTimer(true);

              timerDragOffsetRef.current = {
                x: event.clientX - timerPosition.x,
                y: event.clientY - timerPosition.y,
              };
            }}
          >
            <h2 className="text-3xl font-bold text-[#6b3f24]">상담 타이머</h2>

            <div className="flex gap-2">
              <button
                type="button"
                onClick={(event) => {
                  event.stopPropagation();
                  restoreTimerPosition();
                }}
                onMouseDown={(event) => event.stopPropagation()}
                className="rounded-xl bg-white px-3 py-2 text-lg font-bold text-[#6b3f24] shadow-sm"
                title="단축키: Ctrl + Alt + T"
              >
                위치
              </button>

              <button
                type="button"
                onClick={(event) => {
                  event.stopPropagation();
                  setTimerOpen(false);
                }}
                onMouseDown={(event) => event.stopPropagation()}
                className="rounded-xl bg-white px-3 py-2 text-xl font-bold text-[#6b3f24] shadow-sm"
              >
                닫기
              </button>
            </div>
          </div>

         <div className="rounded-2xl bg-white p-4 shadow-inner">
  <div className="flex items-center justify-center gap-3">
    <div className="flex flex-col gap-2">
      {[1, 5, 10].map((minute) => (
        <button
          key={`minus-${minute}`}
          type="button"
          onClick={() =>
            setTimerRemainingSeconds((prev) =>
              Math.max(0, prev - minute * 60),
            )
          }
          className="rounded-lg bg-red-50 px-2 py-1 text-lg font-bold text-red-700 hover:bg-red-100"
        >
          -{minute}
        </button>
      ))}
    </div>

    <div className="min-w-[150px] text-center text-5xl font-bold text-black">
      {formatTimerTime(timerRemainingSeconds)}
    </div>

    <div className="flex flex-col gap-2">
      {[1, 5, 10].map((minute) => (
        <button
          key={`plus-${minute}`}
          type="button"
          onClick={() =>
            setTimerRemainingSeconds((prev) => prev + minute * 60)
          }
          className="rounded-lg bg-zinc-100 px-2 py-1 text-lg font-bold text-[#6b3f24] hover:bg-[#f3e1cf]"
        >
          +{minute}
        </button>
      ))}
    </div>
  </div>
</div>

          <div className="mt-4 grid grid-cols-3 gap-2">
            {[5, 10, 30].map((minutes) => (
              <button
                type="button"
                key={minutes}
                onClick={() => applyTimerPreset(minutes)}
                className="rounded-xl bg-white py-3 text-2xl font-bold text-[#6b3f24] shadow-sm transition hover:bg-[#f3e1cf]"
              >
                {minutes}분
              </button>
            ))}
          </div>

          <div className="mt-3 flex items-center gap-2">
            <input
              type="number"
              min="1"
              step="1"
              value={timerInputMinutes}
              onChange={(event) => {
                const nextValue = event.target.value;

                setTimerFinished(false);
                setTimerBlink(false);
                setTimerInputMinutes(nextValue);

                if (!timerRunning) {
                  const nextMinutes = Number(nextValue);

                  setTimerRemainingSeconds(
                    Number.isFinite(nextMinutes) && nextMinutes > 0
                      ? Math.round(nextMinutes * 60)
                      : 0,
                  );
                }
              }}
              className="min-w-0 flex-1 rounded-xl border border-[#ead8c4] bg-white p-3 text-2xl font-bold text-black outline-none"
            />

            <span className="text-2xl font-bold text-[#6b3f24]">분</span>
          </div>

          <div className="mt-4 grid grid-cols-3 gap-2">
            <button
              type="button"
              onClick={startCountdownTimer}
              disabled={timerRemainingSeconds <= 0 && getTimerInputSeconds() <= 0}
              className="rounded-xl bg-[#6b3f24] py-3 text-2xl font-bold text-white shadow-sm disabled:opacity-40"
            >
              시작
            </button>

            <button
              type="button"
              onClick={() => setTimerRunning(false)}
              className="rounded-xl bg-zinc-700 py-3 text-2xl font-bold text-white shadow-sm"
            >
              정지
            </button>

            <button
              type="button"
              onClick={resetCountdownTimer}
              className="rounded-xl bg-red-700 py-3 text-2xl font-bold text-white shadow-sm"
            >
              초기화
            </button>
          </div>

          <div className="mt-3 text-center text-xl font-bold text-[#6b3f24]">
            위쪽 제목줄을 잡고 드래그 · 위치복구 Ctrl+Alt+T
          </div>
        </div>
      )}

      {!drawingBoardOpen && (
        <button
          type="button"
          onClick={openDrawingBoard}
          className="fixed bottom-6 left-6 z-[60] rounded-2xl bg-black px-5 py-4 text-2xl font-bold text-white shadow-2xl transition hover:bg-zinc-800"
        >
          그림판 열기
        </button>
      )}

      {!drawingBoardOpen && (
        <button
          type="button"
          onClick={() => setMemoOpen((prev) => !prev)}
          className="fixed right-6 top-[240px] z-[10000] rounded-l-2xl rounded-r-md bg-[#6b3f24] px-4 py-4 text-2xl font-bold text-white shadow-2xl transition hover:bg-[#4a2f1c]"
        >
          {memoOpen ? "메모 닫기" : "메모 열기"}
        </button>
      )}

      {memoOpen && !drawingBoardOpen && (
        <button
          type="button"
          onClick={restoreMemoPosition}
          className="fixed right-6 top-[320px] z-[10000] rounded-l-2xl rounded-r-md bg-[#6b3f24] px-4 py-4 text-2xl font-bold text-white shadow-2xl transition hover:bg-[#4a2f1c]"
          title="메모장이 화면 밖으로 나갔을 때 복구합니다. 단축키: Ctrl + Alt + M"
        >
          메모 위치복구
        </button>
      )}

      {memoOpen && (
        <div
          className="fixed z-[60] flex w-[420px] flex-col rounded-3xl border border-[#ead8c4] bg-[#fffaf3] p-4 shadow-2xl"
          style={{
            left: memoPosition.x,
            top: memoPosition.y,
            height: "min(720px, calc(100vh - 40px))",
          }}
        >
          <div
            className="mb-3 flex cursor-move items-center justify-between rounded-xl bg-[#f3e1cf] p-2"
            onMouseDown={(event) => {
              setDraggingMemo(true);
              memoDragOffsetRef.current = {
                x: event.clientX - memoPosition.x,
                y: event.clientY - memoPosition.y,
              };
            }}
          >
            <h2 className="text-3xl font-bold text-[#6b3f24]">상담 메모장</h2>

            <div className="flex gap-2">
              <button
                type="button"
                onClick={(event) => {
                  event.stopPropagation();
                  setMemoOpen(false);
                }}
                onMouseDown={(event) => event.stopPropagation()}
                className="rounded-xl bg-white px-3 py-2 text-xl font-bold text-[#6b3f24] shadow-sm"
              >
                닫기
              </button>
            </div>
          </div>

          <textarea
            value={memoText}
            onChange={(e) => setMemoText(e.target.value)}
            placeholder="상담 중 메모를 입력하세요. 새로고침 후에도 유지됩니다."
            className="min-h-0 flex-1 resize-none rounded-2xl border border-[#ead8c4] bg-white p-4 text-2xl leading-relaxed text-black outline-none"
          />

          <div className="mt-3 text-right text-xl font-bold text-[#6b3f24]">
            제목줄 드래그 · 자동 저장 · 위치복구 Ctrl+Alt+M
          </div>
        </div>
      )}
    </main>
    </>
  );
}
