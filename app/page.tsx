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
  const [tarotSearch, setTarotSearch] = useState("");


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

  const TAROT_SOURCE_DETAILS: Record<string, { description: string; oneCard: string }> = {
    "major-0": { description: "한 젊은이가 왼 손엔 흰 장미를, 오른쪽 어깨엔 배낭을 메고 벼랑 끝에 서 있다. 그의 곁에는 하얀 개 한 마리(믿음직한 후견인)가 있고, 그의 머리 위에는 태양(신의 은총)이 밝게 빛 나고 있다. 그가 옷 속에 입고 있는 하얀 옷과 흰 장미는 그 의 순수함을 나타내고, 머리의 월계수는 성공을, 붉은색 깃털 은 열정을, 장화는 자유를 의미한다. 뒤쪽의 산은 아직 다다르지 않은 정신적 세계를 주위의 노 란색은 지성과 의지력, 광기의 요소를 의미한다. 언덕위에 있 다는 것은 언덕 아래 어떤 방향으로든 떠날 수 있다는 것을 의미하면서 또한 그 길이 항상 좋지만은 않다는 것을 의미하 기도 한다. 이 카드는 인생에서 새로운 세계로 첫발을 내딛는 초보자이 며, 비록 벼랑 끝에 서 있지만 그의 표정에는 두려움이 없으 며 희망으로 가득 차 있으나 어이없는 실수에 대한 경고도 함께 하고 있다.", oneCard: "Q: 취직이 될까? A: 아직 취직할 정도의 스펙이 아니다. 더욱 실력을 갈고 닦아야 한다. Q: 며칠째 속이 쓰린 병이 있는데? A: 너무 쉽게 생각하다가 큰 병으로 발전할 수 있으니 세심한 관찰이 필요하다. Q: 남자와 사귄지 몇 년 되었는데 결혼은 언제? A: 아직 준비가 안되었으니 기다려야 한다. 남자는 자유연애주의일 수도 있으니 조심하라." },
    "major-1": { description: "한 젊은이가 마법사 복장을 하고 자신감이 넘치는 모습을 하고 있다. 그의 머리 위에는 영원의 상징인 뫼비우스의 띠가 있으며, 그의 허리에는 뱀 모양의 허리띠가 있는데 이것은 내 면의 사악함을 나타낸다. 그가 입고 있는 붉고 흰 옷은 열정 과 순수함을 나타내고 탁자위에 있는 4가지 상징들은 불, 물, 공기, 흙의 4원소를 상징한다. 이것은 그가 모든 것을 컨트롤 할 수 있는 능력의 소유자임을 나타낸다. 오른손은 하늘을 향 해 있고 왼손은 땅을 가리키고 있는데 이것은 자신이 하늘과 인간의 매개체 (영매)를 의미하기도 하고, 위의 것은 끌어내리 고 밑의 것은 끌어 올리는 것을 의미한다. 일단 마법사의 가 장 큰 의미는 무엇이든지 알고 있는 사람, 능력 있는 사람이 다. 그러나 자신의 능력을 과신해서 일을 그르치는 경우도 있 음을 주의해야 한다.", oneCard: "Q: 취직이 될까? A: 능력 있는 사람이고, 스펙도 좋으므로 원하는 회사에 취업하겠다. Q: 사업을 하는데 자금회전이 원활치 못한데 언제쯤 나아질까? A: 당신은 위험에서 벗어날 충분한 능력이 있으며 헤쳐 나갈 수 있는 잠재력이 충분하다. 걱정하지 말라. Q: 어머니가 암 선고를 받으셨는데 나을 수 있나? A: 그 분야에서 최고로 능력 있는 의사를 만나 병을 고칠 수 있을 것이다." },
    "major-2": { description: "아름다운 여인이 3가지 모양의 달을 의미하는 왕관을 쓰고, 물과 무의식 그리고 성모마리아를 상징하는 푸른색 베일을 드리운 채 두루마리를 들고, 초승달을 밟고서 바다를 뒤로 한 채 왕좌에 앉아있다. 그녀의 양쪽에는 두 개의 기둥이 서 있 고, 뒤에는 석류와 대추야자가 그려진 베일이 걸려있다. 양쪽 에 보이는 두 개의 기둥은 선과 악, 빛과 어둠, 진실과 거짓, 긍정과 부정 등 이중성을 상징한다. 여제 뒤에 걸려있는 베일은 무의식의 왕국으로 들어가는 입 구를 상징한다. 거기에 그려진 석류는 그리스신화에 나오는 지하왕국의 여신인 페르세포네를 나타내는 것이며, 비옥함을 뜻하기도 한다. 이것은 여제가 무의식과 깊은 연관이 있다는 것을 나타낸다. 여제가 들고 있는 두루마리는 숨겨진 진실을, TORA라는 글자는 성서라는 뜻을 나타낸다. 이 여인은 냉정하고도 도도해 보이는데, 그런 면이 연애를 잘 못하게 하는 원인이 되기도 한다. 또한 여사제의 경우는 자신의 신분으로 인해 행동에 제약이 많고, 그런 제약 속에 있으므로 비밀스러운 면도 내포하고 있다.", oneCard: "Q: 친구가 빌린 돈을 갚지 않고 있는데 받을 수 있나? A: 받기 힘들다. 왜냐하면 여사제는 정신적 일을 하는 사람이라 금전과 인연이 약하기 때문이다. Q: 최근 들어 소화가 잘 안되는데 큰 병인가? A: 신경성 질환을 의심할 수 있지만, 종합검진을 받기를 권한다. (밝혀내기 어려운 질병일 수 있다.) Q: 최근에 만난 여자 친구와 밀월 여행을 가고 싶은데 가능한가? A: 여자 친구는 당신을 그저 친구로 대하고 있으며, 거짓말을 못하는 성격이라 부모님께 허락 받기도 힘들 것이다. (여사제는 종교인이므로 연애와는 무관한 사람이다.)" },
    "major-3": { description: "임신한 듯이 보이는 한 여인이 풍요로운 들판에 앉아 있다. 그녀 주위의 풍성한 밀밭은 풍요로움을 나타내고, 그녀의 옷 에 그려져 있는 석류 무늬는 여성의 생식력과 다산을 상징한 다. 하트모양 쿠션에 그려진 비너스의 상징은 그녀의 여성스 러움과 아름다움을 상징한다. 성공을 상징하는 월계관 위의 마법의 힘을 상징하는 열두 개의 육각별로 이루어진 관은 열 두 개의 별자리를 상징하며, 성모 마리아의 상징이기도 하다. 그러므로 그녀가 매우 고귀 하고 우아한 여성임을 나타낸다. 그녀의 헐거운 옷과 진주목걸이는 그녀의 순수함과 평화로 움을 나타내며, 숲 사이로 보이는 조그만 폭포는 남성성의 상 징이며, 그 물로 인해 밀과 나무가 자라고 풍요롭게 됨으로 그로부터 보호받고 안정을 느끼는 여성을 나타낸다. 여왕 카 드는 앞의 여사제 카드와 더불어 대표적인 여성 캐릭터이지 만 뚜렷한 차이가 있다. 여사제가 무의식과 영혼, 순수함 등을 나타낸다면 여왕은 다산과 풍요로움을 상징한다.", oneCard: "Q: 회사를 그만 두고 창업을 하고 싶은데, 어떤 직종이 맞을지? A: 여성을 상대하는 화장품가게나 란제리(속옷)전문점을 권하고 싶다. 여의치 않으면 식당을 해보는 것도 가능하겠다. Q: 오늘 소개팅한 여자의 성격을 알고 싶다. A: 엄마와 같이 푸근하고, 당신의 고민을 충분히 들어 줄 수 있는 여자이다. 성격은 아주 좋다. Q: 호프집을 오픈하였는데 잘 될까? A: 당신이 워낙 서비스 정신이 좋아서 손님도 많을 것이고 금방 부자가 될 것이다." },
    "major-4": { description: "지엄하고 나이가 있어 보이는 사람이 황금빛 왕관을 쓰고 오른손엔 십자가 모양의 앙크를, 왼손엔 구슬을 들고 있다. 권 위의 단단함을 나타내는 암석으로 된 의자와 그 의자 앞에 있는 양의 머리는 독단적이면서 강인함을 상징한다. 하얀 수 염에서 경험이 풍부하다는 것을 알 수 있고, 쇠로 만든 부츠 는 어떤 길도 갈 수 있는 강인함을 보여준다. 뒤의 산은 황제 의 포부와 야망을 잘 나타내고 있으며, 그 밑으로 가늘게 흐 르는 강은 현실적이고 이성적인 면이 강하다는 것을 나타낸 다. 황제는 강한 가부장적 인물을 상징하며 규칙들을 만들고, 비록 자신은 지키지 않을지라도, 모든 사람들이 그 규칙들을 지키기를 바란다. 또한 직장 상사나 부모, 사업 동료와 같이 힘 있는 위치에 있는 사람일 수도 있다.", oneCard: "Q: 남자친구와 교제중인데 그 사람과 결혼이 가능할까? A: 남자는 보수적이고 카리스마 넘치고, 의지력도 강해 본인을 책임지려고 하니 결혼이 가능하다. Q: 취업을 준비 중인데 입사 시험에 합격이 될까? A: 합격된다. 어느 정도 자기 자리가 확실하게 있기 때문에 능력발휘를 할 수 있는 자리에 갈 것이고 적극적으로 업무처리도 잘 할 것이다. Q: 사업을 시작하려는데 잘 될까? A: 모든 준비는 다 되어 있으니 행동으로 옮기면 된다." },
    "major-5": { description: "두 개의 커다란 무채색 기둥사이에 권력의 상징인 왕관을 쓰고, 왼손에는 십자가를, 그의 발밑에는 열쇠가 있으며, 그의 앞에는 두 명의 사제가 무릎을 꿇고 있다. 이 카드는 황제카 드처럼 남성적이다. 황제카드가 현실적인 아버지의 카드라면 이 카드는 영혼적인 아버지를 상징한다. 교황이 들고 있는 십 자가 모양은 교황의 상징이며 권위를 나타내며 그 아래에 있 는 두 사람은 안내와 지시, 그리고 지혜와 이해를 나타낸다. 붉은 장미는 물질적 세계와 정열을, 흰 백합은 정신적 세계와 순수함을 나타낸다. 이 카드는 존경하는 사람을 상징하며, 개인적인 문제들을 털 어놓고 얘기할 수 있는 사람, 조언을 구하고 따르고자 하는 사람을 가리킨다.", oneCard: "Q: 현재 여자 친구가 없는데 어떻게 하면 생길까? A: 누군가가 다리를 놔줄 사람이 필요하다. 가령 윗사람이 중매를 한다든가, 결혼상담소의 도움을 받아야 한다. Q: 신경성으로 건강이 많이 안 좋아 졌는데 어떻게 하면 좋아질까? A: 명상을 통해서든, 기도를 통해서든, 종교를 가지면서 믿음으로써 마음에 위안을 찾는 것이 중요하다 Q: 친구한테 돈을 차용해 준 것이 있는데 받을 수 있을까? A: 받기 힘들다. 누군가가 중간에서 중재해 주는 사람이 있어야 한다. 공증을 해놨다던가 하면 괜찮다." },
    "major-6": { description: "그림의 맨 위 중앙에 밝은 해가 햇살을 비추고 있고, 그 밑 에는 날개달린 천사가 두 팔을 벌리고 자신의 영향력을 아래 를 향해 보고 있다. 이 천사는 라파엘을 나타낸다. 라파엘은 대단한 지혜의 소유자로 인간의 영혼을 주관하며 사랑을 나 타내는 천사이기도 하다. 그 아래 있는 벌거벗은 남녀는 아담과 이브를 나타낸다. 벌 거벗은 모습은 젊음, 처녀성, 순진함과 물질적인 탐욕으로 물 들기 전의 사랑을 나타낸다. 여자 뒤에 있는 사과나무는 선과 악의 나무이며, 그 나무를 감싸고 있는 뱀은 이브를 유혹한 사탄을 나타낸다. 남자 뒤에 있는 12개의 불꽃이 있는 나무는 생명나무 또는 12개의 별자 리를 나타낸다. 남자의 시선은 정면의 여자쪽을 바라보고 있고, 여인은 하늘 을 바라보고 있다. 혹자는 이것을 남성의 현실성과 여성의 감 성표현이라 하고, 또는 남자는 여자를 바라보지만 여자는 다 른 곳을 바라보고 있기 때문에 완벽한 사랑은 아니라는 뜻으 로 해석하기도 한다. 이 카드는 정신적 사랑과 육체적 관계 사이의 선택을 암시 할 수 있다.", oneCard: "Q: 오늘 미팅을 하는데 남자와 잘 될까? A: 잘된다. 당신에게 깊은 사랑을 느낄 수 있는 사람일 것이다. Q: 지금 여자 친구와 교제중인데 오래갈까? A: 장담할 수 없다. 남자는 그 여자를 사랑하고 있지만 여자는 꿈과 이상이 높다. 그래서 여자는 또 다른 유혹에 빠질 수 있는 선택에 기로에서 이 남자를 선택할지 저 남자를 선택할지 고민 할 수 있다. Q: 사업을 하려는데 동업을 해도 될까? A: 솔직하고 서로 믿음직한 관계이며 좋은 파트너가 될 수 있으니 동업해도 괜찮다." },
    "major-7": { description: "왕자와도 같은 남자가 오른손에는 완즈와 비슷한 것을 들고, 마차를 타고 의기양양하게 서 있다. 그는 마치 태양의 마차를 모는 아폴로와도 같다. 마차위에는 별모양이 그려진 하늘색 커튼이 쳐져 있는데, 별의 모양은 팔각, 육각, 오각별이다. 이 는 각각 우주의 에너지, 마법의 힘과 사랑을 뜻한다. 마차의 중앙에는 날개달린 원이 있는데 이것을 승리의 상징이나 영 혼의 바퀴라고 했으며, 이는 태양과 하늘의 지배자를 나타내 기도 한다. 마차를 이끄는 두 마리의 스핑크스는 하나는 검은색이고, 하 나는 흰색인데 이 두 마리는 얼굴을 모두 정면을 바라보고 있는데 이것은 음과 양, 이중성과 통일을 나타내는 것이도 하 다. 전차의 주인이 앞으로 나아가기 위해서는 두 마리의 스핑 크스 힘을 잘 분배해서 앞으로 나아가게 하는 것이다.", oneCard: "Q: 직장을 이직하고 싶은데 가능할까? A: 이직 가능하다. 역마 카드이므로 본인이 생각한 대로 밀고 나가면 된다. 그러나 너무 성급하게 추진하다 보면 실수 할 수 있으니 조심은 해야 한다. Q: 일도 하면서 공부도 하려는데 잘할 수 있을까? A: 강한 성향의 의지를 갖고 있기에 두 가지 일을 병행하면서 잘할 수 있다. Q: 친구와 싸워서 냉전 상태인데 어떻게 하면 좋을지? A: 전차는 앞만 보고 달리기에 내 의견만 내세울 수 있다. 내가 먼저 다가가 화해를 청해야 한다. 적극적인 자세로 그 친구의 의견도 듣고, 그 친구를 이해하려는 노력이 필요하다." },
    "major-8": { description: "한 아름다운 여성의 머리위에 영원의 상징이며, 성령의 신비 로움을 나타내는 뫼비우스의 띠가 있고, 그녀는 장미를 엮어 만든 화환과 허리띠를 두르고 있다. 이는 그녀의 내적인 힘이 무의식을 압도할 정도로 강력하다는 것을 상징한다. 순수와 초월을 나타내는 흰옷을 입은 이 여인은 아무렇지도 않은 듯 무시무시한 사자의 턱 아래와 콧등을 어루만지고 있 다. 사자는 그녀의 아름다움과 부드러움에 압도당한 듯, 꼬리 를 다리사이에 감추고 마치 애교 부리는 강아지처럼 그녀에 게 응석부리는 것처럼 보인다. 사자는 그녀에게 굴복했다. 마치 부드러움이 강함을 흡수하 듯 단지 굴복한 것뿐만 아니라 그녀에게 순종하고 있다.", oneCard: "Q: 좋아하는 사람이 있는데 그 사람과 잘 될까? A: 자신감과 용기를 가지고 부드러운 포용력으로 그 사람을 지혜롭게 잘 다스리면 시간은 걸리지만 잘 될 수 있다. Q: 대학에 진학하려는데 원하는데 들어갈 수 있을까? A: 포기하지 말고, 인내와 끈기를 갖고 꾸준히 노력하면 들어갈 수 있다. Q: 내 여자 친구는 어떤 여자인가? A: 남자를 잘 다스릴 수 있는 내공이 있는 여자이다. 내가 이기기 힘든 여자이니 잘 대처해야 한다." },
    "major-9": { description: "현자처럼 보이는 자가 왼손에는 지팡이를, 오른손에는 등불 을 높이 들고 눈 쌓인 산위에 고요히 서있다. 그가 들고 있는 등불 안에는 육각별이 밝게 빛나고 있다. 이 별은 마법의 힘, 더 큰 세상을 넘나드는 통제권을 상징하며, 등불은 지혜와 영 적인 빛을 상징한다. 이 카드는 내적으로의 여정에서 답을 찾 아주는 빛을 비춰주어 길을 찾게 도와준다는 의미이다. 왼손에 든 지팡이와 회색의 옷은 그가 내적인 힘을 간직한 현자라는 것을 의미한다. 세상과 떨어진 산속에 홀로 불을 밝 히고 있는 것은 그가 세속을 떠난 은둔자임을 보여준다. 또한 눈을 감고 모자를 쓴 채 고개를 숙이고 있는 것으로 우리는 그가 말로써가 아닌, 마음 또는 영적으로 인도하는 안 내자라는 것을 알 수 있으며, 바깥세상에는 관심이 없고 내면 적인 성찰에 관심이 있음을 나타낸다.", oneCard: "Q: 남자친구와 헤어졌는데 다시 연락이 올까? A: 마음의 문을 닫아 버렸다. 문자를 해도 대답이 없을 수 있으니 새로운 사람을 만나는 것이 좋다. Q: 소개팅을 하러갈 건데 어떤 사람이 나올지? A: 나이 차이가 많이 날수 있다. 아니면 보수적이고 철학적인 사고와 자기 자신을 잘 드러내지 않는 차분하면서 신중한 성향의 사람이 나올 수 있다. Q: 집을 사려는데 사도 좋을지? A: 문서, 서류, 계약 좋다. 윗사람의 조언을 잘 듣고 사도 좋다." },
    "major-10": { description: "카드 가운데 오렌지색 바퀴가 있다. 그 바퀴는 하늘 가운데 에 위치해 있다. 그 바퀴를 받치고 있는 어두운 오렌지색의 생물이 있는데, 사람의 몸에 자칼의 머리를 갖고 있다. 이는 죽은 영혼을 인도하는 이집트의 신이며, 부활의 상징인 “아누 비스”를 나타낸다. 바퀴의 위쪽에는 파란 스핑크스가 검을 손에 들고 있다. 스 핑크스는 균형과 현명함을 나타내며, 이는 이집트의 부활의 신인 “호루스”를 나타낸다. 카드의 네 모퉁이에 네 개의 생물 이 책을 읽고 있는데 이 책은 진실의 기록, 기억, 지혜 등을 의미한다. 바퀴의 안쪽에는 8개의 방향으로 선이 그어져 있고, 바깥쪽 에는 4방위의 방향에 북쪽으로부터 시계방향으로 TARO라는 글자가 쓰여져 있다. 안쪽에 8개의 바퀴살이 모여 바퀴의 중 심을 이루며, 바퀴는 행운과 불운이 무작위로 뒤섞인 운명의 회전을 나타낸다. 그래서 수레바퀴의 각각의 회전은 새로운 시작, 더 나은 쪽 으로의 전환, 그리고 어떤 종류의 완성을 가리키며, 중립적인 카드이기도 하다.", oneCard: "Q: 남자친구가 나를 어떻게 생각할까? A: 하늘에서 보내준 운명적인 사람이라고 생각한다. (운명적 만남, 너는 내 운명) Q: 지금 다니고 있는 직장에서 다른 쪽으로 이직을 해도 되는지? A: 매일같이 반복되는 일들이 힘들어 지쳐 있을 수 있다. 다른 쪽으로 이동하는 하나의 기회의 전환점이 온 것이다. Q: 부모님이 하시던 일을 제가 해도 좋을지? A: 해도 괜찮다. 집안의 가업을 이어받는 일이고 큰돈은 안 되더라도 현상유지 할 수 있으며, 성실하게 노력하는 모습을 보여 드리면 유산도 물려받을 수 있다." },
    "major-11": { description: "한 여인이 오른손엔 검을, 왼손엔 저울을 들고 두 개의 회색 기둥 가운데에 놓인 의자에 앉아 있다. 오른손엔 진취적인 권 위와 정의를 상징하는 검을 들고 있고, 정적인 내면을 상징하 는 왼손의 저울은 어느 쪽으로도 치우치지 않는 공명정대함 을 나타낸다. 그녀 뒤의 보라색 베일은 의식의 세계에서는 볼 수 없는 것 을 나타내며, 베일에 쳐진 기둥은 이러한 숨겨진 지식이나 능 력을 수호함을 상징한다. 그녀가 입고 있는 붉은 빛깔의 옷 위에 초록빛 옷을 걸쳐 입은 것은 그녀가 열정이 넘치고, 용기 있는 결단의 소유자이 지만 풍요로운 마음으로 잘 다스리고 긍정적인 마인드로 잘 조화시켜 나갈 수 있는 능력의 소유자임을 보여준다. 정의카드는 균형과 조화를 나타내며, 어느 한쪽으로도 치우 치지 않는 공명정대함을 뜻한다.", oneCard: "Q: 거래처로부터 받을 돈이 있는데 어떻게 하면 받을 수 있는지? A: 본인이 돈을 받으려면 판단에 망설이지 말고 법적인 절차를 밟아야 받을 수 있다. 서로의 감정으로 부딪치지 말고 냉정하게 심사숙고해서 법으로 해결하는 것이 좋다. Q: 연애가 왜 잘 안될까? A: 너무 옮고 그름을 따질 수 있고 원칙을 중요시 생각하기 때문에 상대방의 잘못에도 그냥 넘어가지 않으려 한다. 그래서 재미없는 사람으로 비쳐줄 수 있다. Q: 짝퉁 가방을 유통하는 사람인데 법에 걸리지 않고 잘 할 수 있을까? A: 불법적인 것은 법의 공정한 판결을 받을 수 있다. 법에 어긋나거나 상식에서 벗어나는 일은 하지 말고 절차를 밟아서 정식적인 상품을 유통하는 것이 좋다." },
    "major-12": { description: "한 남자가 T자형 십자가에 거꾸로 매달려 있다. 그의 한쪽 발과 양 손은 묶여져 있다. 양다리는 십자가형으로 오른쪽 다 리가 묶여 있고, 왼쪽 다리는 그 뒤쪽에 가리워져 있으며, 양 손은 삼각형 모양으로 묶여 있다. 이 모양은 십자가 밑에 삼 각형이 있는 황금새벽회의 상징을 나타낸 것이며 매달린 남 자의 모양은 목성의 심볼을 거꾸로 해 놓은 모양과 같다. 목성의 뜻은 행운을 나타내지만 거꾸로 되니 목성의 심볼을 타로카드의 역방향식으로 해석하면 가치를 받아들이지 못하 고 정체되어 있는 상태를 나타낸다. 그의 머리에 빛나고 있는 후광은 지금은 묶여있어 현실적으로 아무것도 할 수 없는 정 체된 상태임에도 그의 무의식은 자유롭다는 것을 나타내며, 그의 신성함은 가릴 수 없음을 상징한다. 중요한 것은 이 십자가의 나무가 죽은 나무가 아니라 아직 살아있는 나무인데 앞으로 이러한 정체기에서 벗어날 것임을 시사해 주는 것이기도 하다. 그는 죽어가고 있는 것이 아니라 적당한 때를 기다리고 있는 것이다. 지금은 아직 때가 아님을 알고 조용히 마음을 가다듬으면서 그 시간이 오기를 잠잠히 기다리고 있는 것이다.", oneCard: "Q: 사업을 해도 괜찮을까? A: 사업 수완은 없다. 그러나 전문적이고 장인 정신이 뛰어 나면서 자기 일에 충실한 사람이니 누가 뭐라 하든 흔들림 없이 자기만의 일을 하는 경우는 권장할 만하다. Q: 지금 살던 곳에서 이사를 하려는데 가능한가? A: 매달려 있는 상황이라 집을 내놔도 당장 나가지는 않는다. 조금 때를 기다려야 할 상황이다. Q: 친구관계로 만난 지 3년 됐는데 헤어질까? A: 헤어지는 것은 어렵다. 조금만 더 희생하면 희생에 대한 보답은 받을 수 있으니 그때까지 힘들더라도 기다려라." },
    "major-13": { description: "검은 갑옷을 입은 사신이 검은 고삐를 두른 백마를 타고 있 다. 백마를 검은 고삐로 두르고 있다는 것은 자연의 순수한 힘을 죽음의 고삐로 매어 놓았다는 것을 뜻한다. 이것은 변화 의 상징이기도 하다. 그가 들고 있는 검은 바탕에 백장미가 그려진 깃발은 고통을 통과하고 난 후 얻어지는 부활을 뜻한 다. 검은 갑옷의 사신이 백마를 타고 있는 것도 같은 뜻을 가 진다. 뒤에 보이는 말라가는 강은 생명이 다해감을 뜻하지만, 저 멀리 두 개의 기둥 사이로 떠오르는 태양은 생명의 통로 로 서서히 들어오는 부활을 상징한다. 더 이상 과거와 같지 않은 새로운 세상이 온다는 뜻이기도 하다. 사신의 발밑에 보이는 네 사람 중 왕은 죽어있고 처녀는 기 진맥진해 있으며, 어린 아이는 아무것도 모르는 듯 꽃을 들고 앉아 있다. 교황인 듯 보이는 자만이 사신을 맞이하고 있는데 이는 현실적인 자아(왕)를 버리고 신의 뜻을 받아들이는 자(교 황)만이 부활을 자신의 것으로 완벽히 받아들일 수 있다는 것 을 의미한다. 이 카드는 육체적인 죽음을 나타내는 무서운 의 미만이 아니라, 모든 것을 쓸어버린 후에 받아들이는 재탄생 을 의미한다.", oneCard: "Q: 남편과 이혼이 잘 될까? A: 이혼은 할 수 있다. 그러나 남편이 폭력을 행할 수도 있고 죽을 만큼 힘든 고생이 따른다. 곤조로 안 해주고 애먹일 수 있으니 마음의 각오를 단단히 해야 할 것이다. Q: 시험관 아이를 가지려고 하는데 잘 될까? A: 한 번에 될 것 같지는 않으니, 너무 조급히 서둘지 마라. 분명 희망은 있다. Q: 부모님이 얼마 전에 암수술을 받으셨는데 회복이 가능한지? A: 실력 있는 의사를 만났을 것이고 수술 경과도 점점 호전 되어 가고 있으니 걱정하지 않아도 된다." },
    "major-14": { description: "한 천사가 한쪽 발은 물에 담그고, 다른 한쪽 발은 땅에 딛 고 서 있다. 그는 눈을 감은 채로 양손에 컵을 들고 왼쪽 컵 에서 오른쪽 컵으로 물을 붓고 있다. 이것은 이성과 감성을 조합하여 안정됨을 조성한다는 뜻을 담고 있다. 이마 중앙의 원 모양은 태양의 상징이고, 그의 가슴에 보이는 사각형 안의 삼각형 모양은 3각과 4각을 합한 수인 7과 연결된다. 신의 숫 자이자 최상급의 의미인 3인 조화와 자연계의 숫자이자 최초 의 합성수인 4가 합하여 나온 7이란 수는 완벽, 자기절제를 나타내는 수이다. 그리고 이 도형은 신비주의와도 연관이 깊 다. 그의 양쪽 발은 의식과 무의식의 연결을 나타낸다. 그의 옆에 피어 있는 붓꽃은 신으로부터의 메시지와 내적인 안내 를 뜻하며, 그가 한쪽 발을 담그고 있는 물은 연결을 뜻한다. 러므로 이것은 신이 인간에게 주는 메시지와 안내를 상징하 며 타협, 평정, 정서적 성숙을 의미한다. 절제 카드가 나타내 는 중요한 메시지는 변화에 흔들리지 않고 감정을 잘 다스리 며 조화와 균형을 맞추는 것, 즉 중용과 인내를 의미한다.", oneCard: "Q: 현재 학생인데 저녁에 아르바이트를 하려는데 어떤 일이 가능한가? A: 술집, 바텐더, 바리스타, 음식점등 물과 관련된 일이면 더욱 좋다. Q: 소송건이 있는데 잘 해결 될 수 있는지? A: 해결은 되겠지만 시간은 오래 걸릴 수 있다. 오히려 서로 합의를 보는 것이 이익일 수 있다. Q: 직장에서 해외발령을 원하고 있는데, 이번 4월달 인사이동에 가능한가? A: 해외발령 인사이동은 언제라도 가능하다." },
    "major-15": { description: "염소 모양의 머리를 한 짐승의 등엔 박쥐의 날개가 달렸고, 얼굴은 염소의 모습을 하고 있고, 다른 부위는 인간의 모습을 하고 있다. 오른쪽에 있는 여성의 꼬리는 포도송이가 달려 있 는데, 이것은 종교적인 신성한 교류가 잘못 사용된다는 뜻이 다. 왼쪽에 있는 남성의 꼬리는 횃불로 되어 있는데, 이것은 영적인 열의가 잘못 사용된다는 뜻이다. 또한 꼬리는 인간의 이성보다는 숨어 있는 동물적인 욕구가 더 강하게 작용한다 는 상징이기도 하다. 그들은 악마에게 잡혀 있는 것이 아니라, 그들의 의지로 그곳에 있는 것이다. 여기에는 보답 없는 사랑 의 고통, 채워지지 않는 욕망, 절실한 필요, 외로움, 최악의 집 착 등 치명적으로 끌리는 것들이 있는 카드이다.", oneCard: "Q: 지금 교제 중인 남자친구가 혹시 다른 여자가 있는지? A: 여자가 있을 수 있다. 나는 놔주지 않으면서 다른 여자와 성적의 쾌락에 빠져 있을 수 있으니 지금 남자친구를 주위 깊게 지켜봐야 할 것이다. Q: 친구에게 돈을 빌려주었는데 받을 수 있는지? A: 상대방의 금전 상황이 별로 나아지지 않을 가능성이 높아 금방 받기는 어려울 것이다. Q: 공무원 시험 준비를 하고 있는데, 금년에 합격이 가능한지? A: 악마 카드가 나왔으므로, 합격을 기대하기는 곤란하다." },
    "major-16": { description: "험한 산꼭대기의 탑에 하늘에서 번개가 내리치고 탑에는 불 길이 치솟고 있으며, 탑에 있던 왕족같이 보이는 두 사람이 탑 아래로 떨어지고 있다. 이 탑은 인간들이 신의 영역에 가 까이 가려는 야망의 산물이었던 바벨탑이며, 바벨탑이 신의 진노로 무너져 내린 것과 같이 신의 노여움으로 재앙이 닥친 것이다. 떨어지는 금빛 왕관은 왕권의 상실, 즉 신이 인간의 그릇된 야망에 대한 노여움으로 인간의 권위를 내리쳐 버린 것이다. 하늘의 먹구름과 활활 타오르는 불 그리고 번개 등은 하늘의 노여움을 상징하는 재앙과 파괴이다. 이 카드는 22장의 메이저 카드 중 가장 부정적인 뜻이 강한 카드이며, 벗어날 방법이 없는 두려움의 상징이기도 하다. 탑 카드는 보통 마른하늘에 벼락이 치듯이 커다란 변화가 일어 날 것이라는 것을 가리키며, 원조나 도움이 갑자기 끊길 수도 있다. 바깥세상과 관련하여 탑은 우리의 계획이나 일을 망치 는 뜻밖의 사건을 상징한다.", oneCard: "Q: 사업을 하고 있는데 금전적으로 힘든 상황이다. 계속 사업을 해야 하나? A: 접어야 한다. 사업을 계속 진행 할 경우 재정적인 손실이 크게 올수 있다. 파산되기 전에 빨리 정리하여 안정을 찾는 것이 좋다. Q: 남자친구와 헤어지지 않고 잘 만날 수 있는지? A: 갑작스런 사건이나 상황에 의해서 남자친구와 헤어질 수 있다. (입대영장이나 외국으로 유학) Q: 올해 직장에서 승진 할수 있는지? A: 별 문제가 없는데도 승진하기 힘들다. 주위 사람들과도 대립이나 말다툼을 할 수도 있으니 조심해야 할 것이다." },
    "major-17": { description: "별 카드는 14. 절제카드와 비슷한 점이 많다. 우선 절제 카 드는 양손에 든 컵으로 물을 주고받는 모습인데 반해, 별 카 드는 한쪽 컵의 물은 샘물에 붓고 있고, 다른 컵의 물은 땅 위에 붓고 있다. 샘물은 무의식과 잠재의식 그리고 정신세계 를 의미하고, 땅은 그와 대립되는 세계인 의식과 물질적이고 현실적인 세계를 의미한다. 따라서 절제 카드가 개인과 개인 의 절제와 조합이라면, 별 카드는 세상과 나의 결합과 조합을 의미한다. 한편 벌거벗은 여인은 자유로운 인간을 의미하며, 연꽃은 자연의 재생력을 상징하여 최악의 순간에서도 피어나 는 희망을 상징한다. 별 카드는 희망과 행운을 암시하며 긍 정적인 결과를 가리킨다. 이 카드는 행운과 재생의 카드이기 도 하다.", oneCard: "Q: 남편이 몇 달째 휴직 상태인데 재취업이 가능한지? A: 이 카드는 미래에 대한 강한 희망이 있음을 의미하므로 재취업이 가능하다. Q: 현재 임신 계획이 있는데 가능한지? A: 환한 별과 샘물이 있으므로 임신이 가능하다. Q: 딸아이가 연예인이 되기를 희망하는데 가능한지? A: 이 카드는 직업이나 적성에서 예술 관련된 일과 인연이 많으므로 가능하다. 열심히 뒷바라지 해주면 된다." },
    "major-18": { description: "사람의 옆모습을 한 달이 밤하늘에 떠 있는데 마치 잠을 자 는 듯 눈을 감고 있다. 달빛 아래에서는 사물의 형체를 뚜렷 하게 볼 수 없듯이, 이 카드의 달은 실체가 무엇인지 정확하 게 파악할 수 없는 불확실하고 불안정한 상황을 의미한다. 또 다른 의미로 달은 인간의 내면과 여성의 생리주기에 영향을 미치고, 인간을 정화한다는 의미가 있다. 길옆에 있는 개는 사 람의 절친한 친구이자 동반자라는 뜻을 가지고 있으며, 늑대 는 야성의 본능이 있는 길들여지지 않은 위험성이 숨겨져 있 다. 물은 무의식과 미스터리를 상징하고, 갑각류가 물 위로 올 라오고 있는 것은 숨겨진 무엇이 막 드러나려는 상황을 상징 한다. 이 카드에는 길이 험할 뿐만 아니라 목적지가 보이지 않는 다. 두 개의 탑은 통로이자 관문이다. 여기서 보이는 두 탑 까지의 거리는 결코 멀지는 않지만, 막막해 보이고 바람까지 불고 있다. 이는 가도 가도 목적지가 보이지 않는 막막한 상 태와 두려움의 상태를 나타낸다. 달 카드가 나오면 변화, 변 동, 불확실성, 삼각관계, 속임수를 의미하는 카드이다.", oneCard: "Q: 가게 계약을 하려는데 잘 될지? A: 계약이 힘들다. 내가 계약하려고 하면 제 3자가 방해하거나, 상대 또한 신뢰하기 힘든 사람이라 거래가 무산된다. Q: 1년 동안 만난 남자친구와 계속 갈수 있는지? A: 계속 가기 힘들다. 상대에 대한 내 마음이 불분명한 상태이고, 서로가 애매모호한 관계로 갈 수 있으 니 서로의 관계를 확실하게 하는 것이 좋다. (예를 들어 결혼 할 건지) Q: 직장 동료가 있는데 믿을 만한 사람인지? A: 간사하고 교활하며 남을 잘 속일 수 있는 성격의 소유자 일수 있으니 조심해야 할 것이다. 나를 중상모략 하거나 안 좋은 소문들을 퍼트릴 수도 있다." },
    "major-19": { description: "머리에 화관과 붉은 깃털을 쓴 어린아이가 왼손에 붉은 깃 발을 들고 백마를 타고 있다. 어린아이와 백마는 순수함과 생 명력을 나타내며, 붉은 깃발과 아이 머리의 붉은 깃털은 행동 력과 자유로움을 상징한다. 아이는 태양의 보호 아래 안전하고 행복한 상태이며, 더 없 는 환희의 상태라는 것을 알 수 있다. 아이는 표정을 비롯하 여 모든 것이 안정되어 있으며, 행복과 만족감에 넘친다. 태양은 모든 창조의 근원이자 생명의 궁극적인 상징이며, 어 떤 상황이나 사건, 목표를 나타낼 때 이 카드는 확실한 성공 을 의미하며, 22장의 메이저 카드 중 가장 긍정적이며, 행복 하고 만족스런 카드이다.", oneCard: "Q: 내 남자 친구에게 다른 여자가 있나? A: 있을 수 있다. 밝고 명랑한 성향에 사교적인 사람이라서 주변의 인기를 한 몸에 받기 때문이다. Q: 사업을 확장 하려는데 해도 되는지? A: 해도 괜찮다. 사업에 강력한 힘을 발휘하며 일이 잘 진행될 것이다. 물질적인 풍요와 기대 이상으로 큰 수익이 나며, 주변의 도움도 받을 수 있으며 본인의 능력도 충분히 발휘 할 수 있다. Q: 임신을 준비하고 있는데 잘 될지? A: 잘 된다. 건강하고 잘 생긴 사내 아이를 가질 수 있다. 단, 유산되지 않게 각별히 몸 조심은 해야 한다. (말에서 떨어질 수 있는 불안한 상황)" },
    "major-20": { description: "하늘엔 천사가 구름에 둘러싸여 아래를 향해 나팔을 불고 있으며, 땅에는 죽었던 사람들이 무덤에서 일어나 하늘에 있 는 천사를 향해 두 팔을 벌리고 있다. 죽음과 부활의 천사이 며, 최후의 심판 때 나팔을 불 임무를 맡은 천사이기도 하다. 절제를 미덕으로 삼으며 물의 힘을 지니고 백합이 상징인 천 사이다. 다른 메이저 카드들과 마찬가지로 이 카드 또한 기독 교적인 상징들이 들어 있지만, 그 중에서도 그 특성이 특히 강한 카드라고 볼 수 있다. 심판이라고 하는 것은 그 결과가 어떻게 나온다 하더라도 그건 자신이 해온 노력이나 결과에 대한 평가이다. 심판 카드 가 배열에 나올 때, 우리는 올바른 길 위에 있고 모든 장애물 들을 이미 빠져 나왔다는 것을 안다. 주된 메시지는 우리가 기꺼이 과거를 청산하고, 그 동안 배운 교훈에 비추어 새롭게 시작하고, 그 교훈들이 가리키는 결정들을 받아들이는 것이다.", oneCard: "Q: 집 나간 지 한 달된 남편이 다시 가정으로 돌아올까? A: 돌아온다. 내가 좀 더 기다리면 소식이 올수 있고, 다시 남편은 가정으로 돌아올 수 있다. Q: 작년에 승진시험에 떨어졌었는데 올해는 될 수 있는지? A: 올해는 승진이 가능하다. 새롭게 부활하는 카드이므로 과거에는 승진시험에 떨어졌더라도 올해는 기대해 볼만 하다. Q: 여자친구와 헤어진지 3개월 되었는데 다시 만날 수 있는지? A: 다시 만날 수 있다. 헤어졌던 연인들은 다시 연락이 오거나 만날 수 있고, 재회의 기쁨과 사랑이 되살아난다." },
    "major-21": { description: "나체의 모습을 한 여인이 보라색 천을 몸에 휘감고 양 손엔 완즈를 들고서 춤을 추고 있다. 월계수 잎이 타원형으로 그녀 를 둘러싸고 있으며, 그 위쪽과 아래쪽에는 붉은 색으로 된 장식이 뫼비우스 띠 모양을 하고 있다. 월계수는 승리와 순환을 나타내며, 붉은 뫼비우스 띠 모양 또한 무한한 열정과 영원한 삶을 상징한다. 그녀가 양 손에 들고 있는 두 개의 완즈는 그녀가 순수하지만 강력한 힘의 소유자라는 것을 보여 준다. 모든 여정이 끝났고, 모든 단계의 완성을 나타내며 다시 처음으로 돌아가는 새로운 시작을 의 미한다. 세계 카드는 모든 인생길을 마무리하는 모습일 수 있 고 좋은 결과로 끝나는 것을 의미하며 또한 모든 것을 끝내 고 툭툭 털고 일어나 여행을 떠나는 모습으로도 볼 수 있다.", oneCard: "Q: 연애를 10년 동안 해왔는데 그 사람과 결혼이 가능한가? A: 가능하다. 인생(연애)의 마지막 종착점에 와 있고 연애를 끝마치고, 새로운 인생의 시작하는 결혼의 시기가 다가온 것이다. Q: 해외로 이민가고 싶은데 가면 잘 적응할 수 있는지? A: 잘 적응할 수 있다. 새로운 세계로 나아가 삶을 새롭게 시작 할 수 있으며, 성공과 행복이 기다리고 있다. Q: 연예인이 되고 싶은데 될 수 있는지? A: 될 수 있다. 충분한 끼와 재능이 있고 아름다운 자신의 모습을 충분히 표현할 수 있으니, 연예인이 될 자질이 충분히 있다." },
    "wands-ace": { description: "구름 속에서 뻗어 나온 손이, 활기찬 창조력과 남성다운 에 너지의 상징인 나뭇잎이 달려 있는 막대기를 당당히 들고 있 다. 그 막대기에서 떨어져 나온 8개의 나뭇잎은 새로운 시작 을 나타낸다. 오른편 아래에 보이는 세 그루의 나무는 창조력과 성장을 상징하며, 왼편에 보이는 작은 언덕 위의 성은 미래의 성공을 나타낸다. 지팡이는 일이기도 하지만 생각, 창조적인 능력 또한 의미하 고 있다. 일에 대한 새로운 아이디어와 창조력이 충만해 있는 상태를 가장 잘 표현하고 있다. 예를 들면 막 창업을 시작한 혹은 창업을 준비 중인 상태이다. 모든 에이스 카드들과 마찬가지로 구름 속에서 나온 손이 지팡이 하나를 내밀고 있다. 이것은 신의 섭리, 신성의 개입, 도움을 주는 손길을 암시한다.", oneCard: "" },
    "wands-2": { description: "한 남자가 붉은 모자를 쓰고, 갈색 계통의 옷을 입고, 오른손 에는 지구본을 왼손에는 지팡이를 들고 서 있다. 지구본은 그 가 이미 성취한 것을 나타낸다. 그는 홀로 성벽 난간에 서서 해안가를 내려다보고 있다. 갈색 옷을 입고 있는 것과 성 밖이 아닌 성안에 있는 것으 로 그가 안전을 중요시하고 있음을 알 수 있다. 붉은 모자는 열정적이지만 그로 인한 스트레스도 많음을 알 수 있다. 두 개의 지팡이. 두 가지 일을 한다? 물론 그건 아니다. 에이 스가 두 개 모여서 더 큰일을 꿈꾸는 상황이다. 두 배로 큰일 을 하려는 사람. 두 개의 지팡이는 사업에 있어서 용기 있는 사람, 원대한 꿈을 꾸는 사람을 의미한다. 어떤 상황에서든지 이 카드의 메시지는 문자적으로든 어떤 상황 안에서든 그대로 머무는 것과 이동하는 것 사이에 선택 이 있다는 것이다.", oneCard: "" },
    "wands-3": { description: "귀족처럼 보이는 한 남자가 언덕 위에 서서 저 아래 바다를 지나가는 배들을 바라보고 있다. 배들은 그의 명령을 따라 어 디론가 이동 중인 것으로 보인다. 전투를 위한 항해가 아닌, 무언가 좋은 일을 맞이하러 가는 듯 보인다. 일에서 어느 정 도의 성공을 의미한다. 그러나 완벽한 완성의 느낌은 아니다. 창조적인 능력이 어느 정도 성과를 보여서 사업적 통찰력을 갖게 되고, 사업에서의 기회와 결과를 얻어낼 수 있다. 이 카드에서는 3이라는 숫자가 주를 이룬다. 막대기도 3개, 배도 3척, 막대기에 난 나뭇잎도 각각 3개씩이다. 3이라는 숫 자는 최초의 완성을 나타내는 숫자이며, 구조의 설립과 성장 과 성취의 숫자이다. 이 카드는 고향을 그리워하는 마음이나 멀리 있는 어떤 사람과 함께 있지만, 당분간 멀리 떨어져 있 어야 한다.", oneCard: "" },
    "wands-4": { description: "나란히 세워진 4개의 거대한 막대기 위에 커다란 화환이 걸 려 있다. 그 화환은 여러 종류의 꽃들과 풍요를 나타내는 포 도송이들로 이루어져 있다. 화환 아래는 두 명의 여인이 3개 의 꽃다발을 머리 위로 높이 올려 누군가를 환영하고 있는 듯하다. 마치 갈색 옷의 여인은 안전을 수호하는 여신이고, 푸 른 옷의 여인은 평화를 수호하는 여신인 듯하다. 다른 사람들 은 아마도 흥겹게 파티를 즐기는 듯하다. 숫자 4는 안정과 번 영을 상징하며, 네 개의 지팡이는 노동으로 수확한 뒤의 축하 파티라고 생각하면 된다. 풍년을 축하하는 추석이나, 추수감사 절을 카드로 표현한다면 이 카드가 될 수 있다. 지팡이 4번 카드를 ‘행복한 귀향’이라 부르기도 하고 어떤 재결합을 준비 하고 있을 때 나타나는 경우가 많은데, 동반자나 가족과의 재 결합일 수도 있고, 또는 휴가나 휴식을 위한 때를 나타낼수도 있다.", oneCard: "" },
    "wands-5": { description: "5명의 젊은이들이 각기 다른 색의 옷을 입고 긴 막대기를 휘두르며 서로 싸우고 있는 듯하다. 하지만 그들의 표정은 그 리 심각하지도 않으며, 아무도 다치지 않았다. 그들의 막대기 는 서로 부딪치고는 있지만, 상대방에게 직접 일격을 가하고 있지는 않다. 사공이 많으면 산으로 간다. 자신의 주장을 내세워 목표 자 체는 안중에도 없다. 분열과 투쟁, 그리고 장애물만 남아 있을 뿐이다. 이 카드는 진행 중인 논쟁이나 직장에서의 이해관계, 다른 사람들의 다툼 또는 어떤 종류의 중상모략에 연루되거나 둘 러싸일 때 나오는 경향이 있다. 분열과 다툼을 나타내며, 불화 와 옹졸함 그리고 사소한 말썽을 가리킨다.", oneCard: "" },
    "wands-6": { description: "한 남자가 승리의 상징인 월계관을 쓰고 갈색 망토를 걸친 채 연두색 천을 두른 백마를 타고 있다. 그의 손은 월계관이 달린 막대기를 들고 있다. 그는 여러 명의 추종자들과 함께 노력의 댓가로 이룬 승리를 만끽하고 있다. 그가 입고 있는 갈색 망토는 안전과 힘을 나타내며, 백마가 두르고 있는 연두 색 천은 풍요와 안정을 뜻한다. 그는 승리를 통해 풍요와 안 정을 가지고 돌아온 영웅이다. 이 카드는 일에서의 성공을 나타내며, 혼자 이루어낸 성공이 아니라 우두머리로서 일을 주도했을지는 몰라도 주변에 도와 주는 사람들과 함께 한 성공이다. 찬사와 인정을 받는 성취를 가리키며 결혼식, 시험이나 운전면허 시험의 합격, 혹은 힘든 상황에서의 성공적인 결과 등을 보여 주기도 한다. 그리고 당 신이 문제를 겪거나 시간이 걸림에도 불구하고 승리한다는 것을 의미하며 고된 노동에 대한 지지나 보상을 의미하기도 한다.", oneCard: "" },
    "wands-7": { description: "젊은이가 험한 벼랑 위에서 막대기 하나로 다른 6개의 막대 기들의 공격에 용감하게 맞서고 있다. 6개의 막대기가 그를 향해 공격을 하고 있지만, 그의 용맹함 때문에 감히 그에게 다가서지 못하고 있다. 그는 민첩하고, 뛰어난 잠재력을 갖고 있다. 아직은 어리고 미숙하지만 장차 훌륭한 남성이 될 충분 한 자질을 갖고 있다. 자신의 신념을 굽히지 않기에 다른 사 람들과의 마찰이 생긴다. 자신의 신념을 굽히지 않는다는 것 은 좋게는 자신의 뜻을 관철시키는 강한 정신력을 의미하지 만 일면 고집스럽고, 주변과의 융통성이 없다는 것을 나타낼 수도 있다. 힘들고 어려운 상황이라도 용기를 내어 맞선다면 성공할 수 있음을 의미하며, 자신감 있게 대응 하라는 뜻의 카드이다.", oneCard: "" },
    "wands-8": { description: "8개의 막대기가 그들의 목적지인 땅을 향해 하늘을 날고 있 다. 그 모습이 무척 날렵하고도 안정되어 보인다. 이 카드는 에이스를 제외한 마이너 카드 중에서 유일하게 인물이 등장 하지 않는 카드이다. 또한 모든 지팡이가 활성화되어 일의 빠 른 진행과 공부의 효율을 의미한다. 어떤 일을 서둘러 진행해 야 하거나 발 빠른 대응이 필요할 때 자주 등장한다. 이 카드는 침체기를 벗어나 새롭게 시작하고 행동을 할 때 임을 암시하며, 어디론가 이동하거나 여행을 떠나는 것도 의 미한다. 스트레스에서 벗어나는 시기를 나타낼 수도 있다.", oneCard: "" },
    "wands-9": { description: "힘겨운 싸움을 치룬 한 남자가 머리에 상처를 입은 채 막대 기에 기대어 서 있다. 그의 뒤에 서 있는 8개의 막대기는 아 직 건재하다. 이미 힘겨운 일을 잘 겪어 왔고, 잘 지켜 왔다. 하지만 아직은 마음을 놓을 때가 아니다. 이 카드는 자신이 만들어 놓거나 이룩해 놓은 일들을 다른 사람이 망칠까봐 두 려워하고 있다. 8번 카드가 올려놓은 속도를 늦추어야 할 필 요성도 있다. 이 카드는 무엇을 얻고 지키기 위해서는 그 만 큼의 노력이 필요하다는 것을 나타낸다. 부상당한 채 지팡이 를 잡고 싸움에 지쳐 서 있는 인물이 경계하고 있는 모습은 방심하지 말고 계속 경계하면서 자기를 방어할 필요가 있다 는 것을 암시한다.", oneCard: "" },
    "wands-10": { description: "혼자 나르기엔 무거워 보이는 10개의 막대기를 나르고 있는 남자의 뒷모습이 매우 힘겨워 보인다. 앞에서는 목적지인 마 을이 보이지만 그에게는 멀게만 느껴진다. 비록 얼굴은 보이 지 않지만 무거운 짐에 얼굴을 묻고 있는 모습에서 육체적, 정신적 부담감에 매우 힘겨워하는 것이 느껴진다. 책임감에 너무 많은 일들을 한꺼번에 처리하려고 하기 때문에 힘들다. 자신의 능력의 한계를 알고, 그것에 맞춰서 일을 조절해야 한 다. 이 카드는 목적지에 다다르면 누르고 있던 무거운 짐이 곧 제거되고, 부담감도 해소될 수 있음을 나타낸다. 하지만 그러 기 위해 힘들다는 생각보다는 긍정적인 생각을 가져야 할 것 이다. 그리고 이 카드는 업무 압박감이나 특정한 상황 또는 책임이 너무 벅차서 감당하기 힘들 때 나타나는 경우가 많다.", oneCard: "" },
    "wands-page": { description: "지팡이 소년은 믿음직스럽고 꾸준하지만 변화에는 다소 약 한 면이 있다. 하지만 여기에 완즈라는 불의 성향이 섞여, 열 정을 갖고 배우며 진보적이고 미래지향적인 성향을 나타낸다. 놀기를 좋아할뿐더러, 사람을 사귀고 만나는 것도 좋아하여 여러 가지 정보에 빠르다. 하지만 불과 같이 확 타올랐다 사 라지는 성향이 있어, 쉽게 흥미를 갖고 덤비지만 또한 쉽게 질려 버리는 경향도 있다. 직관적이고 본능적이어서 눈치가 빠르고 섹시한 구석도 있다. 따라서 이 사람은 인기 있고 자 발적이고 스스로 시작하는 사람이며, 다른 사람들에게 동기를 부여하고 영향을 미치는 방법을 알고 있는 사람이며, 모험적 이고 외국인이거나 외국과 관련된 사람일 수 있다.", oneCard: "" },
    "wands-knight": { description: "막 박차고 달리려는 말 위에 한 젊은이가 앉아 있다. 그는 불꽃 모양의 주황색 깃털이 머리에 달린 갑옷을 입고 오른손 에는 막대기를 쥐고 있다. 갑옷 위에 불의 상징인 샐러맨더가 그려진 노란 옷을 입고 있다. 여기서 불꽃 모양의 주황색 깃 털은 행동이 따르는 용기를 상징한다. 이 사람은 정력적이고, 모험을 즐기며, 솔직하고, 충동적인 사람이다. 그는 소년과 비 슷하게 전형적으로 불같은 특성을 보인다. 기사는 고집이 세 거나 비현실적인 특성처럼 예측할 수 없는 불같은 성격을 가 지고 있다. 이런 사람과 함께 있으면 매우 활기차고 흥분될 수 있다. 그러나 그는 끈기와 결의의 수단이 부족하여 일단 시작한 일을 끝맺지 못할 수도 있다. 연애 면에서는 기사는 상대방을 매혹 시킬 수 있지만, 과연 그 사람이 변함없이 곁 에 머물러 줄 것인지는 의문이다. 지팡이의 기사는 소년 카드 와 마찬가지로 해외와 관계가 있을 수 있다.", oneCard: "" },
    "wands-queen": { description: "불의 영향을 받아 그녀는 창조적이고, 생기발랄하며, 자기 확 신이 분명하고 열정적이다. 이와 더불어 물의 영향으로 주위 의 모든 것들을 배려하고 챙길 줄 아는 한마디로 여러 방면 에 다재다능한 팔방미인인 것이다. 왕좌의 등받이 부분은 킹 완즈의 것과 같이 위로 곧게 뻗어 있으며, 팔걸이 부분 밑에 는 불, 그리고 힘과 의지의 상징인 사자가 새겨져 있다. 여기 서 우리는 그녀가 여성적이기보다는 남성적이며, 왕과 같은 파워와 용기의 소유자라는 것을 알 수 있다. 성격 또한 강직 하고 명료함을 알 수 있다. 여왕 앞에 검은 고양이가 당당하게 앉아 있는데 이것은 그 녀의 비밀스러운 사자를 상징하거나, 또는 여왕의 부정적인 면을 나타내기도 한다. 퀸 카드들 중 유일하게 다리를 벌리고 앉아 있는 카드인데, 이것은 ‘욕망’이란 단어와 연관시켜 성적 인 것으로 해석된다. 또한 4개의 퀸 카드들 중에서 정면으로 앉아 있는 단 하나의 카드이기도 한데, 이것은 그녀의 당당하 고도 적극적인 면을 상징한다.", oneCard: "" },
    "wands-king": { description: "불은 직관을 나타내며, 활동적인 양의 기운이며 불은 무사인 데 무사가 무조건 돌격하는 것이 아니라, 지적으로 생각하고 작전을 짜면서 공격하고 방어하는 상황이라고 생각하면 되는 것이다. 창조의 상징인 지팡이를 손에 든 왕이 샐러맨더와 사 자가 그려진 황금빛 왕좌에 앉아 있다. 그는 샐러맨더가 수놓 아진 황금빛 망토를 입고, 불의 형상을 한 왕관을 쓰고 사자 머리 모양의 황금 목걸이를 하고 있다. 이 꼬리를 물고 있는 샐러맨더는 영원을 상징하고, 장애물을 뛰어넘어 전진하는 그 의 파워를 나타낸다. 그의 발밑에는 샐러맨더 한 마리가 왕을 향해 앉아 있다. 불을 상징하는 이 샐러맨더는 왕의 충성스런 부하이자 지지자이다. 이 사람은 빈틈없고, 정력적이고, 유능 하며, 위엄이 있는 사람이다.", oneCard: "" },
    "cups-ace": { description: "구름 속에서 나타난 손에는 황금빛 컵이 들려져 있다. 컵에 서는 우리 영혼의 힘과 오감에서 나오는 영향력을 상징하는 5 개의 물줄기와 히브리 문자로 ‘손’을 나타내는 ‘요드’형태를 한 26개의 물방울들이 신성한 힘을 나타내는 연꽃의 연못으로 흘 러내리고 있다. 컵 중앙에는 거꾸로 된 'M' 또는 ‘W’ 모양이 그려져 있다. 성체를 문채 컵 안으로 날아들고 있는 하얀 비둘기는 물질세 계에서 성령의 실현을 상징한다. 컵은 마음을 상징한다. 마음 을 연다는 것은 누군가를 만날 준비가 되어 있고, 새로운 인 연과 근접해 있다는 것이다. 아무리 인연이 옆에 있어도, 자신 이 마음을 열 준비가 되어 있지 않으면 그 인연은 멀어지기 마련이다. 그런 인연을 잡기 위한 마음의 준비가 되어 있는 상태이다. 컵 에이스 카드는 새로운 사업이나 프로젝트, 특히 마음을 끌거나 들뜨게 하거나 열정을 느끼게 하는 사업에 착 수할 기회가 있음을 가리킬 수 있고, 새로운 관계가 시작되거 나 사랑의 기회가 주어질 것을 암시한다.", oneCard: "" },
    "cups-2": { description: "화관을 쓴 남자와 월계관을 쓴 여자가 서로 마주보고 서서 서로의 컵을 교환하고 있다. 그들은 컵(감정)을 서로 교환함으 로써 사랑의 서약을 한다. 그들의 컵 위에는 날개 달린 붉은 사자머리와 그 아래에는 헤르메스의 지팡이가 있다. 붉은 사 자는 힘과 열정을, 두 마리의 뱀은 음과 양을 상징하지만 이 카드를 만든 웨이트경은 이 두 개의 상징이 이 카드와는 상관 이 없다고 했다. 이 카드는 무르익은 사랑은 아니지만 로맨스 의 시작 또는 우정을 표현하고 있다. 또는 두 가지 상반되는 생각이나 감정의 화합을 통한 갈등의 해소를 뜻하기도 한다. 이 카드는 사랑에 빠지고, 약혼을 하고, 결혼 날짜를 잡거나 굳은 서약을 하는 등을 나타내며 두 사람이 서로 인연을 맺어 우정이든, 사랑이든 호의를 가진 관계가 시작된다. 대개의 경 우 서로 같은 감정을 느끼고 있으며 새로운 연애의 시작을 의 미한다.", oneCard: "" },
    "cups-3": { description: "3명의 아리따운 아가씨들이 컵을 높이 맞댄 채 원을 그리며 춤을 추고 있다. 한 여인은 순수함을, 또 한 여인은 사랑의 정 열을, 다른 한 여인은 지성을 상징하는 아름다운 옷을 입고 있다. 서로가 서로의 팔을 교차하며 술잔을 모으면서 그들의 감정을 나누고 있는 듯이 보인다. 두 명이 모이면 둘만의 관 계가 되지만, 세 명 이상이 모이면 집단화되어 여러 사람들과 의 인간관계가 잘 이루어진다는 의미이다. 그리고 발아래에는 과일과 채소가 덮여 있고, 매우 풍요롭게 보인다. 춤을 추고 있는 3명의 아가씨들은 사랑으로 합쳐진 관계와 희망찬 미래 를 상징하며, 여성들 사이에 이루어진 특별한 관계나 여자 친 구들끼리 즐기는 것에 대해 말할 수도 있으며 약혼이나 결혼, 탄생의 소식을 가리킨다. 이 카드는 즐거움과 큰 기쁨, 파티 분위기를 나타내며 다가 올 행복, 성취, 풍요로움을 상징하는 카드이다.", oneCard: "" },
    "cups-4": { description: "한적한 교외에서 한 남자가 팔과 다리를 꼰 채 나무에 기대 어 앉아 있다. 그는 무언가 불만스럽고 귀찮은 듯이 앞에 놓 인 3개의 컵을 바라보고 있다. 하지만 자세히 보면 그는 컵들 을 바라보는 것이 아니라, 새로움과 유혹을 상징하는 구름 속 손에 들린 컵을 외면하고 있는 듯하다. 그는 지금 그 무엇에 도 관심이 없다. 새로운 컵에 관심이 없는 태도는 네 개의 컵 이 기존의 세 개의 컵을 가지고 있던 관계를 깰 수 있다고 경 계하는 것이다. 이미 기존의 돈독한 관계에 새로운 사람이 왔 을 때 그 사람에 대한 인간이 가진 기본적인 적대감, 배척감 이 이 카드의 주요한 의미이다. 이 카드에서 보여주는 불만과 지루함은 사회나 주위에서 온 것이 아닌, 바로 주인공 자신 안에서 나오는 것이다. 자신을 재정비하기 전까지는 그 무엇도 그의 흥미를 끌지 못 할 것이며, 끝없는 정체기 속에서 헤어 나오기 힘들 것이다. 이 카드는 당신이 어떤 사람에 대한, 대개 당신의 감정에 응 답하지 않는 사람에 대한 감정에 너무나 사로잡혀 있어서, 다 른 기회들 심지어 바로 코앞에 있는 기회들조차 놓치고 있다 는 것을 의미할 수 있다. 이 카드는 어떤 사람이 이미 끝났거 나 심지어 제대로 가져 보지도 못한 관계에 집착하며 과거에 갇혀 있을 때 나오는 경우가 많으며, 실망감으로 괴로워하며 그런 생각에서 벗어나지 못하는 상황에서도 나올 수 있다.", oneCard: "" },
    "cups-5": { description: "검은 망토를 걸친 남자가 자기 앞에 쓰러져 있는 3개의 컵을 바라보며 슬퍼하고 있다. 그의 뒤에 있는 2개의 컵은 똑바로 세워져 있지만 그에게 위안이 되지 못하는 듯하다. 멀리 보이 는 성과 그 남자 사이에는 강이 흐르고 있고, 강 오른쪽에는 성에 무사히 다다를 수 있는 다리가 놓여 저 있다. 여기서 검 은 망토는 슬픔과 침울함을 나타내고, 바닥에 쏟아져 있는 붉 은색과 연두색 액체는 활기와 번영을 나타낸다. 검은 망토를 걸친 주인공은 더 이상 잃을 희망도 용기도 없다는 듯 망연자 실한 모습이다. 5개의 컵이 모두 쓰러진 것은 아니지만 이 카드의 주인공은 쓰러진 3개의 컵에 실망을 하고 있다. 아직 쓰러지지 않은 2 개의 컵은 그의 관심 밖이다. 즉, 남아 있는 컵에서 새로운 시 작을 할 생각보다는, 오직 쓰러진 컵에만 관심을 갖고 슬퍼하 고 있는 것이다. 여기서 우리는 긍정적인 생각의 중요성을 발견할 수 있다. 부정적인 마음으로는 아무것도 다시 시작할 수 없다. 아직 늦 지 않았다. 미래는 당신의 손에 달려 있다.", oneCard: "" },
    "cups-6": { description: "한 소년이 흰 꽃이 가득 담긴 컵 하나를 나이 어린 소녀에게 건네주고 있다. 아마 둘은 친구 사이일 것이다. 마을은 풍요롭 고 포근한 느낌이다. 마을 곳곳에 있는 6개의 컵 안에는 순수 함을 나타내는 하얀 꽃이 가득 담겨 있다. 이것은 오랜 친구 또는 연인과의 새로운 만남을 나타낸다. 이 카드에선 그림의 대부분이 ‘큰 행복’을 상징하는 노란색으로 되어 있는데, 이것 은 과거의 추억이 매우 아름다웠다는 것을 나타낸다. 또 다른 의미로는 너무 아름다웠던 추억을 그리워한 나머지 과거에 빠 져 현실을 잘 보지 못하는 것으로도 해석할 수 있다. 너무 과 거에 빠져 집착하면 안 된다. 이 카드는 옛 애인을 다시 만날 때, 혹은 어릴 적부터 알던 사람이나 적어도 오래전부터 알던 어떤 특별한 사람이 우리의 삶 속으로 다시 들어오는 것을 나타내는 경우가 많이 있다. 현재 진행되고 있는 관계이지만 과거에 강한 뿌리를 가진 관 계를 가리킬 때도 나온다.", oneCard: "" },
    "cups-7": { description: "모호한 회색 구름 속에서 나타난 7개의 컵 안에 보석, 월계 관, 드래곤, 뱀과 성 등 갖가지 환영들이 펼쳐져 있다. 그것을 바라보고 있는 한 남자의 뒷모습에는 어리둥절하여 무엇을 선 택해야 할지 어쩔 줄 몰라 하는 모습이 역력하다. 이 카드는 환상과 같은 실체 없는 일들을 상징한다. 하루빨리 정신을 가 다듬고 현실을 직시할 수 있는 힘을 키워야 할 것이다. 그렇 지 않으면 하루하루를 환상에 빠져 허우적대고 말 것이다. 이 카드는 위험을 감수할 필요가 있고, 당신이 아직 망설이 고 있는 관계나 행동 방향에 대해 입장을 분명히 할 필요가 있음을 암시한다. 이 카드는 거의 항상 혼란스러운 상태를 표 시한다. 남의 떡이 더 커보이듯이 하고 싶은 건 많고, 갖고 싶 은 것도 많지만 그중 어느 것도 구할 수 없다. 환영을 좇다가 는 아무것도 남지 않는다는 경고를 담고 있다. 자신의 목표가 뚜렷하지 않거나 세운 목표가 허황되지 않은 가 생각해 보아야 한다.", oneCard: "" },
    "cups-8": { description: "붉은 옷을 입은 한 남자가 잘 정돈되어 있는 8개의 컵을 뒤 로 한 채 지팡이를 짚으며 험한 길로 가고 있다. 잘 정돈된 8 개의 컵은 그가 시간과 노력을 들여 만들어 놓은 것이며, 그 는 그것을 버리고 새로운 길을 찾아 떠나는 듯하다. 하지만 그의 붉은 옷에서 단지 다른 길을 찾아 떠나는 것일 뿐, 용기 와 힘을 잃은 것이 아니라는 것을 알 수 있다. 캄캄한 하늘에는 초승달과 만월의 모양이 합쳐진 달이 침울 한 표정으로 그를 내려다보고 있다. 이 카드는 5번 컵과 비교 할 수 있다. 5번 컵 카드의 주인공은 완전히 망연자실한 모습 을 하고 있는 반면, 이 카드의 주인공은 다른 목표를 찾아서 제법 씩씩하게 자신의 길을 떠나는 모습이다. 이 카드의 의미 는 실패라고 보기보다는 자신이 버리고 떠나는 것이라고 볼 수 있다. 모든 인간관계를 정리하고 싶은 마음. 심적으로 너 무나 피곤해서 모든 걸 뒤로 하고 휴식을 취하러 떠나는 모습 이다. 다른 어떤 제안도 더 이상 소용없다.", oneCard: "" },
    "cups-9": { description: "뚱뚱한 남자가 매우 만족스러운 표정으로 팔짱을 낀 채 나무 로 만든 의자에 앉아 있다. 그 위에 그를 중심으로 반원형 모 양으로 와인이 가득 담긴 9개의 컵들이 가지런히 한 줄로 놓 여 있다. 그가 않아 있는 나무 의자는 그의 출신성분이 귀족 이 아닌 것을 나타낸다. 그러나 그는 물질적, 정신적으로 매우 풍요롭고 안정된 상태임을 알 수 있다. 이 카드는 물질적, 정신적으로 완벽한 풍요로움을 나타내는 카드이다. 자신이 바라던 모든 것을 완벽하게 손에 넣은 것을 나타낸다. 그리고 사회적 성공보다는 인간관계의 성공을 의미하며, 마 음은 평화를 찾았고, 물질적인 성공뿐만 아니라 심리적인 안 정도 얻을 수 있다. 주변에는 믿음직한 든든한 사람들뿐이고, 마음에는 만족감이 가득하다.", oneCard: "" },
    "cups-10": { description: "부부로 보이는 한 쌍의 남녀가 하늘에 떠 있는 무지개 모양 으로 빛나는 10개의 컵을 경이롭게 올려다보며 서 있다. 남 자의 오른팔은 여자의 허리를 감싸고 있고, 남자의 왼팔과 여 자의 오른팔은 하늘을 향해 뻗어 있다. 그들의 오른편에는 두 명의 아이들이 서로 손을 마주잡고 즐겁게 뛰놀고 있고, 오른 편 멀리에는 그들의 집이 보인다. 한 가족이 행복과 즐거움을 함께 나누고 즐기는 행복한 느낌 의 카드이다. 이 카드에서 10개의 컵과 무지개는 한 가족의 평화와 안정을 약속해 주는 상징이다. 이 카드는 함께 사는 것, 결혼, 자녀들, 깊은 만족 그리고 지 속적인 안정감을 나타낸다. 당신이 소망하는 목표가 어떤 분 야이건 이 카드는 성공과 성취, 축하를 나타낸다.", oneCard: "" },
    "cups-page": { description: "컵 소년은 호기심 많고, 마음이 따뜻하며 사랑스런 수다쟁이 이며 어린아이와 같은 순수함과 유리알 같은 예민한 영적인 능력도 있어, 상상력도 풍부하고 낭만적인 사람이며 때로는 별난 행동과 말을 하기도 한다. 그러나 예민한 데 비해 때론 눈치가 빠르지 못해 남들보다 한 템포 늦게 알아채기도 한다. 대부분 귀여운 이미지를 갖고 있고 사랑스런 미소의 소유자이 며 감수성이 뛰어난 특성이 있다. 창조성이 뛰어나 예술 방면 에도 관심이 많고 새롭고 신기한 것에 금세 흥미를 갖지만, 그리 오래가지 못하는 편이다. 영적인 정신세계를 상징하는 물고기와의 만남과 대화는 그가 영적인 능력을 갖고 있다는 것을 의미한다. 또한 지구상의 모든 생명체가 바다에서부터 시작되었듯 컵(자궁)에서 나오는 새로운 생명체(아이의 탄생) 의 의미도 된다. 컵에서 올라오고 있는 물고기는 탄생을 상징 할 수 있으며, 끝까지 실행된다면 매우 생산적일 수 있는 아 이디어의 잉태나 임신을 가리킨다. 물고기는 또는 새로운 관 계를 시작하고 있고 감정과 관계의 세상 속으로 돌아가는 법 을 배우고 있는 사람의 재생을 상징할 수 있다.", oneCard: "" },
    "cups-knight": { description: "컵 기사는 청혼과 평화의 카드이다. 기사는 컵을 내밀고 있 으며, 유순한 회색 말을 타고서 조용히 앞으로 나아간다. 이 카드는 열정적이고 급진적인 것보다는 차분하고 감성적인 면 이 우세하다는 것이다. 그는 상냥하고 로맨틱하며, 부드럽고 예의 바르고 잘 생긴 청년이다. 남성적이기 보다 여성적인 면 이 두드러지고, 이성적이기 보다 감성적이어서 분위기 파악도 잘하고 인간관계도 좋다. 하지만, 감정적인 부분이 민감하여 쉽게 질투하고 지나치게 감상적이 되거나, 자제력이 부족한 면도 있다. 이 사람은 애정이 깊고, 성실하며, 섬세한 사람이 다. 이 카드가 전통적으로 멋진 구혼자로 여겨지던 인물, 청혼 을 할 사람 또는 성실하고 헌신적인 관계를 원하는 사람을 나 타내며 신혼 시절이 지나간 뒤에 직면하게 될 관계의 현실적 인 면에 대해서는 아직 진정으로 준비되어 있지 않을 수도 있 다. 컵 기사는 업무적인 제안에서부터 특히 당신의 마음을 끄는 아이디어나 계획들에 이르기까지 모든 종류의 제안을 나타낼 수 있다. 4개의 기사 카드 중에 유일하게 뒷배경에 강이 흐르고 있는 카드이다.", oneCard: "" },
    "cups-queen": { description: "컵 여왕 이러한 사람들은 감정의 기복이 심하고, 무척 예민 하며 무의식이 잘 발달했기 때문에 한번 상처를 받으면 치유 될 때까지 오랜 시간이 걸린다. 따라서 우울증에 걸릴 확률도 높다. 개인적으로는 따뜻하고 예술적이며, 아름다운 것을 좋아 하는, 감수성이 예민한 여성적인 성격이다. 이 사람은 낭만적 인 기질과 풍부한 상상력, 큰 꿈을 갖고 있는 사람이다. 그리 고 머리가 가슴을 지배하는 검의 여왕과는 많은 면에서 정반 대인 사람이다. 컵의 여왕은 검의 여왕과는 반대로 감정의 세 계에서 사는 사람이다. 컵의 여왕은 사랑에 빠져서 온통 희망 에 부풀어 있으며 마법처럼 펼쳐지는 세상에 놀라워하는 사람 을 나타낼 수도 있다. 78장의 카드 중 가장 화려하면서 뚜껑 이 닫혀 있는 이 컵은 유일하게 이 카드에만 등장한다. 직관 력이 네 명의 여왕 중 가장 뛰어나고 순수하고 오직 사랑이며 경제력은 약한 카드이며 강한 모성애를 가졌기에 여성적인 힘 이 강하다고 할 수 있다. 뚜껑이 닫혀 있는 컵은 그녀의 본심을 쉽사리 다른 이들에게 보이지 않는다는 것을 뜻하고, 그녀가 무의식과 깊은 연관이 있음을 보여 준다.", oneCard: "" },
    "cups-king": { description: "이러한 사람들은 다른 사람들과의 관계나 자신을 표현하는 데 있어 간접적이면서 예의바르고 정중하다. 다시 말해서 영 국신사를 떠올리게 하는 인물이다. 주변사람들을 카리스마가 아니라 정으로 이끈다. 항상 주변 사람을 먼저 배려하며, 뚜렷 이 자기주장을 하지 않을 수 있다. 이 사람은 꿈을 꾸지만 그 다지 확신이 없고, 상상력이 풍부하지만 반드시 집중되어 있 지는 않은 사람이다. 왕은 뗏목 위에 실려 표류하고 있으며, 특정한 상황이나 관계에 대해서건 자신의 진로나 인생의 목적 에 대해서건 어찌할 바를 모르고 있다. 이런 사람은 좋은 쪽 으로 보면 사색적이고 감수성이 강하지만, 최악의 경우에는 제멋대로만 하려고 하며 자기 기만적이다. 하지만 평소에는 바다와 같이 넓고 열린 마음으로 다른 사람을 이해하고 받아 들여 뛰어난 협상가로도 유명하다. 그리고 음주가무와 예술을 사랑한다. 이 사람은 바다에 떠 있기 때문에 지속성이나 목적 이 없을 수 있다. 감정의 기복이 심하기 때문에 사랑에 있어 서 일편단심이 힘들며 마음의 평정을 찾아야 하는 카드이다.", oneCard: "" },
    "swords-ace": { description: "구름 속에서 나타난 빛나는 손이 양날의 검을 잡고 있다. 검은 금빛 왕관을 관통하고 있고, 왕관의 좌우에는 각각 승리 의 상징인 야자나무 잎과 평화의 상징인 올리브 나뭇가지가 걸쳐져 있다. 에이스는 기회와 신선한 출발을 의미하며, 새로 운 모험적 프로젝트나 사업을 말하는데, 대개는 시작 단계나 미성숙 단계에 있는 것을 나타낸다. 이 검은 왕관을 쓰고 있 기에 도전의 추구나 야망의 실현은 분명히 성공할 것을 의미 한다. 그리고 진실과 정의의 검도 상징할 수 있으므로 소송을 진행 중이거나 다른 사람을 위해 대신 싸우고 있을 때도 나타내며 용기, 지략, 일편단심의 필요성을 나타내기도 한다. 검은 이성 적인 판단을 해야 하는 상황을 가장 잘 표현하고 있으며 현실 을 직시하며 껍질을 뚫고 그 안의 본질적인 모습을 보려고 노 력하는 모습이다. 뿐만 아니라 어떠한 마음의 결심을 세우는 모습이기도 하다. 이것은 하늘이 주신 능력과 기회라고 할 수 있으며, 여기서 오른손은 긍정적인 힘과 남성다운 힘을 나타 낸다.", oneCard: "" },
    "swords-2": { description: "한 여인이 눈을 가린 채 회색 가운을 입고 의자에 앉아 있 다. 그녀의 손에는 두 개의 긴 검이 들려져 있다. 그녀는 무언 가에 대한 방어적인 자세를 취하고 있는 듯 두 개의 검을 가 슴 쪽에서 X 자형으로 들고 있다. 그녀 뒤에는 바다가 있고, 바다 위에 험한 바위가 군데군데 보인다. 여기서 바다는 감정 을, 바위는 긴장감을 상징한다. 그녀 위에 떠 있는 초승달은 새로운 시작을 알리고 있으며, 바닥과 그녀의 옷과 의자의 회 색은 상응되는 양쪽을 모두 포용한다는 중립적인 의미의 상징 이다. 두 개의 일, 사건을 뜻하고 그 두 개가 똑같이 중요하다. 그 중 어떤 것도 포기할 수 없고, 더 무거워서는 안 된다. 여기서 자연스럽게 힘의 균형을 떠올릴 수 있으며 물론 힘의 균형으 로 끝나는 게 아니라 그중에 하나를 선택하도록 강요받는 경 우가 더 많다. 이 카드는 현재 균형이 잡혀 있는 상태이기는 하지만 불안감 에 더 이상 나아가지 못하고 있음을 나타낸다. 그녀의 손은 묶여 있지 않다. 눈을 가린 천은 마음만 먹으면 언제든 벗어 버릴 수 있다. 즉, 모든 것은 마음먹기에 달렸고, 용기를 갖고 도전하라는 뜻이다.", oneCard: "" },
    "swords-3": { description: "비구름이 가득하고 장대비가 내리는 우중충한 하늘에 붉은 하트가 3개의 날카로운 검에 관통 당했다. 하늘과 구름은 슬 픔과 불확실성을 나타내는 회색이고, 칼에 찔린 하트에서 가 슴 저린 슬픔과 고통을 느낄 수 있다. 비록 심장을 찌르는 듯 한 고통을 나타내지만, 그럼에도 불구하고 상당히 고요하고 심지어는 안정된 것처럼 보인다. 이 카드는 우울한 기분이나 씁쓸한 실망감, 좌절이 엄습하는 카드이면서 최악의 경우에는 당신에게 깊은 상처를 준 어떤 사람 혹은 어떤 것으로 인한 가슴이 찢어지는 슬픔과 아픔, 고통을 가리키는 카드이다. 그리고 세 개의 검은 유일하게 애 정에 관련된 카드이다. 실연, 이루어지지 않는 사랑이나 삼각관계를 의미한다. 검의 상징 중 갈등을 포인트로 하여 그 갈등이 머릿속이 아니라 마 음속에서부터 나온다는 것을 의미하며 그에 따른 실망과 슬픔 에 대한 카드이다.", oneCard: "" },
    "swords-4": { description: "한 기사가 무덤에서 기도하는 형상으로 누워 있다. 벽에는 참회자에게 은총을 내리는 듯한 스테인드글라스가 있고, 그 오른편에는 세 개의 검이 나란히 걸려 있다. 나머지 한 개의 검은 관 옆에 눕혀져 있다. 이 카드의 주인공은 현실과 고통 에서 벗어나 치유를 위한 휴식에 들어가 있다. 지금은 모든 것에서 한 발짝 물러나 시간과 여유를 갖고 휴식을 취해야 할 시기임을 암시하는 카드이다. 이 카드는 회복의 카드 일 때가 많은데, 질병이나 정신적 외 상을 겪은 뒤의 회복과 치유를 나타낸다. 네 개의 검이 가지 느낌은 죽음과 같은 휴식이다. 열심히 일한 뒤의 휴식은 꿀맛 같지 않을까? 피곤에 지쳐 쉬고 있지만 다음에 해야 할 일이 남아 있는지 편안해 보이진 않는다. 많은 일을 한 뒤의 휴식 이나 휴식의 필요성이라 할 수 있다.", oneCard: "" },
    "swords-5": { description: "전투에서 진 두 명의 남자가 검을 버리고 처량하게 뒤돌아 가고 있다. 한편 승자는 의기양양하게 다섯 자루의 검을 모으 며 미소를 띤 채, 패배자들을 조롱하듯 쳐다보고 있다. 하늘은 패자들의 마음을 나타내는 듯 회색 구름이 이리저리 흩어져 있다. 이 카드에서 주목할 점은 주인공이 승자가 아닌 패자라 는 것이다. 이 카드는 경고의 카드이다. 여기서 보여 주는 것과 같은 패배자가 되지 않기 위한 노력 이 필요하다는 것을 알려준다. 대적하기에 강한 적이나 얻고자 하는 것을 주지 않는 상대와 부딪쳤음을 의미하며 싸움이 있었는지, 또한 일어날 것인지는 카드의 배치를 보고 유추할 수 있다. 또한 심각한 타툼을 나 타내며 성격의 충돌이나 의지간의 싸움일 수도 있다. 주된 메 세지는 파괴적인 상황에서 벗어나는 것이다. 이유는 확실하게 패배할 싸움을 하고 있기 때문이다. 만약 승리자의 입장이라 도 무의미한 승리일 수도 있고 승리의 맛도 씁쓸한 마음으로 바뀔 수도 있다.", oneCard: "" },
    "swords-6": { description: "이미 큰 슬픔을 경험한 듯 옷으로 가린 채 웅크리고 있는 한 여인과 아이가 배에 타고 있다. 그 배를 뱃사공이 노를 저어 강 건너편으로 향하고 있다. 배에는 무거운 짐 대신 여섯 개 의 검이 꽂혀 있다. 배 오른편의 강물은 거칠지만, 왼편의 목 적지로 흐르는 강물은 잔잔하다. 이것은 고통에서 안정된 미 래로의 여행을 암시한다. 이 카드는 현재의 상황보다 더 나은 미래로의 이동 혹은 여 행을 나타낸다. 지금까지의 걱정과 근심으로부터의 벗어남을 암시한다. 그리고 카드에 나타난 슬픔의 느낌은 슬픈 사연으 로 인해 여행을 한다는 것을 의미한다. 또한 이 여행은 필수적이며 만약 여행을 떠나지 않는다면 더 이상 도움이 되거나 건강하거나 살아 갈 수가 없는 상황에 계 속 머물게 될 것이다. 더 나은 곳으로 간다면 이별의 슬픔은 치유될 것이다. 또한 감정보다는 이성을 앞세워야 하는 경우 가 많고 자기발전과 향상의 카드이다.", oneCard: "" },
    "swords-7": { description: "한 남자가 가까운 캠프로부터 다섯 개의 검을 서둘러 나르고 있다. 나머지 두 개의 검은 아직 캠프에 남아 있는데, 이 남자 는 아마도 일곱 개의 검을 다 나르려고 했던 것 같다. 카드의 주인공은 자신감에 찬 얼굴로 검을 나르고 있지만, 서두르다 그만 두 개의 검을 빼먹은 것이다. 이 카드는 성급히 서두르는 것에 대한 경고의 카드이다. 급 히 서두르다 보면 잊고 지나가는 중요한 것들이 분명히 있기 에 어떤 상황에 대해서 결전을 벌이는 것이 아니라 그 상황을 피해 도망치는 카드이다. 즉 어려운 상황에서 탈출하기 위해 잔꾀를 쓰고 몰래 하는 행위를 떠올릴 수 있다. 어떤 상황들 에서는 손을 떼고 물러나는 것이 가장 좋은 대처 방법이다. 또한 어떤 일을 처리한 방법에 대해 죄의식을 느낄 때 나오 기도 한다. 아마도 어떤 사람을 저버렸다고 느끼기 때문일 수 있으며, 즉 단호함이 결여되어 있고 문제를 직면하지 않고 몰 래 가져가 버릴 수도 있다. 만약 다른 사람을 가리키는 것이 라면 그 사람은 죄의식을 느끼고 있거나 어떤 것에 대한 당신 의 대응을 두려워하고 있음을 의미할 수도 있다. 아니면 자신 의 책임을 회피하거나 피하고 있을 지도 모른다. 어떤 종류의 직면은 유용할 수 있다. 문제들을 바깥으로 끄집어 낼 수 있 는 방법에 대해 의논하는 것이 좋다.", oneCard: "" },
    "swords-8": { description: "눈이 가려지고 손은 뒤로 묶인 여인이 늪지대에 서 있다. 그녀 뒤편으로 벼랑 위의 성이 있다. 그녀 바로 뒤에는 8개의 검이 한 줄로 꽂혀 있어, 그녀가 성으로 가지 못하게 막고 있 다. 하지만 그녀의 앞길은 열려 있다. 단지 그녀의 눈이 가려 져 있기 때문에 앞으로 갈 수 있는 것을 모를 뿐이다. 이 카드는 현실을 똑바로 인식할 수 없다는 것을 가르쳐 준 다. 주인공의 눈이 가려 있고 손이 묶여 있어 어디로 어떻게 빠져 나가야 하는지 알지 못하지만, 그녀의 앞길은 열려 있다. 그녀의 발도 묶여 있지 않다. 그녀가 마음만 단단히 먹는다면 빠져 나올 수 있는 현실임을 보여 준다. 이 카드는 어떤 환경이나 누군가에 의해 구석에 몰리거나 갇 혀 있다고 느낀다는 것을 의미한다. 그러나 에워싸여 있다고 느끼더라도 벗어날 수 있다는 것을 깨달기만 하면 빠져나갈 수 있다. 즉 자신이 할 수 있는 건 아무것도 없다고 포기하고 있지만 사실 자신을 묶고 있는 것은 강한 힘이 아니다. 스스 로 벗어나서 두려움을 떨치려 한다면 얼마든지 할 수 있지만 그러한 노력조차도 두려움 때문에 시도해 보지 못한다. 현실 을 직시하는 것이 필요한 카드이다.", oneCard: "" },
    "swords-9": { description: "한 여인이 자다 일어나 두 손으로 얼굴을 가린 채 침대에서 울고 있다. 그녀 옆에는 9개의 검이 나란히 걸려 있다. 침대 옆면에는 한 사람이 다른 사람을 칼로 공격하는 그림이 새겨 져 있다. 그녀가 덮고 있는 이불에는 붉은 꽃들과 별자리, 행 성들의 기호가 그려져 있다. 여기서 주목해야 할 것은 주인공 이 겁에 질려 있지만, 아홉 개의 칼(근심)은 그녀를 위협하거 나 건드리고 있지 않다는 것이다. 이 카드는 외면할 수 없는 현실의 두려움이 있지만, 그것이 주인공에게 직접적인 영향력을 행사하고 있지는 않다는 것을 의미한다. 마음을 굳게 먹고 대면하면 얼마든지 헤쳐 나갈 수 있으며, 지금 현재의 상황이 생각하는 것보다는 덜 심각하다 는 것을 보여 준다. 다만 두려워하는 것이 상상하는 것만큼 나쁘지 않을 수 있지만 진짜 문제는 혼자 고립되어 있다는 점 이다. 즉 혼자 문제를 해결해야 한다고 느끼며 돕거나 완화시 켜 줄 사람이 아무도 없다고 느끼고 있다. 만약 불면증이나 우울증, 공황이나 편집증을 겪고 있다면 문제를 더 악화시킬 수 있다. 검 9번 카드의 숨겨진 메시지는 머리를 감싼 손을 풀고 괴롭 히는 문제가 무엇이건 그것을 치유하는 쪽으로 나아갈 수 있 도록 주위를 둘러보라는 점이다. 검 9번 카드는 일종의 교착 상태나 마비상태를 가리키는 검 2번 카드와 비슷하다.", oneCard: "" },
    "swords-10": { description: "한 남자가 땅에 죽은 듯 엎드려 있고, 그 위에 열 개의 검이 꽂혀 있다. 앞에 보이는 바다는 평온해 보이며, 산 너머에는 새로운 날의 태양이 떠오르고 있다. 주인공은 죽었다. 즉, 더 이상 나쁜 상황은 없음을 나타낸다. 또한 멀리 새벽이 밝아 오는 것으로 희망이 서서히 다가오고 있음을 알 수 있다. 이 카드는 극도로 안 좋은 상황이 종료되고, 새로운 주기가 시작되었음을 알려준다. 보기에는 섬뜩하고 나쁜 의미의 카드 같지만, 그리 나쁘지만은 않은 희망을 주는 카드이다. 어떤 일 들이 고통스럽게 명백한 최후에 다다랐음을 의미한다. 인간관계에서는 필사적으로 잃어버린 사랑을 되찾으려 하거 나 힘든 동반자 관계를 유지하기 위해 노력하지만 슬픈 현실 은 그 관계가 끝났다는 것이다. 즉 인간관계이든 직장이건 다 시 되돌릴 방법은 전혀 없으며 끝내는 것이 최선의 방법일 때 가 많다. 그러나 죽음의 상징들처럼 그 이면에는 재생이라는 것을 명 심해야 한다. 검 10번 카드의 배경의 지평선은 동이 트면서 황금빛 줄무늬로 물들고 있다. 재생의 이미지는 고통이나 상 실을 과소평가할 목적으로 사용되어서는 결코 안되며 힘과 희 망을 주는 용도로 사용된다. 즉 어떤 것의 현재 상태가 끝나는 것이지 인생의 끝은 아니 라는 것이다.", oneCard: "" },
    "swords-page": { description: "비장한 표정의 소년이 언덕 위에서 두 손에 검을 높이 쥐고, 어딘가에서 불쑥 나타날 적들을 경계하는 듯이 뒤를 돌아보고 있다. 어디서 무슨 일이 생기면 당장이라도 달려갈 것 같은 자세이다. 그는 자주색의 옷을 입고, 붉은 부츠를 신고 있다. 이것은 그가 충성심이 충만하며, 충만한 에너지로 자신이 필 요한 곳이면 어디든지 달려간다는 의지를 나타낸다. 검 소년은 재빠르고 예민하지만, 무정할 수도 있다. 검은 공기의 원소와 관계있으므로 공기의 성격, 즉 이해가 빠르고 민첩하며 활기찬 특성들을 지니고 있다. 그러나 다른 사람을 공감하고 그들의 입장을 고려하고 자신의 행동이 상대 방에게 어떤 감정을 초래하는 지에 대해 배려하는 마음이 필 요하다. 또한 검의 소년은 조급하게 서둘러 일을 처리하는 것을 나타 낼 수 있다. 즉 행동을 실행하기 전에 신중하게 고려해 볼 필요가 있음을 의미하므로 원하는 최종결과가 무엇인지를 생각해 보아야 한 다. 검의 소년 카드는 믿음성이 없는 인물, 고의적이든 경솔해 서든 남의 험담을 하거나 말썽을 일으키는 사람들을 나타내는 경우도 있다.", oneCard: "" },
    "swords-knight": { description: "백마를 탄 젊은이가 머리 부분에 붉은 깃털이 달린 갑옷을 입고 오른손에는 하늘높이 검을 치켜든 채 바람과 같이 달리 고 있다. 검 기사는 냉철하고 냉정하게 사리 분별을 할 줄 알 며, 자신의 소신을 굽히지 않는다. 이성적인 판단이 행동으로 먼저 나와 돌발적인 행동을 할 수 있는 카드이다. 네 개의 기사 카드들 중에서 말을 가장 빠르게 몰고 달리는 것이 바로 이 검 기사이다. 검을 높이 치켜들고서 바람을 가 르고 쏜살같이 어디론가 돌진해 가는 듯한 그의 모습에서 기 사 컵의 정적인 모습과는 정반대의 돌격하는 기사의 에너지를 느낄 수 있다. 카드 전체에서 강풍이 세차게 불어 구름과 나 무가 흩어지고, 전속력으로 달리는 말과 기사가 모두 조금은 흥분한 상태임을 느낄 수 있다. 검의 기사는 쉽게 매혹시킬 수 있다. 그러나 유능하고 대담 하지만 무모한 사람일 수도 있고 가는 곳마다 소동을 일으킬 수 있으며 카리스마는 대단하지만 믿음성이 부족할 수 있다. 강력한 동맹자가 될 수 있지만 마음에 들 때일 뿐일 수 있 다. 너무 쉽게 신뢰하지 않도록 조심할 필요가 있다. 즉 검의 기사는 갑작스럽게 일어나거나 뜻밖의 사건을 말해주며 잘 대 처할 사람이 없으면 상황이 나빠져서 혼란에 빠질 위험이 있 다.", oneCard: "" },
    "swords-queen": { description: "엄격해 보이는 여성이 옆모습을 보이며 회색 옥좌에 앉아 있 다. 검 여왕은 전문적인 여성을 말하며 지적이고 자신감에 차 있으며 비판적인 정신과 날카로운 혀를 가지고 있다. 감정보 다는 이성적으로 판단하며, 어떤 면에서는 역경이나 주변의 난관에 강하게 대항할 수 있는 힘이 있는 여성을 나타낸다. 거칠고 독립적인 성격을 가지며 자신을 원하는 것을 위하여 싸우거나 강경한 태도를 취하는 사람이며 공정하고 빈틈없고 논리적이며 사무적인 인물이다. 전통적으로는 과부로 보지만 이혼했거나 별거중인 사람 또는 오랫동안 혼자 지냈고 다른 사람들과 쉽게 친해지지 못하는 사람을 나타낼 수도 있다. 겉 으로는 날카롭고 거칠어 보이지만 외롭고 수줍은 사람이거나 혼자서 세상을 살아가야 한다고 느끼는 사람을 가리킨다. 검의 여왕은 유리한 조건을 이끌어 내기 위해 애써야 하는 힘든 거래나 강력한 반대를 나타낼 수 있어 결과를 쉽게 얻기 는 힘든 것을 의미한다. 일들은 공정하게 규정에 따라 엄격히 처리될 것이며 감상이나 개인적인 것들이 끼어 들 여지는 없 는 것을 의미한다.", oneCard: "" },
    "swords-king": { description: "파란 하늘에는 구름이 떠 있고, 위로 치솟은 돌로 된 왕좌에 왕이 앉아 있다. 검 왕은 공정하고, 논리적이고, 통제되고, 엄 격하며, 규칙을 준수하고, 때로는 가까이 하기 어려운 사람이 며 흑백논리의 성향이 강한 사람이기도 하다. 검 왕은 분석적 이고 효율적이며, 평균 이상으로 지성적인 사람을 가리킨다. 판단을 흐리게 하는 감정의 개입을 허용하지 않으므로 법률 가나 컨설턴트 같은 전문가를 나타낼 때가 많다. 이런 면에서 그는 남을 돕는 데는 유리하지만, 삶에 대해 냉철하게 접근하 기 때문에 인간관계 면에서는 불행한 결과를 초래할 수 있다. 컵 왕이나 컵 기사는 로맨티스 라고 한다면 검 왕은 완벽주의 자 이면서 카리스마 강한 냉정한 남자를 뜻한다. 검 왕으로서 권위나 지성, 판단, 힘이라는 검과 잘 어울리며 정신력이 강하고 영리한 인물을 말하며 자신보다 약한 사람들 을 압도한다. 좋은 경우에는 이성적이고 공명정대한 사람일 수 있으나 반대의 경우에는 감정을 무시한 채 편협한 합리주 의적 관점인 사람을 나타낼 수도 있다. 또한 자기 불신이나 남의 의견을 경청하려는 자세를 의미할 수도 있다. 역으로 나 오면 남들을 지배하거나 이득을 얻기 위해 지적인 힘을 오용 할 수 있는 독재자이거나 폭군일 수 있다.", oneCard: "" },
    "pentacles-ace": { description: "구름 속에서 나타난 손이 황금빛으로 빛나는 펜타클을 받쳐 들고 있다. 그 아래는 노력의 대가를 상징하는 아름다운 정원 이 보이고, 정원에는 붉은 꽃과 흰 백합이 만발해 있다. 아치 형 넝쿨을 통해 멀리 산을 향해 작은 길이 나 있다. 저 멀리 보이는 산은 우리의 목적지이다. 출발지인 정원은 풍요롭게만 보이고, 가는 길도 험하지 않다. 여기서 우리는 목 표(물질적 안정과 이득 그리고 풍요)를 향해 나아가는 시작이 풍족하다는 것을 알 수 있다. 펜타클 에이스는 돈과 함께 하는 신선한 출발을 나타낸다. 예를 들어 월급인상, 대출이나 증여 또는 유산형태로 큰돈이 들어 올 수도 있다. 또한 번창하게 될 새로운 일자리나 사업 기회를 나타낼 수 있다. 또는 보석, 금 같은 큰 값어치 있는 선물을 의미하기도 한다. 반면에 에이스 카드가 문제를 상징하는 카드와 함께 나오면, 돈 문제와 관련하여 실망하거나 퇴보할 수 있으므로 유의해야 한다.", oneCard: "" },
    "pentacles-2": { description: "한 남자가 양 손에 펜타클을 들고 곡예를 하듯이 한쪽 발을 들고 서 있다. 두 개의 동전은 8자 모양의 줄로 연결되어 있 으며, 이 상태에서 동전을 떨어뜨리지 않으려면 균형을 잘 잡 아야 한다. 그러나 그런 곤란함 속에서 균형 있는 대처만이 이 난관을 탈출할 수 있다. 대부분의 경우 두 가지 일을 잘 해나간다고 본다. 펜타클 2번은 자금 압박을 받거나 수지 균형을 맞추기 위해 애쓰고 있을 때, 또는 빚을 갚아야 하는 상황(카드 돌려막기) 에 있을 때, 혹은 두 가지 직업을 가지고 일할 때에도 나타난 다. 또한 돈 문제와 관련된 변화를 가리키기도 한다. 때로는 상 황이 매우 위험해질 가능성이 있다는 것을 알지만 현재로는 다른 선택의 대안이 없거나 현실과 직면하고 싶지 않기 때문 에 등을 돌리고 무시하고 있는지도 모른다. 만약 재정적인 문 제와 상관이 없어 보인다면 재산을 유지하기 위해 모든 정력 을 소비하고 있는 사람을 가리킬 수도 있다. 펜타클 2번 카드는 실용적이고 실제적이고 해결책을 찾는 등 딜레마를 해결하도록 도울 것이라는 것을 의미한다. 그렇게 되면 소극적으로 현상을 유지하는 데에만 정력을 낭비하는 것 을 볼 수 있다.", oneCard: "" },
    "pentacles-3": { description: "두 명의 성직자들이 손에 설계도를 들고 있다. 그들은 한 명 의 장인과 성당의 세부적인 조각 작업에 대해 이야기하고 있 다. 뾰족한 아치 기둥 위에는 3개의 펜타클이 새겨져 있고, 그 아래에는 장미십자가 문양이 있다. 이 카드는 습득한 기술, 혹 은 전문 직종을 통한 물질적인 이득을 얻게 됨을 나타낸다. 혹은 두 사람 이상의 동업을 의미하기도 한다. 펜타클 3번 카드는 사업 시작이나 승진 등 돈을 버는 것과 관련된 협상이 진행되고 있다는 것을 나타낸다. 또는 불어난 돈을 자산구입이나 투자 등에 사용하는 것을 가 리킬 수 있다. 또한 필요한 전문 기술이나 지식을 갖추고 있 는 제 3자의 필요성이나 존재를 가리킨다. 만약 새로운 벤처 사업이나 프로젝트를 의미할 경우에는 전 문가의 도움이나 안내를 받는 것이 좋다는 것을 상징한다. 왜냐하면 3번 카드는 기초를 쌓고 시간을 들여 조사하고 협 상의 기술을 배우는 것을 가리키기 때문이다.", oneCard: "" },
    "pentacles-4": { description: "한 남자가 펜타클 왕관을 쓴 채 가슴에 펜타클을 꼭 안고 앉 아 있다. 그의 발 양쪽 밑에 각각 한 개씩의 펜타클이 있다. 머리 위의 펜타클은 그의 사고방식이 물질적이고 이해 타산적 이라는 것을 의미하고, 가슴에 안긴 펜타클은 그의 마음 또한 인색하다는 것을 보여 준다. 발아래와 품속의 펜타클은 금전 에 대한 그의 집착을 보여 준다. 이 주인공은 마치 고리대금 업자 혹은 구두쇠와 같은 이미지로 비친다. 이 카드는 인색함과 물질에 대한 애착을 버리고, 너그러움과 베푸는 마음을 가져야 함을 경고해 주는 카드이다. 좋게 보면 이 카드는 자기의 자산을 잘 지키고 모은 것을 굳게 붙잡고 있으라는 경고의 카드이기도 하다. 그러나 최악의 경우는 부유하지만 함께 나누려 하지 않는 인 색함을 나타낸다. 즉 중요한 것은 재물뿐이므로 재물이 줄어 드는 것을 두려워하여 재물의 안전에 집착하여 두려워하는 것 이다. 그래서 금전적인 상황이나 다른 상황들이 걱정되어 꼼짝 못 하고 있을 수 있다. 이런 상태에 있는 한 아무 것도 바뀔 수 없다. 즉 적당히 놓아 버리고 마음을 편안하게 하여 에너지나 현금이 잘 흐르도록 하는 것이 요구된다.", oneCard: "" },
    "pentacles-5": { description: "눈보라치는 추운 거리를 두 명의 걸인이 걸어가고 있고, 그 들 옆에는 교회의 불빛이 보인다. 남자는 다리를 다쳐 목발을 짚은 채 힘겹게 걷고 있고, 여자는 추위를 이기기 위해 숄을 머리에 뒤집어쓰고 한껏 움츠린 채 맨발로 걷고 있다. 그들은 헐벗고 굶주렸다. 하지만 그들은 따뜻한 교회의 불빛을 보지 못하고 지나치고 있다. 이 카드는 어렵고 각박한 현실에 휩싸여 자신들에게 내미는 따사로운 구원의 손길을 알아차리지 못하고 지나칠 수 있다는 것을 보여주며 금전적인 궁핍을 나타내고 일자리를 잃었거나 병이나 장애 때문에 일할 능력이 없을 때, 아니면 단순히 수 지를 맞추기 위해 안간힘을 쓰고 있을 때도 나온다. 또한 영적 빈곤을 의미할 수도 있으며 인생의 의미가 없어졌 다고 느낄 때 나오기도 한다. 그리고 차가운 세계에 내쳐졌다 는 의미도 있다. 다른 사람에게 도움을 구했지만 거절당했을 수도 있고 다른 사람과 함께 하기를 원하지만 그 사람이 마음 의 문을 굳게 닫고 열어 주지 않아 절망적인 외로움을 느낄 수도 있다. 어떤 상황이든 도움이 필요하고 금전적이든 정서 적이든 문제를 바로 잡기 위해 어떤 조치가 필요하다는 것은 분명하다.", oneCard: "" },
    "pentacles-6": { description: "부유한 남자가 저울을 사용하여 가난한 자들에게 균등하게 재물을 나눠주고 있다. 이 관대한 남자 양 옆에는 두 명의 남 자가 무릎을 꿇고 손을 벌린 채 존경과 기쁨에 찬 얼굴로 이 남자를 우러러보고 있다. 펜타클 4번에서 스크루지처럼 아까 워하는 구두쇠를 볼 수 있었다면, 펜타클 6번에 나오는 사람 은 크리스마스 뒤에 변한 스크루지의 모습이라고 생각하면 비 슷하다. 우리는 이 카드에서 자비, 친절, 그리고 관대함을 느 낄 수 있다. 이 카드는 나눠주는 기쁨을 나타낸다. 그리고 앞으로 다가올 금전적인 도움이나 혜택도 의미하며, 빚을 갚는 카드이기도 하다. 빌려 간 돈을 갚거나 빌린 돈을 갚을 필요가 있음을 나 타낸다. 즉 채권과 채무를 청산하는 경우이다. 돈은 주는 인물이 유복하고 행운을 나누어 주는 위치에 있으 므로 자선의 의미도 있다. 천칭은 공정함과 균형의 상징이므 로 어떤 것을 돌려주거나 베풀어 준 지원이나 호의에 감사를 표현해야 할 수도 있다. 도움을 베푸는 사람이 없다면 대출을 받거나 준비할 필요가 있음을 알려 줄 수도 있다.", oneCard: "" },
    "pentacles-7": { description: "한 젊은이가 긴 괭이에 턱을 기대고 포도넝쿨을 물끄러미 바 라보고 있다. 수확이 가까워진 풍성한 포도넝쿨이다. 포도넝쿨 에는 포도 대신 펜타클 6개가 달려 있고, 나머지 한 개는 그 의 발아래 땅에 떨어져 있다. 그는 포도를 수확하기 전에 무 언가를 계획하고 있는 듯하다. 이 카드는 그동안의 땀과 노력의 대가로 앞으로 얻게 될 물 질적 수확에 대한 몰두와 계획을 나타내며, 식상하다거나 권 태로운 모습이기도 하다. 그리고 어떤 상황 속에서 얼마나 멀 리 왔는지를 돌이켜 볼 때도 많지만 실제로 성취를 나타낼 때 도 많으며, 아직은 해야 할 일이 더 남아있고 가야 할 길이 많이 남아 있어 잠시 일을 멈춘 것뿐이다. 즉 어떤 갈림길을 나타내는 경우가 많으며, 같은 길로 계속 가야할 지 아니면 새로운 길로 가야 할지를 선택해야 할 수도 있다. 노력의 결실을 맺거나 보람이 있을지는 더 두고 봐야 한다. 주된 메시지는 계속하라는 것이며 장기간의 멈춤은 잘못 일 수 있다는 것이다. 그러므로 이미 해 놓은 일을 이용할 때이 다.", oneCard: "" },
    "pentacles-8": { description: "한 젊은 장인이 열심히 별 모양을 새기는 일에 집중하고 있 다. 옆에 있는 나무로 된 기둥에는 이미 완성된 5개의 펜타클 이 자랑스럽게 걸려져 있다. 두 개는 의자 밑에, 나머지 하나 는 아직 작업 중이다. 이 카드는 부지런히 같은 모양의 펜타클을 만들면서 얻어가 는 자신만의 기술과 능력이, 미래의 큰 자산이 될 것이라는 의미를 담고 있으며, 자신이 가진 능력을 더욱 개발하여 한 단계 더 성장하기 위한 노력의 모습이다. 새로운 일을 배우는 누군가의 제자로 들어가서 제일 밑바닥 일부터 다시 배우기 시작한다. 처음부터 다시 시작하는 노력이 필요한 카드이기도 하다. 그리고 이 카드는 대부분 열심히 일하는 기간을 나타내지만, 반드시 길고 고된 작업을 의미하는 것은 아니다. 즉 연마한 기술을 완성시키거나 더 큰 수익을 가져올 새로운 기술을 연 마할 때이다. 그러므로 종종 일할 기회나 새로운 프로젝트를 암시하며 이미 이룬 일에 대한 자부심을 나타내기도 한다.", oneCard: "" },
    "pentacles-9": { description: "아름다운 옷차림을 한 부유한 여성이 근사하고 풍요로운 배 경 속에서 편안히 있는 모습을 보여 주고 있다. 이 카드의 분 위기는 풍부함과 고요함 그리고 안정을 나타낸다. 컵 카드와 마찬가지로 여기에서도 펜타클 9번, 10번 카드가 가장 좋은 카드이다. 이 여성의 유일한 친구는 그녀의 손목에 앉아 있는 새이다. 새 사냥은 부유한 계층의 스포츠이므로 이것은 그녀가 여유로 운 상류 계층이라는 것을 암시한다. 따라서 진정한 부유함이 나 대대로 전해진 재산, 또는 돈은 문제되지 않는다는 것을 의미한다. 또한 고독이라는 느낌은 신체적으로 혼자 있는 것을 나타내 며, 잠시 배우자가 부재중일 수 있음을 의미한다. 혼자 있지만 긍정적인 방식으로 즐기며 자기 뜻대로 살아가는 생활의 이점 과 혜택들을 누릴 수도 있다. 즉 안전한 생활을 마련해 놓고 풍요로운 삶을 나누기에 좋은 사람을 즐거운 마음으로 기다릴 수 있는 사람일 때가 많다. 내적인 평화나 만족은 우리 자신에게서 온다는 것을 일깨워준 다. 즉 자신의 세계를 다른 사람으로 채우기 전에 자신의 환 경과 조화를 이루는 법을 배울 수 있는 생활의 공간이 필요할 수도 있다.", oneCard: "" },
    "pentacles-10": { description: "강아지를 쓰다듬고 있는 부유해 보이는 한 노인 앞에 한 쌍 의 남녀가 정겹게 이야기를 나누고 있다. 그들 뒤에는 한 소 년이 흰 강아지의 꼬리를 장난스럽게 만지고 있다. 그들의 머 리 위 커다란 돌 아치에는 성 그림이 그려진 깃발과 저울이 그려진 깃발이 걸려 있는데, 이것은 노인이 지금까지 공명정 대하고 성공적인 삶을 살았음을 보여 준다. 노인이 입고 있는 옷에는 풍요로움의 상징인 포도 넝쿨이 그 려져 있어, 그가 매우 부유하다는 것을 보여 준다. 그리고 두 마리의 흰 강아지가 노인을 향해 친근감을 표시하고 있는 것 은 그를 믿고 따르는 충성스런 이들이 많다는 것을 나타낸다. 그리고 이 카드는 가족 안에서의 부유함과 관대함의 결과인 경우가 많으며, 즉 오랫동안 성실히 일한 결과로 얻게 된 돈 이나 재산을 상징하기도 한다. 또한 가족으로부터 금전적인 도움이 있다는 것을 의미하며 정신적인 지원일 수도 있다. 결국 10번 카드는 금전적 상황이 좋아질 것을 의미하며 일을 통해서라면 새로운 일자리나 큰 사업기회, 놀랄 만한 승진의 가능성이 있다. 이 카드는 물질적, 감정적인 풍요로움과 함께 행복하고 안정적인 단란한 가정을 나타내는 카드이다.", oneCard: "" },
    "pentacles-page": { description: "드넓은 초원 한복판에 초록색 옷과 붉은 모자를 쓴 젊은이가 두 손에 펜타클을 받쳐 들고 서 있다. 뒤에 보이는 나무와 풀 밭은 풍성해 보이고, 그 뒤로 보이는 산은 평탄하다. 이것으로 우리는 그가 풍요로운 삶을 살고 있고, 험난한 시련을 겪지 않았던 것을 알 수 있다. 그가 입고 있는 초록색 옷은 성장과 풍요를 나타내고, 쓰고 있는 붉은 모자는 에너지와 열정을 나 타낸다. 펜타클은 물질적인 풍요로움을 나타내는데, 그것을 받들고 있는 그는 실용적이고 물질적인 것을 동경하고 있는 것이다. 펜타클 소년은 사려 깊고 신중한 사람이며, 흙의 성격인 믿 을 수 있고 솜씨가 좋으며 물질세계와 화합하고 돈의 가치를 인식하는 특성을 지닌다. 즉 새로운 프로젝트나 사업의 시작을 나타내며 돈을 버는 것 과 관련되 아이디어에 대해 좋은 징조이다. 틀에 박힌 생활의 사람에게 직업을 바꿀 필요가 있는 경우, 특히 자신에게 맞는 직업을 찾는 사람에게 나오는 경우도 있다.", oneCard: "" },
    "pentacles-knight": { description: "펜타클 기사는 다른 기사들에 비해 유일하게 정지해 있는 기 사이며, 가만히 앉아서 오른손에 들고 있는 펜타클을 바라보 고 있다. 펜타클은 흙의 성향을 받아 꾸준하고 참을성이 많고 계획대로 실천하는 철저하고 신용할 수 있는 사람을 의미한 다. 이 사람은 인내심이 강하고 확고부동하며 목적에 집중하는 사람이기도 하다. 부정적인 면은 단조롭고, 완고하며, 다른 기 사들에 비해 행동이 느린 사람이다. 보통은 자발적으로 일을 실행하기보다는 행동하도록 밀어 붙여야 한다. 펜타클의 기사 는 진전되어 온 상황이 진척이 없거나 교착상태에 빠졌더라도 끝내는 긍정적인 결과를 낳을 것이라는 점을 의미한다. 결과 가 빨리 이루어지지는 않지만 오래 지속되며 기다릴만한 가치 가 있다. 그래도 그의 깊은 곳에는 야망이 있고 정열도 있다. 철저하 고 조심성도 많아 세부적인 것들까지 세세히 챙기는 꼼꼼한 면이 있지만, 큰 일 앞에서는 용감하게 맞서 돌파하지 못하는 소심한 면도 있다. 한번 결심한 일은 무슨 일이 있건 해 나가 야 하고, 지킬 수 없는 말은 함부로 내뱉지 않는다. 사람들과 쉽게 친해지지는 못하지만, 한번 인연을 맺으면 꾸준히 오래 가는 믿을 수 있는 사람들이 많다. 또한 대개 정이 많다. 이런 사람은 큰 그룹을 이끄는 리더보다는, 그 안에서 실무를 맡는 믿을 수 있는 참모가 적격이다.", oneCard: "" },
    "pentacles-queen": { description: "장미 덩굴의 그늘 아래에 여왕이 옥좌에 앉아 무릎위에 펜타 클을 바라보고 있다. 자신이 이룬 부와 성공을 바라보는 모습 이다. 그녀가 앉아 있는 옥좌에는 천사, 염소와 과일이 새겨져 있다. 번식을 상징하는 염소의 머리 바로 아래쪽 옆에 한 아 이가 어머니의 태중에 있는 포즈로 있는 그림이 새겨져 있다. 그녀를 둘러싼 주위의 비옥한 땅에는 비옥과 다산의 상징인 토끼 한 마리가 숲속에서 뛰어 나오고 있다. 풍요와 다산의 상징, 그리고 태중의 아이라는 상징을 통해서 임신이라는 키 워드를 얻을 수 있다. 또한 여왕이 앉아 있는 옥좌는 꽃과 식 물이 무성하게 우거진 들판 위에 놓여 있는데 이는 풍부함과 비옥함을 나타내며 흙의 원소와 연관을 보여준다. 그녀는 넉넉하고 편안하며 풍요로운 땅에서 삶을 즐기고 있 는 사람으로 안정되고 건설적이고 감사하는 사람이다. 자연친 화적이며 물질세계를 잘하는 메이저 카드의 여황제와 비슷하 다. 펜타클의 여왕은 정원에 있는 모든 것이 장밋빛이라는 것 을 나타내며 금전적으로나 정서적으로 안정된 상태를 나타내 거나 앞으로 더 안정되고 상황들이 좋아질 것이라는 것을 가 리킨다.", oneCard: "" },
    "pentacles-king": { description: "부유한 왕이 덩굴에 주렁주렁 열려 있는 포도송이들로 장식 된 옥좌에 앉아 있고, 그의 오른손에는 왕홀이 들려져 있고, 왼손에는 펜타클을 잡고 있다. 이는 그가 왕의 존엄성을 가지 고 있을 뿐만 아니라 물질적으로 풍요롭고 안정된 상태임을 나타낸다. 이 카드는 펜타클의 여왕처럼 비옥함과 풍부함의 상징이다. 이 사람은 믿을 수 있고, 책임감이 있고, 부유하며, 안정적인 사람이다. 이 왕은 삶에 대하여 꾸준하고 실용적으로 접근하 고, 타당한 이유를 확인한 뒤에 일을 시작하며, 충동적으로 행 동하지 않는다. 엄숙함과 권위를 발산하는 사람이며, 부모나 성공적인 사업가, 은행가, 자선을 베푸는 사람일 수도 있다. 그러나 다른 한편으로는 자신이 가진 것을 주변 인물들에게 자랑하며 물질적인 장식물을 즐기는 사람이기도 하다. 펜타클 왕은 안정을 유지하려는 습성이 있어, 더 발전할 수 있는 여러 가지 능력을 갖고 있음에도 불구하고, 안주하여 더 이상의 발전 없이 현재에 눌러 앉아 버리는 경우도 있다. 때 로는 육체적인 쾌락과 즐거움에 대한 추구와 지배가 두드러 져, 그러한 생활이 계속되다 보면 무기력이나 권태에 빠질 위 험도 있는 카드이다.", oneCard: "" },
  };

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
    const selectedDetails = selectedCard ? TAROT_SOURCE_DETAILS[selectedCard.id] : null;
    const searchKeyword = tarotSearch.trim().toLowerCase();

    const matchesTarotSearch = (card: TarotCard) => {
      if (!searchKeyword) return true;

      const detail = TAROT_SOURCE_DETAILS[card.id];
      const searchable = [
        card.number,
        card.name,
        card.korean,
        card.group,
        card.upright,
        card.reversed,
        detail?.description || "",
        detail?.oneCard || "",
      ]
        .join(" ")
        .toLowerCase();

      return searchable.includes(searchKeyword);
    };

    const filteredCount = TAROT_CARDS.filter(matchesTarotSearch).length;

    return (
      <section className="rounded-3xl border border-[#ead8c4] bg-[#fffaf3] p-5 shadow-inner">
        <div className="mb-5">
          <h2 className={`${FONT.sectionTitle} ${WEIGHT.sectionTitle} ${COLOR.sectionTitle}`}>
            타로 카드 해석
          </h2>

          <div className="mt-4">
            <input
              type="text"
              value={tarotSearch}
              onChange={(event) => setTarotSearch(event.target.value)}
              placeholder="카드명 또는 해석 검색 (예: 바보, Fool, 취직, 사랑, 금전)"
              className="w-full rounded-2xl border-2 border-[#d7c4ad] bg-white px-5 py-4 text-3xl font-bold text-black outline-none placeholder:text-zinc-400 focus:border-[#6b3f24]"
            />
            <div className="mt-2 text-xl font-bold text-[#9a7657]">
              {searchKeyword ? `검색 결과 ${filteredCount}장` : "78장 전체 카드"}
            </div>
          </div>

          <p className={`mt-3 ${FONT.body} ${WEIGHT.body} ${COLOR.body}`}>
            카드를 클릭하면 카드 설명, 정방향·역방향 해석, 원 카드 리딩법을 확인할 수 있습니다.
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

            <div className="mt-5 rounded-2xl border border-[#ead8c4] bg-[#fffaf3] p-5">
              <div className="text-3xl font-bold text-[#6b3f24]">카드 설명</div>
              <div className="mt-3 whitespace-pre-wrap text-2xl font-semibold leading-relaxed text-black">
                {selectedDetails?.description || "교안의 카드 설명이 없습니다."}
              </div>
            </div>

            <div className="mt-4 grid grid-cols-2 gap-4">
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

            <div className="mt-4 rounded-2xl border border-[#d7c4ad] bg-white p-5">
              <div className="text-3xl font-bold text-[#6b3f24]">원 카드 리딩법</div>
              {selectedDetails?.oneCard ? (
                <div className="mt-3 whitespace-pre-wrap text-2xl font-semibold leading-relaxed text-black">
                  {selectedDetails.oneCard}
                </div>
              ) : (
                <div className="mt-3 text-2xl font-semibold leading-relaxed text-zinc-600">
                  이 교안의 마이너 카드 부분에는 별도의 원 카드 Q&amp;A 예시가 수록되어 있지 않습니다. 위 카드 설명과 정·역방향 의미를 중심으로 한 장 리딩에 적용합니다.
                </div>
              )}
            </div>
          </div>
        )}

        <div className="space-y-7">
          {TAROT_GROUPS.map((group) => {
            const cards = TAROT_CARDS.filter(
              (card) => card.group === group && matchesTarotSearch(card),
            );

            if (cards.length === 0) return null;

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
                        onClick={() => {
                          setSelectedTarotCard(card.id);
                          requestAnimationFrame(() => window.scrollTo({ top: 0, behavior: "smooth" }));
                        }}
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

          {filteredCount === 0 && (
            <div className="rounded-2xl bg-white p-8 text-center text-3xl font-bold text-zinc-400">
              검색 결과가 없습니다.
            </div>
          )}
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
