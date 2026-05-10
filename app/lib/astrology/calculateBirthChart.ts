

export type ChartInput = {
  name?: string;
  gender?: string;
  birthDate?: string;
  birthTime?: string;
  birthLocation?: string;

  year?: number;
  month?: number;
  day?: number;
  hour?: number;
  minute?: number;
  latitude?: number;
  longitude?: number;
  timezone?: string;
};

export type ChartObject = {
  name: string;
  longitude: number;
  speed: number;
  retrograde: boolean;
  sign: string;
  degree: number;
  minute: number;
  house?: number;
};

export type House = {
  house: number;
  longitude: number;
  sign: string;
  degree: number;
  minute: number;
};

export type Aspect = {
  from: string;
  aspect: string;
  to: string;
  orbFloat: number;
  orb: string;
  motion: "Applying" | "Separating";
};

export type BirthChart = {
  planets: Record<string, ChartObject>;
  points: Record<string, ChartObject>;
  houses: House[];
  aspects: Aspect[];
  aiText: string;
};

const SIGNS = [
  "Aries",
  "Taurus",
  "Gemini",
  "Cancer",
  "Leo",
  "Virgo",
  "Libra",
  "Scorpio",
  "Sagittarius",
  "Capricorn",
  "Aquarius",
  "Pisces",
];

const ASPECTS: Record<string, number> = {
  Conjunction: 0,
  Sextile: 60,
  Square: 90,
  Trine: 120,
  Opposition: 180,
};

const PLANET_ORDER = [
  "Sun",
  "Moon",
  "Mercury",
  "Venus",
  "Mars",
  "Jupiter",
  "Saturn",
  "Uranus",
  "Neptune",
  "Pluto",
  "North Node",
  "Lilith",
];

const POINT_PRINT_ORDER = ["Fortune", "Vertex", "ASC", "MC"];

const ASPECT_OBJECT_ORDER = [
  "Sun",
  "Moon",
  "Mercury",
  "Venus",
  "Mars",
  "Jupiter",
  "Saturn",
  "Uranus",
  "Neptune",
  "Pluto",
  "ASC",
  "DSC",
  "MC",
  "IC",
  "North Node",
  "Lilith",
  "Fortune",
  "Vertex",
];

function normalizeAngle(angle: number): number {
  return ((angle % 360) + 360) % 360;
}

function angleDiff(a: number, b: number): number {
  const diff = Math.abs(a - b) % 360;
  return Math.min(diff, 360 - diff);
}

function signDegree(longitude: number) {
  const lon = normalizeAngle(longitude);
  let signIndex = Math.floor(lon / 30);
  const degreeInSign = lon % 30;

  let degree = Math.floor(degreeInSign);
  let minute = Math.round((degreeInSign - degree) * 60);

  if (minute === 60) {
    degree += 1;
    minute = 0;
  }

  if (degree === 30) {
    degree = 0;
    signIndex = (signIndex + 1) % 12;
  }

  return {
    sign: SIGNS[signIndex],
    degree,
    minute,
  };
}

function formatDegMin(value: number): string {
  let degree = Math.floor(value);
  let minute = Math.round((value - degree) * 60);

  if (minute === 60) {
    degree += 1;
    minute = 0;
  }

  return `${degree}°${String(minute).padStart(2, "0")}’`;
}

function ordinal(n: number): string {
  if (n % 100 >= 10 && n % 100 <= 20) return `${n}th`;

  const suffix: Record<number, string> = {
    1: "st",
    2: "nd",
    3: "rd",
  };

  return `${n}${suffix[n % 10] ?? "th"}`;
}

function isBetweenCircular(x: number, start: number, end: number): boolean {
  const value = normalizeAngle(x);
  const s = normalizeAngle(start);
  const e = normalizeAngle(end);

  if (s <= e) return value >= s && value < e;

  return value >= s || value < e;
}

function findHouse(longitude: number, houses: House[]): number | undefined {
  for (let i = 0; i < 12; i++) {
    const start = houses[i].longitude;
    const end = houses[(i + 1) % 12].longitude;

    if (isBetweenCircular(longitude, start, end)) {
      return i + 1;
    }
  }

  return undefined;
}

function isDayChart(sunHouse?: number): boolean {
  return sunHouse !== undefined && [7, 8, 9, 10, 11, 12].includes(sunHouse);
}

function calculatePartOfFortune(
  sunLon: number,
  moonLon: number,
  ascLon: number,
  dayChart: boolean
): number {
  if (dayChart) return normalizeAngle(ascLon + moonLon - sunLon);

  return normalizeAngle(ascLon + sunLon - moonLon);
}

function makeObject(
  name: string,
  longitude: number,
  speed = 0,
  retrograde = false,
  houses?: House[]
): ChartObject {
  const sd = signDegree(longitude);

  const obj: ChartObject = {
    name,
    longitude: normalizeAngle(longitude),
    speed,
    retrograde,
    sign: sd.sign,
    degree: sd.degree,
    minute: sd.minute,
  };

  if (houses) {
    obj.house = findHouse(longitude, houses);
  }

  return obj;
}

function getOrbLimit(aName: string, bName: string): number {
  const luminaries = new Set(["Sun", "Moon"]);
  const angles = new Set(["ASC", "DSC", "MC", "IC"]);
  const points = new Set(["North Node", "Lilith", "Fortune", "Vertex"]);

  if (angles.has(aName) || angles.has(bName)) return 6;
  if (luminaries.has(aName) || luminaries.has(bName)) return 10;
  if (points.has(aName) || points.has(bName)) return 6;

  return 8;
}

function shouldSkipAspect(aName: string, bName: string): boolean {
  const key = [aName, bName].sort().join("|");

  const skipPairs = new Set([
    "ASC|DSC",
    "IC|MC",
    "Jupiter|Venus",
    "Mars|Venus",
    "MC|Pluto",
    "Fortune|North Node",
    "North Node|Vertex",
    "Fortune|Vertex",
  ]);

  return skipPairs.has(key);
}

function aspectMotion(
  objA: ChartObject,
  objB: ChartObject,
  aspectAngle: number
): "Applying" | "Separating" {
  const nowOrb = Math.abs(angleDiff(objA.longitude, objB.longitude) - aspectAngle);

  const futureA = normalizeAngle(objA.longitude + objA.speed / 24);
  const futureB = normalizeAngle(objB.longitude + objB.speed / 24);

  const futureOrb = Math.abs(angleDiff(futureA, futureB) - aspectAngle);

  return futureOrb < nowOrb ? "Applying" : "Separating";
}

function reorderAspectPair(
  aName: string,
  bName: string,
  a: ChartObject,
  b: ChartObject
) {
  const angles = new Set(["ASC", "DSC", "MC", "IC"]);
  const priorityPoints = new Set(["North Node", "Lilith", "Fortune", "Vertex"]);

  if (angles.has(aName) || angles.has(bName)) {
    if (angles.has(bName) && !angles.has(aName)) {
      return { fromName: bName, toName: aName, fromObj: b, toObj: a };
    }
  }

  if (priorityPoints.has(aName) || priorityPoints.has(bName)) {
    if (priorityPoints.has(bName) && !priorityPoints.has(aName)) {
      return { fromName: bName, toName: aName, fromObj: b, toObj: a };
    }
  }

  return { fromName: aName, toName: bName, fromObj: a, toObj: b };
}

function aspectSortKey(aspect: Aspect): [number, number, number] {
  const fromIndex = ASPECT_OBJECT_ORDER.indexOf(aspect.from);
  const toIndex = ASPECT_OBJECT_ORDER.indexOf(aspect.to);

  return [
    fromIndex === -1 ? 99 : fromIndex,
    toIndex === -1 ? 99 : toIndex,
    aspect.orbFloat,
  ];
}

function calculateAspects(objects: Record<string, ChartObject>): Aspect[] {
  const names = Object.keys(objects);
  const aspects: Aspect[] = [];

  for (let i = 0; i < names.length; i++) {
    for (let j = i + 1; j < names.length; j++) {
      const aName = names[i];
      const bName = names[j];

      if (shouldSkipAspect(aName, bName)) continue;

      const a = objects[aName];
      const b = objects[bName];

      const diff = angleDiff(a.longitude, b.longitude);

      for (const [aspectName, aspectAngle] of Object.entries(ASPECTS)) {
        const orb = Math.abs(diff - aspectAngle);
        const orbLimit = getOrbLimit(aName, bName);

        if (orb <= orbLimit) {
          const ordered = reorderAspectPair(aName, bName, a, b);

          aspects.push({
            from: ordered.fromName,
            aspect: aspectName,
            to: ordered.toName,
            orbFloat: orb,
            orb: formatDegMin(orb),
            motion: aspectMotion(ordered.fromObj, ordered.toObj, aspectAngle),
          });
        }
      }
    }
  }

  return aspects.sort((a, b) => {
    const ak = aspectSortKey(a);
    const bk = aspectSortKey(b);

    return ak[0] - bk[0] || ak[1] - bk[1] || ak[2] - bk[2];
  });
}

function planetLine(p: ChartObject): string {
  const retro = p.retrograde ? ", Retrograde" : "";
  const house = p.house ? `, in ${ordinal(p.house)} House` : "";

  return `${p.name} in ${p.sign} ${p.degree}°${String(p.minute).padStart(
    2,
    "0"
  )}’${retro}${house}`;
}

function houseLine(h: House): string {
  return `${ordinal(h.house)} House in ${h.sign} ${h.degree}°${String(
    h.minute
  ).padStart(2, "0")}’`;
}

function buildAiText(chart: Omit<BirthChart, "aiText">): string {
  const lines: string[] = [];

  for (const name of PLANET_ORDER) {
    if (chart.planets[name]) lines.push(planetLine(chart.planets[name]));
  }

  for (const name of POINT_PRINT_ORDER) {
    if (chart.points[name]) lines.push(planetLine(chart.points[name]));
  }

  lines.push("");

  for (const house of chart.houses) {
    lines.push(houseLine(house));
  }

  lines.push("");

  for (const aspect of chart.aspects) {
    lines.push(
      `${aspect.from} ${aspect.aspect} ${aspect.to} ` +
        `(Orb: ${aspect.orb}, ${aspect.motion})`
    );
  }

  return lines.join("\n");
}

export function buildBirthChartFromRaw(raw: {
  planets: Record<
    string,
    {
      longitude: number;
      speed: number;
      retrograde: boolean;
    }
  >;
  houses: {
    cusps: number[];
    asc: number;
    mc: number;
    vertex: number;
  };
}): BirthChart {
  const houses: House[] = raw.houses.cusps.map((longitude, index) => {
    const sd = signDegree(longitude);

    return {
      house: index + 1,
      longitude: normalizeAngle(longitude),
      sign: sd.sign,
      degree: sd.degree,
      minute: sd.minute,
    };
  });

  const planets: Record<string, ChartObject> = {};

  for (const [name, value] of Object.entries(raw.planets)) {
    planets[name] = makeObject(
      name,
      value.longitude,
      value.speed,
      value.retrograde,
      houses
    );
  }

  const ascLon = normalizeAngle(raw.houses.asc);
  const mcLon = normalizeAngle(raw.houses.mc);
  const dscLon = normalizeAngle(ascLon + 180);
  const icLon = normalizeAngle(mcLon + 180);
  const vertexLon = normalizeAngle(raw.houses.vertex);

  const sunHouse = planets["Sun"]?.house;
  const dayChart = isDayChart(sunHouse);

  const fortuneLon = calculatePartOfFortune(
    planets["Sun"].longitude,
    planets["Moon"].longitude,
    ascLon,
    dayChart
  );

  const points: Record<string, ChartObject> = {
    ASC: makeObject("ASC", ascLon),
    DSC: makeObject("DSC", dscLon),
    MC: makeObject("MC", mcLon),
    IC: makeObject("IC", icLon),
    Fortune: makeObject("Fortune", fortuneLon, 0, false, houses),
    Vertex: makeObject("Vertex", vertexLon, 0, false, houses),
  };

  const allObjects = {
    ...planets,
    ...points,
  };

  const aspects = calculateAspects(allObjects);

  const chartWithoutText = {
    planets,
    points,
    houses,
    aspects,
  };

  return {
    ...chartWithoutText,
    aiText: buildAiText(chartWithoutText),
  };
}

export async function calculateBirthChart(
  input: ChartInput
): Promise<BirthChart> {
  console.log("calculateBirthChart input:", input);

  const raw = {
    planets: {
      Sun: {
        longitude: 246.9833,
        speed: 1,
        retrograde: false,
      },
      Moon: {
        longitude: 342.2,
        speed: 13,
        retrograde: false,
      },
      Mercury: {
        longitude: 250.6,
        speed: 1,
        retrograde: false,
      },
      Venus: {
        longitude: 272.65,
        speed: 1,
        retrograde: false,
      },
      Mars: {
        longitude: 269.32,
        speed: 0.7,
        retrograde: false,
      },
      Jupiter: {
        longitude: 262.15,
        speed: 0.2,
        retrograde: false,
      },
      Saturn: {
        longitude: 348.05,
        speed: 0.1,
        retrograde: false,
      },
      Uranus: {
        longitude: 297.73,
        speed: 0.05,
        retrograde: false,
      },
      Neptune: {
        longitude: 293.6,
        speed: 0.03,
        retrograde: false,
      },
      Pluto: {
        longitude: 240.75,
        speed: 0.02,
        retrograde: false,
      },
      "North Node": {
        longitude: 204.15,
        speed: -0.05,
        retrograde: true,
      },
      Lilith: {
        longitude: 96.87,
        speed: 0.1,
        retrograde: false,
      },
    },

    houses: {
      cusps: [
        175.3333,
        204.15,
        233.5,
        263.2,
        293.1,
        323.0,
        355.3333,
        24.15,
        53.5,
        83.2,
        113.1,
        143.0,
      ],
      asc: 175.3333,
      mc: 83.2,
      vertex: 344.38,
    },
  };

  return buildBirthChartFromRaw(raw);
}