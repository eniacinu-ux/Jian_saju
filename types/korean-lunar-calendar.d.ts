declare module "korean-lunar-calendar" {
  type LunarDate = {
    year: number;
    month: number;
    day: number;
    intercalation: boolean;
  };

  export default class KoreanLunarCalendar {
    setSolarDate(year: number, month: number, day: number): boolean;
    getLunarCalendar(): LunarDate;
  }
}