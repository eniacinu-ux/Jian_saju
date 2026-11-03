"use client";
export type Alert = { id: string; kind: "chat" | "payment"; createdAt: number };
let enabledAt = 0;
const seen = new Set<string>();
export function enableAlerts() {
  if (!("speechSynthesis" in window)) throw new Error("이 브라우저는 음성 알림을 지원하지 않습니다.");
  enabledAt = Date.now();
  speak("음성 알림을 켰습니다");
}
export function disableAlerts() { enabledAt = 0; window.speechSynthesis?.cancel(); }
function speak(text: string) {
  const voice = new SpeechSynthesisUtterance(text);
  voice.lang = "ko-KR";
  voice.rate = 1;
  const korean = window.speechSynthesis.getVoices().find(v => v.lang.startsWith("ko"));
  if (korean) voice.voice = korean;
  window.speechSynthesis.speak(voice);
}
export function receiveAlerts(alerts: Alert[]) {
  for (const alert of alerts) {
    if (seen.has(alert.id)) continue;
    seen.add(alert.id);
    if (!enabledAt || alert.createdAt < enabledAt || alert.createdAt < Date.now() - 120000) continue;
    if (alert.kind === "chat") speak("채팅창 확인바랍니다");
    if (alert.kind === "payment") speak("입금 확인 바랍니다");
  }
  if (seen.size > 500) { const recent = [...seen].slice(-100); seen.clear(); recent.forEach(id => seen.add(id)); }
}
