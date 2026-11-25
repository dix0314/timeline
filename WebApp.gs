// =============================================================
// 모바일 웹앱 v32 — DATA / MEMO 시트 기반 · 화면은 GitHub(앱), 여기는 데이터 서버(API)
//  · 일정: DATA 시트  (B 날짜 | C 시작시간 | D 진행시간·종류 | E 일정명)  5행부터
//          D 예) "2h" · "FROM" · "UNTIL" · "2h MAJOR"(중요) · "1h TBD"(미정)
//  · 메모: MEMO 시트  (B 날짜 | C 번호(1~3) | D 완료("DONE") | E 내용) 5행부터
//  · 핵심 일정: TIMELINE 아래 표 (기존 그대로)
//  · 날짜 탭: 이번 주 일요일부터 5주 (시트를 찾지 않고 날짜로 계산)
// Code.gs 의 getTodayZero(), getTimelineLastRow(), TL_* 상수를 사용합니다.
//  · 변경 번호(RD_VER): 앱에서 저장하거나 시트를 직접 고치면(Code.gs onEdit) 올라감
//    → 자동 갱신 때 번호가 같으면 시트를 읽지 않고 "변경 없음"만 돌려줌
//  · 지난 일정·메모 정리는 하루 한 번 (그날 첫 저장 때)
//  · 세방고딕 2.0 글꼴은 GitHub 저장소에 파일로 올려서 씀 (서버는 관여하지 않음)
//  · 앱 ↔ 서버: 앱이 doPost 로 { key, fn, args } 를 보내면 비밀 키 확인 후 해당 함수 결과를 JSON 으로 돌려줌
//    배포: 실행 = 나 · 액세스 = 모든 사용자 (키가 맞아야만 동작) · 키는 시트 메뉴 [RUNDOWN] → 앱 키 보기
// =============================================================
const WEB_TITLE = "TIMELINE";
const DATA_SHEET = "DATA";
const MEMO_SHEET = "MEMO";
const WEB_START_ROW = 5;      // 5행부터 데이터 (그 위는 제목·머리글)
// DATA 시트 B열(날짜) 표시 서식: "1124. 화" — 값은 날짜 그대로, 보이는 모양만 (바꾸고 싶으면 여기만 고치면 됨)
//  · MM = 월, dd = 일, ddd = 요일 (시트 언어가 한국어면 "화")
const DATA_DATE_FORMAT = 'MMdd". "ddd';
const WEB_DAYS = 3660;    // 오늘부터 약 10년치 일정을 불러옴 (달력을 옆으로 넘기면 끝없이 이어지므로 사실상 전부)
const MEMO_SLOTS = 3;
const WEB_DEFAULT_DURATIONS = ["FROM", "30m", "60m", "90m", "2h", "3h", "4h", "5h", "6h", "7h", "8h", "9h", "10h", "11h", "12h", "UNTIL"];

// 문서 잠금 안에서 실행 (저장 함수 공통)
function webLocked_(fn) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// -------------------------------------------------------------
// 변경 번호: 시트 내용이 바뀔 때마다 올라감
//  · 날짜도 같이 붙임 → 자정이 지나면 (D-DAY 수식 등) 무조건 새로 읽음
// -------------------------------------------------------------
function webVer_(tz) {
  const v = PropertiesService.getDocumentProperties().getProperty("RD_VER") || "0";
  return v + "|" + webKey_(getTodayZero(), tz);
}
// 저장 후 / 시트 직접 수정 후 호출 (Code.gs 의 onEdit 에서도 씀)
function webBumpVer_() {
  // 시각 + 무작위 글자 → 같은 순간에 두 번 바뀌어도 번호가 겹치지 않음
  PropertiesService.getDocumentProperties().setProperty("RD_VER", Date.now().toString(36) + Math.random().toString(36).slice(2, 7));
}

// -------------------------------------------------------------
// 전체 데이터 (웹앱이 호출)
//  · since: 앱이 마지막으로 받은 변경 번호. 지금과 같으면 시트를 읽지 않고 { same:true } 만 돌려줌
// -------------------------------------------------------------
function getAll(since) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const ver = webVer_(tz);                                   // 읽기 전에 번호부터 (읽는 도중 바뀌면 다음 갱신 때 다시 읽힘)
  if (since && since === ver) return { same: true, ver: ver };
  const today = getTodayZero();

  // 어제(전날에서 넘어오는 일정용) ~ WEB_DAYS 일 뒤
  const from = new Date(today); from.setDate(from.getDate() - 1);
  const to = new Date(today); to.setDate(to.getDate() + WEB_DAYS);
  const fromKey = webKey_(from, tz), toKey = webKey_(to, tz);

  const data = ss.getSheetByName(DATA_SHEET);
  const sched = [];
  if (data) {
    webReadRows_(data, 4).forEach(function (r) {
      const date = webDateKey_(r.vals[0], r.disp[0], tz);
      const name = String(r.disp[3]).trim();
      if (!date || !name || date < fromKey || date > toKey) return;
      const start = webParseTime_(r.disp[1]);
      const dm = webSplitDur_(r.disp[2]);                    // "2h MAJOR" → 진행시간 "2h" + 중요
      const dur = dm.dur || "30m";
      const kind = webKind_(dur);
      sched.push({
        id: r.row,
        date: date,
        name: name,
        start: start == null ? 0 : start,
        noTime: start == null,
        dur: dur,
        len: webDurMin_(dur),
        kind: kind,
        mark: kind === "norm" ? dm.mark : ""                 // 중요(imp) / 미정(tbd) / 빈칸
      });
    });
  }

  const memo = ss.getSheetByName(MEMO_SHEET);
  const memos = [];
  if (memo) {
    webReadRows_(memo, 4).forEach(function (r) {
      const date = webDateKey_(r.vals[0], r.disp[0], tz);
      const slot = parseInt(r.disp[1], 10);
      const done = /^(DONE|TRUE)$/i.test(String(r.disp[2]).trim());    // D열: 완료
      const text = String(r.vals[3] == null ? "" : r.vals[3]);           // E열: 내용
      if (!date || !(slot >= 1 && slot <= MEMO_SLOTS) || date < fromKey || date > toKey) return;
      memos.push({ id: r.row, date: date, slot: slot, text: text, done: done });
    });
  }

  const tl = ss.getSheetByName("TIMELINE");
  return {
    ver: ver,
    sched: sched,
    memos: memos,
    holidays: webGetHolidays_(),
    lvdays: webGetLvDays_(),
    events: tl ? webReadTimelineTable_(tl) : [],
    durations: webDurationsCached_(data)
  };
}

// -------------------------------------------------------------
// 일정 추가 / 수정 / 삭제  →  DATA 시트
// req = { mode: "add"|"edit"|"delete", id, old:{date,name,start,dur,mark}, date:"yyyy-MM-dd", name, start(분), dur, mark:""|"imp"|"tbd" }
// -------------------------------------------------------------
function saveEvent(req) {
  webLocked_(function () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const sheet = ss.getSheetByName(DATA_SHEET);
    if (!sheet) throw new Error("DATA 시트가 없어요.");

    let row = Number(req.id);
    const now = Date.now();
    let origAbs = null;                                        // 원래 일정의 시각 (지난 시간 규칙용)
    if (req.mode === "edit" || req.mode === "delete") {
      if (!(row >= WEB_START_ROW)) throw new Error("잘못된 일정이에요.");
      const one = sheet.getRange(row, 2, 1, 4);              // 한 번에 읽기 (B~E)
      const disp = one.getDisplayValues()[0];
      const vals = one.getValues()[0];
      const o = req.old || {};
      const dm = webSplitDur_(disp[2]);
      const oldDur = dm.dur || "30m";
      const oKey = webDateKey_(vals[0], disp[0], tz), oSt = webParseTime_(disp[1]);
      const same = oKey === o.date &&
        String(disp[3]).trim() === o.name &&
        (oSt == null ? 0 : oSt) === Number(o.start) &&
        oldDur === o.dur &&
        (webKind_(oldDur) === "norm" ? dm.mark : "") === (o.mark || "");
      if (!same) throw new Error("시트 내용이 바뀌었어요. 새로고침 후 다시 시도해 주세요.");
      // 이미 끝난 일정은 수정 · 삭제 불가 (진행 중인 일정은 가능)
      origAbs = webKeyToDate_(oKey).getTime() + (oSt == null ? 0 : oSt) * 60000;
      if (webIsPast_(oKey, origAbs, oldDur, tz, now)) throw new Error("종료된 일정입니다.");
    }

    if (req.mode === "delete") {
      sheet.getRange(row, 2, 1, 4).clearContent();
    } else {
      const name = String(req.name || "").trim();
      if (!name) throw new Error("일정 이름을 입력해 주세요.");
      const date = webKeyToDate_(req.date);
      if (!date) throw new Error("날짜를 확인해 주세요.");
      const start = Number(req.start);
      if (!(start >= 0 && start < 1440)) throw new Error("시작 시간을 확인해 주세요.");
      const dur = webSplitDur_(req.dur).dur || "30m";
      // 중요 / 미정은 기본 일정(BASIC)에만 붙음 → D열에 "2h MAJOR" / "1h TBD" 처럼 같이 적음
      const mark = webKind_(dur) === "norm" ? webMarkText_(req.mark) : "";
      const dCell = mark ? dur + " " + mark : dur;

      // 지난 시간 규칙
      //  · 새 일정: 지난 시간이 조금이라도 들어가면 저장 안 함
      //  · 진행 중인 일정 수정: 원래 시작보다 더 과거로는 못 옮기고, 끝은 지금 이후여야 함
      const at = date.getTime() + start * 60000, kind = webKind_(dur);
      const floor = req.mode === "edit" && origAbs != null ? Math.min(origAbs, now) : now;
      if (at < (kind === "out" ? now : floor) - WEB_PAST_GRACE) throw new Error("지난 시간입니다.");
      if (kind === "norm" && at + webDurMin_(dur) * 60000 <= now - WEB_PAST_GRACE) throw new Error("끝나는 시간이 이미 지났어요.");

      // 1분이라도 겹치는 일정이 있으면 저장하지 않음 (시트는 한 번만 읽어서 겹침 검사 · 빈 줄 찾기에 같이 씀)
      const rows = webReadRows_(sheet, 4);
      const hit = webFindOverlap_(rows, tz, date, start, dur, req.mode === "edit" ? row : 0);
      if (hit) throw new Error("겹치는 일정: " + hit);

      if (req.mode === "add") row = webFreeRowFrom_(sheet, rows);
      // 칸 서식은 그대로 둠 (예: 맞춤 서식 " "@ 이 풀리지 않게)
      //  · 시간 칸: 이미 시간 서식이면 손대지 않음 (서식이 없을 때만 HH:mm)
      //  · 진행시간 · 일정명: 글자로 저장한 뒤 원래 서식으로 되돌림
      webTimeFmt_(sheet.getRange(row, 3));
      sheet.getRange(row, 2, 1, 2).setValues([[date, webHM_(start)]]);
      webSetText_(sheet.getRange(row, 4, 1, 2), [[dCell, name]]);
    }
    // 지난 일정·메모 삭제는 하루 한 번, 그 외에는 이 시트 순서만 확인
    webTidyMaybe_(ss, tz, DATA_SHEET);
    webFormatDates_(sheet);                                    // 정리·정렬이 끝난 뒤 날짜 칸 서식을 한 번에 맞춤
    webBumpVer_();
    SpreadsheetApp.flush();
  });
  return getAll();
}

// -------------------------------------------------------------
// 메모 저장 →  MEMO 시트  (빈 내용이면 지움)
// req = { date:"yyyy-MM-dd", slot:1~3, text }
// -------------------------------------------------------------
function saveMemo(req) {
  webLocked_(function () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const sheet = ss.getSheetByName(MEMO_SHEET);
    if (!sheet) throw new Error("MEMO 시트가 없어요.");
    const slot = Number(req.slot);
    if (!(slot >= 1 && slot <= MEMO_SLOTS)) throw new Error("잘못된 메모 칸이에요.");
    const date = webKeyToDate_(req.date);
    if (!date) throw new Error("날짜를 확인해 주세요.");
    const text = String(req.text || "").replace(/\s+$/, "");

    const rows = webReadRows_(sheet, 4);
    const found = webFindMemo_(rows, tz, req.date, slot, false);

    if (!text) {
      if (found) sheet.getRange(found, 2, 1, 4).clearContent();
    } else if (found) {
      webSetText_(sheet.getRange(found, 5), [[text]]);                       // E열: 내용 (서식 그대로)
    } else {
      const row = webFreeRowFrom_(sheet, rows);
      sheet.getRange(row, 2, 1, 2).setValues([[date, slot]]);                // B 날짜 | C 번호
      webSetText_(sheet.getRange(row, 4, 1, 2), [["", text]]);              // D 완료(빈칸) | E 내용 (서식 그대로)
    }
    // 지난 일정·메모 삭제는 하루 한 번, 그 외에는 이 시트 순서만 확인
    webTidyMaybe_(ss, tz, MEMO_SHEET);
    webBumpVer_();
    SpreadsheetApp.flush();
  });
  return getAll();
}

// -------------------------------------------------------------
// 메모 완료 표시 (꾹 눌러서 회색) →  MEMO 시트 D열에 "DONE" / 빈칸
// req = { date:"yyyy-MM-dd", slot:1~3, done:true|false }
// -------------------------------------------------------------
function setMemoDone(req) {
  return webLocked_(function () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const sheet = ss.getSheetByName(MEMO_SHEET);
    if (!sheet) throw new Error("MEMO 시트가 없어요.");
    const found = webFindMemo_(webReadRows_(sheet, 4), tz, req.date, Number(req.slot), true);   // 내용이 있는 메모만
    if (!found) throw new Error("메모를 찾을 수 없어요. 새로고침해 주세요.");
    sheet.getRange(found, 4).setValue(req.done ? "DONE" : "");   // D열: 완료 (서식은 건드리지 않음)
    webBumpVer_();
    SpreadsheetApp.flush();
    return true;
  });
}

// 그 날짜 · 번호의 메모 행 (없으면 0) — needText: E열(내용)이 있는 행만
function webFindMemo_(rows, tz, dateKey, slot, needText) {
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (webDateKey_(r.vals[0], r.disp[0], tz) === dateKey && parseInt(r.disp[1], 10) === slot &&
      (!needText || String(r.disp[3]).trim() !== "")) return r.row;
  }
  return 0;
}

// -------------------------------------------------------------
// 빨간 날 (달력에서 꾹 눌러 지정) — 문서 속성에 날짜 목록으로 저장
// req = { date:"yyyy-MM-dd", on:true|false }
// -------------------------------------------------------------
// 문서 속성에 저장한 날짜 목록 (빨간 날 HOLIDAYS · 연두·보라 날 LVDAYS)
function webGetDays_(prop) {
  try { return JSON.parse(PropertiesService.getDocumentProperties().getProperty(prop) || "[]"); }
  catch (e) { return []; }
}
function webGetHolidays_() { return webGetDays_("HOLIDAYS"); }
function webGetLvDays_() { return webGetDays_("LVDAYS"); }

// 목록에서 그 날짜를 켜고 끔 (지난 날짜는 같이 정리) → 바뀐 목록
function webToggleDay_(prop, req) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(req.date || ""))) throw new Error("날짜를 확인해 주세요.");
  const todayKey = webKey_(getTodayZero(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone());
  const list = webGetDays_(prop).filter(function (k) { return k >= todayKey && k !== req.date; });
  if (req.on) list.push(req.date);
  list.sort();
  PropertiesService.getDocumentProperties().setProperty(prop, JSON.stringify(list));
  return list;
}

function setHoliday(req) {
  return webLocked_(function () {
    const list = webToggleDay_("HOLIDAYS", req);
    webBumpVer_();
    return list;
  });
}

// -------------------------------------------------------------
// 연두·보라 날 (달력에서 꾹 누른 채 계속 → 지이잉 → 뾱) — 문서 속성에 날짜 목록으로 저장
// req = { date:"yyyy-MM-dd", on:true|false }
// -------------------------------------------------------------
function setLvDay(req) {
  return webLocked_(function () {
    const list = webToggleDay_("LVDAYS", req);
    // 연두·보라 날이 되면 그날 빨간 날 지정은 풀림 (끄면 검정 · 일요일이면 빨강으로 돌아감)
    if (req.on) webToggleDay_("HOLIDAYS", { date: req.date, on: false });
    webBumpVer_();
    return list;
  });
}

// =============================================================
// 내부 헬퍼
// =============================================================

// -------------------------------------------------------------
// 칸 서식을 그대로 두고 글자 쓰기
//  · 잠깐 글자(@) 서식으로 바꿔서 쓰고 → "2h" · "1/2" 같은 글자가 숫자·날짜로 바뀌지 않음
//  · 쓰고 나면 칸마다 원래 서식(예: " "@)으로 바로 되돌림 → 값은 글자 그대로, 서식은 시트에 정해 둔 그대로
// -------------------------------------------------------------
function webSetText_(range, values) {
  const fmts = range.getNumberFormats();
  range.setNumberFormat("@").setValues(values);
  for (let i = 0; i < fmts.length; i++) {
    for (let j = 0; j < fmts[i].length; j++) {
      if (fmts[i][j] && fmts[i][j] !== "@") range.getCell(i + 1, j + 1).setNumberFormat(fmts[i][j]);
    }
  }
}
// 시간 칸: 이미 시간 서식(h 가 들어간 서식)이면 그대로, 아니면 HH:mm
function webTimeFmt_(cell) {
  if (!/h/i.test(String(cell.getNumberFormat() || ""))) cell.setNumberFormat("HH:mm");
}

// -------------------------------------------------------------
// 지난 시간 규칙 (앱에서도 같은 규칙으로 막지만, 서버에서 한 번 더 확인)
//  · 앱 → 서버 전달 시간을 감안해 2분 여유
// -------------------------------------------------------------
const WEB_PAST_GRACE = 120000;
// 이미 끝난 일정인지: 기본 = 끝 시각이 지남 / UNTIL = 그 시각이 지남 / FROM = 시작 후 60분이 지남 (60분짜리 일정처럼 · 앱과 같은 규칙)
const WEB_FROM_LIVE = 60;
function webIsPast_(key, abs, dur, tz, now) {
  const kind = webKind_(dur);
  if (kind === "norm") return abs + webDurMin_(dur) * 60000 <= now - WEB_PAST_GRACE;
  if (kind === "out") return abs <= now - WEB_PAST_GRACE;
  return abs + WEB_FROM_LIVE * 60000 <= now - WEB_PAST_GRACE;
}

// 시작 행부터 마지막 행까지 B열부터 ncols 칸 읽기 (완전히 빈 줄은 건너뜀)
function webReadRows_(sheet, ncols) {
  const last = sheet.getLastRow();
  if (last < WEB_START_ROW) return [];
  const n = last - WEB_START_ROW + 1;
  const rng = sheet.getRange(WEB_START_ROW, 2, n, ncols);
  const vals = rng.getValues();
  const disp = rng.getDisplayValues();
  const out = [];
  for (let i = 0; i < n; i++) {
    if (disp[i].every(function (v) { return String(v).trim() === ""; })) continue;
    out.push({ row: WEB_START_ROW + i, vals: vals[i], disp: disp[i] });
  }
  return out;
}

// -------------------------------------------------------------
// 겹침 검사: [시작, 끝) 구간이 1분이라도 겹치면 그 일정 설명을 돌려줌
//  · 끝과 시작이 딱 맞닿는 건 (12:00 끝 / 12:00 시작) 겹침 아님
//  · FROM / UNTIL 은 그 시각 1분짜리로 봄
// -------------------------------------------------------------
function webSpan_(dayMs, start, dur) {
  const len = webKind_(dur) === "norm" ? webDurMin_(dur) : 1;
  const s = dayMs + start * 60000;
  return { s: s, e: s + len * 60000 };
}

function webFindOverlap_(rows, tz, date, start, dur, skipRow) {
  const me = webSpan_(date.getTime(), start, dur);
  let found = "";
  rows.some(function (r) {
    if (r.row === skipRow) return false;
    const key = webDateKey_(r.vals[0], r.disp[0], tz);
    const st = webParseTime_(r.disp[1]);
    const name = String(r.disp[3]).trim();
    if (!key || st == null || !name) return false;
    const d = String(r.disp[2]).trim() || "30m";
    const o = webSpan_(webKeyToDate_(key).getTime(), st, d);
    if (o.s < me.e && me.s < o.e) {
      found = name + " (" + key.slice(5).replace("-", "/") + " " + webHM_(st) + ")";
      return true;
    }
    return false;
  });
  return found;
}

// -------------------------------------------------------------
// 정리: 지난 일정·메모 완전 삭제 → 정렬
//  · DATA: 끝나는 시각이 오늘 0시 이전인 일정만 삭제
//          (어제 시작해서 오늘까지 이어지는 일정은 남겨둠)
//  · MEMO: 날짜가 어제 이전인 메모 삭제
// -------------------------------------------------------------
function webTidy_(ss, tz) {
  // TIMELINE 핵심 이벤트 표: 날짜가 지난 행은 행 자체를 삭제
  const tl = ss.getSheetByName("TIMELINE");
  if (tl) webTidyTimeline_(tl, getTodayZero().getTime());

  const data = ss.getSheetByName(DATA_SHEET);
  if (data) webTidySheet_(data, tz, true, true);
  const memo = ss.getSheetByName(MEMO_SHEET);
  if (memo) webTidySheet_(memo, tz, false, true);
}

// 정리는 하루 한 번만 (그날 첫 저장 때) — 나머지 저장 때는 방금 고친 시트의 순서만 확인
//  · 지난 일정이 하루 동안 시트에 남아 있어도 앱은 날짜로 걸러서 보여 주므로 화면에는 영향 없음
function webTidyMaybe_(ss, tz, touched) {
  const props = PropertiesService.getDocumentProperties();
  const todayKey = webKey_(getTodayZero(), tz);
  if (props.getProperty("TIDY_DAY") !== todayKey) {
    webTidy_(ss, tz);
    props.setProperty("TIDY_DAY", todayKey);               // 정리가 끝까지 된 뒤에만 기록
    return;
  }
  const sh = ss.getSheetByName(touched);
  if (sh) webTidySheet_(sh, tz, touched === DATA_SHEET, false);
}

// 시트 하나 정리 + 정렬
//  · isData: DATA 면 (날짜, 시작시간) / MEMO 면 (날짜, 번호) 순서
//  · clearPast: 지난 줄을 지울지 (false 면 순서만 확인해서 어긋났을 때만 정렬)
//  · DATA: 끝나는 시각이 오늘 0시 이전인 일정만 삭제 (어제 시작해서 오늘까지 이어지는 일정은 남겨둠)
//  · MEMO: 날짜가 오늘 이전인 메모 삭제
function webTidySheet_(sheet, tz, isData, clearPast) {
  const today = getTodayZero(), todayMs = today.getTime(), todayKey = webKey_(today, tz);
  const rows = webReadRows_(sheet, 4), clear = [], keep = [];
  rows.forEach(function (r) {
    const key = webDateKey_(r.vals[0], r.disp[0], tz);
    if (isData) {
      const start = webParseTime_(r.disp[1]);
      if (clearPast && key) {                                  // 날짜를 못 읽는 줄은 건드리지 않음
        const days = Math.round((todayMs - webKeyToDate_(key).getTime()) / 86400000);
        const end = (start == null ? 0 : start) + webDurMin_(r.disp[2]);
        if (days > 0 && days * 1440 >= end) { clear.push(r.row); return; }   // 오늘·미래는 남김
      }
      keep.push({ row: r.row, a: key, b: start });
    } else {
      if (clearPast && key && key < todayKey) { clear.push(r.row); return; }
      keep.push({ row: r.row, a: key, b: parseInt(r.disp[1], 10) });
    }
  });
  webClearRows_(sheet, clear, 4);
  if (clear.length || !webInOrder_(keep)) webSort_(sheet, 4, [{ column: 2, ascending: true }, { column: 3, ascending: true }]);
}

// DATA B열 날짜 칸 전체를 같은 서식으로 (시트에 직접 넣어서 "2026-12-01" 처럼 보이는 줄도 같이 맞춰짐)
//  · 글자를 바꾸는 게 아니라 서식만 바꿈 → 날짜 값은 그대로라 앱이 읽는 데 영향 없음
//  · 한 번의 호출로 범위 전체에 적용 (줄마다 따로 하지 않음)
function webFormatDates_(sheet) {
  const last = sheet.getLastRow();
  if (last < WEB_START_ROW) return;
  sheet.getRange(WEB_START_ROW, 2, last - WEB_START_ROW + 1, 1).setNumberFormat(DATA_DATE_FORMAT);
}

// 여러 줄을 한 번에 비우기 (줄마다 따로 지우면 느림)
function webClearRows_(sheet, rows, ncols) {
  if (!rows.length) return;
  const last = String.fromCharCode(65 + ncols);                 // B 부터 ncols 칸 → 끝 열 글자
  sheet.getRangeList(rows.map(function (r) { return "B" + r + ":" + last + r; })).clearContent();
}

// 이미 시작 행부터 빈 줄 없이 (날짜, 두 번째 값) 순서대로면 true → 정렬을 건너뜀
//  · 날짜를 못 읽는 줄이 있으면 예전처럼 정렬함
function webInOrder_(list) {
  for (let i = 0; i < list.length; i++) {
    const x = list[i];
    if (x.row !== WEB_START_ROW + i || !x.a) return false;
    if (i > 0) {
      const p = list[i - 1];
      const pb = p.b == null ? 1e9 : p.b, xb = x.b == null ? 1e9 : x.b;   // 빈 칸은 시트 정렬처럼 맨 뒤
      if (p.a > x.a || (p.a === x.a && pb > xb)) return false;
    }
  }
  return true;
}

// TIMELINE 표에서 날짜가 지난 행 삭제 (행 자체를 지움)
//  · 날짜가 비어 있는 행은 건드리지 않음 (입력 중일 수 있어서)
//  · 표에 행이 하나만 남으면 서식 보존을 위해 삭제 대신 내용만 비움 (메뉴의 "추가"가 이 행 서식을 복사해서 씀)
function webTidyTimeline_(sheet, todayMs) {
  const last = getTimelineLastRow(sheet);
  if (last < TL_START_ROW) return;
  const n = last - TL_START_ROW + 1;
  const dates = sheet.getRange(TL_START_ROW, TL_COL_DATE, n, 1).getValues();
  let remaining = n, changed = false;
  for (let i = n - 1; i >= 0; i--) {
    const v = dates[i][0];
    if (!(v instanceof Date) || isNaN(v.getTime())) continue;
    const d = new Date(v);
    d.setHours(0, 0, 0, 0);
    if (d.getTime() >= todayMs) continue;
    const row = TL_START_ROW + i;
    if (remaining > 1) {
      sheet.deleteRow(row);
      remaining--;
    } else {
      sheet.getRange(row, TL_COL_DATE, 1, TL_NUM_COLS - (TL_COL_DATE - TL_FIRST_COL)).clearContent();
    }
    changed = true;
  }
  if (changed) refreshDdayFormulas(sheet);
}

// 시작 행부터 마지막 줄까지 정렬 (빈 줄은 맨 아래로 내려감)
function webSort_(sheet, ncols, spec) {
  const last = sheet.getLastRow();
  if (last <= WEB_START_ROW) return;
  sheet.getRange(WEB_START_ROW, 2, last - WEB_START_ROW + 1, ncols).sort(spec);
}

// 이미 읽어 둔 줄 목록으로 비어 있는 첫 줄 찾기 (시트를 다시 읽지 않음)
function webFreeRowFrom_(sheet, rows) {
  let expect = WEB_START_ROW;
  for (let i = 0; i < rows.length; i++) {
    if (rows[i].row !== expect) return expect;                  // 중간에 빈 줄
    expect++;
  }
  // 읽은 줄 목록은 마지막 줄까지 빈 줄을 뺀 것 → 이어진 줄 바로 다음 줄이 첫 빈 줄 (예전 webFreeRow_ 와 같은 결과)
  webEnsureRows_(sheet, expect);
  return expect;
}

// 시트 행이 모자라면 맨 아래에 행 추가 (바로 위 행의 서식을 이어받음)
function webEnsureRows_(sheet, lastNeeded) {
  const max = sheet.getMaxRows();
  if (lastNeeded > max) sheet.insertRowsAfter(max, lastNeeded - max);
}

function webKey_(d, tz) {
  return Utilities.formatDate(d, tz, "yyyy-MM-dd");
}

function webKeyToDate_(key) {
  const m = String(key || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
}

// 날짜 칸 → "yyyy-MM-dd"
function webDateKey_(v, disp, tz) {
  if (v instanceof Date && !isNaN(v.getTime())) return webKey_(v, tz);
  const m = String(disp || "").match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  if (!m) return null;
  return m[1] + "-" + ("0" + m[2]).slice(-2) + "-" + ("0" + m[3]).slice(-2);
}

// "10:15", "오후 3:00", "3:00 PM" → 분
function webParseTime_(s) {
  s = String(s || "").trim();
  if (!s) return null;
  const m = s.match(/(\d{1,2})(?::(\d{2}))?(?::\d{2})?/);
  if (!m) return null;
  let h = Number(m[1]);
  const mi = Number(m[2] || 0);
  if (/(PM|오후)/i.test(s) && h < 12) h += 12;
  if (/(AM|오전)/i.test(s) && h === 12) h = 0;
  if (h > 23 || mi > 59) return null;
  return h * 60 + mi;
}

function webHM_(min) {
  return ("0" + Math.floor(min / 60)).slice(-2) + ":" + ("0" + (min % 60)).slice(-2);
}

// "30m", "2h", "1h30m", "90" → 분 (FROM/UNTIL · 예전 IN/OUT 은 0)
function webDurMin_(t) {
  const s = webSplitDur_(t).dur.toUpperCase().replace(/\s/g, "");
  if (webKind_(s) !== "norm") return 0;
  let mins = 0, found = false;
  const h = s.match(/(\d+(?:\.\d+)?)H/);
  if (h) { mins += Math.round(parseFloat(h[1]) * 60); found = true; }
  const m = s.match(/(\d+)M/);
  if (m) { mins += Number(m[1]); found = true; }
  if (!found && /^\d+$/.test(s)) { mins = Number(s); found = true; }
  return found && mins > 0 ? mins : 30;
}

// D열 → { dur: 진행시간 글자, mark: "imp"(중요) / "tbd"(미정) / "" }
//  · "2h MAJOR", "MAJOR 2h", "1h30m TBD" 처럼 진행시간과 같이 적음 (순서 · 띄어쓰기 상관없음)
//  · "MAJOR" / "TBD" 만 적으면 진행시간은 기본 30m
const WEB_MARK_RE = /(MAJOR|TBD|중요|미정)/gi;
function webSplitDur_(t) {
  const raw = String(t || "");
  let mark = "";
  const m = raw.match(WEB_MARK_RE);
  if (m) mark = /^(MAJOR|중요)$/i.test(m[0]) ? "imp" : "tbd";
  const dur = raw.replace(WEB_MARK_RE, " ").replace(/\s+/g, " ").trim();
  return { dur: dur, mark: mark };
}
// "imp" / "tbd" → 시트에 적을 글자
function webMarkText_(m) {
  return m === "imp" ? "MAJOR" : m === "tbd" ? "TBD" : "";
}

function webKind_(t) {
  const s = webSplitDur_(t).dur.toUpperCase();
  // FROM = 시작만 있는 일정 / UNTIL = 끝만 있는 일정 (예전 IN / OUT 도 같은 뜻으로 인식)
  return (s === "FROM" || s === "IN") ? "in" : (s === "UNTIL" || s === "OUT") ? "out" : "norm";
}

// 진행시간 드롭다운 목록 (DATA D열 첫 데이터 칸의 데이터 확인 규칙) — 30분 캐시 (목록을 바꾸면 30분 안에 반영)
function webDurationsCached_(sheet) {
  const cache = CacheService.getScriptCache();
  const hit = cache.get("web_durations_v6");
  if (hit) return JSON.parse(hit);
  let list = WEB_DEFAULT_DURATIONS;
  try {
    const dv = sheet ? sheet.getRange(WEB_START_ROW, 4).getDataValidation() : null;
    if (dv) {
      const type = dv.getCriteriaType();
      const vals = dv.getCriteriaValues();
      if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST && vals[0] && vals[0].length) {
        list = vals[0].map(String);
      } else if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE && vals[0]) {
        const l = [];
        vals[0].getDisplayValues().forEach(function (r) {
          r.forEach(function (v) { if (String(v).trim() !== "") l.push(String(v).trim()); });
        });
        if (l.length) list = l;
      }
    }
  } catch (e) { }
  // 드롭다운에 "2h MAJOR" 같은 값이 있어도 앱에는 진행시간만 (중복 없이)
  const seen = {};
  list = list.map(function (v) { return webSplitDur_(v).dur; })
    .filter(function (v) { if (!v || seen[v]) return false; seen[v] = true; return true; });
  if (!list.length) list = WEB_DEFAULT_DURATIONS;
  cache.put("web_durations_v6", JSON.stringify(list), 1800);
  return list;
}

// TIMELINE 하단 표 (B: VIEW | C: DATE | D: TITLE | E: IMG | F: LINK | G: D-DAY)
//  · 열 위치는 Code.gs 의 TL_COL_* 상수에서 가져옴 → 표 열이 또 바뀌면 Code.gs 상수만 고치면 됨
function webReadTimelineTable_(sheet) {
  const last = getTimelineLastRow(sheet);
  if (last < TL_START_ROW) return [];

  const n = last - TL_START_ROW + 1;
  const rng = sheet.getRange(TL_START_ROW, TL_FIRST_COL, n, TL_NUM_COLS);
  const vals = rng.getValues();
  const disp = rng.getDisplayValues();
  // 한 줄 안에서 각 칸의 위치 (B 가 0)
  const cV = TL_COL_VIEW - TL_FIRST_COL, cD = TL_COL_DATE - TL_FIRST_COL, cT = TL_COL_TITLE - TL_FIRST_COL,
    cI = TL_COL_IMG - TL_FIRST_COL, cL = TL_COL_LINK - TL_FIRST_COL, cY = TL_COL_DDAY - TL_FIRST_COL;
  // 수식은 IMG · LINK 칸만, 링크 서식은 LINK 칸만 읽음 (표 전체를 읽으면 느림)
  //  · IMG ~ LINK 를 한 번에 읽음 (LINK 가 IMG 오른쪽에 있다는 전제)
  const fw = TL_COL_LINK - TL_COL_IMG;
  const forms = sheet.getRange(TL_START_ROW, TL_COL_IMG, n, fw + 1).getFormulas();     // [0] = IMG, [fw] = LINK
  let rich = null;
  try { rich = sheet.getRange(TL_START_ROW, TL_COL_LINK, n, 1).getRichTextValues(); } catch (e) { }

  const out = [];
  for (let i = 0; i < n; i++) {
    const dateV = vals[i][cD];
    const title = String(disp[i][cT]).trim();
    const hasDate = dateV instanceof Date && !isNaN(dateV.getTime());
    if (!hasDate && title === "") continue;

    let time = null;
    if (hasDate) {
      const d = new Date(dateV);
      d.setHours(0, 0, 0, 0);
      time = d.getTime();
    }

    out.push({
      view: vals[i][cV] === true,
      time: time,
      dateText: String(disp[i][cD]).trim(),
      title: title,
      img: webImageUrl_(sheet, vals[i][cI], forms[i][0]),
      link: webLinkUrl_(forms[i][fw], rich ? rich[i][0] : null, disp[i][cL]),
      linkText: String(disp[i][cL]).trim(),
      dday: String(disp[i][cY]).trim()
    });
  }
  return out;
}

// IMAGE("url") / IMAGE(셀참조) 수식 또는 셀 안 이미지에서 주소 꺼내기
function webImageUrl_(sheet, value, formula) {
  if (formula) {
    const m = String(formula).match(/IMAGE\(\s*"([^"]+)"/i);
    if (m) return m[1];
    const r = String(formula).match(/IMAGE\(\s*(?:'?[^'!]+'?!)?\$?([A-Z]{1,3})\$?(\d+)/i);
    if (r) {
      try {
        const v = String(sheet.getRange(r[1] + r[2]).getDisplayValue()).trim();
        if (/^https?:\/\//i.test(v)) return v;
      } catch (e) { }
    }
  }
  // 셀 안에 넣은 이미지: 원본 주소(getUrl)가 있으면 그걸 먼저, 없으면 구글 주소(getContentUrl)
  if (value && typeof value === "object" && typeof value.getContentUrl === "function") {
    try { const u = value.getUrl && value.getUrl(); if (u && /^https?:\/\//i.test(u)) return u; } catch (e) { }
    try { return value.getContentUrl() || ""; } catch (e) { }
  }
  if (typeof value === "string" && /^https?:\/\//i.test(value.trim())) return value.trim();
  return "";
}

// -------------------------------------------------------------
// 이미지를 서버에서 받아서 data URL로 전달
// -------------------------------------------------------------
function getImages(urls) {
  const cache = CacheService.getScriptCache();
  const data = {}, err = {};
  const list = (urls || []).filter(function (u) { return /^https?:\/\//i.test(u); }).slice(0, 30);
  if (!list.length) return { data: data, err: err };

  const hits = cache.getAll(list.map(webImgKey_));
  const need = [];
  list.forEach(function (u) {
    const h = hits[webImgKey_(u)];
    if (h) data[u] = h; else need.push(u);
  });
  if (!need.length) return { data: data, err: err };

  // 1차: 로그인 정보 없이
  const resps = webFetchAll_(need.map(function (u) { return { url: u, muteHttpExceptions: true, followRedirects: true }; }));

  // 2차: 실패한 구글 주소만 로그인 정보를 붙여 다시 (다른 사이트에는 절대 보내지 않음)
  //  · 주소 이름이 정확히 google.com 등이거나 그 하위 주소(xxx.google.com)일 때만
  //    (예전 검사는 evilgoogle.com 같은 남의 주소도 통과시켰음)
  const retry = [];
  need.forEach(function (u, i) {
    if (!webIsImage_(resps[i]) && WEB_GOOGLE_HOST_RE.test(u)) retry.push(i);
  });
  if (retry.length) {
    const token = ScriptApp.getOAuthToken();
    const r2 = webFetchAll_(retry.map(function (i) {
      return { url: need[i], muteHttpExceptions: true, followRedirects: true, headers: { Authorization: "Bearer " + token } };
    }));
    retry.forEach(function (i, k) { if (webIsImage_(r2[k]) || !resps[i] || resps[i].__err) resps[i] = r2[k]; });
  }

  need.forEach(function (u, i) {
    const res = resps[i];
    try {
      if (!res || res.__err) { err[u] = "접속 실패: " + (res ? res.__err : "응답 없음"); return; }
      const code = res.getResponseCode();
      const blob = res.getBlob();
      const type = blob.getContentType() || "";
      if (code !== 200) { err[u] = "HTTP " + code; return; }
      if (!/^image\//i.test(type)) { err[u] = "이미지 아님 (" + type + ")"; return; }
      const bytes = blob.getBytes();
      if (bytes.length > 5 * 1024 * 1024) { err[u] = "5MB 초과"; return; }
      const d = "data:" + type + ";base64," + Utilities.base64Encode(bytes);
      data[u] = d;
      if (d.length < 100000) cache.put(webImgKey_(u), d, 21600);
    } catch (e) {
      err[u] = "오류 " + e.message;
    }
  });
  return { data: data, err: err };
}

const WEB_GOOGLE_HOST_RE = /^https:\/\/([a-z0-9-]+\.)*(googleusercontent\.com|google\.com|ggpht\.com)(:\d+)?\//i;

function webFetchAll_(reqs) {
  try {
    return UrlFetchApp.fetchAll(reqs);
  } catch (e) {
    return reqs.map(function (r) { try { return UrlFetchApp.fetch(r.url, r); } catch (e2) { return { __err: e2.message }; } });
  }
}

function webIsImage_(res) {
  try { return !!res && res.getResponseCode() === 200 && /^image\//i.test(res.getBlob().getContentType() || ""); }
  catch (e) { return false; }
}

function webImgKey_(u) {
  return "img_" + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, u));
}

// 셀 링크 · HYPERLINK("url") 수식 · 칸에 그냥 적은 주소 순서로 꺼내기
function webLinkUrl_(formula, richText, text) {
  if (richText) {
    try {
      const u = richText.getLinkUrl();
      if (u) return u;
      const runs = richText.getRuns();
      for (let i = 0; i < runs.length; i++) {
        const u2 = runs[i].getLinkUrl();
        if (u2) return u2;
      }
    } catch (e) { }
  }
  if (formula) {
    const m = String(formula).match(/HYPERLINK\(\s*"([^"]+)"/i);
    if (m) return m[1];
  }
  // 링크 서식 없이 주소 글자만 넣은 경우
  const t = String(text || "").trim();
  if (/^https?:\/\/\S+$/i.test(t)) return t;
  return "";
}