/**
 * 3-3 회식 메뉴 조사 - 구글 시트 저장용 Apps Script
 *
 * - "응답" 시트: 학번마다 한 줄 (같은 학번이 다시 제출하면 마지막 제출로 덮어씀)
 * - "기록" 시트: 모든 제출을 시간순으로 전부 남김 (확인용)
 */

const HEADERS = ["학번", "이름", "파", "치킨 브랜드", "치킨 메뉴", "뼈/순살", "음료", "엽떡 메뉴", "엽떡 맵기", "제출 시각", "제출 횟수"];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const id = String(d.id || "").replace(/\D/g, "");
    const name = String(d.name || "").trim();
    if (!/^\d{5}$/.test(id) || !name) return json({ ok: false, error: "학번/이름 오류" });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getSheet(ss, "응답");
    const log = getSheet(ss, "기록");
    const now = new Date();
    const row = [id, name, d.party || "", d.brand || "", d.flavor || "", d.bone || "", d.drink || "", d.yMenu || "", d.ySpicy || "", now];

    log.appendRow(row);

    const last = sheet.getLastRow();
    let target = -1, count = 0;
    if (last > 1) {
      const ids = sheet.getRange(2, 1, last - 1, 1).getDisplayValues();
      for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) { target = i + 2; break; }
    }
    if (target > 0) {
      count = Number(sheet.getRange(target, 11).getValue()) || 1;
      sheet.getRange(target, 1, 1, 11).setValues([row.concat([count + 1])]);
    } else {
      target = last + 1;
      sheet.getRange(target, 1).setNumberFormat("@");
      sheet.getRange(target, 1, 1, 11).setValues([row.concat([1])]);
    }
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json({ ok: true, msg: "3-3 회식 조사 서버 작동 중" });
}

function getSheet(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight("bold");
    sh.getRange("A:A").setNumberFormat("@");
    sh.setFrozenRows(1);
  }
  return sh;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
