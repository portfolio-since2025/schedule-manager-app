function handleEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();

  // 1行目は見出し
  if (row === 1) return;

  // A列(日にち)かB列(予定の詳細)だけ反応
  if (col !== 1 && col !== 2) return;

  const calendar = CalendarApp.getDefaultCalendar();
  const lastRow = sheet.getLastRow();

  // --- 1. 全体の重複判定を更新 ---
  const details = sheet.getRange(2, 2, Math.max(lastRow - 1, 1)).getValues().flat();

  for (let i = 0; i < details.length; i++) {
    const currentRow = i + 2;
    const detailValue = details[i];
    const detailCell = sheet.getRange(currentRow, 2);
    const judgeCell = sheet.getRange(currentRow, 3);

    if (!detailValue) {
      detailCell.setBackground(null);
      judgeCell.setValue("");
      continue;
    }

    const count = details.filter(v => v === detailValue).length;

    if (count > 1) {
      detailCell.setBackground("#ffcccc");
      judgeCell.setValue("重複");
    } else {
      detailCell.setBackground(null);
      judgeCell.setValue("新規");
    }
  }

  // --- 2. 編集した行の情報 ---
  const dateValue = sheet.getRange(row, 1).getValue();   // A列: 日にち
  const detailValue = sheet.getRange(row, 2).getValue(); // B列: 予定の詳細
  const judgeValue = sheet.getRange(row, 3).getValue();  // C列: 判定
  const calendarCell = sheet.getRange(row, 4);           // D列: カレンダー
  const eventIdCell = sheet.getRange(row, 5);            // E列: eventId

  // 予定タイトルが空なら、その行の eventId の予定だけ消す
  if (!detailValue) {
    const eventId = eventIdCell.getValue();
    if (eventId) {
      try {
        const event = calendar.getEventById(eventId);
        if (event) event.deleteEvent();
      } catch (error) {
        Logger.log("削除失敗: " + error);
      }
    }
    calendarCell.setValue("");
    eventIdCell.setValue("");
    return;
  }

  // 重複ならカレンダー登録しない
  if (judgeValue === "重複") {
    return;
  }

  // 日にちが空なら、その行の eventId の予定だけ消す
  if (!dateValue) {
    const eventId = eventIdCell.getValue();
    if (eventId) {
      try {
        const event = calendar.getEventById(eventId);
        if (event) event.deleteEvent();
      } catch (error) {
        Logger.log("削除失敗: " + error);
      }
    }
    calendarCell.setValue("");
    eventIdCell.setValue("");
    return;
  }

  // --- 3. 同じタイトルの予定をカレンダーから全部消す ---
  deleteAllEventsByTitle(calendar, detailValue);

  // --- 4. 新しい日に1件だけ作る ---
  const newEvent = calendar.createAllDayEvent(detailValue, new Date(dateValue));
  calendarCell.setValue("済");
  eventIdCell.setValue(newEvent.getId());
}

function deleteAllEventsByTitle(calendar, title) {
  const start = new Date(2025, 0, 1);   // 2025/01/01
  const end = new Date(2030, 11, 31);   // 2030/12/31
  const events = calendar.getEvents(start, end);

  for (const event of events) {
    if (event.getTitle() === title) {
      event.deleteEvent();
    }
  }
}

function testCalendar() {
  const calendar = CalendarApp.getDefaultCalendar();
  calendar.createAllDayEvent("テスト予定", new Date());
}
