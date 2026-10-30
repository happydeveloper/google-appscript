
function 콘텐츠일정등록() {
  // 1. 스프레드시트와 기본 캘린더 가져오기
  const 시트 = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("콘텐츠일정");

  const 캘린더 = CalendarApp.getDefaultCalendar();

  // 2. 제목행을 제외한 데이터 가져오기
  const 마지막행 = 시트.getLastRow();
  if (마지막행 < 2) return;

  const 데이터 = 시트.getRange(2, 1, 마지막행 - 1, 5).getValues();

  // 3. 각 행을 확인하며 일정 등록
  데이터.forEach((행, i) => {
    const [날짜, 채널, 제목, 담당자, 상태] = 행;

    // 이미 등록된 일정은 건너뛰기
    if (String(상태).trim() === "등록완료") return;

    // 날짜나 제목이 없으면 건너뛰기
    if (!(날짜 instanceof Date) || isNaN(날짜.getTime()) || !제목) return;

    // 캘린더 일정 제목 만들기
    const 일정제목 = `[${채널}] ${제목}`;

    try {
      // 4. Google 캘린더에 종일 일정 생성
      캘린더.createAllDayEvent(일정제목, 날짜, {
        description: `담당자: ${담당자}`
      });

      // 5. 등록 성공 시 E열 상태 변경
      시트.getRange(i + 2, 5).setValue("등록완료");

      Logger.log(`${일정제목} 등록 성공`);

    } catch (오류) {
      // 오류 발생 시 상태를 변경하지 않음
      Logger.log(`${i + 2}행 등록 실패: ${오류.message}`);
    }
  });
}
