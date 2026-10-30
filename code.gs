
const 담당자이메일 = "admin@example.com"; // 담당자 이메일

function 이벤트응모알림(e) {

  // 직접 실행한 경우 오류 방지
  if (!e || !e.range) {
    Logger.log("폼 제출 트리거로 실행해주세요.");
    return;
  }
  // 새로 제출된 응답의 시트 확인
  const 시트 = e.range.getSheet();
  if (시트.getName() !== "이벤트응모") return;

  // 새로 입력된 행의 데이터 가져오기
  const 행 = e.range.getRow();
  const 데이터 = 시트.getRange(행, 1, 1, 6).getValues()[0];

  const 이메일 = 데이터[1];   // B열: 이메일
  const 이름 = 데이터[2];     // C열: 이름
  const 관심분야 = 데이터[3]; // D열: 관심분야
  const 수신동의 = 데이터[5]; // F열: 수신동의

  // 수신 동의한 응모자에게만 확인 메일 발송
  if (String(수신동의).trim() === "동의" && 이메일) {
    MailApp.sendEmail({
      to: 이메일,
      subject: "이벤트 응모가 접수되었습니다",
      body: `안녕하세요, ${이름}님!\n\n이벤트 응모가 접수되었습니다.\n관심분야: ${관심분야}\n\n감사합니다.`
    });
  }

  // 담당자에게 신규 응모 알림 (동의 여부와 무관)
  MailApp.sendEmail({
    to: 담당자이메일,
    subject: "새로운 이벤트 응모 알림",
    body: `새로운 응모가 접수되었습니다.\n\n이름: ${이름}\n이메일: ${이메일}\n관심분야: ${관심분야}\n수신동의: ${수신동의}`
  });
}




const 내이메일 = "myemail@gmail.com"; // 받을 이메일 주소

function 일일지표발송() {
  // 1. 시트와 데이터 가져오기
  const 시트 = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("일일지표");

  const 마지막행 = 시트.getLastRow();
  if (마지막행 < 2) return; // 데이터가 없으면 종료

  const 데이터 = 시트.getRange(2, 1, 마지막행 - 1, 6).getValues();

  // 2. 날짜가 있는 행만 추출하고 최신 날짜순 정렬
  const 유효데이터 = 데이터
    .filter(행 => 행[0] instanceof Date && !isNaN(행[0].getTime()))
    .sort((a, b) => b[0].getTime() - a[0].getTime());

  // 날짜 데이터가 없으면 메일 발송 없이 종료
  if (유효데이터.length === 0) return;

  // 3. 가장 최근 날짜의 데이터
  const [날짜, 방문자수, 신규가입, 주문건수, 매출, 광고비] = 유효데이터[0];

  // 4. 매출 대비 광고비 비율 계산
  // 매출이 0이면 0으로 나누는 오류 방지
  const 광고비율 = Number(매출) > 0
    ? (Number(광고비) / Number(매출) * 100).toFixed(2) + "%"
    : "계산 불가";

  // 5. 날짜 형식 변환
  const 날짜문자열 = Utilities.formatDate(
    날짜,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );

  // 6. 이메일 본문 작성
  const 본문 = `
[일일지표 보고서]

날짜: ${날짜문자열}
방문자 수: ${방문자수}
신규 가입: ${신규가입}
주문 건수: ${주문건수}
매출: ${Number(매출).toLocaleString()}원
광고비: ${Number(광고비).toLocaleString()}원

매출 대비 광고비 비율: ${광고비율}
`;

  // 7. 이메일 발송
  MailApp.sendEmail({
    to: 내이메일,
    subject: `[일일지표] ${날짜문자열} 보고서`,
    body: 본문
  });

  Logger.log("일일지표 이메일 발송 완료");
}



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

