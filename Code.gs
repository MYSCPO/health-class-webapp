/**
 * 감염병 예방과 면역 - 보건 수업 웹앱 서버 측 코드 (Code.gs)
 * 작성자: 에듀테크 보건교사 & 수석 개발자
 */


// 1. 웹앱 웹페이지 렌더링
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('보건 수업: 감염병 예방과 면역')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// 2. 구글 시트 초기화 함수 (최초 1회 실행 권장)
function setupSheetHeader() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
 
  // '제출응답' 시트 설정
  let responseSheet = ss.getSheetByName('제출응답');
  if (!responseSheet) {
    responseSheet = ss.insertSheet('제출응답');
  } else {
    responseSheet.clear(); // 기존 내용 초기화
  }
 
  const headers = [
    '제출시간', '학번', '이름',
    '[활동1] 전파고리 끊기', '[활동2-1] 지구온난화 영향', '[활동2-2] 환경보호 실천',
    '퀴즈1 답', '퀴즈2 답', '퀴즈3 답', '퀴즈 총점',
    '배운 점', '느낀 점', '실천할 점'
  ];
 
  // 헤더 작성 및 보라색 스타일 적용 (Indigo tone)
  responseSheet.appendRow(headers);
  const headerRange = responseSheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4F46E5'); // Tailwind Indigo-600
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  responseSheet.setFrozenRows(1);
 
  // '명렬표' 시트 설정 (없을 경우 예시 데이터 추가)
  let studentSheet = ss.getSheetByName('명렬표');
  if (!studentSheet) {
    studentSheet = ss.insertSheet('명렬표');
    studentSheet.appendRow(['학번', '이름']);
    studentSheet.getRange(1, 1, 1, 2).setBackground('#E0E7FF').setFontWeight('bold');
   
    // 예시 샘플 데이터 삽입
    const sampleStudents = [
      ['10101', '강보건'],
      ['10102', '김면역'],
      ['10103', '박예방'],
      ['10104', '이백신'],
      ['10105', '최청결']
    ];
    sampleStudents.forEach(row => studentSheet.appendRow(row));
  }
 
  return "초기화가 완료되었습니다. '제출응답' 및 '명렬표' 시트를 확인하세요.";
}


// 3. 학번으로 이름 자동 조회 함수
function findStudentName(studentId) {
  if (!studentId) return '';
 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentSheet = ss.getSheetByName('명렬표');
 
  if (!studentSheet) return '명렬표 시트 없음';
 
  const data = studentSheet.getDataRange().getValues();
  // 첫 번째 행(헤더)을 제외하고 검색
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(studentId).trim()) {
      return data[i][1]; // 이름 반환
    }
  }
  return ''; // 일치하는 학번이 없을 경우
}


// 4. 학생 제출 데이터 처리 및 퀴즈 자동 채점 함수
function submitData(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let responseSheet = ss.getSheetByName('제출응답');
   
    if (!responseSheet) {
      setupSheetHeader();
      responseSheet = ss.getSheetByName('제출응답');
    }


    // 형성평가 정답지 지정 및 채점 (Q1: 3, Q2: 2, Q3: 4) - 점수: 33, 33, 34
    const answerKey = { q1: '3', q2: '2', q3: '4' };
    let score = 0;
   
    if (formData.q1 === answerKey.q1) score += 33;
    if (formData.q2 === answerKey.q2) score += 33;
    if (formData.q3 === answerKey.q3) score += 34;


    // 저장할 데이터 배열 생성
    const rowData = [
      new Date(),                     // 제출 시간
      formData.studentId,            // 학번
      formData.studentName,          // 이름
      formData.act1,                 // 탐구1: 전파고리 끊기
      formData.act2_1,               // 탐구2-1: 지구온난화 영향
      formData.act2_2,               // 탐구2-2: 환경보호 실천
      formData.q1,                   // 퀴즈 1 선택답
      formData.q2,                   // 퀴즈 2 선택답
      formData.q3,                   // 퀴즈 3 선택답
      score,                         // 퀴즈 총점
      formData.reflection_learn,     // 배운 점
      formData.reflection_feel,      // 느낀 점
      formData.reflection_act        // 실천할 점
    ];


    // 행 추가
    responseSheet.appendRow(rowData);


    // 성공 메시지와 획득 점수 반환
    return {
      success: true,
      score: score,
      message: `제출이 성공적으로 완료되었습니다! 형성평가 점수: ${score}점 / 100점`
    };


  } catch (error) {
    return {
      success: false,
      message: '저장 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

