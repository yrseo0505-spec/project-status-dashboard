/**
 * project-status-dashboard 전용 Google Apps Script (GAS) 백엔드
 * 대상 스프레드시트 ID: 1DJHA-GBPy6PCG7ghXvOBI9YmSq93SDlp-oaOtibVPg0
 *
 * [지원 action 목록]
 * 1. updateProject : '프로젝트 현황' 시트 F열(Project) 검색 -> 첫 번째 일치 행의 J열(Status) 덮어쓰기, O열(금주 진행현황) 줄바꿈(\n) 이어 붙이기
 * 2. addIssue      : 'IssueTracker' 시트 하단에 새 행 추가 [발생 일자, 관련 모델/프로젝트명, 이슈 내용]
 * 3. addMeeting    : 'MeetingMinutes' 시트 하단에 새 행 추가 [일자, 참석자, 회의 내용]
 */

const SPREADSHEET_ID = '1DJHA-GBPy6PCG7ghXvOBI9YmSq93SDlp-oaOtibVPg0';

/**
 * 대상 스프레드시트 객체를 반환합니다.
 * 시트에 바인딩된 스크립트면 getActiveSpreadsheet()를 우선 사용하고,
 * 독립 스크립트 환경이면 SPREADSHEET_ID로 엽니다.
 */
function getTargetSpreadsheet() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/**
 * POST 요청 엔드포인트 (GitHub Pages 프론트엔드에서 text/plain JSON으로 호출)
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    // 동시 요청 충돌 방지를 위해 최대 10초 대기
    lock.waitLock(10000);

    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({
        success: false,
        message: '전달된 요청 본문(body) 데이터가 없습니다.'
      });
    }

    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const ss = getTargetSpreadsheet();

    switch (action) {
      case 'updateProject':
        return handleUpdateProject(ss, payload);
      case 'addIssue':
        return handleAddIssue(ss, payload);
      case 'addMeeting':
        return handleAddMeeting(ss, payload);
      default:
        return createJsonResponse({
          success: false,
          message: '지원하지 않는 action 값입니다: ' + action
        });
    }
  } catch (error) {
    return createJsonResponse({
      success: false,
      message: '서버 처리 중 오류가 발생했습니다: ' + error.message
    });
  } finally {
    try {
      lock.releaseLock();
    } catch (_) {}
  }
}

/**
 * 웹앱 URL을 브라우저에서 직접 접속(GET)했을 때 상태 확인용 응답
 */
function doGet() {
  return createJsonResponse({
    success: true,
    message: 'project-status-dashboard GAS API가 정상 동작 중입니다. POST 요청으로 데이터를 전송해 주세요.'
  });
}

/**
 * 1. Project 진행현황 업데이트
 * - 시트명: '프로젝트 현황'
 * - 입력값: projectName(프로젝트명), status(진행 상태: 컨셉/디자인/개발/판매), details(세부 내용)
 * - 로직:
 *   F열(6번째 열, Project)을 검색하여 일치하는 첫 번째 행을 찾음 ->
 *   해당 행의 J열(10번째 열, Status)을 진행 상태로 덮어씀 ->
 *   O열(15번째 열, 금주 주간현황)의 기존 데이터에 줄바꿈(\n)을 하여 세부 내용을 이어 붙임
 */
function handleUpdateProject(ss, data) {
  const sheetName = '프로젝트 현황';
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return createJsonResponse({
      success: false,
      message: `'${sheetName}' 시트를 찾을 수 없습니다.`
    });
  }

  const targetProject = String(data.projectName || '').trim();
  const newStatus = String(data.status || '').trim();
  const newDetails = String(data.details || '').trim();

  if (!targetProject) {
    return createJsonResponse({
      success: false,
      message: '프로젝트명을 입력해 주세요.'
    });
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 1) {
    return createJsonResponse({
      success: false,
      message: `'${sheetName}' 시트에 데이터가 없습니다.`
    });
  }

  // F열(6번째 열) 전체 값 조회 (1행 ~ 마지막 행)
  const fColValues = sheet.getRange(1, 6, lastRow, 1).getValues();
  let targetRowIndex = -1;

  for (let i = 0; i < fColValues.length; i++) {
    const cellText = String(fColValues[i][0] || '').trim();
    if (cellText === targetProject) {
      targetRowIndex = i + 1; // 1-based 행 인덱스 (일치하는 첫 번째 행)
      break;
    }
  }

  if (targetRowIndex === -1) {
    return createJsonResponse({
      success: false,
      message: `'${sheetName}' 시트의 F열에서 '${targetProject}' 프로젝트를 찾을 수 없습니다.`
    });
  }

  // J열(10번째 열, Status): 전달받은 진행 상태로 덮어쓰기
  if (newStatus) {
    sheet.getRange(targetRowIndex, 10).setValue(newStatus);
  }

  // O열(15번째 열, 금주 진행현황): 기존 데이터에 줄바꿈(\n)하여 세부 내용 이어 붙이기
  let updatedDetails = '';
  if (newDetails) {
    const detailCell = sheet.getRange(targetRowIndex, 15);
    const existingText = String(detailCell.getValue() || '').trim();
    updatedDetails = (existingText && existingText !== '-')
      ? (existingText + '\n' + newDetails)
      : newDetails;
    detailCell.setValue(updatedDetails);
  }

  SpreadsheetApp.flush();

  return createJsonResponse({
    success: true,
    message: `'${targetProject}' 프로젝트(${targetRowIndex}행)의 Status(J열) 및 세부 내용(O열)이 업데이트되었습니다.`,
    updatedRow: targetRowIndex,
    projectName: targetProject,
    status: newStatus,
    details: updatedDetails
  });
}

/**
 * 2. Issue Tracker (새 행 추가)
 * - 시트명: 'IssueTracker' (없으면 자동 생성 및 헤더 추가)
 * - 입력값: issueDate(발생 일자), projectName(관련 모델/프로젝트명), issueContent(이슈 내용)
 * - 로직: 시트 가장 아래에 appendRow로 추가
 */
function handleAddIssue(ss, data) {
  const sheetName = 'IssueTracker';
  let sheet = ss.getSheetByName(sheetName);

  // 시트가 없을 경우 자동 생성 및 기본 헤더 설정
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['발생 일자', '관련 모델/프로젝트명', '이슈 내용', '등록 일시']);
  }

  const issueDate = String(data.issueDate || '').trim();
  const projectName = String(data.projectName || '').trim();
  const issueContent = String(data.issueContent || '').trim();
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

  sheet.appendRow([issueDate, projectName, issueContent, timestamp]);
  SpreadsheetApp.flush();

  return createJsonResponse({
    success: true,
    message: `'${projectName}' 관련 이슈가 IssueTracker 시트에 추가되었습니다.`
  });
}

/**
 * 3. 회의록 입력 (새 행 추가)
 * - 시트명: 'MeetingMinutes' (없으면 자동 생성 및 헤더 추가)
 * - 입력값: meetingDate(일자), attendees(참석자), meetingContent(회의 내용)
 * - 로직: 시트 가장 아래에 appendRow로 추가
 */
function handleAddMeeting(ss, data) {
  const sheetName = 'MeetingMinutes';
  let sheet = ss.getSheetByName(sheetName);

  // 시트가 없을 경우 자동 생성 및 기본 헤더 설정
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['일자', '참석자', '회의 내용', '등록 일시']);
  }

  const meetingDate = String(data.meetingDate || '').trim();
  const attendees = String(data.attendees || '').trim();
  const meetingContent = String(data.meetingContent || '').trim();
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

  sheet.appendRow([meetingDate, attendees, meetingContent, timestamp]);
  SpreadsheetApp.flush();

  return createJsonResponse({
    success: true,
    message: `${meetingDate} 회의록이 MeetingMinutes 시트에 저장되었습니다.`
  });
}

/**
 * JSON 응답 생성 헬퍼
 */
function createJsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
