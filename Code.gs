/**
 * 프로젝트 현황판 (project-status-dashboard) 통합 Google Apps Script (Code.gs)
 * 대상 스프레드시트 ID: 1DJHA-GBPy6PCG7ghXvOBI9YmSq93SDlp-oaOtibVPg0
 *
 * [지원 기능]
 * - doGet(e)                 : GAS 웹앱 직접 접속 시 Index.html 렌더링 (또는 API 상태 JSON 반환)
 * - doPost(e) 분기 (action)  :
 *   1) updateProject         : '프로젝트 현황' 시트 F열(Project) 검색 -> 첫 번째 일치 행의 J열(Status) 덮어쓰기, O열(금주 현황) 줄바꿈(\n) 추가
 *   2) addIssue              : 'IssueTracker' 시트 하단에 새 행 추가 (appendRow)
 *   3) addMeeting            : 'MeetingMinutes' 시트 하단에 새 행 추가 (appendRow)
 *   4) getLiveSheetData      : '프로젝트 현황' 시트 최신 데이터 동기화 조회
 */

const SHEET_ID = '1DJHA-GBPy6PCG7ghXvOBI9YmSq93SDlp-oaOtibVPg0';
const SHEET_NAME = '프로젝트 현황';

function getTargetSpreadsheet() {
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (_) {}
  return SpreadsheetApp.openById(SHEET_ID);
}

/**
 * 웹앱 접속(GET) 시 Index.html이 있으면 대시보드를 렌더링하고, 없으면 API 안내 JSON을 반환합니다.
 */
function doGet(e) {
  try {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('프로젝트 현황판 (Project Status Dashboard)')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (_) {
    return createJsonResponse({
      success: true,
      message: 'project-status-dashboard GAS 백엔드 API가 정상 작동 중입니다.'
    });
  }
}

/**
 * GitHub Pages 등 외부 정적 웹 대시보드에서 보내는 비동기 POST 요청 처리 (text/plain;charset=utf-8)
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
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
      case 'getLiveSheetData':
        return createJsonResponse(getLiveSheetData(payload.password));
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
 * 1. Project 진행현황 업데이트
 * - 시트명: '프로젝트 현황'
 * - 입력값: projectName(프로젝트명), status(진행 상태), details(세부 내용)
 * - 로직: F열(6번째 열)에서 프로젝트명을 검색하여 일치하는 첫 번째 행을 찾고,
 *         J열(10번째 열, Status)은 덮어쓰고, O열(15번째 열)은 기존 내용에 줄바꿈(\n)하여 이어 붙입니다.
 */
function handleUpdateProject(ss, data) {
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    return createJsonResponse({
      success: false,
      message: `'${SHEET_NAME}' 시트를 찾을 수 없습니다.`
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
      message: `'${SHEET_NAME}' 시트에 데이터가 없습니다.`
    });
  }

  // F열(6번째 열) 전체 값 조회 (1행 ~ 마지막 행)
  const fColValues = sheet.getRange(1, 6, lastRow, 1).getValues();
  let targetRowIndex = -1;

  for (let i = 0; i < fColValues.length; i++) {
    const cellText = String(fColValues[i][0] || '').replace(/\n+/g, ' ').trim();
    const rawCellText = String(fColValues[i][0] || '').trim();
    if (rawCellText === targetProject || cellText === targetProject) {
      targetRowIndex = i + 1; // 일치하는 첫 번째 행 (1-based)
      break;
    }
  }

  if (targetRowIndex === -1) {
    return createJsonResponse({
      success: false,
      message: `'${SHEET_NAME}' 시트의 F열에서 '${targetProject}' 프로젝트를 찾을 수 없습니다.`
    });
  }

  // J열(10번째 열, Status) 덮어쓰기
  if (newStatus) {
    sheet.getRange(targetRowIndex, 10).setValue(newStatus);
  }

  // O열(15번째 열) 기존 데이터에 줄바꿈(\n) 후 세부 내용 이어 붙이기
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
    message: `'${targetProject}' 프로젝트(${targetRowIndex}행)의 진행 상태(J열) 및 세부 내용(O열)이 업데이트되었습니다.`,
    updatedRow: targetRowIndex,
    projectName: targetProject,
    status: newStatus,
    details: updatedDetails
  });
}

/**
 * 2. Issue Tracker (새 행 추가)
 * - 시트명: 'IssueTracker'
 * - 입력값: 발생 일자(issueDate), 관련 모델/프로젝트명(projectName), 이슈 내용(issueContent)
 * - 로직: 시트 가장 아래에 새로운 행(appendRow) 추가
 */
function handleAddIssue(ss, data) {
  const sheetName = 'IssueTracker';
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['발생 일자', '관련 모델/프로젝트명', '이슈 내용']);
  }

  const issueDate = String(data.issueDate || '').trim();
  const projectName = String(data.projectName || '').trim();
  const issueContent = String(data.issueContent || '').trim();

  sheet.appendRow([issueDate, projectName, issueContent]);
  SpreadsheetApp.flush();

  return createJsonResponse({
    success: true,
    message: `'${projectName}' 관련 이슈가 IssueTracker 시트에 등록되었습니다.`
  });
}

/**
 * 3. 회의록 입력 (새 행 추가)
 * - 시트명: 'MeetingMinutes'
 * - 입력값: 일자(meetingDate), 참석자(attendees), 회의 내용(meetingContent)
 * - 로직: 시트 가장 아래에 새로운 행(appendRow) 추가
 */
function handleAddMeeting(ss, data) {
  const sheetName = 'MeetingMinutes';
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['일자', '참석자', '회의 내용']);
  }

  const meetingDate = String(data.meetingDate || '').trim();
  const attendees = String(data.attendees || '').trim();
  const meetingContent = String(data.meetingContent || '').trim();

  sheet.appendRow([meetingDate, attendees, meetingContent]);
  SpreadsheetApp.flush();

  return createJsonResponse({
    success: true,
    message: `${meetingDate} 회의록이 MeetingMinutes 시트에 저장되었습니다.`
  });
}

/**
 * 실시간 구글 시트 전체 동기화 데이터 조회 (웹앱 및 GitHub Pages 공용)
 */
function getLiveSheetData(password) {
  if (password !== 'diofit88*') {
    return { ok: false, success: false, error: '보안 암호가 일치하지 않습니다.' };
  }

  const ss = getTargetSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(16, sheet.getLastColumn());
  const values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();

  // 병합 셀(KSF, 44W, 43W 등) 값 채우기
  const mergedRanges = sheet.getRange(4, 1, Math.max(1, lastRow - 3), lastCol).getMergedRanges();
  mergedRanges.forEach(function(rng) {
    const topVal = rng.getDisplayValue();
    const startR = rng.getRow();
    const endR = rng.getLastRow();
    const startC = rng.getColumn();
    const endC = rng.getLastColumn();
    if (startC === endC && topVal) {
      for (let r = startR; r <= endR; r++) {
        if (!values[r - 1][startC - 1]) {
          values[r - 1][startC - 1] = topVal;
        }
      }
    }
  });

  const weekHeader = (values[2][14] || '44W (10/26~)').trim();
  const prevWeekHeader = (values[2][15] || '43W (10/19~)').trim();
  const rowsOut = [];
  let currentGroupKey = '';
  let currentGidNum = 0;
  let lastStatus = '판매';

  for (let r = 3; r < values.length; r++) {
    const row = values[r];
    const seq = (row[0] || '').trim();
    const noRaw = (row[1] || '').trim();
    const start = (row[2] || '').trim();
    const end = (row[3] || '').trim();
    const buyer = (row[4] || '').trim();
    const project = (row[5] || '').replace(/\n+/g, ' ').trim();
    const item = (row[6] || '').replace(/\n+/g, ' ').trim();
    const model = (row[7] || '').trim();
    let status = (row[9] || '').trim();
    const prodDev = (row[10] || '').replace(/\n+/g, ' / ').trim();
    const pkgDev = (row[11] || '').replace(/\n+/g, ' / ').trim();
    const massProd = (row[12] || '').replace(/\n+/g, ' / ').trim();
    const ksf = (row[13] || '').split(/\r?\n/).map(s => s.trim()).filter(Boolean).join('\n');
    const w44 = (row[14] || '').split(/\r?\n/).map(s => s.trim()).filter(Boolean).join('\n');
    const w43 = (row[15] || '').split(/\r?\n/).map(s => s.trim()).filter(Boolean).join('\n');

    if (!buyer && !project && !model) continue;
    if (status) {
      lastStatus = status;
    } else {
      status = lastStatus;
    }

    const groupKey = [buyer, project, start, end, status].join('||');
    if (groupKey !== currentGroupKey) {
      currentGidNum++;
      currentGroupKey = groupKey;
    }

    rowsOut.push({
      seq: seq,
      no: parseInt(noRaw, 10) || (rowsOut.length + 1),
      start: start,
      end: end,
      buyer: buyer,
      project: project,
      item: item,
      model: model,
      image: '',
      status: status,
      prod_dev: prodDev,
      pkg_dev: pkgDev,
      mass_prod: massProd,
      dev: '제품:' + (prodDev || '-') + ' · 패키지:' + (pkgDev || '-') + ' · 양산:' + (massProd || '-'),
      ksf: ksf,
      w44: w44,
      w43: w43,
      group_id: 'G' + currentGidNum
    });
  }

  const now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
  return {
    ok: true,
    success: true,
    rows: rowsOut,
    week_header: weekHeader,
    prev_week_header: prevWeekHeader,
    synced_at: now
  };
}

function createJsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
