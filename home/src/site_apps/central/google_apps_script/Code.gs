/**
 * 한국지역난방공사 중앙지사 24시간 모바일 현장 입력 시스템 (Google Apps Script)
 * 컴퓨터 전원이 꺼져 있어도 스마트폰에서 언제든 저장 가능합니다.
 * 구글 드라이브 스프레드시트와 실시간 연동됩니다.
 */

// 1. 기본 설정 및 마스터 데이터
var APP_CONFIG = {
  companies: ["한국지역난방공사 중앙지사", "한국지역난방공사", "중앙지사"],
  loc_types: ["열수송관", "플랜트(관리소)", "열배관", "핫태핑"],
  users: [
    "부장 주진철", "부장 이병국", "부장 박광복", "차장 권인경",
    "과장 주영광", "대리 우명광", "계장 박상태", "계장 이준우",
    "계장 장승대", "주임 김성렬", "주임 김진환"
  ],
  vehicles: [
    "81두1580 (탑차)", "89보4028 (탑차)", "90너4889 (탑차)",
    "71고4405 (스타렉스)", "76마3422 (스타렉스)", "81도5958 (스타렉스)",
    "81도8458 (스타렉스)", "81루5100 (스타렉스)", "84저1259 (스타렉스)",
    "90서8616 (스타렉스)", "91주8839 (스타렉스)", "95가0175 (스타렉스)",
    "95가0200 (스타렉스)", "개인차량 (스타렉스)", "개인차량-이병국 (스타렉스)"
  ],
  equipments: ["PAUT MX2(OMNI2-103085)", "Se-75", "Yoke", "PT 약품"],
  aux_equipments: [
    "PAUT SCANNER (MANUAL)", "COBRA Manual Scanner", "JIREH MICROBE BTG006",
    "JIREH Scanner", "ADR, SURVEY, 콜리메타", "ADR; SURVEY; 콜리메타",
    "ADR", "SURVEY", "콜리메타", "자분살포기", "조도계", "자계강도계", "온습도계", "휴대용 UV메타", "없음"
  ],
  methods: ["PAUT", "RT", "MT", "PT"],
  pipe_sizes: [
    "1100A", "1000A", "900A", "850A", "800A", "750A", "700A", "650A",
    "600A", "550A", "500A", "450A", "400A", "350A", "300A", "250A",
    "200A", "150A", "125A", "100A", "80A", "65A", "50A", "40A",
    "32A", "25A", "20A"
  ],
  size_lengths: {
    '1100A': 3.511,  '1000A': 3.1919, '900A': 2.8727, '850A': 2.7131,
    '800A':  2.5535, '750A':  2.3939, '700A': 2.2343, '650A': 2.0747,
    '600A':  1.9151, '550A':  1.7555, '500A': 1.5959, '450A': 1.4363,
    '400A':  1.2767, '350A':  1.1172, '300A': 1.0006, '250A': 0.8401,
    '200A':  0.6795, '150A':  0.519,  '125A': 0.4392, '100A': 0.3591,
    '80A':   0.2799, '65A':   0.2397, '50A':  0.1901, '40A':  0.1527,
    '32A':   0.1341, '25A':   0.1068, '20A':  0.0855
  },
  base_prices: {
    "PAUT": 325451,
    "RT": 205412,
    "MT": 110000,
    "PT": 110000
  },
  units: ["M", "매", "EA", "P,M,I/D", "Point", "SET", "CAN", "KG", "Inch", "Dia", "I/D"],
  test_items: ["Pipe", "Piping", "Tank", "Vessel", "Plate", "Beam", "Structure", "기타"],
  applied_codes: ["지역난방시방서", "KS", "ASME", "AWS", "API", "KEPIC", "기타"],
  materials: [
    'Carestream AA400-3⅓*12"', 'Carestream AA400-3⅓*17"', 'Carestream AA400-4½*12"',
    'Carestream AA400-10*12"', 'Carestream M100-3⅓*12"', 'Carestream M100-10*12"',
    'Carestream M100-14*17"', 'Carestream MX125-3⅓*6"', 'Carestream MX125-3⅓*12"',
    'Carestream MX125-4½*12"', 'Carestream MX125-10*12"', 'Carestream MX125-14*17"',
    'Carestream T200-3⅓*12"', 'Carestream T200-3⅓*17"', 'Carestream T200-4½*12"',
    'Carestream T200-10*12"', 'Carestream T200-14*17"'
  ],
  worktimes: [
    "09:00~18:00", "09:00~19:00", "09:00~20:00", "09:00~21:00", "09:00~22:00",
    "09:00~23:00", "09:00~24:00", "09:00~익일01:00", "09:00~익일02:00", "09:00~익일03:00",
    "09:00~익일04:00", "09:00~익일05:00", "09:00~익일06:00", "09:00~익일07:00", "09:00~익일08:00", "09:00~익일09:00",
    "10:00~17:00", "10:00~19:00", "12:00~익일08:00", "13:00~22:00", "13:00~익일02:00",
    "18:00~익일01:00", "18:00~익일02:00", "18:00~익일03:00", "18:00~익일04:00", "18:00~익일05:00", "18:00~익일06:00", "18:00~익일07:00", "18:00~익일08:00",
    "20:00~익일02:00", "20:00~익일04:00", "20:00~익일05:00", "20:00~익일06:00",
    "21:00~익일03:00", "21:00~익일05:00", "21:00~익일06:00", "21:00~익일07:00",
    "22:00~익일01:00", "22:00~익일02:00", "22:00~익일03:00", "22:00~익일04:00", "22:00~익일05:00", "22:00~익일06:00", "22:00~익일07:00", "22:00~익일08:00", "22:00~익일09:00"
  ]
};

var HEADERS = [
  'Date', 'Site', 'MaterialID', 'Usage', 'Note', 'EntryTime',
  'RTK_센터미스', 'RTK_농도', 'RTK_마킹미스', 'RTK_필름마크', 'RTK_취급부주의', 'RTK_고객불만', 'RTK_기타',
  'User', '장비명', '검사량', '단가', '출장비', '일식', '검사비', '검사방법', 'FilmCount',
  'WorkTime', 'OT', 'User2', 'WorkTime2', 'OT2', 'User3', 'WorkTime3', 'OT3',
  'User4', 'WorkTime4', 'OT4', 'User5', 'WorkTime5', 'OT5', 'User6', 'WorkTime6', 'OT6',
  'User7', 'WorkTime7', 'OT7', 'User8', 'WorkTime8', 'OT8', 'User9', 'WorkTime9', 'OT9',
  'User10', 'WorkTime10', 'OT10',
  'NDT_형광자분', 'NDT_흑색자분', 'NDT_백색페인트', 'NDT_침투제', 'NDT_세척제', 'NDT_현상제', 'NDT_형광침투제',
  '차량번호', '주행거리', '차량점검', '차량비고', '품목명', 'RTK_총계', '회사코드', '업체명',
  '작업자', '작업시간', 'OT시간', 'OT금액', '입력시간', '(Full작업자)', 'RTK_RTK총계',
  '검사품명', '적용코드', '성적서번호', '검사자', 'Unit', 'Material', '작업형태',
  '조건1', '조건2', '제경비율', '기술료율', '보정계수', '환산물량', '재료비', '인건비', '제경비', '기술료',
  'Meal', 'Meal2', 'Meal3', 'Meal4', 'Meal5', 'Meal6', 'Meal7', 'Meal8', 'Meal9', 'Meal10',
  '검사구분', '조인트수', '불량수', '관경(Inch)', '구분', 'Remark', 'DailyMemo',
  '핫태핑차수', '핫태핑개소', '핫태핑관경', '보조장비', '수량', 'ClientRequestID', 'ClientRequestHash'
];

/**
 * 모바일 웹 페이지 접속 (doGet)
 */
function doGet(e) {
  var api = e && e.parameter ? e.parameter.api : null;

  if (api === 'config') {
    return createJsonResponse(getConfig());
  } else if (api === 'recent') {
    var date = e.parameter.date || '';
    return createJsonResponse(getRecentRecords(date));
  } else if (api === 'progress') {
    var date = e.parameter.date || '';
    return createJsonResponse(getProgressSummary(date));
  } else if (api === 'sync' || api === 'all_records') {
    return createJsonResponse(getAllRecords());
  }

  var htmlOutput = HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('한국지역난방공사 중앙지사 모바일 현장 입력 (24시간 클라우드)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  return htmlOutput;
}

function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    return "";
  }
}

/**
 * 작업 저장 요청 처리 (doPost)
 */
function doPost(e) {
  try {
    var payload;
    if (e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else {
      payload = e.parameter;
    }

    var action = payload.action || 'save';
    if (action === 'batch_sync' || action === 'batch') {
      var res = batchSyncRecords(payload);
      return createJsonResponse(res);
    } else if (action === 'update' || action === 'updateRecord' || (payload.id !== undefined && payload.is_update)) {
      var res = updateRecord(payload);
      return createJsonResponse(res);
    } else if (action === 'delete' || action === 'deleteRecord') {
      var res = deleteRecord(payload);
      return createJsonResponse(res);
    }

    var res = saveRecord(payload);
    return createJsonResponse(res);
  } catch (err) {
    return createJsonResponse({ status: 'error', detail: err.toString() });
  }
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 활성 스프레드시트의 'DailyUsage' 시트 획득 또는 자동 생성
 */
function getDailyUsageSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("DailyUsage");
  if (!sheet) {
    sheet = ss.insertSheet("DailyUsage");
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#e8f0fe");
    sheet.setFrozenRows(1);
  } else {
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#e8f0fe");
      sheet.setFrozenRows(1);
    }
  }
  var current=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var missing=HEADERS.filter(function(name){return current.indexOf(name)<0;});
  if(missing.length)sheet.getRange(1,current.length+1,1,missing.length).setValues([missing]);
  return sheet;
}

/**
 * 설정 데이터 획득
 */
function getConfig() {
  var managed = typeof CENTRAL_REFERENCE !== 'undefined' ? CENTRAL_REFERENCE.config : {};
  var cfg = JSON.parse(JSON.stringify(APP_CONFIG));
  Object.keys(managed).forEach(function(key) {cfg[key] = managed[key];});
  return cfg;
}

// Historical company names are available to inquiry, not new-entry choices.
function getHistoricalCompanies() {
  var companies = getConfig().companies.slice();
  try {
    var sheet = getDailyUsageSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var compColIdx = headerRow.indexOf("업체명");
      if (compColIdx !== -1) {
        var values = sheet.getRange(2, compColIdx + 1, lastRow - 1, 1).getValues();
        for (var i = 0; i < values.length; i++) {
          var c = String(values[i][0]).trim();
          if (c && c !== 'nan' && companies.indexOf(c) === -1) {
            companies.push(c);
          }
        }
      }
    }
  } catch (err) {
    // fallback to default config if sheet read fails
  }

  return companies;
}

/**
 * 작업 레코드 맵 구성 (신규 및 수정 공용)
 */
function buildRowMap(req, preserveEntryTime) {
  var nowStr = preserveEntryTime || (req.entry_time ? String(req.entry_time) : Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss.SSS"));

  // PC 프로그램에서 직접 보낸 row_data 맵이 있는 경우 그대로 매핑
  if (req.row_data && typeof req.row_data === 'object') {
    var directMap = {};
    for (var k in req.row_data) {
      if (req.row_data.hasOwnProperty(k)) {
        directMap[k] = req.row_data[k];
      }
    }
    directMap['EntryTime'] = nowStr;
    directMap['입력시간'] = nowStr;
    return directMap;
  }

  var calculatedFee = Number(req.total_fee || 0);
  if (calculatedFee <= 0) {
    calculatedFee = Math.trunc((Number(req.qty || 0) * Number(req.unit_price || 0)) + Number(req.travel_cost || 0));
  }

  var rowMap = {
    'Date': req.date || Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd"),
    'Site': req.site || '중앙지사',
    '작업구분': req.loc_type || '열수송관',
    '구분': req.loc_type || '열수송관',
    '검사구분': req.loc_type || '열수송관',
    '업체명': req.company || '한국지역난방공사 중앙지사',
    '회사코드': req.company_code || '한국지역난방공사',
    '검사방법': req.method || 'PAUT',
    '규격': req.pipe || '',
    '관경(Inch)': req.pipe || '',
    '검사품명': req.item_name || 'Pipe',
    '자재명': req.material || '',
    'MaterialID': req.material_id || '',
    'ClientRequestID': req.client_request_id || '',
    'ClientRequestHash': req._request_hash || '',
    'Material': req.material || '',
    '품목명': req.material || '',
    '적용코드': req.applied_code || '지역난방시방서',
    '성적서번호': req.report_no || '',
    '조인트수': Number(req.joints || req.ori_joints || 0),
    '불량수': Number(req.rej_qty || 0),
    '수량': Number(req.qty || 0),
    '검사량': Number(req.qty || 0),
    'Usage': Number(req.qty || 0),
    '단위': req.unit || 'M',
    'Unit': req.unit || 'M',
    '단가': Number(req.unit_price || 0),
    '출장비': Number(req.travel_cost || 0),
    '일식': req.lump_sum || '',
    '검사비': calculatedFee,
    '장비명': req.equip_main || '',
    '보조장비': req.equip_aux || '',
    '작업형태': req.work_shape || '주간',
    '차량번호': req.vehicle_no || '',
    '주행거리': req.mileage || '',
    '차량점검': '양호',
    '차량비고': req.vehicle_note || '',
    'Note': req.note || '',
    'Remark': req.note || '',
    'DailyMemo': req.note || '',
    'EntryTime': nowStr,
    '입력시간': nowStr,
    'FilmCount': (req.method === 'RT') ? Number(req.qty || 0) : 0,
    // 지역난방공사 단가계약 비율
    '제경비율': 1.10,
    '기술료율': 0.20,
    // 핫태핑
    '핫태핑차수': req.hot_tap_stage || '',
    '핫태핑개소': req.hot_tap_joints || '',
    '핫태핑관경': req.hot_tap_pipe || '',
    // RTK Defect Categories (8)
    'RTK_센터미스': Number(req.rtk_center_miss || 0),
    'RTK_농도': Number(req.rtk_density || 0),
    'RTK_마킹미스': Number(req.rtk_marking_miss || 0),
    'RTK_필름마크': Number(req.rtk_film_mark || 0),
    'RTK_취급부주의': Number(req.rtk_handling || 0),
    'RTK_고객불만': Number(req.rtk_complaint || 0),
    'RTK_기타': Number(req.rtk_etc || 0),
    'RTK_총계': Number(req.rtk_total || 0),
    'RTK_RTK총계': Number(req.rtk_total || 0),
    // NDT Chemicals (7)
    'NDT_형광자분': Number(req.chem_fl_mag || 0),
    'NDT_흑색자분': Number(req.chem_black_mag || 0),
    'NDT_백색페인트': Number(req.chem_white_paint || 0),
    'NDT_침투제': Number(req.chem_penetrant || 0),
    'NDT_세척제': Number(req.chem_cleaner || 0),
    'NDT_현상제': Number(req.chem_developer || 0),
    'NDT_형광침투제': Number(req.chem_fl_pen || 0)
  };

  // Reset worker columns
  for (var u = 1; u <= 10; u++) {
    var uKey = (u === 1) ? 'User' : ('User' + u);
    var wtKey = (u === 1) ? 'WorkTime' : ('WorkTime' + u);
    var otKey = (u === 1) ? 'OT' : ('OT' + u);
    rowMap[uKey] = '';
    rowMap[wtKey] = '';
    rowMap[otKey] = 0;
    rowMap['Meal' + (u === 1 ? '' : u)] = 0;
  }

  // 참여 작업자 할당 (1 ~ 10)
  if (req.workers && req.workers.length > 0) {
    rowMap['User'] = req.workers[0].user || '';
    rowMap['검사자'] = req.workers[0].user || '';
    rowMap['작업자'] = req.workers[0].user || '';

    var allWorkersDesc = [];
    for (var i = 0; i < req.workers.length && i < 10; i++) {
      var w = req.workers[i];
      var idx = i + 1;
      var uKey = (idx === 1) ? 'User' : ('User' + idx);
      var wtKey = (idx === 1) ? 'WorkTime' : ('WorkTime' + idx);
      var otKey = (idx === 1) ? 'OT' : ('OT' + idx);
      rowMap[uKey] = w.user || '';
      rowMap[wtKey] = (w.shift ? '(' + w.shift + ') ' : '') + (w.time || '');
      rowMap[otKey] = Number(String(w.ot || '0').replace(/,/g, ''));
      rowMap['Meal' + (idx === 1 ? '' : idx)] = Number(String(w.meal || '0').replace(/,/g, ''));
      allWorkersDesc.push((w.user || '') + '(' + (w.time || '') + ')');
    }
    rowMap['(Full작업자)'] = allWorkersDesc.join(', ');
  }

  return rowMap;
}

/**
 * 작업 레코드 시트에 저장
 */
function saveRecordCore(req) {
  var sheet = getDailyUsageSheet();
  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var rowMap = buildRowMap(req);

  var newRow = [];
  for (var h = 0; h < headerRow.length; h++) {
    var headerName = headerRow[h];
    newRow.push(rowMap.hasOwnProperty(headerName) ? rowMap[headerName] : "");
  }

  sheet.appendRow(newRow);

  return {
    status: 'success',
    message: '중앙지사 구글 클라우드에 성공적으로 저장되었습니다.'
  };
}

function formatCellTime(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss.SSS");
  }
  return String(val).trim();
}

/**
 * 작업 레코드 수정
 */
function updateRecordCore(req) {
  var sheet = getDailyUsageSheet();
  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var entryTimeCol = headerRow.indexOf("EntryTime");
  var lastRow = sheet.getLastRow();
  var targetRow = findMobileRow(sheet, headerRow, req);

  if (targetRow === -1) {
    return { status: 'error', detail: '수정할 일보 내역을 찾을 수 없습니다.' };
  }

  validateMobileRecord(req);
  var oldRow = sheet.getRange(targetRow,1,1,headerRow.length).getValues()[0];
  var rowMap = buildRowMap(req, formatCellTime(oldRow[entryTimeCol]));
  rowMap.ClientRequestID = oldRow[headerRow.indexOf('ClientRequestID')] || req.client_request_id || '';
  var rowValues = [];
  for (var j = 0; j < headerRow.length; j++) {
    var colName = headerRow[j];
    rowValues.push(rowMap.hasOwnProperty(colName) ? rowMap[colName] : '');
  }
  sheet.getRange(targetRow, 1, 1, headerRow.length).setValues([rowValues]);
  return { status: 'success', message: '성공적으로 수정되었습니다.' };
}

/**
 * 작업 레코드 삭제
 */
function deleteRecordCore(req) {
  var sheet = getDailyUsageSheet();
  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var entryTimeCol = headerRow.indexOf("EntryTime");
  var lastRow = sheet.getLastRow();
  var targetRow = findMobileRow(sheet, headerRow, req);

  if (targetRow === -1) {
    return { status: 'error', detail: '삭제할 일보 내역을 찾을 수 없습니다.' };
  }

  sheet.deleteRow(targetRow);
  return { status: 'success', message: '성공적으로 삭제되었습니다.' };
}

/**
 * PC 프로그램 일괄 동기화 (추가, 수정, 삭제 일괄 처리)
 */
function batchSyncRecords(payload) {
  var sheet = getDailyUsageSheet();
  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var entryTimeCol = headerRow.indexOf("EntryTime");

  var deletesCount = 0;
  var updatesCount = 0;
  var savesCount = 0;

  // 1. 삭제 처리
  var deletes = payload.deletes || [];
  if (deletes.length > 0 && entryTimeCol !== -1) {
    for (var d = 0; d < deletes.length; d++) {
      var targetTime = String(deletes[d]).split('.')[0].trim();
      var curLastRow = sheet.getLastRow();
      if (curLastRow <= 1) break;
      var timeVals = sheet.getRange(2, entryTimeCol + 1, curLastRow - 1, 1).getValues();
      for (var r = timeVals.length - 1; r >= 0; r--) {
        var cellTime = formatCellTime(timeVals[r][0]);
        if (cellTime === targetTime) {
          sheet.deleteRow(r + 2);
          deletesCount++;
          break;
        }
      }
    }
  }

  // 2. 수정 처리
  var updates = payload.updates || [];
  for (var u = 0; u < updates.length; u++) {
    var resU = updateRecord(updates[u]);
    if (resU.status === 'success') updatesCount++;
  }

  // 3. 신규 추가 처리
  var saves = payload.saves || [];
  for (var s = 0; s < saves.length; s++) {
    var resS = saveRecord(saves[s]);
    if (resS.status === 'success') savesCount++;
  }

  return {
    status: 'success',
    added_count: savesCount,
    updated_count: updatesCount,
    deleted_count: deletesCount,
    message: '일괄 동기화 완료 (추가: ' + savesCount + '건, 수정: ' + updatesCount + '건, 삭제: ' + deletesCount + '건)'
  };
}

function safeNum(val, defaultVal) {
  if (val === undefined || val === null || val === '') return (defaultVal !== undefined ? defaultVal : 0);
  var n = Number(String(val).replace(/,/g, '').trim());
  return isNaN(n) ? (defaultVal !== undefined ? defaultVal : 0) : n;
}

/**
 * 최근 저장된 내역 조회 (당일 또는 최근 내역 자동 표시)
 */
function getRecentRecords(targetDate) {
  var sheet = getDailyUsageSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return { records: [], is_fallback: false };
  }

  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var startRow = Math.max(2, lastRow - 30);
  var numRows = lastRow - startRow + 1;
  var data = sheet.getRange(startRow, 1, numRows, headerRow.length).getValues();


  var matchedRecords = [];
  var allRecent = [];

  for (var r = 0; r < data.length; r++) {
    var rec = mobileRecordFromRow(headerRow, data[r], startRow + r);
    allRecent.push(rec);
    if (!targetDate || rec.date === targetDate) {
      matchedRecords.push(rec);
    }
  }

  var isFallback = false;
  var finalRecords = matchedRecords;
  if (matchedRecords.length === 0 && targetDate && allRecent.length > 0) {
    finalRecords = allRecent.slice(-10);
    isFallback = true;
  }

  finalRecords.reverse();
  return {
    records: finalRecords.slice(0, 15),
    is_fallback: isFallback
  };
}

/**
 * 전체 저장 데이터 조회 (중앙지사 PC 프로그램 동기화용 API)
 */
function getAllRecords() {
  var sheet = getDailyUsageSheet();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow <= 1) {
    return { status: 'success', count: 0, headers: HEADERS, rows: [] };
  }

  var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = data[0];
  var rows = [];

  for (var r = 1; r < data.length; r++) {
    var rowObj = {};
    for (var c = 0; c < headers.length; c++) {
      var val = data[r][c];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss.SSS");
      }
      rowObj[headers[c]] = val;
    }
    rows.push(rowObj);
  }

  return {
    status: 'success',
    count: rows.length,
    headers: headers,
    rows: rows
  };
}

