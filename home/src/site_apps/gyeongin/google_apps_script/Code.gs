/**
 * 경인사무소 24시간 모바일 현장 입력 시스템 (Google Apps Script)
 * 컴퓨터 전원이 꺼져 있어도 스마트폰에서 언제든 저장 가능합니다.
 * 구글 드라이브 스프레드시트와 실시간 연동됩니다.
 */

// 1. 기본 설정 및 마스터 데이터
var APP_CONFIG = {
  companies: ["경인사무소"],
  users: [
    "부장 주진철", "대리 우명광", "주임 김진환", "계장 장승대",
    "주임 김성렬", "부장 박광복", "과장 주영광", "계장 박상태"
  ],
  vehicles: [
    "경인 1호차 (스타렉스)", "경인 2호차 (카니발)", "경인 3호차 (포터)", "자가용/대중교통"
  ],
  equipments: ["Se-75", "PAUT MX2(OMNI2-103085)", "Yoke", "PT 약품"],
  aux_equipments: [
    "ADR, SURVEY, 콜리메타", "ADR; SURVEY; 콜리메타", "ADR", "SURVEY", "콜리메타",
    "PAUT SCANNER (MANUAL)", "COBRA Manual Scanner", "JIREH MICROBE BTG006",
    "JIREH Scanner", "자분살포기", "조도계", "자계강도계", "온습도계", "휴대용 UV메타", "없음"
  ],
  pipe_sizes: [
    "15A", "20A", "25A", "32A", "40A", "50A", "65A", "80A", "90A", "100A",
    "125A", "150A", "200A", "250A", "300A", "350A", "400A", "450A", "500A",
    "550A", "600A", "650A", "700A", "750A", "800A", "850A", "900A", "950A",
    "1000A", "1050A", "1100A", "1150A", "1200A"
  ],
  units: ["매", "M", "P,M,I/D", "EA", "Point", "SET", "CAN", "KG", "Inch", "Dia", "I/D", "Meter"],
  test_items: ["Piping", "Tank", "Vessel", "Plate", "Beam", "Structure", "Pipe", "기타"],
  applied_codes: ["KS", "ASME", "AWS", "API", "KEPIC", "기타"],
  materials: [
    'Carestream AA400-3⅓*12"', 'Carestream AA400-3⅓*17"', 'Carestream AA400-4½*12"',
    'Carestream AA400-10*12"', 'Carestream M100-3⅓*12"', 'Carestream M100-10*12"',
    'Carestream M100-14*17"', 'Carestream MX125-3⅓*6"', 'Carestream MX125-3⅓*12"',
    'Carestream MX125-4½*12"', 'Carestream MX125-10*12"', 'Carestream MX125-14*17"',
    'Carestream T200-3⅓*12"', 'Carestream T200-3⅓*17"', 'Carestream T200-4½*12"',
    'Carestream T200-10*12"', 'Carestream T200-14*17"'
  ],
  worktimes: [
    "09:00~18:00", "09:00~19:00", "09:00~20:00", "09:00~21:00",
    "09:00~22:00", "09:00~23:00", "09:00~24:00", "09:00~익일01:00",
    "09:00~익일02:00", "09:00~익일03:00", "18:00~익일02:00",
    "18:00~익일03:00", "22:00~익일09:00", "12:00~익일08:00"
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
  '검사품명', '적용코드', '성적서번호', 'Unit', 'Material', '작업형태',
  '검사구분', '조인트수', '불량수', '관경(Inch)', '구분', 'Remark', 'DailyMemo', '보조장비', '수량'
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
  } else if (api === 'sync' || api === 'all_records') {
    return createJsonResponse(getAllRecords());
  }

  // 모바일 웹 화면 반환 (HTML)
  var htmlOutput = HtmlService.createHtmlOutputFromFile('index')
    .setTitle('경인사무소 모바일 현장 입력 (24시간 클라우드)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
  return htmlOutput;
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
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#e6f4ea");
    sheet.setFrozenRows(1);
  } else {
    // 헤더가 비어있으면 헤더 작성
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#e6f4ea");
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

/**
 * 설정 데이터 획득 (google.script.run 및 API 공용)
 */
function getConfig() {
  var sheet = getDailyUsageSheet();
  var companies = APP_CONFIG.companies.slice();
  
  // 시트에 등록된 기존 업체명 동적 수집
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

  var cfg = JSON.parse(JSON.stringify(APP_CONFIG));
  cfg.companies = companies;
  return cfg;
}

/**
 * 작업 레코드 시트에 저장
 */
function saveRecord(req) {
  var sheet = getDailyUsageSheet();
  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var nowStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
  var calculatedFee = Number(req.total_fee || 0);
  if (calculatedFee <= 0) {
    calculatedFee = Math.round((Number(req.qty || 0) * Number(req.unit_price || 0)) + Number(req.travel_cost || 0));
  }

  var rowMap = {
    'Date': req.date || Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd"),
    'Site': req.site || '경인사무소',
    '작업구분': req.loc_type || '경인사무소',
    '구분': req.loc_type || '경인사무소',
    '검사구분': req.loc_type || '경인사무소',
    '업체명': req.company || '경인사무소',
    '회사코드': req.company_code || req.company || '경인사무소',
    '검사방법': req.method || 'RT',
    '규격': req.pipe || '',
    '관경(Inch)': req.pipe || '',
    '검사품명': req.item_name || 'Piping',
    '자재명': req.material || '',
    'MaterialID': req.material || '',
    'Material': req.material || '',
    '품목명': req.material || '',
    '적용코드': req.applied_code || 'KS',
    '성적서번호': req.report_no || '',
    '조인트수': req.joints || req.ori_joints || 0,
    '불량수': req.rej_qty || 0,
    '수량': Number(req.qty || 0),
    '검사량': Number(req.qty || 0),
    'Usage': Number(req.qty || 0),
    '단위': req.unit || '매',
    'Unit': req.unit || '매',
    '단가': Number(req.unit_price || 0),
    '출장비': Number(req.travel_cost || 0),
    '일식': req.lump_sum || '',
    '검사비': calculatedFee,
    '장비명': req.equip_main || '',
    '보조장비': req.equip_aux || '',
    '작업형태': req.work_shape || '일반',
    '차량번호': req.vehicle_no || '',
    '주행거리': req.mileage || '',
    '차량점검': '양호',
    '차량비고': req.vehicle_note || '',
    'Note': req.note || '',
    'Remark': req.note || '',
    'DailyMemo': req.note || '',
    'EntryTime': nowStr,
    'FilmCount': (req.method === 'RT') ? Number(req.qty || 0) : 0,
    // RTK (8)
    'RTK_센터미스': Number(req.rtk_center_miss || 0),
    'RTK_농도': Number(req.rtk_density || 0),
    'RTK_마킹미스': Number(req.rtk_marking_miss || 0),
    'RTK_필름마크': Number(req.rtk_film_mark || 0),
    'RTK_취급부주의': Number(req.rtk_handling || 0),
    'RTK_고객불만': Number(req.rtk_complaint || 0),
    'RTK_기타': Number(req.rtk_etc || 0),
    'RTK_총계': Number(req.rtk_total || 0),
    // NDT Chemicals (7)
    'NDT_형광자분': Number(req.chem_fl_mag || 0),
    'NDT_흑색자분': Number(req.chem_black_mag || 0),
    'NDT_백색페인트': Number(req.chem_white_paint || 0),
    'NDT_침투제': Number(req.chem_penetrant || 0),
    'NDT_세척제': Number(req.chem_cleaner || 0),
    'NDT_현상제': Number(req.chem_developer || 0),
    'NDT_형광침투제': Number(req.chem_fl_pen || 0)
  };

  // 참여 작업자 할당 (1 ~ 10)
  if (req.workers && req.workers.length > 0) {
    rowMap['User'] = req.workers[0].user || '';
    for (var i = 0; i < req.workers.length && i < 10; i++) {
      var w = req.workers[i];
      var idx = i + 1;
      var uKey = (idx === 1) ? 'User' : ('User' + idx);
      var wtKey = (idx === 1) ? 'WorkTime' : ('WorkTime' + idx);
      var otKey = (idx === 1) ? 'OT' : ('OT' + idx);
      rowMap[uKey] = w.user || '';
      rowMap[wtKey] = w.time || '';
      var cleanOt = Number(String(w.ot || 0).replace(/,/g, '').trim()) || 0;
      rowMap[otKey] = cleanOt;
    }
  }

  // 헤더 순서대로 배열 생성
  var newRow = [];
  for (var h = 0; h < headerRow.length; h++) {
    var headerName = headerRow[h];
    newRow.push(rowMap.hasOwnProperty(headerName) ? rowMap[headerName] : "");
  }

  sheet.appendRow(newRow);

  return {
    status: 'success',
    message: '구글 드라이브 클라우드에 성공적으로 저장되었습니다.'
  };
}

/**
 * 최근 저장된 내역 조회 (당일 또는 최근 15건)
 */
function getRecentRecords(targetDate) {
  var sheet = getDailyUsageSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return { records: [] };
  }

  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var startRow = Math.max(2, lastRow - 30);
  var numRows = lastRow - startRow + 1;
  var data = sheet.getRange(startRow, 1, numRows, headerRow.length).getValues();

  var records = [];
  for (var r = 0; r < data.length; r++) {
    var row = data[r];
    var item = {};
    for (var c = 0; c < headerRow.length; c++) {
      item[headerRow[c]] = row[c];
    }

    var rowDateStr = "";
    if (item['Date'] instanceof Date) {
      rowDateStr = Utilities.formatDate(item['Date'], "Asia/Seoul", "yyyy-MM-dd");
    } else {
      rowDateStr = String(item['Date'] || '').substring(0, 10);
    }

    if (targetDate && rowDateStr !== targetDate) {
      continue;
    }

    // 작업자 요약
    var workersArr = [];
    for (var u = 1; u <= 10; u++) {
      var uName = item[(u === 1) ? 'User' : ('User' + u)];
      if (uName && String(uName).trim()) {
        var otPay = Number(item[(u === 1) ? 'OT' : ('OT' + u)] || 0);
        workersArr.push(String(uName).trim() + " (" + otPay.toLocaleString() + "원)");
      }
    }

    // 약품 / RTK 요약
    var chemList = [];
    if (Number(item['NDT_형광자분'] || 0) > 0) chemList.push("형자:" + item['NDT_형광자분']);
    if (Number(item['NDT_흑색자분'] || 0) > 0) chemList.push("흑자:" + item['NDT_흑색자분']);
    if (Number(item['NDT_백색페인트'] || 0) > 0) chemList.push("백페:" + item['NDT_백색페인트']);
    if (Number(item['NDT_침투제'] || 0) > 0) chemList.push("침투:" + item['NDT_침투제']);
    if (Number(item['NDT_세척제'] || 0) > 0) chemList.push("세척:" + item['NDT_세척제']);
    if (Number(item['NDT_현상제'] || 0) > 0) chemList.push("현상:" + item['NDT_현상제']);
    if (Number(item['NDT_형광침투제'] || 0) > 0) chemList.push("형침:" + item['NDT_형광침투제']);

    var rtkTotal = Number(item['RTK_총계'] || 0);
    if (rtkTotal > 0) chemList.push("RTK:" + rtkTotal);

    records.push({
      id: startRow + r,
      date: rowDateStr,
      company: String(item['업체명'] || '경인사무소'),
      method: String(item['검사방법'] || ''),
      item_name: String(item['검사품명'] || 'Piping'),
      material: String(item['자재명'] || item['MaterialID'] || item['품목명'] || ''),
      loc_type: String(item['작업구분'] || item['구분'] || '경인사무소'),
      pipe: String(item['관경(Inch)'] || item['규격'] || ''),
      qty: Number(item['수량'] || item['Usage'] || item['검사량'] || 0),
      unit: String(item['Unit'] || item['단위'] || '매'),
      unit_price: Number(item['단가'] || 0),
      travel_cost: Number(item['출장비'] || 0),
      total_fee: Number(item['검사비'] || 0),
      applied_code: String(item['적용코드'] || ''),
      report_no: String(item['성적서번호'] || ''),
      extra: chemList.join(" "),
      workers: workersArr.join(", ")
    });
  }

  records.reverse();
  return { records: records.slice(0, 15) };
}

/**
 * 전체 저장 데이터 조회 (PC 프로그램 동기화용 API)
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
        val = Utilities.formatDate(val, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
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
