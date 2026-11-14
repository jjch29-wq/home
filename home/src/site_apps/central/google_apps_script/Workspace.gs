/** Mobile workspace: managed configuration, retry-safe queue, history and stock. */
function include(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }

function findMobileRow(sheet, headers, req) {
  if (sheet.getLastRow() < 2) return -1;
  var rows = sheet.getRange(2,1,sheet.getLastRow()-1,headers.length).getValues();
  var timeCol=headers.indexOf('EntryTime'), clientCol=headers.indexOf('ClientRequestID');
  var targetTime=String(req.entry_time || '').trim(), matches=[], exactTimes=[];
  rows.forEach(function(row,index) {
    if(req.client_request_id && clientCol>=0 && String(row[clientCol])===req.client_request_id) matches.push(index+2);
    else if(!req.client_request_id && targetTime) {
      var stored=formatCellTime(row[timeCol]);
      if(stored===targetTime)exactTimes.push(index+2);
      if(stored.split('.')[0]===targetTime.split('.')[0])matches.push(index+2);
    }
  });
  if(!req.client_request_id && exactTimes.length)matches=exactTimes;
  if(matches.length===1)return matches[0];
  var id=Number(req.id);
  if(matches.indexOf(id)>=0)return id;
  return -1; // Never guess from a shifted sheet row number.
}

function validateMobileRecord(req) {
  if (!req) throw new Error('저장할 작업이 없습니다.');
  if (req.row_data) return; // PC sync already supplies a complete ledger row.
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(req.date || ''))) throw new Error('검사 날짜를 확인하세요.');
  var parsed=new Date(req.date+'T00:00:00Z');
  if (!isFinite(parsed.getTime()) || parsed.toISOString().slice(0,10)!==req.date) throw new Error('검사 날짜를 확인하세요.');
  if (getConfig().methods.indexOf(req.method) < 0) throw new Error('검사방법을 확인하세요.');
  ['qty', 'unit_price', 'travel_cost', 'total_fee'].forEach(function(key) {
    var value = Number(req[key] || 0);
    if (!isFinite(value) || value < 0) throw new Error('수량·금액은 0 이상의 숫자로 입력하세요.');
  });
  if (req.method === 'RT' && Number(req.qty) % 1) throw new Error('RT 매수는 정수로 입력하세요.');
  if ((req.workers || []).length > 10) throw new Error('작업자는 최대 10명입니다.');
  (req.workers || []).forEach(function(worker) {
    ['ot', 'meal'].forEach(function(key) {
      var v = Number(String(worker[key] || 0).replace(/,/g, ''));
      if (!isFinite(v) || v < 0) throw new Error('작업자 OT·일비를 확인하세요.');
    });
  });
}

function saveRecord(req) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    validateMobileRecord(req);
    var sheet = getDailyUsageSheet();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var idColumn = headers.indexOf('ClientRequestID');
    if (idColumn < 0) {
      idColumn = headers.length;
      sheet.getRange(1, idColumn + 1).setValue('ClientRequestID');
    }
    headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    var hashColumn=headers.indexOf('ClientRequestHash');
    if(hashColumn<0){hashColumn=headers.length;sheet.getRange(1,hashColumn+1).setValue('ClientRequestHash');}
    var map=buildRowMap(req), payload={};
    Object.keys(map).sort().forEach(function(key){if(key!=='EntryTime' && key!=='입력시간' && key!=='ClientRequestHash')payload[key]=map[key];});
    req._request_hash=Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(payload)));
    if (req.client_request_id && sheet.getLastRow() > 1) {
      var ids = sheet.getRange(2, idColumn + 1, sheet.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === req.client_request_id) {
          var oldHash=sheet.getRange(i+2,hashColumn+1).getValues()[0][0];
          if(oldHash && String(oldHash)!==req._request_hash)throw new Error('같은 요청이 이미 다른 내용으로 저장되었습니다. 기록 조회에서 저장된 작업을 확인하세요.');
          return {status:'success', duplicate:true, request_id:req.client_request_id};
        }
      }
    }
    var result = saveRecordCore(req);
    SpreadsheetApp.flush();
    result.request_id = req.client_request_id || '';
    return result;
  } finally { lock.releaseLock(); }
}

function saveRecordsBatch(payload) {
  var records = payload && payload.records;
  if (!Array.isArray(records) || !records.length || records.length > 30) throw new Error('작업 목록은 1~30건으로 저장하세요.');
  records.forEach(function(row) {
    if (!row.client_request_id) throw new Error('작업 식별자가 없습니다. 목록을 다시 확인하세요.');
    validateMobileRecord(row);
  });
  var results = [];
  records.forEach(function(row) {
    try { results.push(saveRecord(row)); }
    catch (error) { results.push({status:'error', request_id:row.client_request_id, detail:String(error)}); }
  });
  return {status:results.every(function(r) {return r.status === 'success';}) ? 'success':'partial', results:results};
}

function updateRecord(req) {
  var lock=LockService.getScriptLock();lock.waitLock(30000);
  try {return updateRecordCore(req);} finally {lock.releaseLock();}
}

function deleteRecord(req) {
  var lock=LockService.getScriptLock();lock.waitLock(30000);
  try {return deleteRecordCore(req);} finally {lock.releaseLock();}
}

function queryRecords(filters) {
  filters = filters || {};
  var sheet = getDailyUsageSheet();
  if (sheet.getLastRow() < 2) return {status:'success', records:[], total:0};
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var data = sheet.getRange(2, 1, sheet.getLastRow()-1, headers.length).getValues();
  var records = [];
  data.forEach(function(row, i) {
    var record = mobileRecordFromRow(headers, row, i + 2);
    if (filters.start && record.date < filters.start) return;
    if (filters.end && record.date > filters.end) return;
    if (filters.company && record.company !== filters.company) return;
    if (filters.method && record.method !== filters.method) return;
    var term = String(filters.search || '').toLowerCase();
    if (term && JSON.stringify(record).toLowerCase().indexOf(term) < 0) return;
    records.push(record);
  });
  records.sort(function(a,b) {return (b.date + b.entry_time).localeCompare(a.date + a.entry_time);});
  var offset = Math.max(0, Number(filters.offset) || 0);
  return {status:'success', records:records.slice(offset, offset+50), total:records.length, offset:offset};
}

function getInventory(filters) {
  filters = filters || {};
  var reference = typeof CENTRAL_REFERENCE !== 'undefined' ? CENTRAL_REFERENCE : {inventory:[], generated_at:''};
  var term = String(filters.search || '').toLowerCase();
  var rows = reference.inventory.filter(function(row) {
    return (!term || JSON.stringify(row).toLowerCase().indexOf(term) >= 0) &&
      (!filters.warehouse || row.warehouse === filters.warehouse);
  });
  var offset = Math.max(0, Number(filters.offset) || 0);
  return {status:'success', records:rows.slice(offset,offset+50), total:rows.length,
          generated_at:reference.generated_at, snapshot:true, offset:offset,
          unassigned_chemicals:reference.unassigned_chemicals || 0};
}

function mobileRecordFromRow(headerRow, row, rowNumber) {
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

  var workersArr = [];
  var workerList = [];
  for (var u = 1; u <= 10; u++) {
    var uName = item[(u === 1) ? 'User' : ('User' + u)];
    if (uName && String(uName).trim()) {
      var rawTime = String(item[(u === 1) ? 'WorkTime' : ('WorkTime' + u)] || '09:00~18:00');
      var shiftMatch = rawTime.match(/^\((.*?)\)\s*(.*)$/);
      var shiftVal = shiftMatch ? shiftMatch[1] : '주간';
      var cleanTime = shiftMatch ? shiftMatch[2] : rawTime;
      var otPay = safeNum(item[(u === 1) ? 'OT' : ('OT' + u)], 0);

      workersArr.push(String(uName).trim() + " (" + otPay.toLocaleString() + "원)");
      workerList.push({
        user: String(uName).trim(),
        shift: shiftVal,
        time: cleanTime,
        ot: String(otPay), meal: String(safeNum(item['Meal' + (u === 1 ? '' : u)], 0))
      });
    }
  }

  var chemList = [];
  if (safeNum(item['NDT_형광자분']) > 0) chemList.push("형자:" + item['NDT_형광자분']);
  if (safeNum(item['NDT_흑색자분']) > 0) chemList.push("흑자:" + item['NDT_흑색자분']);
  if (safeNum(item['NDT_백색페인트']) > 0) chemList.push("백페:" + item['NDT_백색페인트']);
  if (safeNum(item['NDT_침투제']) > 0) chemList.push("침투:" + item['NDT_침투제']);
  if (safeNum(item['NDT_세척제']) > 0) chemList.push("세척:" + item['NDT_세척제']);
  if (safeNum(item['NDT_현상제']) > 0) chemList.push("현상:" + item['NDT_현상제']);
  if (safeNum(item['NDT_형광침투제']) > 0) chemList.push("형침:" + item['NDT_형광침투제']);

  var rtkTotal = safeNum(item['RTK_총계'] || item['RTK_RTK총계'], 0);
  if (rtkTotal > 0) chemList.push("RTK:" + rtkTotal);

  var entryTimeVal = formatCellTime(item['EntryTime'] || item['입력시간']);

  return {
    id: rowNumber,
    entry_time: entryTimeVal,
    client_request_id: String(item['ClientRequestID'] || ''),
    material_id: String(item['MaterialID'] || ''),
    date: rowDateStr,
    site: String(item['Site'] || '중앙지사'),
    company: String(item['업체명'] || '한국지역난방공사 중앙지사'),
    method: String(item['검사방법'] || ''),
    item_name: String(item['검사품명'] || 'Pipe'),
    material: String(item['Material'] || item['자재명'] || item['품목명'] || ''),
    loc_type: String(item['작업구분'] || item['구분'] || '열수송관'),
    pipe: String(item['관경(Inch)'] || item['규격'] || ''),
    hot_tap_stage: String(item['핫태핑차수'] || '1차'),
    hot_tap_joints: safeNum(item['핫태핑개소'], 1),
    hot_tap_pipe: String(item['핫태핑관경'] || ''),
    joints: safeNum(item['조인트수'], 0),
    ori_joints: safeNum(item['조인트수'], 0),
    rep_joints: safeNum(item['rep_joints'], 0),
    rej_qty: safeNum(item['불량수'], 0),
    qty: safeNum(item['수량'] || item['Usage'] || item['검사량'], 0),
    unit: String(item['Unit'] || item['단위'] || 'M'),
    unit_price: safeNum(item['단가'], 0),
    travel_cost: safeNum(item['출장비'], 0),
    lump_sum: String(item['일식'] || ''),
    total_fee: safeNum(item['검사비'], 0),
    equip_main: String(item['장비명'] || ''),
    equip_aux: String(item['보조장비'] || ''),
    vehicle_no: String(item['차량번호'] || ''),
    mileage: String(item['주행거리'] || ''),
    vehicle_note: String(item['차량비고'] || ''),
    note: String(item['Note'] || item['Remark'] || item['DailyMemo'] || ''),
    applied_code: String(item['적용코드'] || ''),
    report_no: String(item['성적서번호'] || ''),
    rtk_center_miss: safeNum(item['RTK_센터미스'], 0),
    rtk_density: safeNum(item['RTK_농도'], 0),
    rtk_marking_miss: safeNum(item['RTK_마킹미스'], 0),
    rtk_film_mark: safeNum(item['RTK_필름마크'], 0),
    rtk_handling: safeNum(item['RTK_취급부주의'], 0),
    rtk_complaint: safeNum(item['RTK_고객불만'], 0),
    rtk_etc: safeNum(item['RTK_기타'], 0),
    rtk_total: rtkTotal,
    chem_fl_mag: safeNum(item['NDT_형광자분'], 0),
    chem_black_mag: safeNum(item['NDT_흑색자분'], 0),
    chem_white_paint: safeNum(item['NDT_백색페인트'], 0),
    chem_penetrant: safeNum(item['NDT_침투제'], 0),
    chem_cleaner: safeNum(item['NDT_세척제'], 0),
    chem_developer: safeNum(item['NDT_현상제'], 0),
    chem_fl_pen: safeNum(item['NDT_형광침투제'], 0),
    extra: chemList.join(" "),
    workers: workersArr.join(", "),
    worker_list: workerList
  };
}

/**
 * 작업 물량 및 누계/공정률 현황 집계 (PC 작업/감독일보 '1. 작업 물량 및 누계 현황' 연동)
 */
function classifyMethodSpec(method, pipeStr, shiftStr) {
  var sizeMatch = String(pipeStr || '').match(/\d+/);
  var sizeVal = sizeMatch ? parseInt(sizeMatch[0], 10) : 0;
  var isNight = String(shiftStr || '').indexOf('야간') >= 0;
  var m = String(method || '').toUpperCase().trim();

  if (m === 'RT') {
    var spec = (sizeVal <= 80 ? '80A이하' : '150A~100A') + (isNight ? '-야간' : '');
    return 'RT_' + spec;
  } else if (m === 'PAUT') {
    var spec = '200A';
    if (sizeVal >= 300) spec = '300A이상';
    else if (sizeVal === 250) spec = '250A';
    else if (sizeVal === 200) spec = '200A';
    else if (sizeVal >= 125 && sizeVal <= 150) spec = '150A~125A';
    else if (sizeVal <= 100) spec = '100A이하';
    if (isNight) spec += '-야간';
    return 'PAUT_' + spec;
  } else if (m === 'MT') {
    return 'MT_' + (isNight ? '전체(야간)' : '전체(주간)');
  } else if (m === 'PT') {
    return 'PT_' + (isNight ? '전체(야간)' : '전체(주간)');
  }
  return '';
}

function getProgressSummary(targetDate) {
  targetDate = String(targetDate || '').trim();
  if (!targetDate || !/^\d{4}-\d{2}-\d{2}$/.test(targetDate)) {
    targetDate = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd");
  }

  var expectedQtyMap = (typeof CENTRAL_REFERENCE !== 'undefined' && CENTRAL_REFERENCE.default_expected_qty) || {
    'PAUT_300A이상': 129, 'PAUT_300A이상-야간': 624,
    'PAUT_250A': 4, 'PAUT_250A-야간': 1, 'PAUT_200A': 4,
    'PAUT_200A-야간': 2, 'PAUT_150A~125A': 1, 'PAUT_150A~125A-야간': 1,
    'PAUT_100A이하': 1, 'PAUT_100A이하-야간': 1, 'PAUT_소계': 768,
    'RT_150A~100A': 293, 'RT_150A~100A-야간': 43,
    'RT_80A이하': 105, 'RT_80A이하-야간': 49, 'RT_소계': 490,
    'MT_전체(주간)': 26, 'MT_전체(야간)': 1,
    'PT_전체(주간)': 26, 'PT_전체(야간)': 1
  };

  var QTY_ROWS = [
    ['PAUT', '300A이상'], ['PAUT', '300A이상-야간'], ['PAUT', '250A'], ['PAUT', '250A-야간'],
    ['PAUT', '200A'], ['PAUT', '200A-야간'], ['PAUT', '150A~125A'], ['PAUT', '150A~125A-야간'],
    ['PAUT', '100A이하'], ['PAUT', '100A이하-야간'], ['PAUT', '소계'],
    ['RT', '150A~100A'], ['RT', '150A~100A-야간'], ['RT', '80A이하'], ['RT', '80A이하-야간'],
    ['RT', '소계'],
    ['MT', '전체(주간)'], ['MT', '전체(야간)'],
    ['PT', '전체(주간)'], ['PT', '전체(야간)']
  ];

  var history = (typeof CENTRAL_REFERENCE !== 'undefined' && CENTRAL_REFERENCE.daily_progress_history) || {};
  var histDates = Object.keys(history).sort();
  var pastDates = histDates.filter(function(d) { return d < targetDate; });
  var latestPastDate = pastDates.length > 0 ? pastDates[pastDates.length - 1] : null;

  var prevCumulative = {};
  var prevDefects = {};
  var todayQty = {};
  var todayDefects = {};

  // 1. PC 일보 히스토리에서 기준 전일누계 로드
  if (latestPastDate && history[latestPastDate] && history[latestPastDate].qty_data) {
    var pData = history[latestPastDate].qty_data;
    Object.keys(pData).forEach(function(k) {
      prevCumulative[k] = safeNum(pData[k]['총누계'] || pData[k]['전일누계'], 0);
      prevDefects[k] = safeNum(pData[k]['불량'], 0);
    });
  }

  // 2. Google Spreadsheet 실시간 DailyUsage 시트 반영
  var sheet = getDailyUsageSheet();
  var lastRow = sheet.getLastRow();
  var weatherVal = '';
  var remarksList = [];

  if (lastRow > 1) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var record = mobileRecordFromRow(headers, row, i + 2);
      var recDate = record.date;
      if (!recDate) continue;

      // 기준일 이전 데이터는 이미 prevCumulative에 포함되어 있으므로 건너뜀
      if (latestPastDate && recDate <= latestPastDate) {
        continue;
      }

      var shiftStr = '주간';
      if (record.worker_list && record.worker_list.length > 0) {
        for (var w = 0; w < record.worker_list.length; w++) {
          if (record.worker_list[w].shift === '야간' || String(record.worker_list[w].time).indexOf('야간') >= 0) {
            shiftStr = '야간';
            break;
          }
        }
      }

      var compKey = classifyMethodSpec(record.method, record.pipe, shiftStr);
      if (!compKey) continue;

      var val = safeNum(record.qty, 0);
      var rej = safeNum(record.rej_qty, 0);

      if (recDate < targetDate) {
        prevCumulative[compKey] = (prevCumulative[compKey] || 0) + val;
        prevDefects[compKey] = (prevDefects[compKey] || 0) + rej;
      } else if (recDate === targetDate) {
        todayQty[compKey] = (todayQty[compKey] || 0) + val;
        todayDefects[compKey] = (todayDefects[compKey] || 0) + rej;
        if (record.note && remarksList.indexOf(record.note) < 0) {
          remarksList.push(record.note);
        }
      }
    }
  }

  // 만약 targetDate가 과거 히스토리에 완전히 존재하는 경우 히스토리의 날씨/비고도 참조
  if (history[targetDate]) {
    weatherVal = history[targetDate].weather || weatherVal;
    if (history[targetDate].remarks && remarksList.indexOf(history[targetDate].remarks) < 0) {
      remarksList.unshift(history[targetDate].remarks);
    }
  }

  // 3. 소계 계산
  var pautKeys = [
    '300A이상', '300A이상-야간', '250A', '250A-야간', '200A', '200A-야간',
    '150A~125A', '150A~125A-야간', '100A이하', '100A이하-야간'
  ];
  var pautPrev = 0, pautToday = 0;
  pautKeys.forEach(function(s) {
    pautPrev += (prevCumulative['PAUT_' + s] || 0);
    pautToday += (todayQty['PAUT_' + s] || 0);
  });
  prevCumulative['PAUT_소계'] = pautPrev;
  todayQty['PAUT_소계'] = pautToday;

  var rtKeys = ['150A~100A', '150A~100A-야간', '80A이하', '80A이하-야간'];
  var rtPrev = 0, rtToday = 0, rtDefects = 0;
  rtKeys.forEach(function(s) {
    var k = 'RT_' + s;
    rtPrev += (prevCumulative[k] || 0);
    rtToday += (todayQty[k] || 0);
    rtDefects += ((prevDefects[k] || 0) + (todayDefects[k] || 0));
  });
  prevCumulative['RT_소계'] = rtPrev;
  todayQty['RT_소계'] = rtToday;

  function fmtVal(key, v) {
    if (v === 0 || !v) return '';
    if (key.indexOf('PAUT') === 0 || key.indexOf('MT') === 0 || key.indexOf('PT') === 0) {
      return Number(v).toFixed(4);
    }
    return (v % 1) ? Number(v).toFixed(1) : String(Math.round(v));
  }

  // 4. 결과 테이블 행 생성
  var resultRows = [];
  QTY_ROWS.forEach(function(item) {
    var method = item[0];
    var spec = item[1];
    var compKey = method + '_' + spec;
    var expected = safeNum(expectedQtyMap[compKey], 0);
    var prev = safeNum(prevCumulative[compKey], 0);
    var today = safeNum(todayQty[compKey], 0);
    var total = prev + today;
    var rate = expected > 0 ? (total / expected * 100) : 0;

    var defects = (prevDefects[compKey] || 0) + (todayDefects[compKey] || 0);
    if (compKey === 'RT_소계') defects = rtDefects;
    var rejectRate = (total > 0 && defects > 0) ? (defects / total * 100) : 0;

    resultRows.push({
      method: method,
      spec: spec,
      comp_key: compKey,
      expected: expected ? String(expected) : '',
      prev: fmtVal(compKey, prev),
      today: fmtVal(compKey, today),
      total: fmtVal(compKey, total),
      rate: total > 0 ? rate.toFixed(1) : '',
      defects: defects > 0 ? String(defects) : '',
      reject_rate: rejectRate > 0 ? rejectRate.toFixed(1) : '',
      is_subtotal: spec === '소계',
      num_rate: rate,
      num_total: total
    });
  });

  var pautSubtotal = resultRows.find(function(r) { return r.comp_key === 'PAUT_소계'; }) || {};
  var rtSubtotal = resultRows.find(function(r) { return r.comp_key === 'RT_소계'; }) || {};
  var mtDay = resultRows.find(function(r) { return r.comp_key === 'MT_전체(주간)'; }) || {};
  var ptDay = resultRows.find(function(r) { return r.comp_key === 'PT_전체(주간)'; }) || {};

  // 5. 기성 정산 요약 및 세부 레코드 (PC 기성 정산(NDT) '1. 일별 작업 기록 및 목록' 연동)
  var baseBilling = (typeof CENTRAL_REFERENCE !== 'undefined' && CENTRAL_REFERENCE.billing_records) || [];
  var billingList = [];
  var billingSummary = {
    total_amount: 0,
    paut_amount: 0,
    rt_amount: 0,
    mt_amount: 0,
    pt_amount: 0,
    paut_qty: 0,
    rt_qty: 0,
    mt_qty: 0,
    pt_qty: 0,
    records_count: 0
  };

  baseBilling.forEach(function(rec) {
    if (!rec.date || rec.date <= targetDate) {
      billingList.push(rec);
    }
  });

  billingList.forEach(function(r) {
    var amt = safeNum(r.subtotal, 0);
    var q = safeNum(r.qty, 0);
    var t = String(r.ndt_type || '').toUpperCase();
    billingSummary.total_amount += amt;
    if (t.indexOf('PAUT') >= 0 || t.indexOf('UT') >= 0) {
      billingSummary.paut_amount += amt;
      billingSummary.paut_qty += q;
    } else if (t.indexOf('RT') >= 0) {
      billingSummary.rt_amount += amt;
      billingSummary.rt_qty += q;
    } else if (t.indexOf('MT') >= 0) {
      billingSummary.mt_amount += amt;
      billingSummary.mt_qty += q;
    } else if (t.indexOf('PT') >= 0) {
      billingSummary.pt_amount += amt;
      billingSummary.pt_qty += q;
    }
  });
  billingSummary.records_count = billingList.length;

  return {
    status: 'success',
    date: targetDate,
    weather: weatherVal,
    remarks: remarksList.join('; '),
    rows: resultRows,
    billing: billingSummary,
    billing_records: billingList,
    kpi: {
      paut_rate: pautSubtotal.rate || '0.0',
      paut_total: pautSubtotal.total || '0',
      paut_expected: pautSubtotal.expected || '768',
      rt_rate: rtSubtotal.rate || '0.0',
      rt_total: rtSubtotal.total || '0',
      rt_expected: rtSubtotal.expected || '490',
      mt_rate: mtDay.rate || '0.0',
      mt_total: mtDay.total || '0',
      mt_expected: mtDay.expected || '26',
      pt_rate: ptDay.rate || '0.0',
      pt_total: ptDay.total || '0',
      pt_expected: ptDay.expected || '26'
    }
  };
}
