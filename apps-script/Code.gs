/*******************************************************************
 *  ㈜아이에스에이연구원 · 견적요청 백엔드 (Google Apps Script)
 *  ---------------------------------------------------------------
 *  기능 : 견적 신청 접수 → 구글 시트 저장 + 이메일 발송,
 *         접수번호 + 비밀번호로 본인 견적 조회 / 수정 (제품 여러 개 지원)
 *
 *  설치 방법은 "견적요청_설치안내.md" 파일을 참고하세요.
 *******************************************************************/

/* ===================== 설정 (필요 시 수정) ===================== */
var RECIPIENT_EMAIL = 'isalab20181101@gmail.com';   // 알림 받을 회사 이메일 (여러 개면 콤마)
var SEND_CONFIRM_TO_APPLICANT = true;               // 신청자에게 접수 확인메일 발송
var SALT = 'isa-lab-quote-secret-2026';             // 비밀번호 암호화용 소금값
var SHEET_NAME = '견적요청';                          // 데이터가 저장될 시트 이름
var ORG_NAME = '㈜아이에스에이연구원';
/* ============================================================== */

// 시트 컬럼 순서 (비밀번호는 항상 마지막)
var HEADERS = [
  '접수번호','신청일시','처리상태','허가종류','업체명','담당자','연락처','이메일',
  '검사제품목록','긴급처리','기타요청사항','자가품질도우미','제품목록JSON','비밀번호(암호화)'
];

function doGet(e) {
  return jsonOut({ ok: true, msg: ORG_NAME + ' 견적요청 API 정상 작동중' });
}

function doPost(e) {
  var out;
  try {
    var data = JSON.parse(e.postData.contents);
    switch (data.action) {
      case 'create': out = handleCreate(data); break;
      case 'lookup': out = handleLookup(data); break;
      case 'update': out = handleUpdate(data); break;
      default: out = { ok: false, error: '알 수 없는 요청입니다.' };
    }
  } catch (err) {
    out = { ok: false, error: '서버 오류: ' + err };
  }
  return jsonOut(out);
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- 시트 준비 ---------------- */
function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#1d3158').setFontColor('#ffffff');
  }
  return sh;
}

/* ---------------- 비밀번호 해시 ---------------- */
function hashPw(pw) {
  var raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, SALT + '|' + pw, Utilities.Charset.UTF_8);
  return raw.map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

/* ---------------- 접수번호 생성 ---------------- */
function makeCode(sh) {
  var today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd');
  var prefix = 'Q' + today + '-';
  var seq = 1;
  var last = sh.getLastRow();
  if (last >= 2) {
    var codes = sh.getRange(2, 1, last - 1, 1).getValues();
    codes.forEach(function (r) {
      var c = String(r[0]);
      if (c.indexOf(prefix) === 0) {
        var n = parseInt(c.substring(prefix.length), 10);
        if (n >= seq) seq = n + 1;
      }
    });
  }
  return prefix + ('00' + seq).slice(-3);
}

function findRow(sh, code) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var codes = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < codes.length; i++) {
    if (String(codes[i][0]) === String(code)) return i + 2;
  }
  return -1;
}

function colIndex(name) { return HEADERS.indexOf(name) + 1; }

/* ---------------- 자가품질 도우미 요약 ---------------- */
function helperSummary(d) {
  var parts = [];
  if (d.foodType) parts.push('식품유형:' + d.foodType);
  if (d.steril) parts.push('살균:' + d.steril);
  if (d.eat) parts.push('섭취:' + d.eat);
  if (d.powder) parts.push('분말가루환:' + d.powder);
  if (d.special) parts.push('특수대상:' + d.special);
  if (d.longStore) parts.push('장기보존:' + d.longStore);
  if (d.extraItems) parts.push('예상추가항목:' + d.extraItems);
  return parts.join(' | ');
}

/* ---------------- 신청 접수 ---------------- */
function handleCreate(d) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = getSheet();
    var code = makeCode(sh);
    var now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
    var productsJSON = JSON.stringify(d.products || []);
    var row = [
      code, now, '접수',
      d.license || '', d.company || '', d.name || '', d.phone || '', d.email || '',
      d.productsText || '', d.urgent || '일반', d.message || '', helperSummary(d),
      productsJSON, hashPw(d.password || '')
    ];
    sh.appendRow(row);
    notifyNew(code, now, d);
    return { ok: true, code: code };
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- 조회 ---------------- */
function handleLookup(d) {
  var sh = getSheet();
  var r = findRow(sh, d.code);
  if (r < 0) return { ok: false, error: '해당 접수번호를 찾을 수 없습니다.' };
  var vals = sh.getRange(r, 1, 1, HEADERS.length).getValues()[0];
  var stored = vals[colIndex('비밀번호(암호화)') - 1];
  if (hashPw(d.password || '') !== stored) return { ok: false, error: '비밀번호가 올바르지 않습니다.' };
  return { ok: true, record: rowToRecord(vals) };
}

/* ---------------- 수정 ---------------- */
function handleUpdate(d) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = getSheet();
    var r = findRow(sh, d.code);
    if (r < 0) return { ok: false, error: '해당 접수번호를 찾을 수 없습니다.' };
    var stored = sh.getRange(r, colIndex('비밀번호(암호화)')).getValue();
    if (hashPw(d.password || '') !== stored) return { ok: false, error: '비밀번호가 올바르지 않습니다.' };

    var map = {
      '업체명': d.company, '담당자': d.name, '연락처': d.phone, '이메일': d.email,
      '기타요청사항': d.message
    };
    if (d.productsText !== undefined) map['검사제품목록'] = d.productsText;
    if (d.products !== undefined) map['제품목록JSON'] = JSON.stringify(d.products || []);
    Object.keys(map).forEach(function (key) {
      if (map[key] !== undefined) sh.getRange(r, colIndex(key)).setValue(map[key]);
    });
    var vals = sh.getRange(r, 1, 1, HEADERS.length).getValues()[0];
    notifyUpdate(d.code, vals);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- 행 → 레코드(비번 제외) ---------------- */
function rowToRecord(vals) {
  var products = [];
  try { products = JSON.parse(vals[colIndex('제품목록JSON') - 1] || '[]'); } catch (e) { products = []; }
  return {
    code: vals[colIndex('접수번호') - 1],
    date: vals[colIndex('신청일시') - 1],
    status: vals[colIndex('처리상태') - 1] || '접수',
    company: vals[colIndex('업체명') - 1],
    name: vals[colIndex('담당자') - 1],
    phone: vals[colIndex('연락처') - 1],
    email: vals[colIndex('이메일') - 1],
    message: vals[colIndex('기타요청사항') - 1],
    productsText: vals[colIndex('검사제품목록') - 1],
    products: products
  };
}

/* ---------------- 이메일 : 신규 접수 ---------------- */
function notifyNew(code, now, d) {
  var lines = [
    '새 견적 신청이 접수되었습니다.', '',
    '■ 접수번호 : ' + code,
    '■ 신청일시 : ' + now, '',
    '[신청자]',
    '· 허가종류 : ' + (d.license || '-'),
    '· 업체/기관 : ' + (d.company || '-'),
    '· 담당자 : ' + (d.name || '-'),
    '· 연락처 : ' + (d.phone || '-'),
    '· 이메일 : ' + (d.email || '-'),
    '· 긴급처리 : ' + (d.urgent || '일반'), '',
    '[검사 제품 목록]',
    (d.productsText || '-'), '',
    '[자가품질 도우미] ' + (helperSummary(d) || '-'), '',
    '[기타 요청사항]',
    (d.message || '-'), '',
    '────────────────────────',
    '※ 전체 목록은 구글 시트에서 확인하실 수 있습니다.'
  ];
  MailApp.sendEmail({
    to: RECIPIENT_EMAIL,
    subject: '[견적신청] ' + (d.company || '무기명') + ' / ' + code,
    body: lines.join('\n'),
    replyTo: d.email || RECIPIENT_EMAIL
  });

  if (SEND_CONFIRM_TO_APPLICANT && d.email && /@/.test(d.email)) {
    var a = [
      (d.name || '고객') + '님, 안녕하세요. ' + ORG_NAME + '입니다.', '',
      '견적 신청이 정상 접수되었습니다.', '',
      '■ 접수번호 : ' + code,
      '■ 신청일시 : ' + now, '',
      '[신청하신 제품]',
      (d.productsText || '-'), '',
      '접수번호와 신청 시 설정하신 비밀번호로 홈페이지 [견적·상담 > 견적 조회·수정]에서',
      '신청 내용을 확인·수정하실 수 있습니다.', '',
      '담당자가 1영업일 이내에 견적을 안내해 드리겠습니다.',
      '문의 : 031-429-0620 / 070-8831-7376', '',
      ORG_NAME
    ];
    MailApp.sendEmail({ to: d.email, subject: '[' + ORG_NAME + '] 견적 신청 접수 완료 (' + code + ')', body: a.join('\n') });
  }
}

/* ---------------- 이메일 : 수정 알림 ---------------- */
function notifyUpdate(code, vals) {
  var body = [
    '견적 신청 내용이 신청자에 의해 수정되었습니다.', '',
    '■ 접수번호 : ' + code,
    '· 업체/기관 : ' + vals[colIndex('업체명') - 1],
    '· 담당자 : ' + vals[colIndex('담당자') - 1],
    '· 연락처 : ' + vals[colIndex('연락처') - 1],
    '· 이메일 : ' + vals[colIndex('이메일') - 1], '',
    '[검사 제품 목록]',
    (vals[colIndex('검사제품목록') - 1] || '-'), '',
    '[기타 요청사항]',
    (vals[colIndex('기타요청사항') - 1] || '-'), '',
    '※ 자세한 내용은 구글 시트에서 확인하세요.'
  ];
  MailApp.sendEmail({ to: RECIPIENT_EMAIL, subject: '[견적수정] ' + code, body: body.join('\n') });
}

/* ---------------- (선택) 테스트용 ---------------- */
function testCreate() {
  var res = handleCreate({
    action: 'create', license: '식품제조·가공업', company: '테스트업체', name: '홍길동',
    phone: '031-000-0000', email: RECIPIENT_EMAIL,
    products: [
      { cat: '자가품질검사', name: '과자류 / ○○쿠키', items: '산가, 세균수', qty: '1' },
      { cat: '영양성분', name: '○○시리얼', items: '9대 영양성분', qty: '1' }
    ],
    productsText: '1) [자가품질검사] 과자류 / ○○쿠키 / 항목: 산가, 세균수 / 수량: 1\n2) [영양성분] ○○시리얼 / 항목: 9대 영양성분 / 수량: 1',
    foodType: '과자', steril: '비살균', eat: '그대로섭취', powder: '예', special: '영유아', longStore: '레토르트',
    extraItems: '금속성이물(쇳가루) 별표2', urgent: '일반', message: '테스트입니다.', password: '1234'
  });
  Logger.log(res);
}
