/** @OnlyCurrentDoc */

function setup1() {
  setFastroCookieFor('1', 'PHPSESSID=noeospqnpdhufiuprpdli80pr6; hd_pops_78=1; hd_pops_77=1');
  Logger.log('완료');
}

function myFunction() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('H65').activate();
  spreadsheet.getActiveRangeList().setBackground('#ffff00');
};
function myFunction1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().setBackground('#c9daf8');
};
function myFunction2() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().setBackground('#ff9900');
};
function _4() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().setBackground('#00ff00');
};
function myFunction3() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().setBackground('#ffff00');
};
function myFunction4() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().setBackground('#ffffff');
};
function myFunction5() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().setBackground('#00ff00');
};
function myFunction6() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().setBackground('#b6d7a8');
};
function myFunction7() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().setBackground('#93c47d');
};
function LCL() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().setBackground('#ff00ff');
};
function myFunction8() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E46').activate();
  spreadsheet.getActiveRangeList().setBackground('#e06666');
};
function myFunction9() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().setBackground('#e06666');
};
function myFunction10() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().setBackground('#ea9999');
};
function myFunction11() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().setBackground('#ffffff');
};
function myFunction12() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().setBackground('#ea9999');
};


// 사업자별 쇼핑몰 이름 결정
//  - 1~4: 항상 '모다mall'
//  - 5:   항상 '수기mall'
//  - 6:   판매처에 '수기' 포함 시 '수기mall', 그 외 '모다mall'
function getMallName_(ssId, source) {
  var bid = String(ssId || '1');
  if (bid === '5') return '수기mall';
  if (bid === '6') {
    return String(source || '').indexOf('수기') >= 0 ? '수기mall' : '모다mall';
  }
  return '모다mall';
}

function buildMsg_(row, ssId) {
  var NL=String.fromCharCode(10), TAB=String.fromCharCode(9);
  var n=String(row[7]).trim(), o=String(row[6]).trim(), p=String(row[8]).trim();
  var opt=String(row[9]).trim(), q=String(row[10]).trim()||'1', sh=String(row[5]).trim();
  var di=sh.indexOf('-'), src=di>=0?sh.substring(di+1).trim():sh.trim();
  var mall = getMallName_(ssId, sh);
  var msg=n+'님'+NL+'안녕하세요. 쇼핑몰 '+mall+'입니다.'+NL+'주문해주신 제품 정상 접수되었습니다.'+NL+NL;
  if(src) msg+='[구입처] : '+src+NL;
  if(o)   msg+='[주문번호] : '+o+NL;
  if(p)   msg+='[제품] : '+p+NL;
  msg+='[옵션] : '+(opt||'')+TAB+'수량:'+q+'개';
  return msg;
}

// ─────────────────────────────────────────────────────────
// ★ 로그인 쿠키 — PropertiesService 사용
// ─────────────────────────────────────────────────────────
function getLoginCookie_() {
  var props = PropertiesService.getScriptProperties();
  var cookie = props.getProperty('LOGIN_COOKIE');
  if (cookie) return cookie;
  var fallback = 'connect.sid=s%3Agu4c_Li3M-EbBV0g-b6bnWdTKDUkWM91.gn4BvUAq50TAFydPCKKaOhb1qIW8fWiaO1DJQ5xeCTw';
  props.setProperty('LOGIN_COOKIE', fallback);
  return fallback;
}

function setLoginCookie(newCookie) {
  if (!newCookie) {
    Logger.log('사용법: setLoginCookie("connect.sid=...")');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('LOGIN_COOKIE', newCookie);
  Logger.log('쿠키 갱신 완료');
}

// ─────────────────────────────────────────────────────────
// ★ 공통 헬퍼: 현재 월 시트 또는 가장 최근 월 시트 반환
// ─────────────────────────────────────────────────────────
function getCurrentMonthSheet_(ss) {
  var now = new Date();
  var sheetName = String(now.getFullYear()).slice(2) + '년 ' + (now.getMonth()+1) + '월';
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) return { sheet: sheet, sheetName: sheetName };

  var monthSheets = ss.getSheets().filter(function(s) {
    return /^\d{2}년 \d+월$/.test(s.getName());
  });
  if (monthSheets.length === 0) return { sheet: null, sheetName: sheetName };

  monthSheets.sort(function(a, b) {
    var pa = a.getName().match(/(\d+)년 (\d+)월/);
    var pb = b.getName().match(/(\d+)년 (\d+)월/);
    if (!pa || !pb) return 0;
    return (Number(pb[1])*100 + Number(pb[2])) - (Number(pa[1])*100 + Number(pa[2]));
  });
  return { sheet: monthSheets[0], sheetName: monthSheets[0].getName() };
}

// 공통 헬퍼: 주문번호로 행 찾기
function findRowByOrder_(sheet, orderNo) {
  var data = sheet.getDataRange().getValues();
  // ★ G열만 display values로 따로 가져오기 (큰 숫자 정밀도 손실 방지)
  var lastRow = sheet.getLastRow();
  var displayG = lastRow >= 1 ? sheet.getRange(1, 7, lastRow, 1).getDisplayValues() : [];
  var searchOrder = String(orderNo).trim();
  for (var i = 0; i < data.length; i++) {
    // 1차: 일반 값 매칭
    if (String(data[i][6]).trim() === searchOrder) {
      return { row: i + 1, data: data };
    }
    // 2차: display value 매칭 (숫자형 셀일 경우)
    if (displayG[i] && String(displayG[i][0]).trim() === searchOrder) {
      return { row: i + 1, data: data };
    }
  }
  return { row: -1, data: data };
}

// ═══════════════════════════════════════════════════════════════
// ★ 패스트로(fast-ro.com) 연동 v2 — 사업자별 쿠키 관리
// ═══════════════════════════════════════════════════════════════

function getFastroCookie_(ssId) {
  var props = PropertiesService.getScriptProperties();
  var perSs = props.getProperty('FASTRO_COOKIE_' + ssId);
  if (perSs) return perSs;
  return props.getProperty('FASTRO_COOKIE') || '';
}

function setFastroCookieFor(ssId, newCookie) {
  if (!ssId || !newCookie) {
    Logger.log('사용법: setFastroCookieFor("1", "PHPSESSID=...")');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('FASTRO_COOKIE_' + ssId, newCookie);
  Logger.log('ssId=' + ssId + ' 패스트로 쿠키 등록 완료');
}

function setFastroCookie(newCookie) {
  if (!newCookie) {
    Logger.log('사용법: setFastroCookie("PHPSESSID=xxx; ...")');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('FASTRO_COOKIE', newCookie);
  Logger.log('패스트로 공통 쿠키 갱신 완료');
}

function listFastroCookies() {
  var props = PropertiesService.getScriptProperties();
  var common = props.getProperty('FASTRO_COOKIE');
  Logger.log('공통 쿠키: ' + (common ? '등록됨 (' + common.length + '자)' : '없음'));
  for (var i = 1; i <= 6; i++) {
    var c = props.getProperty('FASTRO_COOKIE_' + i);
    var mb = props.getProperty('FASTRO_MBID_' + i);
    Logger.log('ssId=' + i + ' → 쿠키:' + (c ? '있음' : '없음') + ', mb_id:' + (mb || '미등록'));
  }
}

function getFastroMbId_(ssId) {
  var props = PropertiesService.getScriptProperties();
  var key = 'FASTRO_MBID_' + ssId;
  var stored = props.getProperty(key);
  if (stored) return stored;
  var defaults = {
    '1': 'kjm6421',
    '2': 'kjm5696421',
    '3': 'kjm5696423',
    '4': 'kjm5696424',
    '5': 'kjm5697425',
    '6': 'kjm5696426'
  };
  return defaults[ssId] || '';
}

function setFastroMbId(ssId, mbId) {
  if (!mbId) {
    PropertiesService.getScriptProperties().deleteProperty('FASTRO_MBID_' + ssId);
    Logger.log('ssId=' + ssId + ' mb_id 삭제 완료');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('FASTRO_MBID_' + ssId, mbId);
  Logger.log('ssId=' + ssId + ' → mb_id=' + mbId + ' 등록 완료');
}

// ─── 검색 인덱스 캐시 무효화 (모든 월수 키 일괄 제거) ───
// 검색 캐시 키: search_index_v4_ss{ssId}, ..._m1, ..._m3, ..._m6
// 어디든 데이터 변경 후 호출하면 다음 검색 때 신선한 데이터로 새로 인덱싱
function clearSearchCache_(ssId) {
  if (!ssId) return;
  try {
    var c = CacheService.getScriptCache();
   var keys = [
  'search_index_v3_ss' + ssId,
  'search_index_v3_ss' + ssId + '_m1',
  'search_index_v3_ss' + ssId + '_m3',
  'search_index_v3_ss' + ssId + '_m6',
  'search_index_v1_ss' + ssId,
  'search_index_v4_ss' + ssId,
  'search_index_v4_ss' + ssId + '_m1',
  'search_index_v4_ss' + ssId + '_m3',
  'search_index_v4_ss' + ssId + '_m6'
];
    c.removeAll(keys);
  } catch (e) {}
}

// ═══════════════════════════════════════════════════════
// ★ 패스트로 상세페이지 구조 파악 디버그
// 사용법: 함수 드롭다운에서 debugFastroDetail 선택 → ▶ 실행
//        → 실행 로그에 패스트로 상세페이지 HTML 분석 결과 표시
// ═══════════════════════════════════════════════════════
function debugFastroDetail() {
  var ssId = '1';
  var cookie = getFastroCookie_(ssId);
  if (!cookie) {
    Logger.log('❌ 쿠키 없음');
    return;
  }
  
  // 1) 목록에서 최근 SH번호 하나 가져오기
  var mbId = getFastroMbId_(ssId);
  var today = new Date();
  var past = new Date(today.getTime() - 30 * 24 * 60 * 60 * 1000);
  var fmt = function(d) {
    return d.getFullYear() + '-' + 
           String(d.getMonth()+1).padStart(2,'0') + '-' + 
           String(d.getDate()).padStart(2,'0');
  };
  var listResult = fastroFetch_(ssId, mbId, fmt(past), fmt(today), 20);
  if (!listResult.ok) {
    Logger.log('❌ 목록 조회 실패: ' + listResult.error);
    return;
  }
  
  var records = parseFastroHtml_(listResult.html);
  if (records.length === 0) {
    Logger.log('❌ 최근 30일 SH번호 없음');
    return;
  }
  
  // 첫 3개 SH번호로 테스트
  var testCount = Math.min(3, records.length);
  
  for (var ri = 0; ri < testCount; ri++) {
    var rec = records[ri];
    var shNo = rec.sh;
    if (!shNo) continue;
    
    var url = 'https://www.fast-ro.com/service/service_03_apply.php?code=' + shNo;
    
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('🎯 테스트 ' + (ri+1) + '/' + testCount);
    Logger.log('   SH번호: ' + shNo);
    Logger.log('   URL: ' + url);
    Logger.log('   IT개수: ' + rec.items.length);
    if (rec.items.length > 0) {
      Logger.log('   첫 IT: ' + rec.items[0].itCode);
      Logger.log('   첫 타오주문: ' + rec.items[0].orderNo);
    }
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    
    var options = {
      method: 'get',
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        Cookie: cookie,
        Referer: 'https://www.fast-ro.com/mypage/service_list.php',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'ko-KR,ko;q=0.9'
      }
    };
    
    try {
      var res = UrlFetchApp.fetch(url, options);
      var sc = res.getResponseCode();
      var html = res.getContentText('UTF-8');
      
      Logger.log('   status: ' + sc);
      Logger.log('   응답크기: ' + html.length + '자');
      
      if (sc !== 200 || html.length < 1000) {
        Logger.log('   ❌ 응답 작거나 에러');
        // 응답 일부 보여주기
        Logger.log('   --- 응답 시작 ---');
        Logger.log('   ' + html.substring(0, 500));
        continue;
      }
      
      // 키워드 검사
      var hasTaobao = html.indexOf('taobao.com') >= 0;
      var hasTmall = html.indexOf('tmall.com') >= 0;
      var has1688 = html.indexOf('1688.com') >= 0;
      var hasItemHtm = html.indexOf('item.htm') >= 0;
      var hasDetailUrl = html.indexOf('상세URL') >= 0;
      var hasTrackingNo = html.indexOf('트레킹번호') >= 0;
      
      Logger.log('   ✓ 타오바오 포함: ' + hasTaobao);
      Logger.log('   ✓ 티몰 포함: ' + hasTmall);
      Logger.log('   ✓ 1688 포함: ' + has1688);
      Logger.log('   ✓ item.htm 포함: ' + hasItemHtm);
      Logger.log('   ✓ "상세URL" 라벨: ' + hasDetailUrl);
      Logger.log('   ✓ "트레킹번호" 라벨: ' + hasTrackingNo);
      
      // 정규식 매칭 시도
      var patterns = [
        { name: 'taobao item.htm',  re: /https?:\/\/item\.taobao\.com\/item\.htm\?[^"'\s<>)]+/g },
        { name: 'tmall detail',     re: /https?:\/\/detail\.tmall\.com\/item\.htm\?[^"'\s<>)]+/g },
        { name: '1688 detail',      re: /https?:\/\/detail\.1688\.com\/offer\/[^"'\s<>)]+/g },
        { name: 'm.taobao',         re: /https?:\/\/m\.intl\.taobao\.com\/[^"'\s<>)]+/g },
        { name: 'world.taobao',     re: /https?:\/\/world\.taobao\.com\/[^"'\s<>)]+/g },
      ];
      
      Logger.log('   ─── 정규식 매칭 ───');
      var foundAny = false;
      for (var pi = 0; pi < patterns.length; pi++) {
        var p = patterns[pi];
        var matches = html.match(p.re);
        if (matches && matches.length > 0) {
          foundAny = true;
          Logger.log('   🎯 [' + p.name + '] ' + matches.length + '개 발견');
          for (var mi = 0; mi < Math.min(2, matches.length); mi++) {
            Logger.log('      ' + (mi+1) + '. ' + matches[mi].substring(0, 250));
          }
        }
      }
      
      if (!foundAny) {
        Logger.log('   ⚠ 위 패턴으로 매칭 없음');
      }
      
      // "상세URL" 글자 주변 HTML
      var detailIdx = html.indexOf('상세URL');
      if (detailIdx >= 0) {
        Logger.log('   ─── "상세URL" 주변 800자 ───');
        var snippet = html.substring(Math.max(0, detailIdx - 100), detailIdx + 700);
        Logger.log('   ' + snippet);
      }
      
      // input 태그 검사 (상세URL이 input value에 있을 수 있음)
      var inputRe = /<input[^>]*value\s*=\s*["']([^"']*taobao[^"']*)["'][^>]*>/g;
      var inputMatches = [];
      var im;
      while ((im = inputRe.exec(html)) !== null) {
        inputMatches.push(im[1]);
      }
      if (inputMatches.length > 0) {
        Logger.log('   ─── <input value="..."> 안의 타오바오 URL ───');
        for (var ii = 0; ii < Math.min(3, inputMatches.length); ii++) {
          Logger.log('   ' + (ii+1) + '. ' + inputMatches[ii].substring(0, 250));
        }
      }
      
      // textarea 안에 있을 수도
      var taRe = /<textarea[^>]*>[\s\S]*?taobao[\s\S]*?<\/textarea>/g;
      var taMatches = html.match(taRe);
      if (taMatches && taMatches.length > 0) {
        Logger.log('   ─── <textarea> 안의 타오바오 URL ───');
        Logger.log('   ' + taMatches[0].substring(0, 500));
      }
      
    } catch(e) {
      Logger.log('   ❌ 예외: ' + e.message);
    }
    
    Utilities.sleep(500);
  }
  
  Logger.log('\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('✅ 디버그 완료. 로그 캡처해서 Claude에게 보내주세요.');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
}



// ★ 디버그: 출고완료 건의 운송료/선불 패턴 확인
function debugFastroFee() {
  var ssId = '1';
  var mbId = getFastroMbId_(ssId);
  var today = new Date();
  var past = new Date(today.getTime() - 60 * 24 * 60 * 60 * 1000);  // 60일 (출고완료 건 찾기)
  var fmt = function(d) {
    return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
  };
  var result = fastroFetch_(ssId, mbId, fmt(past), fmt(today), 200);
  if (!result.ok) { Logger.log('실패: ' + result.error); return; }
  
  var html = result.html;
  
  // KRW 단어 검색
  var krwIdx = html.indexOf('KRW');
  if (krwIdx < 0) {
    Logger.log('❌ KRW 단어 없음 - 운송료 패턴이 다를 수 있음');
    // "원" 검색
    var wonIdx = html.indexOf('선불');
    if (wonIdx >= 0) {
      Logger.log('✅ "선불" 단어 발견 위치: ' + wonIdx);
      var snippet = html.substring(Math.max(0, wonIdx - 200), wonIdx + 200);
      Logger.log('--- 선불 주변 HTML ---');
      Logger.log(snippet);
    }
    return;
  }
  
  Logger.log('✅ KRW 단어 발견 위치: ' + krwIdx);
  var snippet1 = html.substring(Math.max(0, krwIdx - 300), krwIdx + 300);
  Logger.log('--- KRW 주변 HTML ---');
  Logger.log(snippet1);
  
  // 선불 검색
  var preIdx = html.indexOf('선불');
  if (preIdx >= 0) {
    Logger.log('--- 선불 주변 HTML ---');
    var snippet2 = html.substring(Math.max(0, preIdx - 100), preIdx + 200);
    Logger.log(snippet2);
  } else {
    Logger.log('❌ "선불" 단어 없음');
  }
}
// ─── 디버그: 패스트로 응답 HTML을 검사해서 경동택배 패턴 찾기 ───
// 사용법: 함수 드롭다운에서 debugFastroKdong 선택 → ▶ 실행
// → 실행 로그에서 경동택배 주변 HTML 확인
function debugFastroKdong() {
  var ssId = '1';  // ssId=1 사용 (필요시 변경)
  var mbId = getFastroMbId_(ssId);
  var today = new Date();
  var past = new Date(today.getTime() - 30 * 24 * 60 * 60 * 1000);
  var fmt = function(d) {
    return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
  };
  var result = fastroFetch_(ssId, mbId, fmt(past), fmt(today), 200);
  if (!result.ok) {
    Logger.log('실패: ' + result.error);
    return;
  }
  var html = result.html;
  Logger.log('전체 HTML 크기: ' + html.length + ' 자');
  // '경동' 단어가 있는지 확인
  var keywordIdx = html.indexOf('경동');
  if (keywordIdx < 0) {
    Logger.log('❌ "경동" 단어가 응답에 없음 - 출고완료 건이 30일 안에 없거나 경동 사용 안 함');
    return;
  }
  Logger.log('✅ "경동" 단어 발견 위치: ' + keywordIdx);
  // 경동 주변 800자 출력 (이걸로 정규식 패턴 확인)
  var snippet = html.substring(Math.max(0, keywordIdx - 100), keywordIdx + 700);
  Logger.log('--- 경동 주변 HTML ---');
  Logger.log(snippet);
  // 모든 경동 위치 찾기
  var allIdx = [];
  var idx = 0;
  while ((idx = html.indexOf('경동', idx)) >= 0) {
    allIdx.push(idx);
    idx++;
    if (allIdx.length > 5) break;
  }
  Logger.log('--- 경동 단어 위치 (최대 5개) ---');
  Logger.log(allIdx.join(', '));
}

// ─── 6개 사업자 쿠키 일괄 등록 ───
function setupAll() {
  var sameCookie = 'PHPSESSID=noeospqnpdhufiuprpdli80pr6; hd_pops_78=1; hd_pops_77=1';
  setFastroCookieFor('1', sameCookie);
  setFastroCookieFor('2', sameCookie);
  setFastroCookieFor('3', sameCookie);
  setFastroCookieFor('4', sameCookie);
  setFastroCookieFor('5', sameCookie);
  setFastroCookieFor('6', sameCookie);
  Logger.log('전체 등록 완료');
  listFastroCookies();
}

// ─── 쿠키 만료 시 갱신용 ───
// 사용법: 새 PHPSESSID 받아서 따옴표 안 값만 교체하고 실행
function refreshAllCookies() {
  var newCookie = 'PHPSESSID=여기에_새값; hd_pops_78=1; hd_pops_77=1';
  for (var i = 1; i <= 6; i++) {
    setFastroCookieFor(String(i), newCookie);
  }
  Logger.log('전체 쿠키 갱신 완료');
  listFastroCookies();
}

// ─── 테스트 함수들 (시트 안 건드림) ───
function testFastroDryRun()    { _testDry_('1'); }
function testFastroDryRun_2()  { _testDry_('2'); }
function testFastroDryRun_3()  { _testDry_('3'); }
function testFastroDryRun_4()  { _testDry_('4'); }
function testFastroDryRun_5()  { _testDry_('5'); }
function testFastroDryRun_6()  { _testDry_('6'); }
function _testDry_(ssId) {
  var fakeEvent = { parameter: { action: 'fastroDryRun', ssId: ssId } };
  var result = doGet(fakeEvent);
  Logger.log('=== ssId=' + ssId + ' ===');
  Logger.log(result.getContent());
}

// 6개 사업자 일괄 DryRun (간단 결과만)
function testAllDryRun() {
  for (var i = 1; i <= 6; i++) {
    var ssId = String(i);
    var fakeEvent = { parameter: { action: 'fastroDryRun', ssId: ssId } };
    try {
      var result = doGet(fakeEvent);
      var parsed = JSON.parse(result.getContent());
      if (parsed.success) {
        Logger.log('✅ ssId=' + ssId + ' (' + parsed.mbId + '): ' + parsed.recordCount + '건 발견');
      } else {
        Logger.log('❌ ssId=' + ssId + ': ' + (parsed.error || '알 수 없음'));
      }
    } catch (err) {
      Logger.log('❌ ssId=' + ssId + ' 예외: ' + err.message);
    }
    Utilities.sleep(1000);
  }
}

// ─── 실제 동기화 함수들 ───
function runFastroSync_ssId1() { _runSync_('1'); }
function runFastroSync_ssId2() { _runSync_('2'); }
function runFastroSync_ssId3() { _runSync_('3'); }
function runFastroSync_ssId4() { _runSync_('4'); }
function runFastroSync_ssId5() { _runSync_('5'); }
function runFastroSync_ssId6() { _runSync_('6'); }
function _runSync_(ssId) {
  var fakeEvent = { parameter: { action: 'fastroSync', ssId: ssId } };
  var result = doGet(fakeEvent);
  Logger.log('=== ssId=' + ssId + ' ===');
  Logger.log(result.getContent());
}

// 6개 사업자 일괄 동기화 (트리거가 매시간 실행)
function runFastroSyncAll() {
  for (var i = 1; i <= 6; i++) {
    var ssId = String(i);
    var fakeEvent = { parameter: { action: 'fastroSync', ssId: ssId } };
    try {
      var result = doGet(fakeEvent);
      Logger.log('=== ssId=' + ssId + ' ===');
      Logger.log(result.getContent());
    } catch (err) {
      Logger.log('ssId=' + ssId + ' 오류: ' + err.message);
    }
    Utilities.sleep(1000);
  }
}

// ═══════════════════════════════════════════════════════════════
// ★ 이전 패스트로 동기화로 잘못 칠해진 H열 색깔 복원
// ═══════════════════════════════════════════════════════════════
// - 이전 버전이 H열을 #b6d7a8(초록) 또는 #ff9800(주황)으로 칠함
// - 기존 색깔 시스템과 충돌 → 흰색으로 리셋
// - 노트는 유지 (정보 보존)
function resetFastroColors_(ssId) {
  var SS_MAP = {
    '1': '10hfa94OWjpyvF4eK5w2FYhqsIRbfsn8pPlISoj9N9Jg',
    '2': '1u2kIj8-c6YVtv1XOg40-Im1dL09sGiIxehatDKD8Cu0',
    '3': '137pCqtBMsOgX752pVoglwBk3NyK2aXtAy1R40lLJFiM',
    '4': '1AVe0jVZXFLloVEbJHh65OSru8ZbWnXwplHxAmbaXkGE',
    '5': '1IZEnu3o7WcpRUXJP3AcxPQyNOsUj1J2bTSjvpX14oGo',
    '6': '1FCvAfy4AMClQ-0pdJ_AjdwAOqh05Ldh80fTSXknI6Nc'
  };
  var ss;
  if (ssId === '1') {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } else {
    ss = SpreadsheetApp.openById(SS_MAP[ssId]);
  }
  var monthSheets = ss.getSheets().filter(function(s) {
    return /^\d{2}년 \d+월$/.test(s.getName());
  });
  monthSheets.sort(function(a, b) {
    var pa = a.getName().match(/(\d+)년 (\d+)월/);
    var pb = b.getName().match(/(\d+)년 (\d+)월/);
    if (!pa || !pb) return 0;
    return (Number(pb[1]) * 100 + Number(pb[2])) - (Number(pa[1]) * 100 + Number(pa[2]));
  });
  monthSheets = monthSheets.slice(0, 3);

  var totalReset = 0;
  monthSheets.forEach(function(sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    var hRange = sheet.getRange(2, 8, lastRow - 1, 1);
    var bgs = hRange.getBackgrounds();
    var notes = hRange.getNotes();
    var changed = false;
    for (var i = 0; i < bgs.length; i++) {
      var bg = (bgs[i][0] || '').toLowerCase().replace('#', '');
      var note = notes[i][0] || '';
      // 패스트로가 칠한 색이고, 노트에 [패스트로]가 있는 경우만 흰색으로 복원
      if (note.indexOf('[패스트로]') >= 0 && (bg === 'b6d7a8' || bg === 'ff9800')) {
        bgs[i][0] = '#ffffff';
        changed = true;
        totalReset++;
      }
    }
    if (changed) hRange.setBackgrounds(bgs);
  });
  return totalReset;
}

// ★ 6개 사업자 모두 색깔 복원 (한 번만 실행하면 됨!) ★
function resetAllFastroColors() {
  var grandTotal = 0;
  for (var i = 1; i <= 6; i++) {
    try {
      var n = resetFastroColors_(String(i));
      grandTotal += n;
      Logger.log('ssId=' + i + ' → 복원된 셀: ' + n + '개');
    } catch (err) {
      Logger.log('ssId=' + i + ' 오류: ' + err.message);
    }
    Utilities.sleep(500);
  }
  Logger.log('═══════════════════');
  Logger.log('전체 복원 완료: ' + grandTotal + '개 셀 (노트는 유지됨)');
}

// ═══════════════════════════════════════════════════════════════
// ★ 기존 LCL 버튼이 F열에 남긴 'LCL출고' 메모 정리
// ═══════════════════════════════════════════════════════════════
// - 옛 LCL 버튼이 F열(판매처) 셀에 setNote('LCL출고') 했음
// - 이젠 패스트로해운 드롭다운으로 통합되었으므로 메모 제거 필요
function resetLclNotes_(ssId) {
  var SS_MAP = {
    '1': '10hfa94OWjpyvF4eK5w2FYhqsIRbfsn8pPlISoj9N9Jg',
    '2': '1u2kIj8-c6YVtv1XOg40-Im1dL09sGiIxehatDKD8Cu0',
    '3': '137pCqtBMsOgX752pVoglwBk3NyK2aXtAy1R40lLJFiM',
    '4': '1AVe0jVZXFLloVEbJHh65OSru8ZbWnXwplHxAmbaXkGE',
    '5': '1IZEnu3o7WcpRUXJP3AcxPQyNOsUj1J2bTSjvpX14oGo',
    '6': '1FCvAfy4AMClQ-0pdJ_AjdwAOqh05Ldh80fTSXknI6Nc'
  };
  var ss;
  if (ssId === '1') {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } else {
    ss = SpreadsheetApp.openById(SS_MAP[ssId]);
  }
  var monthSheets = ss.getSheets().filter(function(s) {
    return /^\d{2}년 \d+월$/.test(s.getName());
  });
  var totalReset = 0;
  monthSheets.forEach(function(sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    // F열(6번째)의 메모 가져오기
    var fRange = sheet.getRange(2, 6, lastRow - 1, 1);
    var notes = fRange.getNotes();
    var changed = false;
    for (var i = 0; i < notes.length; i++) {
      var note = (notes[i][0] || '').trim();
      // 'LCL출고' 메모만 제거
      if (note === 'LCL출고') {
        notes[i][0] = '';
        changed = true;
        totalReset++;
      }
    }
    if (changed) fRange.setNotes(notes);
  });
  // 캐시 무효화
  clearSearchCache_(ssId);
  return totalReset;
}

// ★ 6개 사업자 모두 LCL출고 메모 정리 (한 번만 실행하면 됨!) ★
function resetAllLclNotes() {
  var grandTotal = 0;
  for (var i = 1; i <= 6; i++) {
    try {
      var n = resetLclNotes_(String(i));
      grandTotal += n;
      Logger.log('ssId=' + i + ' → 정리된 메모: ' + n + '개');
    } catch (err) {
      Logger.log('ssId=' + i + ' 오류: ' + err.message);
    }
    Utilities.sleep(500);
  }
  Logger.log('═══════════════════');
  Logger.log('전체 LCL출고 메모 정리 완료: ' + grandTotal + '개');
}

// ═══════════════════════════════════════════════════════════════
// ★ 기존 tao트래킹발급 액션이 G열에 칠한 주황색 정리
// ═══════════════════════════════════════════════════════════════
// - 옛 tao트래킹발급 버튼이 G열(주문번호)을 #FF9800 으로 칠함
// - 이젠 그 액션이 제거되어 더 이상 칠해지지 않으므로 정리
// - 노란색(#FFFF00)이나 다른 색은 그대로 유지
function resetOrangeG_(ssId) {
  var SS_MAP = {
    '1': '10hfa94OWjpyvF4eK5w2FYhqsIRbfsn8pPlISoj9N9Jg',
    '2': '1u2kIj8-c6YVtv1XOg40-Im1dL09sGiIxehatDKD8Cu0',
    '3': '137pCqtBMsOgX752pVoglwBk3NyK2aXtAy1R40lLJFiM',
    '4': '1AVe0jVZXFLloVEbJHh65OSru8ZbWnXwplHxAmbaXkGE',
    '5': '1IZEnu3o7WcpRUXJP3AcxPQyNOsUj1J2bTSjvpX14oGo',
    '6': '1FCvAfy4AMClQ-0pdJ_AjdwAOqh05Ldh80fTSXknI6Nc'
  };
  var ss;
  if (ssId === '1') {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } else {
    ss = SpreadsheetApp.openById(SS_MAP[ssId]);
  }
  var monthSheets = ss.getSheets().filter(function(s) {
    return /^\d{2}년 \d+월$/.test(s.getName());
  });
  var totalReset = 0;
  // 주황색 후보들 (Material Design 주황 계열)
  var orangeColors = ['ff9800', 'ff6d00', 'ff7043', 'ff5722', 'e65100', 'f57c00', 'fb8c00', 'ff8a65'];
  monthSheets.forEach(function(sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    // G열(7번째)
    var gRange = sheet.getRange(2, 7, lastRow - 1, 1);
    var bgs = gRange.getBackgrounds();
    var changed = false;
    for (var i = 0; i < bgs.length; i++) {
      var bg = (bgs[i][0] || '').toLowerCase().replace('#', '');
      if (orangeColors.indexOf(bg) >= 0) {
        bgs[i][0] = '#ffffff';
        changed = true;
        totalReset++;
      }
    }
    if (changed) gRange.setBackgrounds(bgs);
  });
  // 캐시 무효화
  clearSearchCache_(ssId);
  return totalReset;
}

// ★ 6개 사업자 모두 G열 주황색 정리 (한 번만 실행하면 됨!) ★
function resetAllOrangeG() {
  var grandTotal = 0;
  for (var i = 1; i <= 6; i++) {
    try {
      var n = resetOrangeG_(String(i));
      grandTotal += n;
      Logger.log('ssId=' + i + ' → 정리된 주황 셀: ' + n + '개');
    } catch (err) {
      Logger.log('ssId=' + i + ' 오류: ' + err.message);
    }
    Utilities.sleep(500);
  }
  Logger.log('═══════════════════');
  Logger.log('전체 G열 주황색 정리 완료: ' + grandTotal + '개');
}
function parseFastroHtml_(html) {
  var results = [];
  var parts = html.split(/<td colspan="4" class="new_hap"/);
  for (var i = 1; i < parts.length; i++) {
    var part = parts[i];
    var grM = part.match(/GR(\d+)<font[^>]*>(\d+)<\/font>/);
    if (!grM) continue;
    var gr = 'GR' + grM[1] + grM[2];
    var shM = part.match(/SH<\/font>(\d+)<font[^>]*>(\d+)<\/font>/);
    var sh = shM ? ('SH' + shM[1] + shM[2]) : '';
    var stM = part.match(/class="color_\d+">([^<]+?)<\/span>/);
    var state = stM ? stM[1].trim() : '';
    var invM = part.match(/invoice=(\d+)/);
    var invoice = invM ? invM[1] : '';
    // ★ 경동택배 번호
    var kdongInvoice = '';
    var kdongPatterns = [
      /경동택배[\s\S]*?<a[^>]*id=["']?(\d{10,14})["']?/i,
      /경동택배[\s\S]*?id=["']?(\d{10,14})["']?/i,
      /경동택배[\s\S]*?>(\d{10,14})<\/a>/,
      /경동택배[\s\S]{0,300}?(\d{10,14})/
    ];
    for (var pi = 0; pi < kdongPatterns.length; pi++) {
      var km = part.match(kdongPatterns[pi]);
      if (km && km[1] && km[1] !== invoice) {
        kdongInvoice = km[1];
        break;
      }
    }
    var dtM = part.match(/(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})/);
    var regDate = dtM ? dtM[1] : '';
    
    // ★★★ 운송료 추출 (예: 16,100KRW) ★★★
    var shipFee = 0;
    // ★★★ 운송료 추출 (예: <b>6,600</b>KRW) ★★★
    var shipFee = 0;
    var feeM = part.match(/<b>([\d,]+)<\/b>\s*KRW/);
    if (feeM) {
      shipFee = parseInt(feeM[1].replace(/,/g, ''), 10) || 0;
    }
    
    // ★★★ 선불 추출 (예: <b>선불12800원</b>) ★★★
    var prepaid = 0;
    var prePatterns = [
      /선불([\d,]+)원/,
      /<b>선불([\d,]+)원<\/b>/,
      /선불\s*([\d,]+)\s*원/
    ];
    for (var pp = 0; pp < prePatterns.length; pp++) {
      var preM = part.match(prePatterns[pp]);
      if (preM) {
        prepaid = parseInt(preM[1].replace(/,/g, ''), 10) || 0;
        if (prepaid > 0) break;
      }
    }
    
    var itemRe = /IT(\d+)<font[^>]*>(\d+)<\/font>[\s\S]*?주문번호\s*<span>(\d+)<\/span>/g;
    var items = [];
    var im;
    while ((im = itemRe.exec(part)) !== null) {
      items.push({ itCode: 'IT' + im[1] + im[2], orderNo: im[3] });
    }
    results.push({
      gr: gr, sh: sh, state: state, invoice: invoice, kdongInvoice: kdongInvoice,
      regDate: regDate, items: items,
      shipFee: shipFee,    // ★ 운송료
      prepaid: prepaid     // ★ 선불 (0이면 선불 아님)
    });
  }
  return results;
}

function buildFastroNote_(rec, item) {
  var lines = [];
  lines.push('[패스트로]');
  if (rec.gr)      lines.push('GR: ' + rec.gr);
  if (rec.sh)      lines.push('SH: ' + rec.sh);
  if (item.itCode) lines.push('IT: ' + item.itCode);
  if (rec.invoice) lines.push('한국운송장: ' + rec.invoice);
  if (rec.kdongInvoice) lines.push('경동택배: ' + rec.kdongInvoice);
  if (rec.state)   lines.push('상태: ' + rec.state);
  if (rec.regDate) lines.push('신청일: ' + rec.regDate);
  lines.push('동기화: ' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'MM-dd HH:mm'));
  return lines.join('\n');
}

function fastroFetch_(ssId, mbId, sdate, edate, limit) {
  var cookie = getFastroCookie_(ssId);
  if (!cookie) {
    return { ok: false, error: '쿠키 없음. setFastroCookieFor("' + ssId + '", "...") 실행 필요.' };
  }
  var url = 'https://www.fast-ro.com/elpisbbs/ajax.nt_order_list_member.php' +
            '?last=0&limit=' + (limit || 200) +
            '&find=&value=&or_de_no=&state=' +
            '&sdate=' + sdate + '&edate=' + edate +
            '&mb_id=' + encodeURIComponent(mbId) +
            '&type=&last_code=&it_code=&dtype=' +
            '&gr_output_stay_type=&gr_var5=&gr_unipass_result=' +
            '&gr_fltno=&gr_fltno2=';
  var options = {
    method: 'post',
    muteHttpExceptions: true,
    headers: {
      Cookie: cookie,
      Referer: 'https://www.fast-ro.com/mypage/service_list.php',
      Origin:  'https://www.fast-ro.com',
      'X-Requested-With': 'XMLHttpRequest',
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36',
      'Accept': 'text/html, */*; q=0.01',
      'Accept-Language': 'ko-KR,ko;q=0.9'
    }
  };
  try {
    var res = UrlFetchApp.fetch(url, options);
    var sc = res.getResponseCode();
    if (sc !== 200) return { ok: false, error: '응답 status=' + sc };
    var html = res.getContentText('UTF-8');
    if (html.indexOf('GR') === -1 && html.length < 500) {
      return { ok: false, error: '로그인 만료 또는 권한 없음 (응답 ' + html.length + '자, GR없음). 쿠키 갱신 필요.' };
    }
    return { ok: true, html: html };
  } catch (err) {
    return { ok: false, error: 'fetch 예외: ' + err.message };
  }
}

/**
 * @OnlyCurrentDoc false
 */

function doGet(e) {
  // ★ 웹 주소로 접속 시 (action 파라미터 없으면) HTML 페이지 표시 ★
  if (!e || !e.parameter || !e.parameter.action) {
    const tpl = HtmlService.createTemplateFromFile('index');
    // 현재 웹 앱 URL을 HTML에 주입 (HTML에서 GAS_BASE_AUTO로 사용 가능)
    tpl.gasUrl = ScriptApp.getService().getUrl();
    return tpl.evaluate()
      .setTitle('MODA Hub - 통합 판매관리')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }

  const callback = e.parameter.callback || '';
  const respond = (data) => {
    const json = JSON.stringify(data);
    if (callback) {
      return ContentService.createTextOutput(callback + '(' + json + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  };

  // 사업자별 스프레드시트 선택
  const SS_MAP = {
    '1': '10hfa94OWjpyvF4eK5w2FYhqsIRbfsn8pPlISoj9N9Jg',
    '2': '1u2kIj8-c6YVtv1XOg40-Im1dL09sGiIxehatDKD8Cu0',
    '3': '137pCqtBMsOgX752pVoglwBk3NyK2aXtAy1R40lLJFiM',
    '4': '1AVe0jVZXFLloVEbJHh65OSru8ZbWnXwplHxAmbaXkGE',
    '5': '1IZEnu3o7WcpRUXJP3AcxPQyNOsUj1J2bTSjvpX14oGo',
    '6': '1FCvAfy4AMClQ-0pdJ_AjdwAOqh05Ldh80fTSXknI6Nc',
  };
  const ssId = e.parameter.ssId || '1';
  const targetId = SS_MAP[ssId] || SS_MAP['1'];
  let ss;
  try {
    ss = ssId === '1'
      ? SpreadsheetApp.getActiveSpreadsheet()
      : SpreadsheetApp.openById(targetId);
  } catch(openErr) {
    return respond({ error: '스프레드시트 접근 실패: ' + openErr.message });
  }
  const action = e.parameter.action || 'list';

  if (action === 'list') {
    const smsSheet = ss.getSheetByName('문자발송');
    if (!smsSheet) return respond({ error: '문자발송 시트를 찾을 수 없습니다.' });
    const data = smsSheet.getDataRange().getValues();
    const customers = [];
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const colB = String(row[1]).trim();
      const colC = String(row[2]).trim();
      const colD = String(row[3]).trim();
      const colE = String(row[4]).trim();
      const colG = String(row[6]).trim();
      const colH = String(row[7]).trim();
      if (colB.length >= 10 && /^\d+$/.test(colB) && colC !== '') {
        let phone = '';
        if (colH.startsWith('010')) phone = colH.replace(/-/g, '');
        else if (colG.startsWith('010')) phone = colG.replace(/-/g, '');
        if (!phone) continue;
        let source = '', customs = '';
        for (let j = i + 1; j < Math.min(i + 5, data.length); j++) {
          const hVal = String(data[j][7]).trim();
          const dVal = String(data[j][3]).trim();
          const eVal = String(data[j][4]).trim();
          if (hVal && !hVal.startsWith('010') && hVal.length > 0) source = hVal;
          if (dVal === '통관부호' && eVal) customs = eVal;
        }
        const skipWords = ['주소록 담당자','사업자','주문번호','자동'];
        let msgLines = [], msgStarted = false;
        for (let j = i + 1; j < Math.min(i + 40, data.length); j++) {
          const nextB = String(data[j][1]).trim();
          const nextC = String(data[j][2]).trim();
          if (nextB.length >= 10 && /^\d+$/.test(nextB)) break;
          if (skipWords.includes(nextB)) { if (!msgStarted) continue; else break; }
          if (!msgStarted) { if (nextB.endsWith('님')) { msgStarted = true; msgLines.push(nextB); } continue; }
          if (/^010-?\d{4}-?\d{4}$/.test(nextB)) continue;
          if (nextB !== '') {
            if (nextB.startsWith('[옵션]') && nextC !== '') msgLines.push(nextB + '\t' + nextC);
            else msgLines.push(nextB);
          } else {
            if (!nextC.startsWith('수량') && msgLines.length > 0 && msgLines[msgLines.length - 1] !== '') msgLines.push('');
          }
        }
        while (msgLines.length > 0 && msgLines[msgLines.length - 1] === '') msgLines.pop();
        customers.push({
          idx: customers.length + 1,
          orderNo: colB, name: colC, phone, product: colD, option: colE,
          source, customs, message: msgLines.join("\n")
        });
      }
    }
    return respond({ customers });
  }

  if (action === 'scripts') {
    const smsSheet = ss.getSheetByName('문자발송');
    if (!smsSheet) return respond({ error: '문자발송 시트를 찾을 수 없습니다.' });
    const data = smsSheet.getDataRange().getValues();
    const scripts = [];
    for (let i = 0; i < data.length; i++) {
      const colA = String(data[i][0]).trim();
      const colB = String(data[i][1]).trim();
      if (colA !== '' && colB !== '' && !/^\d{15,}$/.test(colB) &&
          !['주소록 담당자','사업자','주문번호','고객명','자동'].includes(colA)) {
        scripts.push({ idx: scripts.length + 1, title: colA, content: colB });
      }
    }
    return respond({ scripts });
  }

  if (action === 'saveContact') {
    const name    = e.parameter.name    || '';
    const phone   = e.parameter.phone   || '';
    const product = e.parameter.product || '';
    const source  = e.parameter.source  || '';
    const memo_f  = e.parameter.memo_f  || '';
    const memo_j  = e.parameter.memo_j  || '';
    if (!name || !phone) return respond({ error: '이름 또는 전화번호가 없습니다.' });
    try {
      try {
        const searchResult = People.People.searchContacts({
          query: phone.replace(/\D/g, ''),
          readMask: 'phoneNumbers,names'
        });
        if (searchResult.results && searchResult.results.length > 0) return respond({ error: 'DUPLICATE' });
      } catch(e) {}
      let lastName = '';
      const dashIdx = source.indexOf('-');
      if (dashIdx > 0) lastName = source.substring(0, dashIdx).trim();
      const firstName = name + '/' + product;
      const memo = (memo_f ? memo_f : '') + (memo_j ? (memo_f ? '\n' : '') + memo_j : '');
      People.People.createContact({
        names: [{ givenName: firstName, familyName: lastName }],
        phoneNumbers: [{ value: phone, type: 'mobile' }],
        biographies: [{ value: memo, contentType: 'TEXT_PLAIN' }]
      });
      return respond({ success: true });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // 통관부호 검증
  if (action === 'verifyContact') {
    const unipass = (e.parameter.unipass || '').trim().toUpperCase();
    const name    = (e.parameter.name    || '').trim();
    const phone   = (e.parameter.phone   || '').replace(/-/g, '');
    if (!unipass || !name || !phone) return respond({ result: '⚠️ 데이터없음', msg: '이름, 통관번호, 전화번호가 필요합니다.' });
    if (!/^P[0-9]{12}$/.test(unipass)) return respond({ result: '❌ 형식오류', msg: '올바른 형식: P + 12자리 숫자' });
    try {
      const url = 'https://sellochomes.co.kr/api/v1/sellerlife/unipass/unipass';
      const callApi = (ph) => {
        const payload = 'persEcm=' + unipass + '&pltxNm=' + encodeURIComponent(name) + '&cralTelno=' + ph + '&addrNo=';
        const options = {
          method: 'post', contentType: 'application/x-www-form-urlencoded',
          payload, muteHttpExceptions: true,
          headers: { Cookie: getLoginCookie_(), Referer: 'https://sellochomes.co.kr/sellerlife/unipass/' }
        };
        const res  = UrlFetchApp.fetch(url, options);
        const json = JSON.parse(res.getContentText());
        const data = json.data || '';
        const tCntMatch = data.match(/<tCnt>(\d+)<\/tCnt>/);
        const tCnt = tCntMatch ? parseInt(tCntMatch[1]) : 0;
        if (tCnt >= 1) return '✅ 정상';
        if (data.indexOf('존재하지 않습니다') > -1) return '❌ 번호없음';
        if (data.indexOf('불일치') > -1 || data.indexOf('일치하지 않습니다') > -1) return '❌ 불일치';
        return '⚠️ 확인필요';
      };
      let result = callApi(phone);
      const phone2 = (e.parameter.phone2 || '').replace(/-/g, '');
      if (result !== '✅ 정상' && phone2 && phone2 !== phone && /^010/.test(phone2)) {
        const result2 = callApi(phone2);
        if (result2 === '✅ 정상') result = result2;
      }
      const msgMap = {
        '✅ 정상': '이름과 통관번호가 일치합니다.',
        '❌ 번호없음': '등록되지 않은 통관번호입니다.',
        '❌ 불일치': '이름 또는 전화번호가 일치하지 않습니다.',
        '⚠️ 확인필요': '셀러라이프에서 직접 확인이 필요합니다.'
      };
      const orderNo2 = e.parameter.orderNo || '';
      if (orderNo2) {
        try {
          const m = getCurrentMonthSheet_(ss);
          if (m.sheet) {
            const found = findRowByOrder_(m.sheet, orderNo2);
            if (found.row > 0) {
              m.sheet.getRange(found.row, 18).setValue(result);
            }
          }
        } catch(e2) {}
      }
      return respond({ result, msg: msgMap[result] || '' });
    } catch(err) {
      return respond({ result: '⚠️ 오류', msg: err.message });
    }
  }

  // ★★★ fast 액션 ★★★
  if (action === 'fast') {
    try {
      const allSheets = ss.getSheets();
      const monthSheets = allSheets.filter(s => /^\d{2}년 \d+월$/.test(s.getName()));
      monthSheets.sort((a, b) => {
        const pa = a.getName().match(/(\d+)년 (\d+)월/);
        const pb = b.getName().match(/(\d+)년 (\d+)월/);
        if (!pa || !pb) return 0;
        return (Number(pb[1])*100+Number(pb[2])) - (Number(pa[1])*100+Number(pa[2]));
      });
      const monthNames = monthSheets.map(s => s.getName());

      const m = getCurrentMonthSheet_(ss);
      const mSheet = m.sheet;
      if (!mSheet) return respond({ error: '월별 시트를 찾을 수 없습니다.' });

      const mName = mSheet.getName();
      const mData = mSheet.getDataRange().getValues();
      const lastRow = mData.length;
      const allBg = lastRow > 1 ? mSheet.getRange(2, 1, lastRow-1, 12).getBackgrounds() : [];
      const allNotes = lastRow > 1 ? mSheet.getRange(2, 1, lastRow-1, 12).getNotes() : [];

      const monthList = [];
      for (let i = 1; i < lastRow; i++) {
        const row = mData[i];
        const orderNo = String(row[6]).trim();
        const name    = String(row[7]).trim();
        if (!orderNo || !name) continue;

        let phone1 = String(row[12]).trim().replace(/-/g, '');
        let phone2 = String(row[13]).trim().replace(/-/g, '');
        if (phone1 && !phone1.startsWith('0') && phone1.length === 10) phone1 = '0' + phone1;
        if (phone2 && !phone2.startsWith('0') && phone2.length === 10) phone2 = '0' + phone2;
        let phone = phone2.startsWith('010') ? phone2 : phone1.startsWith('010') ? phone1 : (phone2 || phone1);
        if (!phone) continue;

        const noteRow = allNotes[i-1] || [];

        monthList.push({
          row: i+1, sheetName: mName, orderNo, name, phone, phone1, phone2,
          product:  String(row[8]).trim(),
          option:   String(row[9]).trim(),
          orderDate: row[4] || '', 
          qty:      String(row[10]).trim(),
          customs:  String(row[11]).trim(),
          baejiNo:  String(row[2]).trim(),
          status:   String(row[0]).trim(),
          ship:     String(row[5]).trim(),
          source:   String(row[5]).trim(),
          memo2:    String(row[15]).trim(),
          address:  String(row[15]).trim(),
          zipCode:  String(row[14]).trim(),
          deliveryType: String(row[35] || '').trim(),
          message: buildMsg_(row, ssId),
          payAmt:    row[16] || '', settleAmt: row[20] || '',
          taoAmt:    row[22] || '', shipFee:   row[26] || '',
          expShip:   row[27] || '', addShip:   row[28] || '',
          profitRate: row[32] || '', marginRate: row[33] || '',
          eBg: allBg[i-1] ? (allBg[i-1][4]||'').toLowerCase() : '',
          eBg: allBg[i-1] ? (allBg[i-1][4]||'').toLowerCase() : '',  
          fBg: allBg[i-1] ? (allBg[i-1][5]||'').toLowerCase() : '',
          gBg: allBg[i-1] ? (allBg[i-1][6]||'').toLowerCase() : '',
          hBg: allBg[i-1] ? (allBg[i-1][7]||'').toLowerCase() : '',
          iBg: allBg[i-1] ? (allBg[i-1][8]||'').toLowerCase() : '',
          lBg: allBg[i-1] ? (allBg[i-1][11]||'').toLowerCase() : '',
          prodNote:   noteRow[8] || '',
          optNote:    noteRow[9] || '',
          sourceNote: noteRow[5] || '',
          orderNote:  noteRow[6] || '',
          recipientNote: noteRow[7] || '',
          deliveryType: String(row[35] || '').trim()
        });
      }
      monthList.reverse();

      const scripts = [];
      const smsSheet = ss.getSheetByName('문자발송');
      if (smsSheet) {
        const smsData = smsSheet.getDataRange().getValues();
        for (let i = 0; i < smsData.length; i++) {
          const colA = String(smsData[i][0]).trim();
          const colB = String(smsData[i][1]).trim();
          if (colA !== '' && colB !== '' && !/^\d{10,}$/.test(colB) &&
              !['주소록 담당자','사업자','주문번호','고객명','자동'].includes(colA)) {
            scripts.push({ idx: scripts.length + 1, title: colA, content: colB });
          }
        }
      }

      return respond({ monthList, sheetName: mName, monthNames, scripts, customers: [] });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // ★★★ all 액션 ★★★
  if (action === 'all') {
    try {
      const smsSheet = ss.getSheetByName('문자발송');
      const customers = [];
      const scripts = [];
      const msgMap = {};

      if (smsSheet) {
        const smsData = smsSheet.getDataRange().getValues();
        for (let i = 0; i < smsData.length; i++) {
          const colA = String(smsData[i][0]).trim();
          const colB = String(smsData[i][1]).trim();
          if (colA !== '' && colB !== '' && !/^\d{10,}$/.test(colB) &&
              !['주소록 담당자','사업자','주문번호','고객명','자동'].includes(colA)) {
            scripts.push({ idx: scripts.length + 1, title: colA, content: colB });
          }
        }
        for (let i = 0; i < smsData.length; i++) {
          const row = smsData[i];
          const colB = String(row[1]).trim();
          const colC = String(row[2]).trim();
          const colG = String(row[6]).trim();
          const colH = String(row[7]).trim();
          if (colB.length >= 10 && /^\d+$/.test(colB) && colC !== '') {
            let phone = '';
            if (colH.startsWith('010')) phone = colH.replace(/-/g, '');
            else if (colG.startsWith('010')) phone = colG.replace(/-/g, '');
            const skipWords = ['주소록 담당자','사업자','주문번호','자동'];
            let msgLines = [], msgStarted = false;
            for (let j = i+1; j < Math.min(i+40, smsData.length); j++) {
              const nB = String(smsData[j][1]).trim();
              const nC = String(smsData[j][2]).trim();
              if (nB.length >= 10 && /^\d+$/.test(nB)) break;
              if (skipWords.includes(nB)) { if (!msgStarted) continue; else break; }
              if (!msgStarted) { if (nB.endsWith('님')) { msgStarted = true; msgLines.push(nB); } continue; }
              if (/^010-?\d{4}-?\d{4}$/.test(nB)) continue;
              if (nB !== '') {
                if (nB.startsWith('[옵션]') && nC !== '') msgLines.push(nB + '\t' + nC);
                else msgLines.push(nB);
              } else if (!nC.startsWith('수량') && msgLines.length > 0 && msgLines[msgLines.length-1] !== '') {
                msgLines.push('');
              }
            }
            while (msgLines.length > 0 && msgLines[msgLines.length-1] === '') msgLines.pop();
            let src = '';
            for (let j = i+1; j < Math.min(i+5, smsData.length); j++) {
              const hv = String(smsData[j][7]).trim();
              if (hv && !hv.startsWith('010')) { src = hv; break; }
            }
            if (phone) {
              customers.push({
                idx: customers.length+1, orderNo: colB, name: colC, phone,
                product: String(row[3]).trim(), option: String(row[4]).trim(),
                source: src, customs: '', message: msgLines.join('\n')
              });
              msgMap[colB] = { message: msgLines.join('\n'), phone, product: String(row[3]).trim(), option: String(row[4]).trim(), source: src };
            }
          }
        }
      }

      const m = getCurrentMonthSheet_(ss);
      const mSheet = m.sheet;

      const mName = mSheet ? mSheet.getName() : m.sheetName;
      const monthList = [];
      if (mSheet) {
        const mData   = mSheet.getDataRange().getValues();
        const lastRow = mData.length;
        const allBg   = lastRow > 1 ? mSheet.getRange(2, 1, lastRow-1, 12).getBackgrounds() : [];
        const allNotesAll = lastRow > 1 ? mSheet.getRange(2, 1, lastRow-1, 12).getNotes() : [];

        const allSheets2  = ss.getSheets();
        const monthSheets2 = allSheets2.filter(s => /^\d{2}년 \d+월$/.test(s.getName()));
        monthSheets2.sort((a, b) => {
          const pa = a.getName().match(/(\d+)년 (\d+)월/);
          const pb = b.getName().match(/(\d+)년 (\d+)월/);
          if (!pa || !pb) return 0;
          return (Number(pb[1])*100+Number(pb[2])) - (Number(pa[1])*100+Number(pa[2]));
        });
        const monthNames = monthSheets2.map(s => s.getName());

        for (let i = 1; i < lastRow; i++) {
          const row = mData[i];
          const orderNo = String(row[6]).trim();
          const name    = String(row[7]).trim();
          if (!orderNo || !name) continue;

          let phone1 = String(row[12]).trim().replace(/-/g, '');
          let phone2 = String(row[13]).trim().replace(/-/g, '');
          if (phone1 && !phone1.startsWith('0') && phone1.length === 10) phone1 = '0' + phone1;
          if (phone2 && !phone2.startsWith('0') && phone2.length === 10) phone2 = '0' + phone2;
          let phone = phone2.startsWith('010') ? phone2 : phone1.startsWith('010') ? phone1 : (phone2 || phone1);
          if (!phone) continue;

          const msgInfo = msgMap[orderNo] || {};
          const noteRow = allNotesAll[i-1] || [];

          monthList.push({
            row: i+1, sheetName: mName, orderNo, name,
            phone, phone1, phone2,
            product: String(row[8]).trim(),  option:   String(row[9]).trim(),
            orderDate: row[4] || '', 
            qty:     String(row[10]).trim(), customs:  String(row[11]).trim(),
            status:  String(row[0]).trim(),  ship:     String(row[5]).trim(),
            source:  String(row[5]).trim(),  memo2:    String(row[15]).trim(),  address: String(row[15]).trim(),  zipCode: String(row[14]).trim(),
            message: msgInfo.message || '',
            payAmt:    row[16] || '',
            settleAmt: row[20] || '',
            taoAmt:    row[22] || '',
            shipFee:   row[26] || '',
            expShip:   row[27] || '',
            addShip:   row[28] || '',
            profitRate: row[32] || '',
            marginRate: row[33] || '',
            eBg: allBg[i-1] ? (allBg[i-1][4]||'').toLowerCase() : '', 
            fBg: allBg[i-1] ? (allBg[i-1][5]||'').toLowerCase() : '',
            gBg: allBg[i-1] ? (allBg[i-1][6]||'').toLowerCase() : '',
            hBg: allBg[i-1] ? (allBg[i-1][7]||'').toLowerCase() : '',
            iBg: allBg[i-1] ? (allBg[i-1][8]||'').toLowerCase() : '',
            lBg: allBg[i-1] ? (allBg[i-1][11]||'').toLowerCase() : '',
            prodNote:   noteRow[8] || '',
            optNote:    noteRow[9] || '',
            sourceNote: noteRow[5] || '',
            orderNote:  noteRow[6] || '',
          recipientNote: noteRow[7] || '',
          deliveryType: String(row[35] || '').trim()
          });
        }
        monthList.reverse();
        return respond({ customers, scripts, monthList, sheetName: mName, monthNames });
      }

      return respond({ customers, scripts, monthList, sheetName: mName, monthNames: [] });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  if (action === 'monthList') {
    try {
      const m = getCurrentMonthSheet_(ss);
      const sheet = m.sheet;
      if (!sheet) return respond({ error: '월별 시트를 찾을 수 없습니다.' });

      const sheetName = sheet.getName();
      const data = sheet.getDataRange().getValues();
      const list = [];
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const orderNo = String(row[6]).trim();
        const name    = String(row[7]).trim();
        if (!orderNo || !name) continue;

        let phone1 = String(row[12]).trim().replace(/-/g, '');
        let phone2 = String(row[13]).trim().replace(/-/g, '');
        if (phone1 && !phone1.startsWith('0') && phone1.length === 10) phone1 = '0' + phone1;
        if (phone2 && !phone2.startsWith('0') && phone2.length === 10) phone2 = '0' + phone2;
        let phone = phone2.startsWith('010') ? phone2 : phone1.startsWith('010') ? phone1 : (phone2 || phone1);
        if (!phone) continue;

        list.push({
          row: i + 1,
          orderNo, name, phone, phone1, phone2,
          product: String(row[8]).trim(),
          option:  String(row[9]).trim(),
          source:  String(row[5]).trim(),
          customs: String(row[11]).trim(),
          baejiNo: String(row[2]).trim(),
          status:  String(row[0]).trim(),
          qty:     String(row[10]).trim(),
          ship:    String(row[5]).trim(),
          delivery:String(row[1]).trim(),
          memo2:   String(row[15]).trim(),
          address: String(row[15]).trim(),
          zipCode: String(row[14]).trim(),
          message: ''
        });
      }
      list.reverse();
      return respond({ list, sheetName, total: list.length });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // 경동 선택
  if (action === 'kdongSelect') {
    const orderNo = e.parameter.orderNo || '';
    const type    = e.parameter.type    || 'none';
    if (!orderNo) return respond({ error: '주문번호 없음' });
    const kdongMap = {
      'k1': { bg: '#FFE0B2', color: '#E65100', text: '경동포함' },
      'k2': { bg: '#FFCCBC', color: '#BF360C', text: '경동관부포함' },
      'k3': { bg: '#E1F5FE', color: '#01579B', text: '개인용달' },
      'k4': { bg: '#B3D9F5', color: '#01579B', text: '개인용달포함' },
      'k5': { bg: '#C5CAE9', color: '#1A237E', text: '개인용달관부포함' },
      'none': { bg: '#FFFFFF', color: '#000000', text: '' }
    };
    const style = kdongMap[type] || kdongMap['none'];
    try {
      const m = getCurrentMonthSheet_(ss);
      const sheet = m.sheet;
      if (!sheet) return respond({ error: '월별 시트를 찾을 수 없습니다.' });
      const found = findRowByOrder_(sheet, orderNo);
      if (found.row < 0) return respond({ error: '주문번호를 찾을 수 없습니다.' });
      const cell = sheet.getRange(found.row, 1);
      cell.setValue(style.text);
      cell.setBackground(style.bg);
      cell.setFontColor(style.color);
      cell.setFontWeight(style.text ? 'bold' : 'normal');
      sheet.getRange(found.row, 7).setNote(style.text || '');
      return respond({ success: true });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // 전체 시트 검색
  // ★ 버전 체크 - 클라이언트가 시작 시 호출, 새 버전이면 자동 새로고침 ★
  if (action === 'version') {
    return respond({
      version: '1.0.4',  // ⚠ 새 버전 배포 시 이 숫자만 올리세요 (예: 1.0.4, 1.1.0)
      releaseDate: '2026-06-10',
      releaseNote: '고객 식별 기준 전화번호→주문번호 변경, 동일 고객 다중주문 행 겹침 버그 수정'
    });
  }

  if (action === 'searchAll') {
    const query = (e.parameter.q || '').toLowerCase().trim();
    if (!query) return respond({ list: [] });
    // ★ 검색 범위 제한 - 최근 N개월만 (기본: 전체)
    const maxMonths = parseInt(e.parameter.months) || 0; // 0 = 전체
    // ★ 특정 시트만 검색 (현재 보는 월 탭 우선)
    const targetSheetName = e.parameter.sheetName || '';
    // ★ ?refresh=1 → 캐시 무시하고 재인덱싱 (강제 새로고침)
    const forceRefresh = e.parameter.refresh === '1';
    if (forceRefresh) clearSearchCache_(ssId);
    try {
      const cache = CacheService.getScriptCache();
      // 캐시 키에 sheetName도 포함 (다른 시트는 별도 캐시)
      const cacheKey = 'search_index_v3_ss' + ssId
        + (maxMonths ? '_m' + maxMonths : '')
        + (targetSheetName ? '_sh' + targetSheetName.replace(/\s/g, '') : '');
      let index = null;
      const cached = cache.get(cacheKey);
      if (cached) { try { index = JSON.parse(cached); } catch(e) { index = null; } }
      if (!index) {
        index = [];
        const allSheets = ss.getSheets();
        let monthSheets = allSheets.filter(s => /^\d{2}년 \d+월$/.test(s.getName()));
        // 특정 시트만 필터
        if (targetSheetName) {
          monthSheets = monthSheets.filter(s => s.getName() === targetSheetName);
        } else {
          monthSheets.sort((a, b) => {
            const pa = a.getName().match(/(\d+)년 (\d+)월/);
            const pb = b.getName().match(/(\d+)년 (\d+)월/);
            if (!pa || !pb) return 0;
            return (Number(pb[1])*100+Number(pb[2])) - (Number(pa[1])*100+Number(pa[2]));
          });
          // 최근 N개월만 검색 (속도 향상)
          if (maxMonths > 0) monthSheets = monthSheets.slice(0, maxMonths);
        }
        for (const mSheet of monthSheets) {
          const mName = mSheet.getName();
          const mData = mSheet.getDataRange().getValues();
          const lastRow = mData.length;
          const allBg = lastRow > 1 ? mSheet.getRange(2, 1, lastRow-1, 12).getBackgrounds() : [];
          const allNotes = lastRow > 1 ? mSheet.getRange(2, 1, lastRow-1, 12).getNotes() : [];
          for (let i = 1; i < lastRow; i++) {

            const row = mData[i];
            const name    = String(row[7]).trim();
            const orderNo = String(row[6]).trim();
            if (!orderNo || !name) continue;
            let phone1 = String(row[12]).trim().replace(/-/g, '');
            let phone2 = String(row[13]).trim().replace(/-/g, '');
            if (phone1 && !phone1.startsWith('0') && phone1.length === 10) phone1 = '0' + phone1;
            if (phone2 && !phone2.startsWith('0') && phone2.length === 10) phone2 = '0' + phone2;
            let phone = phone2.startsWith('010') ? phone2 : phone1.startsWith('010') ? phone1 : (phone2 || phone1);
            if (!phone) phone = orderNo.replace(/\D/g,'').slice(-11);
            index.push({
              row: i+1, sheetName: mName, orderNo, name, phone, phone1, phone2,
              product: String(row[8]).trim(), option:  String(row[9]).trim(),
              orderDate: row[4] || '', 
              qty: String(row[10]).trim(), customs: String(row[11]).trim(),
              baejiNo: String(row[2]).trim(),
              status: String(row[0]).trim(), ship: String(row[5]).trim(),
              source: String(row[5]).trim(), memo2: String(row[15]).trim(),
              address: String(row[15]).trim(),
              zipCode: String(row[14]).trim(),
              message: '',
              payAmt:    row[16] !== '' ? row[16] : '',
              settleAmt: row[20] !== '' ? row[20] : '',
              taoAmt:    row[22] !== '' ? row[22] : '',
              shipFee:   row[26] !== '' ? row[26] : '',
              expShip:   row[27] !== '' ? row[27] : '',
              addShip:   row[28] !== '' ? row[28] : '',
              profitRate: row[32] !== '' ? row[32] : '',
              marginRate: row[33] !== '' ? row[33] : '',
              eBg: allBg[i-1] ? (allBg[i-1][4]||'').toLowerCase() : '', 
              fBg: allBg[i-1] ? (allBg[i-1][5]||'').toLowerCase() : '',
              gBg: allBg[i-1] ? (allBg[i-1][6]||'').toLowerCase() : '',
              hBg: allBg[i-1] ? (allBg[i-1][7]||'').toLowerCase() : '',
               iBg: allBg[i-1] ? (allBg[i-1][8]||'').toLowerCase() : '',
              lBg: allBg[i-1] ? (allBg[i-1][11]||'').toLowerCase() : '',
              prodNote:   (allNotes[i-1] && allNotes[i-1][8])  || '',
              optNote:    (allNotes[i-1] && allNotes[i-1][9])  || '',
              sourceNote: (allNotes[i-1] && allNotes[i-1][5])  || '',
              orderNote:  (allNotes[i-1] && allNotes[i-1][6])  || '',
              recipientNote: (allNotes[i-1] && allNotes[i-1][7]) || '',
              deliveryType: String(row[35] || '').trim()
            });

          }
        }
        try { cache.put(cacheKey, JSON.stringify(index), 21600); } catch(e) {}
      }
      const q = query.replace(/-/g,'');
      const list = index.filter(c =>
        c.name.toLowerCase().includes(query) ||
        c.phone.includes(q) ||
        c.phone1.includes(q) ||
        c.phone2.includes(q) ||
        c.orderNo.includes(query) ||
        c.product.toLowerCase().includes(query) ||
        c.option.toLowerCase().includes(query) ||
        (c.memo2 && c.memo2.toLowerCase().includes(query)) ||
        (c.zipCode && c.zipCode.includes(query)) ||
        (c.baejiNo && c.baejiNo.toLowerCase().includes(query))
      ).slice(0, 100);
      return respond({ list });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // 판매처 목록
  if (action === 'getSourceList') {
    try {
      const feeSheet = ss.getSheetByName('수수료율 정리');
      if (!feeSheet) return respond({ error: '수수료율 정리 시트 없음' });
      const vals = feeSheet.getRange('B5:B11').getValues();
      const sources = vals.map(r => String(r[0]).trim()).filter(v => v);
      return respond({ sources });
    } catch(e) {
      return respond({ error: e.message });
    }
  }

  if (action === 'clearCache') {
    try {
      clearSearchCache_(ssId);
      return respond({ success: true });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // 주문취소
  if (action === 'cancelOrder') {
    const orderNo = e.parameter.orderNo || '';
    const cancel  = e.parameter.cancel === '1';
    if (!orderNo) return respond({ error: '주문번호 없음' });
    try {
      const allSheets = ss.getSheets();
      const monthSheets = allSheets.filter(s => /^\d{2}년 \d+월$/.test(s.getName()));
      let found = false;
      for (const mSheet of monthSheets) {
        const data = mSheet.getDataRange().getValues();
        for (let i = 0; i < data.length; i++) {
          if (String(data[i][6]).trim() !== orderNo) continue;
          found = true;
          if (cancel) {
            mSheet.getRange(i+1, 1, 1, 17).setBackground('#ea9999');
            mSheet.getRange(i+1, 17).clearContent();
            mSheet.getRange(i+1, 21).clearContent();
            mSheet.getRange(i+1, 23).clearContent();
            mSheet.getRange(i+1, 27).clearContent();
            mSheet.getRange(i+1, 28).clearContent();
            mSheet.getRange(i+1, 29).clearContent();
          } else {
            mSheet.getRange(i+1, 1, 1, 17).setBackground('#ffffff');
          }
        }
      }
      clearSearchCache_(ssId);
      return respond({ success: true, found });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // 금액 전체 한번에 저장
  if (action === 'saveAllAmt') {
    const orderNo   = e.parameter.orderNo || '';
    const sheetName = e.parameter.sheetName || '';
    const vals = {
      17: e.parameter.payAmt,
      21: e.parameter.settleAmt,
      23: e.parameter.taoAmt,
      27: e.parameter.shipFee,
      28: e.parameter.expShip,
      29: e.parameter.addShip
    };
    if (!orderNo) return respond({ error: '주문번호 없음' });
    try {
      const allSheets = ss.getSheets();
      let monthSheets = allSheets.filter(s => /^\d{2}년 \d+월$/.test(s.getName()));
      if (sheetName) monthSheets.sort((a,b) => a.getName()===sheetName?-1:b.getName()===sheetName?1:0);
      for (const mSheet of monthSheets) {
        const data = mSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][6]).trim() === String(orderNo).trim()) {
            Object.entries(vals).forEach(([col, val]) => {
              if (val === undefined || val === null) return;
              const n = parseFloat(val);
              mSheet.getRange(i+1, parseInt(col)).setValue(isNaN(n) ? val : n);
            });
            SpreadsheetApp.flush();
            Utilities.sleep(200);
            const updRow = mSheet.getRange(i+1, 1, 1, 35).getValues()[0];
            return respond({
              success: true, sheet: mSheet.getName(), row: i+1,
              profitRate: updRow[32] || '', marginRate: updRow[33] || ''
            });
          }
        }
      }
      return respond({ error: '주문번호 없음: ' + orderNo });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  if (action === 'saveAmtField') {
    const orderNo  = e.parameter.orderNo || '';
    const col      = parseInt(e.parameter.col) || 0;
    const val      = e.parameter.val || '';
    const sheetName = e.parameter.sheetName || '';
    if (!orderNo || !col) return respond({ error: '파라미터 오류' });
    try {
      const allSheets = ss.getSheets();
      let monthSheets = allSheets.filter(s => /^\d{2}년 \d+월$/.test(s.getName()));
      if (sheetName) monthSheets.sort((a, b) => a.getName() === sheetName ? -1 : b.getName() === sheetName ? 1 : 0);
      for (const mSheet of monthSheets) {
        const data = mSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][6]).trim() === String(orderNo).trim()) {
            const numVal = parseFloat(val);
            mSheet.getRange(i+1, col).setValue(isNaN(numVal) ? val : numVal);
            return respond({ success: true, sheet: mSheet.getName(), row: i+1 });
          }
        }
      }
      return respond({ error: '주문번호를 찾을 수 없습니다: ' + orderNo });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // 엑셀 데이터 행 추가
    if (action === 'appendRow') {
    const sheetName = e.parameter.sheetName || '';
    if (!sheetName) return respond({ error: '시트명 없음' });
    const mSheet = ss.getSheetByName(sheetName);
    if (!mSheet) return respond({ error: '시트를 찾을 수 없습니다: ' + sheetName });
    try {
      const aVals = mSheet.getRange('A:A').getValues();
      let lastRow = 1;
      for (let i = aVals.length - 1; i >= 0; i--) {
        if (aVals[i][0] !== '' && aVals[i][0] !== null) { lastRow = i + 2; break; }
      }
      const cols = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P'];

      // ★★★ 주문번호(G열) 중복 검사 후 자동 제미사 부여 ★★★
      const inputOrderNo = String(e.parameter['G'] || '').trim();
      let finalOrderNo = inputOrderNo;
      if (inputOrderNo) {
        // 현재 시트의 모든 G열 값 가져오기
        const gVals = mSheet.getRange('G:G').getValues();
        const existing = [];
        for (let gi = 0; gi < gVals.length; gi++) {
          const v = String(gVals[gi][0] || '').trim();
          if (v) existing.push(v);
        }
        // 기본번호 자체 또는 -숫자 접미사가 이미 있는지 검사
        // 예: "2026050860908851" 이 있으면 "-1" 부여
        //     "2026050860908851-1" 도 있으면 "-2" 부여
        const baseNo = inputOrderNo.replace(/-\d+$/, '');  // 혹시 입력값에 이미 -숫자 있으면 제거
        const hasBase = existing.indexOf(baseNo) >= 0;
        if (hasBase) {
          // 기존 -숫자 최대값 찾기
          let maxSuffix = 0;
          const re = new RegExp('^' + baseNo.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '-(\\d+)$');
          for (let ei = 0; ei < existing.length; ei++) {
            const m = existing[ei].match(re);
            if (m) {
              const n = parseInt(m[1], 10);
              if (n > maxSuffix) maxSuffix = n;
            }
          }
          // 기본번호만 있고 -숫자 없는 경우
          if (maxSuffix === 0) {
            // 기존 기본번호 행을 찾아서 -1로 바꾸고, 새 행은 -2
            for (let ei = 0; ei < gVals.length; ei++) {
              if (String(gVals[ei][0] || '').trim() === baseNo) {
                mSheet.getRange(ei + 1, 7).setValue(baseNo + '-1');
                break;
              }
            }
            finalOrderNo = baseNo + '-2';
          } else {
            finalOrderNo = baseNo + '-' + (maxSuffix + 1);
          }
        }
      }

      for (let i = 0; i < cols.length; i++) {
        let val = e.parameter[cols[i]] || '';
        if (cols[i] === 'G') val = finalOrderNo;  // ★ G열은 최종 주문번호 사용
        if (val) mSheet.getRange(lastRow, i+1).setValue(val);
      }
      const qVal = e.parameter['Q'] || '';
      if (qVal) {
        const qNum = parseFloat(qVal.replace(/,/g,''));
        mSheet.getRange(lastRow, 17).setValue(isNaN(qNum) ? qVal : qNum);
      }
      const uVal = e.parameter['U'] || '';
      if (uVal) {
        const uNum = parseFloat(uVal.replace(/,/g,''));
        mSheet.getRange(lastRow, 21).setValue(isNaN(uNum) ? uVal : uNum);
      }
      // 캐시 무효화
      clearSearchCache_(ssId);
      return respond({ success: true, row: lastRow, orderNo: finalOrderNo });
    } catch(err) {
      return respond({ error: err.message });
    }
  }
  // ═══════════════════════════════════════════════════════════════
// ★ 기존 시트의 중복 주문번호 일괄 정리 (한 번만 실행하면 됨!)
// ═══════════════════════════════════════════════════════════════
// 같은 주문번호 2개 이상 → 위에서 아래로 -1, -2, -3... 자동 부여
// 이미 -숫자 접미사 있는 행은 건드리지 않음
function deduplicateOrderNos_(ssId) {
  var SS_MAP = {
    '1': '10hfa94OWjpyvF4eK5w2FYhqsIRbfsn8pPlISoj9N9Jg',
    '2': '1u2kIj8-c6YVtv1XOg40-Im1dL09sGiIxehatDKD8Cu0',
    '3': '137pCqtBMsOgX752pVoglwBk3NyK2aXtAy1R40lLJFiM',
    '4': '1AVe0jVZXFLloVEbJHh65OSru8ZbWnXwplHxAmbaXkGE',
    '5': '1IZEnu3o7WcpRUXJP3AcxPQyNOsUj1J2bTSjvpX14oGo',
    '6': '1FCvAfy4AMClQ-0pdJ_AjdwAOqh05Ldh80fTSXknI6Nc'
  };
  var ss;
  if (ssId === '1') {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } else {
    ss = SpreadsheetApp.openById(SS_MAP[ssId]);
  }
  var monthSheets = ss.getSheets().filter(function(s) {
    return /^\d{2}년 \d+월$/.test(s.getName());
  });
  var totalChanged = 0;
  var detail = [];
  monthSheets.forEach(function(sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    var gRange = sheet.getRange(2, 7, lastRow - 1, 1);
    var gVals = gRange.getValues();
    // 1단계: 기본번호(접미사 없음)와 중복카운트 수집
    var baseCount = {};
    var baseFirstRowIdx = {};  // baseNo → 첫 등장 행 인덱스
    var hasSuffix = {};        // baseNo → 이미 -숫자 행이 있는지
    var maxSuffix = {};        // baseNo → 현재 최대 접미사
    for (var i = 0; i < gVals.length; i++) {
      var v = String(gVals[i][0] || '').trim();
      if (!v) continue;
      var sm = v.match(/^(.+)-(\d+)$/);
      if (sm) {
        var base = sm[1];
        var num = parseInt(sm[2], 10);
        hasSuffix[base] = true;
        if (!maxSuffix[base] || num > maxSuffix[base]) maxSuffix[base] = num;
      } else {
        // 접미사 없는 기본번호
        if (!(v in baseCount)) {
          baseCount[v] = 1;
          baseFirstRowIdx[v] = i;
        } else {
          baseCount[v]++;
        }
      }
    }
    // 2단계: 중복(2개 이상)인 기본번호 처리
    var changed = false;
    for (var i = 0; i < gVals.length; i++) {
      var v = String(gVals[i][0] || '').trim();
      if (!v) continue;
      if (/-\d+$/.test(v)) continue;  // 이미 접미사 있으면 스킵
      if (baseCount[v] < 2) continue;  // 중복 아니면 스킵
      // 이 행에 새 접미사 부여
      var base = v;
      if (!(base in maxSuffix)) maxSuffix[base] = 0;
      maxSuffix[base]++;
      var newVal = base + '-' + maxSuffix[base];
      gVals[i][0] = newVal;
      changed = true;
      totalChanged++;
      detail.push(sheet.getName() + ' ' + (i+2) + '행: ' + base + ' → ' + newVal);
    }
    if (changed) gRange.setValues(gVals);
  });
  clearSearchCache_(ssId);
  return { total: totalChanged, detail: detail };
}

// ★ 6개 사업자 모두 일괄 정리 (한 번만 실행)
function deduplicateAllOrderNos() {
  var grandTotal = 0;
  for (var i = 1; i <= 6; i++) {
    try {
      var r = deduplicateOrderNos_(String(i));
      grandTotal += r.total;
      Logger.log('═══════════════════');
      Logger.log('ssId=' + i + ' → 변경된 행: ' + r.total + '개');
      r.detail.slice(0, 20).forEach(function(d) { Logger.log('  ' + d); });
      if (r.detail.length > 20) Logger.log('  ... 외 ' + (r.detail.length - 20) + '개');
    } catch (err) {
      Logger.log('ssId=' + i + ' 오류: ' + err.message);
    }
    Utilities.sleep(500);
  }
  Logger.log('═══════════════════');
  Logger.log('🎉 전체 정리 완료: ' + grandTotal + '개 행 변경');
}

// ★ 한 사업자만 (테스트용)
function deduplicateOrderNos_test1() {
  var r = deduplicateOrderNos_('1');
  Logger.log('ssId=1 → 변경된 행: ' + r.total + '개');
  r.detail.forEach(function(d) { Logger.log('  ' + d); });
}




  // ★★★ 고객 정보 수정 ★★★
  if (action === 'updateInfo') {
    const orderNo   = e.parameter.orderNo || '';
    const sheetName = e.parameter.sheetName || '';
    if (!orderNo) return respond({ error: '주문번호 없음' });
    try {
      const allSheets = ss.getSheets();
      // ★ 2자리(26년) + 4자리(2026년) 년도 모두 매칭
      let monthSheets = allSheets.filter(s => /^\d{2,4}년\s*\d+월$/.test(s.getName()));
      if (sheetName) monthSheets.sort((a,b) => a.getName()===sheetName?-1:1);
      for (const mSheet of monthSheets) {
        const data = mSheet.getDataRange().getValues();
        // ★ G열 display values로 큰 숫자 정밀도 손실 방지
        const lastRowG = mSheet.getLastRow();
        const displayG = lastRowG >= 1 ? mSheet.getRange(1, 7, lastRowG, 1).getDisplayValues() : [];
        const searchOrder = String(orderNo).trim();
        for (let i = 1; i < data.length; i++) {
          const cellVal = String(data[i][6]).trim();
          const displayVal = displayG[i] ? String(displayG[i][0]).trim() : '';
          if (cellVal !== searchOrder && displayVal !== searchOrder) continue;
          if (e.parameter.name    !== undefined) mSheet.getRange(i+1, 8).setValue(e.parameter.name);
          // ★ 휴대폰/통관/우편번호는 텍스트 형식 + 앞 0 유지
          const setAsText = (row, col, val) => {
            const cell = mSheet.getRange(row, col);
            cell.setNumberFormat('@');  // 텍스트 형식 강제
            cell.setValue(String(val));
          };
          if (e.parameter.phone1  !== undefined) setAsText(i+1, 13, e.parameter.phone1);
          if (e.parameter.phone2  !== undefined) setAsText(i+1, 14, e.parameter.phone2);
          if (e.parameter.customs !== undefined) setAsText(i+1, 12, e.parameter.customs);
          if (e.parameter.zipCode !== undefined) setAsText(i+1, 15, e.parameter.zipCode);
          if (e.parameter.addr    !== undefined) mSheet.getRange(i+1, 16).setValue(e.parameter.addr);
          if (e.parameter.source  !== undefined) mSheet.getRange(i+1, 6).setValue(e.parameter.source);
          const toNum = v => { const n=parseFloat(String(v).replace(/,/g,'')); return isNaN(n)?'':n; };
          if (e.parameter.taoAmt  !== undefined) mSheet.getRange(i+1, 23).setValue(toNum(e.parameter.taoAmt));
          if (e.parameter.shipFee !== undefined) mSheet.getRange(i+1, 27).setValue(toNum(e.parameter.shipFee));
          if (e.parameter.expShip !== undefined) mSheet.getRange(i+1, 28).setValue(toNum(e.parameter.expShip));
          if (e.parameter.addShip !== undefined) mSheet.getRange(i+1, 29).setValue(toNum(e.parameter.addShip));
          if (e.parameter.qty     !== undefined) mSheet.getRange(i+1, 11).setValue(toNum(e.parameter.qty));

          SpreadsheetApp.flush();
          Utilities.sleep(200);
          const updRow = mSheet.getRange(i+1, 1, 1, 35).getValues()[0];
          return respond({
            success: true,
            sheet: mSheet.getName(),
            row: i+1,
            profitRate: updRow[32] !== '' ? updRow[32] : '',
            marginRate: updRow[33] !== '' ? updRow[33] : ''
          });
        }
      }
      return respond({ error: '주문번호를 찾을 수 없습니다.' });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // ★★★ getMonthData ★★★
  if (action === 'getMonthData') {
    const targetSheet = e.parameter.sheetName || '';
    if (!targetSheet) return respond({ error: '시트명 없음' });
    const mSheet = ss.getSheetByName(targetSheet);
    if (!mSheet) return respond({ error: '시트를 찾을 수 없습니다.' });

    const msgMap = {};
    const smsSheet = ss.getSheetByName('문자발송');
    if (smsSheet) {
      const smsData = smsSheet.getDataRange().getValues();
      for (let i = 0; i < smsData.length; i++) {
        const row = smsData[i];
        const colB = String(row[1]).trim();
        const colC = String(row[2]).trim();
        const colG = String(row[6]).trim();
        const colH = String(row[7]).trim();
        if (colB.length >= 10 && /^\d+$/.test(colB) && colC !== '') {
          let phone = '';
          if (colH.startsWith('010')) phone = colH.replace(/-/g, '');
          else if (colG.startsWith('010')) phone = colG.replace(/-/g, '');
          const skipWords = ['주소록 담당자','사업자','주문번호','자동'];
          let msgLines = [], msgStarted = false;
          for (let j = i+1; j < Math.min(i+40, smsData.length); j++) {
            const nB = String(smsData[j][1]).trim();
            const nC = String(smsData[j][2]).trim();
            if (nB.length >= 10 && /^\d+$/.test(nB)) break;
            if (skipWords.includes(nB)) { if (!msgStarted) continue; else break; }
            if (!msgStarted) { if (nB.endsWith('님')) { msgStarted = true; msgLines.push(nB); } continue; }
            if (/^010-?\d{4}-?\d{4}$/.test(nB)) continue;
            if (nB !== '') {
              if (nB.startsWith('[옵션]') && nC !== '') msgLines.push(nB + '\t' + nC);
              else msgLines.push(nB);
            } else if (!nC.startsWith('수량') && msgLines.length > 0 && msgLines[msgLines.length-1] !== '') {
              msgLines.push('');
            }
          }
          while (msgLines.length > 0 && msgLines[msgLines.length-1] === '') msgLines.pop();
          let source = '';
          for (let j = i+1; j < Math.min(i+5, smsData.length); j++) {
            const hv = String(smsData[j][7]).trim();
            if (hv && !hv.startsWith('010')) { source = hv; break; }
          }
          msgMap[colB] = {
            message: msgLines.join('\n'),
            product: String(row[3]).trim(),
            option:  String(row[4]).trim(),
            source:  source
          };
        }
      }
    }

    const list = [];
    const mData   = mSheet.getDataRange().getValues();
    const lastRow = mData.length;
    const allBg   = lastRow > 1 ? mSheet.getRange(2, 1, lastRow-1, 12).getBackgrounds() : [];
    const allNotesMD = lastRow > 1 ? mSheet.getRange(2, 1, lastRow-1, 12).getNotes() : [];

    for (let i = 1; i < lastRow; i++) {
      const row = mData[i];
      const orderNo = String(row[6]).trim();
      const name    = String(row[7]).trim();
      if (!orderNo || !name) continue;

      let phone1 = String(row[12]).trim().replace(/-/g, '');
      let phone2 = String(row[13]).trim().replace(/-/g, '');
      if (phone1 && !phone1.startsWith('0') && phone1.length === 10) phone1 = '0' + phone1;
      if (phone2 && !phone2.startsWith('0') && phone2.length === 10) phone2 = '0' + phone2;
      let phone = phone2.startsWith('010') ? phone2 : phone1.startsWith('010') ? phone1 : (phone2 || phone1);
      if (!phone) phone = orderNo.replace(/\D/g,'').slice(-11);

      const msgInfo = msgMap[orderNo] || {};
      const noteRow = allNotesMD[i-1] || [];

      list.push({
        row: i+1, sheetName: targetSheet, orderNo, name,
        phone, phone1, phone2,
        product:  msgInfo.product || String(row[8]).trim(),
        option:   msgInfo.option  || String(row[9]).trim(),
        orderDate: row[4] || '',
        qty:      String(row[10]).trim(),
        customs:  String(row[11]).trim(),
        baejiNo:  String(row[2]).trim(),
        status:   String(row[0]).trim(),
        ship:     String(row[5]).trim(),
        source:   String(row[5]).trim() || msgInfo.source || '',
        memo2:    String(row[15]).trim(),
        zipCode:  String(row[14]).trim(),
        deliveryType: String(row[35] || '').trim(),
        message:  msgInfo.message || buildMsg_(row, ssId),
        payAmt:    row[16] !== '' ? row[16] : '',
        settleAmt: row[20] !== '' ? row[20] : '',
        taoAmt:    row[22] !== '' ? row[22] : '',
        shipFee:   row[26] !== '' ? row[26] : '',
        expShip:   row[27] !== '' ? row[27] : '',
        addShip:   row[28] !== '' ? row[28] : '',
        profitRate: row[32] !== '' ? row[32] : '',
        marginRate: row[33] !== '' ? row[33] : '',
        fBg: allBg[i-1] ? (allBg[i-1][5]||'').toLowerCase() : '',
        gBg: allBg[i-1] ? (allBg[i-1][6]||'').toLowerCase() : '',
        hBg: allBg[i-1] ? (allBg[i-1][7]||'').toLowerCase() : '',
        iBg: allBg[i-1] ? (allBg[i-1][8]||'').toLowerCase() : '',
        lBg: allBg[i-1] ? (allBg[i-1][11]||'').toLowerCase() : '',
        prodNote:   noteRow[8] || '',
        optNote:    noteRow[9] || '',
        sourceNote: noteRow[5] || '',
        orderNote:  noteRow[6] || '',
        recipientNote: noteRow[7] || '',
          deliveryType: String(row[35] || '').trim()
      });
    }
    return respond({ list, sheetName: targetSheet });
  }

  // 배대지 운송방법
  if (action === 'shipMethod') {
    const orderNo = e.parameter.orderNo || '';
    const type    = e.parameter.type    || 'none';
    if (!orderNo) return respond({ error: '주문번호 없음' });
    const colorMap = {
      'fast'     : '#FF9800',
      'joeun-air': '#8BC34A',
      'joeun-sea': '#9C27B0',
      'lcl-out'  : '#A04545',
      'none'     : '#FFFFFF'
    };
    const color = colorMap[type] || '#FFFFFF';
    try {
      const m = getCurrentMonthSheet_(ss);
      const sheet = m.sheet;
      if (!sheet) return respond({ error: '월별 시트를 찾을 수 없습니다.' });
      const found = findRowByOrder_(sheet, orderNo);
      if (found.row < 0) return respond({ error: '주문번호를 찾을 수 없습니다.' });
      sheet.getRange(found.row, 6).setBackground(color);
      return respond({ success: true });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // ★ 발송지연 / 직접전달 처리 ★
  // G열(주문번호) → 주황색 #FF9800 + 노트 = '발송지연' 또는 '직접전달'
  // type=='' (빈값) → 흰색 + 노트 제거
  if (action === 'delayOrder') {
    const orderNo = e.parameter.orderNo || '';
    const type    = (e.parameter.type || '').trim();  // '발송지연' / '직접전달' / ''
    if (!orderNo) return respond({ error: '주문번호 없음' });
    if (type && type !== '발송지연' && type !== '직접전달') {
      return respond({ error: 'type은 발송지연/직접전달/빈값만 허용' });
    }
    try {
      const targetSheet = e.parameter.sheetName || '';
      let sheet = targetSheet ? ss.getSheetByName(targetSheet) : null;
      if (!sheet) { const m = getCurrentMonthSheet_(ss); sheet = m.sheet; }
      if (!sheet) return respond({ error: '월별 시트를 찾을 수 없습니다.' });
      const found = findRowByOrder_(sheet, orderNo);
      if (found.row < 0) return respond({ error: '주문번호를 찾을 수 없습니다.' });
      const gCell = sheet.getRange(found.row, 7);  // G열 = 7번째
      if (type) {
        gCell.setBackground('#FF9800');  // 주황
        gCell.setNote(type);
      } else {
        gCell.setBackground('#FFFFFF');
        gCell.setNote('');
      }
      clearSearchCache_(ssId);
      return respond({ success: true });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // ★ 발송처리완료 / 해제 ★
  // H열(수령자) → 주황색 #FF9800 + 노트 = '발송처리완료'
  // cancel=1 → 흰색 + 노트 제거
  if (action === 'shipDone') {
    const orderNo = e.parameter.orderNo || '';
    const cancel  = e.parameter.cancel === '1';
    if (!orderNo) return respond({ error: '주문번호 없음' });
    try {
      const targetSheet = e.parameter.sheetName || '';
      let sheet = targetSheet ? ss.getSheetByName(targetSheet) : null;
      if (!sheet) { const m = getCurrentMonthSheet_(ss); sheet = m.sheet; }
      if (!sheet) return respond({ error: '월별 시트를 찾을 수 없습니다.' });
      const found = findRowByOrder_(sheet, orderNo);
      if (found.row < 0) return respond({ error: '주문번호를 찾을 수 없습니다.' });
      const hCell = sheet.getRange(found.row, 8);
      if (cancel) {
        hCell.setBackground('#FFFFFF');
        hCell.setNote('');
      } else {
        hCell.setBackground('#FF9800');
        hCell.setNote('발송처리완료');
      }
      clearSearchCache_(ssId);
      return respond({ success: true });
    } catch(err) {
      return respond({ error: err.message });
    }
  }



  // 배대지신청서
  if (action === 'baejiRequest') {
    const orderNo = e.parameter.orderNo || '';
    const cancel  = e.parameter.cancel === '1';
    const baejiNo = e.parameter.baejiNo || '';
    if (!orderNo) return respond({ error: '주문번호 없음' });
    try {
      const targetSheet = e.parameter.sheetName || '';
      let sheet = targetSheet ? ss.getSheetByName(targetSheet) : null;
      if (!sheet) {
        const m = getCurrentMonthSheet_(ss);
        sheet = m.sheet;
      }
      if (!sheet) return respond({ error: '월별 시트를 찾을 수 없습니다.' });
      const found = findRowByOrder_(sheet, orderNo);
      if (found.row < 0) return respond({ error: '주문번호를 찾을 수 없습니다.' });
      sheet.getRange(found.row, 8).setBackground(cancel ? '#FFFFFF' : '#FFFF00');
      const cCell = sheet.getRange(found.row, 3);
      cCell.setValue(cancel ? '' : baejiNo);
      cCell.setBackground(cancel ? '#FFFFFF' : '#FFFF00');
      // 캐시 무효화
      clearSearchCache_(ssId);
      return respond({ success: true });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // ★ 배대지신청 (G열 노란색 칠 + C열 SH번호 저장) ★
  // 기존 baejiRequest와 다른 점: G열(주문번호) 칠, H열은 건드리지 않음
  if (action === 'baejiRequestG') {
    const orderNo = e.parameter.orderNo || '';
    const cancel  = e.parameter.cancel === '1';
    const baejiNo = e.parameter.baejiNo || '';
    if (!orderNo) return respond({ error: '주문번호 없음' });
    try {
      const targetSheet = e.parameter.sheetName || '';
      let sheet = targetSheet ? ss.getSheetByName(targetSheet) : null;
      if (!sheet) {
        const m = getCurrentMonthSheet_(ss);
        sheet = m.sheet;
      }
      if (!sheet) return respond({ error: '월별 시트를 찾을 수 없습니다.' });
      const found = findRowByOrder_(sheet, orderNo);
      if (found.row < 0) return respond({ error: '주문번호를 찾을 수 없습니다.' });
      // G열(7번째) 노란색/흰색 토글
      sheet.getRange(found.row, 7).setBackground(cancel ? '#FFFFFF' : '#FFFF00');
      // C열(3번째)에 신청서번호 저장 (취소 시 비움)
      sheet.getRange(found.row, 3).setValue(cancel ? '' : baejiNo);
      // 캐시 무효화
      clearSearchCache_(ssId);
      return respond({ success: true, baejiNo: baejiNo });
    } catch(err) {
      return respond({ error: err.message });
    }
  }



  // ★ 배대지 신청서번호만 단순 수정 (색깔 변경 없음, 메모 없음)
  if (action === 'updateBaejiNo') {
    const orderNo = e.parameter.orderNo || '';
    const baejiNo = e.parameter.baejiNo || '';
    if (!orderNo) return respond({ error: '주문번호 없음' });
    try {
      const targetSheet = e.parameter.sheetName || '';
      let sheet = targetSheet ? ss.getSheetByName(targetSheet) : null;
      if (!sheet) {
        const m = getCurrentMonthSheet_(ss);
        sheet = m.sheet;
      }
      if (!sheet) return respond({ error: '월별 시트를 찾을 수 없습니다.' });
      const found = findRowByOrder_(sheet, orderNo);
      if (found.row < 0) return respond({ error: '주문번호를 찾을 수 없습니다.' });
      sheet.getRange(found.row, 3).setValue(baejiNo);
      // 캐시 무효화 (검색 인덱스에서 baejiNo 갱신 위해)
      clearSearchCache_(ssId);
      return respond({ success: true, baejiNo: baejiNo });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

if (action === 'deliverySelect') {
    const orderNo = e.parameter.orderNo || '';
    const type    = e.parameter.type    || '';
    if (!orderNo) return respond({ error: '주문번호 없음' });
    try {
      // ★ 모든 월 시트에서 검색
      const allSheets = ss.getSheets();
      const monthSheets = allSheets.filter(s => /^\d{2,4}년\s*\d+월$/.test(s.getName()));
      let sheet = null;
      let found = null;
      for (const ms of monthSheets) {
        const f = findRowByOrder_(ms, orderNo);
        if (f.row > 0) { sheet = ms; found = f; break; }
      }
      if (!sheet || !found) return respond({ error: '주문번호를 찾을 수 없습니다.' });
      
      // 1. AJ열(36)에 택배사 저장
      sheet.getRange(found.row, 36).setValue(type);
      
      // 2. A~F열(1~6) + G열(7) 색칠
      const bgColor = type ? '#81D4FA' : '#FFFFFF';
      sheet.getRange(found.row, 1, 1, 6).setBackground(bgColor);
      sheet.getRange(found.row, 7).setBackground(type ? '#81D4FA' : '#FF9800');
      
      // 3. ★★★ 항상 패스트로에서 운송료 가져오기 (택배사 선택 시) ★★★
      let shipFee = 0, prepaid = 0, fastroFound = false;
      if (type) {
        const cellSh = String(sheet.getRange(found.row, 3).getValue()).trim();
        if (cellSh && /^SH/i.test(cellSh)) {
          const mbId = getFastroMbId_(ssId);
          if (mbId) {
            const today = new Date();
            const past = new Date(today.getTime() - 60 * 24 * 60 * 60 * 1000);
            const fmt = function(d) {
              return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
            };
            const r = fastroFetch_(ssId, mbId, fmt(past), fmt(today), 200);
            if (r.ok) {
              const records = parseFastroHtml_(r.html);
              const shFull = cellSh;
              const shNum = cellSh.replace(/^SH/i, '');
              for (let i = 0; i < records.length; i++) {
                const rec = records[i];
                const recShNum = (rec.sh || '').replace(/^SH/i, '');
                if (rec.sh === shFull || recShNum === shNum) {
                  shipFee = rec.shipFee || 0;
                  prepaid = rec.prepaid || 0;
                  fastroFound = true;
                  if (shipFee > 0) {
                    sheet.getRange(found.row, 27).setValue(shipFee);
                  }
                  sheet.getRange(found.row, 28).setValue(prepaid > 0 ? prepaid : '');
                  break;
                }
              }
            }
          }
        }
      }
      
      clearSearchCache_(ssId);
      
      return respond({
        success: true,
        type: type,
        fastroFound: fastroFound,
        shipFee: shipFee,
        prepaid: prepaid
      });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // ★ 주문 토글 ★
  if (action === 'orderToggle') {
    const orderNo = e.parameter.orderNo || '';
    const cancel  = e.parameter.cancel === '1';
    if (!orderNo) return respond({ error: '주문번호 없음' });
    try {
      // ★ 모든 월 시트에서 검색
      const allSheets = ss.getSheets();
      const monthSheets = allSheets.filter(s => /^\d{2,4}년\s*\d+월$/.test(s.getName()));
      let sheet = null;
      let found = null;
      for (const ms of monthSheets) {
        const f = findRowByOrder_(ms, orderNo);
        if (f.row > 0) { sheet = ms; found = f; break; }
      }
      if (!sheet || !found) return respond({ error: '주문번호를 찾을 수 없습니다.' });
      sheet.getRange(found.row, 9).setBackground(cancel ? '#FFFFFF' : '#FFFF00');
      sheet.getRange(found.row, 12).setBackground('#FFFFFF');
      return respond({ success: true });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // 문자발송 - G열 진한노란
  if (action === 'markSent') {
    const orderNo   = e.parameter.orderNo   || '';
    const sheetName = e.parameter.sheetName || '';
    if (!orderNo) return respond({ error: '주문번호 없음' });
    try {
      let sheet = sheetName ? ss.getSheetByName(sheetName) : null;
      if (!sheet) {
        const m = getCurrentMonthSheet_(ss);
        sheet = m.sheet;
      }
      if (!sheet) return respond({ error: '시트 없음' });
      const found = findRowByOrder_(sheet, orderNo);
      if (found.row < 0) return respond({ error: '주문번호 없음' });
      sheet.getRange(found.row, 7).setBackground('#FFD600');
      return respond({ success: true });
    } catch(err) { return respond({ error: err.message }); }
  }

  if (action === 'requestOrder') {
    const orderNo = e.parameter.orderNo || '';
    if (!orderNo) return respond({ error: '주문번호 없음' });
    try {
      const cancel = e.parameter.cancel === '1';
      // ★ 모든 월 시트에서 검색 (5월 시트의 주문을 6월에 처리 가능)
      const allSheets = ss.getSheets();
      const monthSheets = allSheets.filter(s => /^\d{2,4}년\s*\d+월$/.test(s.getName()));
      // 우선순위: 요청 시 sheetName이 있으면 그 시트 먼저
      const reqSheetName = e.parameter.sheetName || '';
      if (reqSheetName) {
        monthSheets.sort((a, b) => a.getName() === reqSheetName ? -1 : 1);
      }
      for (const sheet of monthSheets) {
        const found = findRowByOrder_(sheet, orderNo);
        if (found.row > 0) {
          sheet.getRange(found.row, 12).setBackground(cancel ? '#FFFFFF' : '#FF6D00');
          SpreadsheetApp.flush();
          return respond({ success: true, sheet: sheet.getName() });
        }
      }
      return respond({ error: '주문번호를 찾을 수 없습니다.', orderNo: orderNo });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // 상품명/옵션명/수량 저장
  if (action === 'saveDetail') {
    const orderNo = e.parameter.orderNo || '';
    if (!orderNo) return respond({ error: '주문번호 없음' });
    try {
      const targetSheet = e.parameter.sheetName || '';
      let sheet = targetSheet ? ss.getSheetByName(targetSheet) : null;
      if (!sheet) {
        const m = getCurrentMonthSheet_(ss);
        sheet = m.sheet;
      }
      if (!sheet) return respond({ error: '월별 시트를 찾을 수 없습니다.' });
      const data = sheet.getDataRange().getValues();
      let updated = 0;
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][6]).trim() !== orderNo) continue;
        if (e.parameter.product  !== undefined) sheet.getRange(i+1, 9).setValue(e.parameter.product);
        if (e.parameter.option   !== undefined) sheet.getRange(i+1, 10).setValue(e.parameter.option);
        if (e.parameter.qty      !== undefined && e.parameter.qty !== '') sheet.getRange(i+1, 11).setValue(Number(e.parameter.qty));
        if (e.parameter.prodNote !== undefined) sheet.getRange(i+1, 9).setNote(e.parameter.prodNote);
        if (e.parameter.optNote  !== undefined) sheet.getRange(i+1, 10).setNote(e.parameter.optNote);
        updated++;
      }
      return respond({ success: true, updated });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // 주문번호 행 노란색 색칠
  if (action === 'highlightRow') {
    const orderNo = e.parameter.orderNo || '';
    if (!orderNo) return respond({ error: '주문번호 없음' });
    try {
      const m = getCurrentMonthSheet_(ss);
      const sheet = m.sheet;
      if (!sheet) return respond({ error: '월별 시트를 찾을 수 없습니다.' });
      const found = findRowByOrder_(sheet, orderNo);
      if (found.row < 0) return respond({ error: '주문번호를 찾을 수 없습니다: ' + orderNo });
      sheet.getRange(found.row, 7).setBackground('#FFFF00');
      return respond({ success: true });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // ★★★ 통계 데이터 ★★★
  if (action === 'stats') {
    try {
      const allSheets = ss.getSheets();
      const monthSheets = allSheets.filter(s => /^\d{2}년 \d+월$/.test(s.getName()));
      monthSheets.sort((a, b) => {
        const pa = a.getName().match(/(\d+)년 (\d+)월/);
        const pb = b.getName().match(/(\d+)년 (\d+)월/);
        if (!pa || !pb) return 0;
        return (Number(pa[1])*100+Number(pa[2])) - (Number(pb[1])*100+Number(pb[2]));
      });

      const monthData = [];
      for (const mSheet of monthSheets) {
        const mName = mSheet.getName();
        const data = mSheet.getDataRange().getValues();
        const lastRow = data.length;
        const allBg = lastRow > 1 ? mSheet.getRange(2, 1, lastRow-1, 12).getBackgrounds() : [];

        let totalOrders = 0;
        let totalRevenue = 0;
        let totalSettlement = 0;
        let totalProfit = 0;
        let cancelled = 0;
        let completed = 0;
        let unsent = 0;
        let inDelivery = 0;

        for (let i = 1; i < lastRow; i++) {
          const row = data[i];
          const orderNo = String(row[6]).trim();
          const name    = String(row[7]).trim();
          if (!orderNo || !name) continue;

          totalOrders++;

          const gBg = allBg[i-1] ? (allBg[i-1][6]||'').toLowerCase() : '';
          const hBg = allBg[i-1] ? (allBg[i-1][7]||'').toLowerCase() : '';
          const iBg = allBg[i-1] ? (allBg[i-1][8]||'').toLowerCase() : '';

          const isRed = (c) => c.includes('ea9999') || c.includes('e06666');
          if (isRed(gBg) || isRed(hBg) || isRed(iBg)) {
            cancelled++;
            continue;
          }

          const payAmt = parseFloat(row[16]) || 0;
          const settleAmt = parseFloat(row[20]) || 0;
          const profit = parseFloat(row[32]) || 0;

          totalRevenue += payAmt;
          totalSettlement += settleAmt;
          totalProfit += profit;

          if (gBg.includes('87ceeb') || gBg.includes('c9daf8')) {
            completed++;
          }
          else if (hBg.includes('ffff00') || iBg.includes('ffff00')) {
            unsent++;
          }
          else if (gBg.includes('ff9800') || hBg.includes('ff9800') || gBg.includes('ff9900') || hBg.includes('ff9900')) {
            inDelivery++;
          }
        }

        monthData.push({
          month: mName,
          totalOrders,
          totalRevenue,
          totalSettlement,
          totalProfit,
          marginRate: totalRevenue > 0 ? (totalProfit / totalRevenue) : 0,
          cancelled,
          completed,
          unsent,
          inDelivery
        });
      }

      return respond({ success: true, ssId, monthData });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // ★★★ 일자별 통계 ★★★
  if (action === 'dailyStats') {
    try {
      const monthParam = e.parameter.month || '';
      let targetYear, targetMonth;
      if (/^\d{4}-\d{1,2}$/.test(monthParam)) {
        const parts = monthParam.split('-');
        targetYear = parseInt(parts[0]);
        targetMonth = parseInt(parts[1]);
      } else {
        const now = new Date();
        targetYear = now.getFullYear();
        targetMonth = now.getMonth() + 1;
      }

      const sheetName = String(targetYear).slice(2) + '년 ' + targetMonth + '월';
      const mSheet = ss.getSheetByName(sheetName);

      const lastDay = new Date(targetYear, targetMonth, 0).getDate();
      const dailyData = [];
      for (let d = 1; d <= lastDay; d++) {
        dailyData.push({
          date: targetYear + '-' + String(targetMonth).padStart(2,'0') + '-' + String(d).padStart(2,'0'),
          day: d,
          orders: 0,
          revenue: 0,
          settlement: 0,
          profit: 0,
          cancelled: 0
        });
      }

      if (!mSheet) {
        return respond({
          success: true, ssId,
          year: targetYear, month: targetMonth,
          sheetName, sheetExists: false,
          dailyData, summary: {
            totalOrders: 0, totalRevenue: 0, totalSettlement: 0,
            totalProfit: 0, cancelled: 0, marginRate: 0
          }
        });
      }

      const data = mSheet.getDataRange().getValues();
      const lastRow = data.length;
      const allBg = lastRow > 1 ? mSheet.getRange(2, 1, lastRow-1, 12).getBackgrounds() : [];
      // ★ E열(주문일자)은 표시값으로 읽기 (시간대 변환 문제 회피)
      const eDisplayValues = lastRow > 1 ? mSheet.getRange(2, 5, lastRow-1, 1).getDisplayValues() : [];

      // ★ 표시값에서 'YYYY-MM-DD' 추출 (시간대 변환 없는 안전한 방식)
      const parseDisplayDate = (displayStr) => {
        if (!displayStr) return null;
        const s = String(displayStr).trim();
        // 시트에 표시되는 그대로 - 예: "2026-05-20 오전 9:20:55"
        const m = s.match(/(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
        if (m) {
          const y = parseInt(m[1]), mo = parseInt(m[2]), d = parseInt(m[3]);
          if (y >= 2020 && mo >= 1 && mo <= 12 && d >= 1 && d <= 31) {
            return { year: y, month: mo, day: d };
          }
        }
        return null;
      };

      let totalOrders = 0, totalRevenue = 0, totalSettlement = 0;
      let totalProfit = 0, cancelled = 0;
      let unmatched = 0;

      for (let i = 1; i < lastRow; i++) {
        const row = data[i];
        const orderNo = String(row[6]).trim();
        const name    = String(row[7]).trim();
        if (!orderNo || !name) continue;

        // ★ E열 표시값에서 추출 (시트에 보이는 그대로)
        const eDisplay = eDisplayValues[i-1] ? eDisplayValues[i-1][0] : '';
        const dt = parseDisplayDate(eDisplay);

        if (!dt || dt.year !== targetYear || dt.month !== targetMonth) {
          unmatched++;
          continue;
        }
        const dayIdx = dt.day - 1;
        if (dayIdx < 0 || dayIdx >= dailyData.length) continue;

        totalOrders++;
        dailyData[dayIdx].orders++;

        const gBg = allBg[i-1] ? (allBg[i-1][6]||'').toLowerCase() : '';
        const hBg = allBg[i-1] ? (allBg[i-1][7]||'').toLowerCase() : '';
        const iBg = allBg[i-1] ? (allBg[i-1][8]||'').toLowerCase() : '';
        const isRed = (c) => c.includes('ea9999') || c.includes('e06666');
        if (isRed(gBg) || isRed(hBg) || isRed(iBg)) {
          cancelled++;
          dailyData[dayIdx].cancelled++;
          continue;
        }

        const payAmt    = parseFloat(row[16]) || 0;
        const settleAmt = parseFloat(row[20]) || 0;
        const profit    = parseFloat(row[32]) || 0;

        totalRevenue    += payAmt;
        totalSettlement += settleAmt;
        totalProfit     += profit;

        dailyData[dayIdx].revenue    += payAmt;
        dailyData[dayIdx].settlement += settleAmt;
        dailyData[dayIdx].profit     += profit;
      }

      return respond({
        success: true, ssId,
        year: targetYear, month: targetMonth,
        sheetName, sheetExists: true,
        dailyData,
        summary: {
          totalOrders,
          totalRevenue,
          totalSettlement,
          totalProfit,
          cancelled,
          unmatched,
          marginRate: totalRevenue > 0 ? (totalProfit / totalRevenue) : 0
        }
      });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // ═══════════════════════════════════════════════════════════════
  // ★ dayOrders - 특정 날짜의 주문 상세 목록 (일자별 상세 펼침용)
  // ═══════════════════════════════════════════════════════════════
  if (action === 'dayOrders') {
    try {
      const dateStr = e.parameter.date || ''; // 'YYYY-MM-DD'
      if (!/^\d{4}-\d{1,2}-\d{1,2}$/.test(dateStr)) {
        return respond({ success: false, error: 'date 파라미터 필요 (YYYY-MM-DD)' });
      }
      const dParts = dateStr.split('-');
      const targetYear  = parseInt(dParts[0]);
      const targetMonth = parseInt(dParts[1]);
      const targetDay   = parseInt(dParts[2]);

      const sheetName = String(targetYear).slice(2) + '년 ' + targetMonth + '월';
      const mSheet = ss.getSheetByName(sheetName);

      if (!mSheet) {
        return respond({
          success: true, ssId, date: dateStr,
          sheetName, sheetExists: false, orders: []
        });
      }

      const lastRow = mSheet.getLastRow();
      if (lastRow < 2) {
        return respond({ success: true, ssId, date: dateStr, sheetName, orders: [] });
      }

      // 데이터 + 배경색 한번에 읽기 (A~AG = 33개 컬럼)
      const data = mSheet.getRange(2, 1, lastRow - 1, 33).getValues();
      const allBg = mSheet.getRange(2, 1, lastRow - 1, 12).getBackgrounds();
      // ★ E열(주문일자)은 표시값으로 읽기 (시간대 변환 문제 회피)
      const eDisplayValues = mSheet.getRange(2, 5, lastRow - 1, 1).getDisplayValues();

      // ★ 표시값에서 'YYYY-MM-DD' 추출 (시간대 변환 없는 안전한 방식)
      const parseDisplayDate = (displayStr) => {
        if (!displayStr) return null;
        const s = String(displayStr).trim();
        const m = s.match(/(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
        if (m) {
          const y = parseInt(m[1]), mo = parseInt(m[2]), d = parseInt(m[3]);
          if (y >= 2020 && mo >= 1 && mo <= 12 && d >= 1 && d <= 31) {
            return { year: y, month: mo, day: d };
          }
        }
        return null;
      };

      const orders = [];
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const orderNo = String(row[6]).trim();
        const name    = String(row[7]).trim();
        if (!orderNo || !name) continue;

        // ★ E열 표시값에서 추출 (시트에 보이는 그대로)
        const eDisplay = eDisplayValues[i] ? eDisplayValues[i][0] : '';
        const dt = parseDisplayDate(eDisplay);

        if (!dt) continue;
        if (dt.year !== targetYear || dt.month !== targetMonth || dt.day !== targetDay) continue;

        // 취소 여부 (G, H, I열 빨간색)
        const gBg = allBg[i] ? (allBg[i][6]||'').toLowerCase() : '';
        const hBg = allBg[i] ? (allBg[i][7]||'').toLowerCase() : '';
        const iBg = allBg[i] ? (allBg[i][8]||'').toLowerCase() : '';
        const isRedC = (c) => c.includes('ea9999') || c.includes('e06666') ||
                              c.includes('ff0000') || c.includes('cc0000');
        const isCancelled = isRedC(gBg) || isRedC(hBg) || isRedC(iBg);

        orders.push({
          orderNo:   orderNo,
          recipient: name,                        // H열
          product:   String(row[8]||'').trim(),  // I열: 상품
          option:    String(row[9]||'').trim(),  // J열: 옵션
          qty:       row[10] || 1,                // K열: 수량
          phone:     String(row[12]||'').trim(), // M열: phone1
          amount:    parseFloat(row[16]) || 0,    // Q열: 결제금액
          settlement: parseFloat(row[20]) || 0,   // U열: 정산금
          profit:    parseFloat(row[32]) || 0,    // AG열: 이익
          cancelled: isCancelled
        });
      }

      return respond({
        success: true, ssId, date: dateStr, sheetName,
        sheetExists: true, orders: orders, count: orders.length
      });

    } catch (err) {
      return respond({ success: false, error: err.message, stack: err.stack });
    }
  }

  // ═══════════════════════════════════════════════════════════════
  // ★ LCL 진행건 관리 액션들 (lclAdd, lclRemove, lclList)
  // 사용 시트: https://docs.google.com/spreadsheets/d/10hfa94OWjpyvF4eK5w2FYhqsIRbfsn8pPlISoj9N9Jg
  // gid=844259954
  // 헤더 순서: 사업자|주문번호|수령자|전화번호|상품|옵션|수량|금액|주소|우편번호|추가일시|그룹
  // ═══════════════════════════════════════════════════════════════
  const LCL_SS_ID = '10hfa94OWjpyvF4eK5w2FYhqsIRbfsn8pPlISoj9N9Jg';
  const LCL_GID = 844259954;

  // 시트 객체 가져오기 (gid 기준)
  const getLclSheet_ = () => {
    const ss = SpreadsheetApp.openById(LCL_SS_ID);
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId() === LCL_GID) return sheets[i];
    }
    // gid로 못 찾으면 첫 번째 시트 사용 (fallback)
    return sheets[0];
  };

  // LCL 시트에 데이터 추가
  if (action === 'lclAdd') {
    try {
      const orderNo  = e.parameter.orderNo || '';
      const group    = e.parameter.group || 'A';
      const bizName  = e.parameter.bizName || '';
      const name     = e.parameter.name || '';
      const phone    = e.parameter.phone || '';
      const product  = e.parameter.product || '';
      const option   = e.parameter.option || '';
      const qty      = e.parameter.qty || '1';
      const amount   = e.parameter.amount || '0';
      const address  = e.parameter.address || '';
      const zipCode  = e.parameter.zipCode || '';

      if (!orderNo) return respond({ success: false, error: 'orderNo 필요' });

      const lclSheet = getLclSheet_();

      // ★ 중복 체크: 사업자(A) + 수령자(C) + 결제금액(M) 조합
      const lastRow = lclSheet.getLastRow();
      let existingRow = -1;
      if (lastRow > 1) {
        const checkData = lclSheet.getRange(2, 1, lastRow - 1, 13).getValues();
        for (let i = 0; i < checkData.length; i++) {
          const row = checkData[i];
          // A열이 비어있으면 스킵 (실제 데이터 행만 체크)
          if (!String(row[0]).trim()) continue;
          // A=사업자, C=수령자, M=결제금액 (인덱스 0, 2, 12)
          if (String(row[0]).trim() === bizName &&
              String(row[2]).trim() === name &&
              String(row[12]).trim() === String(amount)) {
            existingRow = i + 2;
            break;
          }
        }
      }

      const now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
      // LCL 시트 컬럼 (29개) 매핑
      // A:사업자 B:배송대행지신청서 C:수령자 D:주문일자 E:상품명 F:옵션명 G:수량 H:통관번호
      // I:휴대폰1 J:휴대폰2 K:우편번호 L:배송지주소 M:결제금액 N:배송대행지 ... AC:이익
      const orderDate  = e.parameter.orderDate || '';
      const customs    = e.parameter.customs || '';
      const phone2     = e.parameter.phone2 || '';
      const baejiNo    = e.parameter.baejiNo || '';  // ★ B열: 배송대행지신청서 번호
      // ★ 손익 관련 추가 필드 (수식 컬럼은 시트에서 자동 계산되므로 입력값만 보냄)
      const settleAmt  = parseFloat(e.parameter.settleAmt) || 0;   // 정산예정금액
      const taoAmt     = e.parameter.taoAmt || '';                 // taobao 오더번호
      const shipFee    = parseFloat(e.parameter.shipFee) || 0;     // 배대지배송비 (W)
      const addShip    = parseFloat(e.parameter.addShip) || 0;     // 추가배송비 (Y)
      const amountNum  = parseFloat(amount) || 0;                  // 결제금액 (M)
      // 이익은 시트 수식(AC열)이 자동 계산하므로 코드에서 계산 안 함

      // rowData 배열 (인덱스로 접근하기 위해)
      // 수식 컬럼(O,V,AA,AB,AC)은 빈 문자열로 두고 입력하지 않음
      const rowData = [
        bizName,      // [0] A: 사업자
        baejiNo,      // [1] B: 배송대행지 신청서 ★ 모다허브 C열 값
        name,         // [2] C: 수령자
        orderDate,    // [3] D: 주문일자
        product,      // [4] E: 상품명
        option,       // [5] F: 옵션명
        qty,          // [6] G: 수량
        customs,      // [7] H: 통관번호
        phone,        // [8] I: 휴대폰1
        phone2,       // [9] J: 휴대폰2
        zipCode,      // [10] K: 우편번호
        address,      // [11] L: 배송지주소
        amountNum,    // [12] M: 결제금액
        '',           // [13] N: 배송대행지
        '',           // [14] O: 총결제액 ★ 수식 - 건들지 말 것
        '',           // [15] P: 마켓수수료율
        settleAmt,    // [16] Q: 정산예정금액
        '',           // [17] R: taobao오더번호 (모다허브에 없음 - 빈칸)
        taoAmt,       // [18] S: tao결제액 (모다허브 W열에서 가져옴)
        '',           // [19] T: tao실결제액
        '',           // [20] U: 환율
        '',           // [21] V: 원가계 ★ 수식
        shipFee,      // [22] W: A 배대지배송비
        '',           // [23] X: B 관부가세부담
        addShip,      // [24] Y: C 추가부담배송비
        '',           // [25] Z: D 부가세
        '',           // [26] AA: 지출계 ★ 수식
        '',           // [27] AB: 수수료차감전이익 ★ 수식
        ''            // [28] AC: 이익 ★ 수식
      ];

      let targetRow;
      const NUM_COLS = 29;

      // ★ 수식 컬럼: O(15), V(22), AA(27), AB(28), AC(29) → 절대 건들지 말 것!
      // 입력 가능 컬럼만 셀 단위로 입력
      const setIfData = (row, col, value) => {
        if (value !== '' && value !== null && value !== undefined) {
          try { lclSheet.getRange(row, col).setValue(value); } catch(e) {}
        }
      };

      if (existingRow > 0) {
        // 기존 행 업데이트 - 입력 컬럼만
        targetRow = existingRow;
      } else {
        // ★ A열 기준 마지막 데이터 행 찾기 + 1
        // (다른 컬럼에 수식이나 빈 값이 있어도 A열에 사업자명이 마지막인 행 다음)
        const lastRowSheet = lclSheet.getLastRow();
        let lastDataRow = 1; // 헤더만 있을 때 기본값
        if (lastRowSheet >= 2) {
          // A열 값들 한번에 읽기 (효율적)
          const colA = lclSheet.getRange(2, 1, lastRowSheet - 1, 1).getValues();
          for (let i = colA.length - 1; i >= 0; i--) {
            if (String(colA[i][0]).trim() !== '') {
              lastDataRow = i + 2; // 2부터 시작이므로 +2
              break;
            }
          }
        }
        targetRow = lastDataRow + 1;
        if (targetRow < 2) targetRow = 2;
        Logger.log('[lclAdd] A열 마지막 데이터: ' + lastDataRow + ', 추가할 행: ' + targetRow);
      }

      // 입력 컬럼들 (수식 컬럼 제외)
      // A(1):사업자, C(3):수령자, D(4):주문일자, E(5):상품명, F(6):옵션명,
      // G(7):수량, H(8):통관번호, I(9):휴대폰1, J(10):휴대폰2,
      // K(11):우편번호, L(12):배송지주소, M(13):결제금액
      setIfData(targetRow, 1,  rowData[0]);   // A: 사업자
      setIfData(targetRow, 3,  rowData[2]);   // C: 수령자
      setIfData(targetRow, 4,  rowData[3]);   // D: 주문일자
      setIfData(targetRow, 5,  rowData[4]);   // E: 상품명
      setIfData(targetRow, 6,  rowData[5]);   // F: 옵션명
      setIfData(targetRow, 7,  rowData[6]);   // G: 수량
      setIfData(targetRow, 8,  rowData[7]);   // H: 통관번호
      setIfData(targetRow, 9,  rowData[8]);   // I: 휴대폰1
      setIfData(targetRow, 10, rowData[9]);   // J: 휴대폰2
      setIfData(targetRow, 11, rowData[10]);  // K: 우편번호
      setIfData(targetRow, 12, rowData[11]);  // L: 배송지주소
      setIfData(targetRow, 13, amountNum);    // M: 결제금액
      // O(15) 총결제액 - 수식! 건들지 않음
      setIfData(targetRow, 17, settleAmt);    // Q: 정산예정금액
      // R(18) taobao 오더번호 - 빈칸 (모다허브에 없음)
      setIfData(targetRow, 19, taoAmt);       // S: tao결제액 (모다허브 W열)
      // V(22) 원가계 - 수식!
      setIfData(targetRow, 23, shipFee);      // W: A 배대지배송비
      setIfData(targetRow, 25, addShip);      // Y: C 추가부담배송비
      // AA(27), AB(28), AC(29) - 수식! 건들지 않음

      // 셀 포맷 강제 설정 (수식 아닌 숫자 컬럼만)
      const numericCols = [13, 17, 23, 25]; // M, Q, W, Y
      numericCols.forEach(col => {
        try {
          lclSheet.getRange(targetRow, col).setNumberFormat('#,##0');
        } catch(e) {}
      });

      // 그룹별 색상 (수식 컬럼은 색 안 칠하기 위해 컬럼별 적용)
      // 색상 입힐 컬럼: A~M, N, P, Q, R, S, T, U, W, X, Y, Z (수식 컬럼 O,V,AA,AB,AC 제외)
      const groupColors = { 'A': '#FFE5E5', 'B': '#E5F0FF', 'C': '#E5FFE5' };
      const bgColor = groupColors[group] || '#FFFFFF';
      try {
        // 그냥 전체 행에 색칠 (수식 결과도 색칠되지만 값은 안 변함)
        lclSheet.getRange(targetRow, 1, 1, NUM_COLS).setBackground(bgColor);
      } catch(e) {}

      // 디버그 로그
      Logger.log('[lclAdd] orderNo=' + orderNo + ', amount=' + amountNum + 
                ', settleAmt=' + settleAmt + ', shipFee=' + shipFee +
                ', addShip=' + addShip + ', row=' + targetRow);

      return respond({
        success: true,
        orderNo: orderNo,
        group: group,
        row: targetRow,
        action: existingRow > 0 ? 'updated' : 'added'
      });
    } catch (err) {
      return respond({ success: false, error: err.message });
    }
  }

  // LCL 시트에서 데이터 제거 (사업자 + 수령자 + 금액으로 매칭)
  if (action === 'lclRemove') {
    try {
      const orderNo = e.parameter.orderNo || '';
      const bizName = e.parameter.bizName || '';
      const name    = e.parameter.name || '';
      const amount  = e.parameter.amount || '';
      const amountNum = parseFloat(amount) || 0;

      Logger.log('[lclRemove] 매칭 조건: bizName=' + bizName + ', name=' + name + ', amount=' + amount);

      const lclSheet = getLclSheet_();
      const lastRow = lclSheet.getLastRow();
      if (lastRow < 2) return respond({ success: true, removed: false, reason: '시트 비어있음' });

      // 사업자(A) + 수령자(C) + 금액(M) 조합으로 매칭
      const checkData = lclSheet.getRange(2, 1, lastRow - 1, 13).getValues();
      const tried = []; // 디버그용

      for (let i = checkData.length - 1; i >= 0; i--) {
        const row = checkData[i];
        const rowBiz = String(row[0]).trim();
        const rowName = String(row[2]).trim();
        const rowAmount = row[12];
        const rowAmountNum = parseFloat(rowAmount) || 0;

        tried.push({ row: i + 2, biz: rowBiz, name: rowName, amount: rowAmount });

        // 사업자와 수령자 매칭 + 금액은 숫자 비교 (오차 1원 허용)
        if (rowBiz === bizName &&
            rowName === name &&
            Math.abs(rowAmountNum - amountNum) < 1) {
          lclSheet.deleteRow(i + 2);
          Logger.log('[lclRemove] 삭제 성공: row ' + (i + 2));
          return respond({ success: true, removed: true, row: i + 2 });
        }
      }

      // 매칭 실패 시 첫 5개 시도 정보 반환 (디버그용)
      Logger.log('[lclRemove] 매칭 실패. 시도한 행: ' + JSON.stringify(tried.slice(0, 5)));
      return respond({
        success: true,
        removed: false,
        reason: '매칭 항목 없음',
        searched: { bizName: bizName, name: name, amount: amount },
        tried: tried.slice(0, 5)
      });
    } catch (err) {
      return respond({ success: false, error: err.message });
    }
  }

  // LCL 리스트 조회 - 새 컬럼 구조 반영
  if (action === 'lclList') {
    try {
      const filterGroup = e.parameter.group || '';

      const lclSheet = getLclSheet_();
      const lastRow = lclSheet.getLastRow();
      if (lastRow < 2) return respond({ success: true, list: [] });

      // ★ 성능: A열 기준 마지막 데이터 행만 검사 (빈 행 1000개 스캔 안 함)
      const colA = lclSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      let lastDataRow = 1;
      for (let i = colA.length - 1; i >= 0; i--) {
        if (String(colA[i][0]).trim()) { lastDataRow = i + 2; break; }
      }
      if (lastDataRow < 2) return respond({ success: true, list: [] });

      const numRows = lastDataRow - 1;
      // 전체 컬럼 읽기 (32개) - AE(31)/AF(32)까지 - 착불선불, 경동용달
      const data = lclSheet.getRange(2, 1, numRows, 32).getValues();
      // 배경색으로 그룹 추정 (A=#FFE5E5, B=#E5F0FF, C=#E5FFE5)
      const bgs = lclSheet.getRange(2, 1, numRows, 1).getBackgrounds();

      // ★★★ B열이 비어있는 행을 위한 메인 시트 자동 매칭 (이모지 제거 완료) ★★★
      const BIZ_NAME_TO_SS = {
        '모다mall':   '1',
        '모다mall2':  '2',
        '모다mall3':  '3',
        '모다mall4':  '4',
        '수기mall':   '5',
        '불사자(통합)': '6',
        '불사자':     '6'
      };
      const LCL_SS_MAP = {
        '1': '10hfa94OWjpyvF4eK5w2FYhqsIRbfsn8pPlISoj9N9Jg',
        '2': '1u2kIj8-c6YVtv1XOg40-Im1dL09sGiIxehatDKD8Cu0',
        '3': '137pCqtBMsOgX752pVoglwBk3NyK2aXtAy1R40lLJFiM',
        '4': '1AVe0jVZXFLloVEbJHh65OSru8ZbWnXwplHxAmbaXkGE',
        '5': '1IZEnu3o7WcpRUXJP3AcxPQyNOsUj1J2bTSjvpX14oGo',
        '6': '1FCvAfy4AMClQ-0pdJ_AjdwAOqh05Ldh80fTSXknI6Nc'
      };
      // 사업자명 정규화: 이모지 + 공백 제거
      const normalizeBizName = (s) => {
        return String(s || '')
          .replace(/[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}]/gu, '')
          .replace(/\s+/g, '')
          .trim();
      };

      // 비어있는 B열 항목 → 사업자별 그룹화
      const needMatch = {};
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const bizName = normalizeBizName(row[0]);
        const baeji   = String(row[1] || '').trim();
        const name    = String(row[2] || '').trim();
        if (!bizName || !name) continue;
        if (baeji) continue;  // 이미 값 있으면 스킵
        const ssId = BIZ_NAME_TO_SS[bizName];
        if (!ssId) continue;
        if (!needMatch[ssId]) needMatch[ssId] = [];
        needMatch[ssId].push({ name: name, rowIdx: i });
      }

      // 매칭 수행 (최신 2개월만 - 빠르게)
      const matched = {};  // { rowIdx: baejiNo }
      const matchDebug = {};  // 디버그용
      const ssIdsToSearch = Object.keys(needMatch);
      for (const ssId of ssIdsToSearch) {
        try {
          const targetItems = needMatch[ssId];
          const targetSsId = LCL_SS_MAP[ssId];
          if (!targetSsId) {
            matchDebug['ssId=' + ssId] = 'LCL_SS_MAP 없음';
            continue;
          }
          const targetSs = ssId === '1'
            ? SpreadsheetApp.getActiveSpreadsheet()
            : SpreadsheetApp.openById(targetSsId);
          // 월 시트만 + 최신순으로 2개
          const allMonthSheets = targetSs.getSheets().filter(function(s) {
            return /\d+년 \d+월/.test(s.getName());
          });
          const parseYM = (nm) => {
            const m = nm.match(/(\d+)년\s*(\d+)월/);
            return m ? (parseInt(m[1])*100 + parseInt(m[2])) : 0;
          };
          allMonthSheets.sort((a, b) => parseYM(b.getName()) - parseYM(a.getName()));
          const monthSheets = allMonthSheets.slice(0, 2);
          matchDebug['ssId=' + ssId + '_검색시트'] = monthSheets.map(s => s.getName());
          matchDebug['ssId=' + ssId + '_타겟이름들'] = targetItems.map(t => t.name);

          for (const ms of monthSheets) {
            const msLastRow = ms.getLastRow();
            if (msLastRow < 2) continue;
            // C열(배대지) + H열(수령자)만 - 가벼움
            const msBaeji = ms.getRange(2, 3, msLastRow - 1, 1).getValues();
            const msName  = ms.getRange(2, 8, msLastRow - 1, 1).getValues();
            for (const targetItem of targetItems) {
              if (matched[targetItem.rowIdx]) continue;
              for (let j = 0; j < msName.length; j++) {
                const mName = String(msName[j][0] || '').trim();
                const mBaeji = String(msBaeji[j][0] || '').trim();
                if (!mBaeji) continue;
                if (mName === targetItem.name) {
                  matched[targetItem.rowIdx] = mBaeji;
                  break;
                }
              }
            }
          }
        } catch(searchErr) {
          Logger.log('[lclList] 자동매칭 ssId=' + ssId + ' 오류: ' + searchErr.message);
          matchDebug['ssId=' + ssId + '_오류'] = searchErr.message;
        }
      }

      // ★ 매칭된 결과를 시트 B열에 영구 저장 (다음부터 매칭 안 해도 됨!)
      const matchedKeys = Object.keys(matched);
      if (matchedKeys.length > 0) {
        try {
          // 매칭된 행마다 B열에 값 쓰기 (개별 setValue - 안정적)
          for (const rowIdxStr of matchedKeys) {
            const rowIdx = parseInt(rowIdxStr);
            const realRow = rowIdx + 2;  // 시트 행번호 (헤더 1행 + 인덱스)
            lclSheet.getRange(realRow, 2).setValue(matched[rowIdx]);
            // data 배열도 업데이트 (응답에 반영)
            data[rowIdx][1] = matched[rowIdx];
          }
          Logger.log('[lclList] 자동매칭 ' + matchedKeys.length + '건 시트 B열에 저장 완료');
        } catch(saveErr) {
          Logger.log('[lclList] 시트 저장 실패: ' + saveErr.message);
        }
      }

      const list = [];
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        // 사업자(A) + 수령자(C) 둘 다 비어있으면 스킵
        const bizName = String(row[0]).trim();
        const name = String(row[2]).trim();
        if (!bizName && !name) continue;

        // 배경색으로 그룹 추정
        // A=#FFE5E5(분홍/출고대기), B=#E5F0FF(파랑/출고신청), C=#E5FFE5(초록/출고완료), D=#B3E5FC(하늘/정산완료)
        const bg = (bgs[i][0] || '').toLowerCase();
        let grp = 'A';
        if (bg.includes('b3e5fc')) grp = 'D';        // 정산완료 (하늘)
        else if (bg.includes('e5f0ff')) grp = 'B';   // 출고신청 (파랑)
        else if (bg.includes('e5ffe5')) grp = 'C';   // 출고완료 (초록)
        else if (bg.includes('ffe5e5')) grp = 'A';   // 출고대기 (분홍)

        if (filterGroup && grp !== filterGroup) continue;

        // 수식 결과 읽기 (이익 등)
        const profit = parseFloat(row[28]) || 0;  // AC: 이익 (수식 결과)
        const amount = parseFloat(row[12]) || 0;  // M: 결제금액
        // 마진율 계산 (이익 / 결제금액)
        const marginRate = (amount > 0 && profit > 0) ? (profit / amount) : 0;

        // ★ baeji 값: B열에 있는 값 그대로 사용 (자동매칭 제거됨)
        const baejiVal = String(row[1]).trim();

        list.push({
          bizName:  bizName,                    // A
          baeji:    baejiVal,                   // B: 배송대행지 신청서 (자동 매칭 fallback)
          name:     name,                       // C: 수령자
          orderDate:String(row[3]).trim(),     // D: 주문일자
          product:  String(row[4]).trim(),     // E: 상품명
          option:   String(row[5]).trim(),     // F: 옵션명
          qty:      row[6] || 1,                // G: 수량
          customs:  String(row[7]).trim(),     // H: 통관번호
          phone:    String(row[8]).trim(),     // I: 휴대폰1
          phone2:   String(row[9]).trim(),     // J: 휴대폰2
          zipCode:  String(row[10]).trim(),    // K: 우편번호
          address:  String(row[11]).trim(),    // L: 배송지주소
          amount:   amount,                     // M: 결제금액
          settleAmt: parseFloat(row[16]) || 0,  // Q: 정산예정금액
          taoOrder: String(row[17]).trim(),     // R: taobao 오더번호
          taoPay:   parseFloat(row[18]) || 0,   // S: tao 결제액
          shipFee:  parseFloat(row[22]) || 0,   // W: 배대지배송비
          addShip:  parseFloat(row[24]) || 0,   // Y: 추가배송비
          profit:   profit,                     // AC: 이익
          marginRate: marginRate,               // 계산값
          payType:  String(row[30] || '').trim() || '착불',  // AE: 착불/선불 (기본 착불)
          shipType: String(row[31] || '').trim() || '경동',  // AF: 경동/용달 (기본 경동)
          group:    grp,
          row:      i + 2
        });
      }

      return respond({ success: true, list: list, count: list.length, matchDebug: matchDebug, matchedCount: Object.keys(matched).length });
    } catch (err) {
      return respond({ success: false, error: err.message });
    }
  }

  // ★ LCL 그룹 변경 (A:출고대기 → B:출고신청 → C:출고완료)
  // 매칭: 사업자(A) + 수령자(C) + 금액(M) - lclRemove와 동일
  // 변경: A열 배경색만 변경 (그룹 판정 기준)
  if (action === 'lclChangeGroup') {
    try {
      const orderNo  = e.parameter.orderNo  || '';
      const bizName  = e.parameter.bizName  || '';
      const name     = e.parameter.name     || '';
      const amount   = e.parameter.amount   || '';
      const newGroup = e.parameter.newGroup || '';
      const amountNum = parseFloat(amount) || 0;

      // 그룹 → 배경색 매핑
      const GROUP_COLORS = { 'A': '#FFE5E5', 'B': '#E5F0FF', 'C': '#E5FFE5' };
      const newColor = GROUP_COLORS[newGroup];
      if (!newColor) {
        return respond({ success: false, error: '잘못된 그룹: ' + newGroup });
      }

      Logger.log('[lclChangeGroup] 조건: biz=' + bizName + ', name=' + name + ', amount=' + amount + ', newGroup=' + newGroup);

      const lclSheet = getLclSheet_();
      const lastRow = lclSheet.getLastRow();
      if (lastRow < 2) return respond({ success: false, error: '시트가 비어있음' });

      // 사업자(A) + 수령자(C) + 금액(M) 조합으로 매칭
      const checkData = lclSheet.getRange(2, 1, lastRow - 1, 13).getValues();
      const tried = [];

      for (let i = checkData.length - 1; i >= 0; i--) {
        const row = checkData[i];
        const rowBiz = String(row[0]).trim();
        const rowName = String(row[2]).trim();
        const rowAmount = row[12];
        const rowAmountNum = parseFloat(rowAmount) || 0;

        tried.push({ row: i + 2, biz: rowBiz, name: rowName, amount: rowAmount });

        if (rowBiz === bizName &&
            rowName === name &&
            Math.abs(rowAmountNum - amountNum) < 1) {
          const targetRow = i + 2;
          // 행 전체 (A~AC, 29개 컬럼) 배경색 변경
          const NUM_COLS = 29;
          lclSheet.getRange(targetRow, 1, 1, NUM_COLS).setBackground(newColor);
          Logger.log('[lclChangeGroup] 성공: row ' + targetRow + ', color ' + newColor);
          return respond({ success: true, row: targetRow, newGroup: newGroup });
        }
      }

      Logger.log('[lclChangeGroup] 매칭 실패: ' + JSON.stringify(tried.slice(0, 5)));
      return respond({
        success: false,
        error: '매칭 항목 없음',
        searched: { bizName: bizName, name: name, amount: amount },
        tried: tried.slice(0, 5)
      });
    } catch (err) {
      return respond({ success: false, error: err.message });
    }
  }

  // ★ LCL 정산 - 선택된 행에 배대지(W열) + 관세사(X열) 저장
  // 매칭: 사업자(A) + 수령자(C) + 금액(M)
  // 저장: W열(22) = shipFee, X열(23) = customs
  // 이익(AC열)은 시트 수식이 자동 계산
  if (action === 'lclSettle') {
    try {
      const bizName  = e.parameter.bizName || '';
      const name     = e.parameter.name    || '';
      const amount   = e.parameter.amount  || '';
      const shipFee  = parseFloat(e.parameter.shipFee) || 0;
      const customs  = parseFloat(e.parameter.customs) || 0;
      const amountNum = parseFloat(amount) || 0;

      Logger.log('[lclSettle] 조건: biz=' + bizName + ', name=' + name + ', amount=' + amount + ', ship=' + shipFee + ', customs=' + customs);

      const lclSheet = getLclSheet_();
      const lastRow = lclSheet.getLastRow();
      if (lastRow < 2) return respond({ success: false, error: '시트가 비어있음' });

      // 사업자(A) + 수령자(C) + 금액(M)으로 매칭
      const checkData = lclSheet.getRange(2, 1, lastRow - 1, 13).getValues();

      for (let i = checkData.length - 1; i >= 0; i--) {
        const row = checkData[i];
        const rowBiz = String(row[0]).trim();
        const rowName = String(row[2]).trim();
        const rowAmount = row[12];
        const rowAmountNum = parseFloat(rowAmount) || 0;

        if (rowBiz === bizName &&
            rowName === name &&
            Math.abs(rowAmountNum - amountNum) < 1) {
          const targetRow = i + 2;
          // W열(23) = 배대지배송비
          lclSheet.getRange(targetRow, 23).setValue(shipFee);
          // X열(24) = 관세사 (관부가세부담)
          lclSheet.getRange(targetRow, 24).setValue(customs);
          // ★ D그룹(정산완료) 배경색 #B3E5FC (하늘색)로 변경
          const SETTLE_COLOR = '#B3E5FC';
          const NUM_COLS = 29;
          lclSheet.getRange(targetRow, 1, 1, NUM_COLS).setBackground(SETTLE_COLOR);
          Logger.log('[lclSettle] 성공: row ' + targetRow + ', ship=' + shipFee + ', customs=' + customs);
          return respond({ success: true, row: targetRow, shipFee: shipFee, customs: customs });
        }
      }

      return respond({
        success: false,
        error: '매칭 항목 없음',
        searched: { bizName: bizName, name: name, amount: amount }
      });
    } catch (err) {
      return respond({ success: false, error: err.message });
    }
  }

  // ★ LCL 착불/선불 + 경동/용달 토글 저장
  // 매칭: 사업자(A) + 수령자(C) + 금액(M)
  // 저장: AE열(31) = payType (착불/선불), AF열(32) = shipType (경동/용달)
  if (action === 'lclSetPayShip') {
    try {
      const bizName  = e.parameter.bizName || '';
      const name     = e.parameter.name    || '';
      const amount   = e.parameter.amount  || '';
      const field    = e.parameter.field   || ''; // 'payType' 또는 'shipType'
      const value    = e.parameter.value   || ''; // 새 값
      const amountNum = parseFloat(amount) || 0;

      if (field !== 'payType' && field !== 'shipType') {
        return respond({ success: false, error: '잘못된 field: ' + field });
      }

      const lclSheet = getLclSheet_();
      const lastRow = lclSheet.getLastRow();
      if (lastRow < 2) return respond({ success: false, error: '시트가 비어있음' });

      const checkData = lclSheet.getRange(2, 1, lastRow - 1, 13).getValues();
      for (let i = checkData.length - 1; i >= 0; i--) {
        const row = checkData[i];
        const rowBiz = String(row[0]).trim();
        const rowName = String(row[2]).trim();
        const rowAmount = row[12];
        const rowAmountNum = parseFloat(rowAmount) || 0;

        if (rowBiz === bizName &&
            rowName === name &&
            Math.abs(rowAmountNum - amountNum) < 1) {
          const targetRow = i + 2;
          // AE(31) = payType, AF(32) = shipType
          const targetCol = (field === 'payType') ? 31 : 32;
          lclSheet.getRange(targetRow, targetCol).setValue(value);
          Logger.log('[lclSetPayShip] 성공: row ' + targetRow + ', ' + field + '=' + value);
          return respond({ success: true, row: targetRow, field: field, value: value });
        }
      }

      return respond({
        success: false,
        error: '매칭 항목 없음',
        searched: { bizName: bizName, name: name, amount: amount }
      });
    } catch (err) {
      return respond({ success: false, error: err.message });
    }
  }

  // ★ 주문번호 자동 생성 (중복 체크 포함)
  // 형식: 20yymmddNNN (예: 20260524001)
  if (action === 'generateOrderNo') {
    try {
      const now = new Date();
      const yy = String(now.getFullYear()).slice(2);
      const mm = String(now.getMonth() + 1).padStart(2, '0');
      const dd = String(now.getDate()).padStart(2, '0');
      const prefix = '20' + yy + mm + dd;

      // 이번달 시트 - 짧은 형식 우선
      const fullYear = now.getFullYear();
      const monthNum = now.getMonth() + 1;
      const candidateNames = [
        yy + '년 ' + monthNum + '월',
        yy + '년' + monthNum + '월',
        fullYear + '년 ' + monthNum + '월',
        fullYear + '년' + monthNum + '월'
      ];
      let monthSheet = null;
      // 데이터 있는 시트 우선
      for (const cn of candidateNames) {
        const s = ss.getSheetByName(cn);
        if (s && s.getLastRow() > 1) { monthSheet = s; break; }
      }
      if (!monthSheet) {
        for (const cn of candidateNames) {
          const s = ss.getSheetByName(cn);
          if (s) { monthSheet = s; break; }
        }
      }
      if (!monthSheet) {
        const monthPattern = new RegExp('^(' + fullYear + '|' + yy + ')년\\s*' + monthNum + '월$');
        for (const s of ss.getSheets()) {
          if (monthPattern.test(s.getName())) { monthSheet = s; break; }
        }
      }

      let maxSeq = 0;
      try {
        if (monthSheet) {
          const rawLastRow = monthSheet.getLastRow();
          // ★ A열 기준 진짜 데이터 마지막 행 찾기
          let realLastRow = 1;
          if (rawLastRow >= 2) {
            const colA = monthSheet.getRange(2, 1, rawLastRow - 1, 1).getValues();
            for (let i = colA.length - 1; i >= 0; i--) {
              if (String(colA[i][0]).trim()) {
                realLastRow = i + 2;
                break;
              }
            }
          }
          if (realLastRow >= 2) {
            const orderNoCol = monthSheet.getRange(2, 7, realLastRow - 1, 1).getValues();
            for (let i = 0; i < orderNoCol.length; i++) {
              const ono = String(orderNoCol[i][0] || '').trim();
              if (ono.startsWith(prefix) && ono.length >= prefix.length + 3) {
                const seq = parseInt(ono.substring(prefix.length).slice(0, 3));
                if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
              }
            }
          }
        }
      } catch(scanErr) {
        Logger.log('[generateOrderNo] 스캔 오류: ' + scanErr.message);
      }

      const nextSeq = String(maxSeq + 1).padStart(3, '0');
      const newOrderNo = prefix + nextSeq;
      return respond({ success: true, orderNo: newOrderNo });
    } catch (err) {
      return respond({ success: false, error: err.message });
    }
  }

  // ★ 직접 주문 입력 - 현재 사업자의 이번달 시트에 새 행 추가
  // ★ 수식이 누락된 행에 수식 채우기 (이전에 잘못 들어간 데이터 정리)
  if (action === 'fillMissingFormulas') {
    try {
      const targetSheetName = e.parameter.sheetName || '';
      let fixed = 0;
      const allSheets = ss.getSheets();
      const monthSheets = allSheets.filter(s => /^\d{2,4}년\s*\d+월$/.test(s.getName()));
      const targetSheets = targetSheetName
        ? monthSheets.filter(s => s.getName() === targetSheetName)
        : monthSheets;

      for (const mSheet of targetSheets) {
        const lastRow = mSheet.getLastRow();
        if (lastRow < 3) continue;

        // 수식이 있는 기준 행 찾기 (S열 19 또는 AC열 29)
        let formulaRow = 0;
        for (let r = 2; r <= lastRow; r++) {
          const f = mSheet.getRange(r, 19).getFormula();
          if (f && f.startsWith('=')) { formulaRow = r; break; }
        }
        if (formulaRow === 0) continue;  // 수식 있는 행 없음

        const lastCol = mSheet.getLastColumn();
        // 수식이 없는 행들에 수식만 채우기
        for (let r = 2; r <= lastRow; r++) {
          if (r === formulaRow) continue;
          // A열에 데이터 있는 행만 (실제 데이터 행)
          const hasData = mSheet.getRange(r, 1).getValue();
          if (!hasData) continue;
          // S열 수식이 비어있는지 확인
          const fCheck = mSheet.getRange(r, 19).getFormula();
          if (fCheck && fCheck.startsWith('=')) continue;  // 이미 수식 있음

          // 수식 있는 행에서 수식만 복사
          const sourceRange = mSheet.getRange(formulaRow, 1, 1, lastCol);
          const targetRange = mSheet.getRange(r, 1, 1, lastCol);
          // 수식만 복사 (PASTE_FORMULA 모드)
          sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
          fixed++;
        }
      }
      SpreadsheetApp.flush();
      return respond({ success: true, fixed: fixed, sheets: targetSheets.map(s => s.getName()) });
    } catch(err) {
      return respond({ success: false, error: err.message });
    }
  }

  // ★ 휴대폰 형식 일괄 정리 - 앞 0 복구 + 텍스트 형식 강제
  // 현재 사업자의 이번달 시트에서 M열(휴대폰1), N열(휴대폰2)의 형식 정리
  if (action === 'fixPhoneFormat') {
    try {
      const targetSheetName = e.parameter.sheetName || '';
      let fixed = 0;
      const allSheets = ss.getSheets();
      const monthSheets = allSheets.filter(s => /^\d{2,4}년\s*\d+월$/.test(s.getName()));
      const targetSheets = targetSheetName
        ? monthSheets.filter(s => s.getName() === targetSheetName)
        : monthSheets;

      for (const mSheet of targetSheets) {
        const lastRow = mSheet.getLastRow();
        if (lastRow < 2) continue;
        // M열(13), N열(14) 한꺼번에 처리
        const phoneRange = mSheet.getRange(2, 13, lastRow - 1, 2);
        phoneRange.setNumberFormat('@');  // 텍스트 형식 강제
        const values = phoneRange.getValues();
        const newValues = values.map(row => {
          return [row[0], row[1]].map(v => {
            const s = String(v || '').trim().replace(/-/g, '');
            if (!s) return '';
            // 10자리이고 0으로 시작 안 하면 앞에 0 붙이기
            if (s.length === 10 && !s.startsWith('0')) return '0' + s;
            return s;
          });
        });
        phoneRange.setValues(newValues);
        // 통관번호(L=12), 우편번호(O=15)도 텍스트 형식
        mSheet.getRange(2, 12, lastRow - 1, 1).setNumberFormat('@');
        mSheet.getRange(2, 15, lastRow - 1, 1).setNumberFormat('@');
        fixed += values.length;
      }
      SpreadsheetApp.flush();
      return respond({ success: true, fixed: fixed, sheets: targetSheets.map(s => s.getName()) });
    } catch(err) {
      return respond({ success: false, error: err.message });
    }
  }

  if (action === 'directOrderAdd') {
    try {
      const orderNo   = e.parameter.orderNo || '';
      const orderDate = e.parameter.orderDate || '';
      const name      = e.parameter.name || '';
      const customs   = e.parameter.customs || '';
      const product   = e.parameter.product || '';
      const option    = e.parameter.option || '';
      const qty       = e.parameter.qty || '1';
      const phone     = e.parameter.phone || '';
      const phone2    = e.parameter.phone2 || '';
      const zipCode   = e.parameter.zipCode || '';
      const address   = e.parameter.address || '';
      const amount    = parseFloat(e.parameter.amount) || 0;
      const settleAmt = parseFloat(e.parameter.settleAmt) || 0;  // ★ 정산예정금액

      // ★ 디버그: 어떤 SS를 사용하는지 미리 확인
      const debugSsInfo = {
        receivedSsId: e.parameter.ssId || '(없음)',
        usingSsId: ssId,
        ssName: ss.getName(),
        ssId_actual: ss.getId(),
        ssUrl: ss.getUrl()
      };
      Logger.log('[directOrderAdd] SS 정보: ' + JSON.stringify(debugSsInfo));

      // 이번달 시트 이름 - 여러 형식 시도 (짧은 형식 우선)
      const now = new Date();
      const yy = now.getFullYear();
      const yyShort = String(yy).slice(2);  // 26
      const mm = now.getMonth() + 1;
      const candidateNames = [
        yyShort + '년 ' + mm + '월',     // 26년 5월  ★ 기존 시트가 이 형식일 가능성 높음
        yyShort + '년' + mm + '월',      // 26년5월
        yy + '년 ' + mm + '월',          // 2026년 5월
        yy + '년' + mm + '월'            // 2026년5월
      ];

      let monthSheet = null;
      let usedName = '';
      // ★ 데이터가 있는 시트 우선 (빈 시트 무시)
      for (const cn of candidateNames) {
        const s = ss.getSheetByName(cn);
        if (s && s.getLastRow() > 1) { monthSheet = s; usedName = cn; break; }  // 데이터 있는 시트
      }
      // 그래도 없으면 빈 시트라도 찾기
      if (!monthSheet) {
        for (const cn of candidateNames) {
          const s = ss.getSheetByName(cn);
          if (s) { monthSheet = s; usedName = cn; break; }
        }
      }
      // 정규식으로 다른 형식 시트도 검색
      if (!monthSheet) {
        const allSheets = ss.getSheets();
        const monthPattern = new RegExp('^(' + yy + '|' + yyShort + ')년\\s*' + mm + '월$');
        for (const s of allSheets) {
          if (monthPattern.test(s.getName())) {
            monthSheet = s;
            usedName = s.getName();
            break;
          }
        }
      }

      if (!monthSheet) {
        const allNames = ss.getSheets().map(s => s.getName()).join(', ');
        Logger.log('[directOrderAdd] 시트 못 찾음. 시도한 이름: ' + JSON.stringify(candidateNames) + ', 전체: ' + allNames);
        return respond({
          success: false,
          error: '이번달 시트를 찾을 수 없음',
          tried: candidateNames,
          allSheets: ss.getSheets().map(s => s.getName())
        });
      }

      Logger.log('[directOrderAdd] 시트 찾음: ' + usedName);

      // ★ 진짜 데이터 마지막 행 찾기 - A열(상태) 기준으로 역방향 스캔
      // getLastRow()는 서식만 있어도 카운트해서 빈 공간 뒤로 가는 문제 방지
      let realLastRow = 1;  // 헤더만 있을 때
      const rawLastRow = monthSheet.getLastRow();
      if (rawLastRow >= 2) {
        const colA = monthSheet.getRange(2, 1, rawLastRow - 1, 1).getValues();
        for (let i = colA.length - 1; i >= 0; i--) {
          if (String(colA[i][0]).trim()) {
            realLastRow = i + 2;  // 시트 행번호 (헤더 1행 + 인덱스)
            break;
          }
        }
      }
      Logger.log('[directOrderAdd] rawLastRow=' + rawLastRow + ', realLastRow=' + realLastRow);

      // 중복 주문번호 체크 (실제 데이터 범위만)
      if (realLastRow >= 2) {
        const existingOrderNos = monthSheet.getRange(2, 7, realLastRow - 1, 1).getValues();
        for (let i = 0; i < existingOrderNos.length; i++) {
          if (String(existingOrderNos[i][0]).trim() === orderNo) {
            return respond({ success: false, error: '주문번호 중복: ' + orderNo });
          }
        }
      }

      // 새 행 - 메인 시트 컬럼 매핑
      // A상태 B? C배대지 D? E주문일자 F판매처 G주문번호 H수령자 I상품 J옵션 K수량 L통관 M휴대1 N휴대2 O우편 P주소 Q결제 ... U정산
      const newRow = [];
      newRow[0]  = '정상';
      newRow[1]  = '';
      newRow[2]  = '';
      newRow[3]  = '';
      newRow[4]  = orderDate;
      newRow[5]  = '직접입력';
      newRow[6]  = orderNo;
      newRow[7]  = name;
      newRow[8]  = product;
      newRow[9]  = option;
      newRow[10] = qty;
      newRow[11] = customs;
      newRow[12] = phone;
      newRow[13] = phone2;
      newRow[14] = zipCode;
      newRow[15] = address;
      newRow[16] = amount;     // Q: 결제금액
      newRow[17] = '';
      newRow[18] = '';
      newRow[19] = '';
      newRow[20] = settleAmt;  // U: 정산예정금액 ★
      // 빈 자리 채우기 (21개 컬럼)
      while (newRow.length < 21) newRow.push('');

      const targetRow = realLastRow + 1;

      // ★ 1단계: 수식이 있는 행 찾기 (위에서부터 검색)
      // S열(19) 또는 AC열(29)에 수식이 있는 행을 찾음
      let formulaRow = 0;
      if (realLastRow >= 2) {
        // 마지막 행부터 위로 거슬러 올라가며 수식 있는 행 찾기
        for (let r = realLastRow; r >= 2; r--) {
          const formulaCell = monthSheet.getRange(r, 19).getFormula();  // S열 수식 체크
          if (formulaCell && formulaCell.startsWith('=')) {
            formulaRow = r;
            break;
          }
        }
        // S열 못 찾으면 AC열로 시도
        if (formulaRow === 0) {
          for (let r = realLastRow; r >= 2; r--) {
            const formulaCell = monthSheet.getRange(r, 29).getFormula();  // AC열 수식
            if (formulaCell && formulaCell.startsWith('=')) {
              formulaRow = r;
              break;
            }
          }
        }
      }
      Logger.log('[directOrderAdd] 수식 있는 행: ' + formulaRow);

      // ★ 수식이 있는 행에서 수식 + 서식만 복사 (값은 복사하지 않음)
      if (formulaRow > 0) {
        const lastCol = monthSheet.getLastColumn();
        const sourceRange = monthSheet.getRange(formulaRow, 1, 1, lastCol);
        const targetRange = monthSheet.getRange(targetRow, 1, 1, lastCol);
        // 1) 서식만 복사 (배경색, 글꼴 등)
        sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
        // 2) 수식만 복사 (값은 복사 안 됨 - 새 행은 빈 셀 상태에서 수식만 들어감)
        sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
      }

      // ★ 2단계: 데이터 컬럼만 개별적으로 덮어쓰기 (수식이 있는 빈 컬럼은 건드리지 않음)
      const dataMap = [
        { col: 1,  val: '정상' },              // A: 상태
        { col: 5,  val: orderDate },           // E: 주문일자
        { col: 6,  val: '직접입력' },          // F: 판매처
        { col: 7,  val: orderNo },             // G: 주문번호
        { col: 8,  val: name },                // H: 수령자
        { col: 9,  val: product },             // I: 상품명
        { col: 10, val: option },              // J: 옵션명
        { col: 11, val: qty },                 // K: 수량
        { col: 12, val: customs },             // L: 통관번호
        { col: 13, val: phone },               // M: 휴대폰1
        { col: 14, val: phone2 },              // N: 휴대폰2
        { col: 15, val: zipCode },             // O: 우편번호
        { col: 16, val: address },             // P: 배송지주소
        { col: 17, val: amount },              // Q: 결제금액
        { col: 21, val: settleAmt }            // U: 정산예정금액
      ];
      // 텍스트 형식이 필요한 컬럼들
      const textColSet = {7: true, 12: true, 13: true, 14: true, 15: true};
      dataMap.forEach(item => {
        const cell = monthSheet.getRange(targetRow, item.col);
        if (textColSet[item.col]) {
          cell.setNumberFormat('@');
          cell.setValue(String(item.val || ''));
        } else {
          cell.setValue(item.val);
        }
      });

      SpreadsheetApp.flush();  // ★ 강제 저장

      // ★ 저장 직후 verify - 진짜 저장됐는지 다시 읽어봄
      const verifyOrderNo = String(monthSheet.getRange(targetRow, 7).getValue()).trim();
      const verifyName = String(monthSheet.getRange(targetRow, 8).getValue()).trim();
      const ssName = ss.getName();
      const ssUrl = ss.getUrl();

      Logger.log('[directOrderAdd] 성공: ssId=' + ssId + ', SS=' + ssName + ', 시트=' + usedName + ', row=' + targetRow + ', 검증=' + verifyOrderNo + '/' + verifyName);
      return respond({
        success: true,
        row: targetRow,
        orderNo: orderNo,
        sheet: usedName,
        ssId: ssId,
        ssName: ssName,
        ssUrl: ssUrl,
        debugSsInfo: debugSsInfo,
        verify: {
          orderNo: verifyOrderNo,
          name: verifyName
        }
      });
    } catch (err) {
      return respond({ success: false, error: err.message });
    }
  }

  // ═══════════════════════════════════════════════════════════════
  // ★ 패스트로 연동 액션들 — DryRun (테스트) / Sync (실제 동기화)
  // ═══════════════════════════════════════════════════════════════

  if (action === 'fastroDryRun') {
    try {
      var fr_mbId = getFastroMbId_(ssId);
      if (!fr_mbId) return respond({ error: 'mb_id 미등록 (ssId=' + ssId + ')' });

      var fr_days = parseInt(e.parameter.days) || 30;
      var fr_today = new Date();
      var fr_past = new Date(fr_today.getTime() - fr_days * 24 * 60 * 60 * 1000);
      var fr_fmt = function(d) {
        return d.getFullYear() + '-' +
               String(d.getMonth() + 1).padStart(2, '0') + '-' +
               String(d.getDate()).padStart(2, '0');
      };
      var fr_sdate = fr_fmt(fr_past), fr_edate = fr_fmt(fr_today);

      var fr_result = fastroFetch_(ssId, fr_mbId, fr_sdate, fr_edate, 100);
      if (!fr_result.ok) return respond({ error: fr_result.error });

      var fr_records = parseFastroHtml_(fr_result.html);
      return respond({
        success: true,
        ssId: ssId,
        mbId: fr_mbId,
        sdate: fr_sdate,
        edate: fr_edate,
        htmlLength: fr_result.html.length,
        recordCount: fr_records.length,
        records: fr_records.slice(0, 30),
        prepaidCount: fr_records.filter(function(r) { return r.prepaid > 0; }).length
      });
    } catch (err) {
      return respond({ error: err.message });
    }
  }

  if (action === 'fastroSync') {
    try {
      var fs_mbId = getFastroMbId_(ssId);
      if (!fs_mbId) return respond({ error: 'mb_id 미등록 (ssId=' + ssId + ')' });

      var fs_days = parseInt(e.parameter.days) || 30;
      var fs_limit = parseInt(e.parameter.limit) || 200;
      var fs_today = new Date();
      var fs_past = new Date(fs_today.getTime() - fs_days * 24 * 60 * 60 * 1000);
      var fs_fmt = function(d) {
        return d.getFullYear() + '-' +
               String(d.getMonth() + 1).padStart(2, '0') + '-' +
               String(d.getDate()).padStart(2, '0');
      };
      var fs_sdate = e.parameter.sdate || fs_fmt(fs_past);
      var fs_edate = e.parameter.edate || fs_fmt(fs_today);

      var fs_result = fastroFetch_(ssId, fs_mbId, fs_sdate, fs_edate, fs_limit);
      if (!fs_result.ok) return respond({ error: fs_result.error });

      var fs_records = parseFastroHtml_(fs_result.html);

      var fs_allSheets = ss.getSheets();
      var fs_monthSheets = fs_allSheets.filter(function(s) {
        return /^\d{2}년 \d+월$/.test(s.getName());
      });
      fs_monthSheets.sort(function(a, b) {
        var pa = a.getName().match(/(\d+)년 (\d+)월/);
        var pb = b.getName().match(/(\d+)년 (\d+)월/);
        if (!pa || !pb) return 0;
        return (Number(pb[1]) * 100 + Number(pb[2])) - (Number(pa[1]) * 100 + Number(pa[2]));
      });
      var fs_targetSheets = fs_monthSheets.slice(0, 3);
      var fs_sheetCache = fs_targetSheets.map(function(s) {
        return { sheet: s, data: s.getDataRange().getValues() };
      });

      var fs_updated = 0, fs_unmatched = 0;
      var fs_detail = [];

      // ★ SH번호 기준 매칭 ★ (색깔은 칠하지 않고 노트만)
      fs_records.forEach(function(rec) {
        if (!rec.sh) return;  // SH번호 없으면 스킵
        var shFull = rec.sh;
        var shNum  = rec.sh.replace(/^SH/, '');
        var found = false;
        for (var sc = 0; sc < fs_sheetCache.length; sc++) {
          var sd = fs_sheetCache[sc];
          for (var i = 1; i < sd.data.length; i++) {
            var cellSh = String(sd.data[i][2]).trim();  // C열 값
            if (!cellSh) continue;
            var cellShNum = cellSh.replace(/^SH/, '');
            if (cellSh === shFull || cellShNum === shNum) {
              var firstItem = rec.items.length > 0 ? rec.items[0] : { itCode: '', orderNo: '' };
              var note = buildFastroNote_(rec, firstItem);
              if (rec.items.length > 1) {
                note += '\n\n[전체 ' + rec.items.length + '개 상품]';
                rec.items.forEach(function(it, idx) {
                  note += '\n' + (idx+1) + '. ' + it.itCode + ' / 타오:' + it.orderNo;
                });
              }
              // 노트만 저장 (색깔은 건드리지 않음 - 기존 색깔 시스템과 충돌 방지)
              sd.sheet.getRange(i + 1, 8).setNote(note);
              
             
              fs_updated++;
              found = true;
              fs_detail.push({
                sheet: sd.sheet.getName(),
                row: i + 1,
                cellSh: cellSh,
                gr: rec.gr,
                sh: rec.sh,
                invoice: rec.invoice || '(미발급)',
                state: rec.state,
                itemCount: rec.items.length,
                shipFee: rec.shipFee,
                prepaid: rec.prepaid
              });
              break;
            }
          }
          if (found) break;
        }
        if (!found) fs_unmatched++;
      });

      clearSearchCache_(ssId);

      return respond({
        success: true,
        ssId: ssId,
        mbId: fs_mbId,
        sdate: fs_sdate,
        edate: fs_edate,
        totalRecords: fs_records.length,
        totalItems: fs_records.reduce(function(a, r) { return a + r.items.length; }, 0),
        updated: fs_updated,
        unmatched: fs_unmatched,
        detail: fs_detail.slice(0, 20)
      });
    } catch (err) {
      return respond({ error: err.message, stack: err.stack });
    }
  }
  // ═══════════════════════════════════════════════════════════════
  // ★ 패스트로 쿠키 갱신 (UI에서 호출) ★
  // ═══════════════════════════════════════════════════════════════
  if (action === 'updateFastroCookie') {
    try {
      var newCookie = e.parameter.cookie || '';
      var targetSsId = e.parameter.targetSsId || 'all';  // 'all' or '1'~'6'

      if (!newCookie) {
        return respond({ error: '쿠키값이 비어있습니다' });
      }

      // PHPSESSID 형식 체크
      if (newCookie.indexOf('PHPSESSID=') < 0) {
        return respond({ error: 'PHPSESSID가 포함되지 않은 쿠키입니다' });
      }

      // hd_pops 자동 추가 (없으면)
      if (newCookie.indexOf('hd_pops_78=') < 0) newCookie += '; hd_pops_78=1';
      if (newCookie.indexOf('hd_pops_77=') < 0) newCookie += '; hd_pops_77=1';

      var props = PropertiesService.getScriptProperties();
      var updated = [];

      if (targetSsId === 'all') {
        // 6개 사업자 모두 갱신
        for (var i = 1; i <= 6; i++) {
          props.setProperty('FASTRO_COOKIE_' + i, newCookie);
          updated.push(String(i));
        }
      } else {
        // 특정 사업자만
        props.setProperty('FASTRO_COOKIE_' + targetSsId, newCookie);
        updated.push(targetSsId);
      }

      // 갱신 직후 검증 - 첫 번째 사업자로 테스트 fetch
      var testSsId = updated[0];
      var testMbId = getFastroMbId_(testSsId);
      var today = new Date();
      var past = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
      var fmt = function(d) {
        return d.getFullYear() + '-' +
               String(d.getMonth()+1).padStart(2,'0') + '-' +
               String(d.getDate()).padStart(2,'0');
      };
      var testResult = fastroFetch_(testSsId, testMbId, fmt(past), fmt(today), 10);

      return respond({
        success: true,
        updated: updated,
        updatedCount: updated.length,
        verified: testResult.ok,
        verifyMessage: testResult.ok
          ? '✅ 쿠키 정상 작동 확인'
          : '⚠ 쿠키 저장됐으나 패스트로 응답 이상: ' + testResult.error,
        cookieLength: newCookie.length
      });
    } catch (err) {
      return respond({ error: err.message });
    }
  }

  // ═══════════════════════════════════════════════════════════════
  // ★ 패스트로 쿠키 상태 확인 (UI에서 호출) ★
  // ═══════════════════════════════════════════════════════════════
  if (action === 'checkFastroStatus') {
    try {
      var props = PropertiesService.getScriptProperties();
      var statusList = [];

      for (var i = 1; i <= 6; i++) {
        var sid = String(i);
        var cookie = props.getProperty('FASTRO_COOKIE_' + i) || '';
        var mbId = getFastroMbId_(sid);

        if (!cookie) {
          statusList.push({
            ssId: sid,
            mbId: mbId,
            hasCookie: false,
            valid: false,
            message: '쿠키 미등록'
          });
          continue;
        }

        // 빠른 검증
        var today = new Date();
        var past = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
        var fmt = function(d) {
          return d.getFullYear() + '-' +
                 String(d.getMonth()+1).padStart(2,'0') + '-' +
                 String(d.getDate()).padStart(2,'0');
        };
        var r = fastroFetch_(sid, mbId, fmt(past), fmt(today), 5);

        statusList.push({
          ssId: sid,
          mbId: mbId,
          hasCookie: true,
          valid: r.ok,
          message: r.ok ? '✅ 정상' : ('❌ ' + r.error),
          cookieLength: cookie.length
        });

        Utilities.sleep(300);  // 패스트로 부하 방지
      }

      return respond({
        success: true,
        statusList: statusList
      });
    } catch (err) {
      return respond({ error: err.message });
    }
  }



  return respond({ error: '알 수 없는 요청' });
}
// ═══════════════════════════════════════════════════════════════
// ★ 기존 시트의 중복 주문번호 일괄 정리 (한 번만 실행하면 됨!)
// ═══════════════════════════════════════════════════════════════
// 같은 주문번호 2개 이상 → 위에서 아래로 -1, -2, -3... 자동 부여
// 이미 -숫자 접미사 있는 행은 건드리지 않음
function deduplicateOrderNos_(ssId) {
  var SS_MAP = {
    '1': '10hfa94OWjpyvF4eK5w2FYhqsIRbfsn8pPlISoj9N9Jg',
    '2': '1u2kIj8-c6YVtv1XOg40-Im1dL09sGiIxehatDKD8Cu0',
    '3': '137pCqtBMsOgX752pVoglwBk3NyK2aXtAy1R40lLJFiM',
    '4': '1AVe0jVZXFLloVEbJHh65OSru8ZbWnXwplHxAmbaXkGE',
    '5': '1IZEnu3o7WcpRUXJP3AcxPQyNOsUj1J2bTSjvpX14oGo',
    '6': '1FCvAfy4AMClQ-0pdJ_AjdwAOqh05Ldh80fTSXknI6Nc'
  };
  var ss;
  if (ssId === '1') {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } else {
    ss = SpreadsheetApp.openById(SS_MAP[ssId]);
  }
  var monthSheets = ss.getSheets().filter(function(s) {
    return /^\d{2}년 \d+월$/.test(s.getName());
  });
  var totalChanged = 0;
  var detail = [];
  monthSheets.forEach(function(sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    var gRange = sheet.getRange(2, 7, lastRow - 1, 1);
    var gVals = gRange.getValues();
    // 1단계: 기본번호(접미사 없음)와 중복카운트 수집
    var baseCount = {};
    var baseFirstRowIdx = {};  // baseNo → 첫 등장 행 인덱스
    var hasSuffix = {};        // baseNo → 이미 -숫자 행이 있는지
    var maxSuffix = {};        // baseNo → 현재 최대 접미사
    for (var i = 0; i < gVals.length; i++) {
      var v = String(gVals[i][0] || '').trim();
      if (!v) continue;
      var sm = v.match(/^(.+)-(\d+)$/);
      if (sm) {
        var base = sm[1];
        var num = parseInt(sm[2], 10);
        hasSuffix[base] = true;
        if (!maxSuffix[base] || num > maxSuffix[base]) maxSuffix[base] = num;
      } else {
        // 접미사 없는 기본번호
        if (!(v in baseCount)) {
          baseCount[v] = 1;
          baseFirstRowIdx[v] = i;
        } else {
          baseCount[v]++;
        }
      }
    }
    // 2단계: 중복(2개 이상)인 기본번호 처리
    var changed = false;
    for (var i = 0; i < gVals.length; i++) {
      var v = String(gVals[i][0] || '').trim();
      if (!v) continue;
      if (/-\d+$/.test(v)) continue;  // 이미 접미사 있으면 스킵
      if (baseCount[v] < 2) continue;  // 중복 아니면 스킵
      // 이 행에 새 접미사 부여
      var base = v;
      if (!(base in maxSuffix)) maxSuffix[base] = 0;
      maxSuffix[base]++;
      var newVal = base + '-' + maxSuffix[base];
      gVals[i][0] = newVal;
      changed = true;
      totalChanged++;
      detail.push(sheet.getName() + ' ' + (i+2) + '행: ' + base + ' → ' + newVal);
    }
    if (changed) gRange.setValues(gVals);
  });
  clearSearchCache_(ssId);
  return { total: totalChanged, detail: detail };
}

// ★ 6개 사업자 모두 일괄 정리 (한 번만 실행)
function deduplicateAllOrderNos() {
  var grandTotal = 0;
  for (var i = 1; i <= 6; i++) {
    try {
      var r = deduplicateOrderNos_(String(i));
      grandTotal += r.total;
      Logger.log('═══════════════════');
      Logger.log('ssId=' + i + ' → 변경된 행: ' + r.total + '개');
      r.detail.slice(0, 20).forEach(function(d) { Logger.log('  ' + d); });
      if (r.detail.length > 20) Logger.log('  ... 외 ' + (r.detail.length - 20) + '개');
    } catch (err) {
      Logger.log('ssId=' + i + ' 오류: ' + err.message);
    }
    Utilities.sleep(500);
  }
  Logger.log('═══════════════════');
  Logger.log('🎉 전체 정리 완료: ' + grandTotal + '개 행 변경');
}

// ★ 한 사업자만 (테스트용)
function deduplicateOrderNos_test1() {
  var r = deduplicateOrderNos_('1');
  Logger.log('ssId=1 → 변경된 행: ' + r.total + '개');
  r.detail.forEach(function(d) { Logger.log('  ' + d); });
}


