/**
 * @OnlyCurrentDoc false
 */

function doGet(e) {
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
    '1': '10hfa94OWjpyvF4eK5w2FYhqsIRbfsn8pPlISoj9N9Jg', // 모다mall
    '2': '1u2kIj8-c6YVtv1XOg40-Im1dL09sGiIxehatDKD8Cu0', // 모다mall2
    '3': '137pCqtBMsOgX752pVoglwBk3NyK2aXtAy1R40lLJFiM', // 모다mall3
    '4': '1AVe0jVZXFLloVEbJHh65OSru8ZbWnXwplHxAmbaXkGE',  // 모다mall4
    '5': '1IZEnu3o7WcpRUXJP3AcxPQyNOsUj1J2bTSjvpX14oGo', // 수기mall
    '6': '1FCvAfy4AMClQ-0pdJ_AjdwAOqh05Ldh80fTSXknI6Nc',  // 불사자(통합)
  };
  const ssId = e.parameter.ssId || '1';
  const targetId = SS_MAP[ssId] || SS_MAP['1'];
  let ss;
  try {
    ss = ssId === '1'
      ? SpreadsheetApp.getActiveSpreadsheet()
      : SpreadsheetApp.openById(targetId);
  } catch(openErr) {
    return respond({ error: '스프레드시트 접근 실패: ' + openErr.message + ' (ID: ' + targetId + ')' });
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
        if (colH.startsWith('010')) {
          phone = colH.replace(/-/g, '');
        } else if (colG.startsWith('010')) {
          phone = colG.replace(/-/g, '');
        }
        if (!phone) continue;

        let source = '';
        let customs = ''; // 통관번호
        for (let j = i + 1; j < Math.min(i + 5, data.length); j++) {
          const hVal = String(data[j][7]).trim();
          const dVal = String(data[j][3]).trim(); // D열
          const eVal = String(data[j][4]).trim(); // E열
          if (hVal && !hVal.startsWith('010') && hVal.length > 0) {
            source = hVal;
          }
          // D열이 "통관부호" 이면 E열이 통관번호
          if (dVal === '통관부호' && eVal) {
            customs = eVal;
          }
        }

        const skipWords = ['주소록 담당자', '사업자', '주문번호', '자동'];
        let msgLines = [];
        let msgStarted = false;

        for (let j = i + 1; j < Math.min(i + 40, data.length); j++) {
          const nextRow = data[j];
          const nextB = String(nextRow[1]).trim();
          const nextC = String(nextRow[2]).trim();

          if (nextB.length >= 10 && /^\d+$/.test(nextB)) break;
          if (skipWords.includes(nextB)) { if (!msgStarted) continue; else break; }

          if (!msgStarted) {
            if (nextB.endsWith('님')) { msgStarted = true; msgLines.push(nextB); }
            continue;
          }

          if (/^010-?\d{4}-?\d{4}$/.test(nextB)) continue;

          if (nextB !== '') {
            if (nextB.startsWith('[옵션]') && nextC !== '') {
              msgLines.push(nextB + '\t' + nextC);
            } else {
              msgLines.push(nextB);
            }
          } else {
            if (!nextC.startsWith('수량') && msgLines.length > 0 && msgLines[msgLines.length - 1] !== '') {
              msgLines.push('');
            }
          }
        }

        while (msgLines.length > 0 && msgLines[msgLines.length - 1] === '') msgLines.pop();

        customers.push({
          idx: customers.length + 1,
          orderNo: colB,
          name: colC,
          phone: phone,
          product: colD,
          option: colE,
          source: source,
          customs: customs,
          message: msgLines.join("\n")
        });
      }
    }
    return respond({ customers: customers });
  }

  if (action === 'scripts') {
    const smsSheet = ss.getSheetByName('문자발송');
    if (!smsSheet) return respond({ error: '문자발송 시트를 찾을 수 없습니다.' });

    const data = smsSheet.getDataRange().getValues();
    const scripts = [];

    for (let i = 0; i < data.length; i++) {
      const colA = String(data[i][0]).trim();
      const colB = String(data[i][1]).trim();

      if (
        colA !== '' && colB !== '' &&
        !/^\d{15,}$/.test(colB) &&
        colA !== '주소록 담당자' && colA !== '사업자' &&
        colA !== '주문번호' && colA !== '고객명' && colA !== '자동'
      ) {
        scripts.push({ idx: scripts.length + 1, title: colA, content: colB });
      }
    }
    return respond({ scripts: scripts });
  }

  if (action === 'saveContact') {
    const name    = e.parameter.name    || ''; // H열: 수령자
    const phone   = e.parameter.phone   || ''; // N열: 전화번호
    const product = e.parameter.product || ''; // I열: 상품명
    const source  = e.parameter.source  || ''; // F열: 판매처 (예: 1-스마트스토어)
    const memo_f  = e.parameter.memo_f  || ''; // F열: 판매처
    const memo_j  = e.parameter.memo_j  || ''; // J열: 옵션명
    const orderNo = e.parameter.orderNo || '';

    if (!name || !phone) return respond({ error: '이름 또는 전화번호가 없습니다.' });

    try {
      // 중복 체크
      try {
        const searchResult = People.People.searchContacts({
          query: phone.replace(/\D/g, ''),
          readMask: 'phoneNumbers,names'
        });
        if (searchResult.results && searchResult.results.length > 0) {
          return respond({ error: 'DUPLICATE' });
        }
      } catch(e) {}

      // 성: F열에서 - 앞 숫자 (예: "1-스마트스토어" → "1")
      let lastName = '';
      const dashIdx = source.indexOf('-');
      if (dashIdx > 0) lastName = source.substring(0, dashIdx).trim();

      // 이름: H열/I열
      const firstName = name + '/' + product;

      // 메모: F열 + J열
      const memo = (memo_f ? memo_f : '')
                 + (memo_j ? (memo_f ? '\n' : '') + memo_j : '');

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

    if (!unipass || !name || !phone) {
      return respond({ result: '⚠️ 데이터없음', msg: '이름, 통관번호, 전화번호가 필요합니다.' });
    }

    if (!/^P[0-9]{12}$/.test(unipass)) {
      return respond({ result: '❌ 형식오류', msg: '올바른 형식: P + 12자리 숫자 (예: P123456789012)' });
    }

    try {
      const url = 'https://sellochomes.co.kr/api/v1/sellerlife/unipass/unipass';

      const callApi = (ph) => {
        const payload = 'persEcm=' + unipass +
                        '&pltxNm=' + encodeURIComponent(ph.name || name) +
                        '&cralTelno=' + ph +
                        '&addrNo=';
        const options = {
          method: 'post',
          contentType: 'application/x-www-form-urlencoded',
          payload: payload,
          muteHttpExceptions: true,
          headers: {
            Cookie: getLoginCookie_(),
            Referer: 'https://sellochomes.co.kr/sellerlife/unipass/'
          }
        };
        const res  = UrlFetchApp.fetch(url, options);
        const text = res.getContentText();
        const json = JSON.parse(text);
        const data = json.data || '';
        const tCntMatch = data.match(/<tCnt>(\d+)<\/tCnt>/);
        const tCnt = tCntMatch ? parseInt(tCntMatch[1]) : 0;
        if (tCnt >= 1) return '✅ 정상';
        if (data.indexOf('존재하지 않습니다') > -1) return '❌ 번호없음';
        if (data.indexOf('불일치') > -1 || data.indexOf('일치하지 않습니다') > -1) return '❌ 불일치';
        return '⚠️ 확인필요';
      };

      let result = callApi(phone);

      // 실패 시 전화번호2로 재시도
      const phone2 = (e.parameter.phone2 || '').replace(/-/g, '');
      if (result !== '✅ 정상' && phone2 && phone2 !== phone && /^010/.test(phone2)) {
        const result2 = callApi(phone2);
        if (result2 === '✅ 정상') result = result2;
      }

      const msgMap = {
        '✅ 정상':    '이름과 통관번호가 일치합니다.',
        '❌ 번호없음': '등록되지 않은 통관번호입니다.',
        '❌ 불일치':  '이름 또는 전화번호가 일치하지 않습니다.',
        '⚠️ 확인필요': '셀러라이프에서 직접 확인이 필요합니다.'
      };

      // 월별 시트 R열에 결과 저장
      const orderNo2 = e.parameter.orderNo || '';
      if (orderNo2) {
        try {
          const now2 = new Date();
          const sheetName2 = String(now2.getFullYear()).slice(2) + '년 ' + (now2.getMonth()+1) + '월';
          let sheet2 = ss.getSheetByName(sheetName2);
          if (!sheet2) {
            const sheets2 = ss.getSheets();
            for (let i = 0; i < sheets2.length; i++) {
              if (/^\d{2}년 \d+월$/.test(sheets2[i].getName())) {
                sheet2 = sheets2[i];
              }
            }
          }
          if (sheet2) {
            const data2 = sheet2.getDataRange().getValues();
            for (let i = 0; i < data2.length; i++) {
              if (String(data2[i][6]).trim() === orderNo2) {
                sheet2.getRange(i + 1, 18).setValue(result); // R열 = 18번째
                break;
              }
            }
          }
        } catch(e2) {}
      }

      return respond({ result: result, msg: msgMap[result] || '' });

    } catch(err) {
      return respond({ result: '⚠️ 오류', msg: err.message });
    }
  }

  // 전체 데이터 한번에 반환
  if (action === 'all') {
    try {
      // 1. 문자발송 시트 고객 목록
      const smsSheet = ss.getSheetByName('문자발송');
      const customers = [];
      const scripts = [];
      const msgMap = {};

      if (smsSheet) {
        const smsData = smsSheet.getDataRange().getValues();

        // 스크립트 수집
        for (let i = 0; i < smsData.length; i++) {
          const colA = String(smsData[i][0]).trim();
          const colB = String(smsData[i][1]).trim();
          if (colA !== '' && colB !== '' && !/^\d{10,}$/.test(colB) &&
              !['주소록 담당자','사업자','주문번호','고객명','자동'].includes(colA)) {
            scripts.push({ idx: scripts.length + 1, title: colA, content: colB });
          }
        }

        // 고객 및 문자내용 수집
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
                if (nB.startsWith('[옵션]') && nC !== '') msgLines.push(nB + '	' + nC);
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

      // 2. 금월 시트
      const now = new Date();
      const sheetName = String(now.getFullYear()).slice(2) + '년 ' + (now.getMonth()+1) + '월';
      let mSheet = ss.getSheetByName(sheetName);
      if (!mSheet) {
        const sheets = ss.getSheets();
        for (let i = 0; i < sheets.length; i++) {
          if (/^\d{2}년 \d+월$/.test(sheets[i].getName())) mSheet = sheets[i];
        }
      }

      const mName = mSheet ? mSheet.getName() : sheetName;
      const monthList = [];
      if (mSheet) {
        const mData   = mSheet.getDataRange().getValues();
        const lastRow = mData.length;
        const allBg   = lastRow > 1 ? mSheet.getRange(2, 1, lastRow-1, 12).getBackgrounds() : [];

        // 월 목록 수집
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
          let phone = String(row[13]).trim().replace(/-/g,'');
          if (!phone || !phone.startsWith('010')) phone = String(row[12]).trim().replace(/-/g,'');
          if (!phone) continue;
          const msgInfo = msgMap[orderNo] || {};
          monthList.push({
            row: i+1, sheetName: mName, orderNo, name, phone,
            product: String(row[8]).trim(),  option:   String(row[9]).trim(),
            qty:     String(row[10]).trim(), customs:  String(row[11]).trim(),
            status:  String(row[0]).trim(),  ship:     String(row[5]).trim(),
            source:  String(row[5]).trim(),  memo2:    String(row[15]).trim(),
            message: msgInfo.message || '',
            // 금액 관련 열
            payAmt:    row[16] || '',  // Q열(17) 결제금액
            settleAmt: row[20] || '',  // U열(21) 정산예정금액
            taoAmt:    row[22] || '',  // W열(23) tao결제액
            shipFee:   row[26] || '',  // AA열(27) 배송비
            expShip:   row[27] || '',  // AB열(28) 예상배송비
            addShip:   row[28] || '',  // AC열(29) 추가배송비
            profitRate: row[32] || '', // AG열(33) 이익율
            marginRate: row[33] || '', // AH열(34) 마진율
            gBg: allBg[i-1] ? (allBg[i-1][6]||'').toLowerCase() : '',
            hBg: allBg[i-1] ? (allBg[i-1][7]||'').toLowerCase() : '',
            iBg: allBg[i-1] ? (allBg[i-1][8]||'').toLowerCase() : '',
            lBg: allBg[i-1] ? (allBg[i-1][11]||'').toLowerCase() : '',
            prodNote: mSheet.getRange(i+1, 9).getNote() || '',
            optNote:  mSheet.getRange(i+1, 10).getNote() || ''
          });
        }
        monthList.reverse();
        return respond({ customers, scripts, monthList, sheetName: mName, monthNames });
      }

      return respond({ customers, scripts, monthList, sheetName, monthNames: [] });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // 금월 시트 고객 리스트 (문자내용은 문자발송 시트에서 매칭)
  if (action === 'monthList') {
    try {
      const now = new Date();
      const sheetName = String(now.getFullYear()).slice(2) + '년 ' + (now.getMonth()+1) + '월';
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        const sheets = ss.getSheets();
        for (let i = 0; i < sheets.length; i++) {
          if (/^\d{2}년 \d+월$/.test(sheets[i].getName())) sheet = sheets[i];
        }
      }
      if (!sheet) return respond({ error: '월별 시트를 찾을 수 없습니다.' });

      // 문자발송 시트에서 주문번호별 문자내용 미리 수집
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

            const skipWords = ['주소록 담당자', '사업자', '주문번호', '자동'];
            let msgLines = [];
            let msgStarted = false;

            for (let j = i + 1; j < Math.min(i + 40, smsData.length); j++) {
              const nextB = String(smsData[j][1]).trim();
              const nextC = String(smsData[j][2]).trim();
              if (nextB.length >= 10 && /^\d+$/.test(nextB)) break;
              if (skipWords.includes(nextB)) { if (!msgStarted) continue; else break; }
              if (!msgStarted) {
                if (nextB.endsWith('님')) { msgStarted = true; msgLines.push(nextB); }
                continue;
              }
              if (/^010-?\d{4}-?\d{4}$/.test(nextB)) continue;
              if (nextB !== '') {
                if (nextB.startsWith('[옵션]') && nextC !== '') msgLines.push(nextB + '	' + nextC);
                else msgLines.push(nextB);
              } else {
                if (!nextC.startsWith('수량') && msgLines.length > 0 && msgLines[msgLines.length-1] !== '') msgLines.push('');
              }
            }
            while (msgLines.length > 0 && msgLines[msgLines.length-1] === '') msgLines.pop();

            // 주문번호를 key로 저장 (문자내용 + 소스 정보)
            let source = '';
            for (let j = i + 1; j < Math.min(i + 5, smsData.length); j++) {
              const hVal = String(smsData[j][7]).trim();
              if (hVal && !hVal.startsWith('010') && hVal.length > 0) { source = hVal; break; }
            }

            msgMap[colB] = {
              message: msgLines.join('\n'),
              phone: phone,
              product: String(smsData[i][3]).trim(),
              option: String(smsData[i][4]).trim(),
              source: source
            };
          }
        }
      }

      // 금월 시트 데이터 읽기
      const data = sheet.getDataRange().getValues();
      const list = [];

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const orderNo = String(row[6]).trim();  // G열: 주문번호
        const name    = String(row[7]).trim();  // H열: 수령자
        const product = String(row[8]).trim();  // I열: 상품명
        const option  = String(row[9]).trim();  // J열: 옵션명
        const qty     = String(row[10]).trim(); // K열: 수량
        const customs = String(row[11]).trim(); // L열: 통관번호
        const status  = String(row[0]).trim();  // A열: 상태
        const ship    = String(row[5]).trim();  // F열: 운송방법 (판매처)
        const delivery= String(row[1]).trim();  // B열: 송장번호
        const source  = String(row[4]).trim();  // E열: 판매처
        const memo2   = String(row[15]).trim(); // P열: 추가메모

        // N열: 전화번호 (index 13)
        let phone = String(row[13]).trim().replace(/-/g, '');
        if (!phone || !phone.startsWith('010')) {
          // M열 fallback (index 12)
          phone = String(row[12]).trim().replace(/-/g, '');
        }

        if (!orderNo || !name || !phone) continue;

        // 문자발송 시트에서 문자내용 매칭
        const msgInfo = msgMap[orderNo] || {};

        list.push({
          row: i + 1,
          orderNo, name, phone,
          product: msgInfo.product || product,
          option:  msgInfo.option  || option,
          source:  ship || msgInfo.source || source, // F열 판매처 우선
          customs, status, qty, ship, delivery, memo2,
          message: msgInfo.message || ''
        });
      }

      // 최신순 (역순)
      list.reverse();

      return respond({ list: list, sheetName: sheetName, total: list.length });

    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // 경동 선택 - A열
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
      const now = new Date();
      const sheetName = String(now.getFullYear()).slice(2) + '년 ' + (now.getMonth()+1) + '월';
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        const sheets = ss.getSheets();
        for (let i = 0; i < sheets.length; i++) {
          if (/^\d{2}년 \d+월$/.test(sheets[i].getName())) sheet = sheets[i];
        }
      }
      if (!sheet) return respond({ error: '월별 시트를 찾을 수 없습니다.' });

      const data = sheet.getDataRange().getValues();
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][6]).trim() === orderNo) {
          const cell = sheet.getRange(i + 1, 1); // A열
          cell.setValue(style.text);
          cell.setBackground(style.bg);
          cell.setFontColor(style.color);
          cell.setFontWeight(style.text ? 'bold' : 'normal');
          return respond({ success: true });
        }
      }
      return respond({ error: '주문번호를 찾을 수 없습니다.' });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // LCL 진행여부 - D열
  if (action === 'lclToggle') {
    const orderNo = e.parameter.orderNo || '';
    const cancel  = e.parameter.cancel === '1';
    if (!orderNo) return respond({ error: '주문번호 없음' });

    try {
      const now = new Date();
      const sheetName = String(now.getFullYear()).slice(2) + '년 ' + (now.getMonth()+1) + '월';
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        const sheets = ss.getSheets();
        for (let i = 0; i < sheets.length; i++) {
          if (/^\d{2}년 \d+월$/.test(sheets[i].getName())) sheet = sheets[i];
        }
      }
      if (!sheet) return respond({ error: '월별 시트를 찾을 수 없습니다.' });

      const data = sheet.getDataRange().getValues();
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][6]).trim() === orderNo) {
          const cell = sheet.getRange(i + 1, 4); // D열 = 4번째
          if (cancel) {
            cell.setValue('');
            cell.setBackground('#FFFFFF');
          } else {
            cell.setValue('LCL출고');
            cell.setBackground('#880E4F'); // 진한 자주색
            cell.setFontColor('#FFFFFF');
          }
          return respond({ success: true });
        }
      }
      return respond({ error: '주문번호를 찾을 수 없습니다.' });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // 전체 시트 검색 (캐시 활용)
  if (action === 'searchAll') {
    const query = (e.parameter.q || '').toLowerCase().trim();
    if (!query) return respond({ list: [] });

    try {
      const cache = CacheService.getScriptCache();
      const cacheKey = 'search_index_v1_ss' + ssId;
      let index = null;

      // 캐시에서 인덱스 읽기 (최대 6시간 유지)
      const cached = cache.get(cacheKey);
      if (cached) {
        try { index = JSON.parse(cached); } catch(e) { index = null; }
      }

      // 캐시 없으면 전체 시트 읽어서 인덱스 생성
      if (!index) {
        index = [];
        const allSheets = ss.getSheets();
        const monthSheets = allSheets.filter(s => /^\d{2}년 \d+월$/.test(s.getName()));
        monthSheets.sort((a, b) => {
          const pa = a.getName().match(/(\d+)년 (\d+)월/);
          const pb = b.getName().match(/(\d+)년 (\d+)월/);
          if (!pa || !pb) return 0;
          return (Number(pb[1])*100+Number(pb[2])) - (Number(pa[1])*100+Number(pa[2]));
        });

        for (const mSheet of monthSheets) {
          const mName = mSheet.getName();
          const mData = mSheet.getDataRange().getValues();
          const lastRow = mData.length;
          const allBg = lastRow > 1 ? mSheet.getRange(2, 1, lastRow-1, 12).getBackgrounds() : [];

          for (let i = 1; i < lastRow; i++) {
            const row = mData[i];
            const name    = String(row[7]).trim();
            const orderNo = String(row[6]).trim();
            if (!orderNo || !name) continue;
            let phone = String(row[13]).trim().replace(/-/g,'') || String(row[12]).trim().replace(/-/g,'');
            if (!phone) phone = orderNo.replace(/\D/g,'').slice(-11);
            index.push({
              row: i+1, sheetName: mName, orderNo, name, phone,
              product: String(row[8]).trim(),
              option:  String(row[9]).trim(),
              qty:     String(row[10]).trim(),
              customs: String(row[11]).trim(),
              status:  String(row[0]).trim(),
              ship:    String(row[5]).trim(),
              source:  String(row[5]).trim(),
              memo2:   String(row[15]).trim(),
              message: '',
              gBg: allBg[i-1] ? (allBg[i-1][6]||'').toLowerCase() : '',
              hBg: allBg[i-1] ? (allBg[i-1][7]||'').toLowerCase() : '',
              iBg: allBg[i-1] ? (allBg[i-1][8]||'').toLowerCase() : '',
              lBg: allBg[i-1] ? (allBg[i-1][11]||'').toLowerCase() : ''
            });
          }
        }
        // 캐시에 저장 (21600초 = 6시간)
        try { cache.put(cacheKey, JSON.stringify(index), 21600); } catch(e) {}
      }

      // 검색어로 필터링
      const q = query.replace(/-/g,'');
      const list = index.filter(c =>
        c.name.toLowerCase().includes(query) ||
        c.phone.includes(q) ||
        c.orderNo.includes(query) ||
        c.product.toLowerCase().includes(query) ||
        c.option.toLowerCase().includes(query)
      ).slice(0, 100);

      return respond({ list });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // 캐시 초기화 (데이터 변경 시 호출)
  if (action === 'clearCache') {
    try {
      CacheService.getScriptCache().remove('search_index_v1_ss' + ssId);
      return respond({ success: true });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // 주문취소 - A~Q열 연한빨간색 + 특정 열 내용 삭제
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
            // A~Q열(1~17) 연한빨간색
            mSheet.getRange(i+1, 1, 1, 17).setBackground('#ea9999');
            // Q열(17), U열(21), W열(23), AA열(27), AB열(28), AC열(29) 내용 삭제
            mSheet.getRange(i+1, 17).clearContent();
            mSheet.getRange(i+1, 21).clearContent();
            mSheet.getRange(i+1, 23).clearContent();
            mSheet.getRange(i+1, 27).clearContent();
            mSheet.getRange(i+1, 28).clearContent();
            mSheet.getRange(i+1, 29).clearContent();
          } else {
            // 취소 해제 - A~Q열 흰색으로 복원
            mSheet.getRange(i+1, 1, 1, 17).setBackground('#ffffff');
          }
        }
      }
      // 캐시 초기화
      CacheService.getScriptCache().remove('search_index_v1_ss' + ssId);
      return respond({ success: true, found });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // 금액 필드 저장
  // 금액 전체 한번에 저장 (빠름)
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
            // 저장 후 해당 행의 최신값(이익/마진율 포함) 반환
            SpreadsheetApp.flush(); // 수식 강제 계산
            const updRow = mSheet.getRange(i+1, 1, 1, 35).getValues()[0];
            return respond({
              success: true,
              sheet: mSheet.getName(),
              row: i+1,
              profitRate: updRow[32] || '', // AG열(33) 이익
              marginRate: updRow[33] || ''  // AH열(34) 마진율
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
      // 시트명이 있으면 해당 시트 우선 검색
      if (sheetName) {
        monthSheets.sort((a, b) => a.getName() === sheetName ? -1 : b.getName() === sheetName ? 1 : 0);
      }
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
      // A~P열 + Q열(결제금액) + U열(정산예정)
      // A열 기준 마지막 행 찾기 (텍스트 있는 마지막 행 다음)
      const aVals = mSheet.getRange('A:A').getValues();
      let lastRow = 1;
      for (let i = aVals.length - 1; i >= 0; i--) {
        if (aVals[i][0] !== '' && aVals[i][0] !== null) {
          lastRow = i + 2; // 다음 행
          break;
        }
      }
      const cols = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P'];
      // A~P열 입력 (1~16번 컬럼)
      for (let i = 0; i < cols.length; i++) {
        const val = e.parameter[cols[i]] || '';
        if (val) mSheet.getRange(lastRow, i+1).setValue(val);
      }
      // Q열(17번) - 결제금액 (엑셀 T열)
      const qVal = e.parameter['Q'] || '';
      if (qVal) {
        const qNum = parseFloat(qVal.replace(/,/g,''));
        mSheet.getRange(lastRow, 17).setValue(isNaN(qNum) ? qVal : qNum);
      }
      // U열(21번) - 정산예정 (엑셀 U열)
      const uVal = e.parameter['U'] || '';
      if (uVal) {
        const uNum = parseFloat(uVal.replace(/,/g,''));
        mSheet.getRange(lastRow, 21).setValue(isNaN(uNum) ? uVal : uNum);
      }
      return respond({ success: true, row: lastRow });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // 고객 정보 수정 (H=이름, M=휴대폰1, N=휴대폰2, L=통관부호, P=주소)
  if (action === 'updateInfo') {
    const orderNo   = e.parameter.orderNo || '';
    const sheetName = e.parameter.sheetName || '';
    const name      = e.parameter.name    || '';
    const phone1    = e.parameter.phone1  || '';
    const phone2    = e.parameter.phone2  || '';
    const customs   = e.parameter.customs || '';
    const addr      = e.parameter.addr    || '';
    if (!orderNo) return respond({ error: '주문번호 없음' });
    try {
      const allSheets = ss.getSheets();
      let monthSheets = allSheets.filter(s => /^\d{2}년 \d+월$/.test(s.getName()));
      if (sheetName) monthSheets.sort((a,b) => a.getName()===sheetName?-1:1);
      for (const mSheet of monthSheets) {
        const data = mSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][6]).trim() !== String(orderNo).trim()) continue;
          if (name)    mSheet.getRange(i+1, 8).setValue(name);    // H열
          if (phone1)  mSheet.getRange(i+1, 13).setValue(phone1); // M열
          if (phone2)  mSheet.getRange(i+1, 14).setValue(phone2); // N열
          if (customs) mSheet.getRange(i+1, 12).setValue(customs);// L열
          if (addr)    mSheet.getRange(i+1, 16).setValue(addr);   // P열
          return respond({ success: true, sheet: mSheet.getName(), row: i+1 });
        }
      }
      return respond({ error: '주문번호를 찾을 수 없습니다.' });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // 특정 월 데이터 로드
  if (action === 'getMonthData') {
    const targetSheet = e.parameter.sheetName || '';
    if (!targetSheet) return respond({ error: '시트명 없음' });
    const mSheet = ss.getSheetByName(targetSheet);
    if (!mSheet) return respond({ error: '시트를 찾을 수 없습니다.' });

    const list = [];
    const mData   = mSheet.getDataRange().getValues();
    const lastRow = mData.length;
    const allBg   = lastRow > 1 ? mSheet.getRange(2, 1, lastRow-1, 12).getBackgrounds() : [];

    for (let i = 1; i < lastRow; i++) {
      const row = mData[i];
      const orderNo = String(row[6]).trim();
      const name    = String(row[7]).trim();
      if (!orderNo || !name) continue;
      let phone = String(row[13]).trim().replace(/-/g,'');
      if (!phone) phone = String(row[12]).trim().replace(/-/g,'');
      // 전화번호 없으면 주문번호 뒷자리로 대체 (건너뛰지 않음)
      if (!phone) phone = orderNo.replace(/\D/g,'').slice(-11);
      list.push({
        row: i+1, sheetName: targetSheet, orderNo, name, phone,
        product: String(row[8]).trim(),  option:  String(row[9]).trim(),
        qty:     String(row[10]).trim(), customs: String(row[11]).trim(),
        status:  String(row[0]).trim(),  ship:    String(row[5]).trim(),
        source:  String(row[5]).trim(),  memo2:   String(row[15]).trim(),
        message: '',
        gBg: allBg[i-1] ? (allBg[i-1][6]||'').toLowerCase() : '',
        hBg: allBg[i-1] ? (allBg[i-1][7]||'').toLowerCase() : '',
        iBg: allBg[i-1] ? (allBg[i-1][8]||'').toLowerCase() : '',
        lBg: allBg[i-1] ? (allBg[i-1][11]||'').toLowerCase() : ''
      });
    }
    return respond({ list, sheetName: targetSheet });
  }

  // 배대지 운송방법 - F열 색칠
  if (action === 'shipMethod') {
    const orderNo = e.parameter.orderNo || '';
    const type    = e.parameter.type    || 'none';
    if (!orderNo) return respond({ error: '주문번호 없음' });

    const colorMap = {
      'fast'     : '#FF6D00', // 주황 - 패스트로 항공
      'joeun-air': '#8BC34A', // 연두 - 조은직구 항공
      'joeun-sea': '#9C27B0', // 보라 - 조은직구 위해
      'none'     : '#FFFFFF'  // 흰색 - 취소
    };
    const color = colorMap[type] || '#FFFFFF';

    try {
      const now = new Date();
      const sheetName = String(now.getFullYear()).slice(2) + '년 ' + (now.getMonth()+1) + '월';
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        const sheets = ss.getSheets();
        for (let i = 0; i < sheets.length; i++) {
          if (/^\d{2}년 \d+월$/.test(sheets[i].getName())) sheet = sheets[i];
        }
      }
      if (!sheet) return respond({ error: '월별 시트를 찾을 수 없습니다.' });

      const data = sheet.getDataRange().getValues();
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][6]).trim() === orderNo) {
          sheet.getRange(i + 1, 6).setBackground(color); // F열 = 6번째
          return respond({ success: true });
        }
      }
      return respond({ error: '주문번호를 찾을 수 없습니다.' });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // 배대지신청서 - H열 노란색 + C열 신청서번호
  if (action === 'baejiRequest') {
    const orderNo = e.parameter.orderNo || '';
    const cancel  = e.parameter.cancel === '1';
    const baejiNo = e.parameter.baejiNo || '';
    if (!orderNo) return respond({ error: '주문번호 없음' });

    try {
      const now = new Date();
      const sheetName = String(now.getFullYear()).slice(2) + '년 ' + (now.getMonth()+1) + '월';
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        const sheets = ss.getSheets();
        for (let i = 0; i < sheets.length; i++) {
          if (/^\d{2}년 \d+월$/.test(sheets[i].getName())) sheet = sheets[i];
        }
      }
      if (!sheet) return respond({ error: '월별 시트를 찾을 수 없습니다.' });

      const data = sheet.getDataRange().getValues();
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][6]).trim() === orderNo) {
          // H열(8번째) - 신청시 노란색, 취소시 흰색
          sheet.getRange(i + 1, 8).setBackground(cancel ? '#FFFFFF' : '#FFFF00');
          // C열(3번째) - 신청서번호 기입 + 노란색
          const cCell = sheet.getRange(i + 1, 3);
          cCell.setValue(cancel ? '' : baejiNo);
          cCell.setBackground(cancel ? '#FFFFFF' : '#FFFF00');
          return respond({ success: true });
        }
      }
      return respond({ error: '주문번호를 찾을 수 없습니다.' });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // 배송방법 선택 - AJ열 기입 + G열 하늘색
  if (action === 'deliverySelect') {
    const orderNo = e.parameter.orderNo || '';
    const type    = e.parameter.type    || '';
    if (!orderNo) return respond({ error: '주문번호 없음' });

    try {
      const now = new Date();
      const sheetName = String(now.getFullYear()).slice(2) + '년 ' + (now.getMonth()+1) + '월';
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        const sheets = ss.getSheets();
        for (let i = 0; i < sheets.length; i++) {
          if (/^\d{2}년 \d+월$/.test(sheets[i].getName())) sheet = sheets[i];
        }
      }
      if (!sheet) return respond({ error: '월별 시트를 찾을 수 없습니다.' });

      const data = sheet.getDataRange().getValues();
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][6]).trim() === orderNo) {
          // AJ열(36번째) = 배송방법
          sheet.getRange(i + 1, 36).setValue(type);
          // G열(7번째) = 하늘색 or 흰색
          sheet.getRange(i + 1, 7).setBackground(type ? '#87CEEB' : '#FFFFFF');
          return respond({ success: true });
        }
      }
      return respond({ error: '주문번호를 찾을 수 없습니다.' });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // 주문 토글 - I열 노란색/흰색
  if (action === 'orderToggle') {
    const orderNo = e.parameter.orderNo || '';
    const cancel  = e.parameter.cancel === '1';
    if (!orderNo) return respond({ error: '주문번호 없음' });

    try {
      const now = new Date();
      const sheetName = String(now.getFullYear()).slice(2) + '년 ' + (now.getMonth()+1) + '월';
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        const sheets = ss.getSheets();
        for (let i = 0; i < sheets.length; i++) {
          if (/^\d{2}년 \d+월$/.test(sheets[i].getName())) sheet = sheets[i];
        }
      }
      if (!sheet) return respond({ error: '월별 시트를 찾을 수 없습니다.' });

      const data = sheet.getDataRange().getValues();
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][6]).trim() === orderNo) {
          sheet.getRange(i + 1, 9).setBackground(cancel ? '#FFFFFF' : '#FFFF00'); // I열
          return respond({ success: true });
        }
      }
      return respond({ error: '주문번호를 찾을 수 없습니다.' });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // 주문요청 - L열 주황색
  if (action === 'requestOrder') {
    const orderNo = e.parameter.orderNo || '';
    if (!orderNo) return respond({ error: '주문번호 없음' });

    try {
      const now = new Date();
      const sheetName = String(now.getFullYear()).slice(2) + '년 ' + (now.getMonth()+1) + '월';
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        const sheets = ss.getSheets();
        for (let i = 0; i < sheets.length; i++) {
          if (/^\d{2}년 \d+월$/.test(sheets[i].getName())) sheet = sheets[i];
        }
      }
      if (!sheet) return respond({ error: '월별 시트를 찾을 수 없습니다.' });

      const cancel = e.parameter.cancel === '1';
      const data = sheet.getDataRange().getValues();
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][6]).trim() === orderNo) {
          // L열(12번째) - 요청시 주황색, 취소시 흰색
          sheet.getRange(i + 1, 12).setBackground(cancel ? '#FFFFFF' : '#FF6D00');
          return respond({ success: true });
        }
      }
      return respond({ error: '주문번호를 찾을 수 없습니다.' });
    } catch(err) {
      return respond({ error: err.message });
    }
  }

  // 상품명/옵션명/수량 저장 - I,J,K열
  if (action === 'saveDetail') {
    const orderNo = e.parameter.orderNo || '';
    const product = e.parameter.product || '';
    const option  = e.parameter.option  || '';
    const qty     = e.parameter.qty     || '';
    if (!orderNo) return respond({ error: '주문번호 없음' });

    try {
      const now = new Date();
      const sheetName = String(now.getFullYear()).slice(2) + '년 ' + (now.getMonth()+1) + '월';
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        const sheets = ss.getSheets();
        for (let i = 0; i < sheets.length; i++) {
          if (/^\d{2}년 \d+월$/.test(sheets[i].getName())) sheet = sheets[i];
        }
      }
      if (!sheet) return respond({ error: '월별 시트를 찾을 수 없습니다.' });

      const data = sheet.getDataRange().getValues();
      const prodNote = e.parameter.prodNote || '';
      const optNote  = e.parameter.optNote  || '';
      let updated = 0;
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][6]).trim() === orderNo) {
          if (product !== undefined) sheet.getRange(i + 1, 9).setValue(product);   // I열
          if (option  !== undefined) sheet.getRange(i + 1, 10).setValue(option);   // J열
          if (qty)                   sheet.getRange(i + 1, 11).setValue(Number(qty)); // K열
          sheet.getRange(i + 1, 9).setNote(prodNote);   // I열 메모
          sheet.getRange(i + 1, 10).setNote(optNote);   // J열 메모
          updated++;
        }
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
      // 사업자별 스프레드시트 선택
  const SS_MAP = {
    '1': '10hfa94OWjpyvF4eK5w2FYhqsIRbfsn8pPlISoj9N9Jg', // 모다mall
    '2': '1u2kIj8-c6YVtv1XOg40-Im1dL09sGiIxehatDKD8Cu0', // 모다mall2
    '3': '137pCqtBMsOgX752pVoglwBk3NyK2aXtAy1R40lLJFiM', // 모다mall3
    '4': '1AVe0jVZXFLloVEbJHh65OSru8ZbWnXwplHxAmbaXkGE',  // 모다mall4
    '5': '1IZEnu3o7WcpRUXJP3AcxPQyNOsUj1J2bTSjvpX14oGo', // 수기mall
    '6': '1FCvAfy4AMClQ-0pdJ_AjdwAOqh05Ldh80fTSXknI6Nc',  // 불사자(통합)
  };
  const ssId = e.parameter.ssId || '1';
  const targetId = SS_MAP[ssId] || SS_MAP['1'];
  let ss;
  try {
    ss = ssId === '1'
      ? SpreadsheetApp.getActiveSpreadsheet()
      : SpreadsheetApp.openById(targetId);
  } catch(openErr) {
    return respond({ error: '스프레드시트 접근 실패: ' + openErr.message + ' (ID: ' + targetId + ')' });
  }

      // 현재 월 시트 찾기 (예: "26년 5월")
      const now = new Date();
      const sheetName = String(now.getFullYear()).slice(2) + '년 ' + (now.getMonth()+1) + '월';
      let sheet = ss.getSheetByName(sheetName);

      // 못 찾으면 모든 시트에서 주문번호 검색
      if (!sheet) {
        const sheets = ss.getSheets();
        for (let i = 0; i < sheets.length; i++) {
          const name = sheets[i].getName();
          if (/^\d{2}년 \d+월$/.test(name)) {
            sheet = sheets[i];
          }
        }
      }

      if (!sheet) return respond({ error: '월별 시트를 찾을 수 없습니다.' });

      const data = sheet.getDataRange().getValues();
      let highlighted = 0;

      for (let i = 0; i < data.length; i++) {
        // G열(index 6)에서만 주문번호 찾기
        const gVal = String(data[i][6]).trim();
        if (gVal === orderNo) {
          sheet.getRange(i + 1, 7).setBackground('#FFFF00');
          highlighted++;
          break; // 한 개만 색칠
        }
      }

      if (highlighted > 0) {
        return respond({ success: true, msg: sheetName + ' 시트 ' + highlighted + '행 색칠 완료' });
      } else {
        return respond({ error: '주문번호를 찾을 수 없습니다: ' + orderNo });
      }

    } catch(err) {
      return respond({ error: err.message });
    }
  }

  return respond({ error: '알 수 없는 요청' });
}

function getLoginCookie_() {
  return 'connect.sid=s%3Agu4c_Li3M-EbBV0g-b6bnWdTKDUkWM91.gn4BvUAq50TAFydPCKKaOhb1qIW8fWiaO1DJQ5xeCTw';
}
