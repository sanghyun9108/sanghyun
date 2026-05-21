const { google } = require('googleapis');

const SHEET_PERF = '성과데이터';
const SHEET_KP   = 'KP데이터';
const PERF_HEADERS = ['저장일시','스토어','날짜','상품명','상품코드','광고명','광고비','광고매출','ROAS','노출수','클릭수','클릭률','카트수','카트전환율','구매수','구매전환율'];
const KP_HEADERS   = ['저장일시','스토어','날짜','입찰번호','타입','키워드','전시일자','상품번호','판매상태','타이틀','입찰가순위','입찰가','필요Q캐시','시작일','현재상태'];

function getAuth() {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_KEY);
  return new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
}

function parseNum(v) {
  if (v === '' || v === null || v === undefined) return '';
  const n = parseFloat(String(v).replace(/,/g, ''));
  return isNaN(n) ? v : n;
}

async function ensureSheet(sheets, spreadsheetId, sheetName, headers, color) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  const exists = meta.data.sheets.some(s => s.properties.title === sheetName);
  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{
          addSheet: { properties: { title: sheetName } }
        }]
      }
    });
    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: `${sheetName}!A1`,
      valueInputOption: 'RAW',
      requestBody: { values: [headers] }
    });
  }
}

async function deleteExistingRows(sheets, spreadsheetId, sheetName, store, date) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${sheetName}!B:C`,
  });
  const rows = res.data.values || [];
  const sheetsMeta = await sheets.spreadsheets.get({ spreadsheetId });
  const sheetObj = sheetsMeta.data.sheets.find(s => s.properties.title === sheetName);
  if (!sheetObj) return;
  const sheetId = sheetObj.properties.sheetId;
  const toDelete = [];
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === store && String(rows[i][1]).slice(0,10) === date) {
      toDelete.push(i + 1);
    }
  }
  if (toDelete.length === 0) return;
  const requests = toDelete.map(rowIdx => ({
    deleteDimension: {
      range: { sheetId, dimension: 'ROWS', startIndex: rowIdx - 1, endIndex: rowIdx }
    }
  }));
  await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests } });
}

module.exports = async function handler(req, res) {
  // CORS 헤더
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  const spreadsheetId = process.env.SPREADSHEET_ID;
  if (!spreadsheetId) {
    return res.status(500).json({ error: 'SPREADSHEET_ID 환경변수가 없습니다.' });
  }

  try {
    const auth = getAuth();
    const sheets = google.sheets({ version: 'v4', auth });

    // ── GET: 데이터 읽기 ──
    if (req.method === 'GET') {
      const result = { perf: [], kp: [] };

      try {
        const perfRes = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: `${SHEET_PERF}!A2:P`,
        });
        (perfRes.data.values || []).forEach(r => {
          if (!r[0]) return;
          result.perf.push({
            저장일시:r[0], 스토어:r[1], 날짜:r[2], 상품명:r[3], 상품코드:r[4], 광고명:r[5],
            '광고비 (Qcash)':r[6], '광고 매출':r[7], ROAS:r[8], 노출수:r[9],
            '클릭수(PV)':r[10], 클릭률:r[11], 카트수:r[12], '카트 전환율':r[13], 구매수:r[14], '구매 전환율':r[15]
          });
        });
      } catch(e) {}

      try {
        const kpRes = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: `${SHEET_KP}!A2:O`,
        });
        (kpRes.data.values || []).forEach(r => {
          if (!r[0]) return;
          result.kp.push({
            저장일시:r[0], 스토어:r[1], 날짜:r[2], 입찰번호:r[3], 타입:r[4],
            '키워드/카테고리':r[5], '전시 일자':r[6], '상품번호/기획전번호':r[7],
            판매상태:r[8], 타이틀:r[9], '입찰가 순위':r[10], 입찰가:r[11],
            '필요 Q캐시':r[12], 시작일:r[13], 현재상태:r[14]
          });
        });
      } catch(e) {}

      return res.status(200).json(result);
    }

    // ── POST: 데이터 저장 ──
    if (req.method === 'POST') {
      const { store, date, perf, kp } = req.body;
      const now = new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Tokyo' });

      if (perf && perf.length > 0) {
        await ensureSheet(sheets, spreadsheetId, SHEET_PERF, PERF_HEADERS, '#1D9E75');
        await deleteExistingRows(sheets, spreadsheetId, SHEET_PERF, store, date);
        const rows = perf.map(r => [
          now, store, date,
          r['상품명']||'', r['상품코드']||'', r['광고명']||'',
          parseNum(r['광고비 (Qcash)']), parseNum(r['광고 매출']), parseNum(r['ROAS']),
          parseNum(r['노출수']), parseNum(r['클릭수(PV)']), parseNum(r['클릭률']),
          parseNum(r['카트수']), parseNum(r['카트 전환율']), parseNum(r['구매수']), parseNum(r['구매 전환율'])
        ]);
        await sheets.spreadsheets.values.append({
          spreadsheetId,
          range: `${SHEET_PERF}!A1`,
          valueInputOption: 'RAW',
          requestBody: { values: rows }
        });
      }

      if (kp && kp.length > 0) {
        await ensureSheet(sheets, spreadsheetId, SHEET_KP, KP_HEADERS, '#7F77DD');
        await deleteExistingRows(sheets, spreadsheetId, SHEET_KP, store, date);
        const rows = kp.map(r => [
          now, store, date,
          r['입찰번호']||'', r['타입']||'', r['키워드/카테고리']||'',
          r['전시 일자']||'', r['상품번호/기획전번호']||'', r['판매상태']||'', r['타이틀']||'',
          parseNum(r['입찰가 순위']), parseNum(r['입찰가']), parseNum(r['필요 Q캐시']),
          r['시작일']||'', r['현재상태']||''
        ]);
        await sheets.spreadsheets.values.append({
          spreadsheetId,
          range: `${SHEET_KP}!A1`,
          valueInputOption: 'RAW',
          requestBody: { values: rows }
        });
      }

      return res.status(200).json({ success: true, store, date });
    }

    return res.status(405).json({ error: 'Method not allowed' });

  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message });
  }
};
