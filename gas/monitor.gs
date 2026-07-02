// =====================================================
// 무역구제 모니터링 시스템 v3 (통합본)
// trade_monitor_integrated.gs
//
// 기존 monitor_v2.gs + trade_news_monitor.gs 통합
// 매일 KST 오전 6시 단일 이메일 발송
// =====================================================

var CONFIG_V2 = {
  EMAIL_RECIPIENTS: [
    "rlawnsgus0613@naver.com"
    ],
  SHEET_NAME:            "인도+한국+미국 케이스",
  SEND_EMAIL_IF_EMPTY:   false,
  BRAZIL_MEASURES_SHEET: "브라질 DECOM 조치",
  BRAZIL_INVEST_SHEET:   "브라질 DECOM 조사"
};

var IN_FEEDS = [
  { url: "https://www.dgtr.gov.in/en/anti-dumping-investigation-in-india",  country: "IN", agency: "DGTR", label: "IN-DGTR-AD" },
  { url: "https://www.dgtr.gov.in/en/countervailing-duty-investigation",     country: "IN", agency: "DGTR", label: "IN-DGTR-CVD" },
  { url: "https://www.dgtr.gov.in/en/safe-guard-investigation-in-india",     country: "IN", agency: "DGTR", label: "IN-DGTR-SG" }
];

var US_URL        = "https://www.trade.gov/ec-adcvd-case-announcements";
var CBP_EAPA_URL  = "https://www.cbp.gov/trade/eapa/notices-action";
var CSMS_JSON_URL = "https://content.govdelivery.com/accounts/USDHSCBP/widgets/USDHSCBP_WIDGET_2.json";

var EU_TRON_LIST_URL    = "https://tron.trade.ec.europa.eu/investigations/api/eucase/list/ongoing";
var EU_TRON_DETAIL_URL  = "https://tron.trade.ec.europa.eu/investigations/api/eucase/details/";
var EU_TRACK_SHEET_NAME = "EU 케이스 추적";

var BRAZIL_MEASURES_URL = "https://www.gov.br/mdic/pt-br/assuntos/comercio-exterior/defesa-comercial-e-interesse-publico/medidas-em-vigor/medidas-em-vigor/medidas-de-defesa-comercial-em-vigor";
var BRAZIL_INVEST_URL   = "https://www.gov.br/mdic/pt-br/assuntos/comercio-exterior/defesa-comercial-e-interesse-publico/investigacoes/investigacoes-de-defesa-comercial/investigacoes-de-defesa-comercial-em-curso";
var AU_ADC_URL = "https://www.industry.gov.au/anti-dumping-commission/current-cases-and-electronic-public-record-epr";
var NEWS_SHEET_NAME  = "글로벌 무역구제 뉴스";
var TRUTH_SOCIAL_URL = "https://ix.cnn.io/data/truth-social/truth_archive.json";

// =====================================================
// ── 공통 유틸 ─────────────────────────────────────────
// =====================================================

function makeHash(str) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str)
    .map(function(b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

function getRowColor(eventType) {
  if (eventType.indexOf("조사개시") !== -1) return "#fee2e2";
  if (eventType.indexOf("최종판정") !== -1) return "#dbeafe";
  if (eventType.indexOf("예비판정") !== -1) return "#fef3c7";
  if (eventType.indexOf("일몰재심") !== -1) return "#fde68a";
  if (eventType.indexOf("행정재심") !== -1) return "#f3e8ff";
  return "#ffffff";
}

function getExistingIds(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return {};
  var ids = {};
  sheet.getRange(2, 1, lastRow - 1, 1).getValues().forEach(function(r) { ids[String(r[0])] = true; });
  return ids;
}

function initSheet(sheet) {
  if (sheet.getLastRow() === 0) {
    var headers = ["ID","수집일시","국가","기관","이벤트유형","무역구제유형","제목","공고일","원문링크"];
    sheet.appendRow(headers);
    var hr = sheet.getRange(1, 1, 1, headers.length);
    hr.setBackground("#1e293b"); hr.setFontColor("white"); hr.setFontWeight("bold");
    sheet.setFrozenRows(1);
    [120,140,60,80,100,80,420,90,300].forEach(function(w,i) { sheet.setColumnWidth(i+1, w); });
  }
}

function classifyEventIN(text) {
  var t = text.toLowerCase();
  if (/initiation.*anti-dumping|anti-dumping.*initiation/.test(t)) return "조사개시";
  if (/initiation.*countervailing|countervailing.*initiation/.test(t)) return "조사개시";
  if (/initiation.*safeguard|safeguard.*initiation/.test(t)) return "조사개시";
  if (/preliminary finding|preliminary determination/.test(t)) return "예비판정";
  if (/final finding|final determination/.test(t)) return "최종판정";
  if (/sunset review|ssr/.test(t)) return "일몰재심";
  if (/mid.term review|mtr/.test(t)) return "행정재심";
  return "일반공고";
}

// =====================================================
// ── Gemini 호출 ───────────────────────────────────────
// =====================================================

function callGemini(payload) {
  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
    if (!apiKey) { Logger.log("❌ GEMINI_API_KEY 없음"); return null; }
    var options = {
      method: "POST", contentType: "application/json",
      headers: { "x-goog-api-key": apiKey },
      payload: JSON.stringify(Object.assign({}, payload, {
        generationConfig: Object.assign({ thinkingConfig: { thinkingBudget: 0 } }, payload.generationConfig || {})
      })), muteHttpExceptions: true
    };
    var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent";
    var res = UrlFetchApp.fetch(url, options);
    if (res.getResponseCode() === 503) {
      Utilities.sleep(3000);
      res = UrlFetchApp.fetch(url, options);
    }
    if (res.getResponseCode() === 429) {
      Logger.log("⚠️ Gemini 429 - 60초 대기 후 재시도");
      Utilities.sleep(60000);
      res = UrlFetchApp.fetch(url, options);
    }
    if (res.getResponseCode() !== 200) { Logger.log("❌ Gemini 오류: " + res.getResponseCode()); return null; }
    var cands = JSON.parse(res.getContentText()).candidates;
    if (cands && cands[0] && cands[0].content && cands[0].content.parts && cands[0].content.parts[0]) {
      var raw = cands[0].content.parts[0].text;
      return raw
        .replace(/\*\*/g, "").replace(/\*/g, "")
        .replace(/#{1,6}\s/g, "")
        .replace(/^-\s/gm, "").replace(/`/g, "")
        .trim();
    }
    return null;
  } catch(e) { Logger.log("❌ Gemini 오류: " + e.message); return null; }
}

function summarizeWithGemini(pdfUrl, prompt) {
  var pdfRes = UrlFetchApp.fetch(pdfUrl, { headers: { "User-Agent": "TradeRemedyMonitor/2.0" }, muteHttpExceptions: true });
  if (pdfRes.getResponseCode() !== 200) { Logger.log("❌ PDF 다운로드 실패"); return null; }
  var payload = {
    contents: [{ parts: [
      { inline_data: { mime_type: "application/pdf", data: Utilities.base64Encode(pdfRes.getContent()) } },
      { text: prompt || "이 PDF를 한국어로 3줄로 요약해줘." }
    ]}],
    generationConfig: { temperature: 0.1 },
    systemInstruction: { parts: [{ text: "마크다운 기호(**, *, #, _ 등)를 절대 사용하지 마세요. 순수 텍스트로만 답변하세요." }] }
  };
  var s = callGemini(payload);
  if (s) Logger.log("✅ PDF 요약 완료");
  return s;
}

function summarizeWithGeminiPdfViaScraper(pdfUrl, prompt) {
  try {
    var key = PropertiesService.getScriptProperties().getProperty("SCRAPER_API_KEY");
    if (!key) return null;
    var pdfRes = UrlFetchApp.fetch("https://api.scraperapi.com?api_key=" + key + "&url=" + encodeURIComponent(pdfUrl), { muteHttpExceptions: true });
    if (pdfRes.getResponseCode() !== 200) { Logger.log("❌ PDF 다운로드 실패"); return null; }
    var payload = {
      contents: [{ parts: [
        { inline_data: { mime_type: "application/pdf", data: Utilities.base64Encode(pdfRes.getContent()) } },
        { text: prompt || "이 PDF를 한국어로 3줄로 요약해줘." }
      ]}],
      generationConfig: { temperature: 0.1 },
      systemInstruction: { parts: [{ text: "마크다운 기호(**, *, #, _, - 등)를 절대 사용하지 마세요. 순수 텍스트로만 답변하세요." }] }
    };
    var s = callGemini(payload);
    if (s) Logger.log("✅ PDF 요약 완료 (ScraperAPI)");
    return s;
  } catch(e) { Logger.log("ScraperAPI PDF 오류: " + e.message); return null; }
}

function summarizeWithGeminiText(text, prompt) {
  var payload = {
    contents: [{ parts: [{ text: (prompt || "다음 내용을 한국어로 3줄로 요약해줘.\n\n") + text.substring(0, 10000) }] }],
    generationConfig: { temperature: 0.1 },
    systemInstruction: { parts: [{ text: "마크다운 기호(**, *, #, _, - 등)를 절대 사용하지 마세요. 순수 텍스트로만 답변하세요." }] }
  };
  var s = callGemini(payload);
  if (s) Logger.log("✅ 텍스트 요약 완료");
  return s;
}

// =====================================================
// ── 텍스트/PDF 추출 ───────────────────────────────────
// =====================================================

function extractTextFromTradeGov(url) {
  try {
    var res = UrlFetchApp.fetch(url, { headers: { "User-Agent": "TradeRemedyMonitor/2.0" }, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return null;
    return res.getContentText("UTF-8")
      .replace(/<script[\s\S]*?<\/script>/gi,"").replace(/<style[\s\S]*?<\/style>/gi,"")
      .replace(/<[^>]+>/g," ").replace(/&nbsp;/g," ").replace(/&amp;/g,"&")
      .replace(/&lt;/g,"<").replace(/&gt;/g,">").replace(/\s{2,}/g," ").trim();
  } catch(e) { return null; }
}

function extractTextFromCsms(url) {
  try {
    var res = UrlFetchApp.fetch(url, { headers: { "User-Agent": "TradeRemedyMonitor/2.0" }, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return null;
    return res.getContentText("UTF-8")
      .replace(/<script[\s\S]*?<\/script>/gi,"").replace(/<style[\s\S]*?<\/style>/gi,"")
      .replace(/<[^>]+>/g," ").replace(/&nbsp;/g," ").replace(/&amp;/g,"&")
      .replace(/&lt;/g,"<").replace(/&gt;/g,">").replace(/\s{2,}/g," ").trim();
  } catch(e) { return null; }
}

function extractPdfFromCbpEapa(caseUrl) {
  try {
    var key = PropertiesService.getScriptProperties().getProperty("SCRAPER_API_KEY");
    if (!key) return null;
    var res = UrlFetchApp.fetch("https://api.scraperapi.com?api_key=" + key + "&url=" + encodeURIComponent(caseUrl), { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return null;
    var m = res.getContentText("UTF-8").match(/href="(\/sites\/default\/files\/[^"]+\.pdf)"/i);
    return m ? "https://www.cbp.gov" + m[1] : null;
  } catch(e) { return null; }
}

function extractPdfFromKtcBoard(bbsId, html) {
  try {
    // 패턴 1: boardFileDownload.do?bbs_id=...&seq_no=...
    var re1 = /boardFileDownload\.do\?bbs_id=(\d+)&(?:amp;)?seq_no=(\d+)/g;
    var m, links = [];
    while ((m = re1.exec(html)) !== null) {
      links.push({ bbsId: m[1], seqNo: parseInt(m[2]) });
    }

    // 패턴 2: fileDownload('bbsId', 'seqNo')
    if (!links.length) {
      var re2 = /fileDownload\('(\d+)',\s*'(\d+)'\)/g;
      while ((m = re2.exec(html)) !== null) {
        links.push({ bbsId: m[1], seqNo: parseInt(m[2]) });
      }
    }

    if (!links.length) return null;

    // seqNo가 가장 큰 것이 최신
    links.sort(function(a, b) { return b.seqNo - a.seqNo; });
    var latest = links[0];
    return "https://www.ktc.go.kr/boardFileDownload.do?bbs_id=" + latest.bbsId + "&seq_no=" + latest.seqNo;
  } catch(e) { return null; }
}

function extractTextFromKtcBoard(html) {
  try {
    var m = /class="detail-cont"[\s\S]{0,50}?>([\s\S]+?)<\/div>\s*<!--\s*\/\/게시판/.exec(html);
    return (m ? m[1] : html)
      .replace(/<script[\s\S]*?<\/script>/gi,"").replace(/<style[\s\S]*?<\/style>/gi,"")
      .replace(/<br\s*\/?>/gi,"\n").replace(/<[^>]+]/g," ")
      .replace(/&nbsp;/g," ").replace(/&amp;/g,"&").replace(/&lt;/g,"<").replace(/&gt;/g,">")
      .replace(/\s{2,}/g," ").trim();
  } catch(e) { return null; }
}

function fetchKtcBoardDetail(bbsId) {
  try {
    var key = PropertiesService.getScriptProperties().getProperty("SCRAPER_API_KEY");
    var res = UrlFetchApp.fetch(
      "https://api.scraperapi.com?api_key=" + key +
      "&url=" + encodeURIComponent(
        "https://www.ktc.go.kr/boardView.do?bbs_id=" + bbsId + "&menuId=46&pageIndex=1"
      ),
      { muteHttpExceptions: true }
    );
    return res.getResponseCode() === 200 ? res.getContentText("UTF-8") : null;
  } catch(e) { return null; }
}

function extractPdfFromDGTR(caseUrl) {
  try {
    var res = UrlFetchApp.fetch(caseUrl, {
      headers: { "User-Agent": "TradeRemedyMonitor/2.0" },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) return null;
    var html = res.getContentText("UTF-8");

    // S.No가 가장 큰 행의 PDF 추출
    var rowRe = /<tr>[\s\S]*?<td>(\d+)<\/td>[\s\S]*?href="(https?:\/\/dgtr\.gov\.in[^"]+\.pdf)"[\s\S]*?<\/tr>/g;
    var m, latestPdf = null, latestSno = -1;
    while ((m = rowRe.exec(html)) !== null) {
      var sno = parseInt(m[1]);
      if (sno > latestSno) {
        latestSno = sno;
        latestPdf = m[2];
      }
    }
    return latestPdf;
  } catch(e) { return null; }
}

// =====================================================
// ── 요약 실행 함수들 ──────────────────────────────────
// =====================================================

function summarizeNewIndiaCases(arr) {
  var targets = arr.slice(0, 3);
  if (arr.length > 3) { Logger.log("⚠️ 인도 " + arr.length + "건 중 3건만 요약"); }
  targets.forEach(function(c) {
    Utilities.sleep(2000);
    var pdfUrl = extractPdfFromDGTR(c.url);
    if (!pdfUrl) return;
    var s = summarizeWithGemini(pdfUrl, "이 PDF는 인도 무역구제 공문이야. 조사 대상 품목, 피조사국, 핵심 내용을 한국어로 3줄로 요약해줘.");
    if (s) c.summary = s;
    Utilities.sleep(2000);
  });
}

function summarizeNewUsCases(arr) {
  if (arr.length > 10) { Logger.log("⚠️ 미국 " + arr.length + "건 초과 생략"); return; }
  arr.forEach(function(c) {
    Utilities.sleep(2000);
    var text = extractTextFromTradeGov(c.url);
    if (!text || text.length < 100) return;
    var s = summarizeWithGeminiText(text, "이 내용은 미국 무역구제 케이스야. 조사 대상 품목, 피조사국, 핵심 내용을 한국어로 3줄로 요약해줘.\n\n");
    if (s) c.summary = s;
    Utilities.sleep(2000);
  });
}

function summarizeNewKrNoticeCases(arr) {
  if (arr.length > 10) { Logger.log("⚠️ 한국 " + arr.length + "건 초과 생략"); return; }
  arr.forEach(function(c) {
    Utilities.sleep(2000);
    var html = fetchKtcBoardDetail(c.bbsId);
    if (!html) return;
    var pdfUrl = extractPdfFromKtcBoard(c.bbsId, html);
    if (pdfUrl) {
      var s = summarizeWithGeminiPdfViaScraper(pdfUrl, "이 PDF는 한국 무역위원회 공고야. 조사 대상 품목, 피조사국, 핵심 내용을 한국어로 3줄로 요약해줘.");
      if (s) { c.summary = s; }
      else {
        // PDF 다운로드 실패 시 텍스트로 fallback
        var text = extractTextFromKtcBoard(html);
        if (text && text.length >= 50) {
          var sf = summarizeWithGeminiText(text, "이 내용은 한국 무역위원회 공고야. 조사 대상 품목, 피조사국, 핵심 내용을 한국어로 3줄로 요약해줘.\n\n");
          if (sf) c.summary = sf;
        }
      }
    } else {
      var text = extractTextFromKtcBoard(html);
      if (!text || text.length < 50) return;
      var s2 = summarizeWithGeminiText(text, "이 내용은 한국 무역위원회 공고야. 조사 대상 품목, 피조사국, 핵심 내용을 한국어로 3줄로 요약해줘.\n\n");
      if (s2) c.summary = s2;
    }
    Utilities.sleep(2000);
  });
}

function summarizeNewEapaCases(arr) {
  if (arr.length > 10) { Logger.log("⚠️ EAPA " + arr.length + "건 초과 생략"); return; }
  arr.forEach(function(c) {
    Utilities.sleep(2000);
    var pdfUrl = extractPdfFromCbpEapa(c.url);
    if (!pdfUrl) return;
    var s = summarizeWithGeminiPdfViaScraper(pdfUrl, "이 PDF는 미국 CBP EAPA 조사 공문이야. 조사 대상 기업, 품목, 혐의 내용을 한국어로 3줄로 요약해줘.");
    if (s) c.summary = s;
    Utilities.sleep(2000);
  });
}

function summarizeNewCsmsCases(arr) {
  if (arr.length > 10) { Logger.log("⚠️ CSMS " + arr.length + "건 초과 생략"); return; }
  arr.forEach(function(c) {
    Utilities.sleep(2000);
    var text = extractTextFromCsms(c.url);
    if (!text || text.length < 100) return;
    var s = summarizeWithGeminiText(text, "이 내용은 미국 CBP CSMS 메시지야. 무역 관련 핵심 내용을 한국어로 3줄로 요약해줘.\n\n");
    if (s) c.summary = s;
    Utilities.sleep(2000);
  });
}

// =====================================================
// ── CBP EAPA / CSMS 수집 ──────────────────────────────
// =====================================================

function fetchCbpEapaCases() {
  var key = PropertiesService.getScriptProperties().getProperty("SCRAPER_API_KEY");
  if (!key) { Logger.log("❌ SCRAPER_API_KEY 없음"); return []; }
  var resp;
  try {
    resp = UrlFetchApp.fetch("https://api.scraperapi.com?api_key=" + key + "&url=" + encodeURIComponent(CBP_EAPA_URL), { muteHttpExceptions: true });
  } catch(e) { Logger.log("❌ CBP EAPA 오류: " + e.message); return []; }
  if (resp.getResponseCode() !== 200) return [];
  var html = resp.getContentText();
  var cases = [];
  var s1 = html.indexOf('id="acc33421"'), s2 = html.indexOf('id="acc33423"');
  if (s1 === -1) return [];
  var sec = html.substring(s1, s2 !== -1 ? s2 : s1 + 50000);
  var re = /<a\s+href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/g, m;
  while ((m = re.exec(sec)) !== null) {
    var href = m[1], text = m[2].replace(/<[^>]+>/g,"").trim();
    if (href.indexOf("/document/") === -1) continue;
    var cm = text.match(/^(EAPA\s+(?:Cons\.?\s+)?(?:Case|Investigation|Consolidated Case)[^:]*?\d{4,5}[^:]*?):\s*(.*?)\s*\(([^)]+)\)\s*$/);
    if (!cm) continue;
    var dm = cm[3].match(/([A-Za-z]+ \d+,\s*\d{4})$/);
    var dateStr = dm ? dm[1] : "";
    if (dateStr && parseInt(dateStr.match(/\d{4}$/)[0]) < 2025) break;
    var caseId = cm[1].trim(), nm = caseId.match(/(\d{4,5})/);
    cases.push({ caseNum: nm ? nm[1] : caseId, caseId: caseId, company: cm[2].trim(), date: dateStr, url: "https://www.cbp.gov" + href });
  }
  Logger.log("   CBP EAPA: " + cases.length + "건");
  return cases;
}

function fetchCsmsMessages() {
  try {
    var resp = UrlFetchApp.fetch(CSMS_JSON_URL, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return [];
    return JSON.parse(resp.getContentText());
  } catch(e) { Logger.log("❌ CSMS 오류: " + e.message); return []; }
}

// =====================================================
// ── EU TRON ───────────────────────────────────────────
// =====================================================

function initEuTrackSheet(sheet) {
  if (sheet.getLastRow() === 0) {
    var headers = [
      "케이스ID","케이스번호","품목","피조사국","조사유형",
      "무역구제유형","조사개시일","마지막확인일시","Publication수","마지막Publication유형","상세URL"
    ];
    sheet.appendRow(headers);
    var hr = sheet.getRange(1, 1, 1, headers.length);
    hr.setBackground("#1e293b"); hr.setFontColor("white"); hr.setFontWeight("bold");
    sheet.setFrozenRows(1);
    [80,90,200,200,150,80,100,140,80,200,300].forEach(function(w,i) { sheet.setColumnWidth(i+1, w); });
  }
}

function getEuTrackedCases(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return {};
  var tracked = {};
  sheet.getRange(2, 1, lastRow - 1, 11).getValues().forEach(function(r, i) {
    if (!r[0]) return;
    tracked[String(r[0])] = { rowIndex: i + 2, pubCount: Number(r[8]) || 0 };
  });
  return tracked;
}

function classifyEuTradeType(caseNumber) {
  if (/^AS/.test(caseNumber)) return "CVD";
  if (/^SG/.test(caseNumber)) return "SG";
  return "AD";
}

function classifyEuEventType(caseType) {
  if (!caseType) return "일반공고";
  var t = caseType.toLowerCase();
  if (/initial investigation/.test(t))  return "조사개시";
  if (/expiry review/.test(t))          return "일몰재심";
  if (/interim review/.test(t))         return "행정재심";
  if (/new exporting producer/.test(t)) return "행정재심";
  if (/anti-circumvention/.test(t))     return "조사개시";
  return "일반공고";
}

function classifyEuPubEventType(pubType) {
  if (!pubType) return "일반공고";
  var t = pubType.toLowerCase();
  if (/initiation/.test(t))                                                       return "조사개시";
  if (/provisional measures|provisional anti-dumping|provisional countervailing/.test(t)) return "예비판정";
  if (/definitive measures|definitive anti-dumping|definitive countervailing/.test(t))    return "최종판정";
  if (/expiry review|sunset review/.test(t))                                      return "일몰재심";
  if (/interim review/.test(t))                                                   return "행정재심";
  if (/termination|withdrawal/.test(t))                                           return "조사종료";
  return "일반공고";
}

function fetchEuCaseDetails(caseId) {
  try {
    var res = UrlFetchApp.fetch(EU_TRON_DETAIL_URL + caseId, {
      headers: { "User-Agent": "TradeRemedyMonitor/2.0", "Accept": "application/json" },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) return null;
    return JSON.parse(res.getContentText("UTF-8"));
  } catch(e) { Logger.log("❌ EU 상세 조회 오류 (id=" + caseId + "): " + e.message); return null; }
}

function fetchEuOngoingList() {
  try {
    var res = UrlFetchApp.fetch(EU_TRON_LIST_URL, {
      headers: { "User-Agent": "TradeRemedyMonitor/2.0", "Accept": "application/json" },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) { Logger.log("❌ EU TRON 목록 오류: " + res.getResponseCode()); return []; }
    var data = JSON.parse(res.getContentText("UTF-8"));
    return Array.isArray(data) ? data : [];
  } catch(e) { Logger.log("❌ EU TRON 목록 오류: " + e.message); return []; }
}

function getEuTrackRowIndex(sheet, caseId) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return -1;
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(caseId)) return i + 2;
  }
  return -1;
}

// 반환값: { newCases: [], koreanUpdates: [] }
function runEuCaseTracker(ss, nowStr, newCases) {
  var trackSheet = ss.getSheetByName(EU_TRACK_SHEET_NAME) || ss.insertSheet(EU_TRACK_SHEET_NAME);
  initEuTrackSheet(trackSheet);

  var tracked       = getEuTrackedCases(trackSheet);
  var koreanUpdates = [];
  var euNewCases    = [];

  Logger.log("📡 EU-TRON");
  var ongoingList = fetchEuOngoingList();
  Logger.log("   → ongoing 케이스: " + ongoingList.length + "건");

  // 20건씩 청크로 나눠 병렬 수집
  var chunkSize = 20;
  for (var ci = 0; ci < ongoingList.length; ci += chunkSize) {
    var chunk = ongoingList.slice(ci, ci + chunkSize);

    var detailRequests = chunk.map(function(item) {
      return {
        url: EU_TRON_DETAIL_URL + String(item.id),
        headers: { "User-Agent": "TradeRemedyMonitor/2.0", "Accept": "application/json" },
        muteHttpExceptions: true
      };
    });
    var detailResponses = UrlFetchApp.fetchAll(detailRequests);

    detailResponses.forEach(function(res, j) {
      if (res.getResponseCode() !== 200) return;
      var details;
      try { details = JSON.parse(res.getContentText("UTF-8")); } catch(e) { return; }

      var caseId    = String(chunk[j].id);
      var pubs      = details.publications || [];
      var pubCount  = pubs.length;
      var countries = (details.caseCountries || []).map(function(cc) { return cc.countryName; }).join(", ");
      var caseNum   = details.caseNumber || "";
      var shortName = (details.shortName  || "").trim();
      var caseType  = details.caseType   || "";
      var tradeType = classifyEuTradeType(caseNum);
      var eventType = classifyEuEventType(caseType);
      var detailUrl = "https://tron.trade.ec.europa.eu/investigations/case-view?caseId=" + caseId;

      var initDate = "";
      if (details.initDate) {
        try { initDate = Utilities.formatDate(new Date(details.initDate), "Asia/Seoul", "yyyy-MM-dd"); } catch(e) {}
      }
      var lastPubType = pubs.length > 0 ? (pubs[pubs.length - 1].typeOfPublication || "") : "";

      if (tracked[caseId]) {
        var prevCount = tracked[caseId].pubCount;
        if (pubCount > prevCount) {
          var newPubs = pubs.slice(prevCount);
          Logger.log("🔔 EU 변화: [" + caseNum + "] " + shortName + " +" + (pubCount - prevCount) + "건");

          newPubs.forEach(function(pub) {
            var pubEventType = classifyEuPubEventType(pub.typeOfPublication);
            var pubDate = "";
            if (pub.datePublication) {
              try { pubDate = Utilities.formatDate(new Date(pub.datePublication), "Asia/Seoul", "yyyy-MM-dd"); } catch(e) {}
            }
            var pubUrl = pub.urlTransformed || pub.url || detailUrl;
            if (!pubUrl || pubUrl === "N/A") pubUrl = detailUrl;

            var pubSummary = "";
            if (pub.contents && pub.contents.trim().length > 20) {
              Utilities.sleep(3000);
              pubSummary = summarizeWithGeminiText(
                pub.contents.trim(),
                "이 내용은 EU 무역구제 공고야. 조사 대상 품목, 피조사국, 핵심 내용을 한국어로 3줄로 요약해줘.\n\n"
              ) || "";
            }

            if (/korea/i.test(countries)) {
              koreanUpdates.push({
                caseNum:   caseNum,
                shortName: shortName,
                countries: countries,
                tradeType: tradeType,
                eventType: pubEventType,
                pubType:   pub.typeOfPublication || "",
                pubDate:   pubDate,
                summary:   pubSummary,
                url:       pubUrl,
                detailUrl: detailUrl
              });
            }
          });

          var rowIdx = getEuTrackRowIndex(trackSheet, caseId);
          if (rowIdx > 0) {
            trackSheet.getRange(rowIdx, 8).setValue(nowStr);
            trackSheet.getRange(rowIdx, 9).setValue(pubCount);
            trackSheet.getRange(rowIdx, 10).setValue(lastPubType);
            trackSheet.getRange(rowIdx, 1, 1, 11).setBackground("#FEF9C3");
          }
        } else {
          var rowIdx2 = getEuTrackRowIndex(trackSheet, caseId);
          if (rowIdx2 > 0) {
            trackSheet.getRange(rowIdx2, 8).setValue(nowStr);
            trackSheet.getRange(rowIdx2, 1, 1, 11).setBackground("#FFFFFF");
          }
        }
      } else {
        var newRow = [caseId, caseNum, shortName, countries, caseType, tradeType, initDate, nowStr, pubCount, lastPubType, detailUrl];
        trackSheet.appendRow(newRow);
        var lastRow = trackSheet.getLastRow();
        trackSheet.getRange(lastRow, 1, 1, newRow.length).setBackground("#dcfce7");
        trackSheet.getRange(lastRow, 11).setFormula('=HYPERLINK("' + detailUrl + '","상세보기")');
        tracked[caseId] = { pubCount: pubCount };
        Logger.log("   → 신규 등록: [" + caseNum + "] " + shortName);

        var euSummary = "";
        if (details.initPublication && details.initPublication.contents) {
          Utilities.sleep(3000);
          euSummary = summarizeWithGeminiText(
            details.initPublication.contents.trim(),
            "이 내용은 EU 무역구제 조사 개시 공고야. 조사 대상 품목, 피조사국, 핵심 내용을 한국어로 3줄로 요약해줘.\n\n"
          ) || "";
        }
        euNewCases.push({
          url: detailUrl, title: "[" + caseNum + "] " + shortName + " &larr; " + countries,
          eventType: eventType, tradeType: tradeType, country: "EU", agency: "EC",
          published: initDate, summary: euSummary
        });
      }
    });

    // 청크 사이 간격 (서버 부하 방지)
    if (ci + chunkSize < ongoingList.length) Utilities.sleep(500);
  }

  Logger.log("   → EU 신규: " + euNewCases.length + "건 / 한국 업데이트: " + koreanUpdates.length + "건");
  euNewCases.forEach(function(c) { newCases.push(c); });
  return koreanUpdates;
}

// =====================================================
// ── 브라질 DECOM ──────────────────────────────────────
// =====================================================

function initBrazilMeasuresSheet(sheet) {
  if (sheet.getLastRow() === 0) {
    var headers = ["상세URL","품목","조치유형","원산지","유효기간","마지막업데이트일(사이트)","마지막확인일시"];
    sheet.appendRow(headers);
    var hr = sheet.getRange(1, 1, 1, headers.length);
    hr.setBackground("#1e293b"); hr.setFontColor("white"); hr.setFontWeight("bold");
    sheet.setFrozenRows(1);
    [300,200,150,250,100,140,140].forEach(function(w,i) { sheet.setColumnWidth(i+1, w); });
  }
}

function initBrazilInvestSheet(sheet) {
  if (sheet.getLastRow() === 0) {
    var headers = ["상세URL","품목","조사유형","원산지","담당부서","마지막업데이트일(사이트)","현재상황(SituacaoAtual)","마지막확인일시"];
    sheet.appendRow(headers);
    var hr = sheet.getRange(1, 1, 1, headers.length);
    hr.setBackground("#1e293b"); hr.setFontColor("white"); hr.setFontWeight("bold");
    sheet.setFrozenRows(1);
    [300,200,120,250,80,140,250,140].forEach(function(w,i) { sheet.setColumnWidth(i+1, w); });
  }
}

function fetchBrazilHtml(url) {
  try {
    var key = PropertiesService.getScriptProperties().getProperty("SCRAPER_API_KEY");
    if (!key) { Logger.log("❌ SCRAPER_API_KEY 없음"); return null; }
    var res = UrlFetchApp.fetch("https://api.scraperapi.com?api_key=" + key + "&url=" + encodeURIComponent(url), { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) { Logger.log("❌ fetchBrazilHtml 실패: " + res.getResponseCode()); return null; }
    return res.getContentText("UTF-8");
  } catch(e) { Logger.log("❌ fetchBrazilHtml 오류: " + e.message); return null; }
}

function parseBrazilTable(html, tableIndex) {
  if (!html) return [];
  tableIndex = tableIndex || 0;
  var tableRe = /<table class="plain">([\s\S]*?)<\/table>/g;
  var tables = [], tm;
  while ((tm = tableRe.exec(html)) !== null) { tables.push(tm[1]); }
  if (!tables[tableIndex]) return [];

  var rows = [], rowRe = /<tr[\s\S]*?<\/tr>/g, rm;
  while ((rm = rowRe.exec(tables[tableIndex])) !== null) { rows.push(rm[0]); }

  var result = [];
  rows.forEach(function(row, idx) {
    if (idx === 0) return;
    var tdRe = /<td[^>]*>([\s\S]*?)<\/td>/g, tds = [], tdm;
    while ((tdm = tdRe.exec(row)) !== null) { tds.push(tdm[1]); }
    if (!tds.length) return;
    var hrefMatch = row.match(/href="(https?:\/\/[^"]+)"/);
    var cols = tds.map(function(td) {
      return td.replace(/<[^>]+>/g," ").replace(/&nbsp;/g," ").replace(/&amp;/g,"&")
               .replace(/&lt;/g,"<").replace(/&gt;/g,">").replace(/\s{2,}/g," ").trim();
    });
    if (cols[0]) result.push({ href: hrefMatch ? hrefMatch[1] : "", cols: cols });
  });
  return result;
}

function cleanBrazilText(html) {
  return (html || "").replace(/<[^>]+>/g," ").replace(/&nbsp;/g," ").replace(/&amp;/g,"&")
    .replace(/&lt;/g,"<").replace(/&gt;/g,">").replace(/\s{2,}/g," ").trim();
}

function extractAtualizado(html) {
  if (!html) return "";
  var m = html.match(/Atualizado em[\s\S]{0,200}?<span class="value">([^<]+)<\/span>/);
  return m ? m[1].trim().substring(0, 10) : "";
}

function extractSituacaoAtual(html) {
  if (!html) return "";
  var m = html.match(/Situa(?:ção|&ccedil;&atilde;o|&#231;&#227;o|c[a-z]o) Atual[\s\S]{0,300}?<\/[a-z]+>\s*([\s\S]{0,300}?)<\/[a-z]+>/i);
  if (m) return cleanBrazilText(m[1]).substring(0, 200);
  var m2 = html.match(/Situa[^\n<]{0,30}Atual[^\n<]{0,50}\n?([\s\S]{0,300}?)(?:<\/tr>|<\/div>|<h[0-9])/i);
  return m2 ? cleanBrazilText(m2[1]).substring(0, 200) : "";
}

function isKoreanRelated(origin) {
  return /coreia|cor\u00e9ia/i.test(origin || "");
}

function formatBrazilDate(val) {
  if (!val) return "";
  if (val instanceof Date) return Utilities.formatDate(val, "Asia/Seoul", "dd/MM/yyyy");
  return String(val);
}

function getBrazilSheetData(sheet, hasSituacao) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return {};
  var data = sheet.getRange(2, 1, lastRow - 1, hasSituacao ? 8 : 7).getValues();
  var result = {};
  data.forEach(function(r, i) {
    var url = String(r[0]);
    if (!url) return;
    result[url] = { rowIndex: i + 2, lastUpdated: formatBrazilDate(r[5]), situacao: hasSituacao ? String(r[6] || "") : "" };
  });
  return result;
}

function fetchBrazilMedidas(ss, nowStr, newCases) {
  Logger.log("📡 BR-DECOM 조치 현황");
  var sheet = ss.getSheetByName(CONFIG_V2.BRAZIL_MEASURES_SHEET) || ss.insertSheet(CONFIG_V2.BRAZIL_MEASURES_SHEET);
  initBrazilMeasuresSheet(sheet);
  var existing = getBrazilSheetData(sheet, false);
  var html = fetchBrazilHtml(BRAZIL_MEASURES_URL);
  if (!html) { Logger.log("   → HTML 수집 실패"); return []; }

  var rows = parseBrazilTable(html, 0);
  Logger.log("   → 파싱된 행: " + rows.length + "건");
  var newCount = 0, detailUpdates = [];

  rows.forEach(function(row) {
    var href = row.href, produto = row.cols[0] || "", medida = row.cols[1] || "",
        origem = row.cols[2] || "", prazo = row.cols[3] || "";
    if (!href || !produto) return;

    if (!existing[href]) {
      var newRow = [href, produto, medida, origem, prazo, "", nowStr];
      sheet.appendRow(newRow);
      sheet.getRange(sheet.getLastRow(), 1, 1, newRow.length).setBackground("#dcfce7");
      existing[href] = { rowIndex: sheet.getLastRow(), lastUpdated: "", situacao: "" };
      newCount++;
      var brMedidaSummary = "";
      if (newCount <= 10) {
        Utilities.sleep(3000);
        var brMedidaHtml = fetchBrazilHtml(href);
        if (brMedidaHtml) {
          var brMedidaText = cleanBrazilText(brMedidaHtml).substring(0, 10000);
          if (brMedidaText.length > 100) {
            brMedidaSummary = summarizeWithGeminiText(
              brMedidaText,
              "이 내용은 브라질 무역구제 조치야. 대상 품목, 원산지, 조치 내용을 한국어로 3줄로 요약해줘.\n\n"
            ) || "";
          }
        }
      }
      newCases.push({
        url: href, title: "[신규 조치] " + produto + " &larr; " + origem,
        eventType: "조치발동", tradeType: medida.indexOf("Compensat") !== -1 ? "CVD" : "AD",
        country: "BR", agency: "DECOM", published: prazo, summary: brMedidaSummary
      });
      Logger.log("   → 신규 조치: " + produto);
    }

    if (isKoreanRelated(origem)) {
      var detailHtml = fetchBrazilHtml(href);
      if (!detailHtml) return;
      var atualizado = extractAtualizado(detailHtml);
      var prevUpdated = existing[href] ? existing[href].lastUpdated : "";
      if (atualizado && atualizado !== prevUpdated) {
        var rowIdx = existing[href] ? existing[href].rowIndex : -1;
        if (rowIdx > 1) {
          sheet.getRange(rowIdx, 6).setValue(atualizado);
          sheet.getRange(rowIdx, 7).setValue(nowStr);
          sheet.getRange(rowIdx, 1, 1, 7).setBackground("#FEF9C3");
        }
        if (prevUpdated) {
          detailUpdates.push({ produto: produto, medida: medida, origem: origem,
            prevUpdated: prevUpdated, newUpdated: atualizado, situacao: "", prevSituacao: "",
            url: href, type: "조치" });
          Logger.log("   🔔 조치 업데이트: " + produto + " (" + prevUpdated + " -> " + atualizado + ")");
        } else {
          if (existing[href]) existing[href].lastUpdated = atualizado;
        }
      } else {
        var rowIdxReset = existing[href] ? existing[href].rowIndex : -1;
        if (rowIdxReset > 1) {
          sheet.getRange(rowIdxReset, 7).setValue(nowStr);
          sheet.getRange(rowIdxReset, 1, 1, 7).setBackground("#FFFFFF");
        }
      }
    }
  });

  var currentUrls = {};
  rows.forEach(function(r) { if (r.href) currentUrls[r.href] = true; });
  Object.keys(existing).forEach(function(url) {
    if (!currentUrls[url]) {
      Logger.log("   ⚠️ 브라질 조치 미감지 (종료 처리 생략): " + url);
      // 파싱 실패로 인한 오탐 방지 - 종료 처리 안 함
    }
  });

  Logger.log("   → 조치: 신규 " + newCount + "건 / 한국 상세 업데이트 " + detailUpdates.length + "건");
  return detailUpdates;
}

function fetchBrazilInvestigacoes(ss, nowStr, newCases) {
  Logger.log("📡 BR-DECOM 진행 중 조사");
  var sheet = ss.getSheetByName(CONFIG_V2.BRAZIL_INVEST_SHEET) || ss.insertSheet(CONFIG_V2.BRAZIL_INVEST_SHEET);
  initBrazilInvestSheet(sheet);
  var existing = getBrazilSheetData(sheet, true);
  var html = fetchBrazilHtml(BRAZIL_INVEST_URL);
  if (!html) { Logger.log("   → HTML 수집 실패"); return []; }

  var allRows = [];
  parseBrazilTable(html, 0).forEach(function(r) {
    allRows.push({ href: r.href, produto: r.cols[0]||"", investig: r.cols[1]||"", origem: r.cols[2]||"", coord: r.cols[3]||"" });
  });

  Logger.log("   → 파싱된 행: " + allRows.length + "건");

  var newCount = 0, detailUpdates = [];

  allRows.forEach(function(row) {
    var href = row.href, produto = row.produto, investig = row.investig, origem = row.origem, coord = row.coord;
    if (!href || !produto) return;

    if (!existing[href]) {
      var tradeType = /subsídio|cvd/i.test(investig) ? "CVD" : /salvaguarda|safeguard/i.test(investig) ? "SG" : "AD";
      var newRow = [href, produto, investig, origem, coord, "", "", nowStr];
      sheet.appendRow(newRow);
      sheet.getRange(sheet.getLastRow(), 1, 1, newRow.length).setBackground("#dcfce7");
      existing[href] = { rowIndex: sheet.getLastRow(), lastUpdated: "", situacao: "" };
      newCount++;
      var brInvestSummary = "";
      if (newCount <= 10) {
        Utilities.sleep(3000);
        var brInvestHtml = fetchBrazilHtml(href);
        if (brInvestHtml) {
          var brInvestText = cleanBrazilText(brInvestHtml).substring(0, 10000);
          if (brInvestText.length > 100) {
            brInvestSummary = summarizeWithGeminiText(
              brInvestText,
              "이 내용은 브라질 무역구제 조사야. 대상 품목, 원산지, 조사 내용을 한국어로 3줄로 요약해줘.\n\n"
            ) || "";
          }
        }
      }
      newCases.push({ url: href, title: "[신규 조사] " + produto + " &larr; " + (origem||""),
        eventType: "조사개시", tradeType: tradeType, country: "BR", agency: "DECOM",
        published: "", summary: brInvestSummary });
      Logger.log("   → 신규 조사: " + produto);
    }

    if (isKoreanRelated(origem)) {
      var detailHtml = fetchBrazilHtml(href);
      if (!detailHtml) return;
      var atualizado = extractAtualizado(detailHtml), situacao = extractSituacaoAtual(detailHtml);
      var prevUpdated = existing[href] ? existing[href].lastUpdated : "";
      var prevSituacao = existing[href] ? existing[href].situacao : "";

      if (atualizado && atualizado !== prevUpdated) {
        var rowIdx = existing[href] ? existing[href].rowIndex : -1;
        if (rowIdx > 1) {
          sheet.getRange(rowIdx, 6).setValue(atualizado);
          sheet.getRange(rowIdx, 7).setValue(situacao);
          sheet.getRange(rowIdx, 8).setValue(nowStr);
          sheet.getRange(rowIdx, 1, 1, 8).setBackground("#FEF9C3");
        }
        if (prevUpdated) {
          detailUpdates.push({ produto: produto, investig: investig, origem: origem,
            prevUpdated: prevUpdated, newUpdated: atualizado, situacao: situacao,
            prevSituacao: prevSituacao, url: href, type: "조사" });
          Logger.log("   🔔 조사 업데이트: " + produto + " (" + prevUpdated + " -> " + atualizado + ")");
        } else {
          if (existing[href]) { existing[href].lastUpdated = atualizado; existing[href].situacao = situacao; }
        }
      } else if (situacao && situacao !== prevSituacao && prevSituacao) {
        var rowIdx2 = existing[href] ? existing[href].rowIndex : -1;
        if (rowIdx2 > 1) {
          sheet.getRange(rowIdx2, 7).setValue(situacao);
          sheet.getRange(rowIdx2, 8).setValue(nowStr);
          sheet.getRange(rowIdx2, 1, 1, 8).setBackground("#FEF9C3");
        }
        detailUpdates.push({ produto: produto, investig: investig, origem: origem,
          prevUpdated: prevUpdated, newUpdated: atualizado, situacao: situacao,
          prevSituacao: prevSituacao, url: href, type: "조사(상황변경)" });
      } else {
              var rowIdx3 = existing[href] ? existing[href].rowIndex : -1;
              if (rowIdx3 > 1) {
                sheet.getRange(rowIdx3, 8).setValue(nowStr);
                sheet.getRange(rowIdx3, 1, 1, 8).setBackground("#FFFFFF");
              }
            }
    }
  });

  var currentUrls = {};
  allRows.forEach(function(r) { if (r.href) currentUrls[r.href] = true; });
  Object.keys(existing).forEach(function(url) {
    if (!currentUrls[url]) {
      Logger.log("   ⚠️ 브라질 조사 미감지 (종료 처리 생략): " + url);
      // 파싱 실패로 인한 오탐 방지 - 종료 처리 안 함
    }
  });

  Logger.log("   → 조사: 신규 " + newCount + "건 / 한국 상세 업데이트 " + detailUpdates.length + "건");
  return detailUpdates;
}

// 반환값: koreanBrUpdates 배열
function fetchBrazilCases(ss, nowStr, newCases) {
  Logger.log("========================================");
  Logger.log("📡 BR-DECOM 수집 시작");
  var medidasUpdates  = fetchBrazilMedidas(ss, nowStr, newCases)      || [];
  var investigUpdates = fetchBrazilInvestigacoes(ss, nowStr, newCases) || [];
  var allDetailUpdates = medidasUpdates.concat(investigUpdates);
  Logger.log("📡 BR-DECOM 수집 완료 (한국 상세 업데이트: " + allDetailUpdates.length + "건)");
  Logger.log("========================================");
  return allDetailUpdates;
}

// =====================================================
// ── 호주 ADC ──────────────────────────────────────────
// =====================================================

function fetchAustraliaCases(ss, nowStr, newCases) {
  Logger.log("📡 AU-ADC");
  var key = PropertiesService.getScriptProperties().getProperty("SCRAPER_API_KEY");
  if (!key) { Logger.log("❌ SCRAPER_API_KEY 없음"); return []; }

  // 추적 시트 초기화
  var trackSheet = ss.getSheetByName("호주 ADC 추적") || ss.insertSheet("호주 ADC 추적");
  if (trackSheet.getLastRow() === 0) {
    var headers = ["케이스URL","케이스번호","품목","국가","조사유형","Next Milestone","Last Updated","Status","마지막확인일시"];
    trackSheet.appendRow(headers);
    var hr = trackSheet.getRange(1, 1, 1, headers.length);
    hr.setBackground("#1e293b"); hr.setFontColor("white"); hr.setFontWeight("bold");
    trackSheet.setFrozenRows(1);
    [300,80,250,120,150,150,100,80,140].forEach(function(w,i) { trackSheet.setColumnWidth(i+1, w); });
  }

  // 기존 추적 데이터 로드
  var tracked = {};
  if (trackSheet.getLastRow() > 1) {
    trackSheet.getRange(2, 1, trackSheet.getLastRow()-1, 9).getValues()
      .forEach(function(r, i) {
        if (!r[0]) return;
        var lastUpdated = r[6] instanceof Date
          ? Utilities.formatDate(r[6], "Asia/Seoul", "dd/MM/yyyy")
          : String(r[6] || "");
        tracked[String(r[0])] = {
          rowIndex:    i + 2,
          lastUpdated: lastUpdated,
          status:      String(r[7] || "")
        };
      });
  }
  Logger.log("   기존 등록: " + Object.keys(tracked).length + "건");

  function classifyAuEventType(caseType) {
    var t = (caseType || "").toLowerCase();
    if (/investigation/.test(t))      return "조사개시";
    if (/revocation/.test(t))         return "일몰재심";
    if (/review/.test(t))             return "행정재심";
    if (/anti-circumvention/.test(t)) return "조사개시";
    if (/continuation/.test(t))       return "행정재심";
    if (/exemption/.test(t))          return "행정재심";
    return "일반공고";
  }

  // 사이트 수집
  var res = UrlFetchApp.fetch(
    "https://api.scraperapi.com?api_key=" + key + "&url=" + encodeURIComponent(AU_ADC_URL),
    { muteHttpExceptions: true }
  );
  if (res.getResponseCode() !== 200) {
    Logger.log("   ❌ 수집 실패: " + res.getResponseCode());
    PropertiesService.getScriptProperties().setProperty("MONITOR_AU_FAILED", "Y");
    return [];
  }
  var html = res.getContentText("UTF-8");

  var tbodyMatch = /<tbody>([\s\S]*?)<\/tbody>/.exec(html);
  if (!tbodyMatch) { Logger.log("   ❌ tbody 없음"); return []; }

  var rowRe = /<tr>([\s\S]*?)<\/tr>/g;
  var m, cases = [];
  while ((m = rowRe.exec(tbodyMatch[1])) !== null) {
    var row = m[1];
    var tds = [];
    var tdRe = /<td[^>]*>([\s\S]*?)<\/td>/g, tdm;
    while ((tdm = tdRe.exec(row)) !== null) {
      tds.push(tdm[1].replace(/<[^>]+>/g," ").replace(/&amp;/g,"&").replace(/\s{2,}/g," ").trim());
    }
    var linkMatch = /href="(\/anti-dumping-commission\/current-cases[^"]+)"/.exec(row);
    if (tds.length >= 6 && linkMatch) {
      cases.push({
        url:           "https://www.industry.gov.au" + linkMatch[1],
        caseNum:       tds[1],
        commodity:     tds[0],
        country:       tds[3],
        caseType:      tds[2],
        nextMilestone: tds[4] || "",
        lastUpdated:   tds[5] || "",
        status:        tds[6] || ""
      });
    }
  }
  Logger.log("   파싱: " + cases.length + "건");

  var newCount = 0, updateCount = 0;
  var koreanUpdates = [];

  cases.forEach(function(c) {
    var eventType = classifyAuEventType(c.caseType);
    var isKorean  = /korea/i.test(c.country);

    if (!tracked[c.url]) {
      // 신규 케이스
      var newRow = [c.url, c.caseNum, c.commodity, c.country, c.caseType, c.nextMilestone, c.lastUpdated, c.status, nowStr];
      trackSheet.appendRow(newRow);
      var lastRow = trackSheet.getLastRow();
      trackSheet.getRange(lastRow, 7).setNumberFormat("@").setValue(c.lastUpdated);
      trackSheet.getRange(lastRow, 1, 1, 9).setBackground("#dcfce7");
      tracked[c.url] = { rowIndex: lastRow, lastUpdated: c.lastUpdated, status: c.status };
      newCount++;

      // 신규 케이스 전부 메일 발송
      Logger.log("   🆕 신규: [" + c.caseNum + "] " + c.commodity + " | " + c.country);
      var auNewSummary = "";
      try {
        var auDetailRes = UrlFetchApp.fetch(
          "https://api.scraperapi.com?api_key=" + key + "&url=" + encodeURIComponent(c.url),
          { muteHttpExceptions: true }
        );
        if (auDetailRes.getResponseCode() === 200) {
          var auHtml = auDetailRes.getContentText("UTF-8");
          var auPdfRe = /href="(\/sites\/default\/files\/adc[^"]+\.pdf)"/gi;
          var auPm, auPdfs = [];
          while ((auPm = auPdfRe.exec(auHtml)) !== null) {
            auPdfs.push("https://www.industry.gov.au" + auPm[1]);
          }
          if (auPdfs.length > 0) {
            var getAuPdfNum = function(url) {
              var m1 = url.match(/\/\d+-(\d+)-[a-z]/i);   if (m1) return parseInt(m1[1]);
              var m2 = url.match(/\/\d+_-_(\d+)_-_/i);    if (m2) return parseInt(m2[1]);
              var m3 = url.match(/\/\d+---(\d+)---/i);     if (m3) return parseInt(m3[1]);
              return 0;
            };
            auPdfs.sort(function(a, b) { return getAuPdfNum(b) - getAuPdfNum(a); });
            Logger.log("   📄 호주 신규 PDF: " + auPdfs[0]);
            // 요약은 runIndiaSummaryAndNotify에서 처리
          }
        }
      } catch(e) { Logger.log("   ⚠️ 호주 요약 오류: " + e.message); }
      newCases.push({
        url: c.url, title: "[" + c.caseNum + "] " + c.commodity + " &larr; " + c.country,
        eventType: eventType, tradeType: "AD", country: "AU", agency: "ADC",
        published: c.lastUpdated, summary: ""
      });

    } else {
      var prev = tracked[c.url];
      if (c.lastUpdated !== prev.lastUpdated || c.status !== prev.status) {
        // 업데이트 감지
        Logger.log("   🔄 업데이트: [" + c.caseNum + "] " + c.commodity +
          " | " + prev.lastUpdated + " → " + c.lastUpdated +
          " | " + prev.status + " → " + c.status);
        var rowIdx = prev.rowIndex;
        trackSheet.getRange(rowIdx, 6).setValue(c.nextMilestone);
        trackSheet.getRange(rowIdx, 7).setNumberFormat("@").setValue(c.lastUpdated);
        trackSheet.getRange(rowIdx, 8).setValue(c.status);
        trackSheet.getRange(rowIdx, 9).setValue(nowStr);
        trackSheet.getRange(rowIdx, 1, 1, 9).setBackground("#FEF9C3");
        tracked[c.url].lastUpdated = c.lastUpdated;
        tracked[c.url].status      = c.status;
        updateCount++;

        // 한국 관련 케이스이면 별도 알림
        if (isKorean) {
          koreanUpdates.push({
            caseNum:       c.caseNum,
            commodity:     c.commodity,
            country:       c.country,
            caseType:      c.caseType,
            nextMilestone: c.nextMilestone,
            lastUpdated:   c.lastUpdated,
            prevUpdated:   prev.lastUpdated,
            status:        c.status,
            eventType:     eventType,
            url:           c.url
          });
        }

        // 업데이트된 케이스를 newCases에도 추가 (이메일 Section 2)
        // 요약은 runIndiaSummaryAndNotify에서 처리
        newCases.push({
          url: c.url, title: "[업데이트] [" + c.caseNum + "] " + c.commodity + " &larr; " + c.country,
          eventType: eventType, tradeType: "AD", country: "AU", agency: "ADC",
          published: c.lastUpdated, summary: ""
        });

      } else {
        // 변화 없음
        trackSheet.getRange(prev.rowIndex, 9).setValue(nowStr);
        trackSheet.getRange(prev.rowIndex, 1, 1, 9).setBackground("#FFFFFF");
      }
    }
  });

  // 사라진 케이스 감지
  var currentUrls = {};
  cases.forEach(function(c) { currentUrls[c.url] = true; });
  Object.keys(tracked).forEach(function(url) {
    if (!currentUrls[url]) {
      trackSheet.getRange(tracked[url].rowIndex, 1, 1, 9).setBackground("#fee2e2");
      Logger.log("   ⚠️ 종료: " + url);
    }
  });

  Logger.log("   → 신규: " + newCount + "건 / 업데이트: " + updateCount + "건 / 한국관련: " + koreanUpdates.length + "건");
  return koreanUpdates;
}

// =====================================================
// ── 글로벌 뉴스 수집 ──────────────────────────────────
// =====================================================



var SIGNAL_MAP = [
  // 한국어 무역구제
  { pattern: /제소/,                          score: 3 },
  { pattern: /반덤핑/,                        score: 3 },
  { pattern: /상계관세/,                      score: 3 },
  { pattern: /세이프가드/,                    score: 3 },
  { pattern: /무역구제/,                      score: 3 },
  { pattern: /수입.*급증|급증.*수입/,         score: 2 },
  { pattern: /관세.*부과|부과.*관세/,         score: 2 },
  { pattern: /덤핑.*조사|조사.*덤핑/,         score: 3 },
  { pattern: /조사.*개시|개시.*조사/,         score: 3 },
  { pattern: /보호무역/,                      score: 2 },
  { pattern: /수입규제/,                      score: 3 },
  { pattern: /쿼터.*축소|축소.*쿼터/,         score: 2 },
  { pattern: /덤핑방지/,                      score: 3 },
  { pattern: /긴급수입제한/,                  score: 3 },
  { pattern: /불공정.*수입|수입.*불공정/,     score: 2 },
  // 한국어 통상
  { pattern: /통상/,                          score: 2 },
  { pattern: /무역.*협정|협정.*무역/,         score: 2 },
  { pattern: /관세.*협상|협상.*관세/,         score: 2 },
  { pattern: /수출.*규제|규제.*수출/,         score: 2 },
  { pattern: /무역.*분쟁|분쟁.*무역/,         score: 2 },
  { pattern: /232조/,                         score: 3 },
  { pattern: /301조/,                         score: 3 },
  { pattern: /122조/,                         score: 3 },
  { pattern: /트럼프.*관세|관세.*트럼프/,     score: 3 },
  { pattern: /상호관세/,                      score: 3 },
  { pattern: /무역.*마찰|마찰.*무역/,         score: 2 },
  { pattern: /통상.*압력|압력.*통상/,         score: 2 },
  { pattern: /관세.*인상|인상.*관세/,         score: 2 },
  { pattern: /무역전쟁/,                      score: 3 },
  // 영어
  { pattern: /anti[\s-]dumping/i,            score: 3 },
  { pattern: /countervailing\s+dut/i,        score: 3 },
  { pattern: /safeguard\s+invest/i,          score: 3 },
  { pattern: /trade\s+remed/i,               score: 3 },
  { pattern: /petition\s+filed/i,            score: 3 },
  { pattern: /industry\s+complaint/i,        score: 3 },
  { pattern: /dumping\s+probe/i,             score: 2 },
  { pattern: /trade\s+probe/i,               score: 2 },
  { pattern: /trade\s+defense/i,             score: 2 },
  { pattern: /tariff\s+escalation/i,         score: 3 },
  { pattern: /tariff\s+invest/i,             score: 3 },
  { pattern: /section\s+301/i,               score: 3 },
  { pattern: /section\s+232/i,               score: 3 },
  { pattern: /dumping\s+margin/i,            score: 3 },
  { pattern: /trade\s+barrier/i,             score: 3 },
  { pattern: /import\s+surge/i,              score: 2 },
  { pattern: /import\s+quota/i,              score: 2 },
  { pattern: /steel\s+tariff/i,              score: 2 },
  { pattern: /aluminum\s+tariff/i,           score: 2 },
  { pattern: /trade\s+war/i,                 score: 2 },
  { pattern: /import\s+dut/i,                score: 2 },
  { pattern: /customs\s+dut/i,               score: 2 },
  { pattern: /protectionist/i,               score: 2 },
  { pattern: /trade\s+sanction/i,            score: 3 },
  { pattern: /tariff\s+hike/i,               score: 2 },
  { pattern: /trade\s+restriction/i,         score: 2 },
  { pattern: /import\s+restriction/i,        score: 2 },
  { pattern: /unfair\s+trade/i,              score: 3 },
  { pattern: /trade\s+dispute/i,             score: 2 },
  { pattern: /wto\s+dispute/i,               score: 3 },
  { pattern: /circumvention/i,               score: 3 },
  { pattern: /duty\s+evasion/i,              score: 3 },
  { pattern: /trade\s+deal/i,                score: 2 },
  { pattern: /trade\s+negotiat/i,            score: 2 },
  { pattern: /export\s+control/i,            score: 2 },
  { pattern: /reciprocal\s+tariff/i,         score: 3 }
];

var TRUMP_SIGNAL_MAP = [
  { pattern: /tariff/i,        score: 2 },
  { pattern: /trade/i,         score: 2 },
  { pattern: /dumping/i,       score: 2 },
  { pattern: /investigation/i, score: 2 },
  { pattern: /관세/,            score: 2 },
  { pattern: /무역/,            score: 2 }
];

function calcSignalScore(text, isTrump) {
  var score = 0;
  (isTrump ? TRUMP_SIGNAL_MAP : SIGNAL_MAP).forEach(function(s) {
    if (s.pattern.test(text)) score += s.score;
  });
  return score;
}

function initNewsSheet(sheet) {
  if (sheet.getLastRow() === 0) {
    var headers = ["ID","수집일시","출처","조사국","피조사국","무역구제유형","이벤트유형","제목","요약","시그널점수","공고일","원문링크"];
    sheet.appendRow(headers);
    var hr = sheet.getRange(1, 1, 1, headers.length);
    hr.setBackground("#1e293b"); hr.setFontColor("white"); hr.setFontWeight("bold");
    sheet.setFrozenRows(1);
    [120,140,120,100,100,80,80,300,400,70,90,300].forEach(function(w,i) { sheet.setColumnWidth(i+1, w); });
  }
}

function getNewsExistingIds(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return {};
  var ids = {};
  sheet.getRange(2, 1, lastRow - 1, 1).getValues().forEach(function(r) { ids[String(r[0])] = true; });
  return ids;
}

function getNewsRowColor(eventType) {
  if (eventType.indexOf("제소루머") !== -1) return "#FEF9C3";
  if (eventType.indexOf("조사개시") !== -1) return "#fee2e2";
  if (eventType.indexOf("예비판정") !== -1) return "#fef3c7";
  if (eventType.indexOf("최종판정") !== -1) return "#dbeafe";
  if (eventType.indexOf("관세부과") !== -1) return "#dcfce7";
  if (eventType.indexOf("일몰재심") !== -1) return "#fde68a";
  return "#ffffff";
}

function isWithin48Hours(pubDateStr) {
  if (!pubDateStr) return false;
  try {
    var pubTime = new Date(pubDateStr).getTime();
    var nowTime = new Date().getTime();
    if (isNaN(pubTime)) return false;
    var diffHours = (nowTime - pubTime) / (1000 * 60 * 60);
    return diffHours >= 0 && diffHours <= 48;
  } catch(e) { return false; }
}




function fetchTruthSocialPosts() {
  try {
    var res = UrlFetchApp.fetch(TRUTH_SOCIAL_URL, { headers: { "User-Agent": "TradeRemedyMonitor/2.0" }, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) { Logger.log("❌ Truth Social 오류: " + res.getResponseCode()); return []; }
    var data = JSON.parse(res.getContentText("UTF-8"));
    var posts = Array.isArray(data) ? data : (data.posts || data.data || []);
    var items = [], now = new Date().getTime();
    posts.forEach(function(post) {
      var content = (post.content || post.text || post.status || "").replace(/<[^>]+>/g," ").replace(/\s{2,}/g," ").trim();
      if (!content) return;
      var pubDate = post.created_at || post.date || post.published_at || "";
      try {
        var diffHours = (now - new Date(pubDate).getTime()) / (1000 * 60 * 60);
        if (!isNaN(new Date(pubDate).getTime()) && (diffHours < 0 || diffHours > 24)) return;
      } catch(e) {}
      var tradeKeywords = [
        'tariff', 'tariffs', 'import ban', 'import restriction', 'import fee',
        'duty', 'customs', 'trade', 'trade deal', 'trade war', 'trade deficit',
        'export', 'WTO', 'sanction', 'steel', 'aluminum', 'China'
      ];
      var hasTradeContent = tradeKeywords.some(function(kw) { return content.toLowerCase().indexOf(kw.toLowerCase()) !== -1; });
      if (!hasTradeContent) return;
      items.push({ url: post.url || post.uri || "https://truthsocial.com/@realDonaldTrump", pubDate: pubDate, fullText: content });
    });
    Logger.log("   Truth Social (24h): " + items.length + "건");
    return items;
  } catch(e) { Logger.log("❌ Truth Social 오류: " + e.message); return []; }
}

function summarizeTruthSocialPosts(posts) {
  if (!posts || posts.length === 0) return null;
  var postListText = posts.map(function(p, i) {
    return (i+1) + ". [" + (p.pubDate||"").substring(0,10) + "]\n" + p.fullText;
  }).join("\n\n");
  var payload = {
    contents: [{ parts: [{ text:
      "다음은 트럼프 대통령의 Truth Social 게시물 " + posts.length + "건입니다.\n\n" +
      "각 게시물을 번호(1., 2., ...)로 구분하여 한국어로 요약해주세요.\n" +
      "통상, 무역, 관세, 수입규제와 무관한 게시물은 요약에서 제외하고, 관련 게시물만 선별하여 요약해주세요.\n" +
      "무역, 관세, 통상 관련 내용이 있으면 특히 강조해주세요.\n" +
      "각 항목은 4~5줄 분량으로 핵심 내용과 배경, 의미를 충분히 담아 작성하세요.\n\n" +
      "게시물 목록:\n\n" + postListText
    }] }],
    generationConfig: { temperature: 0.1 },
    systemInstruction: { parts: [{ text: "마크다운 기호(**, *, #, _, - 등)를 절대 사용하지 마세요. 순수 텍스트로만 답변하세요." }] }
  };
  var summary = callGemini(payload);
  if (summary) Logger.log("✅ Truth Social 통합 요약 완료");
  return summary;
}

function analyzeNewsWithGemini(candidates) {
  if (!candidates || candidates.length === 0) return [];
  var listText = candidates.map(function(c, i) {
    return (i+1) + ". [출처: " + c.source + "] [점수: " + c.signalScore + "점]\n   제목: " + c.title + "\n   날짜: " + c.pubDate + "\n   URL: " + c.url;
  }).join("\n\n");

  var prompt =
    "당신은 무역구제 분야 20년 경력의 전문가입니다.\n\n" +
    "아래 뉴스 기사 목록을 분석하여 다음 두 가지를 수행하세요:\n" +
    "1. 무역구제(반덤핑, 상계관세, 세이프가드) 및 통상(관세, 무역협정, 수출규제, 무역분쟁, 232조, 301조, 122조, 트럼프 관세, 상호관세)과 무관한 기사는 제외\n" +
    "2. 동일하거나 매우 유사한 사건을 다룬 기사들은 하나로 그룹핑하여 대표 1건으로 통합\n" +
    "3. 기사 날짜(pub_date)가 오늘로부터 7일 이상 지난 것은 제외\n" +
    "4. pub_date가 비어있거나 불분명한 기사는 제외\n\n" +
    "JSON 배열 형식으로만 응답하세요. 다른 텍스트는 절대 포함하지 마세요.\n\n" +
    "이벤트 유형은 반드시 다음 중 하나: 제소루머, 조사개시, 예비판정, 최종판정, 관세부과, 일몰재심, 통상협상, 수출규제, 기타\n\n" +
    "응답 형식:\n[\n  {\n    \"event_title\": \"한국어 간결한 사건 제목\",\n    \"investigating_country\": \"조사국\",\n    \"target_country\": \"피조사국\",\n    \"trade_type\": \"AD 또는 CVD 또는 SG 또는 관세\",\n    \"event_type\": \"이벤트유형\",\n    \"summary\": \"한국어 3줄 요약\",\n    \"url\": \"대표 기사 URL\",\n    \"source\": \"출처 매체명\",\n    \"pub_date\": \"YYYY-MM-DD\",\n    \"signal_score\": 숫자,\n    \"is_trump_sns\": false,\n    \"grouped_count\": 그룹핑 수\n  }\n]\n\nis_trump_sns는 항상 false로 고정하세요. Google News 기사는 절대 SNS 발언으로 분류하지 마세요.\n\n" +
    "분석할 기사 목록 (" + candidates.length + "건):\n\n" + listText;

  var payload = { contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.1, maxOutputTokens: 65536 } };
  var raw = callGemini(payload);
  if (!raw) { Logger.log("❌ Gemini 분석 실패"); return []; }
  try {
    var clean = raw.replace(/```json/g,"").replace(/```/g,"").trim();
    var start = clean.indexOf("[");
    if (start === -1) return [];
    var jsonStr = clean.substring(start);
    var end = jsonStr.lastIndexOf("}");
    if (end === -1) return [];
    var parsed = JSON.parse(jsonStr.substring(0, end + 1) + "]");
    Logger.log("✅ Gemini 분석 완료: " + parsed.length + "건");
    return parsed;
  } catch(e) { Logger.log("❌ JSON 파싱 오류: " + e.message); return []; }
}

// 반환값: newsItems 배열
function runTradeNewsCollector(ss) {
  var sheet = ss.getSheetByName(NEWS_SHEET_NAME) || ss.insertSheet(NEWS_SHEET_NAME);
  initNewsSheet(sheet);
  var existingIds = getNewsExistingIds(sheet);
  var nowStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm");
  var newsItems = [], saved = 0;

  // 1. Truth Social
  Logger.log("📡 Truth Social 수집 중...");
  var trumpPosts = fetchTruthSocialPosts();
  if (trumpPosts.length > 0) {
    var newTrumpPosts = [];
    trumpPosts.forEach(function(post) {
      var id = makeHash("TRUMP-POST-" + post.url + "-" + post.pubDate);
      if (existingIds[id]) return;
      var pubDate = (post.pubDate || "").substring(0, 10);
      var shortText = post.fullText.substring(0, 200).replace(/"/g, "'");
      var row = [id, nowStr, "Trump/Truth Social", "미국", "", "관세", "SNS발언", shortText, "", 0, pubDate, post.url];
      sheet.appendRow(row);
      sheet.getRange(sheet.getLastRow(), 1, 1, row.length).setBackground("#FEF9C3");
      existingIds[id] = true;
      newTrumpPosts.push({ id: id, rowNum: sheet.getLastRow(), post: post });
      saved++;
    });

    if (newTrumpPosts.length > 0) {
      var combinedSummary = summarizeTruthSocialPosts(newTrumpPosts.map(function(p) { return p.post; }));
      if (combinedSummary && newTrumpPosts[0]) {
        sheet.getRange(newTrumpPosts[0].rowNum, 9).setValue(combinedSummary.replace(/"/g,"'"));
      }
      var latestPost = newTrumpPosts[0].post;
      newsItems.push({
        event_title: "Trump Truth Social 발언 " + newTrumpPosts.length + "건",
        investigating_country: "미국", target_country: "", trade_type: "관세", event_type: "SNS발언",
        summary: combinedSummary || "", url: latestPost.url, source: "Trump/Truth Social",
        pub_date: (latestPost.pubDate || "").substring(0, 10), signal_score: 0,
        is_trump_sns: true, grouped_count: newTrumpPosts.length
      });
    }
  }

// 2. 네이버 뉴스 API (한국어) + FT RSS (영어) 수집
  Logger.log("📡 네이버 API + FT 수집 중...");
  var allNewsItems = [], seenUrls = {};
  var naverCount = 0, ftCount = 0;

  var naverClientId     = PropertiesService.getScriptProperties().getProperty("NAVER_CLIENT_ID");
  var naverClientSecret = PropertiesService.getScriptProperties().getProperty("NAVER_CLIENT_SECRET");

  // ── 2-1. 네이버 뉴스 API (5건씩 청크) ────────────
  var NAVER_QUERIES = [
    // 무역구제
    "반덤핑 조사",    "상계관세 조사",  "세이프가드 조사",
    "무역구제 조사",  "한국산 반덤핑",  "수입규제 조사",
    "반덤핑 제소",    "덤핑 조사",      "무역위원회 조사",
    // 통상
    "통상 협상",      "무역 협정",      "통상 마찰",
    "232조 관세",     "301조 관세",     "122조 관세",
    "트럼프 관세",    "상호관세",       "수출 규제",
    "무역 분쟁",      "보호무역 관세"
  ];

  var NAVER_TITLE_KEYWORDS = [
    // 무역구제
    "반덤핑", "상계관세", "세이프가드", "무역구제",
    "수입규제", "덤핑", "관세 부과", "제소", "조사 개시",
    "무역위원회", "보호무역", "쿼터", "수입 규제",
    "덤핑방지", "긴급수입제한", "불공정 수입",
    // 통상
    "통상", "무역 협정", "무역협정", "관세 협상",
    "수출 규제", "무역 분쟁", "무역분쟁",
    "232조", "301조", "122조",
    "트럼프 관세", "상호관세", "무역 마찰", "통상 마찰",
    "수출 통제", "무역전쟁", "관세 인상", "통상 압력",
    "무역 협상", "FTA", "통상 리스크"
  ];

  if (naverClientId && naverClientSecret) {
    var naverChunkSize = 5;
    for (var ni = 0; ni < NAVER_QUERIES.length; ni += naverChunkSize) {
      var naverChunk = NAVER_QUERIES.slice(ni, ni + naverChunkSize);
      var naverRequests = naverChunk.map(function(q) {
        return {
          url: "https://openapi.naver.com/v1/search/news.json?query=" + encodeURIComponent(q) + "&display=10&sort=date",
          headers: { "X-Naver-Client-Id": naverClientId, "X-Naver-Client-Secret": naverClientSecret },
          muteHttpExceptions: true
        };
      });
      var naverResponses = UrlFetchApp.fetchAll(naverRequests);

      naverResponses.forEach(function(res, i) {
        if (res.getResponseCode() === 429) { Logger.log("⚠️ 네이버 429: " + naverChunk[i]); return; }
        if (res.getResponseCode() !== 200) return;
        var data;
        try { data = JSON.parse(res.getContentText("UTF-8")); } catch(e) { return; }
        (data.items || []).forEach(function(item) {
          var title = (item.title || "")
            .replace(/<[^>]+>/g,"").replace(/&quot;/g,'"')
            .replace(/&amp;/g,"&").replace(/&lt;/g,"<").replace(/&gt;/g,">").trim();
          var url = item.originallink || item.link || "";
          var pubDate = item.pubDate || "";
          if (!title || !url || seenUrls[url]) return;
          if (!isWithin48Hours(pubDate)) return;
          var hasKeyword = NAVER_TITLE_KEYWORDS.some(function(kw) {
            return title.indexOf(kw) !== -1;
          });
          if (!hasKeyword) return;
          seenUrls[url] = true;
          allNewsItems.push({
            title: title, url: url, pubDate: pubDate,
            source: "네이버뉴스", lang: "ko", isTrump: false,
            signalScore: calcSignalScore(title, false)
          });
          naverCount++;
        });
      });

      if (ni + naverChunkSize < NAVER_QUERIES.length) Utilities.sleep(500);
    }
    Logger.log("   → 네이버: " + naverCount + "건");
  } else {
    Logger.log("⚠️ 네이버 API 키 없음 - 스킵");
  }

  // ── 2-2. FT RSS (영어) ───────────────────────────
  var FT_FEEDS = [
    { url: "https://www.ft.com/world?format=rss",           source: "FT" },
    { url: "https://www.ft.com/global-economy?format=rss",  source: "FT" },
    { url: "https://www.ft.com/us-economy?format=rss",      source: "FT" },
    { url: "https://www.ft.com/markets?format=rss",         source: "FT" }
  ];

  var ftRequests = FT_FEEDS.map(function(f) {
    return { url: f.url, muteHttpExceptions: true };
  });
  var ftResponses = UrlFetchApp.fetchAll(ftRequests);

  ftResponses.forEach(function(res, i) {
    if (res.getResponseCode() !== 200) return;
    var xml = res.getContentText("UTF-8");
    var itemRe = /<item>([\s\S]*?)<\/item>/g, m;
    while ((m = itemRe.exec(xml)) !== null) {
      var block = m[1];
      var title = (/<title><!\[CDATA\[([\s\S]*?)\]\]><\/title>/.exec(block)
                || /<title>([\s\S]*?)<\/title>/.exec(block) || ["",""])[1]
                .replace(/<!\[CDATA\[|\]\]>/g,"").replace(/<[^>]+>/g,"").trim();
      var lm = /<link>([\s\S]*?)<\/link>/.exec(block);
      var url = lm ? lm[1].trim() : "";
      var pubDate = (/<pubDate>([\s\S]*?)<\/pubDate>/.exec(block) || ["",""])[1].trim();
      if (!title || !url || seenUrls[url]) continue;
      if (!isWithin48Hours(pubDate)) continue;
      var score = calcSignalScore(title, false);
      if (score < 2) continue;
      seenUrls[url] = true;
      allNewsItems.push({
        title: title, url: url, pubDate: pubDate,
        source: "FT", lang: "en", isTrump: false, signalScore: score
      });
      ftCount++;
    }
  });
  Logger.log("   → FT: " + ftCount + "건");
  Logger.log("📊 뉴스 총 후보: " + allNewsItems.length + "건");

  // 4. Gemini 분석
  if (allNewsItems.length > 0) {
    allNewsItems.sort(function(a, b) { return b.signalScore - a.signalScore; });
    Utilities.sleep(2000);
    var results = analyzeNewsWithGemini(allNewsItems.slice(0, 100));
    if (results && results.length > 0) {
      results.forEach(function(r) {
        if (!r.url) return;
        var id = makeHash("NEWS-" + r.url);
        if (existingIds[id]) return;
        var pubDate = (r.pub_date || "").substring(0, 10);
        var titleWithCount = r.grouped_count > 1 ? (r.event_title||"") + " [" + r.grouped_count + "건 통합]" : (r.event_title||"");
        var row = [id, nowStr, r.source||"", r.investigating_country||"", r.target_country||"",
          r.trade_type||"", r.event_type||"", titleWithCount.replace(/"/g,"'"),
          (r.summary||"").replace(/"/g,"'"), r.signal_score||0, pubDate, r.url];
        sheet.appendRow(row);
        sheet.getRange(sheet.getLastRow(), 1, 1, row.length).setBackground(getNewsRowColor(r.event_type||""));
        existingIds[id] = true;
        r.event_title = titleWithCount;
        newsItems.push(r);
        saved++;
      });
      Logger.log("✅ 뉴스 저장: " + results.length + "건");
    }
  }

  Logger.log("✅ 뉴스 수집 완료: " + saved + "건");
  return newsItems;
}

// =====================================================
// ── 통합 이메일 발송 ──────────────────────────────────
// =====================================================

function buildNewsSection(newsItems) {
  if (!newsItems || newsItems.length === 0) return "";
  var bodyHtml = "";

  var trumpItems  = newsItems.filter(function(c) { return c.is_trump_sns; });
  var normalItems = newsItems.filter(function(c) { return !c.is_trump_sns; });
  var sections = [];
  if (trumpItems.length  > 0) sections.push({ label: "Trump / Truth Social", items: trumpItems,  isTrump: true  });
  if (normalItems.length > 0) sections.push({ label: "글로벌 무역구제 뉴스",   items: normalItems, isTrump: false });

  sections.forEach(function(sec) {
    var secCount = sec.isTrump
      ? sec.items.reduce(function(s, c) { return s + (c.grouped_count || 1); }, 0)
      : sec.items.length;

    bodyHtml +=
      "<tr>" +
        "<td colspan=\"2\" bgcolor=\"#F5F5F5\" style=\"padding:8px 16px;border-top:2px solid #86BC25;\">" +
          "<font face=\"Arial, sans-serif\" size=\"2\" color=\"#222222\"><b>" + sec.label + "</b></font>" +
          " <font face=\"Arial, sans-serif\" size=\"1\" color=\"#999999\">(" + secCount + "건)</font>" +
        "</td>" +
      "</tr>";

    sec.items.forEach(function(c) {
      var tradeType  = c.trade_type            || "";
      var eventType  = c.event_type            || "";
      var pubDate    = (c.pub_date || "").substring(0, 10);
      var source     = c.source               || "";
      var countryLine = (c.investigating_country && c.target_country)
        ? c.investigating_country + " → " + c.target_country
        : c.investigating_country || c.target_country || "";

      var safeTitle   = (c.event_title || c.title || "").replace(/&(?!amp;|lt;|gt;|quot;|#)/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
      var safeSummary = (c.summary || "").replace(/&(?!amp;|lt;|gt;|quot;|#)/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/\n/g,"<br>");

      var titlePrefix = c.is_trump_sns
        ? "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#DC2626\"><b>[SNS 발언]</b></font><br>"
        : "";

      var postBody = c.is_trump_sns
        ? "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#555555\">" + safeTitle + "</font>"
        : "<font face=\"Arial, sans-serif\" size=\"2\"><a href=\"" + (c.url||"#") + "\" style=\"color:#1A1A1A;text-decoration:underline;\">" + safeTitle + "</a></font>";

      bodyHtml +=
        "<tr bgcolor=\"#FFFFFF\">" +
          "<td valign=\"top\" width=\"110\" bgcolor=\"#FFFFFF\" style=\"padding:10px 8px 4px 16px;\">" +
            "<font face=\"Arial, sans-serif\" size=\"2\" color=\"#111111\"><b>" + tradeType + "</b></font><br>" +
            "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#555555\">" + eventType + "</font><br>" +
            "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#888888\">" + countryLine + "</font><br>" +
            "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#AAAAAA\">" + pubDate + "</font>" +
          "</td>" +
          "<td valign=\"top\" bgcolor=\"#FFFFFF\" style=\"padding:10px 16px 4px 8px;\">" +
            titlePrefix + postBody + "<br>" +
            "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#999999\">" + source + "</font>" +
          "</td>" +
        "</tr>";

      if (safeSummary) {
        bodyHtml +=
          "<tr bgcolor=\"#FFFFFF\">" +
            "<td bgcolor=\"#FFFFFF\" style=\"padding:0 8px 0 16px;\"></td>" +
            "<td bgcolor=\"#F7FBF2\" style=\"padding:6px 16px 12px 12px;border-left:3px solid #86BC25;border-bottom:1px solid #EEEEEE;\">" +
              "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#86BC25\"><b>AI 요약</b></font><br>" +
              "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#444444\">" + safeSummary + "</font>" +
            "</td>" +
          "</tr>";
      } else {
        bodyHtml += "<tr bgcolor=\"#FFFFFF\"><td colspan=\"2\" style=\"padding:0;border-bottom:1px solid #EEEEEE;\"></td></tr>";
      }
    });
  });
  return bodyHtml;
}

function buildNewCasesSection(newCases) {
  if (!newCases || newCases.length === 0) return "";
  var bodyHtml = "";

  var sections = [
    { label: "인도 — DGTR",             filterFn: function(c) { return c.country === "IN"; } },
    { label: "한국 — KTC",              filterFn: function(c) { return c.country === "KR"; } },
    { label: "미국 — DOC/ITA",         filterFn: function(c) { return c.country === "US" && c.agency === "DOC/ITA"; } },
    { label: "미국 — CBP EAPA",        filterFn: function(c) { return c.agency  === "CBP" && c.tradeType === "EAPA"; } },
    { label: "미국 — CBP CSMS",        filterFn: function(c) { return c.agency  === "CBP" && c.tradeType === "CSMS"; } },
    { label: "EU — European Commission", filterFn: function(c) { return c.country === "EU"; } },
    { label: "브라질 — DECOM",          filterFn: function(c) { return c.country === "BR"; } },
    { label: "호주 — ADC",              filterFn: function(c) { return c.country === "AU"; } }
  ];

  sections.forEach(function(sec) {
    var items = newCases.filter(sec.filterFn);
    if (items.length === 0) return;

    bodyHtml +=
      "<tr>" +
        "<td colspan=\"2\" bgcolor=\"#F5F5F5\" style=\"padding:8px 16px;border-top:2px solid #86BC25;\">" +
          "<font face=\"Arial, sans-serif\" size=\"2\" color=\"#222222\"><b>" + sec.label + "</b></font>" +
          " <font face=\"Arial, sans-serif\" size=\"1\" color=\"#999999\">(" + items.length + "건)</font>" +
        "</td>" +
      "</tr>";

    items.forEach(function(c) {
      var pubDate = c.published ? c.published.substring(0, 10) : "";
      var safeTitle = (c.title || "").replace(/&(?!amp;|lt;|gt;|quot;|larr;|#)/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
      var safeSummary = c.summary
        ? c.summary.replace(/&(?!amp;|lt;|gt;|quot;|#)/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/\n/g,"<br>")
        : "";

      bodyHtml +=
        "<tr bgcolor=\"#FFFFFF\">" +
          "<td valign=\"top\" width=\"110\" bgcolor=\"#FFFFFF\" style=\"padding:10px 8px 6px 16px;\">" +
            "<font face=\"Arial, sans-serif\" size=\"2\" color=\"#111111\"><b>" + c.country + " / " + c.agency + "</b></font><br>" +
            "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#666666\">" + (c.tradeType||"") + "</font><br>" +
            "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#AAAAAA\">" + pubDate + "</font>" +
          "</td>" +
          "<td valign=\"top\" bgcolor=\"#FFFFFF\" style=\"padding:10px 16px 6px 8px;\">" +
            "<font face=\"Arial, sans-serif\" size=\"2\"><a href=\"" + c.url + "\" style=\"color:#1A1A1A;text-decoration:underline;\">" + safeTitle + "</a></font>" +
          "</td>" +
        "</tr>";

      if (safeSummary) {
        bodyHtml +=
          "<tr bgcolor=\"#FFFFFF\">" +
            "<td bgcolor=\"#FFFFFF\" style=\"padding:0 8px 0 16px;\"></td>" +
            "<td bgcolor=\"#F7FBF2\" style=\"padding:6px 16px 12px 10px;border-left:3px solid #86BC25;border-bottom:1px solid #EEEEEE;\">" +
              "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#86BC25\"><b>AI 요약</b></font><br>" +
              "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#444444\">" + safeSummary + "</font>" +
            "</td>" +
          "</tr>";
      } else {
        bodyHtml += "<tr bgcolor=\"#FFFFFF\"><td colspan=\"2\" style=\"padding:0;border-bottom:1px solid #EEEEEE;\"></td></tr>";
      }
    });
  });
  return bodyHtml;
}

function buildKoreanUpdatesSection(euUpdates, brUpdates) {
  var allUpdates = (euUpdates || []).concat(brUpdates || []);
  if (allUpdates.length === 0) return "";
  var bodyHtml = "";

  // EU 업데이트
  if (euUpdates && euUpdates.length > 0) {
    bodyHtml +=
      "<tr>" +
        "<td colspan=\"2\" bgcolor=\"#F5F5F5\" style=\"padding:8px 16px;border-top:2px solid #86BC25;\">" +
          "<font face=\"Arial, sans-serif\" size=\"2\" color=\"#222222\"><b>EU — 한국 관련 케이스 업데이트</b></font>" +
          " <font face=\"Arial, sans-serif\" size=\"1\" color=\"#999999\">(" + euUpdates.length + "건)</font>" +
        "</td>" +
      "</tr>";

    euUpdates.forEach(function(c) {
      var safeShortName = (c.shortName||"").replace(/&(?!amp;|lt;|gt;|quot;|#)/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
      var safeSummary = (c.summary||"").replace(/&(?!amp;|lt;|gt;|quot;|#)/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/\n/g,"<br>");

      bodyHtml +=
        "<tr bgcolor=\"#FFFFFF\">" +
          "<td valign=\"top\" width=\"110\" bgcolor=\"#FFFFFF\" style=\"padding:10px 8px 4px 16px;\">" +
            "<font face=\"Arial, sans-serif\" size=\"2\" color=\"#111111\"><b>" + (c.tradeType||"") + "</b></font><br>" +
            "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#555555\">" + (c.eventType||"") + "</font><br>" +
            "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#888888\">" + (c.countries||"") + "</font><br>" +
            "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#AAAAAA\">" + (c.pubDate||"") + "</font>" +
          "</td>" +
          "<td valign=\"top\" bgcolor=\"#FFFFFF\" style=\"padding:10px 16px 4px 8px;\">" +
            "<font face=\"Arial, sans-serif\" size=\"2\"><a href=\"" + c.detailUrl + "\" style=\"color:#1A1A1A;text-decoration:underline;\">[" + c.caseNum + "] " + safeShortName + "</a></font><br>" +
            "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#999999\">" + (c.pubType||"") + "</font>" +
          "</td>" +
        "</tr>";

      if (safeSummary) {
        bodyHtml +=
          "<tr bgcolor=\"#FFFFFF\">" +
            "<td bgcolor=\"#FFFFFF\" style=\"padding:0 8px 0 16px;\"></td>" +
            "<td bgcolor=\"#F7FBF2\" style=\"padding:6px 16px 12px 10px;border-left:3px solid #86BC25;border-bottom:1px solid #EEEEEE;\">" +
              "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#86BC25\"><b>AI 요약</b></font><br>" +
              "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#444444\">" + safeSummary + "</font>" +
            "</td>" +
          "</tr>";
      } else {
        bodyHtml += "<tr bgcolor=\"#FFFFFF\"><td colspan=\"2\" style=\"padding:0;border-bottom:1px solid #EEEEEE;\"></td></tr>";
      }
    });
  }

  // 브라질 업데이트
  if (brUpdates && brUpdates.length > 0) {
    bodyHtml +=
      "<tr>" +
        "<td colspan=\"2\" bgcolor=\"#F5F5F5\" style=\"padding:8px 16px;border-top:2px solid #86BC25;\">" +
          "<font face=\"Arial, sans-serif\" size=\"2\" color=\"#222222\"><b>브라질 DECOM — 한국 관련 케이스 업데이트</b></font>" +
          " <font face=\"Arial, sans-serif\" size=\"1\" color=\"#999999\">(" + brUpdates.length + "건)</font>" +
        "</td>" +
      "</tr>";

    brUpdates.forEach(function(u) {
      var safeProduto  = (u.produto||"").replace(/&(?!amp;|lt;|gt;|quot;|#)/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
      var safeOrigem   = (u.origem ||"").replace(/&(?!amp;|lt;|gt;|quot;|#)/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
      var safeInvestig = (u.investig||u.medida||"").replace(/&(?!amp;|lt;|gt;|quot;|#)/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
      var typeText     = u.type === "조치" ? "[조치]" : "[조사]";

      bodyHtml +=
        "<tr bgcolor=\"#FFFFFF\">" +
          "<td valign=\"top\" width=\"110\" bgcolor=\"#FFFFFF\" style=\"padding:10px 8px 4px 16px;\">" +
            "<font face=\"Arial, sans-serif\" size=\"2\" color=\"#111111\"><b>" + typeText + "</b></font><br>" +
            "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#555555\">" + safeInvestig + "</font><br>" +
            "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#888888\">" + safeOrigem + "</font>" +
          "</td>" +
          "<td valign=\"top\" bgcolor=\"#FFFFFF\" style=\"padding:10px 16px 4px 8px;\">" +
            "<font face=\"Arial, sans-serif\" size=\"2\"><a href=\"" + u.url + "\" style=\"color:#1A1A1A;text-decoration:underline;\">" + safeProduto + "</a></font><br>" +
            "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#999999\">업데이트: " + u.prevUpdated + " → " + u.newUpdated + "</font>" +
          "</td>" +
        "</tr>";

      if (u.type !== "조치" && u.situacao) {
        var safeSituacao     = u.situacao.replace(/&(?!amp;|lt;|gt;|quot;|#)/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
        var safePrevSituacao = (u.prevSituacao||"").replace(/&(?!amp;|lt;|gt;|quot;|#)/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
        bodyHtml +=
          "<tr bgcolor=\"#FFFFFF\">" +
            "<td bgcolor=\"#FFFFFF\" style=\"padding:0 8px 0 16px;\"></td>" +
            "<td bgcolor=\"#F7FBF2\" style=\"padding:6px 16px 12px 10px;border-left:3px solid #86BC25;border-bottom:1px solid #EEEEEE;\">" +
              "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#86BC25\"><b>현재 상황 (Situação Atual)</b></font><br>" +
              (safePrevSituacao ? "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#999999\">이전: " + safePrevSituacao + "</font><br>" : "") +
              "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#444444\">현재: " + safeSituacao + "</font>" +
            "</td>" +
          "</tr>";
      } else {
        bodyHtml += "<tr bgcolor=\"#FFFFFF\"><td colspan=\"2\" style=\"padding:0;border-bottom:1px solid #EEEEEE;\"></td></tr>";
      }
    });
  }

  return bodyHtml;
}

function sendCombinedEmail(newsItems, newCases, euKoreanUpdates, brKoreanUpdates) {
  var today = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd");

  var newsCount    = newsItems   ? newsItems.length   : 0;
  var casesCount   = newCases    ? newCases.length     : 0;
  var euUpdCount   = euKoreanUpdates ? euKoreanUpdates.length : 0;
  var brUpdCount   = brKoreanUpdates ? brKoreanUpdates.length : 0;
  var hasContent   = newsCount > 0 || casesCount > 0 || euUpdCount > 0 || brUpdCount > 0;

  if (!hasContent) {
    Logger.log("✅ 발송할 내용 없음 - 이메일 생략");
    return;
  }

  // ── 섹션별 본문 생성 ──────────────────────────────
  var sec1Html = newsCount   > 0 ? buildNewsSection(newsItems)                        : "";
  var sec2Html = (casesCount > 0 || euUpdCount > 0 || brUpdCount > 0)
    ? buildNewCasesSection(newCases) + buildKoreanUpdatesSection(euKoreanUpdates, brKoreanUpdates)
    : "";
  var sec3Html = "";

  // 섹션 헤더 타이틀 행 생성 함수
  function sectionTitle(title) {
    return "<tr><td colspan=\"2\" bgcolor=\"#000000\" style=\"padding:6px 16px;\">" +
      "<font face=\"Arial, sans-serif\" size=\"2\" color=\"#86BC25\"><b>" + title + "</b></font>" +
      "</td></tr>";
  }

  var bodyHtml = "";
  if (sec1Html) bodyHtml += sectionTitle("Section 1. 글로벌 통상 뉴스") + sec1Html;
  if (sec2Html) bodyHtml += sectionTitle("Section 2. 반덤핑/상계관세/세이프가드 모니터링") + sec2Html;

  // ── 요약 배너 텍스트 ──────────────────────────────
  var summaryParts = [];
  if (newsCount  > 0) summaryParts.push("뉴스 " + newsCount + "건");
  if (casesCount > 0) summaryParts.push("신규케이스 " + casesCount + "건");
  if (euUpdCount + brUpdCount > 0) summaryParts.push("한국관련 " + (euUpdCount + brUpdCount) + "건");
  var summaryText = summaryParts.join(" &nbsp;|&nbsp; ");

  var html =
    "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\">" +
    "<html xmlns=\"http://www.w3.org/1999/xhtml\">" +
    "<head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\" />" +
    "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\" /></head>" +
    "<body bgcolor=\"#E8E8E8\" style=\"margin:0;padding:0;\">" +
    "<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" bgcolor=\"#E8E8E8\">" +
    "<tr><td align=\"center\" style=\"padding:24px 0;\">" +
    "<table width=\"600\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" bgcolor=\"#FFFFFF\">" +

    // 헤더
    "<tr><td colspan=\"2\" bgcolor=\"#000000\" style=\"padding:22px 24px 18px 24px;\">" +
      "<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tr>" +
        "<td valign=\"bottom\">" +
          "<font face=\"Arial, sans-serif\" color=\"#FFFFFF\"><b><font size=\"4\">Trade Remedy Intelligence</font></b></font><br>" +
          "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#86BC25\">Trade Compliance &amp; Strategy</font>"
        "</td>" +
        "<td align=\"right\" valign=\"bottom\">" +
          "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#777777\">" + today + "</font>" +
        "</td>" +
      "</tr></table>" +
    "</td></tr>" +
    "<tr><td colspan=\"2\" bgcolor=\"#86BC25\" style=\"font-size:1px;line-height:4px;height:4px;\">&nbsp;</td></tr>" +

    // 요약 배너
    "<tr><td colspan=\"2\" bgcolor=\"#FAFAFA\" style=\"padding:12px 24px;border-bottom:1px solid #E0E0E0;\">" +
      "<font face=\"Arial, sans-serif\" size=\"2\" color=\"#555555\">Daily Trade Intelligence&nbsp;&nbsp;</font>" +
      "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#BBBBBB\">" + summaryText + "</font>" +
    "</td></tr>" +

    // 컬럼 헤더
    "<tr bgcolor=\"#F0F0F0\">" +
      "<td width=\"110\" bgcolor=\"#F0F0F0\" style=\"padding:7px 8px 7px 16px;border-bottom:1px solid #CCCCCC;\">" +
        "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#666666\"><b>유형 / 국가</b></font>" +
      "</td>" +
      "<td bgcolor=\"#F0F0F0\" style=\"padding:7px 16px 7px 8px;border-bottom:1px solid #CCCCCC;\">" +
        "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#666666\"><b>제목</b></font>" +
      "</td>" +
    "</tr>" +

    bodyHtml +

    // 푸터 링크
    "<tr><td colspan=\"2\" bgcolor=\"#F5F5F5\" style=\"padding:12px 24px;border-top:1px solid #E0E0E0;\">" +
      "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#555555\">원본 데이터 확인: &nbsp;</font>" +
      "<font face=\"Arial, sans-serif\" size=\"1\">" +
        "<a href=\"YOUR_GOOGLE_SHEETS_URL\" style=\"color:#86BC25;text-decoration:underline;\">Trade Remedy Monitoring Dashboard &rarr;</a>" 
+
      "</font>" +
    "</td></tr>" +
    "<tr><td colspan=\"2\" bgcolor=\"#86BC25\" style=\"font-size:1px;line-height:3px;height:3px;\">&nbsp;</td></tr>" +
    "<tr><td colspan=\"2\" bgcolor=\"#000000\" style=\"padding:16px 24px;\">" +
      "<font face=\"Arial, sans-serif\" size=\"1\" color=\"#777777\">" +
        "&copy; 2026 <font color=\"#86BC25\"><b>Trade Remedy Intelligence</b></font>" +
        " &nbsp;|&nbsp; Automated Trade Monitoring System" +
      "</font>" +
    "</td></tr>" +

    "</table></td></tr></table></body></html>";

  var subject = today + " Report from Deloitte Trade Risk Intelligence";

  CONFIG_V2.EMAIL_RECIPIENTS.forEach(function(email) {
    GmailApp.sendEmail(email, subject, "", {
      htmlBody: html, name: "Deloitte Trade Risk Intelligence", charset: "UTF-8"
    });
    Logger.log("✅ 통합 메일 발송 → " + email);
  });
}

// =====================================================
// ── 메인 실행 ─────────────────────────────────────────
// =====================================================

function runMonitorV2() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var sheet  = ss.getSheetByName(CONFIG_V2.SHEET_NAME) || ss.insertSheet(CONFIG_V2.SHEET_NAME);
  initSheet(sheet);

  var existingIds      = getExistingIds(sheet);
  var newCases         = [];
  var newIndiaCases    = [];
  var newUsCases       = [];
  var newKrNoticeCases = [];
  var newEapaCases     = [];
  var newCsmsCases     = [];
  var nowStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm");

  // ── 1. 글로벌 뉴스 수집 ───────────────────────────
  var newsItems = [];
  try {
    newsItems = runTradeNewsCollector(ss) || [];
  } catch(e) { Logger.log("❌ 뉴스 수집 오류: " + e.message); }

  // ── 2. 인도 DGTR (신규/업데이트 통합) ────────────────
  Logger.log("📡 IN-DGTR");
  try {
    var IN_FEEDS_DEF = [
      { url: "https://www.dgtr.gov.in/en/anti-dumping-investigation-in-india",  type: "AD"  },
      { url: "https://www.dgtr.gov.in/en/countervailing-duty-investigation",     type: "CVD" },
      { url: "https://www.dgtr.gov.in/en/safe-guard-investigation-in-india",     type: "SG"  }
    ];

    var reAdCvd = /<div class="investigation">[\s\S]*?<span class="status">[^<]+<\/span>[\s\S]*?<a href="(\/en\/[^"]+)"[^>]*>([\s\S]*?)<\/a>/g;
    var reSg    = /<div class="investigation">[\s\S]*?<p><span class="status">[^<]+<\/span>\s*<a href="(\/en\/[^"]+)"[^>]*>([\s\S]*?)<\/a>/g;

    // 추적 시트 초기화
    var inTrackSheet = ss.getSheetByName("인도 DGTR 추적") || ss.insertSheet("인도 DGTR 추적");
    if (inTrackSheet.getLastRow() === 0) {
      var inHeaders = ["케이스URL", "제목", "최신S.No", "최신PDF URL", "마지막확인일시"];
      inTrackSheet.appendRow(inHeaders);
      var inHr = inTrackSheet.getRange(1, 1, 1, inHeaders.length);
      inHr.setBackground("#1e293b"); inHr.setFontColor("white"); inHr.setFontWeight("bold");
      inTrackSheet.setFrozenRows(1);
      [300,300,80,400,140].forEach(function(w,i) { inTrackSheet.setColumnWidth(i+1, w); });
    }

    // 기존 추적 데이터 로드
    var inTracked = {};
    if (inTrackSheet.getLastRow() > 1) {
      inTrackSheet.getRange(2, 1, inTrackSheet.getLastRow()-1, 5).getValues()
        .forEach(function(r, i) {
          if (r[0]) inTracked[String(r[0])] = {
            rowIndex: i + 2,
            sno: Number(r[2]) || 0,
            pdf: String(r[3] || "")
          };
        });
    }
    Logger.log("   기존 등록: " + Object.keys(inTracked).length + "건");

    var inNewCount = 0, inUpdateCount = 0;

    IN_FEEDS_DEF.forEach(function(feed) {
      // 1페이지만 수집
      var res = UrlFetchApp.fetch(feed.url + "?page=0", {
        headers: { "User-Agent": "TradeRemedyMonitor/2.0" },
        muteHttpExceptions: true
      });
      if (res.getResponseCode() !== 200) return;
      var html = res.getContentText("UTF-8");

      var re = feed.type === "SG" ? reSg : reAdCvd;
      re.lastIndex = 0;
      var items = [], m;
      while ((m = re.exec(html)) !== null) {
        var url = "https://www.dgtr.gov.in" + m[1].trim();
        var title = m[2].replace(/&quot;/g,'"').replace(/&amp;/g,'&').replace(/&#039;/g,"'").replace(/<[^>]+>/g,"").trim();
        items.push({ url: url, title: title, type: feed.type });
      }

      // 상세페이지 5건씩 청크로 병렬 수집
      var inChunkSize = 5;
      var itemsToCheck = items.slice(0, 20);
      for (var ci = 0; ci < itemsToCheck.length; ci += inChunkSize) {
        var chunk = itemsToCheck.slice(ci, ci + inChunkSize);
        var detailRequests = chunk.map(function(item) {
          return { url: item.url, headers: { "User-Agent": "TradeRemedyMonitor/2.0" }, muteHttpExceptions: true };
        });
        var detailResponses = UrlFetchApp.fetchAll(detailRequests);

        detailResponses.forEach(function(detailRes, i) {
          var item = chunk[i];
          var latestSno = 0, latestPdf = "";

          if (detailRes.getResponseCode() === 200) {
            var detailHtml = detailRes.getContentText("UTF-8");
            var rowRe = /<tr>[\s\S]*?<td>(\d+)<\/td>[\s\S]*?href="(https?:\/\/dgtr\.gov\.in[^"]+\.pdf)"[\s\S]*?<\/tr>/g;
            var rm;
            while ((rm = rowRe.exec(detailHtml)) !== null) {
              var sno = parseInt(rm[1]);
              if (sno > latestSno) { latestSno = sno; latestPdf = rm[2]; }
            }
          }

          var tradeType = item.type;
          var eventType = classifyEventIN(item.title);

          if (!inTracked[item.url]) {
            // 시트에 없는 URL → S.No로 신규/업데이트 판단
            inTrackSheet.appendRow([item.url, item.title, latestSno, latestPdf, nowStr]);
            inTrackSheet.getRange(inTrackSheet.getLastRow(), 1, 1, 5).setBackground("#dcfce7");
            inTracked[item.url] = { rowIndex: inTrackSheet.getLastRow(), sno: latestSno, pdf: latestPdf };

            if (latestSno === 1) {
              // 진짜 신규 조사건
              Logger.log("   🆕 신규: " + item.title.substring(0, 50));
              var o = { url: item.url, title: item.title, eventType: "조사개시", tradeType: tradeType, country: "IN", agency: "DGTR", published: "" };
              newCases.push(o); newIndiaCases.push(o); inNewCount++;

              // 메인 시트에도 등록
              var id = makeHash(item.url);
              if (!existingIds[id]) {
                var row = [id, nowStr, "IN", "DGTR", "조사개시", tradeType, item.title, "", item.url];
                sheet.appendRow(row);
                sheet.getRange(sheet.getLastRow(), 1, 1, row.length).setBackground(getRowColor("조사개시"));
                sheet.getRange(sheet.getLastRow(), 9).setFormula('=HYPERLINK("' + item.url + '","원문보기")');
                existingIds[id] = true;
              }
            }
            // S.No > 1인 미등록 케이스는 조용히 시트에만 등록

          } else {
            // 시트에 있는 URL → S.No 비교
            var prevSno = inTracked[item.url].sno;
            if (latestSno > prevSno) {
              Logger.log("   🔄 업데이트: S.No " + prevSno + " → " + latestSno + " | " + item.title.substring(0, 40));

              var rowIdx = inTracked[item.url].rowIndex;
              inTrackSheet.getRange(rowIdx, 3).setValue(latestSno);
              inTrackSheet.getRange(rowIdx, 4).setValue(latestPdf);
              inTrackSheet.getRange(rowIdx, 5).setValue(nowStr);
              inTrackSheet.getRange(rowIdx, 1, 1, 5).setBackground("#FEF9C3");
              inTracked[item.url].sno = latestSno;

              var o2 = { url: item.url, title: "[업데이트] " + item.title, eventType: eventType, tradeType: tradeType, country: "IN", agency: "DGTR", published: "" };
              newCases.push(o2); newIndiaCases.push(o2); inUpdateCount++;

              var id2 = makeHash("IN-update-" + item.url + "-sno" + latestSno);
              if (!existingIds[id2]) {
                var row2 = [id2, nowStr, "IN", "DGTR", eventType, tradeType, "[업데이트] " + item.title, "", item.url];
                sheet.appendRow(row2);
                sheet.getRange(sheet.getLastRow(), 1, 1, row2.length).setBackground(getRowColor(eventType));
                sheet.getRange(sheet.getLastRow(), 9).setFormula('=HYPERLINK("' + item.url + '","원문보기")');
                existingIds[id2] = true;
              }
            } else {
              // 변화 없음
              var rowIdx2 = inTracked[item.url].rowIndex;
              inTrackSheet.getRange(rowIdx2, 5).setValue(nowStr);
              inTrackSheet.getRange(rowIdx2, 1, 1, 5).setBackground("#FFFFFF");
            }
          }
        }); // detailResponses.forEach 끝

        
      } // 청크 루프 끝
    });

    Logger.log("   → 신규: " + inNewCount + "건 / 업데이트: " + inUpdateCount + "건");
  } catch(e) { Logger.log("❌ IN-DGTR: " + e.message); }

  // ── 3. 한국 KTC ───────────────────────────────────
  Logger.log("📡 KR-KTC");
  try {
    var krCount = 0, noticeCount = 0;
    var ktcKey = PropertiesService.getScriptProperties().getProperty("SCRAPER_API_KEY");

    // 조사 목록 + 공고 목록 병렬 수집
    var ktcRequests = [
      {
        url: "https://api.scraperapi.com?api_key=" + ktcKey + "&url=" + encodeURIComponent(
          "https://www.ktc.go.kr/investArticle.do?menuId=11&pageIndex=1&process_step_code=11000005&invstg_type_code=10100001"
        ),
        muteHttpExceptions: true
      },
      {
        url: "https://api.scraperapi.com?api_key=" + ktcKey + "&url=" + encodeURIComponent(
          "https://www.ktc.go.kr/boardList.do?bbs_id=0&menuId=46&bbsTypeCode=&searchCondition=S&searchKeyword=&pageIndex=1"
        ),
        muteHttpExceptions: true
      }
    ];
    var ktcResponses = UrlFetchApp.fetchAll(ktcRequests);

    // 조사 목록 파싱
    if (ktcResponses[0].getResponseCode() === 200) {
      var reKtc = /viewInvest\('(\d+)','10100001'\)[\s\S]{0,200}?<i>([^<]+)<\/i>[\s\S]{0,300}?<td>([^<]+)<\/td>[\s\S]{0,100}?<td>([^<]+)<\/td>[\s\S]{0,100}?<td>([^<]+)<\/td>/g;
      var mKtc;
      while ((mKtc = reKtc.exec(ktcResponses[0].getContentText("UTF-8"))) !== null) {
        var caseId = mKtc[1], title = mKtc[2].trim(), caseNo = mKtc[3].trim(), startDate = mKtc[5].trim();
        var url = "https://www.ktc.go.kr/viewInvest.do?masterId=" + caseId + "&invstg_type_code=10100001";
        var id = makeHash("KR-invest-" + caseId);
        if (existingIds[id]) continue;
        var eventType = /종료재심/.test(title) ? "일몰재심" : /재심/.test(title) ? "행정재심" : "조사개시";
        var row = [id, nowStr, "KR", "KTC", eventType, "AD", title + " [" + caseNo + "]", startDate, url];
        sheet.appendRow(row);
        sheet.getRange(sheet.getLastRow(), 1, 1, row.length).setBackground(getRowColor(eventType));
        sheet.getRange(sheet.getLastRow(), 9).setFormula('=HYPERLINK("' + url + '","원문보기")');
        existingIds[id] = true;
        newCases.push({ url: url, title: title + " [" + caseNo + "]", eventType: eventType, tradeType: "AD", country: "KR", agency: "KTC", published: startDate });
        krCount++;
      }
    }

    // 공고 목록 파싱
    if (ktcResponses[1].getResponseCode() === 200) {
      var reKtc2 = /boardView\('(\d+)'\)[\s\S]{0,200}?<i>([^<]+)<\/i>/g, m2;
      while ((m2 = reKtc2.exec(ktcResponses[1].getContentText("UTF-8"))) !== null) {
        var boardId = m2[1], title2 = m2[2].trim();
        var url2 = "https://www.ktc.go.kr/boardView.do?bbs_id=" + boardId;
        var id2 = makeHash("KR-notice-" + boardId);
        if (existingIds[id2]) continue;
        var tradeType2 = /보조금|상계/.test(title2) ? "CVD" : /세이프가드/.test(title2) ? "SG" : "AD";
        var eventType2 = /잠정덤핑|잠정관세/.test(title2) ? "예비판정" : /부과.*규칙|부과.*고시/.test(title2) ? "최종판정" : /재심/.test(title2) ? "행정재심" : "일반공고";
        var row2 = [id2, nowStr, "KR", "KTC", eventType2, tradeType2, title2, "", url2];
        sheet.appendRow(row2);
        sheet.getRange(sheet.getLastRow(), 1, 1, row2.length).setBackground(getRowColor(eventType2));
        sheet.getRange(sheet.getLastRow(), 9).setFormula('=HYPERLINK("' + url2 + '","원문보기")');
        existingIds[id2] = true;
        var o2 = { url: url2, title: title2, eventType: eventType2, tradeType: tradeType2, country: "KR", agency: "KTC", published: "", bbsId: boardId };
        newCases.push(o2); newKrNoticeCases.push(o2);
        krCount++; noticeCount++;
      }
    }

    Logger.log("   → KTC 합계: " + krCount + "건 (공고 " + noticeCount + "건)");
  } catch(e) { Logger.log("❌ KR-KTC: " + e.message); }
  if (newKrNoticeCases.length > 0) summarizeNewKrNoticeCases(newKrNoticeCases);

  // ── 4. 미국 DOC/ITA ───────────────────────────────
  Logger.log("📡 US-TRADEGOV");
  try {
    var res3 = UrlFetchApp.fetch(US_URL, { headers: { "User-Agent": "TradeRemedyMonitor/2.0" }, muteHttpExceptions: true });
    if (res3.getResponseCode() === 200) {
      var re3 = /ita-static-cards__title">([^<]+)<\/div>[\s\S]{0,200}?ita-static-cards__summary">([^<]+)<\/div>[\s\S]{0,200}?href="([^"]+)"/g;
      var m3, count3 = 0;
      while ((m3 = re3.exec(res3.getContentText("UTF-8"))) !== null) {
        var title3 = m3[1].trim().replace(/&amp;#039;/g,"'"), path3 = m3[3].trim();
        var url3 = path3.indexOf("http") === 0 ? path3 : "https://www.trade.gov" + path3;
        var id3 = makeHash("US-tradegov-" + path3);
        if (existingIds[id3]) continue;
        var tradeType3 = /countervailing/i.test(title3) ? "CVD" : /safeguard/i.test(title3) ? "SG" : "AD";
        var eventType3 = /final/i.test(title3) ? "최종판정" : /preliminary/i.test(title3) ? "예비판정" : /initiat/i.test(title3) ? "조사개시" : "일반공고";
        var row3 = [id3, nowStr, "US", "DOC/ITA", eventType3, tradeType3, title3, m3[2].trim(), url3];
        sheet.appendRow(row3);
        sheet.getRange(sheet.getLastRow(), 1, 1, row3.length).setBackground(getRowColor(eventType3));
        sheet.getRange(sheet.getLastRow(), 9).setFormula('=HYPERLINK("' + url3 + '","원문보기")');
        existingIds[id3] = true;
        var o3 = { url: url3, title: title3, eventType: eventType3, tradeType: tradeType3, country: "US", agency: "DOC/ITA", published: m3[2].trim() };
        newCases.push(o3); newUsCases.push(o3); count3++;
      }
      Logger.log("   → " + count3 + "건");
    }
  } catch(e) { Logger.log("❌ US-TRADEGOV: " + e.message); }
  if (newUsCases.length > 0) summarizeNewUsCases(newUsCases);

  // ── 5. CBP EAPA ───────────────────────────────────
  Logger.log("📡 US-CBP-EAPA");
  try {
    var eapaCount = 0;
    fetchCbpEapaCases().forEach(function(c) {
      var id = makeHash("US-CBP-EAPA-" + c.caseNum);
      if (existingIds[id]) return;
      var title = c.caseId + ": " + c.company;
      var row = [id, nowStr, "US", "CBP", "조사개시", "EAPA", title, c.date, c.url];
      sheet.appendRow(row);
      sheet.getRange(sheet.getLastRow(), 1, 1, row.length).setBackground(getRowColor("조사개시"));
      sheet.getRange(sheet.getLastRow(), 9).setFormula('=HYPERLINK("' + c.url + '","원문보기")');
      existingIds[id] = true;
      var o = { url: c.url, title: title, eventType: "조사개시", tradeType: "EAPA", country: "US", agency: "CBP", published: c.date };
      newCases.push(o); newEapaCases.push(o); eapaCount++;
    });
    Logger.log("   → " + eapaCount + "건");
  } catch(e) { Logger.log("❌ US-CBP-EAPA: " + e.message); }
  if (newEapaCases.length > 0) summarizeNewEapaCases(newEapaCases);

  // ── 6. CBP CSMS ───────────────────────────────────
  Logger.log("📡 US-CBP-CSMS");
  try {
    var csmsCount = 0;
    fetchCsmsMessages().forEach(function(item) {
      var id = makeHash("US-CBP-CSMS-" + item.subject);
      if (existingIds[id]) return;
      var dm = item.pub_date.match(/^(\d{2}\/\d{2}\/\d{4})/);
      var dateStr = dm ? dm[1] : item.pub_date;
      var row = [id, nowStr, "US", "CBP", "일반공고", "CSMS", item.subject, dateStr, item.href];
      sheet.appendRow(row);
      sheet.getRange(sheet.getLastRow(), 1, 1, row.length).setBackground("#ffffff");
      sheet.getRange(sheet.getLastRow(), 9).setFormula('=HYPERLINK("' + item.href + '","원문보기")');
      existingIds[id] = true;
      var o = { url: item.href, title: item.subject, eventType: "일반공고", tradeType: "CSMS", country: "US", agency: "CBP", published: dateStr };
      newCases.push(o); newCsmsCases.push(o); csmsCount++;
    });
    Logger.log("   → " + csmsCount + "건");
  } catch(e) { Logger.log("❌ US-CBP-CSMS: " + e.message); }
  if (newCsmsCases.length > 0) summarizeNewCsmsCases(newCsmsCases);

  // ── 7. EU TRON ────────────────────────────────────
  var euKoreanUpdates = [];
  try {
    euKoreanUpdates = runEuCaseTracker(ss, nowStr, newCases) || [];
  } catch(e) { Logger.log("❌ EU-TRON: " + e.message); }

  // ── 8. 브라질 DECOM ───────────────────────────────
  var brKoreanUpdates = [];
  try {
    brKoreanUpdates = fetchBrazilCases(ss, nowStr, newCases) || [];
  } catch(e) { Logger.log("❌ Brazil-DECOM: " + e.message); PropertiesService.getScriptProperties().setProperty("MONITOR_BR_FAILED", "Y"); }

  // ── 9. 호주 ADC ───────────────────────────────────
  var auKoreanUpdates = [];
  try {
    auKoreanUpdates = fetchAustraliaCases(ss, nowStr, newCases) || [];
  } catch(e) { Logger.log("❌ AU-ADC: " + e.message); PropertiesService.getScriptProperties().setProperty("MONITOR_AU_FAILED", "Y"); }

  // ── 10. 결과 저장 ─────────────────────────────────  ← 번호 변경
  Logger.log("✅ 뉴스: " + newsItems.length + "건 / 신규케이스: " + newCases.length + "건 / EU업데이트: " + euKoreanUpdates.length + "건 / BR업데이트: " + brKoreanUpdates.length + "건 / AU업데이트: " + auKoreanUpdates.length + "건");
  var props = PropertiesService.getScriptProperties();
  props.setProperty("MONITOR_NEWS",     JSON.stringify(newsItems));
  props.setProperty("MONITOR_CASES",    JSON.stringify(newCases));
  props.setProperty("MONITOR_EU",       JSON.stringify(euKoreanUpdates));
  props.setProperty("MONITOR_BR",       JSON.stringify(brKoreanUpdates));
  props.setProperty("MONITOR_AU_CASES", JSON.stringify(
    newCases.filter(function(c) { return c.country === "AU"; })
  ));
  props.setProperty("MONITOR_IN_CASES", JSON.stringify(newIndiaCases));
  props.setProperty("MONITOR_DONE",     "Y");
  Logger.log("✅ 데이터 저장 완료 - 트리거에서 요약 후 발송 예정");
}

// =====================================================
// ── 트리거 설정 ───────────────────────────────────────
// =====================================================

function setTriggerV2() {
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });

  ScriptApp.newTrigger("runMonitorV2")
    .timeBased().everyDays(1).atHour(5).nearMinute(45)
    .inTimezone("Asia/Seoul").create();

  ScriptApp.newTrigger("runIndiaSummaryAndNotify")
    .timeBased().everyDays(1).atHour(5).nearMinute(55)
    .inTimezone("Asia/Seoul").create();

  Logger.log("✅ 트리거 설정: 메인 5시45분 / 인도요약+발송 5시55분");
}

// ── Gemini 테스트 ─────────────────────────────────────
function testGemini() {
  var apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  var res = UrlFetchApp.fetch(
    "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent",
    { method: "POST", contentType: "application/json", headers: { "x-goog-api-key": apiKey },
      payload: JSON.stringify({ contents: [{ parts: [{ text: "안녕하세요. 한국어로 짧게 인사해주세요." }] }] }),
      muteHttpExceptions: true }
  );
  Logger.log("HTTP: " + res.getResponseCode());
  Logger.log(res.getContentText().substring(0, 500));
}

function testKtcParsing() {
  var key = PropertiesService.getScriptProperties().getProperty("SCRAPER_API_KEY");
  var res = UrlFetchApp.fetch(
    "https://api.scraperapi.com?api_key=" + key +
    "&render=true" +   // JS 렌더링 추가
    "&url=" + encodeURIComponent(
      "https://www.ktc.go.kr/investArticle.do?menuId=11&pageIndex=1&process_step_code=11000005&invstg_type_code=10100001"
    ),
    { muteHttpExceptions: true }
  );
  Logger.log("응답코드: " + res.getResponseCode());
  var html = res.getContentText("UTF-8");
  Logger.log("HTML 길이: " + html.length);
  Logger.log("HTML 미리보기: " + html.substring(0, 500));

  var reKtc = /viewInvest\('(\d+)','10100001'\)[\s\S]{0,200}?<i>([^<]+)<\/i>[\s\S]{0,300}?<td>([^<]+)<\/td>[\s\S]{0,100}?<td>([^<]+)<\/td>[\s\S]{0,100}?<td>([^<]+)<\/td>/g;
  var mKtc, count = 0;
  while ((mKtc = reKtc.exec(html)) !== null) {
    Logger.log("masterId=" + mKtc[1] + " / 제목=" + mKtc[2].trim() + " / 개시일=" + mKtc[5].trim());
    count++;
  }
  Logger.log("파싱된 케이스: " + count + "건");
}


function testKtcProcFileList2() {
  var key = PropertiesService.getScriptProperties().getProperty("SCRAPER_API_KEY");

  var res = UrlFetchApp.fetch(
    "https://api.scraperapi.com?api_key=" + key +
    "&url=" + encodeURIComponent(
      "https://www.ktc.go.kr/procFileListView.do?invstg_mast_id=00000856&invstg_seq_no=0003"
    ),
    { muteHttpExceptions: true }
  );
  var html = res.getContentText("UTF-8");
  Logger.log("전체 길이: " + html.length + "자");

  // 파일 관련 키워드 검색
  var keywords = ["pdf", "PDF", "hwp", "HWP", "fileDown", "download", "fileName", "file_name", "orgFileName", "seq_no", "첨부", "파일"];
  keywords.forEach(function(kw) {
    var idx = html.indexOf(kw);
    if (idx !== -1) {
      Logger.log("키워드 [" + kw + "] 발견:");
      Logger.log(html.substring(Math.max(0, idx-50), idx + 400));
      Logger.log("---");
    }
  });

  // HTML 중간부분 출력
  Logger.log("=== HTML 중간 (1000~3000자) ===");
  Logger.log(html.substring(1000, 3000));
}

function testKtcNoticeSummary() {
  var boardId = "00002104";

  Logger.log("=== 1. 공고 상세 HTML 수집 ===");
  var html = fetchKtcBoardDetail(boardId);
  if (!html) { Logger.log("❌ HTML 수집 실패"); return; }
  Logger.log("✅ HTML 수집 성공, 길이: " + html.length);

  Logger.log("=== 2. PDF 링크 추출 ===");
  var pdfUrl = extractPdfFromKtcBoard(boardId, html);
  Logger.log("PDF URL: " + pdfUrl);

  Logger.log("=== 3. 텍스트 추출 ===");
  var text = extractTextFromKtcBoard(html);
  Logger.log("텍스트 길이: " + (text ? text.length : 0));
  if (text) Logger.log("텍스트 미리보기: " + text.substring(0, 200));

  Logger.log("=== 4. AI 요약 시도 ===");
  if (pdfUrl) {
    Logger.log("PDF 요약 시도: " + pdfUrl);
    var s = summarizeWithGeminiPdfViaScraper(pdfUrl, "이 PDF는 한국 무역위원회 공고야. 조사 대상 품목, 피조사국, 핵심 내용을 한국어로 3줄로 요약해줘.");
    Logger.log("PDF 요약 결과: " + s);
  } else if (text && text.length >= 50) {
    Logger.log("텍스트 요약 시도");
    var s2 = summarizeWithGeminiText(text, "이 내용은 한국 무역위원회 공고야. 조사 대상 품목, 피조사국, 핵심 내용을 한국어로 3줄로 요약해줘.\n\n");
    Logger.log("텍스트 요약 결과: " + s2);
  } else {
    Logger.log("❌ PDF도 없고 텍스트도 부족");
  }
}

function testKtcPdfDownload() {
  var key = PropertiesService.getScriptProperties().getProperty("SCRAPER_API_KEY");
  var sessionNum = "456";

  // 1. 먼저 공고 페이지 접근해서 세션 수립
  UrlFetchApp.fetch(
    "https://api.scraperapi.com?api_key=" + key +
    "&url=" + encodeURIComponent(
      "https://www.ktc.go.kr/boardView.do?bbs_id=00002104&menuId=46&pageIndex=1"
    ) + "&session_number=" + sessionNum,
    { muteHttpExceptions: true }
  );

  // 2. POST 방식으로 PDF 다운로드 시도
  var pdfParams = "bbs_id=00002104&seq_no=2";
  var res = UrlFetchApp.fetch(
    "https://api.scraperapi.com?api_key=" + key +
    "&url=" + encodeURIComponent("https://www.ktc.go.kr/boardFileDownload.do") +
    "&method=POST&body=" + encodeURIComponent(pdfParams) +
    "&session_number=" + sessionNum,
    { muteHttpExceptions: true }
  );
  Logger.log("응답코드: " + res.getResponseCode());
  Logger.log("응답 내용: " + res.getContentText("UTF-8").substring(0, 500));
  Logger.log("PDF 크기: " + res.getContent().length + " bytes");
}

function runIndiaSummaryAndNotify() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty("MONITOR_DONE") !== "Y") {
    Logger.log("❌ 메인 수집이 완료되지 않음");
    return;
  }

  // 재시도 횟수 확인 (최대 3회)
  var retryCount = parseInt(props.getProperty("INDIA_RETRY_COUNT") || "0");
  if (retryCount >= 3) {
    Logger.log("❌ 최대 재시도 횟수(3회) 초과 - 인도 요약 없이 이메일 발송");
    props.deleteProperty("INDIA_RETRY_COUNT");
    var newsItems       = JSON.parse(props.getProperty("MONITOR_NEWS")  || "[]");
    var newCases        = JSON.parse(props.getProperty("MONITOR_CASES") || "[]");
    var euKoreanUpdates = JSON.parse(props.getProperty("MONITOR_EU")    || "[]");
    var brKoreanUpdates = JSON.parse(props.getProperty("MONITOR_BR")    || "[]");
    sendCombinedEmail(newsItems, newCases, euKoreanUpdates, brKoreanUpdates);
    props.deleteProperty("MONITOR_NEWS");
    props.deleteProperty("MONITOR_CASES");
    props.deleteProperty("MONITOR_EU");
    props.deleteProperty("MONITOR_BR");
    props.deleteProperty("MONITOR_IN_CASES");
    props.deleteProperty("MONITOR_DONE");
    return;
  }

  try {
    // 저장된 데이터 불러오기
    var newsItems       = JSON.parse(props.getProperty("MONITOR_NEWS")     || "[]");
    var newCases        = JSON.parse(props.getProperty("MONITOR_CASES")    || "[]");
    var euKoreanUpdates = JSON.parse(props.getProperty("MONITOR_EU")       || "[]");
    var brKoreanUpdates = JSON.parse(props.getProperty("MONITOR_BR")       || "[]");
    var newIndiaCases   = JSON.parse(props.getProperty("MONITOR_IN_CASES") || "[]");

    // 호주 재시도
    if (props.getProperty("MONITOR_AU_FAILED") === "Y") {
      Logger.log("📡 호주 ADC 재시도 수집");
      try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var auRetryUpdates = fetchAustraliaCases(ss, Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm"), newCases);
        props.deleteProperty("MONITOR_AU_FAILED");
        props.setProperty("MONITOR_CASES", JSON.stringify(newCases));
        props.setProperty("MONITOR_AU_CASES", JSON.stringify(newCases.filter(function(c) { return c.country === "AU"; })));
        Logger.log("✅ 호주 재시도 완료");
      } catch(e) { Logger.log("⚠️ 호주 재시도 실패: " + e.message); }
    }

    // 브라질 재시도
    if (props.getProperty("MONITOR_BR_FAILED") === "Y") {
      Logger.log("📡 브라질 DECOM 재시도 수집");
      try {
        var ss = ss || SpreadsheetApp.getActiveSpreadsheet();
        var brRetry = fetchBrazilCases(ss, Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm"), newCases);
        brKoreanUpdates = brKoreanUpdates.concat(brRetry);
        props.deleteProperty("MONITOR_BR_FAILED");
        props.setProperty("MONITOR_CASES", JSON.stringify(newCases));
        props.setProperty("MONITOR_BR", JSON.stringify(brKoreanUpdates));
        Logger.log("✅ 브라질 재시도 완료");
      } catch(e) { Logger.log("⚠️ 브라질 재시도 실패: " + e.message); }
    }

    // 인도 PDF 요약 실행
    if (newIndiaCases.length > 0) {
      var targetCases = newIndiaCases;
      Logger.log("📡 인도 PDF 요약 시작: " + targetCases.length + "건");

      // 1단계: 상세 페이지 병렬 수집
      var pageRequests = targetCases.map(function(c) {
        return { url: c.url, headers: { "User-Agent": "TradeRemedyMonitor/2.0" }, muteHttpExceptions: true };
      });
      var pageResponses = UrlFetchApp.fetchAll(pageRequests);

      // PDF URL 추출 (S.No 기반)
      var pdfInfos = [];
      pageResponses.forEach(function(res, i) {
        if (res.getResponseCode() !== 200) return;
        var html = res.getContentText("UTF-8");
        var rowRe = /<tr>[\s\S]*?<td>(\d+)<\/td>[\s\S]*?href="(https?:\/\/dgtr\.gov\.in[^"]+\.pdf)"[\s\S]*?<\/tr>/g;
        var m, latestPdf = null, latestSno = -1;
        while ((m = rowRe.exec(html)) !== null) {
          var sno = parseInt(m[1]);
          if (sno > latestSno) { latestSno = sno; latestPdf = m[2]; }
        }
        if (latestPdf) pdfInfos.push({ index: i, pdfUrl: latestPdf });
      });

      // 2단계: PDF 병렬 다운로드
      if (pdfInfos.length > 0) {
        var pdfRequests = pdfInfos.map(function(p) {
          return { url: p.pdfUrl, headers: { "User-Agent": "TradeRemedyMonitor/2.0" }, muteHttpExceptions: true };
        });
        var pdfResponses = UrlFetchApp.fetchAll(pdfRequests);

        // 3단계: Gemini 요약 (순차)
        pdfResponses.forEach(function(pdfRes, j) {
          if (pdfRes.getResponseCode() !== 200) return;
          var targetIndex = pdfInfos[j].index;
          var payload = {
            contents: [{ parts: [
              { inline_data: { mime_type: "application/pdf", data: Utilities.base64Encode(pdfRes.getContent()) } },
              { text: "이 PDF는 인도 무역구제 공문이야. 조사 대상 품목, 피조사국, 핵심 내용을 한국어로 3줄로 요약해줘." }
            ]}],
            generationConfig: { temperature: 0.1 },
            systemInstruction: { parts: [{ text: "마크다운 기호(**, *, #, _ 등)를 절대 사용하지 마세요. 순수 텍스트로만 답변하세요." }] }
          };
          var s = callGemini(payload);
          if (s) {
            targetCases[targetIndex].summary = s;
            Logger.log("✅ 인도 PDF 요약 완료 (" + (j+1) + "/" + pdfResponses.length + ")");
          }
        });

        // 요약된 인도 케이스를 newCases에 반영
        targetCases.forEach(function(c) {
          if (!c.summary) return;
          for (var i = 0; i < newCases.length; i++) {
            if (newCases[i].url === c.url) { newCases[i].summary = c.summary; break; }
          }
        });
      }
    }
    // 호주 요약 (케이스 메타데이터 기반 - 추가 HTTP 요청 없음)
    var newAuCases = JSON.parse(props.getProperty("MONITOR_AU_CASES") || "[]");
    if (newAuCases.length > 0) {
      Logger.log("📡 호주 요약 시작: " + newAuCases.length + "건");
      newAuCases.forEach(function(c) {
        try {
          // title에서 케이스 정보 추출 (예: "[693] Steel ← Korea")
          var titleText = (c.title || "").replace(/&larr;/g, "←").replace(/&amp;/g, "&");
          var prompt = "호주 반덤핑위원회(ADC) 무역구제 케이스 정보:\n" +
            "케이스 제목: " + titleText + "\n" +
            "조사 유형: " + (c.eventType || "") + "\n" +
            "무역구제 유형: " + (c.tradeType || "") + "\n" +
            "공고일: " + (c.published || "") + "\n\n" +
            "위 정보를 바탕으로 이 케이스의 의미와 핵심 내용을 한국어로 2줄로 설명해줘.";
          var s = summarizeWithGeminiText("", prompt);
          if (s) {
            for (var i = 0; i < newCases.length; i++) {
              if (newCases[i].url === c.url) { newCases[i].summary = s; break; }
            }
            Logger.log("✅ 호주 요약 완료: " + c.url);
          }
        } catch(e) { Logger.log("⚠️ 호주 요약 오류: " + e.message); }
      });
    }
    // 통합 이메일 발송
    sendCombinedEmail(newsItems, newCases, euKoreanUpdates, brKoreanUpdates);

    // 완료 후 저장 데이터 초기화
    props.deleteProperty("MONITOR_NEWS");
    props.deleteProperty("MONITOR_CASES");
    props.deleteProperty("MONITOR_EU");
    props.deleteProperty("MONITOR_BR");
    props.deleteProperty("MONITOR_AU_CASES");
    props.deleteProperty("MONITOR_IN_CASES");
    props.deleteProperty("MONITOR_DONE");
    props.deleteProperty("INDIA_RETRY_COUNT");
    Logger.log("✅ 이메일 발송 완료 및 임시 데이터 초기화");

  } catch(e) {
    // 오류 발생 시 재시도 예약
    retryCount++;
    props.setProperty("INDIA_RETRY_COUNT", String(retryCount));
    Logger.log("❌ 오류 발생 (" + retryCount + "회): " + e.message);
    Logger.log("⏰ 5분 후 재시도 예약");

    // 기존 재시도 트리거 삭제 후 새로 등록
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getHandlerFunction() === "runIndiaSummaryAndNotify") {
        ScriptApp.deleteTrigger(t);
      }
    });
    ScriptApp.newTrigger("runIndiaSummaryAndNotify")
      .timeBased().after(5 * 60 * 1000)
      .create();
  }
}


