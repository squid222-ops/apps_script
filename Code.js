/**
 * 매일 아침 전일 급등주(한국거래소 KRX 공개 시세) 및 대표뉴스(Google RSS)를 수집하여
 * 스프레드시트에 기록하고 HTML 브리핑 메일을 발송하는 메인 함수
 * (외부 API 키 미사용 / 거래대금 항목 제외)
 */
function runMorningBriefing() {
  // 주말(토: 6, 일: 0)에는 실행하지 않고 종료
  const now = new Date();
  const day = now.getDay();
  if (day === 0 || day === 6) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("설정_및_요약");
  const briefingSheet = ss.getSheetByName("급등주_브리핑");
  const logSheet = ss.getSheetByName("발송_이력");

  // 1. '설정_및_요약' 시트의 설정값 위치(B16~B22)에서 값 읽기
  const recipientEmail = String(configSheet.getRange("B16").getValue()).trim();

  const rawUpper = Number(configSheet.getRange("B18").getValue());
  const raw15 = Number(configSheet.getRange("B19").getValue());
  const raw8 = Number(configSheet.getRange("B20").getValue());

  const limitUpper = rawUpper
    ? rawUpper < 1
      ? rawUpper * 100
      : rawUpper
    : 29.5;
  const limit15 = raw15 ? (raw15 < 1 ? raw15 * 100 : raw15) : 15.0;
  const limit8 = raw8 ? (raw8 < 1 ? raw8 * 100 : raw8) : 8.0;

  const maxItemsPerGroup = Number(configSheet.getRange("B21").getValue()) || 7;
  const maxNewsPerItem = Number(configSheet.getRange("B22").getValue()) || 2;

  const timestamp = Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd HH:mm");

  // 2. 한국거래소(KRX) 최근 영업일 전 종목 시세 수집
  const krxResult = fetchKrxRecentRiseStocks(limit8);
  const allStocks = krxResult.stocks;

  // 3. 노이즈 필터링 (우선주, 스팩, ETF/ETN 제외) 및 등락률 내림차순 정렬
  const filteredStocks = allStocks
    .filter((stock) => isValidStock(stock))
    .sort((a, b) => b.changeRate - a.changeRate);

  // 4. 3단계 등급 분류 및 등급별 상위 7종목 제한
  const upperGroup = filteredStocks
    .filter((s) => s.changeRate >= limitUpper)
    .slice(0, maxItemsPerGroup)
    .map((s, idx) => ({ ...s, grade: "상한가", rank: idx + 1 }));

  const midGroup = filteredStocks
    .filter((s) => s.changeRate >= limit15 && s.changeRate < limitUpper)
    .slice(0, maxItemsPerGroup)
    .map((s, idx) => ({ ...s, grade: "15% 이상", rank: idx + 1 }));

  const lowGroup = filteredStocks
    .filter((s) => s.changeRate >= limit8 && s.changeRate < limit15)
    .slice(0, maxItemsPerGroup)
    .map((s, idx) => ({ ...s, grade: "8% 이상", rank: idx + 1 }));

  const targetStocks = [...upperGroup, ...midGroup, ...lowGroup];

  // 5. 각 종목별 Google 뉴스 RSS 수집 (대표 뉴스 최대 2개)
  targetStocks.forEach((stock) => {
    stock.news = fetchGoogleNewsRss(stock.name, maxNewsPerItem);
    Utilities.sleep(200);
  });

  // 6. '급등주_브리핑' 시트 기존 데이터 초기화 후 새 데이터 기록 (3행부터 기록)
  const lastRow = briefingSheet.getLastRow();
  if (lastRow >= 3) {
    briefingSheet.getRange(3, 1, lastRow - 2, 12).clearContent();
  }

  if (targetStocks.length > 0) {
    const rows = targetStocks.map((s) => [
      timestamp,
      s.grade,
      s.rank,
      s.code,
      s.name,
      s.price,
      s.changeRate / 100,
      "-",
      s.news[0] ? s.news[0].title : "관련 뉴스 없음",
      s.news[0] ? s.news[0].link : "",
      s.news[1] ? s.news[1].title : "",
      s.news[1] ? s.news[1].link : "",
    ]);
    briefingSheet.getRange(3, 1, rows.length, 12).setValues(rows);
  }

  // 7. HTML 이메일 발송 및 '발송_이력' 시트 기록
  let status = "성공";
  let note = `정상 [KRX 기준일:${krxResult.tradeDate || "없음"}, 상태:${krxResult.httpStatus}, 원본:${allStocks.length}건 → 필터후:${filteredStocks.length}건]`;

  try {
    if (!recipientEmail || !recipientEmail.includes("@")) {
      throw new Error(
        "설정_및_요약 탭(B16)의 수신 이메일 주소를 확인해주세요.",
      );
    }
    const htmlBody = buildEmailHtml(
      timestamp,
      krxResult.tradeDate,
      upperGroup,
      midGroup,
      lowGroup,
    );
    const subject = `[장전 브리핑] 전일 급등주 요약 (상한가 ${upperGroup.length} / 15%↑ ${midGroup.length} / 8%↑ ${lowGroup.length})`;

    MailApp.sendEmail({
      to: recipientEmail,
      subject: subject,
      htmlBody: htmlBody,
    });
  } catch (e) {
    status = "실패";
    note = e.message;
  }

  const nextLogRow = logSheet.getLastRow() + 1;
  logSheet
    .getRange(nextLogRow, 1, 1, 8)
    .setValues([
      [
        timestamp,
        recipientEmail,
        upperGroup.length,
        midGroup.length,
        lowGroup.length,
        `=SUM(C${nextLogRow}:E${nextLogRow})`,
        status,
        note,
      ],
    ]);
}

/**
 * 한국거래소(KRX) 정보데이터시스템에서 가장 최근 영업일의 코스피·코스닥 상승 종목 조회
 */
function fetchKrxRecentRiseStocks(minRate) {
  const url = "http://data.krx.co.kr/comm/bldAttendant/getJsonData.cmd";
  const today = new Date();
  let lastHttpStatus = 0;

  // 오늘부터 최대 7일 전까지 역순 탐색하여 시세가 확정된 최근 영업일 데이터 확보
  for (let offset = 0; offset < 7; offset++) {
    const targetDate = new Date(today.getTime() - offset * 24 * 60 * 60 * 1000);
    const dayOfWeek = targetDate.getDay();
    if (dayOfWeek === 0 || dayOfWeek === 6) continue; // 주말 건너뛰기

    const trdDd = Utilities.formatDate(targetDate, "Asia/Seoul", "yyyyMMdd");
    const options = {
      method: "post",
      muteHttpExceptions: true,
      headers: {
        Referer:
          "http://data.krx.co.kr/contents/MDC/MDI/mdiLoader/index.cmd?menuId=MDC0201020101",
        "User-Agent":
          "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
      },
      payload: {
        bld: "dbms/MDC/STAT/standard/MDCSTAT01501",
        locale: "ko_KR",
        mktId: "ALL",
        trdDd: trdDd,
        share: "1",
        money: "1",
        csvxls_isNo: "false",
      },
    };

    try {
      const response = UrlFetchApp.fetch(url, options);
      lastHttpStatus = response.getResponseCode();
      if (lastHttpStatus !== 200) continue;

      const json = JSON.parse(response.getContentText("UTF-8"));
      const list = json.OutBlock_1 || [];

      // 해당 날짜에 실제 거래(등락률) 데이터가 존재하는지 확인
      const hasValidPrice = list.some(
        (item) =>
          item.TDD_CLSPRC &&
          item.TDD_CLSPRC !== "-" &&
          item.FLUC_RT &&
          item.FLUC_RT !== "-",
      );
      if (!hasValidPrice) continue;

      const stocks = [];
      list.forEach((item) => {
        const market = String(item.MKT_NM || "").trim();
        if (market !== "KOSPI" && market !== "KOSDAQ") return; // 코넥스(KONEX) 제외

        const code = String(item.ISU_SRT_CD || "").trim();
        const name = String(item.ISU_ABBRV || "").trim();
        const price =
          Number(String(item.TDD_CLSPRC || "0").replace(/,/g, "")) || 0;
        const changeRate =
          parseFloat(String(item.FLUC_RT || "0").replace(/,/g, "")) || 0;

        if (code && name && price > 0 && changeRate >= minRate) {
          stocks.push({ code, name, price, changeRate });
        }
      });

      return { stocks, tradeDate: trdDd, httpStatus: lastHttpStatus };
    } catch (e) {
      Logger.log(`KRX 조회 오류(${trdDd}): ${e.message}`);
    }
  }

  return { stocks: [], tradeDate: "", httpStatus: lastHttpStatus };
}

/**
 * 단타 부적합 종목(우선주, 스팩, ETF/ETN) 제외 필터
 */
function isValidStock(stock) {
  if (!stock.code || !stock.code.endsWith("0")) return false;
  const excludeKeywords = [
    "스팩",
    "ETF",
    "ETN",
    "KODEX",
    "TIGER",
    "KBSTAR",
    "ACE",
    "SOL",
    "ARIRANG",
    "HANARO",
    "KOSEF",
    "인버스",
    "레버리지",
  ];
  for (let kw of excludeKeywords) {
    if (stock.name.includes(kw)) return false;
  }
  return true;
}

/**
 * Google 뉴스 RSS를 통해 종목 관련 최신 대표 뉴스 수집
 */
function fetchGoogleNewsRss(stockName, maxCount) {
  const query = encodeURIComponent(
    `${stockName} 특징주 OR 급등 OR 상한가 when:2d`,
  );
  const rssUrl = `https://news.google.com/rss/search?q=${query}&hl=ko&gl=KR&ceid=KR:ko`;

  try {
    const xmlText = UrlFetchApp.fetch(rssUrl, {
      muteHttpExceptions: true,
    }).getContentText("UTF-8");
    const document = XmlService.parse(xmlText);
    const items = document
      .getRootElement()
      .getChild("channel")
      .getChildren("item");

    const newsList = [];
    for (let i = 0; i < Math.min(items.length, maxCount); i++) {
      const title = items[i].getChildText("title");
      const link = items[i].getChildText("link");
      newsList.push({ title, link });
    }
    return newsList;
  } catch (e) {
    return [];
  }
}

/**
 * 3단계 색상 구분(빨강/주황/노랑)이 적용된 HTML 이메일 본문 생성
 */
function buildEmailHtml(timestamp, tradeDate, upperGroup, midGroup, lowGroup) {
  const renderSection = (title, bgColor, badgeColor, list) => {
    if (list.length === 0) {
      return `<h3 style="color:${badgeColor}; margin-top:20px;">■ ${title} (0종목)</h3><p style="color:#666;">해당 조건의 종목이 없습니다.</p>`;
    }
    let rowsHtml = list
      .map((s) => {
        const news1 = s.news[0]
          ? `<a href="${s.news[0].link}" target="_blank" style="color:#1a73e8; text-decoration:none;">• ${s.news[0].title}</a>`
          : "-";
        const news2 = s.news[1]
          ? `<br><a href="${s.news[1].link}" target="_blank" style="color:#1a73e8; text-decoration:none;">• ${s.news[1].title}</a>`
          : "";
        return `
        <tr style="border-bottom:1px solid #e0e0e0;">
          <td style="padding:8px; text-align:center; background:${bgColor}; color:${badgeColor}; font-weight:bold;">${s.grade}</td>
          <td style="padding:8px; font-weight:bold;">${s.name}<br><span style="font-size:11px; color:#777;">(${s.code})</span></td>
          <td style="padding:8px; text-align:right;">${s.price.toLocaleString()}원</td>
          <td style="padding:8px; text-align:right; color:#d93025; font-weight:bold;">+${s.changeRate.toFixed(2)}%</td>
          <td style="padding:8px; font-size:13px; line-height:1.5;">${news1}${news2}</td>
        </tr>`;
      })
      .join("");

    return `
      <h3 style="color:${badgeColor}; margin-top:24px; margin-bottom:8px;">■ ${title} (상위 ${list.length}종목)</h3>
      <table style="width:100%; border-collapse:collapse; font-size:14px;">
        <thead>
          <tr style="background:#f1f3f4; border-top:2px solid ${badgeColor}; border-bottom:1px solid #ccc;">
            <th style="padding:8px; width:80px;">등급</th>
            <th style="padding:8px; width:140px; text-align:left;">종목명</th>
            <th style="padding:8px; width:95px; text-align:right;">전일종가</th>
            <th style="padding:8px; width:85px; text-align:right;">등락률</th>
            <th style="padding:8px; text-align:left;">대표 뉴스 (최대 2건)</th>
          </tr>
        </thead>
        <tbody>${rowsHtml}</tbody>
      </table>`;
  };

  return `
    <div style="font-family:'Malgun Gothic', sans-serif; max-width:900px; margin:0 auto; color:#222;">
      <h2 style="border-bottom:2px solid #344479; padding-bottom:10px; color:#344479;">전일 급등주 및 핵심 재료 브리핑</h2>
      <p style="font-size:13px; color:#555;">수집 기준 시각: ${timestamp} (시세 기준일: ${tradeDate || "-"}) | 등급별 최대 7종목 · 대표 뉴스 최대 2건</p>
      ${renderSection("상한가 종목", "#FEE2E2", "#991B1B", upperGroup)}
      ${renderSection("15% 이상 급등 종목", "#FFEDD5", "#C2410C", midGroup)}
      ${renderSection("8% 이상 상승 종목", "#FEF9C3", "#854D0E", lowGroup)}
    </div>`;
}
