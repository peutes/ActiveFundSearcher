/**
 * みんかぶのシャープレシオランキング（5年版）からデータを取得し、
 * Googleスプレッドシートに書き込むGoogle Apps Scriptです。
 * ランキングページはPhantomJS Cloudを、各ファンドの/returnページは直接URLフェッチを使用します。
 *
 * 取得するデータ:
 * - 順位
 * - 会社名
 * - ファンド名
 * - ファンドのURL
 * - 各ファンドの /return ページからの「5年シャープレシオ」
 * - 各ファンドの /return ページからの「5年リターン」
 *
 * スプレッドシートはアクティブなスプレッドシートを使用し、
 * シート名は「YYYY/MM」（例: 2025/07）となります。
 *
 * ページを繰り返し取得し、最大100位までのデータを収集します。
 */
function getMinkabuSharpeRatioRanking() {
  const baseUrl = "https://itf.minkabu.jp/ranking/sharpe_ratio?term=60"; // 5年版シャープレシオランキングの基本URL
  const ss = SpreadsheetApp.getActiveSpreadsheet(); // アクティブなスプレッドシートを取得
  const MAX_RANK_TO_FETCH = 100; // 取得したい最大順位を100に設定

  // PhantomJS Cloud の設定を再度有効化
  const PHANTOMJS_CLOUD_API_KEY = "ak-7twx8-xw8e8-yaydg-havtz-5r3yk"; // APIキーを更新しました
  const PHANTOMJS_CLOUD_ENDPOINT = `https://phantomjscloud.com/api/browser/v2/${PHANTOMJS_CLOUD_API_KEY}/`;

  let allRankingData = [];
  // ヘッダー行を更新: 順序を「5年シャープレシオ」, 「5年リターン」に変更
  allRankingData.push(['順位', '会社名', 'ファンド名', 'ファンドURL', '5年シャープレシオ', '5年リターン']);

  let page = 1;
  let fetchedCount = 0;

  try {
    while (fetchedCount < MAX_RANK_TO_FETCH) {
      const targetUrl = `${baseUrl}&page=${page}`; // ページパラメータを追加したターゲットURL
      Logger.log(`PhantomJS Cloud を使用して URL をフェッチ中 (ランキングページ): ${targetUrl}`);

      // PhantomJS Cloud へのリクエストボディを構築
      const payload = {
        url: targetUrl,
        renderType: "html",
        renderSettings: {
          resourceTimeout: 10000 // 10秒でタイムアウト
        }
      };

      const options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true // エラーが発生しても例外をスローしない
      };

      // PhantomJS Cloud を介してウェブページをフェッチ
      const response = UrlFetchApp.fetch(PHANTOMJS_CLOUD_ENDPOINT, options);
      const responseCode = response.getResponseCode();
      const htmlContent = response.getContentText("UTF-8");

      if (responseCode !== 200) {
        Logger.log(`ランキングページからエラーレスポンス (${responseCode}): ${htmlContent}`);
        throw new Error(`ランキングページの取得に失敗しました。ステータスコード: ${responseCode}`);
      }

      // HTMLからランキングデータを解析
      const currentPageData = parseMinkabuHtml(htmlContent);

      if (currentPageData.length === 0) {
        // 現在のページにデータがない場合、これ以上ページがないと判断してループを終了
        Logger.log(`ページ ${page} で新しいデータが見つかりませんでした。取得を終了します。`);
        break;
      }

      // 取得したデータを全体データに追加（ヘッダーは最初に追加済みなので、ここではデータ部分のみ）
      // 5年シャープレシオと5年リターンのプレースホルダーを追加
      const dataWithPlaceholders = currentPageData.map(row => [...row, '', '']);
      allRankingData = allRankingData.concat(dataWithPlaceholders);
      fetchedCount += currentPageData.length;
      page++;
    }

    // 取得したデータがMAX_RANK_TO_FETCHを超える場合、指定された順位までで切り詰める
    if (allRankingData.length > MAX_RANK_TO_FETCH + 1) { // +1はヘッダー行のため
      allRankingData = allRankingData.slice(0, MAX_RANK_TO_FETCH + 1);
      Logger.log(`${MAX_RANK_TO_FETCH}件のデータに切り詰めました。`);
    }

    // 各ファンドの /return ページからデータを取得
    // ヘッダー行を除外してループ
    for (let i = 1; i < allRankingData.length; i++) {
      const fundUrl = allRankingData[i][3]; // ファンドURLはインデックス3
      if (fundUrl) {
        const returnPageUrl = fundUrl + '/return'; // /fund/ID を /fund/ID/return に変換
        Logger.log(`URLを直接フェッチ中 (リターンページ): ${returnPageUrl}`);

        try {
          const returnOptions = {
            muteHttpExceptions: true
          };
          const returnResponse = UrlFetchApp.fetch(returnPageUrl, returnOptions);
          const returnResponseCode = returnResponse.getResponseCode();
          const returnHtmlContent = returnResponse.getContentText("UTF-8");

          if (returnResponseCode !== 200) {
            Logger.log(`リターンページからエラーレスポンス (${returnResponseCode}): ${returnHtmlContent}`);
            allRankingData[i][4] = `エラー: ページ取得失敗 (${returnResponseCode})`; // 5年シャープレシオ
            allRankingData[i][5] = `エラー: ページ取得失敗 (${returnResponseCode})`; // 5年リターン
          } else {
            const extractedReturnAndSharpeData = parseReturnPageData(returnHtmlContent);
            allRankingData[i][4] = extractedReturnAndSharpeData.fiveYearSharpeRatio; // 5年シャープレシオ (インデックス4)
            allRankingData[i][5] = extractedReturnAndSharpeData.fiveYearReturn; // 5年リターン (インデックス5)
            Logger.log(`リターンページデータ取得完了 for ${fundUrl}: 5年シャープレシオ: ${extractedReturnAndSharpeData.fiveYearSharpeRatio}, 5年リターン: ${extractedReturnAndSharpeData.fiveYearReturn}`);
          }
        } catch (e) {
          Logger.log(`リターンページ (${returnPageUrl}) の取得または解析中にエラー: ${e.message}`);
          allRankingData[i][4] = `エラー: ${e.message}`;
          allRankingData[i][5] = `エラー: ${e.message}`;
        }
      }
    }

    // スプレッドシートにデータを書き込む
    writeToSpreadsheet(ss, allRankingData);

    Logger.log("シャープレシオランキングの取得と書き込みが完了しました。");

  } catch (e) {
    Logger.log("エラーが発生しました: " + e.message);
    // エラーメッセージをユーザーに表示するカスタムメッセージボックスは削除されました
    // SpreadsheetApp.getUi().alert('エラー', 'データの取得または書き込み中にエラーが発生しました。\n詳細はログを確認してください。\n' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * HTMLコンテンツからシャープレシオランキングデータを解析します。
 * この関数は、ヘッダー行を含まないデータ行のみを返します。
 * @param {string} html HTMLコンテンツの文字列
 * @returns {Array<string>} 順位、会社名、ファンド名、ファンドURLを含む配列
 */
function parseMinkabuHtml(html) {
  const data = [];

  // まず、テーブルのデータ部分（通常は<tbody>内）を抽出することを試みます。
  // これにより、ヘッダー行やフッター行など、不要な<tr>を除外できます。
  const tbodyMatch = /<tbody[^>]*>(.*?)<\/tbody>/gs.exec(html);
  let contentToParse = html; // デフォルトはHTML全体
  if (tbodyMatch && tbodyMatch[1]) {
    contentToParse = tbodyMatch[1]; // <tbody>の内容を解析対象とする
    Logger.log("<tbody> タグを検出しました。その内容を解析します。");
  } else {
    Logger.log("<tbody> タグが見つかりませんでした。HTML全体から<tr>を検索します。");
  }

  // ランキングの各行を抽出するための正規表現
  const rowRegex = /<tr[^>]*>(.*?)<\/tr>/gs;
  let rowMatch;

  while ((rowMatch = rowRegex.exec(contentToParse)) !== null) {
    const rowHtml = rowMatch[1]; // 各行のHTMLコンテンツ

    let rank = '';
    let companyName = '';
    let fundName = '';
    let fundUrl = '';

    // 各<td>要素を個別に抽出
    const tdRegex = /<td[^>]*>(.*?)<\/td>/gs;
    const tds = [];
    let tdMatch;
    while ((tdMatch = tdRegex.exec(rowHtml)) !== null) {
      tds.push(tdMatch[1]); // <td>タグ内のコンテンツ
    }

    if (tds.length === 0) {
      Logger.log(`No <td> tags found in row: ${rowHtml}`);
      continue; // <td>タグが見つからない行はスキップ
    }

    // --- 順位の抽出 (最初の<td>内の最初の数字) ---
    if (tds[0]) {
      const rankSpanMatch = /<span class="text-\\[1\.09375rem\\] leading-1\.5\\">(\d+)<\/span>/s.exec(tds[0]);
      if (rankSpanMatch && rankSpanMatch[1]) {
        rank = rankSpanMatch[1].trim();
      } else {
        // スパンタグが見つからない場合、td内のテキスト全体から最初の数字を試す
        const td0Text = tds[0].replace(/<[^>]*>/g, '').trim();
        const rankNumberMatch = /(\d+)/.exec(td0Text);
        if (rankNumberMatch && rankNumberMatch[1]) {
          rank = rankNumberMatch[1].trim();
        }
      }
    }

    // --- 会社名の抽出 (最初の<td>内の<div class="text-itf-xs text-itf-gray-sub-header">) ---
    if (tds[0]) {
      const companyNameMatch = /<div class="text-itf-xs text-itf-gray-sub-header">(.*?)<\/div>/s.exec(tds[0]);
      if (companyNameMatch && companyNameMatch[1]) {
        companyName = companyNameMatch[1].trim();
      }
    }

    // --- ファンド名とURLの抽出 (最初の<td>内の<a>) ---
    if (tds[0]) {
      const fundLinkMatch = /<a data-turbo="false" class="text-itf-sm text-itf-blue-link font-bold hover:underline hover:text-itf-blue-light" href="([^"]+)">([^<]+)<\/a>/s.exec(tds[0]);
      if (fundLinkMatch) {
        fundUrl = "https://itf.minkabu.jp" + fundLinkMatch[1].trim();
        fundName = fundLinkMatch[2].trim();
      }
    }

    // 順位、会社名、ファンド名、ファンドURLが取得できた場合のみ追加
    if (rank && companyName && fundName && fundUrl) {
      data.push([rank, companyName, fundName, fundUrl]);
    } else {
      // 解析できなかった行はログに出力してデバッグに役立てる
      Logger.log(`Skipping incomplete row (Rank: "${rank}", Company: "${companyName}", Fund: "${fundName}", URL: "${fundUrl}")`);
      Logger.log(`Original row HTML for skipped row: ${rowHtml}`);
    }
  }

  return data; // ヘッダー行を含まないデータのみを返す
}

/**
 * /return ページから「5年リターン」と「5年シャープレシオ」のデータを解析します。
 * @param {string} html HTMLコンテンツの文字列
 * @returns {Object} 5年リターンと5年シャープレシオを含むオブジェクト、または「N/A」
 */
function parseReturnPageData(html) {
  let fiveYearReturn = "N/A";
  let fiveYearSharpeRatio = "N/A";

  // 「リターンとリスク」のセクション全体をIDで特定
  const returnAndRiskSectionMatch = /<div id="return_and_risk"[^>]*>(.*?)<\/div>/s.exec(html);

  if (!returnAndRiskSectionMatch || !returnAndRiskSectionMatch[1]) {
    Logger.log("ID 'return_and_risk' のセクションが見つかりませんでした。");
    return { fiveYearReturn: "N/A (セクションなし)", fiveYearSharpeRatio: "N/A (セクションなし)" };
  }

  const sectionContent = returnAndRiskSectionMatch[1];

  // セクション内からテーブルを特定
  const tableMatch = /<table[^>]*>(.*?)<\/table>/s.exec(sectionContent);

  if (!tableMatch || !tableMatch[1]) {
    Logger.log("「リターンとリスク」セクション内にテーブルが見つかりませんでした。");
    return { fiveYearReturn: "N/A (テーブルなし)", fiveYearSharpeRatio: "N/A (テーブルなし)" };
  }

  const tableHtml = tableMatch[1];

  // ヘッダー行から期間のインデックスをマッピング
  const headerRowMatch = /<tr[^>]*>(.*?)<\/tr>/s.exec(tableHtml);
  let fiveYearColumnIndex = -1;

  if (headerRowMatch && headerRowMatch[1]) {
    const headerCells = headerRowMatch[1].match(/<th[^>]*>(.*?)<\/th>/gs);
    if (headerCells) {
      // ヘッダーの最初の<th>は「期間」なので、tdのインデックスと合わせるために-1する
      for (let i = 0; i < headerCells.length; i++) {
        const headerText = headerCells[i].replace(/<[^>]*>/g, '').trim();
        if (headerText === '5年') {
          fiveYearColumnIndex = i -1; // <th>期間</th> を除外したtdのインデックスに調整
          break;
        }
      }
    }
  }

  if (fiveYearColumnIndex === -1) {
    Logger.log("「5年」の期間ヘッダーが見つかりませんでした。");
    return { fiveYearReturn: "N/A (5年ヘッダーなし)", fiveYearSharpeRatio: "N/A (5年ヘッダーなし)" };
  }

  // 「リターン」の行を見つける
  const returnRowRegex = /<tr[^>]*>\s*<th[^>]*>リターン<\/th>(.*?)<\/tr>/s;
  const returnRowMatch = returnRowRegex.exec(tableHtml);

  if (returnRowMatch && returnRowMatch[1]) {
    const returnRowContent = returnRowMatch[1];
    const tdRegex = /<td[^>]*>(.*?)<\/td>/gs;
    const tds = [];
    let tdMatch;
    while ((tdMatch = tdRegex.exec(returnRowContent)) !== null) {
      tds.push(tdMatch[1]);
    }

    if (tds.length > fiveYearColumnIndex && tds[fiveYearColumnIndex]) {
      const rawValue = tds[fiveYearColumnIndex];
      // HTMLタグと<a>タグをすべて除去し、連続する空白を1つのスペースに変換後、前後の空白をトリム
      // 順位情報 (例: (1 位)) も除去
      const cleanedValue = rawValue.replace(/<[^>]*>/g, '').replace(/\(\s*\d+\s*位\)/g, '').replace(/\s+/g, ' ').trim();
      fiveYearReturn = cleanedValue;
    }
  } else {
    Logger.log("「リターン」の行が見つかりませんでした。");
  }

  // 「シャープレシオ」の行を見つける
  const sharpeRatioRowRegex = /<tr[^>]*>\s*<th[^>]*>シャープレシオ<\/th>(.*?)<\/tr>/s;
  const sharpeRatioRowMatch = sharpeRatioRowRegex.exec(tableHtml);

  if (sharpeRatioRowMatch && sharpeRatioRowMatch[1]) {
    const sharpeRatioRowContent = sharpeRatioRowMatch[1];
    const tdRegex = /<td[^>]*>(.*?)<\/td>/gs;
    const tds = [];
    let tdMatch;
    while ((tdMatch = tdRegex.exec(sharpeRatioRowContent)) !== null) {
      tds.push(tdMatch[1]);
    }

    if (tds.length > fiveYearColumnIndex && tds[fiveYearColumnIndex]) {
      const rawValue = tds[fiveYearColumnIndex];
      // HTMLタグと<a>タグをすべて除去し、連続する空白を1つのスペースに変換後、前後の空白をトリム
      // 順位情報 (例: (1 位)) も除去
      const cleanedValue = rawValue.replace(/<[^>]*>/g, '').replace(/\(\s*\d+\s*位\)/g, '').replace(/\s+/g, ' ').trim();
      fiveYearSharpeRatio = cleanedValue;
    }
  } else {
    Logger.log("「シャープレシオ」の行が見つかりませんでした。");
  }

  return { fiveYearReturn: fiveYearReturn, fiveYearSharpeRatio: fiveYearSharpeRatio };
}


/**
 * 取得したデータをスプレッドシートに書き込みます。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss 書き込み先のスプレッドシートオブジェクト
 * @param {Array<Array<string>>} data 書き込むデータ（ヘッダー行を含む）
 */
function writeToSpreadsheet(ss, data) {
  // シート名を「YYYY/MM」形式で生成
  const today = new Date();
  const year = today.getFullYear();
  const month = (today.getMonth() + 1).toString().padStart(2, '0');
  const sheetName = `${year}/${month}`; // シート名をYYYY/MM形式に変更

  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    // シートが存在しない場合は新規作成
    sheet = ss.insertSheet(sheetName);
    Logger.log(`新しいシート '${sheetName}' を作成しました。`);
  } else {
    // シートが存在する場合は既存の内容をクリア
    sheet.clearContents();
    Logger.log(`既存のシート '${sheetName}' の内容をクリアしました。`);
  }

  // データをシートに書き込む
  if (data.length > 0) {
    const range = sheet.getRange(1, 1, data.length, data[0].length);
    range.setValues(data);
    Logger.log(`${data.length - 1}件のランキングデータをシート '${sheetName}' に書き込みました。`);
  } else {
    Logger.log("書き込むデータがありません。");
  }

  // 列幅を自動調整
  sheet.autoResizeColumns(1, data[0].length);
}
