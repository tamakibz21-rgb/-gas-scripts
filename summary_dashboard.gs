// Googleスプレッドシートの売上データを集計し、月次サマリーとチャートを作成

/**
 * メインの集計処理
 * 売上データシートを読み込み、月次サマリーを作成し、グラフを生成
 */
function runDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("売上データ");
  const summarySheet = ss.getSheetByName("月次サマリー");

  if (!sourceSheet || !summarySheet) {
    Logger.log("エラー：「売上データ」または「月次サマリー」シートが見つかりません");
    return;
  }

  // 売上データを取得
  const data = sourceSheet.getDataRange().getValues();

  // 月ごとの集計結果を保存するオブジェクト
  const monthlyData = {};

  // 1行目はヘッダーなのでスキップ
  for (let i = 1; i < data.length; i++) {
    const date = new Date(data[i][0]);
    const amount = data[i][3]; // D列の金額

    // 無効なデータはスキップ
    if (!date || isNaN(date) || !amount) {
      continue;
    }

    // 年月をキーにする（例：2026-01）
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const monthKey = `${year}-${month}`;
    const monthLabel = `${year}年${month}月`;

    // 初期化
    if (!monthlyData[monthKey]) {
      monthlyData[monthKey] = {
        label: monthLabel,
        totalAmount: 0,
        count: 0
      };
    }

    // 集計
    monthlyData[monthKey].totalAmount += amount;
    monthlyData[monthKey].count += 1;
  }

  // 月次サマリーシートをクリア（ヘッダーは残す）
  const lastRow = summarySheet.getLastRow();
  if (lastRow > 1) {
    summarySheet.deleteRows(2, lastRow - 1);
  }

  // ソート済みの月キーを取得
  const sortedMonths = Object.keys(monthlyData).sort();

  // サマリーデータを書き込み
  for (let i = 0; i < sortedMonths.length; i++) {
    const monthKey = sortedMonths[i];
    const data = monthlyData[monthKey];
    const row = i + 2; // 2行目から開始

    summarySheet.getRange(row, 1).setValue(data.label);
    summarySheet.getRange(row, 2).setValue(data.totalAmount);
    summarySheet.getRange(row, 3).setValue(data.count);
  }

  // グラフを作成
  createChart(ss, summarySheet);

  Logger.log("月次サマリーの集計が完了しました");
}

/**
 * 月次推移を表示する棒グラフを作成
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {Sheet} sheet - 月次サマリーシート
 */
function createChart(ss, sheet) {
  // 既存のグラフを削除
  const charts = sheet.getCharts();
  charts.forEach(chart => sheet.removeChart(chart));

  // データの範囲を取得
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log("サマリーデータがないため、グラフを作成できません");
    return;
  }

  // グラフ用のデータ範囲（月とその売上）
  const range = sheet.getRange(1, 1, lastRow, 2);

  // 棒グラフを作成
  const chartBuilder = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(range)
    .setPosition(lastRow + 3, 1, 0, 0)
    .setOption("title", "月次売上推移")
    .setOption("hAxis", {
      title: "月"
    })
    .setOption("vAxis", {
      title: "売上金額"
    })
    .setOption("legend", {
      position: "bottom"
    });

  sheet.insertChart(chartBuilder.build());
  Logger.log("月次売上推移グラフを作成しました");
}

/**
 * 毎朝9時に自動実行するトリガーを設定
 * スクリプトエディタから手動実行して、トリガーを登録
 */
function setDailyTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === "runDashboard") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 新しいトリガーを作成（毎日9時）
  ScriptApp.newTrigger("runDashboard")
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();

  Logger.log("毎朝9時に runDashboard を実行するトリガーを設定しました");
}
