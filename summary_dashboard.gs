// Googleスプレッドシートの売上データを集計し、月次サマリーとチャートを作成

/**
 * シンプルな接続テスト
 */
function simpleTest() {
  Logger.log("シンプルテスト実行");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("OK: " + ss.getName());
}

/**
 * 環境確認テスト（デバッグ用）
 * シートの存在確認とスプレッドシート情報を取得
 */
function testEnvironment() {
  try {
    Logger.log("=== 環境確認テスト開始 ===");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log("スプレッドシート名: " + ss.getName());
    Logger.log("スプレッドシートURL: " + ss.getUrl());

    const sheets = ss.getSheets();
    Logger.log("シート総数: " + sheets.length);

    sheets.forEach((sheet, index) => {
      Logger.log((index + 1) + ". シート名: " + sheet.getName() + " | 最終行: " + sheet.getLastRow());
    });

    Logger.log("=== 環境確認テスト完了 ===");
  } catch (error) {
    Logger.log("環境確認エラー: " + error.message);
  }
}

/**
 * メインの集計処理
 * 売上データシートを読み込み、月次サマリーを作成し、グラフを生成
 */
function runDashboard() {
  try {
    Logger.log("=== runDashboard 開始 ===");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log("スプレッドシート: " + ss.getName());

    // シート名を確認
    const allSheets = ss.getSheets();
    Logger.log("利用可能なシート: " + allSheets.map(s => s.getName()).join(", "));

    const sourceSheet = ss.getSheetByName("売上データー");
    const summarySheet = ss.getSheetByName("月次サマリー");

    if (!sourceSheet) {
      throw new Error("「売上データ」シートが見つかりません");
    }
    if (!summarySheet) {
      throw new Error("「月次サマリー」シートが見つかりません");
    }

    Logger.log("シート確認完了");

    // 売上データを取得
    const lastRow = sourceSheet.getLastRow();
    Logger.log("売上データの行数: " + lastRow);

    if (lastRow < 2) {
      Logger.log("データがありません（ヘッダーのみ）");
      return;
    }

    const data = sourceSheet.getRange(1, 1, lastRow, 4).getValues();
    Logger.log("データ取得完了: " + data.length + "行");

    // 月ごとの集計結果を保存するオブジェクト
    const monthlyData = {};
    let processedCount = 0;

    // 1行目はヘッダーなのでスキップ
    for (let i = 1; i < data.length; i++) {
      const dateValue = data[i][0];
      const amount = data[i][3]; // D列の金額

      // データが空の場合はスキップ
      if (!dateValue || !amount) {
        continue;
      }

      // 日付を解析（Google Sheetsの日付シリアル値対応）
      let date;
      if (dateValue instanceof Date) {
        date = new Date(dateValue);
      } else if (typeof dateValue === "number") {
        // シリアル値の場合（Excelからの変換用）
        date = new Date((dateValue - 25569) * 86400 * 1000);
      } else if (typeof dateValue === "string") {
        date = new Date(dateValue);
      } else {
        continue;
      }

      // 無効な日付はスキップ
      if (isNaN(date.getTime())) {
        Logger.log("警告: 無効な日付スキップ (行" + (i + 1) + "): " + dateValue);
        continue;
      }

      // 金額を数値に変換
      const numAmount = Number(amount);
      if (isNaN(numAmount)) {
        Logger.log("警告: 無効な金額スキップ (行" + (i + 1) + "): " + amount);
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
      monthlyData[monthKey].totalAmount += numAmount;
      monthlyData[monthKey].count += 1;
      processedCount += 1;
    }

    Logger.log("処理完了: " + processedCount + "件のデータを集計");
    Logger.log("集計月数: " + Object.keys(monthlyData).length + "ヶ月");

    // 月次サマリーシートをクリア（ヘッダーは残す）
    Logger.log("サマリーシートのクリア処理開始");
    const summaryLastRow = summarySheet.getLastRow();
    if (summaryLastRow > 1) {
      try {
        summarySheet.getRange(2, 1, summaryLastRow - 1, 3).clearContent();
        Logger.log("データ行をクリアしました");
      } catch (clearError) {
        Logger.log("クリア処理エラー: " + clearError.message);
      }
    }

    // ヘッダーを確認・設定
    const headerRow = summarySheet.getRange(1, 1, 1, 3).getValues();
    if (!headerRow[0][0] || headerRow[0][0] === "") {
      summarySheet.getRange(1, 1).setValue("月");
      summarySheet.getRange(1, 2).setValue("合計売上");
      summarySheet.getRange(1, 3).setValue("件数");
      Logger.log("ヘッダーを作成しました");
    }

    // ソート済みの月キーを取得
    const sortedMonths = Object.keys(monthlyData).sort();
    Logger.log("ソート済み月数: " + sortedMonths.join(", "));

    // サマリーデータを書き込み
    for (let i = 0; i < sortedMonths.length; i++) {
      const monthKey = sortedMonths[i];
      const monthData = monthlyData[monthKey];
      const row = i + 2; // 2行目から開始

      summarySheet.getRange(row, 1).setValue(monthData.label);
      summarySheet.getRange(row, 2).setValue(monthData.totalAmount);
      summarySheet.getRange(row, 3).setValue(monthData.count);
    }

    Logger.log("サマリーシートへの書き込み完了");

    // グラフを作成
    try {
      createChart(ss, summarySheet);
    } catch (chartError) {
      Logger.log("グラフ作成中にエラー発生（続行）: " + chartError.message);
    }

    Logger.log("=== runDashboard 完了 ===");
  } catch (error) {
    Logger.log("エラーが発生しました: " + error.message);
    Logger.log("スタックトレース: " + error.stack);
  }
}

/**
 * 月次推移を表示する棒グラフを作成
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {Sheet} sheet - 月次サマリーシート
 */
function createChart(ss, sheet) {
  try {
    Logger.log("グラフ作成開始");

    // 既存のグラフを削除
    const charts = sheet.getCharts();
    Logger.log("既存グラフ数: " + charts.length);
    charts.forEach(chart => sheet.removeChart(chart));

    // データの範囲を取得
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log("サマリーデータがないため、グラフを作成できません");
      return;
    }

    Logger.log("グラフデータ行数: " + lastRow);

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
  } catch (error) {
    Logger.log("グラフ作成エラー: " + error.message);
  }
}

/**
 * 毎朝9時に自動実行するトリガーを設定
 * スクリプトエディタから手動実行して、トリガーを登録
 */
function setDailyTrigger() {
  try {
    Logger.log("トリガー設定開始");

    // 既存のトリガーを削除
    const triggers = ScriptApp.getProjectTriggers();
    Logger.log("既存トリガー数: " + triggers.length);

    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === "runDashboard") {
        ScriptApp.deleteTrigger(trigger);
        Logger.log("既存トリガーを削除しました");
      }
    });

    // 新しいトリガーを作成（毎日9時）
    ScriptApp.newTrigger("runDashboard")
      .timeBased()
      .atHour(9)
      .everyDays(1)
      .create();

    Logger.log("毎朝9時に runDashboard を実行するトリガーを設定しました");
  } catch (error) {
    Logger.log("トリガー設定エラー: " + error.message);
  }
}
