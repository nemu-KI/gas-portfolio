// =====================
// ヘルパー関数
// =====================
function showMessage(message) {
  console.log(message);
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch(e) {
    // エディタから直接実行した場合はスキップ
  }
}

// =====================
// CSV読み込み・転記
// =====================
function importCSV() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 「読み込み設定」シートのB3からファイル名を取得
  const configSheet = ss.getSheetByName("読み込み設定");
  if (!configSheet) {
    showMessage("エラー: 「読み込み設定」シートが見つかりません。");
    return;
  }

  let fileName = configSheet.getRange("B3").getValue().toString().trim();
  if (!fileName) {
    showMessage("エラー: 「読み込み設定」シートのB3セルにファイル名を入力してください。");
    return;
  }

  // ユーザーが拡張子(.csv)を未入力の場合は自動で付与
  if (!fileName.toLowerCase().endsWith(".csv")) {
    fileName += ".csv";
  }

  // このスプレッドシートの親フォルダを取得
  const ssId = ss.getId();
  const ssFile = DriveApp.getFileById(ssId);
  const parentFolders = ssFile.getParents();

  if (!parentFolders.hasNext()) {
    showMessage("エラー: スプレッドシートの親フォルダが見つかりません。");
    return;
  }
  const parentFolder = parentFolders.next();

  // 親フォルダ内からファイル名でCSVを検索
  const files = parentFolder.getFilesByName(fileName);
  if (!files.hasNext()) {
    showMessage(`エラー: フォルダ内に「${fileName}」という名前のファイルが見つかりません。`);
    return;
  }
  const csvFile = files.next();

  // CSVをテキストとして読み込む
  let csvDataText = csvFile.getBlob().getDataAsString('UTF-8');
  
  // BOM対策：もしテキストの先頭に見えないBOM（\ufeff）があれば自動で削除する
  if (csvDataText.startsWith('\ufeff')) {
    csvDataText = csvDataText.substring(1);
  }

  // 改行およびカンマで分割して2次元配列に変換
  const csvValues = Utilities.parseCsv(csvDataText);

  if (csvValues.length === 0 || csvValues[0].length === 0) {
    showMessage("確認: CSVファイルが空、またはデータが含まれていません。");
    return;
  }

  // 「データ」シートをクリアして書き込む
  const dataSheet = ss.getSheetByName("データ");
  if (!dataSheet) {
    showMessage("エラー: 「データ」シートが見つかりません。");
    return;
  }

  dataSheet.clearContents();
  dataSheet.getRange(1, 1, csvValues.length, csvValues[0].length).setValues(csvValues);

  showMessage("完了: CSVデータの転記が成功しました！");
}

// =====================
// データ絞り込み
// =====================
function filterCSVData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 「読み込み設定」シートのE3（都道府県）とF3（購入金額）から条件を取得
  const configSheet = ss.getSheetByName("読み込み設定");
  if (!configSheet) {
    showMessage("エラー: 「読み込み設定」シートが見つかりません。");
    return;
  }

  // 都道府県条件の取得
  const prefCondition = configSheet.getRange("E3").getValue().toString().trim();
  // 購入金額条件の取得
  const amountCondition = configSheet.getRange("F3").getValue();

  // 「データ」シートの全データを読み込む
  const dataSheet = ss.getSheetByName("データ");
  if (!dataSheet) {
    showMessage("エラー: 「データ」シートが見つかりません。");
    return;
  }

  const lastRow = dataSheet.getLastRow();
  const lastColumn = dataSheet.getLastColumn();

  // データが1行もない（ヘッダーのみ、または完全空）場合は処理を終了
  if (lastRow <= 1) {
    showMessage("確認: 「データ」シートに絞り込み対象のデータがありません。");
    return;
  }

  // 全データを2次元配列として取得
  const allRows = dataSheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const header = allRows[0];         // 1行目（ヘッダー行）
  const dataRows = allRows.slice(1); // 2行目以降（実際のデータ）


  // ヘッダー（1行目）から文字列で列の位置（インデックス）を動的に検索
  // indexOf() を使って、指定した文字列がヘッダーの何番目にあるかを検索（見つからない場合は -1 が返る）
  const prefColumnIndex = header.indexOf("都道府県");
  const amountColumnIndex = header.indexOf("購入金額");

  // 安全対策: 万が一、指定したヘッダー名が存在しない場合はエラーを出して処理を中断する
  if (prefColumnIndex === -1) {
    showMessage("エラー: 「データ」シートの1行目に「都道府県」という列名が見つかりません。");
    return;
  }
  if (amountColumnIndex === -1) {
    showMessage("エラー: 「データ」シートの1行目に「購入金額」という列名が見つかりません。");
    return;
  }

  // 条件に合う行だけを絞り込む
  const filteredRows = dataRows.filter(row => {
    // 見つけた列インデックスを使って、動的に値を取得
    const rowPref = row[prefColumnIndex] ? row[prefColumnIndex].toString().trim() : "";
    const rowAmount = row[amountColumnIndex] !== "" ? Number(row[amountColumnIndex]) : 0;

    // 都道府県条件の判定
    const isPrefMatch = (prefCondition === "全県" || prefCondition === "" || rowPref === prefCondition);

    // 購入金額条件の判定
    const isAmountMatch = (amountCondition === "" || rowAmount >= Number(amountCondition));

    // 両方の条件を満たす行だけを残す
    return isPrefMatch && isAmountMatch;
  });

  //「ソート結果」シートをクリアして出力する
  const resultSheet = ss.getSheetByName("ソート結果");
  if (!resultSheet) {
    showMessage("エラー: 「ソート結果」シートが見つかりません。");
    return;
  }

  // 既存データをクリア
  resultSheet.clearContents();

  // 出力用データ配列の作成（先頭にヘッダーを戻す）
  const outputValues = [header, ...filteredRows];

  // シートへの書き込み
  resultSheet.getRange(1, 1, outputValues.length, outputValues[0].length).setValues(outputValues);

  showMessage(`完了: 条件に合うデータを絞り込みました。（該当: ${filteredRows.length}件）`);
}

// =====================
// メール生成
// =====================
function createMail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  //スプレッドシートの親フォルダを取得 ➔ 「メール」フォルダを検索
  const ssId = ss.getId();
  const ssFile = DriveApp.getFileById(ssId);
  const parentFolders = ssFile.getParents();

  if (!parentFolders.hasNext()) {
    showMessage("エラー: スプレッドシートの親フォルダが見つかりません。");
    return;
  }
  const parentFolder = parentFolders.next();

  // 「メール」フォルダを探す。なければ自動で新規作成する
  const folderName = "メール";
  const subFolders = parentFolder.getFoldersByName(folderName);
  let mailFolder;

  if (subFolders.hasNext()) {
    mailFolder = subFolders.next();
  } else {
    // フォルダが見つからなかった場合は自動で作成
    mailFolder = parentFolder.createFolder(folderName);
  }

  // 「ソート結果」シートの全データを読み込み ➔ 0件チェック
  const resultSheet = ss.getSheetByName("ソート結果");
  if (!resultSheet) {
    showMessage("エラー: 「ソート結果」シートが見つかりません。先に絞り込みを実行してください。");
    return;
  }

  const lastRow = resultSheet.getLastRow();
  const lastColumn = resultSheet.getLastColumn();

  // データが1行もない（ヘッダーのみ、または完全空）場合はエラー
  if (lastRow <= 1) {
    showMessage("エラー: 「ソート結果」シートにデータがありません。先に絞り込みを実行してください。");
    return;
  }

  // 全データを2次元配列として取得
  const allRows = resultSheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const header = allRows[0];         // 1行目（ヘッダー行）
  const dataRows = allRows.slice(1); // 2行目以降（ソートされた実際のデータ）


  // ヘッダー行から各列のインデックスを取得
  const idColumnIndex = header.indexOf("ユーザーID");
  const nameColumnIndex = header.indexOf("氏名");
  const amountColumnIndex = header.indexOf("購入金額");

  // 必要な列名が見つからない場合は処理を中断
  if (idColumnIndex === -1 || nameColumnIndex === -1 || amountColumnIndex === -1) {
    showMessage("エラー: 「ソート結果」シートの1行目に「ユーザーID」「氏名」「購入金額」のいずれかが見つかりません。");
    return;
  }

  // 2行目以降のデータを1件ずつループ処理
  let outputCount = 0;

  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];

    // 動的に値を取得（空白、または行データが足りない場合への安全対策）
    const userId = row[idColumnIndex] ? row[idColumnIndex].toString().trim() : `UNKNOWN_${i+1}`;
    const name = row[nameColumnIndex] ? row[nameColumnIndex].toString().trim() : "お客様";
    const amount = row[amountColumnIndex] !== "" ? Number(row[amountColumnIndex]).toLocaleString() : "0";

    // メールの文面を作成 
    const mailBody = `件名：ご購入ありがとうございます
${name} 様

この度は${amount}円のご購入ありがとうございます。
またのご利用をお待ちしております。

送信元：サンプルショップ`;

    // ── ファイル名を「ユーザーID_氏名.txt」の形式で生成 ──
    const textFileName = `${userId}_${name}.txt`;

    // ── 「メール」フォルダにテキストファイルとして出力 ──
    // mimeTypeをPLAIN_TEXTに指定して保存します
    mailFolder.createFile(textFileName, mailBody, MimeType.PLAIN_TEXT);
    outputCount++;
  }


  // 完了メッセージを表示（出力件数を含む）
  showMessage(`完了: 「メール」フォルダに ${outputCount} 件のテキストファイルを出力しました！`);
}
