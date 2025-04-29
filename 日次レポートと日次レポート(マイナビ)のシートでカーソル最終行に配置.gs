function onOpen() {
    Logger.log("[INFO] スプレッドシートが開かれました。");

    const sheetNames = ["日次レポート", "日次レポート (マイナビ)"];
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    sheetNames.forEach(sheetName => {
        let sheet = ss.getSheetByName(sheetName);
        if (sheet) {
            moveToLastFilledRow(sheet);
        } else {
            Logger.log(`[ERROR] シートが見つかりません: ${sheetName}`);
        }
    });
}

function moveToLastFilledRow(sheet) {
    Logger.log(`[INFO] スクリプト開始: ${sheet.getName()}`);

    SpreadsheetApp.flush(); // **キャッシュクリア**

    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();

    if (lastRow < 1 || lastColumn < 1) {
        Logger.log(`[WARNING] シート「${sheet.getName()}」にデータがありません。処理を終了します。`);
        return;
    }

    Logger.log(`[INFO] 最終行（推定）: ${lastRow}`);

    // **最終2000行の範囲を取得 (パフォーマンス最適化)**
    const checkRowStart = Math.max(lastRow - 2000, 1);
    let checkRange = sheet.getRange(checkRowStart, 1, lastRow - checkRowStart + 1, lastColumn);
    let values = checkRange.getDisplayValues(); // **表示データのみ取得**
    
    let foundLastValidRow = -1;

    // **最終行の検索**
    for (let i = values.length - 1; i >= 0; i--) {
        let hasValue = values[i].some(cell => cell !== "" && cell !== null);
        if (hasValue) {
            foundLastValidRow = checkRowStart + i;
            Logger.log(`[INFO] 最終行（値あり）を発見: ${foundLastValidRow} 行目`);
            break;
        }
    }

    if (foundLastValidRow === -1) {
        Logger.log(`[ERROR] シート「${sheet.getName()}」で最終行を特定できませんでした。処理を終了します。`);
        return;
    }

    let targetCell = sheet.getRange(foundLastValidRow, 1);
    sheet.setActiveSelection(targetCell);
    Logger.log(`[INFO] カーソルを ${foundLastValidRow} 行目のA列に移動しました`);

    Logger.log(`[INFO] スクリプト完了: ${sheet.getName()}`);
}
