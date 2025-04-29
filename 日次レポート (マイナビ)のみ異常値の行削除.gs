function deleteRowsForMynaviReport() {
    Logger.log("[INFO] スクリプト開始: 日次レポート (マイナビ)");

    const sheetName = "日次レポート (マイナビ)"; // 対象シート
    const numHeaderRow = 2; // ヘッダー行を除外
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

    if (!sheet) {
        Logger.log("[ERROR] シートが見つかりません: " + sheetName);
        return;
    }
    Logger.log("[INFO] シート取得成功: " + sheet.getName());

    const lastRow = sheet.getLastRow();
    const lastColumn = 15; // A列～O列まで取得

    if (lastRow <= numHeaderRow) {
        Logger.log("[WARNING] データが不足しているためスキップ");
        return;
    }
    Logger.log(`[INFO] 最終行: ${lastRow}`);

    // **A列～O列のデータを取得**
    const range = sheet.getRange(numHeaderRow + 1, 1, lastRow - numHeaderRow, lastColumn).getValues();
    Logger.log(`[INFO] データ取得成功, 取得行数: ${range.length}`);

    // **基準日 (前日) の取得**
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const baseDate = new Date(today);
    baseDate.setDate(today.getDate() - 1);
    const baseDateStr = Utilities.formatDate(baseDate, "JST", "yyyy/MM/dd");
    Logger.log(`[INFO] 基準日: ${baseDateStr}`);

    let rowsToDelete = [];
    let lastValidRow = -1; // 基準行を格納する変数

    // **基準行を見つける**
    for (let i = range.length - 1; i >= 0; i--) {
        let rowValue = range[i][0]; // A列の値
        let parsedDate = parseDate(rowValue);

        if (parsedDate) {
            parsedDate.setHours(0, 0, 0, 0);
            const parsedDateStr = Utilities.formatDate(parsedDate, "JST", "yyyy/MM/dd");

            if (parsedDateStr === baseDateStr) {
                lastValidRow = numHeaderRow + i + 1; // **1-based index**
                Logger.log(`[INFO] 基準行を発見: 行 ${lastValidRow}, 日付=${parsedDateStr}`);
                break;
            }
        }
    }

    if (lastValidRow === -1) {
        Logger.log("[WARNING] 基準行が見つかりませんでした。処理を終了します。");
        return;
    }

    // **基準行の1つ下の行から本当の最終行までスキャン**
    for (let i = range.length - 1; i >= 0; i--) {
        let currentRowNum = numHeaderRow + i + 1;
        if (currentRowNum <= lastValidRow) {
            break; // **基準行に到達したら処理を停止**
        }

        let rowValue = range[i][0]; // A列の値
        let parsedDate = parseDate(rowValue);
        let shouldDelete = false;

        if (parsedDate) {
            parsedDate.setHours(0, 0, 0, 0);
            if (parsedDate.getTime() < baseDate.getTime()) {
                Logger.log(`[INFO] 削除対象 (基準日より前): 行 ${currentRowNum}, ${Utilities.formatDate(parsedDate, "JST", "yyyy/MM/dd")}`);
                shouldDelete = true;
            }
        }

        // **B～O列に 1 または 0 がある場合も削除**
        if (!shouldDelete) {
            for (let j = 1; j < range[i].length; j++) {
                if (range[i][j] == 1 || range[i][j] == 0) {
                    Logger.log(`[INFO] 削除対象 (B～O列に1または0あり): 行 ${currentRowNum}`);
                    shouldDelete = true;
                    break;
                }
            }
        }

        if (shouldDelete) {
            rowsToDelete.push(currentRowNum);
        }
    }

    // **削除処理 (逆順で削除)**
    if (rowsToDelete.length > 0) {
        Logger.log(`[INFO] 削除対象行: ${JSON.stringify(rowsToDelete)}`);
        rowsToDelete.sort((a, b) => b - a);
        rowsToDelete.forEach(row => {
            sheet.deleteRow(row);
            Logger.log(`[INFO] ${row} 行目を削除しました`);
        });
        SpreadsheetApp.flush();
    } else {
        Logger.log("[INFO] 削除対象の行はありません");
    }

    Logger.log("[INFO] スクリプト完了: 日次レポート (マイナビ)");
}

/**
 * **A列の日付を解析**
 * @param {any} value - セルの値
 * @returns {Date|null} - 変換された Date オブジェクト (失敗時は null)
 */
function parseDate(value) {
    if (!value) return null;

    if (value instanceof Date) {
        value.setHours(0, 0, 0, 0);
        return isNaN(value.getTime()) ? null : value;
    }

    if (typeof value === "number") {
        // **Excel シリアル値を日付に変換**
        const date = new Date(1899, 11, 30);
        date.setDate(date.getDate() + value);
        date.setHours(0, 0, 0, 0);
        return date;
    }

    if (typeof value === "string") {
        const cleanedValue = value.trim();
        const parsedDate = new Date(cleanedValue);
        return isNaN(parsedDate.getTime()) ? null : parsedDate;
    }

    return null;
}
