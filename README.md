【日次レポート (マイナビ)のみ異常値の行削除の処理概要】
・これは、「日次レポート (マイナビ)」シートの中から特定の条件に合致する行（基準日より前のデータや異常値を含む行）を削除する Google Apps Script（GAS）のコードです。以下のような処理を行います。

・(1)スクリプトの概要
・Google スプレッドシートの「日次レポート (マイナビ)」シートを開き、データを取得する。
・「基準日（前日）」のデータを探し、その日より前のデータを削除する。
・B列～O列の範囲に「1」または「0」がある場合も削除する。
・不要なデータを削除した後、シートを更新する。

・(2)コードの動作詳細
・①スクリプトの開始
・Logger.log("[INFO] スクリプト開始: 日次レポート (マイナビ)"); → スクリプトが開始されたことをログに記録する。
・const sheetName = "日次レポート (マイナビ)"; → 操作対象のシート名を「日次レポート (マイナビ)」に設定する。
・const numHeaderRow = 2; → ヘッダー行（2行）を除外する。

・②シートの取得と存在確認
・const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName); → 「日次レポート (マイナビ)」シートを取得。
・if (!sheet) { Logger.log("[ERROR] シートが見つかりません: " + sheetName); return; } → シートが見つからない場合、エラーメッセージを出力して処理を終了する。

・③最終行とデータ範囲の取得
・const lastRow = sheet.getLastRow(); → シート内の最後のデータがある行番号を取得。
・const lastColumn = 15; → A列からO列までの範囲（15列）を対象とする。
・if (lastRow <= numHeaderRow) { Logger.log("[WARNING] データが不足しているためスキップ"); return; } → データ行が2行以下の場合、スクリプトをスキップする。
・const range = sheet.getRange(numHeaderRow + 1, 1, lastRow - numHeaderRow, lastColumn).getValues(); → データ範囲（A列～O列）を取得。
・Logger.log("[INFO] データ取得成功, 取得行数: " + range.length); → 取得したデータの行数をログに出力する。

・④基準日の取得
・const today = new Date(); → 現在の日付を取得。
・today.setHours(0, 0, 0, 0); → 時間を 0:00 にリセットする。
・const baseDate = new Date(today); baseDate.setDate(today.getDate() - 1); → 基準日（前日） を取得。
・const baseDateStr = Utilities.formatDate(baseDate, "JST", "yyyy/MM/dd"); → 日付を「yyyy/MM/dd」形式に変換してログに記録。

・⑤基準行（前日データの行）を見つける
・for (let i = range.length - 1; i >= 0; i--) { → 最後の行から逆方向にスキャンする。
・let rowValue = range[i][0]; → A列（1列目）の値を取得。
・let parsedDate = parseDate(rowValue); → A列の値を日付形式に変換。
・if (parsedDate) { → 有効な日付であれば処理を継続。
・if (parsedDateStr === baseDateStr) { → A列の日付が基準日と一致する場合 に基準行を特定。
・Logger.log("[INFO] 基準行を発見: 行 " + lastValidRow + ", 日付=" + parsedDateStr); → 基準行をログに記録。

・⑥削除対象の行を特定
・for (let i = range.length - 1; i >= 0; i--) { → 最後の行から逆方向にスキャン。
・if (currentRowNum <= lastValidRow) { break; } → 基準行に到達したら処理を停止。
・let rowValue = range[i][0]; let parsedDate = parseDate(rowValue); → A列の日付を取得して変換。
・if (parsedDate.getTime() < baseDate.getTime()) { → 基準日より前の日付のデータを削除対象とする。
・for (let j = 1; j < range[i].length; j++) { if (range[i][j] == 1 || range[i][j] == 0) { → B列～O列に「1」または「0」がある場合も削除対象にする。

・⑦行の削除
・rowsToDelete.sort((a, b) => b - a); → 削除を行う際に「大きい行番号から削除」する（小さい行から削除すると、行番号がずれるため）。
・rowsToDelete.forEach(row => { sheet.deleteRow(row); }); → 実際に削除を実行。
・SpreadsheetApp.flush(); → スプレッドシートの変更を確定。

・(3)サブ関数「parseDate」
・parseDate(value) → A列の日付データを適切に解析する関数。
・値の形式ごとに処理を分岐
・Date 型ならそのまま処理。
・数値型（シリアル値）なら Excel 方式の日付に変換。
・文字列型なら日付に変換。

・(4)このスクリプトのまとめ
・(1)スプレッドシートの「日次レポート (マイナビ)」シートからデータを取得。
・(2)A列の日付を解析し、前日（基準日）のデータを探す。
・(3)基準日より前の日付の行を削除対象にする。
・(4)B～O列に「1」「0」のデータが含まれる行も削除対象にする。
・(5)削除リストを作成し、大きい行番号から順に削除する。
・(6)スプレッドシートの変更を確定し、処理を完了する。

・このスクリプトを実行することで、「日次レポート (マイナビ)」シートの中から 基準日より前の古いデータや不要なデータを自動的に削除 することができます。

【日次レポートと日次レポート(マイナビ)_実行時間設定の処理概要】

【1】全体概要
このスクリプトは、Google スプレッドシートに関連する Apps Script であり、異常データ削除関数を1日複数回自動実行するための時間トリガーを、毎日自動で作成・更新する目的で構成されています。
初期設定関数 setupDailyTriggerForTimeTriggers() を一度実行することで、以降は完全自動で対象処理が動作し続けます。

本スクリプトは、過去に発生した「トリガーが単発（1日限り）で終わってしまい、翌日以降は何も実行されない」問題の再発を防ぐために構築されたものです。
当初は .at(triggerTime) を用いてトリガーを設定していましたが、この記述方法ではトリガーが1回限りしか動作せず、日次で自動的に再作成されることはありませんでした。
そのため、翌日以降に処理が実行されない状態に気づかず、異常データ削除が漏れるという問題が発生しました。
この問題には、GASのトリガーの仕様（.at() は「単発トリガー」）への理解が不足していた点も影響しています。

問題に気づいたきっかけは、日曜日に処理がまったく実行されなかったことです。
トリガー一覧を確認すると、実行済みのトリガーしか存在せず、以降の予定がないことにより「これは単発トリガーなのでは」と疑いを持ちました。
ログを確認し、.at() の使用が1回限りの実行しか意味しないことを明確に把握したことで、根本的な設計の見直しを決定しました。

その結果、setTimeTriggers() を毎日01:00に自動実行する「永続的な親トリガー（繰り返し型）」を導入し、当日分の .at() トリガー（削除処理用）を毎日再生成する構成へと改善しました。
これにより、毎日正確に処理が実行され、トリガーの管理もシンプルになるよう設計されています。

【2】処理構成と機能概要

【2-1】setupDailyTriggerForTimeTriggers
・この関数は、初期設定用関数です。
・一度だけ手動で実行すれば、以降は自動処理が永続的に継続されます。
・役割は、毎日01:00に setTimeTriggers() を実行する「親トリガー（繰り返し型）」を作成することです。
・その前に、誤動作を避けるために deleteExistingTriggers() によりすべての既存トリガーを削除します。

【2-2】setTimeTriggers
・この関数は、毎日01:00に実行され、当日分の削除処理トリガーを作成します。
・最初に deleteTimeBasedTriggersOnly() により、既存の「時間ベース（CLOCK）」のトリガーを削除します。
・その後、以下の時間と関数に対応したトリガーを1回限りで作成します。
　※10:30〜14:05まで、5分おきで交互に実行する構成です。

　・10:30 → deleteRowsForNichiReport
　・10:35 → deleteRowsForMynaviReport
　・11:00 → deleteRowsForNichiReport
　・11:05 → deleteRowsForMynaviReport
　・12:00 → deleteRowsForNichiReport
　・12:05 → deleteRowsForMynaviReport
　・13:00 → deleteRowsForNichiReport
　・13:05 → deleteRowsForMynaviReport
　・14:00 → deleteRowsForNichiReport
　・14:05 → deleteRowsForMynaviReport

・この処理により、毎日必要な時間だけ deleteRowsFor〜 関数が1回ずつ実行される状態が維持されます。
・なお、.at() を使用して作成されたトリガーは「1回限り」であるため、この関数が毎日自動実行されることで、日々のトリガーが最新状態で再生成され続けます。

【2-3】createTimeTrigger(functionName, hour, minute)
・この関数は、指定された関数を指定された時刻に1回だけ実行する時間トリガーを作成します。
・現在時刻を取得し、指定時間が過去だった場合は翌日に設定されるよう調整されます。
・作成されたトリガーは .at(triggerTime) により「単発トリガー」として登録されます。
・この処理は setTimeTriggers() によって毎日実行されるため、単発でも問題なく日々の処理が実現されます。
・作成結果はログに記録され、トラブルシュートにも活用できます。

【2-4】deleteTimeBasedTriggersOnly
・この関数は、現在のプロジェクトに設定されているトリガーのうち、時間主導（CLOCK）のトリガーのみを削除します。
・フォーム送信やメニュー操作など、その他のトリガーには影響しません。
・これにより、毎日トリガーを再生成する際に重複が発生しないよう制御しています。

【2-5】deleteExistingTriggers
・この関数は、プロジェクトに存在するすべてのトリガー（種類問わず）を削除します。
・setupDailyTriggerForTimeTriggers() の初回実行時に使用され、完全な初期化を行うための関数です。
・意図しないトリガーを排除する目的で使用します。

【3】運用上の注意事項

【3-1】初期設定の実行は必ず共有アカウントで行う
・トリガーは「実行したアカウントに紐づく」ため、共有アカウントで実行しないと、トリガーが無効になります。
・setTimeTriggers のトリガーのオーナーが「他のユーザー」や「無効」になっている場合は、共有アカウントで再実行が必要です。

【3-2】deleteAllTriggers について（補足用関数）
・運用に混乱が生じた場合、以下のコードを別途定義して実行することで、すべてのトリガー（無効なもの含む）を一括削除可能です。

javascript

コピーする編集する

function deleteAllTriggers() { const triggers = ScriptApp.getProjectTriggers(); triggers.forEach(trigger => { ScriptApp.deleteTrigger(trigger); }); Logger.log("[INFO] 全トリガー削除完了（無効含む）"); } 

【3-3】正しく登録された状態の例
・Apps Script トリガー画面にて、以下の状態になっていれば正常です。

　・setTimeTriggers：時間ベース、毎日01:00、オーナー＝自分（共有アカウント）
　・deleteRowsForNichiReport / deleteRowsForMynaviReport：当日の日付、指定時間、オーナー＝自分（共有アカウント）

【4】導入・引継ぎフローまとめ

【4-1】初期セットアップ時（1回のみ）
・共有アカウントでログイン
・スプレッドシートを開く
・拡張機能 > Apps Script からエディタを開く
・setupDailyTriggerForTimeTriggers() を実行
・権限の承認ダイアログで「許可」選択
・setTimeTriggers のトリガーが 01:00 に作成されていることを確認

【4-2】運用中の確認ポイント（毎日または週1確認）
・当日分の deleteRowsFor〜 トリガーが10:30〜14:05に設定されていることを確認
・異常行が削除されていることをスプレッドシート上で確認
・トリガーが日付ごとに更新されているかを確認

【日次レポートと日次レポート(マイナビ)のシートでカーソル最終行に配置_Enhouse(本番用のシート)の処理概要】
・このスクリプトは、Google スプレッドシートの「日次レポート」と「日次レポート (マイナビ)」の2つのシートにおいて、データが存在する最終行にカーソルを自動で移動させる Google Apps Script（GAS）のコードです。以下のような処理を行います。

・(1)スクリプトの概要
・Google スプレッドシートを開いたときに、自動で「日次レポート」と「日次レポート (マイナビ)」の2つのシートを処理対象とする。
・各シートの中でデータがある最終行を検索する。
・最終行が見つかった場合、その行のA列にカーソルを移動させる。

・(2)コードの動作詳細

・①スクリプトの開始
・function onOpen() { → スプレッドシートを開いたときに自動で実行される関数。
・Logger.log("[INFO] スプレッドシートが開かれました。"); → スクリプトの開始をログに記録する。
・const sheetNames = ["日次レポート", "日次レポート (マイナビ)"]; → 処理対象のシートを指定（2つのシート名を配列に格納）。
・const ss = SpreadsheetApp.getActiveSpreadsheet(); → 現在開いているスプレッドシートを取得。

・②シートの存在確認と処理の実行
・sheetNames.forEach(sheetName => { → 2つのシートをループ処理（1つずつ処理を実行）。
・let sheet = ss.getSheetByName(sheetName); → シートを取得。
・if (sheet) { moveToLastFilledRow(sheet); } → シートが存在する場合、「moveToLastFilledRow」関数を実行。
・else { Logger.log([ERROR] シートが見つかりません: ${sheetName}); } → シートが見つからない場合、エラーメッセージを出力。

・③最終行の検索とカーソル移動
・function moveToLastFilledRow(sheet) { → 最終行を特定してカーソルを移動する関数。
・Logger.log([INFO] スクリプト開始: ${sheet.getName()}); → 処理開始時にログを記録。
・SpreadsheetApp.flush(); → キャッシュをクリアし、最新のデータを確実に取得。

・④最終行の取得
・const lastRow = sheet.getLastRow(); → シート内の最後のデータがある行番号を取得。
・const lastColumn = sheet.getLastColumn(); → シート内の最後のデータがある列番号を取得。
・if (lastRow < 1 || lastColumn < 1) { Logger.log([WARNING] シート「${sheet.getName()}」にデータがありません。処理を終了します。); return; } → データがない場合、処理を終了。
・Logger.log([INFO] 最終行（推定）: ${lastRow}); → 取得した最終行の行番号をログに記録。

・⑤最終行を見つけるための処理（パフォーマンス最適化）
・const checkRowStart = Math.max(lastRow - 2000, 1); → 最後の2000行だけをスキャン対象にしてパフォーマンスを向上。
・let checkRange = sheet.getRange(checkRowStart, 1, lastRow - checkRowStart + 1, lastColumn); → チェック範囲を指定。
・let values = checkRange.getDisplayValues(); → セルの表示値を取得（計算式ではなく実際に表示されている値）。

・⑥最終行の特定
・let foundLastValidRow = -1; → データがある最終行の初期値を設定。
・for (let i = values.length - 1; i >= 0; i--) { → データのある最終行を探すため、後ろからスキャン。
・let hasValue = values[i].some(cell => cell !== "" && cell !== null); → その行にデータが存在するかチェック。
・if (hasValue) { foundLastValidRow = checkRowStart + i; Logger.log([INFO] 最終行（値あり）を発見: ${foundLastValidRow} 行目); break; } → データがある場合、その行を「最終行」として記録。

・⑦カーソルの移動
・if (foundLastValidRow === -1) { Logger.log([ERROR] シート「${sheet.getName()}」で最終行を特定できませんでした。処理を終了します。); return; } → 最終行が特定できなかった場合、エラーログを出力。
・let targetCell = sheet.getRange(foundLastValidRow, 1); → 最終行のA列（1列目）のセルを取得。
・sheet.setActiveSelection(targetCell); → カーソルをそのセルに移動。
・Logger.log([INFO] カーソルを ${foundLastValidRow} 行目のA列に移動しました); → カーソル移動のログを記録。
・Logger.log([INFO] スクリプト完了: ${sheet.getName()}); → 処理完了のログを記録。

・(3)このスクリプトのまとめ
・(1)スプレッドシートが開かれると自動的に処理が開始される。
・(2)「日次レポート」と「日次レポート (マイナビ)」のシートを取得し、それぞれ処理を実行する。
・(3)各シートでデータのある最終行を特定する。
・(4)カーソルをその最終行のA列に移動させる。
・(5)スプレッドシートの更新を反映し、処理を完了する。

・このスクリプトを実行することで、「日次レポート」と「日次レポート (マイナビ)」シートにおいて、データのある最終行にカーソルを自動で移動させることができます。


【日次レポートのみ異常値の行削除_Enhouse(本番用のシート)の処理概要】
・これは、「日次レポート」シートの中から特定の条件に合致する行（基準日より前のデータや異常値を含む行）を削除する Google Apps Script（GAS）のコードです。以下のような処理を行います。

・(1)スクリプトの概要
・Google スプレッドシートの「日次レポート」シートを開き、データを取得する。
・「基準日（前日）」のデータを探し、その日より前のデータを削除する。
・B列～O列の範囲に「1」または「0」がある場合も削除する。
・不要なデータを削除した後、シートを更新する。

・(2)コードの動作詳細
・①スクリプトの開始
・Logger.log("[INFO] スクリプト開始"); → スクリプトが開始されたことをログに記録する。
・const sheetName = "日次レポート"; → 操作対象のシート名を「日次レポート」に設定する。
・const numHeaderRow = 2; → ヘッダー行（2行）を除外する。

・②シートの取得と存在確認
・const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets(); → スプレッドシート内の全シートを取得。
・const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName); → 「日次レポート」シートを取得。
・if (!sheet) { Logger.log("[ERROR] シートが見つかりません: " + sheetName); return; } → シートが見つからない場合、エラーメッセージを出力して処理を終了する。

・③最終行とデータ範囲の取得
・const lastRow = sheet.getLastRow(); → シート内の最後のデータがある行番号を取得。
・const lastColumn = 15; → A列からO列までの範囲（15列）を対象とする。
・if (lastRow <= numHeaderRow) { Logger.log("[WARNING] データが不足しているためスキップ"); return; } → データ行が2行以下の場合、スクリプトをスキップする。
・const range = sheet.getRange(numHeaderRow + 1, 1, lastRow - numHeaderRow, lastColumn).getValues(); → データ範囲（A列～O列）を取得。
・Logger.log("[INFO] データ取得成功, 取得行数: " + range.length); → 取得したデータの行数をログに出力する。

・④基準日の取得
・const today = new Date(); → 現在の日付を取得。
・today.setHours(0, 0, 0, 0); → 時間を 0:00 にリセットする。
・const baseDate = new Date(today); baseDate.setDate(today.getDate() - 1); → 基準日（前日） を取得。
・const baseDateStr = Utilities.formatDate(baseDate, "JST", "yyyy/MM/dd"); → 日付を「yyyy/MM/dd」形式に変換してログに記録。

・⑤基準行（前日データの行）を見つける
・for (let i = range.length - 1; i >= 0; i--) { → 最後の行から逆方向にスキャンする。
・let rowValue = range[i][0]; → A列（1列目）の値を取得。
・let parsedDate = parseDate(rowValue); → A列の値を日付形式に変換。
・if (parsedDate) { → 有効な日付であれば処理を継続。
・if (parsedDateStr === baseDateStr) { → A列の日付が基準日と一致する場合 に基準行を特定。
・Logger.log("[INFO] 基準行を発見: 行 " + lastValidRow + ", 日付=" + parsedDateStr); → 基準行をログに記録。

・⑥削除対象の行を特定
・for (let i = range.length - 1; i >= 0; i--) { → 最後の行から逆方向にスキャン。
・if (currentRowNum <= lastValidRow) { break; } → 基準行に到達したら処理を停止。
・let rowValue = range[i][0]; let parsedDate = parseDate(rowValue); → A列の日付を取得して変換。
・if (parsedDate.getTime() < baseDate.getTime()) { → 基準日より前の日付のデータを削除対象とする。
・for (let j = 1; j < range[i].length; j++) { if (range[i][j] == 1 || range[i][j] == 0) { → B列～O列に「1」または「0」がある場合も削除対象にする。

・⑦行の削除
・rowsToDelete.sort((a, b) => b - a); → 削除を行う際に「大きい行番号から削除」する（小さい行から削除すると、行番号がずれるため）。
・rowsToDelete.forEach(row => { sheet.deleteRow(row); }); → 実際に削除を実行。
・SpreadsheetApp.flush(); → スプレッドシートの変更を確定。

・(3)サブ関数「parseDate」
・parseDate(value) → A列の日付データを適切に解析する関数。
・値の形式ごとに処理を分岐
・Date 型ならそのまま処理。
・数値型（シリアル値）なら Excel 方式の日付に変換。
・文字列型なら日付に変換。

・(4)このスクリプトのまとめ
・(1)スプレッドシートの「日次レポート」シートからデータを取得。
・(2)A列の日付を解析し、前日（基準日）のデータを探す。
・(3)基準日より前の日付の行を削除対象にする。
・(4)B～O列に「1」「0」のデータが含まれる行も削除対象にする。
・(5)削除リストを作成し、大きい行番号から順に削除する。
・(6)スプレッドシートの変更を確定し、処理を完了する。

・このスクリプトを実行することで、「日次レポート」シートの中から 基準日より前の古いデータや不要なデータを自動的に削除 することができます。
