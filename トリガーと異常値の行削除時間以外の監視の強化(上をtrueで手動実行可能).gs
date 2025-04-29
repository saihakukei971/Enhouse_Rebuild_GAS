function manageTriggers(force = true) {
  const now = new Date();
  const hour = now.getHours();

  // **10:00～15:00 の間は監視スキップ（手動実行時は例外）**
  if (hour >= 10 && hour < 15 && !force) {
    Logger.log("[SKIP] 10:00～15:00 の時間帯なので監視をスキップ");
    return;
  }

  Logger.log("[INFO] トリガー管理開始");

  // **現在のトリガー一覧をログ出力**
  logAllTriggers();

  // **不要なトリガーを削除**
  deleteAllTriggersExceptSetTimeTriggers();

  // **01:00 に setTimeTriggers を作成（または手動実行時に強制作成）**
  deleteTrigger("setTimeTriggers");
  createDailyTrigger("setTimeTriggers", 1);
  Logger.log("[INFO] setTimeTriggers を削除し、再作成しました");

  // **削除処理のトリガーを設定**
  if (hour === 1 || force) {
    setupDailyDeleteTriggers();
  }

  // **2時間ごとの監視トリガーをセット**
  deleteTrigger("manageTriggers");
  setupRecurringTrigger();
}

/**
 * **setTimeTriggers を除くすべてのトリガーを削除**
 */
function deleteAllTriggersExceptSetTimeTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() !== "setTimeTriggers") {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`[DELETE] ${trigger.getHandlerFunction()} のトリガーを削除`);
    }
  });
  Logger.log("[INFO] setTimeTriggers を除く不要なトリガーを削除しました");
}

/**
 * setTimeTriggers の削除処理を1日ごとに作成
 */
function setupDailyDeleteTriggers() {
  Logger.log("[INFO] 削除処理トリガーの設定開始");

  // **既存の削除系トリガーを削除**
  deleteSpecificTriggers(["deleteRowsForNichiReport", "deleteRowsForMynaviReport"]);

  const schedule = [
    [10, 30], [10, 35],
    [11, 0],  [11, 5],
    [12, 0],  [12, 5],
    [13, 0],  [13, 5],
    [14, 0],  [14, 5]
  ];

  schedule.forEach(([hour, minute]) => {
    createOneTimeTrigger("deleteRowsForNichiReport", hour, minute);
    createOneTimeTrigger("deleteRowsForMynaviReport", hour, minute + 5);
  });

  Logger.log("[INFO] 削除処理トリガーの設定完了");
}

/**
 * 2時間ごとの監視トリガーを設定（修正済み）
 */
function setupRecurringTrigger() {
  Logger.log("[INFO] 2時間ごとの監視トリガーを設定開始");

  ScriptApp.newTrigger("manageTriggers")
    .timeBased()
    .everyHours(2)  // **← 修正！**
    .create();

  Logger.log("[INFO] 2時間ごとの監視トリガーの設定完了");
}

/**
 * 指定トリガーを削除（関数名を指定）
 */
function deleteTrigger(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`[DELETE] ${functionName} のトリガーを削除`);
    }
  });
}

/**
 * 指定された関数のトリガーをすべて削除
 */
function deleteSpecificTriggers(functionNames) {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (functionNames.includes(trigger.getHandlerFunction())) {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`[DELETE] ${trigger.getHandlerFunction()} のトリガーを削除`);
    }
  });
  Logger.log("[INFO] 不要な削除系トリガーを削除しました");
}

/**
 * **現在のトリガー一覧をログに出力**
 */
function logAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.length === 0) {
    Logger.log("[INFO] 現在、登録されているトリガーはありません");
    return;
  }

  Logger.log("[INFO] 現在のトリガー一覧:");
  triggers.forEach(trigger => {
    Logger.log(`- 関数: ${trigger.getHandlerFunction()}, タイプ: ${trigger.getTriggerSource()}`);
  });
}

/**
 * 指定時刻に1回だけ実行するトリガーを作成（過去のトリガーを作らない）
 */
function createOneTimeTrigger(functionName, hour, minute) {
  const now = new Date();
  let triggerTime = new Date(now.getFullYear(), now.getMonth(), now.getDate(), hour, minute, 0);

  if (triggerTime < now) {
    Logger.log(`[SKIP] 過去の時間 (${triggerTime.toLocaleString()}) のためスキップ`);
    return;
  }

  ScriptApp.newTrigger(functionName)
    .timeBased()
    .at(triggerTime)
    .create();

  Logger.log(`[INFO] ${functionName} の1回限りのトリガーを作成 → ${triggerTime.toLocaleString()}`);
}
