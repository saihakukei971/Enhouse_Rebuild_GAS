/**
 * 初期設定関数：一度実行するだけで、以後は毎日トリガーが自動更新される
 */
function setupDailyTriggerForTimeTriggers() {
  // setTimeTriggers のトリガーがあるか確認し、なければ作成
  if (!isTriggerExists("setTimeTriggers")) {
    ScriptApp.newTrigger("setTimeTriggers")
      .timeBased()
      .everyDays(1)
      .atHour(1)
      .create();
    Logger.log("[INFO] setTimeTriggers の日次トリガーを作成しました（毎日01:00）");
  } else {
    Logger.log("[INFO] setTimeTriggers の日次トリガーはすでに存在しています");
  }

  logAllTriggers(); // 現在のトリガー一覧を記録
}

/**
 * 毎日実行されて、当日の削除処理トリガーをセットする関数
 */
function setTimeTriggers() {
  Logger.log("[INFO] トリガー設定開始");

  // 時間ベーストリガーのみ削除（setTimeTriggers 自体は削除しない）
  deleteTimeBasedTriggersOnly();

  // 指定時刻に実行される1回限りのトリガーを追加
  createTimeTrigger("deleteRowsForNichiReport", 10, 30);
  createTimeTrigger("deleteRowsForMynaviReport", 10, 35);
  createTimeTrigger("deleteRowsForNichiReport", 11, 0);
  createTimeTrigger("deleteRowsForMynaviReport", 11, 5);
  createTimeTrigger("deleteRowsForNichiReport", 12, 0);
  createTimeTrigger("deleteRowsForMynaviReport", 12, 5);
  createTimeTrigger("deleteRowsForNichiReport", 13, 0);
  createTimeTrigger("deleteRowsForMynaviReport", 13, 5);
  createTimeTrigger("deleteRowsForNichiReport", 14, 0);
  createTimeTrigger("deleteRowsForMynaviReport", 14, 5);

  Logger.log("[INFO] トリガー設定完了（本日または翌日分を作成）");
  logAllTriggers(); // 設定後のトリガー一覧を記録
}

/**
 * 指定した関数を、指定時刻に「1回限り」で実行するトリガーを作成
 */
function createTimeTrigger(functionName, hour, minute) {
  const now = new Date();
  let triggerTime = new Date(now.getFullYear(), now.getMonth(), now.getDate(), hour, minute, 0);

  if (triggerTime < now) {
    triggerTime.setDate(triggerTime.getDate() + 1);
  }

  ScriptApp.newTrigger(functionName)
    .timeBased()
    .at(triggerTime)
    .create();

  Logger.log(`[INFO] トリガー作成: ${functionName} → ${triggerTime.toLocaleString()}`);
}

/**
 * 時間ベースのトリガーだけを削除（setTimeTriggers のトリガーは残す）
 */
function deleteTimeBasedTriggersOnly() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getTriggerSource() === ScriptApp.TriggerSource.CLOCK && trigger.getHandlerFunction() !== "setTimeTriggers") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  Logger.log("[INFO] setTimeTriggers を除く時間ベース（CLOCK）の既存トリガーを削除しました");
}

/**
 * すべてのトリガーを削除（フルリセット用途）
 */
function deleteExistingTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
  Logger.log("[INFO] すべての既存トリガーを削除しました（完全初期化）");
}

/**
 * トリガーが存在するか確認
 */
function isTriggerExists(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  return triggers.some(trigger => trigger.getHandlerFunction() === functionName);
}

/**
 * 現在のトリガー一覧をログに出力
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
