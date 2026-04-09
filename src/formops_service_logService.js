/**
 * LogService
 *
 * 設計思想：
 * - ログ出力の一元管理
 * - スプレッドシートへ記録
 * - シンプルな構造（過剰な機能は持たない）
 *
 * 対象シート：
 * Logs
 *
 * カラム構成：
 * [datetime, level, message, data]
 */
class FormOpsLogService {

  /**
   * ログ出力
   *
   * @param {string} level   ログレベル（INFO / WARN / ERROR など）
   * @param {string} message メッセージ
   * @param {string} data    任意データ
   */
  static log(level, message, data = '') {
  
    let sheet;
  
    try {
      sheet = FormOpsSheetRepository.getSheet('Logs');
    } catch (e) {
      /**
       * 🔹 fallback（初期化前）
       */
      Logger.log(`[${level}] ${message} ${data}`);
      return;
    }
  
    sheet.appendRow([
      new Date(),
      level,
      message,
      data
    ]);
  }
}