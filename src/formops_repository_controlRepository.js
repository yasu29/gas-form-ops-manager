/**
 * ControlRepository
 *
 * 設計思想：
 * - Controlシートを「設定ストア」として扱う
 * - key-value形式で設定を取得
 * - Service層から直接Spreadsheetを触らせない
 */
class FormOpsControlRepository {

  /**
   * 設定値取得
   *
   * @param {string} key
   * @returns {string}
   */
  static get(key) {

    const values = FormOpsSheetRepository.getAllValues('Control');

    /**
     * 🔹 ヘッダー除外
     */
    const rows = values.slice(1);

    /**
     * 🔹 key検索
     */
    const row = rows.find(r => r[0] === key);

    /**
     * 🔹 見つからなければ空
     */
    return row ? row[1] : '';
  }
}