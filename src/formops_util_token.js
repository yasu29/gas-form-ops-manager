/**
 * Token生成ユーティリティ
 *
 * 設計思想：
 * - 一意な識別子を生成
 * - URL安全
 */
class FormOpsTokenUtil {

  /**
   * トークン生成
   *
   * @returns {string}
   */
  static generate() {

    const uuid = Utilities.getUuid();

    /**
     * ハイフン除去（URLを短くする）
     */
    return uuid.replace(/-/g, '');
  }
}


/**
 * RunId生成ユーティリティ
 *
 * 設計思想：
 * - 実行単位を一意に識別
 * - ファイル・ログ・フォルダを紐づけるキー
 * - 人間が読める形式（日時ベース）
 */
class FormOpsRunUtil {

  /**
   * RunId生成
   *
   * 例：
   * 20260408_153000
   *
   * @returns {string}
   */
  static generate() {

    const now = new Date();

    return Utilities.formatDate(
      now,
      Session.getScriptTimeZone(),
      'yyyyMMdd_HHmmss'
    );
  }
}