/**
 * ParticipantRepository
 *
 * 設計思想：
 * - 参加者データの取得を一元管理
 * - シート構造を隠蔽
 *
 * 対象シート：
 * Participants
 *
 * カラム構成：
 * [name, email, token, status]
 */
class FormOpsParticipantRepository {

  /**
   * 全参加者取得
   *
   * @returns {Array<Object>}
   */
  static getAll() {

    const sheet = FormOpsSheetRepository.getSheet('Participants');

    /**
     * 🔹 全データ取得
     */
    const data = sheet.getDataRange().getValues();

    const result = [];

    /**
     * 🔹 1行目はヘッダー
     */
    for (let i = 1; i < data.length; i++) {

      const name = data[i][0];
      const email = data[i][1];
      const token = data[i][2];
      const status = data[i][3];

      /**
       * 🔹 メール未設定はスキップ
       */
      if (!email) continue;

      result.push({
        name: name,
        email: email,
        token: token,
        status: status,
        row: i + 1  // シート行番号
      });
    }

    return result;
  }


  /**
   * ステータス更新
   *
   * @param {number} row
   * @param {string} status
   */
  static updateStatus(row, status) {

    const sheet = FormOpsSheetRepository.getSheet('Participants');

    sheet.getRange(row, 4).setValue(status);
  }
}