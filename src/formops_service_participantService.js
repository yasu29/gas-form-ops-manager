/**
 * ParticipantService
 *
 * 設計思想：
 * - Participantsの状態管理
 * - Token付与責務
 * - 空行は処理対象外とする
 */
class FormOpsParticipantService {

  /**
   * 有効データ判定
   */
  static isValid(row) {
    return row[1] && row[1].toString().trim() !== ''; // emailが存在するか
  }

  /**
   * 全参加者取得（空行除外）
   */
  static getAll() {

    const values = FormOpsSheetRepository.getAllValues('Participants');
    const rows = values.slice(1);

    return rows
      .filter(row => this.isValid(row))
      .map(row => ({
        name: row[0],
        email: row[1],
        token: row[2],
        status: row[3]
      }));
  }

  /**
   * ステータス更新
   */
  static updateStatus(token, status) {

    const sheet = FormOpsSheetRepository.getOrCreateSheet('Participants');
    const values = sheet.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {

      /**
       * 🔹 空行スキップ
       */
      if (!this.isValid(values[i])) continue;

      if (values[i][2] && values[i][2].toString().trim() === token) {

        FormOpsParticipantRepository.updateStatus(i + 1, status);
        return;
      }
    }
  
    throw new Error(`Participant not found for token: ${token}`);
  }

  /**
   * tokenから参加者を特定
   *
   * 設計思想：
   * - token一致で一意に特定
   * - onSubmitの検証で使用
   *
   * @param {string} token
   * @returns {Object|null}
   */
  static findByToken(token) {

    const all = this.getAll();

    return all.find(p =>
      p.token &&
      p.token.toString().trim() === token
    ) || null;
  }

  /**
   * 未回答者取得（空行除外）
   */
  static getPending() {

    const all = this.getAll();

    return all.filter(p =>
      p.status !== 'Submitted'
    );
  }

  /**
   * 参加者データ初期化
   *
   * 設計思想：
   * - 毎回クリーンな状態でアンケート開始
   * - tokenとstatusを強制リセット
   * - 再利用時の事故防止
   */
  static resetAll() {

    const sheet = FormOpsSheetRepository.getOrCreateSheet('Participants');
    const values = sheet.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {

      /**
       * 🔹 空行スキップ
       */
      if (!this.isValid(values[i])) continue;

      /**
       * 🔹 token再生成
       */
      const newToken = FormOpsTokenUtil.generate();

      sheet.getRange(i + 1, 3).setValue(newToken);

      /**
       * 🔹 status初期化
       */
      sheet.getRange(i + 1, 4).setValue('Pending');
    }

    FormOpsLogService.log('INFO', 'Participants reset', 'All tokens regenerated');
  }
}