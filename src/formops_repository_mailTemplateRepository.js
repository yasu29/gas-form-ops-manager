/**
 * MailTemplateRepository
 *
 * 設計思想：
 * - メールテンプレートの取得を一元管理
 * - シート構造に依存させない
 *
 * 対象シート：
 * MailTemplates
 *
 * カラム構成：
 * [templateKey, subject, body]
 */
class FormOpsMailTemplateRepository {

  /**
   * テンプレート取得
   *
   * @param {string} templateKey
   * @returns {{subject: string, body: string}}
   */
  static getTemplate(templateKey) {

    const sheet = FormOpsSheetRepository.getSheet('MailTemplates');

    /**
     * 🔹 全データ取得
     */
    const data = sheet.getDataRange().getValues();

    /**
     * 🔹 1行目はヘッダーのためスキップ
     */
    for (let i = 1; i < data.length; i++) {

      const key = data[i][0];

      if (key === templateKey) {

        return {
          subject: data[i][1],
          body: data[i][2]
        };
      }
    }

    /**
     * 🔹 見つからない場合は例外
     */
    throw new Error(`Template not found: ${templateKey}`);
  }
}