/**
 * FormTemplateEngine
 *
 * 設計思想：
 * - フォーム定義を「データ」として扱う
 * - スプレッドシートからテンプレートを構築する
 *
 * 役割：
 * 1. テンプレート読み込み
 * 2. 質問オブジェクト生成
 */
class FormOpsTemplateEngine {

  /**
   * テンプレート読み込み
   */
  static load() {
  
    const values = FormOpsSheetRepository.getAllValues('FormTemplates');
  
    const headers = values[0];
    const rows = values.slice(1);
  
    /**
     * 🔹 空行・不正行を除外
     */
    const filtered = rows.filter(row => {
      return row[0] !== '' && row[1] !== '' && row[2] !== '';
    });
  
    /**
     * 🔹 質問オブジェクト化
     */
    const questions = filtered.map(row => {
      return {
        order: Number(row[0]),
        type: row[1],
        title: row[2],
        // required: row[3] === 'TRUE',
        required: String(row[3]).toUpperCase() === 'TRUE',
        options: row[4] ? row[4].split(',').map(v => v.trim()) : []
      };
    });
  
    /**
     * 🔹 表示順ソート
     */
    questions.sort((a, b) => a.order - b.order);
  
    return questions;
  }
}