/**
 * SheetRepository
 *
 * 設計思想：
 * - Spreadsheet操作の責務を集約
 * - 他レイヤーから直接SpreadsheetAppを触らせない
 * - Spreadsheetは「注入 or Config」で取得
 */
class FormOpsSheetRepository {

  /**
   * 🔹 Spreadsheetを明示的にセット
   */
  static setSpreadsheet(ss) {
    FormOpsSheetRepository._spreadsheet = ss;
  }

  /**
   * 🔹 Spreadsheet取得
   *
   * 設計思想：
   * - 優先：setされたSpreadsheet
   * - fallback：ScriptProperty
   */
  static getSpreadsheet() {

    /**
     * 🔹 優先：注入されたもの
     */
    if (FormOpsSheetRepository._spreadsheet) {
      return FormOpsSheetRepository._spreadsheet;
    }

    /**
     * 🔹 fallback：Config経由
     */
    const ssId = FormOpsConfig.SPREADSHEET_ID;

    if (!ssId) {
      throw new Error('Spreadsheet ID is not set');
    }

    return SpreadsheetApp.openById(ssId);
  }

  /**
   * 🔹 シート取得
   */
  static getSheet(sheetName) {

    const sheet = this.getSpreadsheet().getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`Sheet not found: ${sheetName}`);
    }

    return sheet;
  }

  /**
   * 🔹 シート取得（なければ作成）
   */
  static getOrCreateSheet(sheetName) {

    const ss = this.getSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    return sheet;
  }

  /**
   * 🔹 全データ取得
   */
  static getAllValues(sheetName) {

    const sheet = this.getSheet(sheetName);

    return sheet.getDataRange().getValues();
  }
}

/**
 * 🔹 クラス外で初期化（重要）
 *
 * GAS互換のためここで定義する
 */
FormOpsSheetRepository._spreadsheet = null;