/**
 * FormOpsConfig
 *
 * 設計思想：
 * - ScriptPropertiesの唯一の入口
 * - 設定値はコードに持たない
 * - 必須チェックをここで行う
 * - 読み書きの両方を一元管理する
 */
class FormOpsConfig {

  /**
   * 管理用スプレッドシートID（必須）
   */
  static get SPREADSHEET_ID() {
    return this.getRequired("FORMOPS_SPREADSHEET_ID");
  }

  /**
   * フォームID（生成後に保存）
   */
  static get FORM_ID() {
    return this.get("FORMOPS_FORM_ID");
  }

  /**
   * token entry ID（必須）
   *
   * 設計思想：
   * - フォームのtoken項目の内部ID
   * - URLプリフィル用
   */
  static get TOKEN_ENTRY_ID() {
    return this.getRequired("FORMOPS_TOKEN_ENTRY_ID");
  }

  /**
   * 出力フォルダ名
   */
  static get OUTPUT_FOLDER_NAME() {
    return 'FormOps_Output';
  }

  // ==============================
  // Setter
  // ==============================

  /**
   * フォームID保存
   *
   * 設計思想：
   * - ScriptPropertiesの唯一の書き込み入口
   */
  static setFormId(formId) {

    if (!formId) {
      throw new Error('Invalid formId');
    }

    PropertiesService
      .getScriptProperties()
      .setProperty('FORMOPS_FORM_ID', formId);
  }

  /**
   * token entry ID保存
   *
   * 設計思想：
   * - フォーム生成後に自動登録
   */
  static setTokenEntryId(entryId) {

    if (!entryId) {
      throw new Error('Invalid entryId');
    }

    PropertiesService
      .getScriptProperties()
      .setProperty('FORMOPS_TOKEN_ENTRY_ID', entryId);
  }

  // ==============================
  // 共通処理
  // ==============================

  static get(key, defaultValue = null) {

    const value = PropertiesService
      .getScriptProperties()
      .getProperty(key);

    return value ?? defaultValue;
  }

  static getRequired(key) {

    const value = this.get(key);

    if (!value) {
      throw new Error(`Missing required config: ${key}`);
    }

    return value;
  }

  // ==============================
  // Environment
  // ==============================
  
  /**
   * 環境状態保存
   */
  static setEnvironmentState(state) {
  
    PropertiesService.getScriptProperties()
      .setProperty('FORMOPS_ENV', JSON.stringify(state));
  }
  
  /**
   * 環境状態取得
   */
  static getEnvironmentState() {
  
    const value = PropertiesService.getScriptProperties()
      .getProperty('FORMOPS_ENV');
  
    return value ? JSON.parse(value) : {};
  }
}