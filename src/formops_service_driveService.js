/**
 * =========================================================
 * 🔹 Drive Service（環境差分吸収）
 * =========================================================
 *
 * ■ 目的
 * ・Google Workspace制限環境でも安全に動作させる
 * ・Drive操作の失敗を吸収する
 *
 * =========================================================
 */

class FormOpsDriveService {

  /**
   * 🔹 Drive利用可否チェック
   */
  static isAvailable() {
  
    const env = FormOpsConfig.getEnvironmentState();
  
    if (env.driveAvailable !== undefined) {
      return env.driveAvailable;
    }
  
    try {
      DriveApp.getRootFolder();
      return true;
    } catch (e) {
      return false;
    }
  }

  /**
   * 🔹 フォルダ作成
   */
  static createFolder(name) {

    if (!this.isAvailable()) {
      console.warn('Drive unavailable: createFolder skipped');
      return null;
    }

    return DriveApp.createFolder(name);
  }

  /**
   * 🔹 ファイル移動
   */
  static moveFile(file, folder) {

    if (!this.isAvailable()) {
      console.warn('Drive unavailable: moveFile skipped');
      return;
    }

    file.moveTo(folder);
  }

  /**
   * 🔹 ルート取得
   */
  static getRootFolder() {

    if (!this.isAvailable()) {
      return null;
    }

    return DriveApp.getRootFolder();
  }

  /**
   * 🔹 ファイル移動（ID指定・安全版）
   *
   * 設計思想：
   * ・Drive制限環境で例外を吸収
   * ・getFileById / moveTo の失敗を握りつぶす
   * ・呼び出し元に影響を与えない
   *
   * @param {string} fileId
   * @param {Folder} folder
   */
  static moveFileById(fileId, folder) {

    /**
     * 🔹 Drive利用不可ならスキップ
     */
    if (!this.isAvailable()) {
      console.warn('Drive unavailable: moveFileById skipped');
      return;
    }

    try {

      /**
       * 🔹 ファイル取得
       */
      const file = DriveApp.getFileById(fileId);

      /**
       * 🔹 フォルダへ移動
       */
      file.moveTo(folder);

    } catch (e) {

      /**
       * 🔸 制限環境では何もしない（重要）
       */
      console.warn('Drive move skipped (restricted): ' + e.message);
    }
  }

  /**
   * 🔹 フォルダ取得 or 作成（安全版）
   *
   * 設計思想：
   * ・Workspace制限でも落ちない
   * ・取得失敗時はnull返却
   */
  static getOrCreateFolder(name) {
  
    if (!this.isAvailable()) {
      console.warn('Drive unavailable: getOrCreateFolder skipped');
      return null;
    }
  
    try {
  
      const folders = DriveApp.getFoldersByName(name);
  
      return folders.hasNext()
        ? folders.next()
        : DriveApp.createFolder(name);
  
    } catch (e) {
  
      console.warn('Drive folder access failed: ' + e.message);
      return null;
    }
  }
}