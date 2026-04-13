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
}