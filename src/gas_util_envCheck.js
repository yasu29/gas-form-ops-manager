/**
 * =========================================================
 * 🔹 GAS Environment Check Utility
 * =========================================================
 *
 * ■ 概要
 * 実行環境で利用可能なAPI・権限を事前にチェックするためのユーティリティ
 *
 * ■ 想定用途
 * ・企業Google Workspace環境での制限検知
 * ・GASツールの事前動作確認
 * ・ポートフォリオの品質向上
 *
 * =========================================================
 *
 * ■ 🔰 基本的な使い方
 *
 * ① そのまま実行
 *
 *   gasEnv_check();
 *
 * → Loggerに結果が出力される
 *
 *
 * ■ 🔧 カスタマイズ方法
 *
 * ① チェック内容を変更する
 *
 *   const checks = [
 *     {
 *       name: 'Spreadsheet.access',
 *       test: () => SpreadsheetApp.getActive()
 *     }
 *   ];
 *
 *   gasEnv_runChecks(checks);
 *
 *
 * ■ 💡 実務での使い方（おすすめ）
 *
 * ・新しいGASプロジェクト作成時に一度実行
 * ・エラーが出る環境で原因切り分け
 *
 *
 * ■ ⚠ 注意点
 *
 * ・テスト用のファイル/フォルダを一時作成します
 * ・処理後はゴミ箱へ移動されます（完全削除ではありません）
 * ・既存ファイルには影響しません（UUIDで識別）
 *
 * =========================================================
 */


/**
 * =========================================================
 * 🔹 安全クリーンアップ
 * =========================================================
 *
 * 作成したリソースのみ削除する
 * ※既存ファイルには影響しない
 */
function gasEnv_safeCleanup(file, folder) {

  try {
    if (file) file.setTrashed(true);
  } catch (e) {}

  try {
    if (folder) folder.setTrashed(true);
  } catch (e) {}
}


/**
 * =========================================================
 * 🔹 コア実行エンジン
 * =========================================================
 *
 * @param {Array} checks - チェック定義配列
 * @returns {Array} 実行結果
 */
function gasEnv_runChecks(checks) {

  return checks.map(check => {

    try {

      const result = check.test();

      return {
        name: check.name,
        status: 'OK',
        message: result || ''
      };

    } catch (e) {

      return {
        name: check.name,
        status: 'NG',
        message: e.message
      };
    }

  });
}


/**
 * =========================================================
 * 🔹 デフォルトチェック（GAS用）
 * =========================================================
 *
 * 必要に応じて編集・拡張可能
 */
function gasEnv_getDefaultChecks() {

  const PREFIX = 'GAS_ENV_TEST_';

  return [

    /**
     * 🔹 Drive: フォルダ作成
     */
    {
      name: 'Drive.createFolder',
      test: () => {

        let folder = null;

        try {

          const name = PREFIX + Utilities.getUuid();
          folder = DriveApp.createFolder(name);

        } finally {

          gasEnv_safeCleanup(null, folder);
        }
      }
    },

    /**
     * 🔹 Drive: ファイル移動
     */
    {
      name: 'Drive.moveTo',
      test: () => {

        let file = null;
        let folder = null;

        try {

          const id = Utilities.getUuid();

          file = DriveApp.createFile(PREFIX + id + '.txt', 'test');
          folder = DriveApp.createFolder(PREFIX + id);

          file.moveTo(folder);

        } finally {

          gasEnv_safeCleanup(file, folder);
        }
      }
    },

    /**
     * 🔹 Drive: ルート取得
     */
    {
      name: 'Drive.getRootFolder',
      test: () => {
        const folder = DriveApp.getRootFolder();
        return folder.getName();
      }
    },

    /**
     * 🔹 Drive: 削除権限
     */
    {
      name: 'Drive.setTrashed',
      test: () => {

        let file = null;

        try {

          const name = PREFIX + Utilities.getUuid() + '.txt';
          file = DriveApp.createFile(name, 'test');

        } finally {

          gasEnv_safeCleanup(file, null);
        }
      }
    },

    /**
     * 🔹 Gmail: アクセス確認
     */
    {
      name: 'Gmail.getAliases',
      test: () => {
        GmailApp.getAliases();
      }
    }

  ];
}


/**
 * =========================================================
 * 🔹 実行エントリ（標準チェック）
 * =========================================================
 *
 * @returns {Array} 実行結果
 */
function gasEnv_check() {

  const checks = gasEnv_getDefaultChecks();
  const results = gasEnv_runChecks(checks);

  /**
   * 🔹 ログ出力
   */
  Logger.log(JSON.stringify(results, null, 2));

  return results;
}