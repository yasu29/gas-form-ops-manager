/**
 * SheetProtectionService
 *
 * 設計思想：
 * - 回答データの「手動編集」を防止する
 * - Google Formの自動書き込みは許可する（GAS/フォームは制限されない）
 * - シート単位で保護することで、過剰なProtection増殖を防ぐ
 *
 * 前提：
 * - フォームの回答先スプレッドシートが既に存在していること
 * - フォームは setDestination() 済みであること
 *
 * 注意：
 * - オーナー（スクリプト実行者）は編集可能のまま残る
 * - 閲覧権限ユーザーはコピー可能（仕様）
 */
class FormOpsSheetProtectionService {

  /**
   * 回答シートを保護する
   *
   * @param {string} spreadsheetId
   *   フォーム回答先のスプレッドシートID
   *
   * 処理概要：
   * 1. スプレッドシートを取得
   * 2. 回答シート（Form Responses系）を特定
   * 3. シート全体にProtectionを設定
   * 4. 編集者を削除（オーナーは除く）
   */
  static protectResponseSheet(spreadsheetId) {

    /**
     * 🔹 スプレッドシート取得
     *
     * - FormApp.getDestinationId() で取得したIDを想定
     */
    const ss = SpreadsheetApp.openById(spreadsheetId);

    /**
     * 🔹 回答シートの特定
     *
     * Google Formの仕様：
     * - 回答シートは "Form Responses 1" のような名前で生成される
     * - 複数存在する場合 "Form Responses 2" などになる
     *
     * 設計方針：
     * - 完全一致ではなく prefix一致で検出
     * - 将来の命名揺れに耐性を持たせる
     */
    const sheet = ss.getSheets().find(s =>
      s.getName().startsWith('Form Responses')
    );

    /**
     * 🔹 フェイルセーフ
     *
     * 想定外：
     * - フォームの回答先設定が未実行
     * - シート名が変更されている
     */
    if (!sheet) {
      throw new Error('Response sheet not found. Ensure form destination is set.');
    }

    /**
     * 🔹 Protection作成
     *
     * - シート全体を対象に保護
     * - 行単位ではなくシート単位（パフォーマンス・管理性のため）
     */
    const protection = sheet.protect();

    protection.setDescription('FormOps: Response sheet protection');

    /**
     * 🔹 編集者の削除
     *
     * removeEditors():
     * - 明示的に追加された編集者を削除
     *
     * 注意：
     * - スプレッドシートのオーナーは削除されない（仕様）
     * - GAS実行ユーザーも実質編集可能
     */
    const editors = protection.getEditors();

    if (editors.length > 0) {
      protection.removeEditors(editors);
    }

    /**
     * 🔹 ログ出力
     *
     * - 実行確認用
     * - トラブルシュート時に有用
     */
    FormOpsLogService.log(
      'INFO',
      'Protected response sheet',
      sheet.getName()
    );    
  }
}