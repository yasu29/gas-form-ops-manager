/**
 * FormOps Starter
 *
 * 設計思想：
 * - ユースケース単位の実行入口（Gateway）
 * - すべての操作はここから実行する
 * - Service層を呼び出すだけに限定
 *
 * 役割：
 * 1. スプレッドシート初期化
 * 2. フォーム生成
 * 3. 初回メール送信
 * 4. リマインド送信
 */


/**
 * 初期セットアップ
 *
 * フロー：
 * - 各シート生成
 * - 初期データ投入
 */
function formops_setupAll() {

  /**
   * 🔹 初期化前はLoggerのみ（安全）
   */
  Logger.log('START: setupAll');

  try {

    formops_setup();

    /**
     * 🔹 Logsシート作成後なのでOK
     */
    FormOpsLogService.log('INFO', 'END: setupAll');


  } catch (e) {

    /**
     * 🔹 フォールバック（確実に残す）
     */
    Logger.log(e.message);
    throw e;
  }
}


/**
 * フォーム生成
 */
function formops_createForm() {

  FormOpsLogService.log('INFO', 'START: createForm');

  try {

    const form = FormOpsFormService.createForm();

    formops_setupFormSubmitTrigger(form.getId());

    const formUrl = FormOpsFormService.getFormUrl();

    FormOpsLogService.log('INFO', 'Form URL (response)', formUrl);

    FormOpsLogService.log('INFO', 'END: createForm');

  } catch (e) {

    FormOpsLogService.log('ERROR', e.message, 'createForm');
    throw e;
  }
}


/**
 * 初回メール送信
 */
function formops_sendInitial() {

  FormOpsLogService.log('INFO', 'START: sendInitial');

  try {

    FormOpsMailService.sendInitial();

    FormOpsLogService.log('INFO', 'END: sendInitial');

  } catch (e) {

    FormOpsLogService.log('ERROR', e.message, 'sendInitial');
    throw e;
  }
}


/**
 * リマインド送信
 */
function formops_sendReminder() {

  FormOpsLogService.log('INFO', 'START: sendReminder');

  try {

    FormOpsMailService.sendReminder();

    FormOpsLogService.log('INFO', 'END: sendReminder');

  } catch (e) {

    FormOpsLogService.log('ERROR', e.message, 'sendReminder');
    throw e;
  }
}


/**
 * フォーム送信トリガー作成（再生成対応版）
 *
 * 設計思想：
 * - フォーム送信時イベントを自動登録
 * - フォーム再作成時の不整合を防ぐ
 * - トリガーは「常に最新フォームに紐づく」状態を保証する
 *
 * 重要ポイント：
 * - 既存トリガーは必ず削除（上書きではなく再生成）
 * - GASのトリガーはフォームIDに紐づくため、使い回し不可
 *
 * @param {string} formId フォームID
 */
function formops_setupFormSubmitTrigger(formId) {

  /**
   * 🔹 入力チェック
   *
   * 設計思想：
   * - フォーム未生成状態での実行を防ぐ
   */
  if (!formId) {
    throw new Error('FORM_ID is not set');
  }

  /**
   * 🔹 既存トリガー取得
   *
   * 設計思想：
   * - 同一ハンドラのトリガーを特定する
   */
  const triggers = ScriptApp.getProjectTriggers();

  /**
   * 🔥 既存トリガー削除（最重要）
   *
   * 理由：
   * - フォームを作り直すとIDが変わる
   * - トリガーは古いフォームに紐づいたままになる
   * - よって毎回削除 → 再作成が正解
   */
  triggers.forEach(t => {

    if (t.getHandlerFunction() === 'formops_onFormSubmit') {
      ScriptApp.deleteTrigger(t);

      FormOpsLogService.log(
        'INFO',
        'Old trigger deleted',
        t.getUniqueId()
      );
    }
  });

  /**
   * 🔹 新規トリガー作成
   *
   * 設計思想：
   * - formIdを直接指定して安定性確保
   * - 最新フォームに確実に紐づける
   */
  ScriptApp.newTrigger('formops_onFormSubmit')
    .forForm(formId)
    .onFormSubmit()
    .create();

  /**
   * 🔹 ログ出力
   */
  FormOpsLogService.log(
    'INFO',
    'Form submit trigger created',
    formId
  );
}