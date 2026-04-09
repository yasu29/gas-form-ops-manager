/**
 * フォーム送信トリガー
 *
 * 設計思想：
 * - 回答時に自動で状態更新
 * - tokenとemailでトレーサビリティ確保
 */
function formops_onFormSubmit(e) {

  try {

    const responses = e.response.getItemResponses();

    /**
     * 🔹 token取得（安全化）
     */
    const tokenItem = responses.find(r =>
      r.getItem().getTitle().trim().toLowerCase() === 'token'
    );

    if (!tokenItem) {
      throw new Error('Token field not found in form responses');
    }

    const token = tokenItem.getResponse().toString().trim();

    /**
     * 🔹 回答者メール取得
     */
    const respondentEmail = e.response.getRespondentEmail() || 'unknown';

    /**
     * 🔹 token検証
     */
    const participant = FormOpsParticipantService.findByToken(token);

    if (!participant) {

      FormOpsLogService.log(
        'ERROR',
        'Invalid token detected',
        `token=${token}, email=${respondentEmail}`
      );

      throw new Error('Invalid token');
    }

    /**
     * 🔹 メール不一致チェック
     */
    if (participant.email !== respondentEmail) {

      FormOpsLogService.log(
        'WARN',
        'Email mismatch detected',
        `token=${token}, expected=${participant.email}, actual=${respondentEmail}`
      );
    }

    /**
     * 🔹 ステータス更新
     */
    FormOpsParticipantService.updateStatus(token, 'Submitted');

    /**
     * 🔥 ログ改善（ここ重要）
     */
    FormOpsLogService.log(
      'INFO',
      'Form submitted',
      `name=${participant.name}, email=${respondentEmail}, token=${token}`
    );

  } catch (e) {

    FormOpsLogService.log(
      'ERROR',
      e.message,
      'onFormSubmit'
    );

    throw e;
  }
}