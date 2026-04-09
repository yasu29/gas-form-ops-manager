/**
 * MailService
 *
 * 設計思想：
 * - メール送信の司令塔
 * - 送信に必要な前処理（Token付与）も内包
 * - Repository経由でデータ取得
 *
 * 役割：
 * 1. Token付与
 * 2. フォームURL取得
 * 3. テンプレート取得
 * 4. 参加者取得
 * 5. メール送信
 * 6. ステータス更新
 * 7. ログ出力
 */
class FormOpsMailService {

  /**
   * 初回メール送信
   */
  static sendInitial() {

    /**
     * 🔥 初期化（最重要）
     *
     * 設計思想：
     * - 毎回新しいアンケートとして扱う
     * - tokenとstatusをリセット
     */
    FormOpsParticipantService.resetAll();

    /**
     * 🔹 テンプレ取得
     */
    const template = FormOpsMailTemplateRepository.getTemplate('initial');

    /**
     * 🔹 テンプレートバリデーション（全体エラー）
     *
     * 設計思想：
     * - 設定ミスは即停止
     * - 無駄なループ防止
     */
    if (!template.subject || !template.subject.trim()) {
      throw new Error('Mail template error: subject is empty');
    }
    
    if (!template.body || !template.body.trim()) {
      throw new Error('Mail template error: body is empty');
    }

    /**
     * 🔹 参加者取得
     */
    const participants = FormOpsParticipantService.getAll();

    participants.forEach(p => {

      try {

        /**
         * 🔹 既回答はスキップ
         */
        if (p.status === 'Submitted') return;

        /**
         * 🔹 URL生成（token付与）
         */
        const url = FormOpsFormService.generatePrefilledUrl(p.token);

        /**
         * 🔹 本文生成
         */
        const body = this.buildBody(template.body, {
          name: p.name,
          url: url
        });
        
        /**
         * 🔹 メール送信
         */
        this.validateMail(p.email);
        this.sendMail(p.email, template.subject, body);

        /**
         * 🔹 ステータス更新
         */
        FormOpsParticipantService.updateStatus(p.token, 'Pending');

        /**
         * 🔹 ログ出力
         */
        FormOpsLogService.log('INFO', 'Mail sent', p.email);

      } catch (e) {

        /**
         * 🔹 エラーログ
         */
        FormOpsLogService.log('ERROR', e.message, p.email);
      }
    });
  }


  /**
   * リマインド送信
   *
   * 設計思想：
   * - 未回答者のみ対象
   * - 初回送信と同一ロジックを再利用
   */
  static sendReminder() {

    /**
     * 🔹 テンプレ取得
     */
    const template = FormOpsMailTemplateRepository.getTemplate('reminder');

    /**
     * 🔹 テンプレートバリデーション（全体エラー）
     *
     * 設計思想：
     * - 設定ミスは即停止
     * - 無駄なループ防止
     */
    if (!template.subject || !template.subject.trim()) {
      throw new Error('Mail template error: subject is empty');
    }
    
    if (!template.body || !template.body.trim()) {
      throw new Error('Mail template error: body is empty');
    }

    /**
     * 🔹 未回答者取得
     */
    const participants = FormOpsParticipantService.getPending();

    participants.forEach(p => {

      try {

        /**
         * 🔹 URL生成
         */
        const url = FormOpsFormService.generatePrefilledUrl(p.token);

        /**
         * 🔹 本文生成
         */
        const body = this.buildBody(template.body, {
          name: p.name,
          url: url
        });

        /**
         * 🔹 メール送信
         */
        this.validateMail(p.email);
        this.sendMail(p.email, template.subject, body);

        /**
         * 🔹 ログ
         */
        FormOpsLogService.log('INFO', 'Reminder sent', p.email);

      } catch (e) {

        FormOpsLogService.log('ERROR', e.message, p.email);
      }
    });
  }


  /**
   * 🔹 宛先バリデーション（個別エラー用）
   *
   * 設計思想：
   * - 参加者単位のチェック
   * - ループ内で使用
   */
  static validateMail(to) {
  
    if (!to || !to.trim()) {
      throw new Error('Mail validation error: to is empty');
    }
  }


  /**
   * 共通メール送信処理
   *
   * 設計：
   * - CC/BCC を Control シートから取得
   * - 全メール送信の一元化
   */
  static sendMail(to, subject, body) {
  
    /**
     * 🔹 バリデーション（最重要）
     */
    this.validateMail(to);

    /**
     * 🔹 CC/BCC取得
     */
    const cc = FormOpsControlRepository.get('mail_cc');
    const bcc = FormOpsControlRepository.get('mail_bcc');
  
    /**
     * 🔹 メール送信
     */
    GmailApp.sendEmail(to, subject, body, {
      cc: cc || undefined,
      bcc: bcc || undefined
    });
  }


  /**
   * 本文生成
   *
   * ${xxx} を置換
   *
   * @param {string} template
   * @param {Object} params
   * @returns {string}
   */
  static buildBody(template, params) {

    let body = template;

    Object.keys(params).forEach(key => {
      body = body.replaceAll(`\${${key}}`, params[key]);
    });

    return body;
  }
}