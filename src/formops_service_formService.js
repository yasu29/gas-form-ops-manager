/**
 * FormService
 *
 * 設計思想：
 * - フォーム生成と取得の司令塔
 * - 設定はControlRepositoryから取得
 * - ScriptPropertiesには直接アクセスしない
 */
class FormOpsFormService {

  /**
   * フォーム生成（単一フォーム前提）
   *
   * @returns {GoogleAppsScript.Forms.Form}
   */
  static createForm() {
  
    /**
     * 🔹 RunId生成（実行単位識別）
     */
    const runId = FormOpsRunUtil.generate();
  
    /**
     * 🔹 Run開始ログ
     */
    FormOpsLogService.log('INFO', 'Run started', runId);

    /**
     * 🔹 フォームタイトル取得（Control）
     */
    const formTitle = FormOpsControlRepository.get('form_title');

    /**
     * 🔹 質問テンプレート取得
     */
    const questions = FormOpsTemplateEngine.load();

    /**
     * 🔹 質問数チェック（重要）
     *
     * 設計思想：
     * - 空フォーム防止
     * - 設定ミスの早期検知
     */
    if (!questions || questions.length === 0) {
    
      /**
       * 🔸 処理停止
       */
      throw new Error(
        'Form creation error: no questions defined. Check FormTemplates sheet.'
      );
    }

    /**
     * 🔹 Runフォルダ作成
     */
    const runFolder = this.createRunFolder(runId, formTitle);
  
    /**
     * 🔹 フォーム生成
     */
    const form = FormApp.create(formTitle || 'アンケート');

    /**
     * 🔹 メールアドレス収集（今回は維持）
     */
    form.setCollectEmail(true);

    /**
     * 🔹 説明セクション
     */
    form.addSectionHeaderItem()
      .setTitle('ご回答前にご確認ください')
      .setHelpText(
        '本アンケートは2ページ構成です。\n\n' +
        '【1ページ目】\n' +
        '・メールアドレスの入力\n' +
        '・識別情報（token）の確認\n\n' +
        '【2ページ目】\n' +
        '・アンケート回答\n\n' +
        '※tokenは変更しないでください'
      );

    /**
     * 🔹 token項目
     */
    form.addTextItem()
      .setTitle('token')
      .setRequired(true)
      .setHelpText('※この項目は変更しないでください');

    /**
     * 🔹 ページ分割
     */
    form.addPageBreakItem()
      .setTitle('アンケートにご回答ください');

    /**
     * 🔹 質問追加
     */
    questions.forEach(q => {
    
      let item;
    
      switch (q.type) {
    
        case 'text':
          item = form.addTextItem();
          break;
    
        case 'paragraph':
          item = form.addParagraphTextItem();
          break;
    
        case 'multiple':
          item = form.addMultipleChoiceItem()
            .setChoiceValues(q.options || []);
          break;
    
        case 'checkbox':
          item = form.addCheckboxItem()
            .setChoiceValues(q.options || []);
          break;
    
        case 'list':
          item = form.addListItem()
            .setChoiceValues(q.options || []);
          break;
    
        case 'scale':
          item = form.addScaleItem()
            .setBounds(1, 5);
          break;
    
        case 'date':
          item = form.addDateItem();
          break;
    
        case 'time':
          item = form.addTimeItem();
          break;
    
        default:
          throw new Error(`Unsupported question type: ${q.type}`);
      }
    
      /**
       * 🔹 共通設定
       */
      item
        .setTitle(q.title)
        // .setRequired(q.required === true || q.required === 'TRUE');
        .setRequired(q.required);
    
    });

    /**
     * 🔹 回答先スプレッドシート作成
     */
    const responseSpreadsheet = SpreadsheetApp.create(`${formTitle}_Responses`);

    form.setDestination(
      FormApp.DestinationType.SPREADSHEET,
      responseSpreadsheet.getId()
    );

    /**
     * 🔹 回答シート保護
     */
    FormOpsSheetProtectionService.protectResponseSheet(responseSpreadsheet.getId());

    /**
     * 🔹 FORM_ID保存
     */
    FormOpsConfig.setFormId(form.getId());

    /**
     * 🔹 entryId取得
     */
    this.saveTokenEntryId();

    /**
     * 🔹 Runフォルダへ格納
     */
    this.moveToFolder(form.getId(), runFolder);
    this.moveToFolder(responseSpreadsheet.getId(), runFolder);

    return form;
  }

  /**
   * ファイルを指定フォルダへ移動
   *
   * 設計思想：
   * - Runフォルダ単位で管理
   * - Form / Spreadsheet 共通処理
   *
   * @param {string} fileId
   * @param {Folder} folder
   */
  static moveToFolder(fileId, folder) {
  
    /**
     * 🔥 Drive反映待ち（重要）
     */
    Utilities.sleep(500);
  
    const file = DriveApp.getFileById(fileId);
  
    /**
     * 🔹 フォルダへ移動（推奨API）
     *
     * 設計思想：
     * - Driveの単一親モデルに準拠
     * - add/removeの組み合わせを廃止
     */
    if (!folder) {
      FormOpsLogService.log('WARN', 'Skip move: folder is null');
      return;
    }
    
    FormOpsDriveService.moveFile(file, folder);
  }

  /**
   * tokenフィールドのentryId取得（完全版）
   */
  static getTokenEntryId() {

  const formId = FormOpsConfig.FORM_ID;

  if (!formId) {
    throw new Error('FORM_ID is not set');
  }

  const form = FormApp.openById(formId);

  const textItems = form.getItems(FormApp.ItemType.TEXT);

  const tokenItem = textItems.find(item =>
    item.getTitle().trim().toLowerCase() === 'token'
  );

  if (!tokenItem) {
    throw new Error('Token field not found');
  }

  const response = form.createResponse();

  const itemResponse = tokenItem
    .asTextItem()
    .createResponse('DUMMY_TOKEN');

  response.withItemResponse(itemResponse);

  const url = response.toPrefilledUrl();

  const match = url.match(/(entry\.\d+)=/);

  if (!match) {
    throw new Error('Failed to extract entryId');
  }

  return match[1];
  }

  /**
   * token entryId保存
   */
  static saveTokenEntryId() {
  
    const entryId = this.getTokenEntryId();
  
    FormOpsConfig.setTokenEntryId(entryId);
  
    FormOpsLogService.log(
      'INFO',
      'Token entryId saved',
      entryId
    );
  }

  /**
   * フォームURL取得（viewform保証）
   *
   * 設計思想：
   * - 常に回答用URLを返す
   * - editURLを誤って使う事故を防ぐ
   */
  static getFormUrl() {
  
    const formId = FormOpsConfig.FORM_ID;
  
    if (!formId) {
      throw new Error('FORM_ID is not set');
    }
  
    const form = FormApp.openById(formId);
  
    const editUrl = form.getEditUrl();
  
    /**
     * 🔹 viewformに変換
     */
    return editUrl.replace('/edit', '/viewform');
  }

  /**
   * token付きprefill URL生成（完全自動）
   *
   * 設計思想：
   * - entryIdに依存しない
   * - Googleの内部ロジックを利用
   * - 常に正しいURLを生成
   *
   * @param {string} token
   * @returns {string}
   */
  static generatePrefilledUrl(token) {
  
    const formId = FormOpsConfig.FORM_ID;
  
    if (!formId) {
      throw new Error('FORM_ID is not set');
    }
  
    const form = FormApp.openById(formId);
  
    /**
     * 🔹 token項目取得
     */
    const textItems = form.getItems(FormApp.ItemType.TEXT);
  
    const tokenItem = textItems.find(item =>
      item.getTitle().trim().toLowerCase() === 'token'
    );
  
    if (!tokenItem) {
      throw new Error('Token field not found');
    }
  
    /**
     * 🔹 回答生成
     */
    const response = form.createResponse();
  
    const itemResponse = tokenItem
      .asTextItem()
      .createResponse(token);
  
    response.withItemResponse(itemResponse);
  
    /**
     * 🔹 prefilled URL生成
     */
    return response.toPrefilledUrl();
  }

  /**
   * Run単位フォルダ作成
   *
   * 設計思想：
   * - 実行単位でファイルをグルーピング
   * - Form / Responses / Master を同一Runで管理
   * - 後から追跡しやすくする
   *
   * @param {string} runId
   * @returns {GoogleAppsScript.Drive.Folder}
   */
  static createRunFolder(runId, title = '') {
  
    /**
     * 🔹 Drive利用不可ならスキップ
     */
    if (!FormOpsDriveService.isAvailable()) {
      FormOpsLogService.log('WARN', 'Drive unavailable: skip run folder creation');
      return null;
    }
  
    const rootName = FormOpsConfig.OUTPUT_FOLDER_NAME;
  
    /**
     * 🔹 フォルダ取得 or 作成
     */
    const folders = DriveApp.getFoldersByName(rootName);
  
    let root;
  
    if (folders.hasNext()) {
      root = folders.next();
    } else {
      root = DriveApp.createFolder(rootName);
    }
  
    /**
     * 🔹 タイトル安全化
     */
    const safeTitle = title
      ? title.replace(/[\\\/:*?"<>|]/g, '').replace(/\s+/g, '_').slice(0, 30)
      : '';
  
    /**
     * 🔹 フォルダ名生成
     */
    const folderName = safeTitle
      ? `run_${runId}_${safeTitle}`
      : `run_${runId}`;
  
    return root.createFolder(folderName);
  }

  /**
   * ファイルをルートフォルダへ移動
   *
   * 設計思想：
   * - Masterなどの恒久ファイル用
   * - Runフォルダとは分離
   *
   * @param {string} fileId
   */
  static moveToRootFolder(fileId) {
  
    /**
     * 🔹 Drive利用不可ならスキップ
     */
    if (!FormOpsDriveService.isAvailable()) {
      FormOpsLogService.log('WARN', 'Drive unavailable: skip moveToRootFolder');
      return;
    }
  
    /**
     * 🔹 Drive反映待ち
     */
    Utilities.sleep(500);
  
    const rootName = FormOpsConfig.OUTPUT_FOLDER_NAME;
  
    /**
     * 🔹 フォルダ取得 or 作成
     */
    const folders = DriveApp.getFoldersByName(rootName);
  
    let root;
  
    if (folders.hasNext()) {
      root = folders.next();
    } else {
      root = DriveApp.createFolder(rootName);
    }
  
    /**
     * 🔹 ファイル取得
     */
    const file = DriveApp.getFileById(fileId);
  
    /**
     * 🔹 移動（DriveService経由）
     */
    FormOpsDriveService.moveFile(file, root);
  }
}