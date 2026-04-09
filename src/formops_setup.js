/**
 * FormOps 初期セットアップ
 *
 * 概要：
 * - FormOpsで使用する全スプレッドシートを初期化
 * - シート作成 / ヘッダー設定 / 書式 / 凡例まで一括構築
 *
 * 設計方針：
 * - シンプル構造（過剰抽象化しない）
 * - UIで入力箇所を明確化
 * - 再実行可能
 */
function formops_setup() {

  /**
   * 🔹 Masterスプレッドシート準備
   *
   * 設計思想：
   * - Masterは単一リソース
   * - Run単位では管理しない
   */
  const ss = formops_prepareMasterSpreadsheet();

  /**
   * 🔹 初期化対象を明示（重要）
   */
  FormOpsSheetRepository.setSpreadsheet(ss);

  setupTemplatesSheet();
  setupParticipantsSheet();
  setupMailTemplatesSheet();
  setupControlSheet();
  setupLogsSheet();
  setupGuideSheet();

  sortSheets();

}


/**
 * フォーマット共通定義
 */
const FormOpsFormatConfig = {
  USER:   { bg: '#eaf7ea', font: '#000000', label: 'ユーザー入力' },
  SYSTEM: { bg: '#f1f3f4', font: '#5f6368', label: 'システム生成' },
  FIXED:  { bg: '#fff4e5', font: '#b06000', label: '固定値' },

  LEGEND: {
    HEADER_BG: '#d9d9d9',
    CARD_BG: '#fafafa'
  }
};


/**
 * FormTemplates
 */
function setupTemplatesSheet() {

  const sheet = FormOpsSheetRepository.getOrCreateSheet('FormTemplates');
  const headers = ['order','type','title','required','options'];

  sheet.clear();
  sheet.getRange(1,1,1,headers.length).setValues([headers]);

  formatHeader(sheet, headers.length);

  /**
   * 🔹 サンプルデータ投入（全タイプ）
   */
  sheet.getRange(2,1,8,5).setValues([
    [1, 'text', 'お名前をご入力ください', 'TRUE', ''],
    [2, 'paragraph', 'ご意見・ご感想をお聞かせください', 'FALSE', ''],
    [3, 'multiple', 'サービスの満足度を教えてください', 'TRUE', '非常に満足,満足,普通,不満,非常に不満'],
    [4, 'checkbox', '利用したサービスを選択してください（複数選択可）', 'FALSE', '商品A,商品B,商品C'],
    [5, 'list', 'お住まいの地域を選択してください', 'TRUE', '関東,関西,中部,九州'],
    [6, 'scale', '総合評価をお願いします', 'TRUE', ''],
    [7, 'date', 'ご利用日を教えてください', 'FALSE', ''],
    [8, 'time', 'ご利用時間帯を教えてください', 'FALSE', '']
  ]);

  /**
   * 🔹 type列：選択式（B列）
   */
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(FormOpsQuestionTypes.LIST)
    .setAllowInvalid(false)
    .build();
  
  sheet.getRange('B2:B100').setDataValidation(typeRule);
  
  /**
   * 🔹 required列：選択式（D列）
   */
  const requiredRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE','FALSE'])
    .setAllowInvalid(false)
    .build();
  
  sheet.getRange('D2:D100').setDataValidation(requiredRule);

  /**
   * 🔹 書式
   */
  formatSheetByColumns(sheet, {
    user: [1,2,3,4,5]
  });
}


/**
 * Participants
 */
function setupParticipantsSheet() {

  const sheet = FormOpsSheetRepository.getOrCreateSheet('Participants');
  const headers = ['name','email','token','status'];

  sheet.clear();
  sheet.getRange(1,1,1,headers.length).setValues([headers]);

  formatHeader(sheet, headers.length);

  /**
   * 🔹 サンプルデータ投入（name / email のみ）
   *
   * 設計思想：
   * - token / status はシステムが管理
   * - ユーザー入力部分のみサンプル提供
   */
  sheet.getRange(2,1,4,2).setValues([
    ['山田 太郎', 'taro@example.com'],
    ['佐藤 花子', 'hanako@example.com'],
    ['鈴木 一郎', 'ichiro@example.com'],
    ['田中 次郎', 'jiro@example.com']
  ]);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending','Submitted'])
    .build();

  sheet.getRange('D2:D100').setDataValidation(rule);

  formatSheetByColumns(sheet, {
    user: [1,2],
    system: [3,4]
  });
}


/**
 * MailTemplates
 */
function setupMailTemplatesSheet() {

  const sheet = FormOpsSheetRepository.getOrCreateSheet('MailTemplates');
  const headers = ['templateKey','subject','body'];

  sheet.clear();

  sheet.getRange(1,1,1,headers.length).setValues([headers]);

  formatHeader(sheet, headers.length);

  if (sheet.getLastRow() <= 1) {
  
    /**
     * 🔹 実用テンプレート投入
     */
    sheet.getRange(2,1,2,3).setValues([
      [
        'initial',
        '【ご回答のお願い】アンケートのご案内',
        '${name} 様\n\n' +
        'お世話になっております。\n\n' +
        '以下のアンケートへのご回答をお願いいたします。\n\n' +
        '▼回答はこちら\n' +
        '${url}\n\n' +
        '※所要時間：2〜3分程度\n\n' +
        'お忙しいところ恐縮ですが、ご協力のほどよろしくお願いいたします。\n\n' +
        '---\n' +
        '本メールはシステムより自動送信されています。'
      ],
      [
        'reminder',
        '【再送】アンケートご回答のお願い',
        '${name} 様\n\n' +
        'お世話になっております。\n\n' +
        '先日ご案内いたしましたアンケートについて、\n' +
        '未回答の方へ再度ご連絡させていただいております。\n\n' +
        '▼回答はこちら\n' +
        '${url}\n\n' +
        '※すでにご回答いただいている場合は、\n' +
        '　本メールと行き違いとなりますこと、何卒ご容赦ください。\n\n' +
        'お手数をおかけいたしますが、\n' +
        'ご回答いただけますと幸いです。\n\n' +
        '---\n' +
        '本メールはシステムより自動送信されています。'
      ]
    ]);  
  }

  formatSheetByColumns(sheet, {
    fixed: [1],
    user: [2,3]
  });

  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 400);

  sheet.getRange(2,2,sheet.getMaxRows()).setWrap(true);
  sheet.getRange(2,3,sheet.getMaxRows()).setWrap(true);
}


/**
 * Control
 */
function setupControlSheet() {

  const sheet = FormOpsSheetRepository.getOrCreateSheet('Control');

  const headers = ['key','value'];

  const rows = [
    ['form_title','顧客満足度アンケート'],
    ['mail_cc','cc1@example.com'],
    ['mail_bcc','bcc1@example.com,bcc2@example.com']
  ];

  sheet.clear();
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  sheet.getRange(2,1,rows.length,headers.length).setValues(rows);

  formatHeader(sheet, headers.length);

  formatSheetByColumns(sheet, {
    fixed: [1],
    user: [2]
  });
}


/**
 * Logs シート初期化
 *
 * 設計思想：
 * - ログ専用シート
 * - appendRow前提のため装飾は最小限
 * - 凡例は持たない
 */
function setupLogsSheet() {

  const sheet = FormOpsSheetRepository.getOrCreateSheet('Logs');

  /**
   * 🔹 完全初期化（値＋書式）
   */
  sheet.clear();

  /**
   * 🔹 ヘッダー定義
   */
  const headers = ['timestamp', 'level', 'message', 'data'];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  /**
   * 🔹 ヘッダー書式
   */
  formatHeader(sheet, headers.length);

  /**
   * 🔹 列幅（シンプル）
   */
  sheet.setColumnWidth(1, 160); // timestamp
  sheet.setColumnWidth(2, 80);  // level
  sheet.setColumnWidth(3, 300); // message
  sheet.setColumnWidth(4, 300); // data

  /**
   * 🔹 固定行
   */
  sheet.setFrozenRows(1);
}


/**
 * Guideシート作成
 *
 * 設計思想：
 * - 説明・凡例・運用手順を集約
 * - 各シートは入力に専念させる
 * - 他シートと同一インターフェースに統一
 */
function setupGuideSheet() {

  const sheet = FormOpsSheetRepository.getOrCreateSheet('Guide');

  /**
   * 🔹 完全初期化
   */
  sheet.clear();

  /**
   * 🔹 内容定義
   */
  const values = [
  ['FormOps 利用ガイド'],
  [''],

  ['■ 運用手順（実行関数と対応）'],
  ['① formops_setupAll()：初期セットアップ（初回のみ実行）'],
  ['　　└ formops_starter.gs'],
  ['② formops_createForm()：フォーム作成'],
  ['　　└ formops_starter.gs'],
  ['③ formops_sendInitial()：初回メール送信'],
  ['　　└ formops_starter.gs'],
  ['④ formops_sendReminder()：未回答者へリマインド送信'],
  ['　　└ formops_starter.gs'],
  [''],

  ['■ 凡例（セルの色の意味）'],
  ['ユーザー入力項目：手動で入力・編集してください'],
  ['システム管理項目：自動更新されます（編集しないでください）'],
  ['固定項目：システムで使用する値（変更しないでください）'],
  [''],

  ['■ FormTemplates'],
  ['order：表示順'],
  ['type：質問形式'],
  ['required：必須回答（TRUE / FALSE）'],
  ['options：選択肢（カンマ区切り）'],
  [''],

  ['type一覧'],
  ['text：短文'],
  ['paragraph：長文'],
  ['multiple：ラジオボタン'],
  ['checkbox：チェックボックス'],
  ['list：プルダウン'],
  ['scale：評価スケール'],
  ['date：日付'],
  ['time：時刻'],
  [''],

  ['■ Participants'],
  ['name：氏名'],
  ['email：メールアドレス'],
  ['※token / status は自動管理'],
  [''],

  ['■ MailTemplates'],
  ['subject：メール件名'],
  ['body：メール本文'],
  [''],
  ['使用可能な変数'],
  ['${name}：宛名（Participants.name）'],
  ['${url}：回答用URL'],
  [''],
  ['※これらはメール送信時にスクリプトによって自動で置換されます'],
  ['※本文を編集する場合も削除せずに残してください'],
  [''],

  ['■ Control'],
  ['form_title：フォームタイトル'],
  ['mail_cc：CC（カンマ区切り）'],
  ['mail_bcc：BCC（カンマ区切り）']
  ];

  sheet.getRange(1, 1, values.length, 1).setValues(values);

  /**
   * 🔹 凡例の色付け（安全・非停止）
   */
  const legendIndex = values.findIndex(r =>
    r[0] && r[0].includes('凡例')
  );
  
  if (legendIndex !== -1) {
  
    const legendStartRow = legendIndex + 2;
  
    sheet.getRange(legendStartRow, 1)
      .setBackground(FormOpsFormatConfig.USER.bg);
  
    sheet.getRange(legendStartRow + 1, 1)
      .setBackground(FormOpsFormatConfig.SYSTEM.bg);
  
    sheet.getRange(legendStartRow + 2, 1)
      .setBackground(FormOpsFormatConfig.FIXED.bg);

  } else {
  
    /**
     * 🔸 警告ログのみ（処理は継続）
     */
    FormOpsLogService.log(
      'WARN',
      'Guide: legend section not found'
    );
  }

  /**
   * 🔹 見出し装飾
   */
  sheet.getRange(1, 1)
    .setFontSize(14)
    .setFontWeight('bold');

  /**
   * 🔹 列幅
   */
  sheet.setColumnWidth(1, 520);

  /**
   * 🔹 折り返し
   */
  sheet.getRange(1,1,values.length,1).setWrap(true);
}


/**
 * ヘッダー書式
 */
function formatHeader(sheet, colCount) {

  const range = sheet.getRange(1,1,1,colCount);

  range.setFontWeight('bold');
  range.setBackground('#444444');
  range.setFontColor('#ffffff');

  sheet.setFrozenRows(1);
}


/**
 * 列フォーマット
 */
function formatSheetByColumns(sheet, config) {

  const maxRows = sheet.getMaxRows();

  (config.user || []).forEach(col => {
    sheet.getRange(2,col,maxRows)
      .setBackground(FormOpsFormatConfig.USER.bg)
      .setFontColor(FormOpsFormatConfig.USER.font);
  });

  (config.system || []).forEach(col => {
    sheet.getRange(2,col,maxRows)
      .setBackground(FormOpsFormatConfig.SYSTEM.bg)
      .setFontColor(FormOpsFormatConfig.SYSTEM.font);
  });

  (config.fixed || []).forEach(col => {
    sheet.getRange(2,col,maxRows)
      .setBackground(FormOpsFormatConfig.FIXED.bg)
      .setFontColor(FormOpsFormatConfig.FIXED.font);
  });
}


/**
 * シート並び替え
 */
function sortSheets() {

  const ss = FormOpsSheetRepository.getSpreadsheet();

  const order = [
    'Guide',
    'Participants',
    'FormTemplates',
    'MailTemplates',
    'Control',
    'Logs'
  ];

  order.forEach((name, i) => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(i + 1);
    }
  });
}


/**
 * Masterスプレッドシート準備
 *
 * 設計思想：
 * - 既存があれば再利用
 * - なければ新規作成
 * - ScriptPropertyと同期
 */
function formops_prepareMasterSpreadsheet() {

  const props = PropertiesService.getScriptProperties();

  const ss = SpreadsheetApp.create('FormOps_Master');

  /**
   * 🔥 追加（超重要）
   */
  Utilities.sleep(1000);

  /**
   * 🔹 Repositoryにセット
   */
  FormOpsSheetRepository.setSpreadsheet(ss);

  const ssId = ss.getId();

  props.setProperty('FORMOPS_SPREADSHEET_ID', ssId);

  /**
   * 🔹 ルートフォルダへ移動
   *
   * 設計思想：
   * - Masterは常設ファイル
   * - run配下には置かない
   */
  FormOpsFormService.moveToRootFolder(ssId);

  return ss;
}