/**
 * 質問タイプ定義
 *
 * 設計思想：
 * - FormTemplates と FormService の整合性を担保
 * - type文字列の散在を防ぐ
 */
const FormOpsQuestionTypes = {

  LIST: [
    'text',        // 短文
    'paragraph',   // 長文
    'multiple',    // ラジオボタン
    'checkbox',    // チェックボックス
    'list',        // プルダウン
    'scale',       // スケール
    'date',        // 日付
    'time'         // 時刻
  ]

};