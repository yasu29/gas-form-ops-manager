# 📋 FormOps（GAS Form Automation Tool）

Google Apps Script（GAS）を利用して
**Googleフォームの配布・回収・リマインドを自動化するツール**です。

アンケート運用における以下の課題を解決します：

* 手動でのフォーム作成
* 個別URL配布の手間
* 未回答者の管理
* リマインド送信の工数

本ツールは
**実務利用を前提とした設計（安全性・保守性・拡張性）**で構築されています。

---

# ✨ Features

* スプレッドシート定義からフォーム自動生成
* tokenによる回答者識別（ユニークURL発行）
* prefilled URLによる個別回答リンク生成
* 初回メール送信（Gmail）
* 未回答者へのリマインド送信
* 回答時ステータス自動更新（Pending → Submitted）
* ログ管理（Logsシート）
* Driveフォルダ自動作成（Run単位管理）
* 回答シートの自動保護（編集防止）
* CC / BCC をControlシートで管理
* 実行環境の自動判定（Google Workspace制限対応）

---

# 🏗 Architecture

責務ごとにレイヤー分離した構成です。

```
Spreadsheet (Config / Data)
        │
        ▼
Repository Layer
        │
        ▼
Service Layer
        │
        ▼
Engine
        │
        ▼
Starter (Entry Point)
```

---

## レイヤー構成

```
Service
 ├── FormService
 ├── MailService
 ├── ParticipantService
 ├── LogService
 ├── SheetProtectionService
 └── DriveService（環境差分吸収）

Repository
 ├── SheetRepository
 ├── MailTemplateRepository
 ├── ParticipantRepository
 └── ControlRepository

Engine
 └── TemplateEngine

Config
 └── ScriptProperties（環境状態含む）
```

---

# 🔄 System Flow

## フォーム作成〜配布フロー

1. FormTemplatesシートから質問定義を取得
2. フォーム生成（1プロジェクト1フォーム）
3. 回答用スプレッドシート作成
4. DriveにRunフォルダ作成・格納（※環境によりスキップ）
5. token項目を自動生成
6. Script PropertiesにID保存

---

## 初回メール送信フロー

1. Participants初期化（token再生成）
2. 個別URL生成（prefilled URL）
3. メールテンプレート取得
4. `${name}` / `${url}` を置換
5. Gmail送信
6. Logsに記録

---

## リマインド送信フロー

1. 未回答者（Pending）抽出
2. URL生成
3. メール送信
4. Logsに記録

---

## 回答時処理（Trigger）

1. token取得
2. 参加者特定
3. ステータス更新（Submitted）
4. Logsに記録

---

# 📁 Project Structure

```
src/

  formops_config.gs
  formops_constants_questionTypes.gs

  formops_engine_templateEngine.gs

  formops_repository_*.gs
  formops_service_*.gs

  formops_setup.gs
  formops_starter.gs
  formops_trigger_onSubmit.gs
  formops_util_token.gs

  gas_util_envCheck.gs
```

---

## 各ファイルの役割

| File           | Responsibility      |
| -------------- | ------------------- |
| formops_config | Script Properties管理 |
| templateEngine | フォーム定義構築            |
| repository系    | シートデータ取得            |
| service系       | 業務ロジック              |
| DriveService   | Drive制限吸収           |
| setup          | 初期セットアップ            |
| starter        | 実行入口                |
| trigger        | 回答時処理               |
| util_token     | token / runId生成     |
| envCheck       | 環境チェック              |

---

# 🚀 Usage

## 1. GASプロジェクト作成

Google Apps Script プロジェクトを作成し、
`src` ディレクトリのファイルを配置します。

---

## 2. 初期セットアップ

```
formops_setupAll()
```

実行内容：

* 環境チェック（自動実行）
* Masterスプレッドシート作成
* 各シート初期化
* サンプルデータ投入
* Script Properties 自動設定

  * FORMOPS_SPREADSHEET_ID
  * FORMOPS_ENV（環境状態）

※ Script Properties は事前設定不要です

---

## 3. フォーム作成

```
formops_createForm()
```

実行内容：

* フォーム生成
* 回答用スプレッドシート生成
* Script Properties 自動設定

  * FORMOPS_FORM_ID
  * FORMOPS_TOKEN_ENTRY_ID

---

## 4. 初回メール送信

```
formops_sendInitial()
```

---

## 5. リマインド送信

```
formops_sendReminder()
```

---

# 📂 Google Drive Structure

## ✅ 環境①：制限なし（個人アカウント等）

```
Google Drive
└ FormOps_Output
    ├ run_YYYYMMDD_HHMMSS_アンケート名
    │   ├ Form
    │   └ Responses Spreadsheet
    └ FormOps_Master
```

---

## ⚠️ 環境②：制限あり（企業Google Workspace）

```
Google Drive（ルート直下）

・フォーム
・回答スプレッドシート
・Masterスプレッドシート
```

---

### 補足

* Drive操作は自動的にスキップされます
* フォルダ整理されないだけで、機能自体は正常動作します

---

# 🧪 Example Behavior

* 各参加者にユニークURLを配布
* 回答時に自動で「Submitted」に更新
* 未回答者のみリマインド送信
* 全操作がLogsシートに記録

---

# 🔧 Design Philosophy

本プロジェクトでは以下を重視しています。

---

## 責務分離（Separation of Concerns）

```
データ取得
業務ロジック
テンプレート処理
外部連携
```

を明確に分離

---

## 副作用の制御

* 検証（Validation）後にリソース生成
* Run単位でファイルを管理
* 不要なファイル生成を防止

---

## 安全性（実務前提）

* tokenによる識別
* メール送信前のバリデーション
* 回答シート保護
* ログによるトレーサビリティ確保

---

## 環境適応設計（今回の強化ポイント）

* 実行環境を自動判定
* Drive制限を透過的に吸収
* 1つのコードで全環境対応

---

## シンプル設計

* 1プロジェクト1フォーム
* 過度な抽象化を排除
* 保守性を優先

---

# ⚠️ Environment Limitations

本ツールは Google Apps Script および Google Drive API を利用しています。

利用環境（特に企業の Google Workspace 環境）によっては、
以下の操作が制限される場合があります：

* フォルダ作成（DriveApp.createFolder）
* ファイル移動（file.moveTo）
* ルートフォルダの取得（DriveApp.getRootFolder）
* ファイル削除（file.setTrashed）

---

## 本ツールの対応

これらの制限がある場合でも：

* Drive操作は自動スキップ
* フォーム / メール / ログ機能は正常動作

---

# 📄 License

MIT License
