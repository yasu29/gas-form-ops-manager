# 📋 FormOps（GAS Form Automation Tool）

Google Apps Script（GAS）を利用して
**Googleフォームの配布・回収・リマインドを自動化するツール**です。

---

# ✨ Features

* フォーム自動生成
* token付きURL配布
* 初回メール / リマインド送信
* ステータス自動更新
* ログ管理
* 回答シート保護
* 環境自動判定（Workspace制限対応）

---

# 🏗 Architecture

```
Repository → Service → Engine → Starter
```

---

# 🚀 Usage

## 1. 初期セットアップ

```
formops_setupAll()
```

実行内容：

* Masterスプレッドシート作成
* 各シート初期化
* 環境チェック（自動）

---

## 2. フォーム作成

```
formops_createForm()
```

---

## 3. 初回送信

```
formops_sendInitial()
```

---

## 4. リマインド

```
formops_sendReminder()
```

---

# 🌐 Environment Behavior（重要）

本ツールは実行環境を自動判定し、
Google Workspaceの制限に対応します。

---

## ✅ パターン①：制限なし（個人アカウント等）

### 動作

* Drive操作：すべて有効

### 出力構成

```
Google Drive
└ FormOps_Output
    ├ run_YYYYMMDD_HHMMSS_タイトル
    │   ├ Form
    │   └ Responses Spreadsheet
    └ FormOps_Master
```

---

## ⚠️ パターン②：制限あり（企業Workspace）

### 制限例

* フォルダ作成不可
* ファイル移動不可
* ルート取得不可

---

### 動作

* Drive操作：自動スキップ
* エラーにはならず処理継続

---

### 出力構成（重要）

```
Google Drive（ルート直下）

・フォーム
・回答スプレッドシート
・Masterスプレッドシート
```

---

### ポイント

* フォルダ整理されないだけで動作は完全に成立
* ログ・メール・フォーム機能はすべて利用可能

---

# 🔧 Environment Check

内部的に以下をチェック：

* Drive.createFolder
* Drive.moveTo
* Drive.getRootFolder
* Drive.setTrashed
* Gmail.getAliases

結果は ScriptProperties に保存され、
Service層で利用されます。

---

# 💡 Design Philosophy

* 環境差分はコードで吸収
* 1コードで全環境対応
* 安全性重視（Validation + Logging）
* 過剰な分岐を排除

---

# ⚠️ 注意事項

* Workspace環境ではDrive機能が制限される場合があります
* その場合、ファイルは自動整理されません

---

# 📄 License

MIT License
