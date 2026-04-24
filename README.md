# 行政書士 案件・売上・経費管理アプリ

GitHub Pages でそのまま配信できる、`HTML / CSS / JavaScript` のみで構成されたシングルページアプリです。
データ保存・認証は Supabase を利用しています。

## 現在の公開構成（GitHub Pages 対応）

- `index.html`
- `app.js`
- `styles.css`

上記 3 ファイルはすべて**相対パス**で参照しています。

- `index.html` -> `styles.css`
- `index.html` -> `app.js`

そのため、リポジトリ直下を GitHub Pages の公開ルートにすれば追加ビルドなしで動作します。

## 重要な前提（`file://` ではなく `https://`）

- 本アプリは Supabase Auth を使うため、**`https://` での配信前提**です。
- `file://` 直開きではログインセッション保持や認証コールバックが不安定になるため、アプリ側でも `file://` 起動時は案内を表示して停止します。

## GitHub Pages 公開手順

### 1. GitHub に push

このリポジトリを GitHub に push します。

### 2. Pages を有効化

1. GitHub リポジトリを開く
2. `Settings` -> `Pages`
3. `Build and deployment` で以下を設定
   - `Source`: `Deploy from a branch`
   - `Branch`: `main`（または公開したいブランチ）
   - `Folder`: `/ (root)`
4. `Save`

数分後、以下のような URL で公開されます。

- `https://<GitHubユーザー名>.github.io/<リポジトリ名>/`

### 3. Supabase Auth の URL 設定（必須）

Supabase ダッシュボードで、GitHub Pages の URL を許可してください。

1. Supabase プロジェクトを開く
2. `Authentication` -> `URL Configuration`
3. 以下を設定
   - `Site URL`: `https://<GitHubユーザー名>.github.io/<リポジトリ名>/`
   - `Redirect URLs`: 上記 URL を追加

> 末尾スラッシュ付きで登録しておくと安全です。

## Supabase 接続情報について

`app.js` 内の Supabase URL / publishable key は既存値をそのまま使用しています。

- URL: `https://ueelzyftlbnvjvpsmpyt.supabase.co`
- Key: `sb_publishable_...`

## 固定費自動計上機能のための追加SQL

固定費の重複計上防止は `expenses.fixed_expense_id` を利用します。未追加の場合は以下を実行してください。

```sql
alter table expenses add column if not exists fixed_expense_id uuid;
```

## 複数端末で同じデータを閲覧する条件

同じ公開 URL（GitHub Pages）を複数端末で開き、**同じ Supabase アカウントでログイン**すれば、同じ Supabase テーブル上のデータを参照するため同一データが表示されます。

- 本アプリは `user_id` でデータを保存・取得しています。
- 端末ローカル保存ではなく Supabase 永続データを読み込むため、端末間で同期されます。

## ローカル確認方法（`https://` 推奨）

簡易サーバーを使って確認してください（`file://` ではなく HTTP/HTTPS 経由）。

例:

```bash
python -m http.server 8000
```

その後 `http://localhost:8000` を開いて動作確認します。


## 案件メモ・次回対応管理のための追加SQL

`cases` テーブルに以下カラムが未追加の場合は実行してください。

```sql
alter table cases add column if not exists work_memo text;
alter table cases add column if not exists next_action_date date;
alter table cases add column if not exists next_action text;
```
