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
alter table expenses add column if not exists payee text;
alter table expenses add column if not exists payment_method text;
```

## 売上（請求）入金管理強化のための追加SQL

`sales` テーブルに未入金・一部入金・入金済を管理するカラムを追加します。

```sql
alter table sales add column if not exists paid_amount bigint default 0;
alter table sales add column if not exists payment_status text;
alter table sales add column if not exists due_date date;
alter table sales add column if not exists invoice_number text;
alter table sales add column if not exists estimate_id uuid;
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
alter table cases add column if not exists estimate_id uuid;
alter table cases add column if not exists client_id uuid;
alter table cases add column if not exists template_id uuid;
alter table cases add column if not exists required_documents text;
alter table cases add column if not exists task_list text;
alter table cases add column if not exists document_url text;
alter table cases add column if not exists invoice_url text;
alter table cases add column if not exists receipt_url text;
alter table cases add column if not exists work_memo text;
alter table cases add column if not exists next_action_date date;
alter table cases add column if not exists next_action text;
```

## 日報機能のための追加SQL

```sql
create table if not exists daily_reports (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null,
  report_date date not null,
  case_id uuid,
  client_id uuid,
  interaction_type text,
  work_content text not null,
  work_minutes int default 0,
  next_action text,
  next_action_date date,
  memo text,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

alter table daily_reports add column if not exists client_id uuid;
alter table daily_reports add column if not exists interaction_type text;
alter table daily_reports add column if not exists next_action text;
alter table daily_reports add column if not exists next_action_date date;
alter table daily_reports add column if not exists memo text;

alter table daily_reports enable row level security;

create policy "daily_reports_own" on daily_reports
for all
using (auth.uid() = user_id)
with check (auth.uid() = user_id);
```

日報の insert / update payload は以下のキーで統一しています（DBに未作成のカラムは先に上記SQLで追加してください）。

```js
{
  user_id,
  report_date,
  client_id,
  case_id,
  interaction_type,
  work_content,
  work_minutes,
  next_action,
  next_action_date,
  memo
}
```

## 見積機能（見積作成・管理・請求データ出力）の追加SQL

```sql
create table if not exists estimates (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null,
  customer_name text not null,
  estimate_title text not null,
  estimate_date date not null,
  valid_until date,
  status text default '作成中',
  memo text,
  subtotal bigint default 0,
  tax bigint default 0,
  total bigint default 0,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

create table if not exists estimate_items (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null,
  estimate_id uuid not null,
  item_name text not null,
  quantity numeric default 1,
  unit_price bigint default 0,
  amount bigint default 0,
  sort_order int default 0,
  created_at timestamptz default now()
);

alter table estimates enable row level security;
alter table estimate_items enable row level security;

create policy "estimates_own" on estimates
for all
using (auth.uid() = user_id)
with check (auth.uid() = user_id);

create policy "estimate_items_own" on estimate_items
for all
using (auth.uid() = user_id)
with check (auth.uid() = user_id);
```

## 顧客台帳・採番・証憑URL対応の追加SQL

```sql
create table if not exists clients (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null,
  name text not null,
  client_type text,
  address text,
  tel text,
  email text,
  referral_source text,
  memo text,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

alter table clients enable row level security;

create policy "clients_own" on clients
for all
using (auth.uid() = user_id)
with check (auth.uid() = user_id);

alter table cases add column if not exists client_id uuid;
alter table estimates add column if not exists client_id uuid;

alter table estimates add column if not exists estimate_number text;
alter table sales add column if not exists invoice_number text;

alter table cases add column if not exists document_url text;
alter table cases add column if not exists invoice_url text;
alter table cases add column if not exists receipt_url text;
alter table expenses add column if not exists receipt_url text;
```

## 業務テンプレート機能の追加SQL

```sql
create table if not exists work_templates (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null,
  name text not null,
  default_due_days int,
  required_documents text,
  default_tasks text,
  memo text,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

alter table work_templates enable row level security;

create policy "work_templates_own" on work_templates
for all
using (auth.uid() = user_id)
with check (auth.uid() = user_id);

alter table cases add column if not exists template_id uuid;
alter table cases add column if not exists required_documents text;
alter table cases add column if not exists task_list text;
```
