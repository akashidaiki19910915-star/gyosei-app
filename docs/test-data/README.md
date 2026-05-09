# 手動テスト用サンプルデータ（PR #213 併用）

このディレクトリは、PR #213 で追加された `docs/manual-test-scenarios.md` の手動シナリオを**実機で最短確認**するためのテストデータです。  
**実装本体は変更せず**、CSV/JSON と確認手順のみを提供します。

## 1. ファイル一覧と用途

- `clients-sample.csv`  
  顧客CSV取込確認用。余計な列（`extra_column`）を含みます。
- `cases-sample.csv`  
  案件CSV取込・案件進行監査確認用。
  - 期限超過
  - 期限接近
  - 次回対応日超過
  - 書類未回収あり
  - 未完了タスクあり
  - 請求未作成
  - 売上未登録
  - 問題なし
- `sales-sample.csv`  
  売上CSV取込・入金確認用。
  - 未入金
  - 一部入金
  - 入金済み
  - カンマ入り金額
  - 空欄を含む行
- `expenses-sample.csv`  
  経費CSV取込確認用。
  - 交通費
  - 郵送費
  - 証紙代
  - 0円
  - 空欄
  - 余計な列（`extra_column`）
- `import-extra-columns-sample.csv`  
  DBに存在しない列を複数含む取込確認用。
  `url_notification` / `internal_temp_flag` / `selected_label` / `dummy_column` を含みます。
- `backup-restore-sample.json`  
  バックアップ復元確認用サンプル。
  `clients` / `cases` / `sales` / `expenses` / `case_documents` / `case_tasks` を含み、未定義カラムも混在させています。

## 2. 取込手順（推奨順）

1. 顧客取込画面で `clients-sample.csv` を取込
2. 案件取込画面で `cases-sample.csv` を取込
3. 売上取込画面で `sales-sample.csv` を取込
4. 経費取込画面で `expenses-sample.csv` を取込
5. 余計な列確認として `import-extra-columns-sample.csv` を取込
6. バックアップ復元画面で `backup-restore-sample.json` を復元

## 3. 期待結果

- 余計な列や未定義カラムを含んでも、取り込み可能な項目は正常に取り込まれること
- 未定義カラムは復元/取込時に除外（無視）されること
- 案件監査表示で、各ステータス（期限超過、期限接近、次回対応日超過、書類未回収、未完了タスク、請求未作成、売上未登録、問題なし）が確認できること
- 売上データで未入金・一部入金・入金済みの表示差分が確認できること

## 4. 注意点

- すべてのデータ名に「テスト」「サンプル」を含めています。本番データとは混同しないでください。
- IDは固定文字列（`*_test_sample_*`）を使用していますが、環境内データと重複しないかは取込前に確認してください。
- 取込時に既存データへ上書き・重複登録されないよう、検証環境で実施してください。

## 5. 禁止事項と運用

- **本番環境では使用しないでください。**
- 検証後に不要なテストデータは、画面操作または管理手順に沿って削除してください。
- `docs/manual-test-scenarios.md`（PR #213）と必ず併用してください。
