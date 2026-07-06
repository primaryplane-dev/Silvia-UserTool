# コンベア一覧表 運用チェックリスト

## 接続切替時
- [ ] `src_コンベア一覧表/Bas_Public.bas` の `P_ConnectString` が想定接続先になっている
- [ ] `src_コンベア一覧表/Bas_List.bas` で `P_ConnectString` を使用している
- [ ] `src_コンベア一覧表/Bas_DbConnection.bas` で `P_ConnectString` を使用している

## 出力確認（org比較）
- [ ] コンベアNoの行数が org と一致（例: 1〜10）
- [ ] 時間列の並びが一致（例: 8:22 / 12:25 / 16:20）
- [ ] 各行の `○/×/－` の入り方が一致
- [ ] ヘッダ担当者名が一致（`SYKJ`基準）
- [ ] `製造終了` 表示が一致

## 条件仕様（維持ポイント）
- [ ] SBFP01読込は直SQL方式（org互換）
- [ ] 隠しキー `BC` 列に `BFKTCD` を保持（`fncFindRow` が参照）
- [ ] 実績取得は `BGHINO = Val(P_HINO)` を使用
- [ ] `BGKJNO` では絞り込まない（org互換）
- [ ] `BGFNCD='1'` を明細から除外しない（org互換）

## 問題発生時の切り分け
- [ ] A5:SQL で `SBGP01` の `BGSDAT` / `BGHINO` 件数確認
- [ ] `BGKTCD` 分布（1〜20 / 21以上）を確認
- [ ] `SBFP01` の `BFKTCD/BFCVNO` と、Listシート `BC/A列` の対応を確認
