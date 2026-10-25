**一覧・取り込みの構造整理（2026-09-10）**

巨大なmain/rendererの整理を続け、一覧の検索計算、メタデータ解析、写真レコード生成、取り込みの進行手順を独立した部品へ移した。既存のレタッチ、保存形式、DBスキーマ、デザインはこの整理で変更していない。

**一覧・検索**

`photo-list-query.js` に12関数を集約し、rendererは現在の検索条件を渡して結果を表示する。検索対象の切替、ラベルAND/OR、お気に入り、向き、日時とIDの並び順、ラベル一覧の重複除外・件数をこの部品が担当する。DOM、選択状態、画面遷移はrendererに残す。

切り出し前の期待値をfixtureへ固定し、1,920通りの検索・絞り込み・並び順、972組の比較、24種の日付×3タイムゾーン等を15の単体テストで確認した。テスト実行時に現行rendererから期待値を作り直さない。配列・写真・検索条件を変更しないことも検証している。

検索文字列の正規化を写真ごとではなく1回の検索につき1回にした。検索する写真情報は毎回読み、更新済みメモ等が古いキャッシュで検索されないようにしている。2万枚の合成データで新旧を交互に計測した参考値は49.06ms→46.82ms。差は小さく、アプリ全体の体感速度の改善や一定の短縮率を保証するものではない。すべての結果IDと順序は一致した。

実画面でも9枚・3か月・2年の合成DBを使い、39シナリオを切り出し前に通した。表示中のカード順・件数・空表示、検索入力の確定とIME、検索範囲、フィルター、選択維持と選択解除、月／年の切替を確認する。

[実画面の切り出し前](C:/Users/tyusk/vrchat-world-photo-manager/out/smoke-photo-list/worldshot-photo-list-VPW7CE/result.json) ／ [切り出し後](C:/Users/tyusk/vrchat-world-photo-manager/out/smoke-photo-list/worldshot-photo-list-57boGb/result.json)。元の合成画像9枚のSHAも不変。

**メタデータとレコード生成**

`photo-metadata.js` はタグ・XMP・World情報・日時・寸法分類・文字列補正を扱う。移動した41関数の構文木が切り出し前と一致することを独立レビューで確認した。

`photo-record.js` はファイル読込・stat・native decode・サムネイル生成・時計を明示的な依存として受け取る。mainには実際のfs/nativeImage等への配線を残す。日時と寸法の取得元の優先順位、hash、レコードの全列、サムネイル開始と完了待ちの順序を維持している。

実PNG/JPEG/WebP、XMP付きPNG/JPEG、Exif日時付きJPEGを小さな合成fixtureとして追加した。旧実装で240の解析ケースと17のレコード生成ケースの期待値を固定し、入力の不変と失敗内容も比較。非同期の完了待ち・statとの並行実行を含む261の単体テストが通過した。

範囲外XML文字参照など、従来例外になる入力も既存挙動として記録している。これらを新しい正常仕様と断言したり、分離に合わせて黙って許容したりはしていない。

**取り込み手順と同時実行**

`photo-import-service.js` に、選択パスの重複除外→前処理→DB登録→不足サムネイルの補完→登録後World情報の取得→結果集計を移した。mainの `importManyFiles` は依存の配線と呼び出しだけになった。共有の結果形式もこのサービスにまとめ、追跡フォルダ更新・明示的再取込でも利用する。

`concurrency.js` は、取り込み・表示前処理・健全性確認が共有する既存の同時実行数制御を保持する。単に全Promiseを同時開始する形には変えていない。

21の取り込み契約テストを旧mainの関数で先に通し、そのまま新サービスでも通した。部分失敗、処理完了順と入力順、同時実行上限、一括DB登録失敗時の再試行、個別登録失敗の伝播、サムネイル補完、移動写真と重複hashの件数、登録後の手動World情報を確認する。さらに共有の同時実行処理を7テストで確認した。

通常の失敗順序や既存の再試行の意味を保つことを優先した変更であり、取り込み全体を新たに1トランザクションへ変更したものではない。

現在の全体構成は [コード構成](./architecture.md)、一括検証は [保守用ハーネス](./maintenance-harness.md) にまとめている。

**最終検証**

[一括16チェック](C:/Users/tyusk/vrchat-world-photo-manager/out/regression-checks/2026-09-10T03-45-57-805Z-4992/result.json) はすべて終了コード0。

- 全JavaScript54ファイルの構文検査、単体383件、既存色補正2,185ケース、翻訳監査の漏れ0を確認。
- 一覧39シナリオと、既存のマスク55・保存競合14・素材競合18・Worker6・画像保存26・座標・ブラシ・プリセットを確認。
- 旧DB互換、バックアップ検証150・実復元23・World情報24が通過。
- 最終保存の [manual](C:/Users/tyusk/vrchat-world-photo-manager/out/smoke-export/worldshot-export-dd90Vv/result.json) ／ [legacy](C:/Users/tyusk/vrchat-world-photo-manager/out/smoke-export/worldshot-export-FQSJLz/result.json) のPNG各5種類は、先行修正前baselineと全画素hashが一致。原本SHAも不変。[プリセット4種](C:/Users/tyusk/vrchat-world-photo-manager/out/smoke-presets/20260910-124629/result.json) も同じbaselineの画素hash・統計と一致。

通常のForge設定で別の隔離出力先へ再ビルドし、[実EXEの検証](C:/Users/tyusk/vrchat-world-photo-manager/out/smoke-packaged/worldshot-packaged-smoke-ggIrZU/result.json) も成功した。新5モジュールを含む15ファイルのASAR／ソースSHA一致、6種類の設定済みFuse、テスト資材の除外、SQLite IPC、実Worker6パターンの全画素一致、同梱U2NetP＋WASMによるマスク生成・Undo/Redoを確認。正常終了コード0、強制終了なし、残留プロセスなし。[ビルド記録](C:/Users/tyusk/vrchat-world-photo-manager/out/package-structure-20260910-OmH94T/build-result.json) では開発環境のSQLiteバイナリーhashも不変。

独立レビューでは、共有の同時実行処理とメタデータ41関数が旧実装と構文木で一致。レコード生成はI/O・時計の受け渡し、取り込み手順は関数名・並列数の受け渡し・未使用変数削除以外の変更がないことを確認した。静的な未宣言参照の候補も到達経路を追い、既存の到達不能な旧更新処理を今回の修正に混ぜていない。

公開、インストール、アプリのバージョン更新は行っていない。別GPU/OS、実写真のAI切り抜き品質、インストーラ更新経路、巨大ライブラリ全体の応答時間はこの検証の保証範囲外。全体の整理は完了しておらず、画面操作・Canvas、サムネイル保管、World情報取得・モデル管理等が次のまとまりとして残る。
