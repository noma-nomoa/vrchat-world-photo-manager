**継続改善の記録（2026-09-10）**

前日の構造改善を土台に、実際に再現したレタッチの競合とWebP登録の不具合を修正し、検証を継続実行できるハーネスを整えた。開始点のバージョンは2.5.3のまま。既存の画像計算・デザインを維持し、依存関係の更新や公開は含めない。

その後、ユーザーの継続指示を受けて [一覧・取り込みの構造整理](./list-import-refactor-2026-09-10.md) まで実施した。現在の部品と依存関係は [コード構成](./architecture.md) を参照。以下はその前段の安全性・ハーネス整備時点の記録。

**画像素材・プレビューの非同期競合**

画像素材の追加ダイアログ／コピー処理を待っている間にAからBへ移ると、Aの結果がBのレイヤー・Undo履歴・ステータスへ入ることを実rendererで再現した。修正前は最初の正常系4件が通過したあと、Bの状態不変の検証で失敗した。

- 素材の追加は開始した編集状態を保持し、完了時にその編集が有効な場合だけレイヤーへ追加する。完了済みの素材一覧は全体へ反映する。同じ編集内で独立に始めた複数の追加は維持する。
- 管理素材の削除確認を待つ間に別写真へ移った場合は、古い確認で削除IPCを始めない。既に始めた削除の完了は一覧へ反映するが、新しい編集のレイヤーやステータスへ干渉しない。
- 別写真が使っている読み込み済み素材キャッシュは保持し、編集中の描画を維持する。通常の削除完了では素材一覧DOMも更新する。
- 古い写真のWorkerプレビューが、新しい写真の元画像読み込み前に届くケースも編集の所有者を確認する。無効な結果のbitmapは閉じ、新しい写真のCanvasへ描かない。

実renderer＋制御した非同期境界で18シナリオが通過。素材の実ファイルコピー・削除ダイアログそのものはこのテストでは置換している。

[修正前の失敗記録](C:/Users/tyusk/vrchat-world-photo-manager/out/smoke-overlay-session/worldshot-overlay-session-BLaPFq/result.json) ／ [修正後の18シナリオ](C:/Users/tyusk/vrchat-world-photo-manager/out/smoke-overlay-session/worldshot-overlay-session-5iLaiM/result.json)

**JPEG/WebP保存とWebP寸法の修正**

既存のPNG5種類の画素検証を保ったまま、JPEG/WebPの品質60・100を、実Worker経路とマスクありの通常描画経路の両方で検証した。計13保存をmanual/legacyの2モードで行う。

検証中、WebPファイルの保存に成功してもDBへ登録する幅・高さがnullになる問題を検出した。WindowsのnativeImageで直接デコードできない場合、既に読み込んでいるExifReaderのWebPヘッダーにある正整数の幅・高さを採用する。nativeImageで取得できる寸法を優先し、ヘッダー情報もない場合は既存のファイル名解析へ戻る。新しい依存関係やDB移行は追加していない。

全26保存で形式の署名、実ファイルのデコード、寸法、全alpha、品質の反映、サムネイル、撮影日時・メモ・World情報、再取込・同期を確認した。JPEG/WebPは不可逆圧縮のため、二色の合成画像に対する平均RGB誤差4/255以下、境界を除く平坦部の最大誤差8/255以下を基準にした。PNGと原本は従来の厳密一致を維持する。

既にDBで寸法が欠落しているWebPは、対象月の再取込で回復することを確認した。一覧を開くだけで既存行を一括補正する処理は追加していない。WebPサムネイルはこのWindows環境の既存Shell経由生成で成功しており、すべてのWindows環境のデコーダー構成を確認したものではない。

**繰り返し実行するためのハーネス**

`npm.cmd run check` で構文・単体・画像演算・翻訳・レタッチ・データ保全の15チェックを順番に実行する。各チェックのログはファイルへ直接接続し、先日の閉じたパイプへのログ出力を避ける。タイムアウト・中断では検証用プロセスツリーだけを終了し、失敗箇所と途中経過をJSONへ保存する。JSON更新は同じディレクトリの一時ファイルから置換し、途中結果の破損を避ける。使い方は [保守用ハーネス](C:/Users/tyusk/vrchat-world-photo-manager/docs/maintenance-harness.md) にまとめた。

`smoke:retouch` と `smoke:data-safety` の既存コマンドもこのハーネス経由にした。ハーネス自体の8テストを加え、単体は合計79件。通常の製品実行へ包括的な例外抑制は追加していない。

**一括検証結果**

[15チェックの実行記録](C:/Users/tyusk/vrchat-world-photo-manager/out/regression-checks/2026-09-09T20-10-59-980Z-66780/result.json) はすべて成功、終了コード0。

- 単体79件、画像演算2,185ケース、翻訳457候補・漏れ0。
- マスク競合55、保存競合14、素材・プレビュー競合18、Worker6、保存26、座標、ブラシ、プリセット。
- 旧DB互換、バックアップ検証150、実復元23、World情報24。

保存の最終記録は [manual](C:/Users/tyusk/vrchat-world-photo-manager/out/smoke-export/worldshot-export-u2cJJo/result.json) ／ [legacy](C:/Users/tyusk/vrchat-world-photo-manager/out/smoke-export/worldshot-export-56uzh4/result.json)。全補正・全レイヤーの組合せ、実写真の切り抜き品質、別GPU/OSの結果を保証するものではない。

PNG5種類×2モードは、先行修正前のbaselineと全画素hashが一致。元画像SHA-256も一致した。プリセット4種も初回baselineと画素hash・画像統計が一致した。

**配布用パッケージと同梱AI**

Forgeの既存設定で、過去の生成物と別の `out/package-validation-20260910-PqwhkR/` へWindowsパッケージを生成した。アプリのバージョンは変更していない検証用ビルドであり、配布・インストール・公開は行っていない。

パッケージ用の検証スクリプトを拡充し、[最終実行記録](C:/Users/tyusk/vrchat-world-photo-manager/out/smoke-packaged/worldshot-packaged-smoke-06zEoK/result.json) で以下を確認した。

- ASAR内の実行時10ファイルが作業中のソースとSHA-256で一致し、テスト・fixture・docs・outが除外されている。
- SQLiteのnative moduleがASAR外へ展開され、実preload/IPCを通した隔離DBの件数取得が成功。Forgeのリビルド前後で開発環境のSQLiteバイナリーhashは不変。
- EXEの6種類の設定済みFuse値が期待どおり。ASAR整合性検証等の本番設定を外さずに検証した。
- ASARから起動した実Workerの6パターンで、各3,072チャンネルが通常の共通画像計算と一致。
- 同梱U2NetPとONNX WASMをローカル資源から読み込み、実際の推論で32×24の合成入力に対するPNGマスク生成とUndo/Redoが成功。モデルや推論結果の置換なし。約1.5秒という時間はこの小さな合成入力での実行記録で、実写真の応答時間や品質を示すものではない。
- AppDataだけでなくhomeと標準フォルダも一時profileへ向け、サムネイル移行を隔離。外部通信要求の記録は0件。
- アプリは終了コード0で正常終了し、強制終了なし・cleanup errorなし。異常終了や強制終了を成功として報告しないようにした。

検証側で終了要求だけを意図的に抑止した試験でも、[期待した失敗記録](C:/Users/tyusk/vrchat-world-photo-manager/out/smoke-packaged/worldshot-packaged-smoke-OeCYvJ/result.json) が `ok: false`・終了コード1となり、対象アプリのプロセスツリーだけを終了できた。製品コード・EXEは変更していない。これはハーネスの異常系確認であり、上記の通常起動が失敗したという意味ではない。検証終了後にアプリのプロセスが残っていないことも確認した。

配布版の再確認には、生成済みEXEの場所を明示できる。

```powershell
npm.cmd run smoke:packaged -- "--exe=C:/Users/tyusk/vrchat-world-photo-manager/out/package-validation-20260910-PqwhkR/WorldShot Log-win32-x64/WorldShotLog.exe"
```

**残る確認範囲**

今回、実モデルの実行経路は確認できたが、実際の人物・髪・衣装の切り抜き品質、追加ダウンロード型モデル、全編集機能の組合せは別の検証範囲。画像素材が読み込めないときの保存結果や、画像取り込み・一覧表示に残る大きな処理も、再現ケースと測定を用意してから個別に扱う。インストーラの新規導入・既存版からの更新・更新配信は未実施。デザイン追加は後段に維持する。
