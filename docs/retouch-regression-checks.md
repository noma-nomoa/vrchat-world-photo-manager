**レタッチの回帰確認とマスク処理の安全性修正（2026-09-09）**

既存ユーザーの機能・出力を維持するため、最初の改修を「古い被写体マスク処理の結果が、別写真や新しい操作に反映される問題」に絞った。作業ブランチは `codex/retouch-session-safety`、開始点はリリース済みv2.5.3の `3b7a73e`。

追記：その後の状態・画像処理・保存境界の分離と性能改善は [構造改善の記録](C:/Users/tyusk/vrchat-world-photo-manager/docs/architecture-improvements.md) を参照。本書の「今回の結果」は最初の不具合修正時点の記録。

最新の確認範囲・候補・未実施項目は [リリース前の確認](./release-candidate-review-2026-09-10.md) を参照。以下の過去の結果・限界は、その時点の記録として残している。

**変更した動作**

- AI開始時の編集状態・入力画像・要求識別子を保持し、成功・失敗時に現在の要求であることを確認する。
- 編集画面の閉じるアニメーションが始まった時点で、結果の反映を無効にする。閉じた後、別写真を開いた後、新しい要求を始めた後にも古い結果を反映しない。
- マスク削除、全リセット、Undo／Redo、マスクファイルの読み込みで置き換えられた要求も無効にする。
- 読み込み中のFileReader／Imageコールバックにも同じ判定を使い、遅れて届くマスクやエラーを別写真へ反映しない。
- Undo／Redoで過去の「AI処理中」を復元しても、存在しない処理を待ち続ける状態に戻さない。マスクありならready、なしならnoneとして扱う。

結果の反映だけを止める修正であり、実行中のONNX計算を強制停止したり、共有推論セッションを解放したりしない。通常の明るさ変更や切り抜き操作では、進行中のAIをキャンセルしない。

本体の変更箇所は [src/renderer.js:9752](C:/Users/tyusk/vrchat-world-photo-manager/src/renderer.js:9752) と関連するマスク・履歴・閉じる処理。画像計算、モデル定義、DB、保存形式、アプリバージョン、デザインの変更は含めていない。

**実行方法**

Windowsで依存関係がインストールされた開発環境から実行する。

```powershell
npm.cmd run smoke:retouch
```

現在は以下の8本を順番に実行し、失敗時はそこで止まる。2026-09-10からログ・タイムアウト管理を [共通ハーネス](C:/Users/tyusk/vrchat-world-photo-manager/docs/maintenance-harness.md) にまとめている。

| コマンド | 確認する内容 |
| --- | --- |
| `npm.cmd run smoke:subject-session` | 実Electronに読み込んだrendererで、AI成功／失敗、画面切替、閉じる直後、要求の完了順序、リセット、履歴、マスク取り込み、古い履歴タイマーを55ケース確認 |
| `npm.cmd run smoke:save-session` | 保存開始時情報の固定、描画／IPC／一覧更新待ち中の写真切替、古い画像読込を14ケース確認 |
| `npm.cmd run smoke:overlay-session` | 素材追加・削除確認・削除IPC・プレビュー待ち中の写真切替、同一写真の再オープンを18ケース確認 |
| `npm.cmd run smoke:worker` | 実Workerのbitmap/PNG bufferを要求し、RGB/HSV/併用の6ケースで通常rendererとの全画素一致を確認 |
| `npm.cmd run smoke:resource-readiness` | 必須の画像素材・マスクの読込失敗、待機中の編集・Undo／Redo・切替、Worker待機中の失効、正常PNG出力の維持 |
| `npm.cmd run smoke:export` | PNG5種類に加え、JPEG/WebP各品質60・100×Worker/通常描画の計13保存を確認。manual／legacyの2モードで全画素または明示した圧縮誤差、寸法・サムネイル・メタデータ・原本不変・再取込・World同期を確認 |
| `npm.cmd run smoke:coordinates -- --skip-screenshots` | 切り抜き位置に対するテキスト・画像・ぼかしの追従、画像オーバーレイの拡縮 |
| `npm.cmd run smoke:mask-brush -- --skip-screenshots` | ブラシ追加・削除、濃度取得、Undo／Redo、ズーム・スクロール・輪郭表示 |
| `npm.cmd run smoke:presets` | エアリーホワイト4種の設定、色調・コントラストの関係、日英韓の表示 |

構文検査・翻訳監査・データ互換性の確認も実施する。変更したJavaScriptはファイルごとに `node --check` を実行する。

```powershell
npm.cmd run audit:i18n -- --strict
npm.cmd run smoke:data
git diff --check
```

**検証の再現性と隔離**

- プリセット用の入力は [scripts/fixtures/retouch-preset-fixture.js](C:/Users/tyusk/vrchat-world-photo-manager/scripts/fixtures/retouch-preset-fixture.js) で決定的に生成する。管理対象外の `img/notuse.png` への依存を除去した。既存のプリセット期待値・アサーションは維持している。
- 新規テストは合成画像と合成マスクを使う。保存先は一時ディレクトリで、元画像のSHA-256が変わらないことを確認する。
- 既存の編集スモークとUIキャプチャでも、AppDataに加えてhome・documents・picturesとアプリ固有のルート環境変数を隔離する。サムネイル移行が実ユーザーのhomeを参照しないようにした。
- 新規の競合テスト・保存テストは外部通信なしで動作する。既存のプリセット・UI系テストは従来どおり外部フォントを利用する場合がある。
- 競合テストは推論Promiseと一部のマスク読込を制御して、結果が届く順序を再現する。実モデルのダウンロード・推論品質は検証しない。計算結果の生成以降の制御は、実際に読み込んだrendererを通す。
- 保存テストはネイティブの保存先選択だけを置換し、rendererの書き出し処理→preload→mainの保存IPC→実ファイル保存・再登録を通す。

**今回の結果**

| 検証 | 結果 |
| --- | --- |
| 修正前の通常AI成功・失敗後の再試行 | 3モデル分岐で成功 |
| 修正前の閉じる直後のAI完了 | 期待する状態不変のアサーションが失敗し、不具合を再現 |
| 修正後の競合・通常操作 | 54ケース成功。失敗・履歴復元後にAIボタンが再び有効になることも確認 |
| レタッチ一括実行 | 5本すべて成功 |
| PNG保存結果の修正前後比較 | 5種類すべて画素ハッシュ一致。元画像ハッシュも一致 |
| プリセットの修正前後比較 | 同じ合成入力に対する4種類すべての画素ハッシュ一致 |
| 構文検査 | src・scripts・fixture・forge設定の計20ファイル成功 |
| 翻訳監査 | 候補457件、漏れ0件 |
| データ互換性 | 旧DB移行・正常バックアップ復元のスモーク成功 |
| 独立したコードレビュー | マスク要求の失効境界・成功／失敗／後始末・正常動作への影響を確認 |

今回の実行結果は以下にある。これらはローカルの生成物で、Git管理対象外。

- [競合テスト54ケース](C:/Users/tyusk/vrchat-world-photo-manager/out/smoke-subject-session/worldshot-subject-session-EdbiFP/result.json)
- [保存結果](C:/Users/tyusk/vrchat-world-photo-manager/out/smoke-export/worldshot-export-65XxgL/result.json)
- [プリセット結果](C:/Users/tyusk/vrchat-world-photo-manager/out/smoke-presets/20260909-211623/result.json)

**確認範囲の限界と次の作業**

今回の一致は、固定した入力・操作・同一Windows環境についての結果。全補正・全プリセット・全レイヤーの組合せ、JPEG／WebP、実モデルによる推論、GPUやフォントが異なる環境、保存ボタンのUI状態遷移、Workerとフォールバックの強制比較はまだ含めていない。関連箇所を改修するときに、その正常系を先に検証へ追加する。

配布パッケージの再生成・インストーラ確認・公開・自動更新配信は行っていない。リリース前には [RELEASE.md](C:/Users/tyusk/vrchat-world-photo-manager/RELEASE.md) に従った配布版の確認が必要。

追記：不正バックアップJSONの拒否と、明示的な再取り込み時の手動World情報保全も実装した。追加のDB列・編集元情報の継承と検証については [データ保全の回帰確認](C:/Users/tyusk/vrchat-world-photo-manager/docs/data-safety-regression-checks.md) を参照。その後も、測定と動作維持の確認を伴って最適化を進め、デザイン追加は最後に行う。
