**バックアップ復元とWorld情報の保全（2026-09-09）**

リリース済みv2.5.3を起点に、既存データを失う不具合2件を修正した。作業ブランチは `codex/retouch-session-safety`。元画像・外観・アプリバージョン・依存関係は変更していない。

**バックアップの復元**

従来は `{}` や欠けたテーブルを空のバックアップとして扱い、登録データが削除される場合があった。現在は、復元ファイルのスキーマ、全5テーブル、行の型・必須列・主キー・一意制約・ラベル参照を検証してから復元する。検証は自動バックアップ作成・ローテーションとDB削除より前に行い、DB APIを直接呼ぶ場合も省略できない。

- 正式な旧バックアップ、裸のsnapshot、`{ data: snapshot }` の互換性を維持する。
- 正しく作られた全テーブル空のバックアップは、意図した空状態として復元できる。
- 旧版で省略可能なnullable列・既定値のある列、未知の追加列を許容する。
- schemaVersionは1。外側のappName・backupSchemaVersionは、存在する場合に照合する。
- countsは補助情報として無視し、復元件数は実際の行から求める。
- 写真パスの実在、日時やWorld IDの書式、設定内部のキーは今回の必須条件に加えていない。
- 行の挿入は既存のトランザクション内で行い、途中のSQLエラーでも削除前へロールバックする。

**再取り込み時のWorld情報**

従来は同じ写真を明示的に再取り込みすると、手動で訂正したWorld ID・URLが画像からの抽出値やnullで上書きされた。写真に `world_identity_source` を1列追加し、ID・URL・対応する自動名称を保全する。

| 由来 | 再取り込み時の動作 |
| --- | --- |
| automatic | 今後の新規抽出。従来どおり抽出値で更新する |
| manual | 手動でリンクを変更、または明示的なWorld情報再読込で選んだ情報。既存値を保持する |
| legacy | 旧DBと旧バックアップで由来が不明。IDまたはURLがあれば保持し、両方空の場合は新しい抽出値で補完する |

手動名称だけの変更・解除は、実際の画面で使う設定保存APIでもリンクの由来を変更しない。既存の「手動設定を解除」やURL空欄の意味も維持する。公式情報の同期は、取り込み前の抽出値ではなく、DBに実際に保存されたWorldを対象にする。レタッチ保存時も、編集元メタデータを反映した最終行から同期対象を組み立てる。

由来は一覧データ→レタッチの編集元情報→保存IPC→編集済み写真にも引き継ぐ。バックアップにも含め、由来列がない旧バックアップはlegacyとして復元する。

**移行上の制約**

過去のDBにはID・URLが手入力か自動取得かの記録がない。そのため、旧データの自動取得済みリンクも保護対象になる。旧リンクを再取り込みだけで自動訂正する動作は抑制されるが、手動編集・明示的な再読込は使える。すでに上書きされて失われた手動値を、この修正だけで復元することはできない。

**検証**

```powershell
npm.cmd run smoke:data-safety
npm.cmd run smoke:retouch
```

| コマンド | 確認内容 |
| --- | --- |
| smoke:data | 既存の旧DB移行・通常バックアップ互換性 |
| smoke:backup-validation | メモリDBで150ケース。正常復元・不正入力拒否と全5テーブルの全行不変、事前検証だけでも拒否できること |
| smoke:backup-restore | 実Electron main/preload経由で復元・エクスポート・取消を23シナリオ。不正復元時の全行・採番状態・設定bytes・自動バックアップ不変と、正常復元直前の自動バックアップ内容 |
| smoke:world-identity | 一時DBで24ケース。単発・バッチ取り込み、旧DB移行、旧・新バックアップ、画面の設定保存APIでの名前追加・解除、明示的再読込、編集元情報継承 |
| smoke:export | manualとlegacyの各5PNGで全画素を比較。main→renderer→編集元情報→保存DBのWorld保全、各保存直後の同期対象、月再取り込み、実同期キューと更新通知を確認 |
| smoke:retouch | 上記保存テスト2モードに加えて、マスク競合、座標、ブラシ、プリセットの既存回帰確認 |

不具合修正前はバックアップ検証の正常8件が通り、不正ケースでデータの変化を再現した。World情報は正常2件が通り、手動リンクの上書きを含む17件が失敗した。修正後は上記のDB検証が通過している。

アプリを起動するテストはhome・AppData・一時フォルダを隔離し、合成画像を使用する。バックアップの実IPCテストと保存テストは外部通信を遮断する。テスト起動はElectron CLIを使い、終了コードまで待つ。PowerShellからGUI実行ファイルを直接起動するとログのpipeが閉じた後もプロセスが残ることがあり、初回の復元テストでEPIPEダイアログが発生した。該当する隔離テストプロセスは終了済み。検証スクリプト内のstdout/stderrでのEPIPEだけログを停止する対策を追加し、真のテスト例外はレポートと非0終了で扱う。製品コードに例外無視は追加していない。

実ユーザーデータでの試行、配布パッケージの再生成、インストール・アップデート、公開は未実施。リリース前に配布版での確認が必要。

**最終確認結果**

- `smoke:data-safety` は終了コード0。バックアップ150件、実復元23件、World情報24件と既存データ互換性が通過した。
- マスク競合54件、保存2モード、座標、ブラシ、プリセットを最終コードで確認した。一括実行中に座標テストの描画待ちが止まったため、その隔離プロセスを終了。テストだけ背景描画制限を解除し、renderer操作に30秒の上限を設けた後、座標・ブラシ・プリセットを再実行して通過した。製品の描画設定は変更していない。
- 保存2モード各5PNGの画素ハッシュと元画像ハッシュは、修正前のbaselineと一致した。
- 翻訳監査457候補・漏れ0、変更JavaScriptの構文確認、`git diff --check` が通過した。テスト終了後にElectronプロセスの残留がないことも確認した。
- 開発版の既存CSP警告と終了時のGPUログは出る場合がある。EPIPEダイアログはテスト起動方法とテスト側ログ処理の修正後には発生していない。

実行結果（Git管理外）：

- [復元23シナリオ](C:/Users/tyusk/vrchat-world-photo-manager/out/smoke-backup-restore/worldshot-backup-restore-GTUTcP/result.json)
- [World情報24ケース](C:/Users/tyusk/AppData/Local/Temp/worldshot-world-identity-VAxAaZ/result.json)
- [保存・再取り込みmanual](C:/Users/tyusk/vrchat-world-photo-manager/out/smoke-export/worldshot-export-KD2r0G/result.json) ／ [legacy](C:/Users/tyusk/vrchat-world-photo-manager/out/smoke-export/worldshot-export-tPh8ct/result.json)
- [マスク競合54ケース](C:/Users/tyusk/vrchat-world-photo-manager/out/smoke-subject-session/worldshot-subject-session-NQjrb6/result.json)
- [座標](C:/Users/tyusk/vrchat-world-photo-manager/out/smoke-coordinate-space/20260909-214057/result.json) ／ [ブラシ](C:/Users/tyusk/vrchat-world-photo-manager/out/smoke-mask-brush/20260909-214136/result.json) ／ [プリセット](C:/Users/tyusk/vrchat-world-photo-manager/out/smoke-presets/20260909-214141/result.json)
