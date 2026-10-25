# WorldShot Log v2.5.4 公開準備

## 2026年10月25日: セキュリティ対応版

未公開だったv2.5.4の変更を維持し、CVE-2026-111497への対応を追加しています。下に残す9月11日の配布物・ハッシュは以前の候補の記録であり、今回の公開用ファイルとは異なります。

- Electronを41.0.2から45.0.0へ更新。Chromiumは156.0.8078.48、内蔵Node.jsは24.22.0です。
- ElectronはNodeのBrotli呼び出しにも[Chromium共用ライブラリを使用](https://github.com/electron/electron/blob/v45.0.0/patches/node/build_add_gn_build_files.patch)します。[採用リビジョン](https://chromium.googlesource.com/chromium/src/+/156.0.8078.48/third_party/brotli/README.chromium) `cfa184a997ac19454d5a3ae97f6d600aac2445a6` は、[今回の修正コミットを含む子孫](https://github.com/google/brotli/compare/0ed4268a4401abcb4472096340477cf4af7d1a2b...cfa184a997ac19454d5a3ae97f6d600aac2445a6)です。Node/Brotliの表示バージョンだけで修正有無を判断していません。
- `check:security`で承認したElectron/Chromium/Nodeの組合せと正常なBrotliデータの往復を検証します。実行ファイルのSHA-256と根拠URLも出力します。不正入力による攻撃成立の再現試験ではありません。また、Nodeの同日発表のすべてのCVEを解消したとする検証ではありません。
- ExifReaderを4.52.0、xmldomを0.9.12、protobufjsを7.6.6へ更新。配布対象依存の`npm audit --omit=dev`は警告0件です。Forgeなど開発依存の警告は別に残っており、全依存が警告0件という意味ではありません。
- better-sqlite3を13.0.3へ更新。ElectronのABIに依存しない公式N-APIバイナリを使用し、配布時はWindows x64版だけを同梱します。既存のDB・バックアップ検証を再実行しています。
- 最終の全体回帰チェックは20/20成功。[実行結果](../out/regression-checks/2026-10-25T04-44-18-446Z-45912/result.json)。ExifReaderのUTF-8処理修正に合わせ、原本の日本語と一致する3件の生タグ期待値のみを更新しました。

### 最終ローカル候補と確認結果

公開用候補は [`out/review-candidate-v9AnlX`](../out/review-candidate-v9AnlX/) です。ビルド元のGitコミットは `1a142703ab712b12eef6881a57afcef4c030fcd1`。上記の回帰チェックを入力にした [ビルド記録](../out/review-candidate-v9AnlX/candidate.json) では、次の6工程がすべて成功しています。

| 工程 | 結果 |
| --- | --- |
| Git差分の検査 | 成功 |
| インストーラ生成処理への既存パッチ適用 | 成功 |
| Windowsアプリ・Squirrel配布物の生成 | 成功 |
| 配布版の起動・機能スモークテスト | 成功 |
| 隔離した確認用プロファイルの生成 | 成功 |
| 確認用プロファイルでの画面・動作スモークテスト | 成功 |

候補記録は `ok: true`、`status: ready-for-user-review` です。候補生成時点の記録にある`releaseReady: false`は変更せず、後続の実インストール・公開・配信確認は以下に別途記録します。

### 今回の配布物

配布用の3ファイルは [`build/make/squirrel.windows/x64`](../out/review-candidate-v9AnlX/build/make/squirrel.windows/x64/) にあります。以下は今回の候補記録と照合した値です。再ビルド時は改めてハッシュと検証結果を確定します。

| ファイル | バイト数 | SHA-256 |
| --- | ---: | --- |
| `WorldShotLogSetup.exe` | 235490304 | `3f7235478efc5e7a3b63311b26d2a67f25ef37410d9b1182a4da89947aed6295` |
| `worldshot_log-2.5.4-full.nupkg` | 235079843 | `97add18356784e8bf69585edfbb4f223d4e764756aab8919fbff1d5827a43db4` |
| `RELEASES` | 84 | `3f5002b65c6a7c5a05a26c5116083d3d206c765d49182a8bd27fb2b82ef584a5` |

### 実インストール検証と公開準備

- GitHubのv2.5.4ドラフトへ配布物3点をアップロード済み。GitHub側のasset digestと上記SHA-256の一致を確認しています。
- GitHubの使い捨てWindows環境による[実インストール検証](https://github.com/noma-nomoa/vrchat-world-photo-manager/actions/runs/40163732485)が3経路すべて成功しました。新規導入、v2.5.3からSetupによる上書き導入、旧Squirrelによる更新で、インストール登録・ショートカット・配布物89ファイルの一致を確認しました。
- 実際の旧版で登録した合成写真、お気に入り、日本語メモ、手動World名、ラベル・色、テーマを新版の起動・再起動後に照合しました。各経路で配布版のSQLite・Worker・同梱AI実推論・Undo/Redoも成功し、通常終了しています。
- 2026年10月25日14:00 JSTに[v2.5.4を正式公開](https://github.com/noma-nomoa/vrchat-world-photo-manager/releases/tag/v2.5.4)しました。Latest、draft=false、prerelease=falseを確認。公開タグは`e4071a59129e5439555d801c02ff4eaabe738f5f`です。
- 通常の`update.electronjs.org` URLで、v2.5.3からv2.5.4がHTTP 200で案内され、RELEASESのパッケージ名・サイズ・SHA-1が配布物と一致しました。v2.5.4自身からはHTTP 204です。キャッシュ回避パラメータは使用していません。[配信確認結果](../out/security-release-2026-10-25/public-feed.json)、[公開結果の一覧](../out/security-release-2026-10-25/release-result.json)。

実インストール用CIは、完成済み配布物を実行し、チェックアウト側にはASAR・Fuse検査用のJavaScript依存だけを`npm ci --ignore-scripts`で用意します。最初の実行は準備段階で不要なSQLite再ビルドが走り、ホストのVisual Studio検出に失敗しました。また写真登録の検証コードを、実際のrendererと同じ`Array<File>`を渡す形式に修正しました。成功したCIのコミットは`b9585090ce207d067641eb27e7502dc24f62e9ad`です。ビルド元から変えたのは配布対象外のCI・検証用コードだけで、配布物は再生成していません。この依存導入手順は開発用ビルド環境の構築手順ではありません。

旧版のUpdate.exeをコピーした隔離環境では、v2.5.3からv2.5.4の選択・取得、89ファイルの展開内容一致、展開後のアプリ起動・SQLite・Worker・実AI処理・Undo/Redo・正常終了が成功しました。[パッケージ検証結果](../out/upgrade-harness-dcd4d941/run-izLwY5/result.json)。HTTPでの取得、破損データのSHA不一致拒否、正常な再取得による回復も成功しています。[HTTP検証結果](../out/upgrade-harness-dcd4d941/http-run-675lG3/result.json)。元のインストール済みファイル・ショートカット・アンインストール登録は前後で不変です。この隔離検証は実際のインストール適用とは別の確認です。

利用中のアプリや元DBへのインストール検証は行いません。以下の9月11日の記録は履歴として保存し、今回の配布物に対する検証結果とは区別します。

## 2026年9月11日: 以前の候補の記録

以下は当時の検証記録です。v2.5.4の配布物を生成し、既存機能とDB互換性を確認しました。当時は公開・タグ作成・push・利用中のアプリへのインストールを行わず、実インストーラによる新規導入・上書き更新は未検証でした。

## 今回の候補

- 候補: [`out/review-candidate-1br0cR`](../out/review-candidate-1br0cR/)
- 確認用起動: [`Open-review.cmd`](../out/review-candidate-1br0cR/Open-review.cmd)
- 変更内容: [v2.5.4 リリースノート](../release-notes/v2.5.4.md)
- ビルド記録: [`candidate.json`](../out/review-candidate-1br0cR/candidate.json)
- 最終照合と全証跡の一覧: [`release-preparation-result.json`](../out/review-candidate-1br0cR/release-preparation-result.json)

`Open-review.cmd` は専用のサンプルDBと6枚の合成画像を使います。利用中のアプリやDBを置き換える操作は行いません。候補フォルダーは現在の場所に置いたまま起動してください。ランチャーは実際に Windows PowerShell 5.1 経由で起動し、タイトル `WorldShot Log v2.5.4`、ウィンドウの応答、通常終了を確認しました。

バージョンは `package.json`、`package-lock.json` の2箇所、`src/index.html` に反映済みです。依存関係の更新や新しいデザインの追加は行っていません。ビルド記録の `versionUnchanged: true` は「ビルド中にチェックアウトのバージョンを変更していない」という意味です。公開中の v2.5.3 と同じバージョンという意味ではありません。

## 確認結果

| 対象 | 結果・確認内容 |
| --- | --- |
| 全体回帰チェック | 19/19 合格。構文、単体、翻訳、一覧、レタッチ、AI、保存、オーバーレイ、画像出力、座標、ブラシ、プリセット、DB、バックアップ、World情報、更新処理 |
| 検証対象とビルドの一致 | 118個の入力ファイルをハッシュで照合。検証・ビルド中のソース変更なし |
| Windows配布物 | Electron Forge / Squirrel の生成成功。アプリ89ファイルとSquirrel補助2ファイルの整合性を確認 |
| 配布版の起動 | ASAR、SQLiteのネイティブモジュール、DB IPC、画像Worker、同梱U2NetPの実推論・Undo/Redoが合格 |
| 確認用画面 | ギャラリー・レタッチの描画と通常終了が合格。合成元画像6枚のハッシュ不変 |
| 実ランチャー | `Open-review.cmd` → Windows PowerShell 5.1 → 配布EXEの起動・応答・通常終了が合格 |
| 現行DB | コピーした6,614枚、11フォルダー、607件のWorldキャッシュ、5ラベルを読み込み。既存列・全行の値とシーケンスを保持 |
| 旧名称のDB | コピーした436枚、1フォルダー、46件のWorldキャッシュ、1ラベルを読み込み。既存データを保持 |
| DB接続・再起動 | 両DBで2回ずつ起動。実IPC経由の写真・年月・フォルダー・ラベル・集計がSQLと一致。整合性検査合格 |
| 元DBの保全 | DB/WAL/SHMの検証前後のハッシュが一致。移行・照合は専用コピー上で実行 |

結果ファイル:

- [全体回帰チェック](../out/regression-checks/2026-09-11T03-49-20-469Z-56928/result.json)
- [配布版・画面のビルド検証](../out/review-candidate-1br0cR/candidate.json)
- [実DB互換性](../out/release-2.5.4-check-okkqjp/result.json)
- [実DBの接続・再起動](../out/release-2.5.4-check-okkqjp/ipc-result.json)
- [実ランチャー](../out/release-2.5.4-check-okkqjp/actual-launcher-result.json)

v2.5.4 のビルド内ソースと全体チェック対象は一致しています。レタッチ・オーバーレイの追加操作や旧版との画像比較については、同じ機能実装に対する [9月10日の詳細記録](./release-candidate-review-2026-09-10.md) も参照してください。

## 旧版で保存したデータ・設定の引き継ぎ

[連続性テストの結果](../out/release-upgrade-data-GYh0sS/result.json) は合格です。公開版 v2.5.3 の変更していない ASAR を、すべてのアプリ保存先を隔離した Electron ハーネスで起動しました。旧版の main / preload / renderer から合成画像6枚の情報を保存して通常終了し、同じプロファイルで今回の v2.5.4 の配布EXEを2回起動・通常終了しました。

2回とも次の情報が一致しました。

- 写真のID、お気に入り、メモ、印刷用メモ、手入力のWorld情報
- ラベル・色・写真との対応、登録フォルダー
- 背景画像の設定JSONと画面への適用
- テーマ、フォント、カード密度、自作レタッチプリセット、最近使った文字フォントなど localStorage の7キー
- 再利用するオーバーレイ素材と、既存写真を開いたレタッチ画面

DBの既存列・全行の値・シーケンス、整合性検査、外部キー検査、バックアップ構造の検証も合格しています。合成元画像のハッシュは不変です。これはアプリのデータ・設定の継承検証であり、旧版の実EXEやインストーラを使った上書きインストール検証ではありません。

## 旧更新プログラムでの取得・展開

公開版 v2.5.3 と同じ旧 `Update.exe` を専用ディレクトリーへコピーし、新しいパッケージの検出・取得を確認しました。ローカルディレクトリーと localhost HTTP の両経路で、現在版 `2.5.3`・更新先 `2.5.4` が一致し、取得したNUPKGのSHA-256も候補と一致しました。

旧Squirrelの実際の展開処理で91ファイルを展開し、すべてを元NUPKGとハッシュ照合しました。展開した新EXEでもDB接続、Workerの18,432チャンネル比較、実U2NetP推論、Undo/Redo、正常終了が合格しています。

破損したHTTP応答を使う追加検証では、旧更新プログラムがチェックサム不一致で取得を失敗させ、正常な応答で再取得すると復旧することを確認しました。ただし公開済みSquirrelには、失敗時に読み取りハンドルを閉じる前に不正キャッシュを削除しようとし、IOExceptionでキャッシュが残る挙動があります。今回の修正で追加された不具合ではなく、正常再取得で置換されました。不正キャッシュの削除まで成功したとは扱いません。

利用中の旧Update.exe・NUPKG・RELEASES、WorldShotのショートカット、HKCUのアンインストール登録値は、検証前後で一致しました。

- [検出・取得・展開・新EXE起動](../out/upgrade-harness-dcd4d941/run-gwJdgY/result.json)
- [HTTP取得・破損検知・正常再取得](../out/upgrade-harness-dcd4d941/http-run-XtyCo2/result.json)
- [実行方法・Squirrel調査元・未検証範囲](../out/upgrade-harness-dcd4d941/README.md)

## 配布物の識別

配布用の3ファイルは [`build/make/squirrel.windows/x64`](../out/review-candidate-1br0cR/build/make/squirrel.windows/x64/) にあります。今回検証した配布物を識別する SHA-256 は次のとおりです。再ビルドした場合は新しいハッシュと検証結果が必要です。

| ファイル | バイト数 | SHA-256 |
| --- | ---: | --- |
| `WorldShotLogSetup.exe` | 219760640 | `db4b1b6fb323b7cf2f4ed4ba85e61cb5760a1faef0a0f859030b69e404e0d6d6` |
| `worldshot_log-2.5.4-full.nupkg` | 219409076 | `889cbefbb283ab1e964157f4a6c481ddfd757870b4272ee65a3c6222a90ee533` |
| `RELEASES` | 84 | `7375523872978232b7b6ae62a43709123b4f9c806519dc9bb72b22bdb5274318` |

公開中の基準版は [v2.5.3](https://github.com/noma-nomoa/vrchat-world-photo-manager/releases/tag/v2.5.3) です。旧NUPKGのSHA-256 `f80dea2d2b9c38a0b11bee7fcb2ae51129902d7f1de3b140bfffa7bc01e426da` が公式Releaseのasset digestと一致しています。

## 公開前に残る確認

このPCは Windows 11 Home で、Windows Sandbox や既存の仮想Windows環境は利用できません。Squirrelの実インストール・更新はユーザーのアンインストール登録やショートカットも変更します。環境変数による保存先変更だけでは、利用中のWindowsから隔離したインストール検証になりません。

専用Windows環境で次を確認する必要があります。

1. 公開版 v2.5.3 を導入し、写真登録・お気に入り・メモ・ラベル・World情報・設定を保存して通常終了する。
2. 今回の v2.5.4 のインストーラで上書き導入し、既存ショートカットから v2.5.4 が起動することを確認する。再起動後も登録情報と設定を保持することを照合する。
3. 別のクリーンな状態では v2.5.4 の新規導入を確認する。アンインストール登録・終了・再起動と、バックアップからのデータ復元も確認する。
4. 自動更新では v2.5.3 の実インストール状態から今回のパッケージを適用し、終了・再起動後に v2.5.4 に切り替わることを確認する。

実パッケージの検出・取得・展開や、同一プロファイルでのデータ引き継ぎは、Windowsへのインストール登録・旧版の削除・ショートカット更新・更新後の自動再起動を保証するものではありません。

上記の確認後、リリース対象をコミットとして確定し、[リリース手順](../RELEASE.md) に従ってタグ・配布物・Release本文を揃えます。現在は未コミットの変更を含む候補であり、`candidate.json` の Git HEAD だけでは内容を再現できません。ソースの識別には保存済みの118ファイルのハッシュを使っています。

公開後には `update.electronjs.org` が v2.5.3 に対して v2.5.4 を返し、v2.5.4 自身には更新なしを返すことも必要です。未公開の現段階では、この配信確認は実施できません。

## 検証の範囲

通常使用の元写真や元DBを変更せず、ローカルの専用ディレクトリーで確認しています。すべての利用者の写真・GPU・外部AIモデルでの動作を保証するものではありません。画面テストは通信を遮断しているため、従来から外部取得するアイコンフォントは画像内で名前の文字列として表示されます。今回の画面テストを、通信可能時のアイコン表示の検証とは扱いません。
