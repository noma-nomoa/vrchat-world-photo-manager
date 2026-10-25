**WorldShot Log — 仕様・コード調査（2026-09-09）**

現行機能には、データ互換性を守る処理や画像編集の性能対策が既にある。次の開発では、確認できたデータ保全・非同期処理の不具合を修正し、性能を計測してから責務を分離する順序が適している。新しい外観テーマを追加する土台はあるが、画面構成の異なるデザインを併存させるには、表示部品と操作・状態を分ける準備が必要。

この文書は調査結果であり、実装変更やデザイン案ではない。ファイルの行番号は以下のコミットを基準とする。

構造改善の最新状況は [状態・共通画像処理・保存境界・性能改善の記録](C:/Users/tyusk/vrchat-world-photo-manager/docs/architecture-improvements.md) を参照。

追記：その後、ユーザーの指示を受けてマスクの非同期競合防止と回帰テスト整備、およびバックアップ・World情報の保全を実装した。最新の変更範囲・結果は [レタッチの回帰確認](C:/Users/tyusk/vrchat-world-photo-manager/docs/retouch-regression-checks.md) と [データ保全の回帰確認](C:/Users/tyusk/vrchat-world-photo-manager/docs/data-safety-regression-checks.md) に記録している。以下の問題・未検証範囲は初回調査時点の記録。

| 調査対象 | 確認結果 |
| --- | --- |
| アプリ | WorldShot Log 2.5.3 |
| HEAD | `3b7a73e` — Release WorldShot Log v2.5.3（2026-07-18） |
| 作業ブランチ | `codex/subject-mask-composition` |
| 開始時の作業ツリー | 追跡対象の未コミット変更なし |
| 主な実行基盤 | Electron 41.0.2、JavaScript、better-sqlite3 12.8.0、onnxruntime-web 1.26.0 |
| 配布 | Windows、Electron Forge 7.11.1、Squirrel、GitHub Releasesを利用する更新機構 |
| 調査範囲 | README、保守・引き継ぎ資料、リリース手順、直近リリースノート、src、検証スクリプト、ビルド設定 |

公開サーバー上の最新リリースや依存ライブラリの最新情報は今回照会していない。バージョンはローカルのコード・インストール済み依存関係についての記録。

**把握した仕様**

| 領域 | 現行の主な機能・制約 |
| --- | --- |
| 写真管理 | ファイル・フォルダ・ドロップから取り込み、登録フォルダ再走査、年月日／ワールド別表示、検索、ラベル、お気に入り、複数選択 |
| 写真情報 | メタデータ解析、World ID／URL、公式情報の補完とキャッシュ、手動編集、メモ、プリントのノート |
| 画像編集 | 色・明るさ補正、RGB／HSVカーブ、プリセット、自動／学習補正、切り抜き・回転・反転、ぼかし・目隠し、テキスト、画像オーバーレイ |
| 被写体選択 | ローカルONNX推論、被写体／背景別の適用範囲、マスクブラシ、背景透過PNG保存。軽量モデル同梱、追加モデルは必要時に取得 |
| 保存 | 元画像を上書きしない編集保存、SQLiteの写真カタログ、設定、サムネイル、モデル・オーバーレイ素材 |
| 保守 | 健全性確認、再取り込み、サムネイル再生成、バックアップ／復元、CSV／JSON書き出し |
| 外観・言語 | dark／light、フォント、背景画像、標準／コンパクト表示、日本語・英語・韓国語 |
| 対象環境 | Windows向け。private／non-public worldの自動情報取得は未対応とREADMEに記載 |

主要な仕様の入口は [README.md](C:/Users/tyusk/vrchat-world-photo-manager/README.md)。独立した最新の包括的な設計仕様書は確認できず、実装と変更履歴を合わせて読む必要がある。

**実装の構成**

| ファイル | 規模の目安 | 担当 |
| --- | --- | --- |
| [src/index.js](C:/Users/tyusk/vrchat-world-photo-manager/src/index.js) | 約8,000行／222 KB | Electron main、IPC、取り込み、画像解析、World情報取得、モデル管理、ファイル操作、バックアップ、更新 |
| [src/preload.js](C:/Users/tyusk/vrchat-world-photo-manager/src/preload.js) | 約200行 | `window.electronAPI` とIPCの橋渡し |
| [src/db.js](C:/Users/tyusk/vrchat-world-photo-manager/src/db.js) | 約1,600行／41 KB | スキーマ補完、SQL、写真・ラベル・フォルダ・Worldキャッシュ、バックアップデータ |
| [src/renderer.js](C:/Users/tyusk/vrchat-world-photo-manager/src/renderer.js) | 約27,800行／815 KB | UI状態、DOM生成、ギャラリー、モーダル、検索、設定、画像編集、AI処理の制御 |
| [src/index.html](C:/Users/tyusk/vrchat-world-photo-manager/src/index.html) | 約2,200行 | 画面とモーダルの静的な土台 |
| [src/index.css](C:/Users/tyusk/vrchat-world-photo-manager/src/index.css) | 約8,900行／217 KB | テーマ、共通部品、各画面、状態別の上書き |
| [src/i18n.js](C:/Users/tyusk/vrchat-world-photo-manager/src/i18n.js) | 約1,800行 | 翻訳辞書とDOMの翻訳適用 |
| [src/photo-editor-worker.js](C:/Users/tyusk/vrchat-world-photo-manager/src/photo-editor-worker.js) | 約1,700行 | Workerでの画像補正 |
| [src/photo-editor-subject-mask.js](C:/Users/tyusk/vrchat-world-photo-manager/src/photo-editor-subject-mask.js)、[src/photo-editor-image-overlay.js](C:/Users/tyusk/vrchat-world-photo-manager/src/photo-editor-image-overlay.js) | 各独立モジュール | マスク・オーバーレイの状態や補助処理 |

基本の流れは、renderer → preloadのAPI → main → DB／ファイル処理 → renderer。画像編集はrenderer側のCanvas処理と専用Workerを組み合わせている。フロントエンドフレームワークへの移行を前提にしなくても、この境界を利用した段階的な分割は可能。

**先に修正したい、再現できた問題**

1. **優先度：高 — 不正なバックアップJSONを復元すると、登録データが空になる。**
   [index.js:1247](C:/Users/tyusk/vrchat-world-photo-manager/src/index.js:1247) はJSONオブジェクトと、存在する場合の `appName` だけを検証する。[db.js:1464](C:/Users/tyusk/vrchat-world-photo-manager/src/db.js:1464) はテーブルがなければ空配列として扱い、既存テーブルを削除して置換する。隔離したメモリ内DBで `restoreBackupSnapshot({})` を実行すると、例外なしで写真件数が1→0になった。実アプリ経路では [index.js:1255](C:/Users/tyusk/vrchat-world-photo-manager/src/index.js:1255) で復元前の自動バックアップを作るため、復旧手段はある。元写真ファイルの削除を再現したものではない。必須構造・スキーマ版・行データの検証を、置換処理の前に追加する必要がある。

2. **優先度：高 — AI処理の完了時に、処理を始めた写真かどうかを確認していない。**
   [renderer.js:9725](C:/Users/tyusk/vrchat-world-photo-manager/src/renderer.js:9725) は推論を待ったあと、現在のグローバルな `photoEditorState` に結果を反映する。編集画面を閉じる処理は [renderer.js:20377](C:/Users/tyusk/vrchat-world-photo-manager/src/renderer.js:20377) で状態をnullにする。対象関数をNodeのVMに抽出し、推論を模したPromiseを保留して状態を写真Bへ差し替えると、写真AのマスクがBへ反映された。状態をnullにしたケースでは `Cannot set properties of null (setting 'subjectMask')` が発生した。実モデルを使ったUI操作全体の再現は未実施。編集セッションと要求IDを保持し、閉じた・切り替えた・新しい処理で置き換わった要求の結果を破棄する必要がある。画像読み込みには既に同種のガードがある（[renderer.js:20422](C:/Users/tyusk/vrchat-world-photo-manager/src/renderer.js:20422)）。

3. **優先度：中 — 再取り込みで手動設定したWorld URL／IDが失われる。**
   手動URLは通常の `world_id`／`world_url` 列に保存される。一方、[db.js:381](C:/Users/tyusk/vrchat-world-photo-manager/src/db.js:381) の同一ハッシュ更新はこれらを画像由来の値で上書きする。画像にWorld情報がないケースを隔離DBで再現すると、手動URL／IDがnullになり、手動表示名とメモは残った。明示的な既存写真の再取り込みは [index.js:6136](C:/Users/tyusk/vrchat-world-photo-manager/src/index.js:6136) で取り込み処理を呼ぶ。手動値と自動値の由来を保持する設計が必要。

**性能改善の候補 — 速度・メモリの効果は未計測**

| 候補 | コードで確認したこと | 次の判断材料 |
| --- | --- | --- |
| 取り込み処理をmainから分離 | [index.js:4865](C:/Users/tyusk/vrchat-world-photo-manager/src/index.js:4865) に全画像読み込み、デコード、メタデータ抽出、SHA-256、サムネイル生成が集まる。Promiseの同時実行数制御はCPU処理の別スレッド化ではない | 取り込み枚数・解像度別の処理時間、mainの応答遅延、ピークメモリを計測し、Worker等へ移す部分を選ぶ |
| 一覧取得のページング・集計SQL化 | [db.js:429](C:/Users/tyusk/vrchat-world-photo-manager/src/db.js:429) の一覧SQLと [index.js:4699](C:/Users/tyusk/vrchat-world-photo-manager/src/index.js:4699) の返却前処理は対象全件を扱う。Worldサイドバーも写真から集計する | 年別・World別の件数増加に対し、DB・前処理・IPC・描画の時間を分けて計測する |
| ギャラリーDOMの上限管理 | [renderer.js:23747](C:/Users/tyusk/vrchat-world-photo-manager/src/renderer.js:23747) は段階的にカードを追加する。長くスクロールすると表示済みDOMが蓄積する | 数千枚以上の表示でDOM数・メモリ・スクロール応答を調べ、必要なら仮想化する。選択範囲とキーボード移動の仕様を保持する |
| 被写体マスク利用時の編集処理 | [renderer.js:16300](C:/Users/tyusk/vrchat-world-photo-manager/src/renderer.js:16300) はマスク有効時にWorker経路を避け、rendererでのCanvas処理へ戻す | マスクを使った補正・ブラシ操作の応答時間とメモリを計測し、マスク合成のWorker対応やCanvas再利用を検討する |
| 検索・整列処理の重複削減 | [renderer.js:4528](C:/Users/tyusk/vrchat-world-photo-manager/src/renderer.js:4528) は写真ごとに検索語を正規化し、フィルター後にも再ソートする | 検索語の正規化を1回にし、検索用文字列・日時の更新時キャッシュと既存の整列順を活用できるか確認する |
| マスクとUndo履歴の分離 | [renderer.js:7828](C:/Users/tyusk/vrchat-world-photo-manager/src/renderer.js:7828) はマスクPNG文字列も含む履歴をJSON化して比較する | ブラシ操作と、マスクを変更しない操作のシリアライズ時間を計測する。画像資産への参照と操作状態を分ける余地がある |

既存の性能ログ設定は [index.js:450](C:/Users/tyusk/vrchat-world-photo-manager/src/index.js:450) の `WORLDSHOT_PERF_LOG`／`WORLDSHOT_PERF_LOG_THRESHOLD_MS`。現時点では高速化率や最大対応枚数を断定できない。

維持したい既存の対策は、SQLite WAL・索引・バッチトランザクション、ラベル一括取得、ハッシュによる重複登録防止、サムネイルキャッシュ、段階描画、画像の遅延読み込み、画像補正Worker、操作中の縮小プレビュー、LUT等のキャッシュ。単純に「最適化がないコード」とは評価できない。

**将来のデザイン追加に向けた整理**

| 追加したい範囲 | 現状 | 準備する変更 |
| --- | --- | --- |
| 配色・質感の違うテーマ | CSS変数と `body[data-theme]` がある。ただし [renderer.js:2375](C:/Users/tyusk/vrchat-world-photo-manager/src/renderer.js:2375) はlight以外をdarkに丸める | テーマID・表示名・明暗属性を定義一覧へまとめ、テーマ別の色・影・角丸等を変数に集約する |
| 画面構成の違うデザイン | [renderer.js:22775](C:/Users/tyusk/vrchat-world-photo-manager/src/renderer.js:22775) でDOMを実行時に移設し、[renderer.js:23066](C:/Users/tyusk/vrchat-world-photo-manager/src/renderer.js:23066) のカード生成が操作処理も担う | ギャラリー、カード、ツールバー、モーダルを部品に分け、データと操作を渡して描画できる境界を作る |

CSSは [index.css:5](C:/Users/tyusk/vrchat-world-photo-manager/src/index.css:5) の変数定義に加え、ライト用の調整が7699行・8072行・8367行などに累積している。テーマ定義、部品、レイアウト、表示状態の順に責務を整理すると、追加テーマの修正範囲を限定できる。

ギャラリーの寸法にも二重管理がある。[renderer.js:917](C:/Users/tyusk/vrchat-world-photo-manager/src/renderer.js:917) の事前描画計算用定数はカード220px／間隔16pxだが、[index.css:8657](C:/Users/tyusk/vrchat-world-photo-manager/src/index.css:8657) のコンパクト表示は148px／12px。新デザインでも描画枚数推定を合わせるため、寸法の取得元を統一する余地がある。現時点で実害は計測していない。

翻訳は日本語本文をキーにし、[i18n.js:1734](C:/Users/tyusk/vrchat-world-photo-manager/src/i18n.js:1734) でDOM変更を監視して適用する。文言やDOM構成を変える際には日英韓で確認し、部品分離に合わせて安定した翻訳キーの使用を増やすと保守しやすい。

今回、現行画面の視覚品質の評価や新デザインの選定は実施していない。

また、rendererとWorkerには同種の色補正処理がある（[renderer.js:12888](C:/Users/tyusk/vrchat-world-photo-manager/src/renderer.js:12888)、[photo-editor-worker.js:713](C:/Users/tyusk/vrchat-world-photo-manager/src/photo-editor-worker.js:713)）。片方だけの修正で出力がずれる保守リスクを減らすため、純粋な計算処理の共通化もUI分離と並行して検討する。現在の出力差を確認したものではない。

**仕様資料と検証基盤の課題**

- ローカルの [AI_MAINTENANCE_GUIDE.md:14](C:/Users/tyusk/vrchat-world-photo-manager/AI_MAINTENANCE_GUIDE.md:14) はv1.2.0を前提にしている。[RELEASE.md:350](C:/Users/tyusk/vrchat-world-photo-manager/RELEASE.md:350) の現行公開版欄もv2.4.2。実コードはv2.5.3なので、そのまま現行仕様として扱えない。
- READMEからリンクされる保守ガイドは [.gitignore:99](C:/Users/tyusk/vrchat-world-photo-manager/.gitignore:99) で管理対象外。新しいチェックアウトで共有できる現行の保守資料が必要。
- [package.json:33](C:/Users/tyusk/vrchat-world-photo-manager/package.json:33) のlintは説明を出すだけ。今回の確認範囲では追跡されたCI設定もない。既存スモークをまとめて実行する検証入口と静的解析の整備が有効。
- [scripts/smoke-photo-editor-presets.js:86](C:/Users/tyusk/vrchat-world-photo-manager/scripts/smoke-photo-editor-presets.js:86) は管理対象外の `img/notuse.png` に依存する。この環境では存在してテストが通ったが、新規チェックアウトでの再現性を整える必要がある。
- mainには意図的に残したlegacy／quarantineブロックがある。分割前に有効な実装と参照先を確認し、古い関数を誤って復活させないようにする。

**実施した確認**

| 確認 | 結果・範囲 |
| --- | --- |
| JS構文検査 | src・scripts・forge設定の17ファイルを1ファイルずつ `node --check`。すべて通過 |
| `git diff --check` | 通過 |
| `npm.cmd ls --depth=0` | トップレベル依存に不足・不整合の報告なし |
| `npm.cmd run audit:i18n -- --strict` | 抽出候補457件、翻訳漏れ0件。全画面・全文言の完全性を保証するものではない |
| `npm.cmd run smoke:data` | 旧DB移行と正常なバックアップ復元を通過。隔離した一時DBを使用 |
| `npm.cmd run smoke:coordinates` | 画像座標への追従と画像オーバーレイのリサイズを通過 |
| `npm.cmd run smoke:mask-brush` | ブラシ追加／削除、濃度サンプル、Undo／Redo、ズーム関連を通過 |
| `npm.cmd run smoke:presets` | エアリーホワイト4種の出力・関係と日英韓ラベルの確認を通過 |
| 追加の隔離DB検証 | 不正JSONによる空化、再取り込み相当の更新で手動World URL／IDが消える問題を再現 |
| 追加の関数単位検証 | 遅延PromiseによるAI処理結果の別写真への反映／閉じた状態への書き込みエラーを再現 |

画像編集テストの結果は以下に保存されている。

- [座標処理の結果](C:/Users/tyusk/vrchat-world-photo-manager/out/smoke-coordinate-space/20260909-204848/result.json)
- [マスクブラシの結果](C:/Users/tyusk/vrchat-world-photo-manager/out/smoke-mask-brush/20260909-204904/result.json)
- [プリセットの結果](C:/Users/tyusk/vrchat-world-photo-manager/out/smoke-presets/20260909-204917/result.json)

実ユーザーの写真コレクションを使う大規模性能計測、実AIモデルによる競合のUI再現、全画面の視覚確認、パッケージ再生成、配布版起動確認、依存関係の脆弱性監査は実施していない。開発版テストではContent-Security-PolicyについてElectron警告が出たため、配布前のセキュリティ設定確認に残す。

**ユーザー指定の改修条件（2026-09-09追記）**

デザイン追加は最後に行う。過去の改修では特にレタッチ機能周辺で不具合が頻発したとの申告があり、今回の最優先条件は、実装済み機能の動作と出力を維持すること。高速化や構造整理は、この条件を満たせた変更から進める。

既存スモークの通過だけでは、レタッチ機能全体を確認できない。現在の3本は、主に切り抜き位置への追従・画像オーバーレイの拡縮、マスクブラシ、エアリーホワイト4種を対象にしている。保存処理にはプレビューとは別の経路がある（[renderer.js:20572](C:/Users/tyusk/vrchat-world-photo-manager/src/renderer.js:20572)）。改修対象に応じ、以下の不足を先に補う。

| 守る動作 | 改修前後に確認すること | 現状の自動検証 |
| --- | --- | --- |
| 保存と元画像の保全 | 実ファイルの書き出し・再読込、寸法、形式、透過、レイヤーの合成結果、キャンセル時の状態、元画像のハッシュが不変であること。再登録後のWorld情報・撮影日・メモ・ラベルも確認 | 編集スモーク3本では保存経路全体を確認していない |
| 切り抜きと座標 | ズーム・移動・90度回転・自由回転・反転に対し、テキスト／画像／ぼかし／マスクが意図した位置に残ること。固定解像度・正方形キャンバスの保存も確認 | ズーム・移動による位置変化と画像拡縮は一部確認。座標テストは主に移動量の下限判定で、正確な位置・組合せ・保存結果の確認は不足 |
| マスクと適用範囲 | 被写体／背景別の補正、ぼかし、テキスト、画像合成。境界、半透明、ブラシ修正後の背景透過保存。目隠しが保存後も対象を覆い、操作ガイドやマスクの表示色が書き出されないこと | ブラシとその履歴は一部確認。合成・保存の組合せは不足 |
| Undo／Redo | 補正・切り抜き・テキスト・画像・ブラシを続けて操作し、戻す／やり直すで値・配置・出力が復元されること。Undo後の新規操作で履歴が適切に分岐すること | マスクの削除に対するUndo／Redoは確認。複合操作は不足 |
| 画像処理の一致 | 同一状態・同一寸法でWorker経路と通常経路の結果を比較する。プレビューと保存でも色・配置・適用範囲を確認 | プリセットのプレビューは一部確認。経路比較は不足 |
| 非同期処理と画面切り替え | AI・素材読込・プレビュー処理中の閉じる／再度開く／別写真への切り替えで、結果や履歴が混入しないこと。失敗時に再操作できること | AIの問題は関数単位で再現済み。継続実行する回帰テストは未整備 |
| 補正・プリセット | 個別補正、RGB／HSVカーブ、既存プリセット、ユーザープリセットの保存・読込、自動／学習補正、強度変更・リセットの動作 | エアリーホワイト4種が中心。他の機能は不足 |

実装時の進め方は以下とする。

- 変更対象の正常動作を固定した入力画像と編集手順で記録し、修正する不具合は修正前に失敗するテストで再現する。既知の不具合を正しい基準として固定しない。
- 1回の変更を1つの不具合または計測済みの改善点に絞る。画像計算、座標、履歴、UIの構造変更をまとめない。ファイルの大きさだけを理由に一括分割しない。
- 変更対象のテストに加え、影響を共有するレタッチ操作・保存結果と既存スモークを確認する。出力寸法や状態値は厳密に比較し、画像比較に許容差が必要な場合は理由と範囲を定める。
- 想定外の出力差・操作不良が出たら、その変更の原因を解消するまで次の最適化へ進めない。検証結果は実行した範囲と未確認の範囲を分けて報告する。
- 固定テスト画像は新規チェックアウトでも利用できるようにし、実ユーザーデータから隔離する。高速化の判断では速度・メモリと出力維持の両方を確認する。

この追記は改修条件と検証範囲の整理であり、新しい回帰テストを実装済みという意味ではない。

**次の開発順序**

1. 検証用画像と既存スモークの再現性を整え、最初に触るレタッチ機能の正常動作・保存結果を確認するテストを先に用意する。
2. 不正バックアップ・AI非同期競合・手動World情報保全を、それぞれの再現テストとともに個別の小さな変更で修正する。
3. 代表的な写真枚数・解像度で起動、一覧、検索、取り込み、マスク利用時の編集応答を計測する。測定結果で最適化の順序を決める。
4. 改善が必要な箇所を、出力と操作を維持する検証を伴って最適化・分割する。現行仕様とデータ保全の条件を共有資料へ反映する。
5. 既存機能の維持を確認した後、最後にテーマ定義・CSS・表示部品を整理して新デザインを追加する。

初回調査で追加したものはこの調査文書のみ。その後の実装作業は、冒頭のレタッチ回帰確認文書を参照。
