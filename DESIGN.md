# DESIGN

## お知らせモーダル（v3.0.0）

### 目的と範囲

- やること: v3.0.0 に更新した人の最初の起動で、「デザインがふえた」「ワールド一覧ができた」の 2 点を、文とスクリーンショットで 1 回だけ伝える。
- やらないこと: バージョン番号の変更、README・リリースノートの更新、`src/index.css` / `src/index.html` の変更、各案の CSS（`src/ui-next.css` など）への追記、新しい IPC の追加、スクリーンショットの撮影（呼び出し元が行う）。

### 確定した判断

| 項目 | 判断 | 理由 |
| --- | --- | --- |
| 画面構成 | 2 ページ（1: デザイン / 2: ワールド一覧）。左に画像、右に文と操作 | 最小の高さ 620px でも画像を幅 660px で見せられる。1 画面に 2 枚並べると画像が半分になる |
| デザインの試し方 | モーダルの中のボタンで、その場で切り替える（`applyUiSkin` を呼ぶ） | 標準がランタンに変わるので、従来のデザインに 1 回で戻れる出口をその場に置く |
| 画像の切り替え | デザインのボタンを押すと、左の画像もそのデザインの見本に替わる | モーダルが画面をおおうので、実画面だけでは全体の見た目が分からない |
| ワールド一覧への行き先 | 2 ページ目の右に「ワールド一覧を見る」（押すと閉じてワールド一覧を開く） | 読んだ直後に新機能へ入れる |
| 閉じ方 | 右上の ×、Esc、背景クリック、2 ページ目の「閉じる」 | 既存の `sub-modal` と同じ動きにそろえる |
| 見たことの記録 | 閉じたときではなく、表示したときに記録する | 途中で終了しても 2 回出さない |
| もう一度見る手段 | 設定 > デザインの節に「v3.0.0 のお知らせを見る」を置く | Esc や背景クリックで閉じやすいので、戻れる場所を用意する。この節は `renderer.js` が組み立てているので `index.html` を触らずに済む |
| 動き | ページ送りも画像の切り替えも即時。アニメーションを足さない | 新デザインの「浮かせない・光らせない・ぼかさない」と、動きを減らす設定の両方を満たす |
| 表示の条件 | お知らせの ID（定数）で判定。アプリのバージョン番号は見ない | package.json は 2.5.4 のままでも動く |

### 表示の判定

- 定数（`src/renderer.js`、ほかの `*_STORAGE_KEY` の近く）
  - `WHATS_NEW_ID = '3.0.0'`
  - `WHATS_NEW_SEEN_STORAGE_KEY = 'vrchat-world-photo-manager-whats-new-seen'`（値は最後に見たお知らせの ID）
  - `WHATS_NEW_IMAGE_BASE = '../img/whats-new/3.0.0'`
- 呼ぶ場所: `bootstrapRenderer` の最後（`finishAppInitialization` の後）。`initializeApp` が成功したときだけ。
- 流れ（`maybeShowWhatsNewOnStartup`）
  1. 保存されている ID が `WHATS_NEW_ID` と同じ → 何もしない。localStorage が読めないときも何もしない。
  2. 初期化が失敗した、または写真の枚数が分からない → 何もしない。記録もしない（次の起動でやり直す）。
  3. 枚数が 0 → 記録だけして、表示しない。
  4. 枚数が 1 以上 → ほかのモーダル（`.sub-modal:not(.hidden)`、写真詳細）が開いていなければ、記録してから開く。開いていたら何もしない（次の起動でやり直す）。
- 写真の枚数は、起動時に読み込み済みの年月一覧のデータ（年ごとの枚数の合計）から取る。
- 設定から開いたときは、記録を読み書きしない。

### 画面構成と DOM

`ensureWhatsNewModal` が 1 回だけ作り、`document.body` に足す。

```
div#whats-new-modal.sub-modal.hidden
  div.sub-modal-backdrop
  div.sub-modal-content.whats-new-content   role="dialog" aria-modal="true" aria-labelledby="whats-new-label"
    button.sub-modal-close                  aria-label="お知らせを閉じる"（中は span.material-symbols-outlined「close」）
    div.whats-new-header
      p#whats-new-label.whats-new-kicker    「v3.0.0 のお知らせ」
    div.sub-modal-body.whats-new-body
      section.whats-new-page[data-whats-new-page="designs"]
        div.whats-new-shot                  role="img" aria-label="写真一覧の画面の見本"
          img.whats-new-shot-image[data-whats-new-skin="<id>"] × 6（alt=""。選択中の 1 枚以外は hidden）
        div.whats-new-text
          h3.whats-new-title                tabindex="-1"
          p.whats-new-lead
          div.whats-new-skin-options        role="group" aria-label="デザイン"
            button.settings-font-option-button.whats-new-skin-option[data-whats-new-skin="<id>"] × 6（aria-pressed）
          p.whats-new-note
      section.whats-new-page[data-whats-new-page="worlds"]（最初は hidden）
        div.whats-new-shot                  role="img" aria-label="ワールド一覧の画面の見本"
          img.whats-new-shot-image          alt=""
        div.whats-new-text
          h3.whats-new-title                tabindex="-1"
          p.whats-new-lead
          ul.whats-new-points > li × 3
          button.primary-link-button.whats-new-open-worlds
          p.whats-new-note
    div.whats-new-footer
      p.whats-new-step[data-i18n-ignore]    「1 / 2」「2 / 2」
      div.whats-new-nav
        button.small-action-button.secondary.whats-new-prev   「戻る」（1 ページ目は hidden）
        button.whats-new-next               1 ページ目「次へ」（primary-link-button）/ 2 ページ目「閉じる」（small-action-button secondary）
```

- デザインのボタンの並びと ID は設定と同じ: `legacy` / `next` / `hc` / `noir` / `chocomint` / `soda`（`UI_SKIN_*` の値）。表示名も設定と同じ文字列を使う。
- 文は 1 要素に 1 つのテキストノードで入れる（途中に `span` や `b` をはさまない）。辞書がテキストノード全体をキーにして訳すため。
- 画像に `alt` は使わず、枠の `aria-label` に書く。`alt` は自動で訳されない（訳されるのは `aria-label` / `title` / `placeholder` / `label`）。

### 状態と挙動

- 状態: `whatsNewPageIndex`（0 か 1）、閉じたあとにフォーカスを戻す要素。選択中のデザインは持たず、`readSavedUiSkin()` から読む。
- 開く（`openWhatsNewModal`）
  - 1 ページ目に戻す。今の言語（`window.WorldShotI18n.getLanguage()`）で全画像の `src` を入れ直す。
  - `readSavedUiSkin()` でボタンと画像を合わせる。開いただけでは `applyUiSkin` を呼ばない（デザインの保存値を書きかえない）。
  - `openSubModalElement` で開き、フォーカスを「次へ」に置く。
- デザインのボタン: `applyUiSkin(id)` → `syncWhatsNewSkinUi(id)`（`aria-pressed` と `is-active`、画像の `hidden` を切り替える）。ページは変えず、フォーカスも押したボタンのまま。
- ページ送り（`setWhatsNewPage`）: `section` の `hidden`、「1 / 2」、フッターのボタンを切り替え、移った先の `h3` にフォーカスを移す。
- 「ワールド一覧を見る」: モーダルを閉じてから、サイドバーの「ワールド一覧」を押したときと同じ処理を呼ぶ。
- 設定の「v3.0.0 のお知らせを見る」: 設定を閉じてから開く（重ねない）。`ensureSettingsUiSkinSection` で、デザインのボタン群の下に `small-action-button secondary` を 1 つ足す。
- 画像が読めないとき: その言語で失敗したら `ja` を 1 回だけ試す。それも失敗したら枠に `is-missing` を付け、枠だけ残す。

### キーボードと読み上げ

| 操作 | 動き |
| --- | --- |
| Esc | 閉じる。`bindGlobalDocumentInteractions` の Esc の並びで、ドロップダウンの判定の直後に入れる |
| ← / → | 前 / 次のページ。端では何もしない |
| Tab / Shift+Tab | モーダルの中だけを回る（先頭と末尾でつなぐ） |
| Enter / Space | フォーカス中のボタンを押す（標準の動き） |

- ← / → で後ろの写真の選択が動かないように、写真一覧の矢印キーの条件（`trackedFolderModal` などを並べている箇所）にこのモーダルを足す。
- 閉じたら、開く前にフォーカスがあった要素に戻す（無くなっていれば何もしない）。
- 役割と名前: ダイアログの名前は「v3.0.0 のお知らせ」、各ページの見出しは `h3`、デザインのボタンは `aria-pressed`。「1 / 2」には `aria-live="polite"` を付ける。
- `h3` はフォーカスの輪郭を出さない（`outline: none`）。ボタンのフォーカス表示は既存のまま。

### 寸法

- モーダル: 幅 `min(1040px, 100vw - 48px)`、最大の高さ `100vh - 40px`。縦に「見出し行 / 本体 / フッター」の 3 段。スクロールするのは本体だけで、フッターは常に見える。
- ページ: 2 列のグリッド `minmax(0, 1fr) 300px`、間隔 24px、上ぞろえ。
- 画像の枠: 縦横比 16:10、幅は列いっぱい（幅 1040px のとき 660 × 412px）。`object-fit: cover`、線 1px、角は `var(--radius-md)`、地は `var(--surface-soft)`。
- 余白: 見出し行は上 20px・左 28px・右 64px（× の分）。本体は 12px 28px 16px。フッターは 12px 28px 16px で、上に線 1px。
- 文字: 見出し行 12px 太字、ページの見出し 22px / 行間 1.3、本文 14px / 1.6、箇条書き 13px / 1.6、注記 12px / 1.5。ボタンは高さ 36px 以上、13px。
- デザインのボタン: 2 列 × 3 行、間隔 8px。ボタンの中は折り返してよい。
- 幅 1100px・高さ 620px のとき: モーダルは 1040 × 約 550px で、スクロールなしで収まる。これより低いときは本体が縦にスクロールする。幅が 900px 未満なら 1 列（画像が上、最大幅 560px）。
- あふれ対策: 文には高さの固定・`white-space: nowrap`・省略記号を使わない。`overflow-wrap: anywhere` を付ける。

### CSS

- 新規 `src/whats-new.css`: 構造と従来デザインの見た目。`initializeRendererUi` の `world-cards.css` の行の隣で `appendUiSkinStylesheetLink('whats-new-stylesheet', './whats-new.css')` として読み込む。
- 色・線・角は従来の変数（`--text-0` / `--text-2` / `--line` / `--surface-soft` / `--accent` / `--radius-md`）だけで書く。新デザインでは `src/ui-skin-common.css` がこれらを `--skin-*` に向け直すので、そのまま各案の見た目になる。ボタンと `h3` も `.sub-modal` の共通の規則で付く。
- 新デザイン向けの直しが要るときだけ、`src/ui-skin-common.css` に節を 1 つ足す。各案の CSS には書かない。
- 流用するクラス（`settings-font-option-button` / `small-action-button` / `primary-link-button`）の従来デザインの見た目が設定画面の外で効かないときは、`src/whats-new.css` に `body:not([data-ui])` で書く。
- `.sub-modal-content` の幅や余白が `index.css` の `!important` に負けるときだけ、`#whats-new-modal` を付けて詳細度を上げる。
- ホバーで変えるのは地と線の色だけ。`transform` / `box-shadow` / `filter` / `backdrop-filter` は書かない。

### 関数の分け方（`src/renderer.js`）

| 関数 | 役割 |
| --- | --- |
| `readWhatsNewSeenId` / `markWhatsNewSeen` | localStorage の読み書き（失敗は握りつぶす） |
| `maybeShowWhatsNewOnStartup` | 起動時の判定（上の流れ） |
| `ensureWhatsNewModal` | DOM を 1 回だけ作り、イベントをつなぐ。閉じる操作は `bindSubModalCloseTriggers` |
| `openWhatsNewModal` / `closeWhatsNewModal` | 開閉。`openSubModalElement` / `closeSubModalElement` を使う |
| `setWhatsNewPage` | ページの切り替えとフッターの更新 |
| `syncWhatsNewSkinUi` | デザインのボタンと画像を合わせる |
| `handleWhatsNewKeydown` | ← / → と Tab の折り返し（モーダルの要素に付ける） |
| `openWorldListFromWhatsNew` | 閉じてワールド一覧を開く |

- `applyUiSkin` と `syncUiSkinOptionButtons` は変えない。
- `src/renderer.js` は BOM 付き UTF-8・CRLF のまま。コメントにアポストロフィを書かない。

### 文言（日本語をキーにして `src/i18n.js` の en / ko に足す）

| 場所 | 日本語 | English | 한국어 |
| --- | --- | --- | --- |
| 見出し行 | v3.0.0 のお知らせ | New in v3.0.0 | v3.0.0 업데이트 안내 |
| × の名前 | お知らせを閉じる | Close announcement | 안내 닫기 |
| 1 見出し | デザインがふえました！ | New designs! | 새 디자인이 추가되었습니다! |
| 1 本文 | あたらしいデザインを 5 つ追加しました。ボタンを押すと、その場で切り替わります。 | Five new designs have been added. Select a button to switch right away. | 새 디자인 5종을 추가했습니다. 버튼을 누르면 바로 바뀝니다. |
| 1 注記 | はじめはランタンになっています。ライト表示は従来のデザインだけです。あとから「設定 > デザイン」でも変えられます。 | The app now opens in Lantern. Light mode is available only in the classic design. You can change the design later in Settings > Design. | 처음에는 랜턴으로 표시됩니다. 라이트 모드는 기존 디자인에서만 쓸 수 있습니다. 나중에 설정 > 디자인에서도 바꿀 수 있습니다. |
| 1 画像の名前 | 写真一覧の画面の見本 | Preview of the photo list | 사진 목록 화면 미리보기 |
| 2 見出し | ワールド一覧ができました | Browse by world | 월드 목록이 생겼습니다 |
| 2 本文 | 撮ったワールドがカードでならびます。カードを押すと、そのワールドの写真がひらきます。 | Every world you have photographed appears as a card. Select a card to open the photos from that world. | 사진을 찍은 월드가 카드로 표시됩니다. 카드를 누르면 그 월드의 사진이 열립니다. |
| 2 箇条書き 1 | 上の検索で、ワールド名からしぼりこめます | Filter by world name with the search box at the top | 위쪽 검색으로 월드 이름을 좁혀 볼 수 있습니다 |
| 2 箇条書き 2 | 枚数順・名前順・日付順でならべかえられます | Sort by count, name, or date | 장수순·이름순·날짜순으로 정렬할 수 있습니다 |
| 2 箇条書き 3 | マウスの戻るボタンで一覧にもどれます | Return to the list with the mouse back button | 마우스의 뒤로 버튼으로 목록에 돌아갈 수 있습니다 |
| 2 ボタン | ワールド一覧を見る | Open Worlds | 월드 목록 보기 |
| 2 注記 | いつでもサイドバーの「ワールド一覧」からひらけます。 | You can open it anytime from Worlds in the sidebar. | 사이드바의 월드 목록에서 언제든 열 수 있습니다. |
| 2 画像の名前 | ワールド一覧の画面の見本 | Preview of the world list | 월드 목록 화면 미리보기 |
| フッター | 次へ | Next | 다음 |
| フッター | 戻る | Back | 이전 |
| 設定のボタン | v3.0.0 のお知らせを見る | View the v3.0.0 announcement | v3.0.0 업데이트 안내 보기 |

- 既存のキーをそのまま使うもの（足さない）: 閉じる、デザイン、従来のデザイン、トランジット、ランタン、ノワール、チョコミント、ソーダ。
- 韓国語の「枚数順・名前順・日付順」「ワールド一覧」は、辞書にある既存の訳語に合わせて置きかえる。
- 英語の文にアポストロフィを入れない（上の表は入れていない）。

### スクリーンショット

- 置き場所: `img/whats-new/3.0.0/<ja|en|ko>/<名前>.jpg`。1 言語 7 枚、合計 21 枚。`forge.config.js` の除外リストには入れない（変更なしで同梱される）。
- 共通の撮り方
  - ウィンドウの中身を 1120 × 700（CSS px、拡大率 100%）にして、中身全体を撮る。切り出しなし。OS のタイトルバーは入れない。
  - JPEG・品質 80・1120 × 700px のまま保存。1 枚 200KB 以下が目安。
  - 状態: ダーク、背景画像なし、表示サイズは標準、選択なし、トーストや進行表示なし、スクロールは一番上、マウスはウィンドウの外。
  - 3 言語とも同じデータ・同じ状態で、言語だけ変える。
- 写すもの

| 名前 | 画面 | デザイン | 状態 |
| --- | --- | --- | --- |
| `design-legacy.jpg` | 写真一覧 | 従来のデザイン（ダーク） | 年月一覧で最新の年をひらき、写真が 12 枚以上ある月を選ぶ。写真が画面の下端まで埋まること |
| `design-next.jpg` | 同上 | トランジット | 同じ月・同じ状態 |
| `design-hc.jpg` | 同上 | ランタン | 同上 |
| `design-noir.jpg` | 同上 | ノワール | 同上 |
| `design-chocomint.jpg` | 同上 | チョコミント | 同上 |
| `design-soda.jpg` | 同上 | ソーダ | 同上 |
| `worlds.jpg` | ワールド一覧（すべてのワールド） | ランタン | 枚数順、検索は空。カードが 12 枚以上で、見えているカードすべてに画像があること |

- デザインの 6 枚は、デザイン以外を 1 つも変えずに撮る（切り替えたときに同じ画面の色ちがいに見えるようにする）。
- `worlds.jpg` はランタンの 1 種類だけにする（全デザイン分は作らない。枚数を 21 枚に抑えるため）。

### 確認すること

- 6 デザイン × 3 言語に、従来デザインのライトを加えた組み合わせで、幅 1100 × 高さ 620 と最大化の両方を見る。文が切れない・重ならない・はみ出さない。
- 更新した人（写真あり）で 1 回だけ出る。新規（写真 0 枚）では出ず、そのあと写真を入れても出ない。
- デザインのボタンで実画面・ボタン・画像がそろって変わる。設定のデザインの選択とも一致する。
- モーダルを開いたまま ← / → を押しても、後ろの写真の選択が動かない。
- `npm.cmd run check`（i18n の監査を含む）が通る。

### 実装時に確かめること

- `showWorldCardView()` だけではサイドバーが「ワールド一覧」の表示に切り替わらない可能性がある。サイドバーの「ワールド一覧」ボタンが呼んでいる処理を使う。
- 写真の総数を取れる既存のデータ（年月一覧の枚数）の変数名と、`initializeApp` の成否の取り方。
- 2 ページ目の注記の「サイドバーの「ワールド一覧」」が、実際のボタンの場所と名前に合っているか。

### 実装して分かったこと（2026-11-12）

- ウィンドウの最小サイズは 1200 × 760（中身は約 1184 × 721）。この大きさでモーダルは 1040 × 544 になり、6 デザイン × 日英韓のどれでもスクロールなしで収まった。上の「幅 1100・高さ 620」は設計時の想定で、実際には起こらない。
- スクリーンショットは、中身 1120 × 700 のウィンドウを作れないので、中身 1280 × 800（同じ 16:10）で撮って 1120 × 700 に縮めている。
- スクリーンショット 21 枚は確定（2026-11-12 にユーザーが承認）。写真の枠は、撮影用に描いた風景風の絵（空・丘・太陽の色の面）。実際の写真は使わない。
- 韓国語は単語の途中で改行されるので、モーダルの文に `word-break: keep-all` を付けた（`html[lang="ko"]` のときだけ）。
- 新デザインのページ見出しは、共通の規則で各案の見出しの大きさになる（トランジット 30px、ほかは 24px）。韓国語は 2 行になることがあるが収まる。

### 未決事項

1. （解決済み）韓国語の文言は、母語話者の確認なしで出す（2026-11-12 に決定）。
2. `DESIGN.md` は CLAUDE.md で「新デザインの方針を書く場所」とされているが、今はこの節だけ。方針の節をあとで足すかどうか。

### 解決済み（2026-11-12）

- スクリーンショットに写すもの: 写真の枠には撮影用に描いた風景風の絵を入れる（`test_pic` の写真やアバターの写った写真は使わない）。ワールド名と作者名は撮影用の架空のもの。公式サムネイルは写さない。
- 新デザインは 5 つとも採用済みで、起動時の標準はランタン（CLAUDE.md の「新デザイン」の節のとおり）。お知らせの文はこの前提で確定。
