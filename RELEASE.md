# WorldShot Log リリース手順

## 0. この資料の目的

この資料は、WorldShot Log の Windows 配布版を安全にリリースするための手順書です。

対象:

- バージョン更新
- Windows ビルド
- git タグ作成
- GitHub Release 公開
- 自動アップデート配信

この資料は、**リリース作業だけに絞って整理したもの**です。

---

## 1. 前提

### 1-1. 現在の配布方式

WorldShot Log は、現状次の構成で配布しています。

- Windows 向け配布のみ
- `electron-forge + maker-squirrel`
- GitHub Releases を公開元として使用

### 1-2. 生成される主要成果物

`npm run make:win` 実行後、主に次の成果物を使います。

- `out/make/squirrel.windows/x64/WorldShotLogSetup.exe`
- `out/make/squirrel.windows/x64/worldshot_log-<version>-full.nupkg`
- `out/make/squirrel.windows/x64/RELEASES`

### 1-3. 命名ルール

- アプリ名: `WorldShot Log`
- npm package 名: `worldshot-log`
- git タグ形式: `v<version>`
- Windows インストーラ名: `WorldShotLogSetup.exe`
- Windows 実行ファイル名: `WorldShotLog`

---

## 2. リリース前チェック

リリース前に最低限確認すること:

1. 対象の改修が `main` にまとまっている
2. 先祖返りや不要差分が入っていない
3. 主要画面の動作確認が済んでいる
4. `README.md` など、必要な資料更新が済んでいる
5. 今回版の変更点を `release-notes/` にまとめている

補足:

- リリースノートは `release-notes/vX.Y.Z.md` 形式で追加する運用を推奨します
- 公開版の差分説明は、そのファイルを GitHub Release の本文に転用するのが安全です

---

## 3. バージョン更新

### 3-1. 更新対象

通常、最低限次を更新します。

- `package.json`
- `package-lock.json`
- `src/index.html`

### 3-2. 反映内容

- `package.json` の `version`
- `package-lock.json` の `version`
- `src/index.html` のタイトル表示

`src/index.js` のウィンドウタイトルは `app.getVersion()` を参照しているため、通常は追加修正不要です。

---

## 4. 動作確認

最低限の確認例:

- アプリ起動
- 既存データ読み込み
- 一覧表示
- カードモーダル表示
- 設定モーダル表示
- 画像取り込み
- 追加した機能の動作

改修後は、構文・単体・画像演算・翻訳・レタッチ・データ保全の回帰検証をまとめて実行します。手動での主要画面確認と併せて結果を確認してください。

```powershell
npm.cmd run check
git diff --check
```

チェック別のログと結果は `out/regression-checks/` に保存されます。失敗した項目や未実行の後続がある場合は、全体の確認完了とは扱いません。個別の実行方法は [保守用ハーネス](./docs/maintenance-harness.md) を参照してください。

公開前の作者確認には、合格した全体結果を指定して専用候補を生成できます。検証時のソースと現在のソースを照合し、独立した出力先にアプリ・インストーラ・隔離したサンプル環境を作ります。このコマンドは公開やインストールを行いません。

```powershell
npm.cmd run review:build -- --checks-result=out/regression-checks/<今回の結果>/result.json
```

今回の候補・検証結果・公開前の残確認は [v2.5.4 の公開準備](./docs/release-preparation-v2.5.4.md) を参照してください。以前の確認記録は [リリース前の確認](./docs/release-candidate-review-2026-09-10.md) に残しています。バージョンを更新した後はソースが変わるため、全体検証と候補生成を再実行します。

---

## 5. Windows ビルド

Windows 配布物は次で作成します。

```powershell
npm run make:win
```

生成物の出力先:

- `out/make/squirrel.windows/x64/`

確認すべきファイル:

- `WorldShotLogSetup.exe`
- `worldshot_log-<version>-full.nupkg`
- `RELEASES`

重要:

- `.exe` だけでは自動アップデート配布は成立しません
- **`.nupkg` と `RELEASES` が必須**です
- ビルド後は `npm run smoke:packaged` で配布版の起動、ASAR内のコード一致、DB/IPC、Worker出力、同梱軽量AIの実推論を確認します。検証は一時profileを使い、通常利用中のDBやサムネイルから隔離します。
- `smoke:packaged` は現在のソースと一致し、home隔離に対応したビルドを対象とします。起動だけの確認とは異なり、古いビルドは拒否します。インストーラの新規導入・更新動作は別に確認してください。

---

## 6. git タグと push

### 6-1. main の push

まず、リリース対象のコミットが `main` にあることを確認して push します。

例:

```powershell
git push -u origin main
```

### 6-2. タグ作成

バージョンに対応する注釈付きタグを作成します。

例:

```powershell
git tag -a vX.Y.Z -m "WorldShot Log vX.Y.Z"
```

### 6-3. タグ push

```powershell
git push origin vX.Y.Z
```

確認:

- GitHub 側に `v<version>` タグが存在すること
- タグが公開対象コミットを指していること

---

## 7. GitHub Release 作成

### 7-1. 基本方針

GitHub Release は、**git タグと同名**で作成します。

例:

- タグ: `vX.Y.Z`
- Release 名: `WorldShot Log vX.Y.Z`

### 7-2. 本文

推奨:

- `release-notes/vX.Y.Z.md` の内容をそのまま使う

### 7-3. 添付ファイル

`out/make/squirrel.windows/x64/` から次を添付します。

- `WorldShotLogSetup.exe`
- `worldshot_log-<version>-full.nupkg`
- `RELEASES`

### 7-4. 公開状態

重要:

- `draft` のままでは自動アップデート対象になりません
- 自動アップデートに乗せるなら、**Release を公開状態にする**必要があります
- Release 公開後、GitHub 側で `WorldShotLogSetup.exe`、`.nupkg`、`RELEASES` の 3 asset が表示されていることを確認します

---

## 8. 自動アップデート成立条件

自動アップデートが成立するための条件:

1. アプリの `version` が、公開中より新しい
2. GitHub Release のタグが `v<version>` と一致している
3. Release に `.nupkg` と `RELEASES` が添付されている
4. Release が公開済みである
5. 利用者が Windows 配布版アプリを使っている

補足:

- `npm start` のような開発実行では自動アップデートは動作しません
- Windows 配布版でのみ有効です
- GitHub Release に asset が存在していても、`update.electronjs.org` が古い情報を返している場合があります
- リリース完了前に、旧バージョンから新バージョンが返ることを必ず確認してください

### 8-1. 更新エンドポイント確認

Release 公開後、旧バージョンから最新版が返ることを確認します。

例:

```powershell
$owner = 'noma-nomoa'
$repo = 'vrchat-world-photo-manager'
$oldVersion = '2.4.1'
$newVersion = '2.4.2'

$updateJsonUrl = "https://update.electronjs.org/$owner/$repo/win32-x64/$oldVersion"
$releasesUrl = "https://update.electronjs.org/$owner/$repo/win32-x64/$oldVersion/RELEASES?id=worldshot_log&localVersion=$oldVersion&arch=amd64"

Invoke-WebRequest $updateJsonUrl -Headers @{ 'User-Agent' = "WorldShot Log/$oldVersion" }
Invoke-WebRequest $releasesUrl -Headers @{ 'User-Agent' = "WorldShot Log/$oldVersion" }
```

確認ポイント:

- `$updateJsonUrl` が `204 No Content` ではなく、`WorldShot Log v<newVersion>` を返す
- `$releasesUrl` が `worldshot_log-<newVersion>-full.nupkg` を指している
- 新バージョン自身から確認した場合は `204 No Content` になる

例:

```powershell
Invoke-WebRequest "https://update.electronjs.org/$owner/$repo/win32-x64/$newVersion" `
  -Headers @{ 'User-Agent' = "WorldShot Log/$newVersion" }
```

### 8-2. update.electronjs.org のキャッシュに注意

`update.electronjs.org` は GitHub Releases を参照しますが、内部キャッシュの影響で一時的に古い Release を返すことがあります。

確認時に古い結果が返る場合は、次を確認します。

- 最新 Release が draft / prerelease になっていないか
- 最新 Release に `.nupkg` と `RELEASES` が添付されているか
- `RELEASES` の中身が最新 `.nupkg` を指しているか
- クエリ付き URL でも結果を確認し、キャッシュが更新されるか

例:

```powershell
$stamp = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
Invoke-WebRequest "$updateJsonUrl?bust=$stamp" `
  -Headers @{ 'User-Agent' = "WorldShot Log/$oldVersion" }
```

この確認が通るまで、リリース完了扱いにしません。

---

## 9. 推奨リリース手順

実務上は次の順番が安全です。

1. 必要なコード・資料更新を完了する
2. `package.json` / `package-lock.json` / `src/index.html` のバージョンを上げる
3. `release-notes/vX.Y.Z.md` を作成する
4. 動作確認・構文確認を行う
5. `main` に commit / push する
6. `npm run make:win` で Windows ビルドを作る
7. `git tag -a vX.Y.Z -m "WorldShot Log vX.Y.Z"` を作る
8. タグを push する
9. GitHub Release を作成する
10. `WorldShotLogSetup.exe`、`.nupkg`、`RELEASES` を添付する
11. release-notes の本文を貼る
12. 公開状態にする
13. `update.electronjs.org` の通常URLで旧バージョンから最新版が返ることを確認する
14. 最新版自身からの更新確認が `204 No Content` になることを確認する

---

## 10. リリース時の更新候補

毎回必須ではありませんが、必要に応じて更新します。

- `README.md`
- `release-notes/vX.Y.Z.md`
- GitHub Release 本文

特に、ユーザー向けの見え方が変わったときは README も更新してください。

---

## 11. よくあるミス

### 11-1. `.exe` だけ公開してしまう

これは自動アップデートに必要な構成が足りません。
必ず `.nupkg` と `RELEASES` も公開してください。

### 11-2. タグ名と version がずれる

例:

- `package.json` は `2.2.1`
- タグは `v2.2.0`

この状態は不正です。タグと version を一致させてください。

### 11-3. draft のまま公開したつもりになる

draft Release はアップデート対象になりません。
公開状態を必ず確認してください。

### 11-4. リリースノートを更新し忘れる

後から差分が追いにくくなります。
毎回 `release-notes/` に残す運用を推奨します。

### 11-5. 更新エンドポイントを確認しない

GitHub Release に配布物があっても、更新サービスが古いキャッシュを返していると自動更新が走りません。

旧バージョンからの確認で、次のどちらかになっている場合はリリース完了にしないでください。

- `204 No Content`
- `RELEASES` が前バージョンの `.nupkg` を指している

---

## 12. 現在の公開版

2026年10月25日に GitHub Releases API と自動更新サービスで確認した公開版:

- [v2.5.4](https://github.com/noma-nomoa/vrchat-world-photo-manager/releases/tag/v2.5.4)

対応する変更点ファイル:

- `release-notes/v2.5.4.md`

全体回帰20項目、配布版の機能確認、新規導入・Setup上書き・Squirrel更新の3経路が成功しています。旧版v2.5.3から最新版の配信、最新版自身では更新なし（HTTP 204）も確認済みです。[今回の検証記録](./docs/release-preparation-v2.5.4.md)を参照してください。

---

## 13. 最後に

リリース作業は、単にビルドを作るだけではなく、**version・タグ・配布物・Release 公開状態を揃えて初めて完了**です。

迷ったときは、次の順番で確認してください。

1. `package.json` の version
2. git タグ
3. `out/make/squirrel.windows/x64/` の成果物
4. GitHub Release の本文と添付
5. Release が draft でないこと
6. `update.electronjs.org` が旧バージョンから最新版を返すこと
