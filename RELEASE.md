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

必要に応じて、構文確認も行います。

例:

```powershell
node --check src/index.js src/preload.js src/renderer.js src/i18n.js src/photo-editor-worker.js
npm run audit:i18n
npm run smoke:data
```

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
- ビルド後は `npm run smoke:packaged` で配布版の起動確認を行います

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

現時点の公開版:

- `v2.4.2`

対応する変更点ファイル:

- `release-notes/v2.4.2.md`

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
