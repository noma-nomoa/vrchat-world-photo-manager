# WorldShot Log リリース手順

## リリース手順

1. `package.json` の `version` を更新する
2. アプリの動作確認を行う
3. `main` にコミットして push する
4. Windows 配布物を作成する

```powershell
npm run make:win
```

5. バージョンに対応する git タグを作成する

```powershell
git tag -a v1.0.1 -m "WorldShot Log v1.0.1"
git push origin v1.0.1
```

6. 同じタグ名で GitHub Release を作成する
7. `out/make/squirrel.windows/x64/` から以下を添付する
   - `WorldShotLogSetup.exe`
   - `worldshot_log-<version>-full.nupkg`
   - `RELEASES`
8. Release を公開する  
   `draft` のままでは自動アップデート対象にならない

## 自動アップデートの条件

- アプリの `version` は、現在公開中のものより大きい必要がある
- GitHub Release のタグは、アプリの version と一致している必要がある  
  例: `v1.0.1`
- `RELEASES` と `.nupkg` は必須  
  `.exe` だけでは自動アップデートできない
- 自動アップデートは、Windows の配布版アプリでのみ動作する
- `npm start` などの開発実行では動作しない

## 現在の命名ルール

- アプリ名: `WorldShot Log`
- タグ形式: `v<version>`
- Windows インストーラ名: `WorldShotLogSetup.exe`
