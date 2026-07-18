(() => {
  const LANGUAGE_STORAGE_KEY = 'vrchat-world-photo-manager-language';
  const SUPPORTED_LANGUAGES = Object.freeze({
    ja: { label: '日本語', htmlLang: 'ja' },
    en: { label: 'English', htmlLang: 'en' },
    ko: { label: '한국어', htmlLang: 'ko' },
  });

  const DICTIONARIES = Object.freeze({
    en: {
      '設定を開く': 'Open settings',
      'テーマを切り替える': 'Toggle theme',
      '更新': 'Refresh',
      'サムネイル再生成': 'Regenerate thumbnails',
      'まだ取り込みはありません': 'No photos imported yet',
      '処理を準備中...': 'Preparing...',
      '処理準備中...': 'Preparing...',
      '処理中...': 'Processing...',
      '年月': 'Timeline',
      '年を押すと年一覧、矢印で月を開閉します':
        'Select a year to browse it, or use arrows to expand months.',
      '写真一覧': 'Photo list',
      'お気に入りのみ表示': 'Show favorites only',
      'お気に入りのみ表示中': 'Showing favorites only',
      '並び順: 新しい順': 'Sort: Newest first',
      '並び順: 古い順': 'Sort: Oldest first',
      '表示サイズ: 標準': 'Size: Standard',
      '表示サイズ: コンパクト': 'Size: Compact',
      '向き: すべて': 'Orientation: All',
      '向き: 横長': 'Orientation: Landscape',
      '向き: 縦長': 'Orientation: Portrait',
      '向き: 正方形': 'Orientation: Square',
      '向きフィルタ': 'Orientation filter',
      '向きフィルタ: すべて': 'Orientation: All',
      '向きフィルタ: 横長': 'Orientation: Landscape',
      '向きフィルタ: 縦長': 'Orientation: Portrait',
      '向きフィルタ: 正方形': 'Orientation: Square',
      '向き': 'Orientation',
      'すべて': 'All',
      '横長': 'Landscape',
      '縦長': 'Portrait',
      '正方形': 'Square',
      'ラベル: すべて': 'Labels: All',
      'ラベルフィルタ: すべて': 'Labels: All',
      'ラベルフィルタ': 'Label filter',
      'ラベル': 'Labels',
      '検索': 'Search',
      '検索を実行': 'Run search',
      '検索をクリア': 'Clear search',
      'クリア': 'Clear',
      'World名を入力': 'Enter world name',
      'メモを入力': 'Enter memo',
      'プリントのノートを入力': 'Enter print note',
      '選択': 'Select',
      'お気に入り': 'Favorite',
      'お気に入り解除': 'Unfavorite',
      '削除': 'Delete',
      '削除する': 'Delete',
      '画像/フォルダをここにドラッグ＆ドロップ':
        'Drag photos or folders here',
      'まだ写真がありません': 'No photos yet',
      'まだ写真がありません。画像 / フォルダをドラッグ&ドロップするか、設定から取り込めます':
        'No photos yet. Drag photos or folders here, or import them from Settings.',
      '表示する年または月を選択してください':
        'Select a year or month to view photos.',
      'このワールドの写真はまだありません':
        'No photos for this world yet.',
      '該当する写真はありません': 'No matching photos.',
      'この年の写真はまだありません': 'No photos for this year yet.',
      'この月の写真はまだありません': 'No photos for this month yet.',
      'お気に入り に一致する写真はありません':
        'No photos match Favorites.',
      '年月一覧': 'Timeline',
      'ワールド一覧': 'Worlds',
      'ワールド一覧を表示': 'Show world list',
      '年月一覧へ戻る': 'Back to timeline',
      '撮影枚数順': 'Photo count',
      '名前順': 'Name',
      'ワールド': 'Worlds',
      'ワールド情報付きの写真はまだありません':
        'No photos with world info yet.',
      'まだ取り込みがありません': 'Nothing imported yet',
      '日付不明': 'Unknown date',
      '日時不明': 'Unknown date/time',
      '時刻不明': 'Unknown time',
      'サムネイル未生成': 'Thumbnail not generated',
      'サムネイル要再生成': 'Thumbnail needs regeneration',

      'VRChatで開く': 'Open in VRChat',
      '元画像を開く': 'Open original image',
      '画像を開く': 'Open image',
      '保存先フォルダを開く': 'Open containing folder',
      '保存先を開く': 'Open folder',
      '画像を加工する': 'Edit image',
      '撮影日時': 'Taken at',
      '解像度': 'Resolution',
      'ワールド名未取得': 'World name unavailable',
      'ワールド名を取得できませんでした': 'Could not get world name',
      'ファイル名': 'File name',
      'ファイル名不明': 'Unknown file name',
      '未取得': 'Unavailable',
      '未設定': 'Not set',
      'Description': 'Description',
      'Tag': 'Tags',
      'メモ': 'Memo',
      '編集': 'Edit',
      '保存': 'Save',
      '自由にメモを残せます': 'Add a free memo',
      '現在の表示名': 'Current display name',
      'お気に入り切り替え': 'Toggle favorite',
      'カードを編集': 'Edit card',
      'この登録を削除': 'Delete this entry',
      '前の画像': 'Previous image',
      '次の画像': 'Next image',
      'プリント': 'Print',
      'プリントのノート': 'Print note',

      '画像編集': 'Image Editor',
      '元に戻す': 'Undo',
      'やり直す': 'Redo',
      '編集内容をリセット': 'Reset edits',
      '編集前と比較': 'Compare with original',
      '編集後の表示に戻す': 'Back to edited preview',
      '比較': 'Compare',
      '編集中': 'Editing',
      '別名で保存': 'Save as',
      '保存中...': 'Saving...',
      '切り抜き': 'Crop',
      'オリジナル': 'Original',
      '回転': 'Rotate',
      '自由回転': 'Free rotate',
      '左右反転': 'Flip horizontal',
      '上下反転': 'Flip vertical',
      'ズーム': 'Zoom',
      '左 右位置': 'Horizontal position',
      '左右位置': 'Horizontal position',
      '上下位置': 'Vertical position',
      'アバターサムネイル': 'Avatar thumbnail',
      '長辺 800px': 'Long edge 800px',
      'VRCギャラリー': 'VRC Gallery',
      '絵文字・ステッカー': 'Emoji / Sticker',
      'プリセット': 'Presets',
      '補正の強さ': 'Strength',
      'プリセット名': 'Preset name',
      'プリセット名を入力': 'Preset name',
      'この設定を保存': 'Save preset',
      '補正': 'Adjustments',
      '明るさ': 'Brightness',
      '露出': 'Exposure',
      'コントラスト': 'Contrast',
      'ハイライト': 'Highlights',
      'シャドウ': 'Shadows',
      'ホワイト': 'Whites',
      'ブラック': 'Blacks',
      'ガンマ': 'Gamma',
      '色温度': 'Temperature',
      '色合い': 'Tint',
      '彩度': 'Saturation',
      '自然な彩度': 'Vibrance',
      '明瞭度': 'Clarity',
      'テクスチャ': 'Texture',
      'シャープ': 'Sharpen',
      'ノイズ低減': 'Denoise',
      'フェード': 'Fade',
      '粒子': 'Grain',
      'ビネット': 'Vignette',
      'トーンカーブ': 'Tone Curve',
      '全体': 'Overall',
      'ぼかし': 'Blur',
      'ぼかしをリセット': 'Reset blur',
      'どこをぼかしにするか': 'Blur target',
      '全体ぼかし': 'Full blur',
      '放射ぼかし': 'Radial blur',
      '放射': 'Radial',
      '全体': 'Overall',
      'ぼかし量': 'Blur amount',
      '範囲': 'Area',
      '四角で範囲選択': 'Select rectangle area',
      '丸で範囲選択': 'Select ellipse area',
      'フリーハンドで範囲選択': 'Select freehand area',
      'ぼかしの濃さ': 'Blur strength',
      'モザイクの濃さ': 'Mosaic strength',
      '塗りつぶしの濃さ': 'Fill strength',
      '色': 'Color',
      '確定': 'Confirm',
      '範囲を編集': 'Edit area',
      '目隠し加工': 'Hide / Mask',
      '目隠し加工をリセット': 'Reset hide / mask',
      'モザイク': 'Mosaic',
      '塗りつぶし': 'Fill',
      'ひとつ戻す': 'Undo one',
      '1つ戻す': 'Undo one',
      '書き出し': 'Export',
      '書き出し設定をリセット': 'Reset export settings',
      '形式': 'Format',
      'サイズ': 'Size',
      '長辺 3840px': 'Long edge 3840px',
      '長辺 2560px': 'Long edge 2560px',
      '長辺 2048px': 'Long edge 2048px',
      '長辺 1600px': 'Long edge 1600px',
      '長辺 1200px': 'Long edge 1200px',
      '長辺 1024px': 'Long edge 1024px',
      '品質': 'Quality',
      'テキスト追加': 'Add Text',
      'テキスト': 'Text',
      'テキストをリセット': 'Reset text',
      'テキストを追加': 'Add text',
      'テキストを削除': 'Delete text',
      '追加': 'Add',
      '文字': 'Text',
      '内容': 'Content',
      'テキストを入力': 'Enter text',
      'フォント': 'Font',
      '太さ': 'Weight',
      '縁': 'Stroke',
      '縁の色': 'Stroke color',
      '縁の種類': 'Stroke type',
      'なし': 'None',
      '縁取り': 'Outline',
      '影': 'Shadow',
      '発光': 'Glow',
      '縁内を透過': 'Transparent fill',
      '適用範囲': 'Apply to',
      '被写体': 'Subject',
      '被写体のみ': 'Subject only',
      '背景のみ': 'Background only',
      '文字サイズ': 'Text size',
      '縁の太さ': 'Stroke width',
      '文字間隔': 'Letter spacing',
      'AI被写体': 'AI Subject',
      'AI被写体選択': 'AI Subject Selection',
      '背景透過で保存': 'Save transparent',
      '背景透過PNGを保存中...': 'Saving transparent PNG...',
      '背景透過で保存するにはAI被写体選択のマスクが必要です':
        'Create an AI subject selection mask before saving with a transparent background.',
      '被写体マスクを使って背景を透過したPNGとして保存します':
        'Save a PNG with the background made transparent using the subject mask.',
      'AI被写体選択でマスクを作成すると使用できます':
        'Create a mask with AI Subject Selection to use this.',
      '背景透過用の描画を準備できませんでした':
        'Could not prepare the transparent background render.',
      '背景透過に使う被写体マスクを読み込めませんでした':
        'Could not load the subject mask for transparent background saving.',
      '補正対象': 'Adjustment target',
      '被写体マスクをリセット': 'Reset subject mask',
      'マスク未作成': 'No mask yet',
      'ダミーマスク': 'Dummy mask',
      '軽量AIマスク': 'Lightweight AI mask',
      '標準AIマスク': 'Standard AI mask',
      '高精度AIマスク': 'High quality AI mask',
      '読み込みマスク': 'Imported mask',
      'マスク': 'Mask',
      '被写体を自動選択': 'Auto select subject',
      '標準AIをダウンロード': 'Download standard AI',
      '高精度AIをダウンロード': 'Download high quality AI',
      '高精度で再検出': 'Detect with high quality',
      'マスク画像を読み込み': 'Load mask image',
      'マスクを表示': 'Show mask',
      'マスク反転': 'Invert mask',
      'マスク修正': 'Mask correction',
      '修正する': 'Correct',
      '消す': 'Erase',
      'ブラシサイズ': 'Brush size',
      '濃度': 'Strength',
      'スポイト': 'Eyedropper',
      'マスク画像を読み込み中です': 'Loading mask image...',
      '現在: 標準AIで実行されます': 'Current: Standard AI will be used',
      '現在: 軽量AIで実行されます': 'Current: Lightweight AI will be used',
      '現在: 軽量AIで実行されます / 標準AIをダウンロードすると検出精度が向上します':
        'Current: Lightweight AI will be used / Download Standard AI for better detection.',
      'AIモデルを利用できません': 'AI model is unavailable',
      '設定から高精度AIモデルをダウンロードしてください':
        'Download the high quality AI model from settings.',
      'AIモデル情報を取得できません': 'Could not get AI model information',
      'ONNX Runtimeを読み込めませんでした': 'Could not load ONNX Runtime',
      'AIマスクの出力を読み取れませんでした':
        'Could not read the AI mask output',
      '標準AIの実行機能を利用できません':
        'Standard AI execution is unavailable',
      'AIモデルの実行に失敗しました': 'AI model execution failed',
      '画像サイズを取得できませんでした': 'Could not read the image size',
      'AI入力画像を作成できませんでした': 'Could not create the AI input image',
      'AIマスクの出力が空でした': 'AI mask output was empty',
      'AIマスクを作成できませんでした': 'Could not create the AI mask',
      'AIマスクを切り出せませんでした': 'Could not crop the AI mask',
      'AIマスクを元画像サイズに戻せませんでした':
        'Could not restore the AI mask to the source image size',
      '高解像度マスクを整えています...': 'Refining high resolution mask...',
      'PNG / JPEG / WebP のマスク画像を選択してください':
        'Select a PNG, JPEG, or WebP mask image.',
      'マスク画像を読み込めませんでした': 'Could not load the mask image',
      '画像オーバーレイ': 'Image Overlay',
      '画像オーバーレイをリセット': 'Reset image overlay',
      '画像オーバーレイ素材を読み込めませんでした':
        'Could not load image overlay asset',
      '画像オーバーレイ素材の追加機能を利用できません':
        'Image overlay import is unavailable',
      '画像を追加': 'Add image',
      '管理素材': 'Managed assets',
      '管理素材から削除': 'Delete from managed assets',
      '管理素材を削除': 'Delete managed asset',
      '管理素材を削除できませんでした': 'Could not delete managed asset',
      '管理素材を削除しました': 'Managed asset deleted',
      'レイヤー': 'Layers',
      '手前へ': 'Forward',
      '奥へ': 'Backward',
      '透明度': 'Opacity',
      '合成方法': 'Blend mode',
      '標準合成': 'Normal',
      '乗算': 'Multiply',
      'スクリーン': 'Screen',
      'オーバーレイ': 'Overlay',
      'ソフトライト': 'Soft Light',
      'ハードライト': 'Hard Light',
      '比較(暗)': 'Darken',
      '比較(明)': 'Lighten',
      '最近使用': 'Recent',
      'システム': 'System',
      '標準': 'Regular',
      '中太': 'Medium',
      'セミボールド': 'Semibold',
      '太字': 'Bold',
      '特太': 'Extra bold',
      '極太': 'Black',
      'ルーラー': 'Ruler',
      '三分割グリッド': 'Rule of thirds',

      '✨ 自動補正': '✨ Smart Auto',
      '学習補正': 'Learned Auto',
      '投稿クリア': 'Post Clear',
      '自然クリア': 'Natural Clear',
      '夜景強調': 'Night Boost',
      'ネオン強調': 'Neon Boost',
      'ふんわり1': 'Soft 1',
      'ふんわり2': 'Soft 2',
      'ふんわり3': 'Soft 3',
      'フィルム風': 'Film Look',
      '高コントラスト1': 'High Contrast 1',
      '高コントラスト2': 'High Contrast 2',
      'クールブルー1': 'Cool Blue 1',
      'クールブルー2': 'Cool Blue 2',
      'スイートピンク1': 'Sweet Pink 1',
      'スイートピンク2': 'Sweet Pink 2',
      'スイートピンク3': 'Sweet Pink 3',
      'エアリーホワイト1': 'Airy White 1',
      'エアリーホワイト2': 'Airy White 2',
      'エアリーホワイト3': 'Airy White 3',
      'エアリーホワイト4': 'Airy White 4',
      '暗部クリア': 'Shadow Clear',
      'サムネ強調': 'Thumbnail Pop',
      'モノクロ': 'Monochrome',

      '設定': 'Settings',
      '設定モーダルを閉じる': 'Close settings',
      '言語設定': 'Language',
      '日本語': '日本語',
      '英語': 'English',
      '韓国語': '한국어',
      'フォント': 'Font',
      '更新対象フォルダの設定': 'Tracked folders',
      'フォルダ追加': 'Add folder',
      'まだ登録されていません': 'Nothing registered yet',
      'AIモデル': 'AI models',
      '画像編集の被写体選択で使うモデルを管理します':
        'Manage models used for subject selection in the image editor.',
      'AIモデル情報を取得できませんでした':
        'Could not get AI model information',
      'AIモデルのダウンロード機能を利用できません':
        'AI model download is unavailable',
      'このAIモデルはダウンロードできません':
        'This AI model cannot be downloaded',
      '画像は外部サーバーに送信されません。AI処理はPC内で実行されます。':
        'Images are not sent to external servers. AI processing runs on this PC.',
      'このモデルは標準AIより処理時間やメモリ使用量が大きくなる場合があります。':
        'This model may take longer and use more memory than Standard AI.',
      'AIモデルをダウンロードできませんでした':
        'Could not download the AI model',
      '保存場所を開けません': 'Could not open the folder',
      '保存場所を開けませんでした': 'Could not open the folder',
      'AIモデルを削除': 'Delete AI model',
      'AIモデルの削除機能を利用できません':
        'AI model deletion is unavailable',
      'AIモデルを削除しています...': 'Deleting AI model...',
      'AIモデルを削除できませんでした': 'Could not delete the AI model',
      'AIモデルを削除しました': 'AI model deleted',
      'ライセンス確認により利用停止中です':
        'Disabled after license review',
      '軽量AI': 'Lightweight AI',
      '標準AI': 'Standard AI',
      '高精度AI': 'High quality AI',
      '高精度AI候補': 'High quality AI candidate',
      '旧モデル': 'Legacy model',
      '利用可能': 'Ready',
      '検証失敗': 'Verification failed',
      '準備中': 'Preparing',
      '利用停止中': 'Disabled',
      '一部不足': 'Incomplete',
      '未ダウンロード': 'Not downloaded',
      '同梱': 'Bundled',
      '管理フォルダ': 'Managed folder',
      '同梱モデル': 'Bundled model',
      'ダウンロード': 'Download',
      '配布元': 'Source',
      '保存場所': 'Folder',
      '未確定': 'TBD',
      'すぐ使える軽量な被写体マスク生成モデルです。複雑な髪型、衣装、羽、尻尾などでは精度が落ちる場合があります。':
        'A lightweight subject mask model ready to use immediately. Accuracy may drop for complex hair, outfits, wings, or tails.',
      'すぐ使える高速モデルです。複雑な髪型・衣装・羽・尻尾などでは精度が落ちる場合があります。':
        'A fast model ready to use immediately. Accuracy may drop for complex hair, outfits, wings, or tails.',
      'VRChatアバター向けの通常モデルです。細かい髪型・衣装・装飾・羽・尻尾などの検出精度が向上します。':
        'The everyday model for VRChat avatars. It improves detection of detailed hair, outfits, accessories, wings, and tails.',
      'VRChatアバター向けの標準モデルです。軽量AIよりも、髪型・衣装・装飾・羽・尻尾などの検出精度向上を狙います。':
        'The standard model for VRChat avatars. It aims to improve detection of hair, outfits, accessories, wings, and tails compared with Lightweight AI.',
      '境界をよりきれいに検出する上位モデルです。背景ぼかし、被写体だけ補正、文字や画像を被写体に合わせる加工に向いています。':
        'An advanced model that detects cleaner boundaries. Useful for background blur, subject-only adjustments, and aligning text or images with the subject.',
      '境界をよりきれいに検出する上位モデルです。背景ぼかし、被写体だけ補正、文字や画像を被写体に合わせる加工に向いています。処理時間やメモリ使用量が大きくなる場合があります。':
        'An advanced model that detects cleaner boundaries. Useful for background blur, subject-only adjustments, and aligning text or images with the subject. Processing time and memory use may be higher.',
      '学習データセットの商用利用条件を確認した結果、現在の推奨構成では使用しません。必要に応じて削除できます。':
        'Disabled in the current recommended setup after reviewing the commercial-use terms of the training dataset. You can delete it if needed.',
      '学習データセットの商用利用条件を確認した結果、現在の推奨構成では使用しません。':
        'Disabled in the current recommended setup after reviewing the commercial-use terms of the training dataset.',
      '旧構成で高精度AIとして扱っていたモデルです。現在の推奨構成では使用しません。':
        'A legacy model that used to be treated as high quality AI. It is not used in the current recommended setup.',
      'データ管理': 'Data management',
      'ラベル、メモ、お気に入り、World情報を保存・書き出しできます':
        'Back up and export labels, memos, favorites, and world information.',
      'データ管理を開く': 'Open data management',
      '操作メニュー': 'Actions',
      'バックアップを作成': 'Create backup',
      '状態チェック': 'Health check',
      '元画像なしを表示': 'Show missing originals',
      'サムネイルなしを表示': 'Show missing thumbnails',
      'World情報未取得を表示': 'Show missing world info',
      'World要確認を表示': 'Show world issues',
      '欠損サムネイルを再生成': 'Regenerate missing thumbnails',
      'World要確認を再取得': 'Refresh world issues',
      'バックアップから復元': 'Restore from backup',
      'CSVエクスポート': 'Export CSV',
      'JSONエクスポート': 'Export JSON',
      'メンテナンス': 'Maintenance',
      '表示中の月を削除': 'Delete current month',
      'サムネイルキャッシュを削除': 'Clear thumbnail cache',
      '既存画像の情報を再取り込み': 'Reimport registered photos',
      '全登録を削除': 'Delete all entries',
      'DBを初期化': 'Reset database',
      'アンインストール': 'Uninstall',
      'アプリ本体を削除します。必要に応じて、保存済みデータも一緒に削除できます。':
        'Remove the app. You can also delete saved data if needed.',
      'データも削除してアンインストール': 'Uninstall and delete data',
      '確認': 'Confirm',
      'この操作を実行しますか？': 'Run this action?',
      'キャンセル': 'Cancel',
      '実行する': 'Run',
      '閉じる': 'Close',
      '概要': 'Overview',
      '写真': 'Photos',
      'フォルダ': 'Folders',
      'ワールド数': 'Worlds',
      '背景': 'Background',
      '画像を選択': 'Choose image',
      '一覧を表示': 'Show list',
      '更新対象フォルダ一覧': 'Tracked folder list',
      'サムネイル再生成の対象月': 'Month for thumbnail regeneration',
      '情報再取り込みの対象月': 'Month for reimport',
      '再生成する月を選択': 'Select month to regenerate',
      '再取り込みする月を選択': 'Select month to reimport',
      '対象月がありません': 'No target month',
      'ラベルを設定': 'Set labels',
      '既存ラベルを再利用したり、新しいラベルを色付きで追加して写真ごとに設定できます。':
        'Reuse existing labels or create colored labels for each photo.',
      '現在のラベル': 'Current labels',
      'ラベルを設定する': 'Choose labels',
      '既存ラベルを選択': 'Choose existing label',
      'この月にはラベルがありません': 'No labels for this month',
      'ラベルを作成する': 'Create label',
      'ラベルの名前を入力してください': 'Enter a label name',
      '色を選択': 'Choose color',
      'この内容で追加': 'Add this label',
      '追加できるラベルはありません': 'No labels available',
      'ラベルはまだ設定されていません': 'No labels set yet',
      '読み込み中...': 'Loading...',
      'ワールド名を編集': 'Edit world name',
      'ワールド名': 'World name',
      '手動設定を解除': 'Clear manual setting',
      '再度ワールド名を自動取得': 'Fetch world name again',
      'World情報を再読み込み': 'Reload world info',
      '特殊文字などで自動取得名が崩れる場合に、表示名を手動で上書きできます。':
        'If the detected name breaks because of special characters, you can override it manually.',
      'ここに画像またはフォルダをドロップ': 'Drop photos or folders here',
      'png / jpg / jpeg / webp に対応': 'Supports png / jpg / jpeg / webp',
      '/ 比較中': '/ Comparing',
      '/ 未確定: 1件': '/ Uncommitted: 1',
      'World: すべて': 'World: All',
      'Worldフィルタ: すべて': 'World filter: All',
      'Worldメタデータ要確認の該当分だけを再取得':
        'Re-fetch only photos needing World metadata review',
      'Worldメタデータ要確認の写真だけを表示':
        'Show only photos needing World metadata review',
      'World情報の自動同期を開始できませんでした':
        'Could not start automatic World info sync',
      'World情報未取得の写真だけを表示':
        'Show only photos without World info',
      'World設定を保存しました': 'World settings saved',
      'World名フィルタ': 'World name filter',
      'アップデートがあります': 'Update available',
      'アップデートの準備ができました': 'Update ready',
      'アップデートは準備済みです。あとで再起動して適用できます':
        'The update is ready. Restart later to apply it.',
      'アップデートは保留しました': 'Update postponed',
      'アップデートを開始できませんでした': 'Could not start update',
      'アップデートを適用できませんでした': 'Could not apply update',
      'アプリデータをJSONでバックアップ': 'Back up app data as JSON',
      'アンインストールの確認を開く': 'Open uninstall confirmation',
      'アンインストールモーダルを閉じる': 'Close uninstall dialog',
      'アンインストールを開始します': 'Starting uninstall',
      'アンインストールを開始できませんでした': 'Could not start uninstall',
      'お気に入りを切り替え': 'Toggle favorite',
      'このフォルダを削除': 'Delete this folder',
      'サムネイルなしの写真だけを表示':
        'Show only photos without thumbnails',
      'サムネイルを再生成しています...': 'Regenerating thumbnails...',
      'サムネイルを再生成する月を選択':
        'Choose the month to regenerate thumbnails',
      'サムネイル再生成が完了しました':
        'Thumbnail regeneration complete',
      'すべての登録を削除': 'Delete all entries',
      'データ管理を閉じる': 'Close data management',
      'データ削除とアンインストールを開始します':
        'Starting data deletion and uninstall',
      'ドロップされた項目に対応画像がありませんでした':
        'No supported photos were found in the dropped items',
      'バックアップJSONから復元': 'Restore from backup JSON',
      'フォルダ内に対応画像がありませんでした':
        'No supported photos were found in the folder',
      'プリセットを保存できませんでした': 'Could not save preset',
      'プリセット名を入力してください': 'Enter a preset name',
      'プレビュー補助': 'Preview guides',
      'メモを保存しました': 'Memo saved',
      'ラベルの保存に失敗しました': 'Could not save label',
      'ラベルを読み込み中...': 'Loading labels...',
      'ラベルを保存しました': 'Label saved',
      'ラベル名を確認してください': 'Check the label name',
      'ラベル名を入力してください': 'Enter a label name',
      'ワールド名を手動修正': 'Edit world name manually',
      '右に90度回転': 'Rotate 90 degrees right',
      '画像パスを再検出して更新しました':
        'Image paths were rediscovered and updated',
      '画像を読み込めませんでした': 'Could not load image',
      '画像処理Workerがタイムアウトしました':
        'Image processing worker timed out',
      '画像処理Workerで失敗しました': 'Image processing worker failed',
      '画像処理Workerを起動できませんでした':
        'Could not start image processing worker',
      '画像処理Workerを利用できません':
        'Image processing worker is unavailable',
      '画像処理プレビューを更新しました':
        'Image processing preview updated',
      '画像読み込みがキャンセルされました': 'Image loading canceled',
      '画像編集モーダルを閉じました': 'Image editor closed',
      '画像編集モーダルを閉じる': 'Close image editor',
      '拡大画像': 'Enlarged photo',
      '確認モーダルを閉じる': 'Close confirmation dialog',
      '学習補正を適用しました': 'Learning correction applied',
      '管理サムネイルとサムネイル参照を削除':
        'Delete managed thumbnails and thumbnail references',
      '欠損しているサムネイルだけを再生成':
        'Regenerate only missing thumbnails',
      '月を選択すると利用できます': 'Select a month to use this',
      '検索対象を切り替え': 'Switch search target',
      '元画像なしの写真だけを表示':
        'Show only photos without original files',
      '現在の絞り込みでは、この日の写真は表示されていません':
        'No photos from this day are shown with the current filters',
      '更新に失敗しました': 'Update failed',
      '更新はキャンセルされました': 'Update canceled',
      '更新確認: 新規0件': 'Update check: 0 new',
      '更新対象のフォルダがまだ登録されていません':
        'No folders are registered for updates yet',
      '更新対象フォルダがまだありません': 'No tracked folders yet',
      '更新対象フォルダを追加': 'Add tracked folder',
      '更新対象フォルダ一覧を表示': 'Show tracked folder list',
      '更新対象フォルダ一覧を閉じる': 'Close tracked folder list',
      '今すぐ更新': 'Update now',
      '左に90度回転': 'Rotate 90 degrees left',
      '再起動して更新': 'Restart to update',
      '再取り込みできる月がありません': 'No month can be reimported',
      '再生成できる月がありません':
        'No month can regenerate thumbnails',
      '最新バージョン': 'Latest version',
      '削除してアンインストール': 'Delete and uninstall',
      '削除するサムネイルがありません': 'No thumbnails to delete',
      '削除する登録がありません': 'No entries to delete',
      '自動同期中は再読み込みできません':
        'Cannot reload during automatic sync',
      '自動補正を解除しました': 'Auto correction cleared',
      '自動補正を適用しました': 'Auto correction applied',
      '写真を選択すると利用できます': 'Select a photo to use this',
      '主な更新内容': 'Main updates',
      '写真一覧をCSVで書き出し': 'Export photo list as CSV',
      '写真一覧をJSONで書き出し': 'Export photo list as JSON',
      '取り込み中です。処理が終わってから次の取り込みを開始してください':
        'Import is running. Start the next import after it finishes.',
      '処理中はアンインストールを開始できません':
        'Cannot start uninstall while processing',
      '処理中はエクスポートできません':
        'Cannot export while processing',
      '処理中はバックアップできません':
        'Cannot back up while processing',
      '処理中はフォルダを削除できません':
        'Cannot delete folders while processing',
      '処理中はフォルダを追加できません':
        'Cannot add folders while processing',
      '処理中は再取得できません': 'Cannot re-fetch while processing',
      '処理中は再生成できません':
        'Cannot regenerate thumbnails while processing',
      '処理中は状態チェックできません':
        'Cannot check status while processing',
      '処理中は抽出できません': 'Cannot extract while processing',
      '処理中は復元できません': 'Cannot restore while processing',
      '初期化するデータがありません': 'No data to reset',
      '情報を再取り込みする月を選択':
        'Choose the month to reimport information',
      '選択を切り替え': 'Toggle selection',
      '選択中の月のサムネイルを再生成':
        'Regenerate thumbnails for the selected month',
      '選択中の月の登録画像を現在の解析ロジックで再取り込み':
        'Reimport registered photos in the selected month with the current parser',
      '全期間': 'All time',
      '追跡フォルダを確認しました': 'Tracked folders checked',
      '追跡フォルダを更新しました': 'Tracked folders updated',
      '追跡フォルダを更新中...': 'Updating tracked folders...',
      '登録・キャッシュ・更新対象フォルダを初期化':
        'Reset entries, cache, and tracked folders',
      '登録されたフォルダがありません': 'No registered folders',
      '登録データの状態をチェック': 'Check registered data status',
      '濃さ': 'Strength',
      '背景画像をクリアしました': 'Background image cleared',
      '背景画像を更新しました': 'Background image updated',
      '編集結果を描画できませんでした':
        'Could not render edited result',
      '保存しました': 'Saved',
      '保存に失敗しました': 'Save failed',
      '保存をキャンセルしました': 'Save canceled',
      '保存機能を利用できません': 'Save is unavailable',
      '保存済みプリセットがないためスマート自動補正を適用しました':
        'No saved presets found, so smart auto correction was applied',
      '本当に削除しますか？': 'Are you sure you want to delete?',
      '例: Waiting / 待ち': 'Example: Waiting / Break',
      '時刻不明': 'Unknown time',
    },
    ko: {
      '設定を開く': '설정 열기',
      'テーマを切り替える': '테마 전환',
      '更新': '새로고침',
      'サムネイル再生成': '썸네일 다시 생성',
      'まだ取り込みはありません': '아직 가져온 사진이 없습니다',
      '処理を準備中...': '준비 중...',
      '処理準備中...': '준비 중...',
      '処理中...': '처리 중...',
      '年月': '날짜',
      '年を押すと年一覧、矢印で月を開閉します':
        '연도를 선택해 목록을 보고, 화살표로 월을 열고 닫습니다.',
      '写真一覧': '사진 목록',
      'お気に入りのみ表示': '즐겨찾기만 표시',
      'お気に入りのみ表示中': '즐겨찾기만 표시 중',
      '並び順: 新しい順': '정렬: 최신순',
      '並び順: 古い順': '정렬: 오래된 순',
      '表示サイズ: 標準': '표시 크기: 기본',
      '表示サイズ: コンパクト': '표시 크기: 컴팩트',
      '向き: すべて': '방향: 전체',
      '向き: 横長': '방향: 가로',
      '向き: 縦長': '방향: 세로',
      '向き: 正方形': '방향: 정사각형',
      '向きフィルタ': '방향 필터',
      '向きフィルタ: すべて': '방향: 전체',
      '向きフィルタ: 横長': '방향: 가로',
      '向きフィルタ: 縦長': '방향: 세로',
      '向きフィルタ: 正方形': '방향: 정사각형',
      '向き': '방향',
      'すべて': '전체',
      '横長': '가로',
      '縦長': '세로',
      '正方形': '정사각형',
      'ラベル: すべて': '라벨: 전체',
      'ラベルフィルタ: すべて': '라벨: 전체',
      'ラベルフィルタ': '라벨 필터',
      'ラベル': '라벨',
      '検索': '검색',
      '検索を実行': '검색 실행',
      '検索をクリア': '검색 지우기',
      'クリア': '지우기',
      'World名を入力': 'World 이름 입력',
      'メモを入力': '메모 입력',
      'プリントのノートを入力': '프린트 노트 입력',
      '選択': '선택',
      'お気に入り': '즐겨찾기',
      'お気に入り解除': '즐겨찾기 해제',
      '削除': '삭제',
      '削除する': '삭제',
      '画像/フォルダをここにドラッグ＆ドロップ':
        '사진 또는 폴더를 여기에 드래그 앤 드롭',
      'まだ写真がありません': '아직 사진이 없습니다',
      'まだ写真がありません。画像 / フォルダをドラッグ&ドロップするか、設定から取り込めます':
        '아직 사진이 없습니다. 사진/폴더를 드래그 앤 드롭하거나 설정에서 가져올 수 있습니다.',
      '表示する年または月を選択してください':
        '표시할 연도 또는 월을 선택하세요.',
      'このワールドの写真はまだありません':
        '이 월드의 사진은 아직 없습니다.',
      '該当する写真はありません': '일치하는 사진이 없습니다.',
      'この年の写真はまだありません': '이 연도의 사진은 아직 없습니다.',
      'この月の写真はまだありません': '이 월의 사진은 아직 없습니다.',
      'お気に入り に一致する写真はありません':
        '즐겨찾기에 일치하는 사진이 없습니다.',
      '年月一覧': '날짜 목록',
      'ワールド一覧': '월드 목록',
      'ワールド一覧を表示': '월드 목록 보기',
      '年月一覧へ戻る': '날짜 목록으로 돌아가기',
      '撮影枚数順': '사진 수 기준',
      '名前順': '이름순',
      'ワールド': '월드',
      'ワールド情報付きの写真はまだありません':
        '월드 정보가 있는 사진은 아직 없습니다.',
      'まだ取り込みがありません': '아직 가져온 항목이 없습니다',
      '日付不明': '날짜 알 수 없음',
      '日時不明': '일시 알 수 없음',
      '時刻不明': '시간 알 수 없음',
      'サムネイル未生成': '썸네일 미생성',
      'サムネイル要再生成': '썸네일 재생성 필요',

      'VRChatで開く': 'VRChat에서 열기',
      '元画像を開く': '원본 이미지 열기',
      '画像を開く': '이미지 열기',
      '保存先フォルダを開く': '저장 폴더 열기',
      '保存先を開く': '폴더 열기',
      '画像を加工する': '이미지 편집',
      '撮影日時': '촬영 일시',
      '解像度': '해상도',
      'ワールド名未取得': '월드 이름 없음',
      'ワールド名を取得できませんでした': '월드 이름을 가져올 수 없습니다',
      'ファイル名': '파일명',
      'ファイル名不明': '파일명 알 수 없음',
      '未取得': '없음',
      '未設定': '미설정',
      'Description': '설명',
      'Tag': '태그',
      'メモ': '메모',
      '編集': '편집',
      '保存': '저장',
      '自由にメモを残せます': '자유롭게 메모를 남길 수 있습니다',
      '現在の表示名': '현재 표시 이름',
      'お気に入り切り替え': '즐겨찾기 전환',
      'カードを編集': '카드 편집',
      'この登録を削除': '이 등록 삭제',
      '前の画像': '이전 이미지',
      '次の画像': '다음 이미지',
      'プリント': '프린트',
      'プリントのノート': '프린트 노트',

      '画像編集': '이미지 편집',
      '元に戻す': '실행 취소',
      'やり直す': '다시 실행',
      '編集内容をリセット': '편집 내용 초기화',
      '編集前と比較': '원본과 비교',
      '編集後の表示に戻す': '편집 미리보기로 돌아가기',
      '比較': '비교',
      '編集中': '편집 중',
      '別名で保存': '다른 이름으로 저장',
      '保存中...': '저장 중...',
      '切り抜き': '자르기',
      'オリジナル': '원본',
      '回転': '회전',
      '自由回転': '자유 회전',
      '左右反転': '좌우 반전',
      '上下反転': '상하 반전',
      'ズーム': '확대',
      '左 右位置': '좌우 위치',
      '左右位置': '좌우 위치',
      '上下位置': '상하 위치',
      'アバターサムネイル': '아바타 썸네일',
      '長辺 800px': '긴 변 800px',
      'VRCギャラリー': 'VRC 갤러리',
      '絵文字・ステッカー': '이모지/스티커',
      'プリセット': '프리셋',
      '補正の強さ': '보정 강도',
      'プリセット名': '프리셋 이름',
      'プリセット名を入力': '프리셋 이름 입력',
      'この設定を保存': '이 설정 저장',
      '補正': '보정',
      '明るさ': '밝기',
      '露出': '노출',
      'コントラスト': '대비',
      'ハイライト': '하이라이트',
      'シャドウ': '그림자',
      'ホワイト': '화이트',
      'ブラック': '블랙',
      'ガンマ': '감마',
      '色温度': '색온도',
      '色合い': '색조',
      '彩度': '채도',
      '自然な彩度': '자연스러운 채도',
      '明瞭度': '명료도',
      'テクスチャ': '텍스처',
      'シャープ': '선명도',
      'ノイズ低減': '노이즈 감소',
      'フェード': '페이드',
      '粒子': '그레인',
      'ビネット': '비네트',
      'トーンカーブ': '톤 커브',
      '全体': '전체',
      'ぼかし': '블러',
      'ぼかしをリセット': '블러 초기화',
      'どこをぼかしにするか': '블러 대상',
      '全体ぼかし': '전체 블러',
      '放射ぼかし': '방사형 블러',
      '放射': '방사형',
      'ぼかし量': '블러 양',
      '範囲': '범위',
      '四角で範囲選択': '사각형 범위 선택',
      '丸で範囲選択': '원형 범위 선택',
      'フリーハンドで範囲選択': '프리핸드 범위 선택',
      'ぼかしの濃さ': '블러 강도',
      'モザイクの濃さ': '모자이크 강도',
      '塗りつぶしの濃さ': '채우기 강도',
      '色': '색상',
      '確定': '확정',
      '範囲を編集': '범위 편집',
      '目隠し加工': '가림 처리',
      '目隠し加工をリセット': '가림 처리 초기화',
      'モザイク': '모자이크',
      '塗りつぶし': '채우기',
      'ひとつ戻す': '하나 되돌리기',
      '1つ戻す': '하나 되돌리기',
      '書き出し': '내보내기',
      '書き出し設定をリセット': '내보내기 설정 초기화',
      '形式': '형식',
      'サイズ': '크기',
      '長辺 3840px': '긴 변 3840px',
      '長辺 2560px': '긴 변 2560px',
      '長辺 2048px': '긴 변 2048px',
      '長辺 1600px': '긴 변 1600px',
      '長辺 1200px': '긴 변 1200px',
      '長辺 1024px': '긴 변 1024px',
      '品質': '품질',
      'テキスト追加': '텍스트 추가',
      'テキスト': '텍스트',
      'テキストをリセット': '텍스트 초기화',
      'テキストを追加': '텍스트 추가',
      'テキストを削除': '텍스트 삭제',
      '追加': '추가',
      '文字': '글자',
      '内容': '내용',
      'テキストを入力': '텍스트 입력',
      'フォント': '폰트',
      '太さ': '굵기',
      '縁': '테두리',
      '縁の色': '테두리 색상',
      '縁の種類': '테두리 종류',
      'なし': '없음',
      '縁取り': '외곽선',
      '影': '그림자',
      '発光': '발광',
      '縁内を透過': '글자 내부 투명',
      '適用範囲': '적용 범위',
      '被写体': '피사체',
      '被写体のみ': '피사체만',
      '背景のみ': '배경만',
      '文字サイズ': '글자 크기',
      '縁の太さ': '테두리 두께',
      '文字間隔': '자간',
      'AI被写体': 'AI 피사체',
      'AI被写体選択': 'AI 피사체 선택',
      '背景透過で保存': '투명 배경으로 저장',
      '背景透過PNGを保存中...': '투명 배경 PNG 저장 중...',
      '背景透過で保存するにはAI被写体選択のマスクが必要です':
        '투명 배경으로 저장하려면 AI 피사체 선택 마스크가 필요합니다.',
      '被写体マスクを使って背景を透過したPNGとして保存します':
        '피사체 마스크를 사용해 배경을 투명하게 만든 PNG로 저장합니다.',
      'AI被写体選択でマスクを作成すると使用できます':
        'AI 피사체 선택으로 마스크를 만들면 사용할 수 있습니다.',
      '背景透過用の描画を準備できませんでした':
        '투명 배경 렌더링을 준비할 수 없었습니다.',
      '背景透過に使う被写体マスクを読み込めませんでした':
        '투명 배경 저장에 사용할 피사체 마스크를 불러올 수 없었습니다.',
      '補正対象': '보정 대상',
      '被写体マスクをリセット': '피사체 마스크 초기화',
      'マスク未作成': '마스크 없음',
      'ダミーマスク': '더미 마스크',
      '軽量AIマスク': '경량 AI 마스크',
      '標準AIマスク': '표준 AI 마스크',
      '高精度AIマスク': '고정밀 AI 마스크',
      '読み込みマスク': '불러온 마스크',
      'マスク': '마스크',
      '被写体を自動選択': '피사체 자동 선택',
      '標準AIをダウンロード': '표준 AI 다운로드',
      '高精度AIをダウンロード': '고정밀 AI 다운로드',
      '高精度で再検出': '고정밀로 다시 감지',
      'マスク画像を読み込み': '마스크 이미지 불러오기',
      'マスクを表示': '마스크 표시',
      'マスク反転': '마스크 반전',
      'マスク修正': '마스크 수정',
      '修正する': '수정',
      '消す': '지우기',
      'ブラシサイズ': '브러시 크기',
      '濃度': '농도',
      'スポイト': '스포이드',
      'マスク画像を読み込み中です': '마스크 이미지를 불러오는 중입니다...',
      '現在: 標準AIで実行されます': '현재: 표준 AI로 실행됩니다',
      '現在: 軽量AIで実行されます': '현재: 경량 AI로 실행됩니다',
      '現在: 軽量AIで実行されます / 標準AIをダウンロードすると検出精度が向上します':
        '현재: 경량 AI로 실행됩니다 / 표준 AI를 다운로드하면 감지 정확도가 향상됩니다.',
      'AIモデルを利用できません': 'AI 모델을 사용할 수 없습니다',
      '設定から高精度AIモデルをダウンロードしてください':
        '설정에서 고정밀 AI 모델을 다운로드하세요.',
      'AIモデル情報を取得できません': 'AI 모델 정보를 가져올 수 없습니다',
      'ONNX Runtimeを読み込めませんでした': 'ONNX Runtime을 불러오지 못했습니다',
      'AIマスクの出力を読み取れませんでした':
        'AI 마스크 출력을 읽을 수 없었습니다',
      '標準AIの実行機能を利用できません':
        '표준 AI 실행 기능을 사용할 수 없습니다',
      'AIモデルの実行に失敗しました': 'AI 모델 실행에 실패했습니다',
      '画像サイズを取得できませんでした': '이미지 크기를 가져올 수 없었습니다',
      'AI入力画像を作成できませんでした': 'AI 입력 이미지를 만들 수 없었습니다',
      'AIマスクの出力が空でした': 'AI 마스크 출력이 비어 있었습니다',
      'AIマスクを作成できませんでした': 'AI 마스크를 만들 수 없었습니다',
      'AIマスクを切り出せませんでした': 'AI 마스크를 잘라낼 수 없었습니다',
      'AIマスクを元画像サイズに戻せませんでした':
        'AI 마스크를 원본 이미지 크기로 되돌릴 수 없었습니다',
      '高解像度マスクを整えています...': '고해상도 마스크를 다듬는 중...',
      'PNG / JPEG / WebP のマスク画像を選択してください':
        'PNG, JPEG 또는 WebP 마스크 이미지를 선택하세요.',
      'マスク画像を読み込めませんでした': '마스크 이미지를 불러올 수 없었습니다',
      '画像オーバーレイ': '이미지 오버레이',
      '画像オーバーレイをリセット': '이미지 오버레이 초기화',
      '画像オーバーレイ素材を読み込めませんでした':
        '이미지 오버레이 소재를 불러오지 못했습니다',
      '画像オーバーレイ素材の追加機能を利用できません':
        '이미지 오버레이 소재 추가 기능을 사용할 수 없습니다',
      '画像を追加': '이미지 추가',
      '管理素材': '관리 소재',
      '管理素材から削除': '관리 소재에서 삭제',
      '管理素材を削除': '관리 소재 삭제',
      '管理素材を削除できませんでした': '관리 소재를 삭제하지 못했습니다',
      '管理素材を削除しました': '관리 소재를 삭제했습니다',
      'レイヤー': '레이어',
      '手前へ': '앞으로',
      '奥へ': '뒤로',
      '透明度': '투명도',
      '合成方法': '합성 방법',
      '標準合成': '일반',
      '乗算': '곱하기',
      'スクリーン': '스크린',
      'オーバーレイ': '오버레이',
      'ソフトライト': '소프트 라이트',
      'ハードライト': '하드 라이트',
      '比較(暗)': '어둡게',
      '比較(明)': '밝게',
      '最近使用': '최근 사용',
      'システム': '시스템',
      '標準': '보통',
      '中太': '중간',
      'セミボールド': '세미볼드',
      '太字': '굵게',
      '特太': '아주 굵게',
      '極太': '블랙',
      'ルーラー': '눈금자',
      '三分割グリッド': '삼분할 그리드',

      '✨ 自動補正': '✨ 스마트 자동',
      '学習補正': '학습 보정',
      '投稿クリア': '게시용 선명',
      '自然クリア': '자연 선명',
      '夜景強調': '야경 강조',
      'ネオン強調': '네온 강조',
      'ふんわり1': '부드럽게 1',
      'ふんわり2': '부드럽게 2',
      'ふんわり3': '부드럽게 3',
      'フィルム風': '필름풍',
      '高コントラスト1': '고대비 1',
      '高コントラスト2': '고대비 2',
      'クールブルー1': '쿨 블루 1',
      'クールブルー2': '쿨 블루 2',
      'スイートピンク1': '스위트 핑크 1',
      'スイートピンク2': '스위트 핑크 2',
      'スイートピンク3': '스위트 핑크 3',
      'エアリーホワイト1': '에어리 화이트 1',
      'エアリーホワイト2': '에어리 화이트 2',
      'エアリーホワイト3': '에어리 화이트 3',
      'エアリーホワイト4': '에어리 화이트 4',
      '暗部クリア': '어두운 부분 선명',
      'サムネ強調': '썸네일 강조',
      'モノクロ': '흑백',

      '設定': '설정',
      '設定モーダルを閉じる': '설정 닫기',
      '言語設定': '언어 설정',
      '日本語': '日本語',
      '英語': 'English',
      '韓国語': '한국어',
      'フォント': '폰트',
      '更新対象フォルダの設定': '갱신 대상 폴더 설정',
      'フォルダ追加': '폴더 추가',
      'まだ登録されていません': '아직 등록되지 않았습니다',
      'AIモデル': 'AI 모델',
      '画像編集の被写体選択で使うモデルを管理します':
        '이미지 편집의 피사체 선택에 사용하는 모델을 관리합니다.',
      'AIモデル情報を取得できませんでした':
        'AI 모델 정보를 가져올 수 없었습니다',
      'AIモデルのダウンロード機能を利用できません':
        'AI 모델 다운로드 기능을 사용할 수 없습니다',
      'このAIモデルはダウンロードできません':
        '이 AI 모델은 다운로드할 수 없습니다',
      '画像は外部サーバーに送信されません。AI処理はPC内で実行されます。':
        '이미지는 외부 서버로 전송되지 않습니다. AI 처리는 PC 안에서 실행됩니다.',
      'このモデルは標準AIより処理時間やメモリ使用量が大きくなる場合があります。':
        '이 모델은 표준 AI보다 처리 시간이 길거나 메모리 사용량이 클 수 있습니다.',
      'AIモデルをダウンロードできませんでした':
        'AI 모델을 다운로드할 수 없었습니다',
      '保存場所を開けません': '저장 위치를 열 수 없습니다',
      '保存場所を開けませんでした': '저장 위치를 열 수 없었습니다',
      'AIモデルを削除': 'AI 모델 삭제',
      'AIモデルの削除機能を利用できません':
        'AI 모델 삭제 기능을 사용할 수 없습니다',
      'AIモデルを削除しています...': 'AI 모델을 삭제하는 중...',
      'AIモデルを削除できませんでした': 'AI 모델을 삭제할 수 없었습니다',
      'AIモデルを削除しました': 'AI 모델을 삭제했습니다',
      'ライセンス確認により利用停止中です':
        '라이선스 확인 결과 사용이 중지되었습니다',
      '軽量AI': '경량 AI',
      '標準AI': '표준 AI',
      '高精度AI': '고정밀 AI',
      '高精度AI候補': '고정밀 AI 후보',
      '旧モデル': '이전 모델',
      '利用可能': '사용 가능',
      '検証失敗': '검증 실패',
      '準備中': '준비 중',
      '利用停止中': '사용 중지',
      '一部不足': '일부 누락',
      '未ダウンロード': '미다운로드',
      '同梱': '동봉',
      '管理フォルダ': '관리 폴더',
      '同梱モデル': '동봉 모델',
      'ダウンロード': '다운로드',
      '配布元': '배포처',
      '保存場所': '저장 위치',
      '未確定': '미확정',
      'すぐ使える軽量な被写体マスク生成モデルです。複雑な髪型、衣装、羽、尻尾などでは精度が落ちる場合があります。':
        '바로 사용할 수 있는 경량 피사체 마스크 생성 모델입니다. 복잡한 헤어스타일, 의상, 날개, 꼬리 등에서는 정확도가 떨어질 수 있습니다.',
      'すぐ使える高速モデルです。複雑な髪型・衣装・羽・尻尾などでは精度が落ちる場合があります。':
        '바로 사용할 수 있는 빠른 모델입니다. 복잡한 헤어스타일, 의상, 날개, 꼬리 등에서는 정확도가 떨어질 수 있습니다.',
      'VRChatアバター向けの通常モデルです。細かい髪型・衣装・装飾・羽・尻尾などの検出精度が向上します。':
        'VRChat 아바타용 일반 모델입니다. 세밀한 헤어스타일, 의상, 장식, 날개, 꼬리 등의 감지 정확도가 향상됩니다.',
      'VRChatアバター向けの標準モデルです。軽量AIよりも、髪型・衣装・装飾・羽・尻尾などの検出精度向上を狙います。':
        'VRChat 아바타용 표준 모델입니다. 경량 AI보다 헤어스타일, 의상, 장식, 날개, 꼬리 등의 감지 정확도 향상을 목표로 합니다.',
      '境界をよりきれいに検出する上位モデルです。背景ぼかし、被写体だけ補正、文字や画像を被写体に合わせる加工に向いています。':
        '경계를 더 깔끔하게 감지하는 상위 모델입니다. 배경 흐림, 피사체만 보정, 문자나 이미지를 피사체에 맞추는 편집에 적합합니다.',
      '境界をよりきれいに検出する上位モデルです。背景ぼかし、被写体だけ補正、文字や画像を被写体に合わせる加工に向いています。処理時間やメモリ使用量が大きくなる場合があります。':
        '경계를 더 깔끔하게 감지하는 상위 모델입니다. 배경 흐림, 피사체만 보정, 문자나 이미지를 피사체에 맞추는 편집에 적합합니다. 처리 시간과 메모리 사용량이 커질 수 있습니다.',
      '学習データセットの商用利用条件を確認した結果、現在の推奨構成では使用しません。必要に応じて削除できます。':
        '학습 데이터셋의 상업적 이용 조건을 확인한 결과, 현재 권장 구성에서는 사용하지 않습니다. 필요하면 삭제할 수 있습니다.',
      '学習データセットの商用利用条件を確認した結果、現在の推奨構成では使用しません。':
        '학습 데이터셋의 상업적 이용 조건을 확인한 결과, 현재 권장 구성에서는 사용하지 않습니다.',
      '旧構成で高精度AIとして扱っていたモデルです。現在の推奨構成では使用しません。':
        '이전 구성에서 고정밀 AI로 취급하던 모델입니다. 현재 권장 구성에서는 사용하지 않습니다.',
      'データ管理': '데이터 관리',
      'ラベル、メモ、お気に入り、World情報を保存・書き出しできます':
        '라벨, 메모, 즐겨찾기, World 정보를 저장하고 내보낼 수 있습니다.',
      'データ管理を開く': '데이터 관리 열기',
      '操作メニュー': '작업 메뉴',
      'バックアップを作成': '백업 만들기',
      '状態チェック': '상태 확인',
      '元画像なしを表示': '원본 없음 표시',
      'サムネイルなしを表示': '썸네일 없음 표시',
      'World情報未取得を表示': 'World 정보 없음 표시',
      'World要確認を表示': 'World 확인 필요 표시',
      '欠損サムネイルを再生成': '누락 썸네일 재생성',
      'World要確認を再取得': 'World 확인 필요 재조회',
      'バックアップから復元': '백업에서 복원',
      'CSVエクスポート': 'CSV 내보내기',
      'JSONエクスポート': 'JSON 내보내기',
      'メンテナンス': '유지 관리',
      '表示中の月を削除': '현재 월 삭제',
      'サムネイルキャッシュを削除': '썸네일 캐시 삭제',
      '既存画像の情報を再取り込み': '기존 사진 정보 다시 가져오기',
      '全登録を削除': '전체 등록 삭제',
      'DBを初期化': 'DB 초기화',
      'アンインストール': '제거',
      'アプリ本体を削除します。必要に応じて、保存済みデータも一緒に削除できます。':
        '앱을 제거합니다. 필요하면 저장된 데이터도 함께 삭제할 수 있습니다.',
      'データも削除してアンインストール': '데이터도 삭제하고 제거',
      '確認': '확인',
      'この操作を実行しますか？': '이 작업을 실행할까요?',
      'キャンセル': '취소',
      '実行する': '실행',
      '閉じる': '닫기',
      '概要': '개요',
      '写真': '사진',
      'フォルダ': '폴더',
      'ワールド数': '월드 수',
      '背景': '배경',
      '画像を選択': '이미지 선택',
      '一覧を表示': '목록 보기',
      '更新対象フォルダ一覧': '갱신 대상 폴더 목록',
      'サムネイル再生成の対象月': '썸네일 재생성 대상 월',
      '情報再取り込みの対象月': '정보 재가져오기 대상 월',
      '再生成する月を選択': '재생성할 월 선택',
      '再取り込みする月を選択': '다시 가져올 월 선택',
      '対象月がありません': '대상 월이 없습니다',
      'ラベルを設定': '라벨 설정',
      '既存ラベルを再利用したり、新しいラベルを色付きで追加して写真ごとに設定できます。':
        '기존 라벨을 재사용하거나 색상이 있는 새 라벨을 만들어 사진별로 설정할 수 있습니다.',
      '現在のラベル': '현재 라벨',
      'ラベルを設定する': '라벨 선택',
      '既存ラベルを選択': '기존 라벨 선택',
      'この月にはラベルがありません': '이 월에는 라벨이 없습니다',
      'ラベルを作成する': '라벨 만들기',
      'ラベルの名前を入力してください': '라벨 이름을 입력하세요',
      '色を選択': '색상 선택',
      'この内容で追加': '이 내용으로 추가',
      '追加できるラベルはありません': '추가할 수 있는 라벨이 없습니다',
      'ラベルはまだ設定されていません': '아직 라벨이 설정되지 않았습니다',
      '読み込み中...': '불러오는 중...',
      'ワールド名を編集': '월드 이름 편집',
      'ワールド名': '월드 이름',
      '手動設定を解除': '수동 설정 해제',
      '再度ワールド名を自動取得': '월드 이름 다시 자동 가져오기',
      'World情報を再読み込み': 'World 정보 다시 읽기',
      '特殊文字などで自動取得名が崩れる場合に、表示名を手動で上書きできます。':
        '특수 문자 등으로 자동 이름이 깨질 때 표시 이름을 직접 덮어쓸 수 있습니다.',
      'ここに画像またはフォルダをドロップ': '사진 또는 폴더를 여기에 놓기',
      'png / jpg / jpeg / webp に対応': 'png / jpg / jpeg / webp 지원',
      '/ 比較中': '/ 비교 중',
      '/ 未確定: 1件': '/ 미확정: 1건',
      'World: すべて': 'World: 전체',
      'Worldフィルタ: すべて': 'World 필터: 전체',
      'Worldメタデータ要確認の該当分だけを再取得':
        'World 메타데이터 확인 필요 항목만 다시 가져오기',
      'Worldメタデータ要確認の写真だけを表示':
        'World 메타데이터 확인 필요 사진만 표시',
      'World情報の自動同期を開始できませんでした':
        'World 정보 자동 동기화를 시작할 수 없습니다',
      'World情報未取得の写真だけを表示':
        'World 정보가 없는 사진만 표시',
      'World設定を保存しました': 'World 설정을 저장했습니다',
      'World名フィルタ': 'World 이름 필터',
      'アップデートがあります': '업데이트가 있습니다',
      'アップデートの準備ができました': '업데이트 준비가 완료되었습니다',
      'アップデートは準備済みです。あとで再起動して適用できます':
        '업데이트가 준비되었습니다. 나중에 다시 시작해 적용할 수 있습니다.',
      'アップデートは保留しました': '업데이트를 보류했습니다',
      'アップデートを開始できませんでした': '업데이트를 시작할 수 없었습니다',
      'アップデートを適用できませんでした': '업데이트를 적용할 수 없었습니다',
      'アプリデータをJSONでバックアップ': '앱 데이터를 JSON으로 백업',
      'アンインストールの確認を開く': '제거 확인 열기',
      'アンインストールモーダルを閉じる': '제거 대화상자 닫기',
      'アンインストールを開始します': '제거를 시작합니다',
      'アンインストールを開始できませんでした': '제거를 시작할 수 없었습니다',
      'お気に入りを切り替え': '즐겨찾기 전환',
      'このフォルダを削除': '이 폴더 삭제',
      'サムネイルなしの写真だけを表示': '썸네일이 없는 사진만 표시',
      'サムネイルを再生成しています...': '썸네일을 다시 생성하는 중...',
      'サムネイルを再生成する月を選択': '썸네일을 다시 생성할 월 선택',
      'サムネイル再生成が完了しました': '썸네일 재생성이 완료되었습니다',
      'すべての登録を削除': '모든 등록 삭제',
      'データ管理を閉じる': '데이터 관리 닫기',
      'データ削除とアンインストールを開始します':
        '데이터 삭제와 제거를 시작합니다',
      'ドロップされた項目に対応画像がありませんでした':
        '놓은 항목에서 지원되는 사진을 찾을 수 없었습니다',
      'バックアップJSONから復元': '백업 JSON에서 복원',
      'フォルダ内に対応画像がありませんでした':
        '폴더에서 지원되는 사진을 찾을 수 없었습니다',
      'プリセットを保存できませんでした': '프리셋을 저장할 수 없었습니다',
      'プリセット名を入力してください': '프리셋 이름을 입력하세요',
      'プレビュー補助': '미리보기 보조',
      'メモを保存しました': '메모를 저장했습니다',
      'ラベルの保存に失敗しました': '라벨 저장에 실패했습니다',
      'ラベルを読み込み中...': '라벨을 불러오는 중...',
      'ラベルを保存しました': '라벨을 저장했습니다',
      'ラベル名を確認してください': '라벨 이름을 확인하세요',
      'ラベル名を入力してください': '라벨 이름을 입력하세요',
      'ワールド名を手動修正': '월드 이름 수동 수정',
      '右に90度回転': '오른쪽으로 90도 회전',
      '画像パスを再検出して更新しました':
        '이미지 경로를 다시 찾아 업데이트했습니다',
      '画像を読み込めませんでした': '이미지를 불러올 수 없었습니다',
      '画像処理Workerがタイムアウトしました':
        '이미지 처리 Worker가 시간 초과되었습니다',
      '画像処理Workerで失敗しました': '이미지 처리 Worker에서 실패했습니다',
      '画像処理Workerを起動できませんでした':
        '이미지 처리 Worker를 시작할 수 없었습니다',
      '画像処理Workerを利用できません': '이미지 처리 Worker를 사용할 수 없습니다',
      '画像処理プレビューを更新しました': '이미지 처리 미리보기를 업데이트했습니다',
      '画像読み込みがキャンセルされました': '이미지 불러오기가 취소되었습니다',
      '画像編集モーダルを閉じました': '이미지 편집기를 닫았습니다',
      '画像編集モーダルを閉じる': '이미지 편집기 닫기',
      '拡大画像': '확대 이미지',
      '確認モーダルを閉じる': '확인 대화상자 닫기',
      '学習補正を適用しました': '학습 보정을 적용했습니다',
      '管理サムネイルとサムネイル参照を削除':
        '관리 썸네일과 썸네일 참조 삭제',
      '欠損しているサムネイルだけを再生成': '누락된 썸네일만 다시 생성',
      '月を選択すると利用できます': '월을 선택하면 사용할 수 있습니다',
      '検索対象を切り替え': '검색 대상 전환',
      '元画像なしの写真だけを表示': '원본 이미지가 없는 사진만 표시',
      '現在の絞り込みでは、この日の写真は表示されていません':
        '현재 필터에서는 이 날짜의 사진이 표시되지 않습니다',
      '更新に失敗しました': '업데이트에 실패했습니다',
      '更新はキャンセルされました': '업데이트가 취소되었습니다',
      '更新確認: 新規0件': '업데이트 확인: 신규 0건',
      '更新対象のフォルダがまだ登録されていません':
        '아직 업데이트 대상 폴더가 등록되지 않았습니다',
      '更新対象フォルダがまだありません': '아직 추적 폴더가 없습니다',
      '更新対象フォルダを追加': '추적 폴더 추가',
      '更新対象フォルダ一覧を表示': '추적 폴더 목록 표시',
      '更新対象フォルダ一覧を閉じる': '추적 폴더 목록 닫기',
      '今すぐ更新': '지금 업데이트',
      '左に90度回転': '왼쪽으로 90도 회전',
      '再起動して更新': '다시 시작해 업데이트',
      '再取り込みできる月がありません': '다시 가져올 수 있는 월이 없습니다',
      '再生成できる月がありません': '다시 생성할 수 있는 월이 없습니다',
      '最新バージョン': '최신 버전',
      '削除してアンインストール': '삭제하고 제거',
      '削除するサムネイルがありません': '삭제할 썸네일이 없습니다',
      '削除する登録がありません': '삭제할 등록이 없습니다',
      '自動同期中は再読み込みできません': '자동 동기화 중에는 다시 읽을 수 없습니다',
      '自動補正を解除しました': '자동 보정을 해제했습니다',
      '自動補正を適用しました': '자동 보정을 적용했습니다',
      '写真を選択すると利用できます': '사진을 선택하면 사용할 수 있습니다',
      '主な更新内容': '주요 업데이트 내용',
      '写真一覧をCSVで書き出し': '사진 목록을 CSV로 내보내기',
      '写真一覧をJSONで書き出し': '사진 목록을 JSON으로 내보내기',
      '取り込み中です。処理が終わってから次の取り込みを開始してください':
        '가져오는 중입니다. 처리가 끝난 뒤 다음 가져오기를 시작하세요.',
      '処理中はアンインストールを開始できません':
        '처리 중에는 제거를 시작할 수 없습니다',
      '処理中はエクスポートできません': '처리 중에는 내보낼 수 없습니다',
      '処理中はバックアップできません': '처리 중에는 백업할 수 없습니다',
      '処理中はフォルダを削除できません': '처리 중에는 폴더를 삭제할 수 없습니다',
      '処理中はフォルダを追加できません': '처리 중에는 폴더를 추가할 수 없습니다',
      '処理中は再取得できません': '처리 중에는 다시 가져올 수 없습니다',
      '処理中は再生成できません': '처리 중에는 다시 생성할 수 없습니다',
      '処理中は状態チェックできません': '처리 중에는 상태를 확인할 수 없습니다',
      '処理中は抽出できません': '처리 중에는 추출할 수 없습니다',
      '処理中は復元できません': '처리 중에는 복원할 수 없습니다',
      '初期化するデータがありません': '초기화할 데이터가 없습니다',
      '情報を再取り込みする月を選択': '정보를 다시 가져올 월 선택',
      '選択を切り替え': '선택 전환',
      '選択中の月のサムネイルを再生成': '선택한 월의 썸네일 다시 생성',
      '選択中の月の登録画像を現在の解析ロジックで再取り込み':
        '선택한 월의 등록 사진을 현재 분석 로직으로 다시 가져오기',
      '全期間': '전체 기간',
      '追跡フォルダを確認しました': '추적 폴더를 확인했습니다',
      '追跡フォルダを更新しました': '추적 폴더를 업데이트했습니다',
      '追跡フォルダを更新中...': '추적 폴더 업데이트 중...',
      '登録・キャッシュ・更新対象フォルダを初期化':
        '등록, 캐시, 추적 폴더 초기화',
      '登録されたフォルダがありません': '등록된 폴더가 없습니다',
      '登録データの状態をチェック': '등록 데이터 상태 확인',
      '濃さ': '강도',
      '背景画像をクリアしました': '배경 이미지를 지웠습니다',
      '背景画像を更新しました': '배경 이미지를 업데이트했습니다',
      '編集結果を描画できませんでした': '편집 결과를 그릴 수 없었습니다',
      '保存しました': '저장했습니다',
      '保存に失敗しました': '저장에 실패했습니다',
      '保存をキャンセルしました': '저장을 취소했습니다',
      '保存機能を利用できません': '저장 기능을 사용할 수 없습니다',
      '保存済みプリセットがないためスマート自動補正を適用しました':
        '저장된 프리셋이 없어 스마트 자동 보정을 적용했습니다',
      '本当に削除しますか？': '정말 삭제할까요?',
      '例: Waiting / 待ち': '예: Waiting / 대기',
      '時刻不明': '알 수 없는 시간',
    },
  });

  const ATTRIBUTES = ['aria-label', 'title', 'placeholder', 'label'];
  const textOriginals = new WeakMap();
  const textLastApplied = new WeakMap();
  const attrOriginals = new WeakMap();
  const attrLastApplied = new WeakMap();
  let currentLanguage = 'ja';
  let observer = null;
  let isApplying = false;

  function getStoredLanguage() {
    try {
      const savedLanguage = window.localStorage?.getItem(LANGUAGE_STORAGE_KEY);
      return SUPPORTED_LANGUAGES[savedLanguage] ? savedLanguage : 'ja';
    } catch {
      return 'ja';
    }
  }

  function storeLanguage(language) {
    try {
      window.localStorage?.setItem(LANGUAGE_STORAGE_KEY, language);
    } catch {
      // Local storage can be unavailable in restricted contexts.
    }
  }

  function getDictionary(language = currentLanguage) {
    return DICTIONARIES[language] || {};
  }

  function preserveOuterWhitespace(original, translated) {
    const match = String(original).match(/^(\s*)([\s\S]*?)(\s*)$/);
    if (!match) {
      return translated;
    }

    return `${match[1]}${translated}${match[3]}`;
  }

  function formatNumber(value, language = currentLanguage) {
    const locale =
      language === 'ko' ? 'ko-KR' : language === 'en' ? 'en-US' : 'ja-JP';
    return new Intl.NumberFormat(locale).format(Number(value) || 0);
  }

  function formatMonth(year, month, language = currentLanguage) {
    if (language === 'en') {
      const date = new Date(Number(year), Number(month) - 1, 1);
      return new Intl.DateTimeFormat('en-US', {
        month: 'short',
        year: 'numeric',
      }).format(date);
    }

    if (language === 'ko') {
      return `${year}년 ${Number(month)}월`;
    }

    return `${year}年${month}月`;
  }

  function formatEnglishCount(value, singular, plural = `${singular}s`) {
    const count = Number(value) || 0;
    const label = count === 1 ? singular : plural;
    return `${formatNumber(count, 'en')} ${label}`;
  }

  function translatePattern(core, language) {
    if (language === 'ja') {
      return core;
    }

    let match = core.match(/^(\d+)枚$/);
    if (match) {
      return language === 'ko'
        ? `${formatNumber(match[1], language)}장`
        : formatEnglishCount(match[1], 'photo');
    }

    match = core.match(/^全(\d+)枚$/);
    if (match) {
      return language === 'ko'
        ? `전체 ${formatNumber(match[1], language)}장`
        : `All ${formatEnglishCount(match[1], 'photo')}`;
    }

    match = core.match(/^(\d+)件$/);
    if (match) {
      return language === 'ko'
        ? `${formatNumber(match[1], language)}건`
        : formatEnglishCount(match[1], 'item');
    }

    match = core.match(/^(\d+)件選択中$/);
    if (match) {
      return language === 'ko'
        ? `${formatNumber(match[1], language)}건 선택됨`
        : `${formatNumber(match[1], language)} selected`;
    }

    match = core.match(/^登録済み (\d+)件$/);
    if (match) {
      return language === 'ko'
        ? `등록됨 ${formatNumber(match[1], language)}건`
        : `${formatNumber(match[1], language)} registered`;
    }

    match = core.match(/^(\d{4})年(\d{1,2})月$/);
    if (match) {
      return formatMonth(match[1], match[2], language);
    }

    match = core.match(/^(\d{4})年(\d{1,2})月 \((\d+)枚\)$/);
    if (match) {
      const monthText = formatMonth(match[1], match[2], language);
      const countText =
        language === 'ko'
          ? `${formatNumber(match[3], language)}장`
          : formatEnglishCount(match[3], 'photo');
      return `${monthText} (${countText})`;
    }

    match = core.match(/^(\d{4})年(\d{1,2})月 (.+)$/);
    if (match) {
      const monthText = formatMonth(match[1], match[2], language);
      const actionText = translateText(match[3], language);
      return `${monthText} ${actionText}`;
    }

    match = core.match(/^(\d{4})年$/);
    if (match) {
      return language === 'ko' ? `${match[1]}년` : match[1];
    }

    match = core.match(/^(\d+)日$/);
    if (match) {
      return language === 'ko'
        ? `${Number(match[1])}일`
        : `Day ${Number(match[1])}`;
    }

    match = core.match(/^(.+)をリセット$/);
    if (match) {
      const target = translateText(match[1], language);
      return language === 'ko' ? `${target} 초기화` : `Reset ${target}`;
    }

    match = core.match(/^(.+)を削除$/);
    if (match) {
      const target = translateText(match[1], language);
      return language === 'ko' ? `${target} 삭제` : `Delete ${target}`;
    }

    match = core.match(/^プリセットを保存しました: (.+)$/);
    if (match) {
      return language === 'ko'
        ? `프리셋 저장 완료: ${match[1]}`
        : `Preset saved: ${match[1]}`;
    }

    match = core.match(/^プリセットを削除しました: (.+)$/);
    if (match) {
      return language === 'ko'
        ? `프리셋 삭제 완료: ${match[1]}`
        : `Preset deleted: ${match[1]}`;
    }

    match = core.match(/^プリセットを適用しました: (.+)$/);
    if (match) {
      return language === 'ko'
        ? `프리셋 적용 완료: ${translateText(match[1], language)}`
        : `Preset applied: ${translateText(match[1], language)}`;
    }

    match = core.match(/^背景透過PNGを保存しました: (.+)$/);
    if (match) {
      return language === 'ko'
        ? `투명 배경 PNG 저장 완료: ${match[1]}`
        : `Transparent PNG saved: ${match[1]}`;
    }

    match = core.match(/^(.+)を管理素材から削除$/);
    if (match) {
      return language === 'ko'
        ? `${match[1]} 관리 소재에서 삭제`
        : `Delete ${match[1]} from managed assets`;
    }

    match = core.match(/^(.+) を管理素材から削除します。現在の編集で使っている同じ画像レイヤーも外します。$/);
    if (match) {
      return language === 'ko'
        ? `${match[1]}을(를) 관리 소재에서 삭제합니다. 현재 편집에서 같은 이미지를 사용하는 레이어도 제거됩니다.`
        : `${match[1]} will be deleted from managed assets. Layers using the same image in this edit will also be removed.`;
    }

    match = core.match(/^画像オーバーレイは(\d+)件までです$/);
    if (match) {
      return language === 'ko'
        ? `이미지 오버레이는 최대 ${formatNumber(match[1], language)}개까지 사용할 수 있습니다`
        : `Up to ${formatNumber(match[1], language)} image overlays can be used`;
    }

    match = core.match(/^画像オーバーレイを追加しました: (\d+)件$/);
    if (match) {
      return language === 'ko'
        ? `이미지 오버레이 추가 완료: ${formatNumber(match[1], language)}개`
        : `Image overlays added: ${formatNumber(match[1], language)}`;
    }

    match = core.match(/^画像オーバーレイの追加に失敗しました: (.*)$/);
    if (match) {
      return language === 'ko'
        ? `이미지 오버레이 추가 실패: ${match[1]}`
        : `Image overlay import failed: ${match[1]}`;
    }

    match = core.match(/^管理素材の削除に失敗しました: (.*)$/);
    if (match) {
      return language === 'ko'
        ? `관리 소재 삭제 실패: ${match[1]}`
        : `Managed asset delete failed: ${match[1]}`;
    }

    match = core.match(/^出力: (.+)$/);
    if (match) {
      let body = match[1]
        .replace(/目隠し/g, language === 'ko' ? '가림' : 'masks')
        .replace(/未確定/g, language === 'ko' ? '미확정' : 'draft')
        .replace(/黒つぶれ/g, language === 'ko' ? '블랙 클리핑' : 'black clipping')
        .replace(/白飛び/g, language === 'ko' ? '화이트 클리핑' : 'white clipping')
        .replace(/画像/g, language === 'ko' ? '이미지' : 'images')
        .replace(/比較中/g, language === 'ko' ? '비교 중' : 'comparing')
        .replace(/件/g, language === 'ko' ? '건' : '');
      return language === 'ko' ? `출력: ${body}` : `Output: ${body}`;
    }

    match = core.match(/^保存しました（登録は未反映）: (.*)$/);
    if (match) {
      return language === 'ko'
        ? `저장 완료(등록에는 반영되지 않음): ${match[1]}`
        : `Saved (not added to catalog): ${match[1]}`;
    }

    match = core.match(/^編集済み画像を保存しました: (.*)$/);
    if (match) {
      return language === 'ko'
        ? `편집한 이미지를 저장했습니다: ${match[1]}`
        : `Edited image saved: ${match[1]}`;
    }

    match = core.match(/^保存に失敗しました: (.*)$/);
    if (match) {
      return language === 'ko'
        ? `저장 실패: ${match[1]}`
        : `Save failed: ${match[1]}`;
    }

    match = core.match(/^(.+): (\d+)件失敗しました$/);
    if (match) {
      const label = translateText(match[1], language);
      return language === 'ko'
        ? `${label}: ${formatNumber(match[2], language)}건 실패`
        : `${label}: ${formatNumber(match[2], language)} failed`;
    }

    return core;
  }

  function translateText(value, language = currentLanguage) {
    const text = String(value ?? '');
    if (language === 'ja') {
      return text;
    }

    const core = text.trim().replace(/\s+/g, ' ');
    if (!core) {
      return text;
    }

    const dictionary = getDictionary(language);
    const translated = dictionary[core] || translatePattern(core, language);
    return preserveOuterWhitespace(text, translated);
  }

  function shouldSkipTextNode(node) {
    const parent = node?.parentElement;
    if (!parent) {
      return true;
    }

    if (['SCRIPT', 'STYLE', 'NOSCRIPT', 'TEXTAREA'].includes(parent.tagName)) {
      return true;
    }

    if (
      parent.tagName === 'OPTION' &&
      parent.closest('.settings-month-select')
    ) {
      return true;
    }

    return Boolean(
      parent.closest(
        '[data-i18n-ignore], .material-symbols-outlined, #photo-editor-text-list'
      )
    );
  }

  function localizeTextNode(node) {
    if (shouldSkipTextNode(node)) {
      return;
    }

    const value = node.nodeValue || '';
    if (!value.trim()) {
      return;
    }

    const lastApplied = textLastApplied.get(node);
    let original = textOriginals.get(node);
    if (original === undefined || (!isApplying && value !== lastApplied)) {
      original = value;
      textOriginals.set(node, original);
    }

    const translated = translateText(original);
    if (value !== translated) {
      isApplying = true;
      node.nodeValue = translated;
      isApplying = false;
    }
    textLastApplied.set(node, translated);
  }

  function getElementAttrMap(store, element) {
    let map = store.get(element);
    if (!map) {
      map = new Map();
      store.set(element, map);
    }
    return map;
  }

  function localizeElementAttributes(element) {
    if (!(element instanceof Element) || element.matches('[data-i18n-ignore]')) {
      return;
    }

    const originals = getElementAttrMap(attrOriginals, element);
    const lastAppliedValues = getElementAttrMap(attrLastApplied, element);

    ATTRIBUTES.forEach((attributeName) => {
      if (!element.hasAttribute(attributeName)) {
        return;
      }

      const value = element.getAttribute(attributeName) || '';
      if (!value.trim()) {
        return;
      }

      const lastApplied = lastAppliedValues.get(attributeName);
      let original = originals.get(attributeName);
      if (original === undefined || (!isApplying && value !== lastApplied)) {
        original = value;
        originals.set(attributeName, original);
      }

      const translated = translateText(original);
      if (value !== translated) {
        isApplying = true;
        element.setAttribute(attributeName, translated);
        isApplying = false;
      }
      lastAppliedValues.set(attributeName, translated);
    });
  }

  function localizeNode(node) {
    if (!node) {
      return;
    }

    if (node.nodeType === Node.TEXT_NODE) {
      localizeTextNode(node);
      return;
    }

    if (node.nodeType !== Node.ELEMENT_NODE) {
      return;
    }

    localizeElementAttributes(node);

    const walker = document.createTreeWalker(
      node,
      NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT
    );

    while (walker.nextNode()) {
      const currentNode = walker.currentNode;
      if (currentNode.nodeType === Node.TEXT_NODE) {
        localizeTextNode(currentNode);
      } else {
        localizeElementAttributes(currentNode);
      }
    }
  }

  function syncLanguageOptionButtons() {
    document.querySelectorAll('[data-language-option]').forEach((button) => {
      const isActive = button.dataset.languageOption === currentLanguage;
      button.classList.toggle('is-active', isActive);
      button.setAttribute('aria-pressed', isActive ? 'true' : 'false');
      button.setAttribute('role', 'radio');
      button.setAttribute('aria-checked', isActive ? 'true' : 'false');
    });
  }

  function applyLanguage(language, { notify = true } = {}) {
    const nextLanguage = SUPPORTED_LANGUAGES[language] ? language : 'ja';
    currentLanguage = nextLanguage;
    storeLanguage(nextLanguage);

    document.documentElement.lang =
      SUPPORTED_LANGUAGES[nextLanguage].htmlLang || nextLanguage;
    document.body?.setAttribute('data-language', nextLanguage);

    localizeNode(document.body);
    syncLanguageOptionButtons();

    if (notify) {
      window.dispatchEvent(
        new CustomEvent('worldshot:languagechange', {
          detail: { language: nextLanguage },
        })
      );
      window.dispatchEvent(new Event('resize'));
    }
  }

  function handleMutations(mutations) {
    mutations.forEach((mutation) => {
      if (mutation.type === 'characterData') {
        localizeTextNode(mutation.target);
        return;
      }

      if (mutation.type === 'attributes') {
        localizeElementAttributes(mutation.target);
        return;
      }

      mutation.addedNodes.forEach((node) => {
        localizeNode(node);
      });
    });
  }

  function initializeObserver() {
    if (observer || !document.body) {
      return;
    }

    observer = new MutationObserver(handleMutations);
    observer.observe(document.body, {
      childList: true,
      subtree: true,
      characterData: true,
      attributes: true,
      attributeFilter: ATTRIBUTES,
    });
  }

  document.addEventListener('click', (event) => {
    const button = event.target.closest('[data-language-option]');
    if (!button) {
      return;
    }

    applyLanguage(button.dataset.languageOption || 'ja');
  });

  window.WorldShotI18n = {
    getLanguage: () => currentLanguage,
    setLanguage: (language) => applyLanguage(language),
    t: (text) => translateText(text),
    localize: () => localizeNode(document.body),
    languages: SUPPORTED_LANGUAGES,
  };

  currentLanguage = getStoredLanguage();
  if (document.body) {
    applyLanguage(currentLanguage, { notify: false });
    initializeObserver();
  } else {
    document.addEventListener(
      'DOMContentLoaded',
      () => {
        applyLanguage(currentLanguage, { notify: false });
        initializeObserver();
      },
      { once: true }
    );
  }
})();
