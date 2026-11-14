# GitHub Pages / Google Apps Script の分離

## 配置

- `index.html` は GitHub Pages の公開ルートに配置します。
- `gas/Code.js` の内容をGASのスクリプトファイル（通常は `Code.gs`）、`gas/ApiBridge.html` をHTMLファイル `ApiBridge.html` として配置します。
- スプレッドシート、Drive、スクリプトプロパティなど、現在のGASプロジェクト設定は引き継いでください。

GitHub Pages とGASの間は、GASのHTML Serviceを隠しiframeとして読み込み、`postMessage`でRPCを行います。GAS側で `google.script.run` を呼び出すため、ブラウザーからGASへ直接 `fetch` する際のCORS制約を避けられます。ブリッジは許可したメソッドだけを実行し、親ページのオリジンも照合します。

## GASの設定とデプロイ

1. Apps Script の「プロジェクトの設定」→「スクリプト プロパティ」に `GITHUB_PAGES_ORIGINS` を追加します。値にはページのオリジンだけを指定します（例: `https://kuki-lawn-tennis.github.io`）。独自ドメインやローカル確認用URLがある場合はカンマ区切りで追加できます。パスや末尾の `/` は含めません。
2. `gas/Code.js` の内容をGASのスクリプトファイル（通常は `Code.gs`）、`gas/ApiBridge.html` をHTMLファイル `ApiBridge.html` に反映します。ウェブアプリとして新しいバージョンをデプロイし、実行ユーザーを自分、アクセスできるユーザーを公開サイトの利用者に設定します。管理操作は既存の管理者セッショントークンで保護されます。
3. デプロイされた `/exec` URLをコピーします。

## GitHub Pages の設定

`index.html` 冒頭のスクリプトにある `GAS_WEB_APP_URL` に、手順2で取得した `/exec` URLを設定してコミット・公開します。GitHub Pages の公開元は `index.html` がルートになるよう設定してください。

ローカルで試す場合はHTTPサーバーからページを開き、そのオリジン（例: `http://localhost:8000`）を `GITHUB_PAGES_ORIGINS` に追加します。`file://` からは利用できません。

GASのHTML Serviceはブリッジを内側のiframeで実行するため、ブリッジから親ページへ接続通知を送り、そのウィンドウと `postMessage` で通信します。`Code.js` または `ApiBridge.html` を変更した場合は、Apps Script で新しいバージョンをデプロイしてください。
