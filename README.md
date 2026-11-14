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

## コート当番フォルダの設定

1. Apps Script のスクリプトプロパティ `COURT_DUTY_FOLDER_ID` に、保存先DriveフォルダのID（またはフォルダURL）を設定します。IDはフォルダURLの `/folders/` の後ろの値です。
2. ウェブアプリの実行ユーザー（上記設定ならデプロイした自分のGoogleアカウント）に、そのフォルダへのアクセス権があることを確認します。共有ドライブの場合も同じアカウントがメンバーである必要があります。
3. Apps Script エディターで `verifyCourtDutyFolderAccess` を選んで実行し、初回のGoogle Drive権限確認を完了します。成功すれば実行ログに確認メッセージが出ます。
4. ウェブアプリを新しいバージョンにデプロイします。

## GitHub Pages の設定

1. リポジトリの **Settings → Secrets and variables → Actions → Variables** で、`GAS_WEB_APP_URL` というRepository variableを作成し、手順2の `/exec` URLを設定します。これは秘密情報ではないため、SecretではなくVariableを使います。
2. **Settings → Pages → Build and deployment → Source** を **GitHub Actions** に設定します。
3. `main`へ`index.html`またはワークフローをpushすると、デプロイ時にURLがHTMLへ注入されます。Actionsタブから手動実行することもできます。値が未設定または不正な場合はデプロイを失敗させます。

ローカルで試す場合は、`index.html`の`GAS_WEB_APP_URL`プレースホルダーを一時的に `/exec` URLへ置き換え、HTTPサーバーからページを開いてください。そのオリジン（例: `http://localhost:8000`）は`GITHUB_PAGES_ORIGINS`に追加します。`file://`からは利用できません。

GASのHTML Serviceはブリッジを内側のiframeで実行するため、ブリッジから親ページへ接続通知を送り、そのウィンドウと `postMessage` で通信します。`Code.js` または `ApiBridge.html` を変更した場合は、Apps Script で新しいバージョンをデプロイしてください。
