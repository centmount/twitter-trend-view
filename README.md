# twitter-trend-view
Get twitter trend ranking data from twitter API every 10 minutes, Save to excel file, Email to your address. 
ツイッタートレンドランキングを定期的にExcelファイルに保存して、添付ファイルをメール送信します。

## /create_empty_sheet/main.py
日時を入れたファイル名を指定した「空のExcelシート」を作成してストレージにアップロードします。
[Google Cloud Platform]の[Cloud Functions]にデプロイして、[Cloud Scheduler]で定期実行します。
[Storge]にExcelファイルを保存します。
※ストレージのバケット名を指定して下さい。

## /to_excel_file/main.py
Twitter APIからトレンドデータを定期的にexcelファイルに保存します。
トップ50のランキングのデータを10分ごとに取得して、日時順に追加保存します。
[Google Cloud Platform]の[Cloud Functions]にデプロイして、[Cloud Scheduler]で定期実行します。
[Storge]にExcelファイルを更新して保存します。
※TwitterAPIの各種ツイッターのキーを入力（必要に応じてシークレットファイルなどを使用してください）

## /send_mail/main.py
Twitter APIから取得したトレンドデータを保存したexcelファイルをメールに添付して、送信します。
[Google Cloud Platform]の[Cloud Functions]にデプロイして、[Cloud Scheduler]で定期実行します。
[Storge]に保存したExcelファイルを取得します。
※送信元のgmailアカウント、パスワード、送信先のメールアドレスの入力（必要に応じてシークレットファイルなどを使用してください）
