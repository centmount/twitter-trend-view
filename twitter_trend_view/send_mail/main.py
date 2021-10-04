#!/usr/bin/env python3

"""
Twitter APIから取得したトレンドデータを保存したexcelファイルをメールに添付して、送信します。
[Google Cloud Platform]の[Cloud Functions]にデプロイして、[Cloud Scheduler]で定期実行します。
[Storge]に保存したExcelファイルを取得します。
"""

# 必要なモジュールをインポート
from smtplib import SMTP
from email.mime.text import MIMEText
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from google.cloud import storage
from datetime import datetime, timedelta
import pytz

# 日本時間を取得
japan_time = datetime.now() + timedelta(hours=9)

# 送信元のgmailアカウント、送信先のメールアドレスを入れてください
def input_mail_address():
    global sender
    global password
    global my_address
    while True:
        sender = input('送信元のgmailアドレスを入力して下さい: ') # 送信元メールアドレスを入力
        password = input('送信元のgoogleアカウントのログインパスワードを入力して下さい: ') # 送信元googleアカウントのログインパスワード入力 
        my_address = input('送信先のメールアドレスを入力して下さい: ')
        ok = input('入力完了しましたか？(完了の場合 OK と入力): ')
        if ok == 'OK':
            break
    return sender, password, my_address

# gmailを使ったメール送信のサンプルプログラム
def sendGmailAttach():
    to = my_address  # 送信先メールアドレス
    sub = 'Twitterトレンドデータ' #メール件名
    body = 'TwitterトレンドデータをExcelファイルにして添付しています。'  # メール本文
    host, port = 'smtp.gmail.com', 587

    # メールヘッダー
    msg = MIMEMultipart()
    msg['Subject'] = sub
    msg['From'] = sender
    msg['To'] = to

    # メール本文
    body = MIMEText(body)
    msg.attach(body)

# 添付ファイルの設定 # nameは添付ファイル名。pathは添付ファイルの場所を指定
    attach_file = {'name': f'trend_data{japan_time.strftime("%Y年%m月%d日")}.xlsx', 'path': f'/tmp/trend_data{japan_time.strftime("%Y年%m月%d日")}.xlsx'} 
    attachment = MIMEBase('application', 'xlsx')
    file = open(attach_file['path'], 'rb+')
    attachment.set_payload(file.read())
    file.close()
    encoders.encode_base64(attachment)
    attachment.add_header("Content-Disposition", "attachment", filename=attach_file['name'])
    msg.attach(attachment)

    # gmailへ接続(SMTPサーバーとして使用)
    gmail=SMTP(host, port)
    gmail.starttls() # SMTP通信のコマンドを暗号化し、サーバーアクセスの認証を通す
    gmail.login(sender, password)
    gmail.send_message(msg)

# Storageからのダウンロード処理
def download_blob(bucket_name, source_blob_name, destination_file_name):
    storage_client = storage.Client()
    bucket = storage_client.bucket(bucket_name)
    blob = bucket.blob(source_blob_name)
    blob.download_to_filename(destination_file_name)

# main関数を実行(GCPのCloud Functions用にevent,contextが引数に)
# トリガーをpub/subに設定、ファイルダウンロードしてメール送信
def main(event, context):
    download_blob('ストレージのバケット名を指定', f'trend_data{japan_time.strftime("%Y年%m月%d日")}.xlsx', f'/tmp/trend_data{japan_time.strftime("%Y年%m月%d日")}.xlsx')
    sendGmailAttach()
    