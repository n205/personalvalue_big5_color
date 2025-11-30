from flask import Flask
import pandas as pd
import gspread
from gspread_dataframe import get_as_dataframe
from google.oauth2 import service_account
import logging

from read_coデータ import read_coデータ
from update_co心理指標 import update_co個人価値観
from update_co心理指標 import update_cobig5
from update_co色 import update_co色番号
from update_co色 import update_co色


# Cloud Logging に出力するよう設定
logging.basicConfig(level=logging.INFO)

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def main():
    logging.info('📥 リクエスト受信')

    # スプレッドシート読込
    worksheet, existing_df, processed_urls = read_coデータ()

    update_co個人価値観(worksheet)
    update_cobig5(worksheet)
    update_co色番号(worksheet)
    update_co色(worksheet)
    
    return 'Cloud Run Function executed.', 200


if __name__ == '__main__':
    logging.info('🚀 アプリ起動')
    app.run(host='0.0.0.0', port=8080)
