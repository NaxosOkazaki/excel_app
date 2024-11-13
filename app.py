from flask import Flask, render_template, request
import pandas as pd

app = Flask(__name__)

# ファイルアップロードの許可される拡張子
ALLOWED_EXTENSIONS = {'xlsx'}

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def upload_files():
    if request.method == 'POST':
        # ファイルが正しくアップロードされたか確認
        if 'x02_file' not in request.files or 'x04_file' not in request.files:
            return 'ファイルが選択されていません'

        x02_file = request.files['x02_file']
        x04_file = request.files['x04_file']

        # ファイルの拡張子を確認
        if not (allowed_file(x02_file.filename) and allowed_file(x04_file.filename)):
            return '許可されていないファイル形式です'

        # Excelファイルをデータフレームに読み込み
        x02_df = pd.read_excel(x02_file)
        x04_df = pd.read_excel(x04_file)


        x02_df['カテゴリー'] = pd.to_numeric(x02_df['カテゴリー'], errors='coerce')
        x04_df['正味'] = pd.to_numeric(x04_df['正味'], errors='coerce')
       x04_df['正味金額'] = pd.to_numeric(x04_df['正味金額'], errors='coerce')

        # 1) X04のA列とX02のB列でマッチング
        merged_df = pd.merge(
            x04_df,
            x02_df[['商品番号', 'カテゴリー']],
            left_on='商品番号',
            right_on='商品番号',
            how='inner'
        )

        # 2) カテゴリーのリスト
        categories = ["00", "13", "31", "38", "42", "44", "73", "91"]

        # 3) カテゴリーごとにE列とH列を合算
        result = merged_df[merged_df['カテゴリー'].isin(categories)].groupby('カテゴリー').agg({
            '正味': 'sum',
            '正味金額': 'sum'
        }).reset_index()

        # 結果をテンプレートに渡す
        return render_template('result.html', tables=[result.to_html(classes='data', header="true", index=False)])

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
