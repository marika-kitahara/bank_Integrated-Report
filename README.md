# 統合レポート作成アプリ

## ファイル
- `app.py`: Streamlit本体
- `output_template.xlsx`: 出力テンプレート
- `requirements.txt`: 依存ライブラリ

## 起動
```bash
pip install -r requirements.txt
streamlit run app.py
```

## 集計仕様
- 運用型実績の非表示シートも、非表示のまま読み込み可能です。
- 運用型の合算値は `【Search_合計】` / `【Display_合計】` の指定セルを利用します。
- 日別の日付列はA列・B列を自動判定します（提供サンプルはB列）。
- CPAは `コスト ÷ CV`、TVAは指定どおり `目標 ÷ 実績` です。
- AFF実績は、選択したAFFプランのA列に存在するサイト名と完全一致する行だけを集計します。
