# 家計簿アプリ "money"　マネー

「money」は、日々の収支を簡単に管理し、自動計算を行うためのデスクトップ家計簿アプリです。
タブ形式のUIで月ごとにシートを切り替えられ、固定費テンプレート機能も搭載しています。

## 主な機能

- **月別シート管理**: 年月（YYYYMM）を入力してシートを簡単に追加・削除可能。
- **3つのリスト表示**: 「収入」「支出」「固定費」を左右に配置し、見やすいレイアウト。
- **自動計算**: 金額を入力するたびに合計金額と収支がリアルタイムで更新されます。
- **データ永続化**: 各月のデータは JSON 形式で自動保存されます。
- **固定費テンプレート**: `categories.txt` を編集することで、毎月の固定費項目を自由にカスタマイズ可能。

## 動作環境

- macOS / Windows / Linux 対応
- Python 3.x (Tkinterを使用)

## インストール・実行方法

### 開発環境の構築
1. リポジトリをクローンまたはダウンロードします。
2. 仮想環境を作成し、依存関係をインストールします。
   ```bash
   python -m venv venv
   # Windowsの場合: .\venv\Scripts\activate
   # Mac/Linuxの場合: source venv/bin/activate
   pip install -r requirements.txt

3.okane.py を実行します。
   python money.py

4アプリケーションの作成 (コンパイル)
 PyInstallerを使用して、実行ファイルを作成できます。
 (※OSに合わせてアイコンファイル形式を変更してください)

Win)
pyinstaller --onefile --windowed --icon=app_icon.ico money.py
Mac)
pyinstaller --onefile --windowed --icon=app_icon.icns money.py
Linux)
pyinstaller --onefile --windowed --icon=app_icon.png money.py


★設定ファイル
    categories.txt: 固定費リストの初期値を管理します（1行1項目）。
    money_data/: 各月の家計簿データ（JSON）が格納されるディレクトリです。
★ライセンス
MIT License

★注意点
    家計データ（money_data/）は個人情報を含むため、
    GitHub等にアップロードする際は .gitignore で除外してください。