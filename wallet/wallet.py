"""
このモジュールは、家計簿アプリケーションを提供します。
機能：
1. 家計簿データの登録
2. データの表示（テーブル形式、グラフ表示）
3. CSVおよびExcelへの保存
"""
# 標準ライブラリ
from datetime import datetime

# サードパーティライブラリ
import csv
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import openpyxl
import openpyxl.workbook
from tabulate import tabulate

# フォントを指定
FONT_PATH = "/System/Library/Fonts/ヒラギノ角ゴシック W6.ttc"
font_prop = fm.FontProperties(fname=FONT_PATH)

# システム内のフォント一覧を表示
plt.rcParams["font.family"] = font_prop.get_name()

# 入力機能


def input_date():
    """日付入力の関数"""
    while True:
        date = input("日付の入力(例:2025/01/01):\n")
        try:
            datetime.strptime(date, "%Y/%m/%d")
            return date
        except ValueError:
            print("日付はYYYY/MM/DDの形で入力してください")


def input_category(categories):
    """カテゴリの入力の関数"""
    print("カテゴリを選択してください:\n")
    for i, cat in enumerate(categories, start=1):
        print(f"{i}. {cat}")
    while True:
        try:
            category_index = int(input("番号を入力してください:\n")) - 1
            if 0 <= category_index < len(categories):
                return categories[category_index]
            else:
                print("番号が無効です。もう一度入力してください。")

        except ValueError:
            print("数字で入力してください。")


def input_amount():
    """金額の入力の関数"""
    while True:
        try:
            return int(float(input("金額を入力してください:\n")))
        except ValueError:
            print("金額は数字で入力してください。")


def add_transaction():
    """
    家計簿データを入力する関数。

    Returns:
        dict: 家計簿データ（日付、カテゴリ、金額）。
    """
    # カテゴリ選択肢
    categories = ["食費", "交通費", "日用品", "趣味/交際費", "その他"]

    date = input_date()
    category = input_category(categories)
    amount = input_amount()
    return {"日付": date, "カテゴリ": category, "金額": amount}


def save_to_csv(transaction_list, filename="wallet_data.csv"):
    """
    入力された家計簿データをCSVへ保存する関数。

    Args:
        transaction_list:家計簿の収支データが入ったリスト。
        filename (str): 読み込むCSVファイル名（デフォルト値: 'wallet_data.csv'）。
    """
    with open(filename, mode="w", newline="", encoding="utf-8")as file:
        writer = csv.DictWriter(file, fieldnames=["日付", "カテゴリ", "金額"])
        writer.writeheader()  # ヘッダーを書き込む
        writer.writerows(transaction_list)  # 収支データを書き込む

# CSV読み込み機能


def load_from_csv(filename="wallet_data.csv"):
    """
    CSVから家計簿データを読み込む関数。

    Args:
        filename (str): 読み込むCSVファイル名（デフォルト値: 'wallet_data.csv'）。

    Returns:
        list[dict]: 読み込んだ家計簿データのリスト。
                    各リスト要素は「日付」「カテゴリ」「金額」のキーを持つ辞書。
                    ファイルが存在しない場合は空のリストを返す。
    """
    try:
        with open(filename, mode="r", encoding="utf-8") as file:
            reader = csv.DictReader(file)
            return [{
                "日付": row["日付"],
                "カテゴリ": row["カテゴリ"],
                "金額": float(row["金額"])
            }
                for row in reader]
    except FileNotFoundError:
        # ファイルがない場合は空のリストを返す
        return []


def display_table(transaction_list):
    """
    家計簿データをテーブル形式で表示する関数。

    Args:
        transaction_list: 家計簿の収支データが入ったリスト。
    """
    if not transaction_list:
        print("家計簿の登録がありません。")
    else:
        print(tabulate(transaction_list, headers="keys", tablefmt="grid"))


def plot_graph(transaction_list):
    """
    家計簿データを元にカテゴリ別の支出を棒グラフで表示する関数。

    Args:
        transaction_list: 家計簿の収支データが入ったリスト。
    """

    if not transaction_list:
        print("家計簿の登録がないため、グラフを表示できません。")
        return

    # カテゴリごとの支出を集計する
    category_totals = {}
    for row in transaction_list:
        category = row["カテゴリ"]
        amount = float(row["金額"])
        if category in category_totals:
            # すでに登録されていたら加算する
            category_totals[category] += amount
        else:
            category_totals[category] = amount
    # グラフを作成
    categories = list(category_totals.keys())
    amounts = list(category_totals.values())

    plt.bar(categories, amounts, color="skyblue")
    plt.title("カテゴリ別支出", fontsize=16)
    plt.xlabel("カテゴリ", fontsize=12)
    plt.ylabel("金額(円)", fontsize=12)
    plt.grid(axis="y", linestyle="dotted", alpha=0.7)
    plt.show()


def save_to_excel(transaction_list, filename="wallet_data.xlsx"):
    """
    入力された家計簿データをエクセルへ保存する関数。

    Args:
        transaction_list:家計簿の収支データが入ったリスト。
        filename (str): 保存するエクセルファイル名（デフォルト値: 'wallet_data.xlsx'）。
    """
    try:
        wb = openpyxl.Workbook()  # 新しいエクセルファイル
        ws = wb.active  # アクティブなsheetを取得
        ws.title = "家計簿アプリ"

        # ヘッダー行を追加
        ws.append(["日付", "カテゴリ", "金額"])

        # データの書き込み
        dc = 0
        for row in transaction_list:
            ws.append([row["日付"], row["カテゴリ"], float(row["金額"])])
            dc += 1

        # ファイルを保存する。
        wb.save(filename)
        print(f"データをExcelファイル({filename})に保存しました。")
        print(f"データは{dc}件です。")

    except Exception as e:  # pylint: disable=broad-exception-caught
        print(f"エクセルへ保存中にエラーが発生しました。:{e}")

# 登録した収支を保存するリスト
# 保存されたCSVがある場合は、ファイルから読み込み
transactions = load_from_csv()

# メニュー
while True:
    print("\n1.収支を登録する")
    print("2.家計簿を表示する")
    print("3.グラフを表示する")
    print("4.エクセルに保存する")
    print("5.終了")

    choice = input("メニューを選択してください(例:1,2,3)\n")

    if choice == "1":
        # 登録
        transaction = add_transaction()
        transactions.append(transaction)
        save_to_csv(transactions)
        print("収支を登録しました")
    elif choice == "2":
        # テーブル形式で表示
        display_table(transactions)
    elif choice == "3":
        # 棒グラフで表示
        plot_graph(transactions)
    elif choice == "4":
        # エクセルで保存
        save_to_excel(transactions)
    elif choice == "5":
        # 終了
        print("アプリを終了します")
        save_to_csv(transactions)
        break
    else:
        print("入力する数値が間違っています。メニュー番号を入力してください。")
