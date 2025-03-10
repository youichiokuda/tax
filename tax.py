import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import pyocr
import pyocr.builders
from PIL import Image
import os
import openai

# ==============================
# 1. Google Sheetsから取引データ取得
# ==============================

# Google Sheets APIの認証　以下２行変更する
SERVICE_ACCOUNT_FILE = "××.json"
SCOPES = ["https://www.××/spreadsheets"]

creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
client = gspread.authorize(creds)

# Google Sheetsの設定。以下２行変更する
SPREADSHEET_ID = "YOUR ID"
SHEET_NAME = "シートの名称を入力"

# データ取得
sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
data = sheet.get_all_records()
df = pd.DataFrame(data)

# ==============================
# 2. OCRによる領収書データの処理
# ==============================

# OCRエンジンのセットアップ
tool = pyocr.get_available_tools()[0]

def extract_text_from_receipt(image_path):
    """領収書画像からテキストを抽出"""
    image = Image.open(image_path)
    text = tool.image_to_string(image, builder=pyocr.builders.TextBuilder())
    return text

# OCRデータを取得（jpegファイル）以下１行変更する
receipt_folder = "領収書を格納するフォルダーの絶対パス入力"
receipt_files = [file for file in os.listdir(receipt_folder) if file.endswith(".jpg")]

receipt_data = []
for file in receipt_files:
    file_path = os.path.join(receipt_folder, file)
    extracted_text = extract_text_from_receipt(file_path)

    lines = extracted_text.split("\n")
    store = lines[0] if len(lines) > 0 else "不明"
    date = lines[1] if len(lines) > 1 else "不明"
    amount = 0
    for line in lines:
        if "¥" in line:
            try:
                amount = int(line.replace("¥", "").strip())
                break
            except ValueError:
                continue

    receipt_data.append({"日付": date, "取引内容": f"{store}での支払い", "金額": amount})

receipts_df = pd.DataFrame(receipt_data)

# ==============================
# 3. 銀行データの取得
# ==============================

def load_bank_data(csv_path):
    """銀行取引データをCSVから取得"""
    return pd.read_csv(csv_path)
#以下１行変更する
bank_df = load_bank_data("銀行データのCSVファイル格納絶対パスを入力/bank_transactions.csv")

# ==============================
# 4. データ統合 & AIによる仕訳
# ==============================

openai.api_key = "YOUR OPENAI API"

# 取引データを統合
merged_df = pd.concat([df, bank_df, receipts_df], ignore_index=True)

def categorize_transaction(description):
    """AIを使って勘定科目を自動分類"""
    response = openai.ChatCompletion.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": "適切な勘定科目を記載してください。また、取引が収益（PLの売上など）に当たるのか、支出（費用や資産購入）に当たるのかを判断し、金額のプラス・マイナスを適切に入力してください。勘定科目には文字の揺れがないよう統一し、BS（貸借対照表）またはPL（損益計算書）のどちらに属するかも明記してください。"},
            {"role": "user", "content": f"取引内容: {description}"}
        ]
    )
    return response["choices"][0]["message"]["content"].strip()

merged_df["勘定科目"] = merged_df["取引内容"].apply(categorize_transaction)

# ==============================
# 5. 勘定科目のフォーマット修正
# ==============================

def clean_account_name(account_name):
    """ AIの出力フォーマットを統一 """
    if pd.isna(account_name):
        return "未分類"

    account_name = account_name.replace("勘定科目: ", "").strip()
    if "\n" in account_name:
        account_name = account_name.split("\n")[0].strip()

    for delimiter in ["または", "、"]:
        if delimiter in account_name:
            account_name = account_name.split(delimiter)[0].strip()

    return account_name

merged_df["勘定科目"] = merged_df["勘定科目"].apply(clean_account_name)
merged_df["勘定科目"] = merged_df["勘定科目"].astype(str).str.strip()

# ==============================
# 6. BS・PL（貸借対照表・損益計算書）の作成
# ==============================

def generate_financial_statements(df):
    """貸借対照表（BS）と損益計算書（PL）を作成"""
    bs_accounts = ["現金", "預金", "売掛金", "買掛金", "固定資産"]
    pl_accounts = ["売上", "仕入", "給与手当", "消耗品費", "交際費", "旅費交通費", "水道光熱費", "電気料金"]

    bs = df[df["勘定科目"].isin(bs_accounts)].groupby("勘定科目", as_index=False)["金額"].sum()
    pl = df[df["勘定科目"].isin(pl_accounts)].groupby("勘定科目", as_index=False)["金額"].sum()

    return bs, pl

bs, pl = generate_financial_statements(merged_df)

# ==============================
# 7. Google Sheetsにアップロード
# ==============================

SHEET_OUTPUT_NAME = "Journal"
output_sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_OUTPUT_NAME)

data_to_upload = [merged_df.columns.tolist()] + merged_df.values.tolist()
output_sheet.update(range_name="A1", values=data_to_upload)

SHEET_BS = "BalanceSheet"
SHEET_PL = "ProfitLoss"

bs_sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_BS)
pl_sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_PL)

bs_data = [bs.columns.tolist()] + bs.values.tolist()
pl_data = [pl.columns.tolist()] + pl.values.tolist()

bs_sheet.update(range_name="A1", values=bs_data)
pl_sheet.update(range_name="A1", values=pl_data)

print("BS・PLデータをGoogle Sheetsにアップロード完了")

