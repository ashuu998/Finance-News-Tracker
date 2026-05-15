print("Project started!")
print("Now we will install libraries")

# Install Libraries


import requests
API_KEY = "204cf77b5d6040ae8fdd9822faacff80" 
url = "https://newsapi.org/v2/everything"

params = {
    "q"        : "stock market india",
    "language" : "en",
    "pageSize" : 10,
    "apiKey"   : API_KEY
}

response = requests.get(url, params=params)
data = response.json()

print(f"Total news recieved: {data['totalResults']}")
print("---")

for article in data["articles"]:
    print(article["title"])
    print("---")


import requests
import pandas as pd
from datetime import datetime

API_KEY = "204cf77b5d6040ae8fdd9822faacff80" 

url = "https://newsapi.org/v2/everything"

params = {
    "q"        : "stock market Of india",
    "language" : "en",
    "pageSize" : 10,
    "apiKey"   : API_KEY
}

response = requests.get(url, params=params)
data = response.json()

# Data stored in list
news_list = []

for article in data["articles"]:
    news_list.append({
        "title"       : article["title"],
        "description" : article["description"],
        "source"      : article["source"]["name"],
        "published"   : article["publishedAt"],
        "url"         : article["url"]
    })

# Created Pandas DataFrame 
df = pd.DataFrame(news_list)

#saved in CSV
df.to_csv("news_data.csv", index=False)

print(f"Total news saved: {len(df)}")
print(df.head())


# =====================
# STEP 3 — SENTIMENT
# =====================
from textblob import TextBlob

def get_sentiment(text):
    if text is None:
        return "Unknown"
    
    score = TextBlob(text).sentiment.polarity
    
    if score > 0.1:
        return "Positive"
    elif score < -0.1:
        return "Negative"
    else:
        return "Neutral"

# Har news ka sentiment check karo
df["sentiment"] = df["title"].apply(get_sentiment)

# Count karo
print("\n--- Sentiment Summary ---")
print(df["sentiment"].value_counts())

# CSV update karo
df.to_csv("news_data.csv", index=False)
print("\nCSV updated with sentiment!")
print(df[["title", "sentiment"]])

# =====================
# STEP 4 — EXCEL REPORT
# =====================
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

wb = Workbook()
ws = wb.active
ws.title = "News Report"

# Header banao
headers = ["Title", "Source", "Published", "Sentiment", "URL"]
ws.append(headers)

# Header styling
for cell in ws[1]:
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill(fill_type="solid", fgColor="000080")
    cell.alignment = Alignment(horizontal="center")

# Data add karo
for _, row in df.iterrows():
    ws.append([
        row["title"],
        row["source"],
        row["published"],
        row["sentiment"],
        row["url"]
    ])

# Sentiment color coding
for row in ws.iter_rows(min_row=2):
    sentiment = row[3].value
    if sentiment == "Positive":
        color = "90EE90"    # Green
    elif sentiment == "Negative":
        color = "FFB6B6"    # Red
    else:
        color = "FFFFE0"    # Yellow

    for cell in row:
        cell.fill = PatternFill(
            fill_type="solid",
            fgColor=color
        )

# Column width
ws.column_dimensions["A"].width = 50
ws.column_dimensions["B"].width = 20
ws.column_dimensions["C"].width = 25
ws.column_dimensions["D"].width = 15
ws.column_dimensions["E"].width = 40

# Save karo
wb.save("news_report.xlsx")
print("Excel report ready!")


# =====================
# STEP 4 — EXCEL REPORT
# =====================
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

wb = Workbook()
ws = wb.active
ws.title = "News Report"

# Created Header 
headers = ["Title", "Source", "Published", "Sentiment", "URL"]
ws.append(headers)

# Header styling
for cell in ws[1]:
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill(fill_type="solid", fgColor="000080")
    cell.alignment = Alignment(horizontal="center")

# Added Data 
for _, row in df.iterrows():
    ws.append([
        row["title"],
        row["source"],
        row["published"],
        row["sentiment"],
        row["url"]
    ])

# Sentiment color coding
for row in ws.iter_rows(min_row=2):
    sentiment = row[3].value
    if sentiment == "Positive":
        color = "90EE90"    # Green
    elif sentiment == "Negative":
        color = "FFB6B6"    # Red
    else:
        color = "FFFFE0"    # Yellow

    for cell in row:
        cell.fill = PatternFill(
            fill_type="solid",
            fgColor=color
        )

# Column width
ws.column_dimensions["A"].width = 50
ws.column_dimensions["B"].width = 20
ws.column_dimensions["C"].width = 25
ws.column_dimensions["D"].width = 15
ws.column_dimensions["E"].width = 40

# Saved
wb.save("news_report.xlsx")
print("Excel report ready!")



# =====================
# STEP 5 — DASHBOARD
# =====================
from datetime import datetime

total    = len(df)
positive = len(df[df["sentiment"] == "Positive"])
negative = len(df[df["sentiment"] == "Negative"])
neutral  = len(df[df["sentiment"] == "Neutral"])

if positive > negative:
    mood  = "BULLISH 📈"
    color = "Market positive!"
elif negative > positive:
    mood  = "BEARISH 📉"
    color = "Market negative!"
else:
    mood  = "NEUTRAL ➡️"
    color = "Market stable!"

print("\n" + "="*45)
print("   FINANCIAL NEWS DASHBOARD")
print("="*45)
print(f"Date          : {datetime.now().strftime('%d %B %Y')}")
print(f"Total News    : {total}")
print("-"*45)
print(f"Positive      : {positive} 🟢")
print(f"Negative      : {negative} 🔴")
print(f"Neutral       : {neutral} 🟡")
print("-"*45)
print(f"Market Mood   : {mood}")
print(f"Analysis      : {color}")
print("="*45)
print(f"Excel Report  : news_report.xlsx")
print(f"CSV Data      : news_data.csv")
print("="*45)
