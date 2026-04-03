import json
import pandas as pd

# 파일 불러오기
with open("all_stock_data.json", "r", encoding="utf-8") as f:
    data = json.load(f)

selected = []

def EMA(series, span):
    return series.ewm(span=span, adjust=False).mean()


for code, info in data.items():
    name = info["name"]
    ohlcv = pd.DataFrame(info["ohlcv"])

    if len(ohlcv) < 60:  # 데이터 부족시 스킵
        continue

 # === [1] 주봉 20 이동평균과 종가 만나는 조건 ===
    ohlcv["date"] = pd.to_datetime(ohlcv["date"], format="%Y%m%d")
    ohlcv.set_index("date", inplace=True)
    weekly = ohlcv.resample("W").last()  # 주봉 데이터 (마지막 거래일 종가)
    weekly["ma20"] = weekly["close"].rolling(20).mean()

    if len(weekly.dropna()) == 0:
        continue

    latest_week_close = weekly["close"].iloc[-1]
    latest_week_ma20 = weekly["ma20"].iloc[-1]

    cond1 = abs(latest_week_close - latest_week_ma20) / latest_week_ma20 < 0.01  # 1% 이내 차이

# === [2] 일봉 5일선 20일선 정배열 ===
    ohlcv["ma5"] = ohlcv["close"].rolling(5).mean()
    ohlcv["ma20"] = ohlcv["close"].rolling(20).mean()
    cond2 = ohlcv["ma5"].iloc[-1] > ohlcv["ma20"].iloc[-1]

# === [3] MACD oscillator > 0 ===
    close = ohlcv["close"]
    macd = EMA(close, 12) - EMA(close, 26)
    signal = EMA(macd, 9)
    osc = macd - signal
    cond3 = osc.iloc[-1] > 0

# 모든 조건 충족
    if cond1 and cond2 and cond3:
        close_price = int(ohlcv["close"].iloc[-1])
        selected.append(f"{code} {name} {close_price}")

# TXT 파일 저장
with open("selected.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(selected))

print("조건에 맞는 종목이 selected.txt에 저장되었습니다.")