import sys
import json
import time
from PyQt5.QtWidgets import QApplication
from PyQt5.QAxContainer import QAxWidget

class Kiwoom:
    def __init__(self):
        self.app = QApplication(sys.argv)
        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")

        self.ocx.OnEventConnect.connect(self._on_login)
        self.ocx.OnReceiveTrData.connect(self._on_receive_tr_data)

        self.login_ok = False
        self.tr_data = None
        self.data_ready = False

    def connect(self):
        self.ocx.dynamicCall("CommConnect()")
        self.app.exec_()


    def _on_login(self, err_code):
        if err_code == 0:
            print(">> 로그인 성공")
            self.login_ok = True
        else:
            print(">> 로그인 실패")
        self.app.quit()

    def get_code_list(self):
        kospi = self.ocx.dynamicCall("GetCodeListByMarket(QString)", ["0"]).split(';')
        kosdaq = self.ocx.dynamicCall("GetCodeListByMarket(QString)", ["10"]).split(';')
        codes = list(filter(None, kospi + kosdaq))
        print(f">> 총 종목 수: {len(codes)}")
        return codes

    def get_stock_name(self, code):
        return self.ocx.dynamicCall("GetMasterCodeName(QString)", [code])


    def is_valid_stock(self, code, name):
        # 1. 종목명 키워드 필터링
        keywords = ['우', 'ETF', 'ETN', '리츠', '스팩']
        if any(k in name for k in keywords):
            return False
        # 2. 종목 상태 필터링
        status = self.ocx.dynamicCall("GetMasterStockState(QString)", [code])
        exclude_status = ["관리", "정지", "환기", "정리", "불성실", "유의"]
        if any(s in status for s in exclude_status):
            return False
        # 3. 종목 구분 필터링(6:증거금100, 12:초저유동성)
        construction = self.ocx.dynamicCall("GetConstructionType(QString)", [code])
        if construction in ["6", "12"]:
            return False
        return True


    def _on_receive_tr_data(self, screen_no, rqname, trcode, recordname, prev_next):
        count = self.ocx.dynamicCall("GetRepeatCnt(QString, QString)", trcode, rqname)
        result = []

        for i in range(count):
            date = self.ocx.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, rqname, i, "일자").strip()
            open_ = self.ocx.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, rqname, i, "시가").strip()
            high = self.ocx.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, rqname, i, "고가").strip()
            low = self.ocx.dynamicCall("GetCommData(QString, QString, int, QString)",  trcode, rqname, i, "저가").strip()
            close = self.ocx.dynamicCall("GetCommData(QString, QString, int, QString)",trcode, rqname, i, "현재가").strip()
            volume = self.ocx.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, rqname, i, "거래량").strip()

            result.append({
                "date": date,
                "open": int(open_),
                "high": int(high),
                "low": int(low),
                "close": int(close),
                "volume": int(volume)
            })

        self.tr_data = result[:360]  # 최대 (360개) 저장
        self.data_ready = True


if __name__ == "__main__":
    kiwoom = Kiwoom()
    kiwoom.connect()

    all_data = {}
    codes = kiwoom.get_code_list()

    for idx, code in enumerate(codes):
        name = kiwoom.get_stock_name(code)
        if not kiwoom.is_valid_stock(code, name):
            continue

        print(f"[{idx+1}/{len(codes)}] {code} {name} 데이터 요청 중...")
        try:
            ohlcv = kiwoom.get_ohlcv(code)
            all_data[code] = {
                "name": name,
                "ohlcv": ohlcv
            }
        except Exception as e:
            print(f"  >> {code} {name} 실패: {e}")
        time.sleep(0.3)  # 과도한 요청 방지

    with open("all_stock_data.json", "w", encoding="utf-8") as f:
        json.dump(all_data, f, indent=2, ensure_ascii=False)

    print("✅ 모든 데이터 저장 완료: all_stock_data.json")
