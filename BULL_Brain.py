import sys                        # system specific parameters and functions : 파이썬 스크립트 관리
from PyQt5.QtWidgets import *     # GUI의 그래픽적 요소를 제어       하단의 terminal 선택, activate py37_32,  pip install pyqt5,   전부다 y
from PyQt5 import uic             # ui 파일을 가져오기위한 함수
from PyQt5.QtCore import *        # QEventLoop를 사용하기 위해 불러오자

##kiwoom.py 스크립트에 있는 Kiwoom 클래스를 가져온다

from kiwoom import Kiwoom          # 키움증권 함수/공용 방 (싱글턴)
from Bthread_1 import Thread1      # 계좌평가잔고내역 가져오기
from Bthread_2 import Thread2      # gitHub에서 오늘의 종목을 불러오는 Thread


#=================== 프로그램 실행 하는 부분 =========================#

form_class = uic.loadUiType("BULL.ui")[0]             # 만들어 놓은 ui 불러오기

class BULL_Brain(QMainWindow, QWidget, form_class):       # QMainWindow, PyQt5에서 윈도우 생성시 필요한 함수정의

    def __init__(self, *args, **kwargs):                      # Main class의 self를 초기화 하고 시작

        print("Login Machine 실행합니다.")
        super(BULL_Brain, self).__init__(*args, **kwargs)
        form_class.__init__(self)                            # 상속 받은 from_class(BULL_Brain에서)를 실행하기 위한 초기값(초기화)
        self.setUI()                                         # ui  초기값을 셋업~~꼭 해주길

        ### 초기 셋팅
        self.label_3.setText(str("총매입금액"))
        self.label_4.setText(str("총평가금액"))
        self.label_5.setText(str("추정예탁자산"))
        self.label_6.setText(str("총평가손익금액"))
        self.label_7.setText(str("총수익률(%)"))

        #### git에서 txt 가지고 오기
        self.selected_url = "주소입력"


        #### 기타 함수
        self.login_event_loop = QEventLoop()  #  QEventLoop()는 block 기능을 가지고 있음...!!

        ####키움증권 로그인 하기
        self.k = Kiwoom()
        self.set_signal_slot()
        self.signal_login_commConnect()

        #####이벤트 생성 및 진행 (Bthread_1의 액션, 계좌정보 확인 버튼 연결)
        self.call_account.clicked.connect(self.B_acc)         # 계좌정보가져오기

        #####깃 허브에서 오늘의 종목 가지고오기
        self.call_selectedlist.clicked.connect(self.load_selected_list)



    def setUI(self):
        self.setupUi(self)                # UI 초기값 셋업

    def set_signal_slot(self):
        self.k.kiwoom.OnEventConnect.connect(self.login_slot)  # 내가 알고 있는 login_slot에다가 특정 값을 던져 준다.

    def signal_login_commConnect(self):
        self.k.kiwoom.dynamicCall("CommConnect()")  # 네트워크적 서버 응용프로그램에 데이터를 전송할 수 있게 만든 함수
        self.login_event_loop.exec_()  # 로그인이 완료될 때까지 계속 반복됨. 꺼지지 않음.

    def login_slot(self, errCode):
        if errCode == 0:
            print("로그인 성공")
            self.statusbar.showMessage("로그인 성공")
            self.get_account_info()                    # 로그인시 계좌정보 가져오기

        elif errCode == 100:
            print("사용자 정보교환 실패")
        elif errCode == 101:
            print("서버접속 실패")
        elif errCode == 102:
            print("버전처리 실패")
        self.login_event_loop.exit()  # 로그인이 완료되면 로그인 창을 닫는다.

    def get_account_info(self):
        account_list = self.k.kiwoom.dynamicCall("GetLoginInfo(String)", "ACCNO")

        for n in account_list.split(';'):
            self.BULLacc.addItem(n)

    def B_acc(self):
        print("선택 계좌 정보 가져오기")
        ##### 1번 일꾼 실행(Bthread_1 실행)
        h1 = Thread1(self)
        h1.start()

    def load_selected_list(self):
            print("오늘의 종목 불러오기")
            ##Bthread_2 실행
            self.thread2 = Thread2(self.selected_url)
            self.thread2.data_ready.connect(self.update_selected_table)
            self.thread2.start()
        
    def update_selected_table(self, data):
        table = self.selectedstockTableWidget_2
        table.clearContents()
        table.setRowCount(len(data))
        table.setColumnCount(3)
        table.setHorizontalHeaderLabels(["종목코드", "종목명", "종가"])

        for row, (code, name, last_close) in enumerate(data):
            table.setItem(row, 0, QTableWidgetItem(code))
            table.setItem(row, 1, QTableWidgetItem(name))
            table.setItem(row, 2, QTableWidgetItem(str(last_close)))




if __name__=='__main__':             # import된 것들을 실행시키지 않고 __main__에서 실행하는 것만 실행 시킨다.
                                     # 즉 import된 다른 함수의 코드를 이 화면에서 실행시키지 않겠다는 의미이다.

    app = QApplication(sys.argv)     # PyQt5로 실행할 파일명을 자동으로 설정, PyQt5에서 자동으로 프로그램 실행
    RS = BULL_Brain()                  # Main 클래스 myApp으로 인스턴스화
    RS.show()                        # myApp에 있는 ui를 실행한다.
    app.exec_()                      # 이벤트 루프
