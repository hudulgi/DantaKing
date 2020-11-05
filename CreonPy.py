import win32com.client
import ctypes


################################################
# PLUS 공통 OBJECT
g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')


################################################
# PLUS 실행 기본 체크 함수
def InitPlusCheck():
    # 프로세스가 관리자 권한으로 실행 여부
    if ctypes.windll.shell32.IsUserAnAdmin():
        print('정상: 관리자권한으로 실행된 프로세스입니다.')
    else:
        print('오류: 일반권한으로 실행됨. 관리자 권한으로 실행해 주세요')
        return False

    # 연결 여부 체크
    if (g_objCpStatus.IsConnect == 0):
        print("PLUS가 정상적으로 연결되지 않음. ")
        return False

    # 주문 관련 초기화
    if (g_objCpTrade.TradeInit(0) != 0):
        print("주문 초기화 실패")
        return False

    return True


################################################
# 현재가 - 한종목 통신
class CpRPCurrentPrice:
    def __init__(self):
        self.objStockMst = win32com.client.Dispatch('DsCbo1.StockMst')
        return

    def Request(self, code):
        self.objStockMst.SetInputValue(0, code)
        ret = self.objStockMst.BlockRequest()
        if self.objStockMst.GetDibStatus() != 0:
            print('통신상태', self.objStockMst.GetDibStatus(), self.objStockMst.GetDibMsg1())
            return False

        item = {}
        item['code'] = code
        item['종목명'] = self.objStockMst.GetHeaderValue(1)  # 종목명
        item['cur'] = self.objStockMst.GetHeaderValue(11)  # 현재가
        item['diff'] = self.objStockMst.GetHeaderValue(12)  # 전일대비
        item['vol'] = self.objStockMst.GetHeaderValue(18)  # 거래량
        item['예상플래그'] = chr(self.objStockMst.GetHeaderValue(58))  # 예상플래그
        item['예상체결가'] = self.objStockMst.GetHeaderValue(55)  # 예상체결가

        '''
        caller.curData['기준가'] = self.objStockMst.GetHeaderValue(27)  # 기준가
        caller.curData['예상플래그'] = self.objStockMst.GetHeaderValue(58)  # 예상플래그
        caller.curData['예상체결가'] = self.objStockMst.GetHeaderValue(55)  # 예상체결가
        caller.curData['예상대비'] = self.objStockMst.GetHeaderValue(56)  # 예상체결대비
        '''

        return item


################################################
# 주식 주문 처리
class CpRPOrder:
    def __init__(self, acc):
        self.acc = acc
        self.accFlag = g_objCpTrade.GoodsList(self.acc, 1)  # 주식상품 구분
        print(self.acc, self.accFlag[0])
        self.objOrder = win32com.client.Dispatch("CpTrade.CpTd0311")  # 매수

    def buy_order(self, code, amount):
        # 주식 매수 주문
        #print("신규 매수", code, price, amount)
        #self.caller.textBrowser.append("매수 %s %d %d" % (code, price, amount))

        self.objOrder.SetInputValue(0, "2")  # 2: 매수
        self.objOrder.SetInputValue(1, self.acc)  # 계좌번호
        self.objOrder.SetInputValue(2, self.accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objOrder.SetInputValue(3, code)  # 종목코드
        self.objOrder.SetInputValue(4, amount)  # 매수수량
        #self.objOrder.SetInputValue(5, price)  # 주문단가
        self.objOrder.SetInputValue(7, "0")  # 주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
        self.objOrder.SetInputValue(8, "03")  # 주문호가 구분코드 - 03: 시장가

        # 매수 주문 요청
        ret = self.objOrder.BlockRequest()
        if ret == 4:
            remainTime = g_objCpStatus.LimitRequestRemainTime
            print('주의: 주문 연속 통신 제한에 걸렸음. 대기해서 주문할 지 여부 판단이 필요 남은 시간', remainTime)
            return False

        rqStatus = self.objOrder.GetDibStatus()
        rqRet = self.objOrder.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False

        return True
