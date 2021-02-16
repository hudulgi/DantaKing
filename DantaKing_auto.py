import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5 import uic
import win32com.client
import ctypes
import pandas as pd
import datetime
import time
from dt_alimi import *
from multiprocessing import Process, Queue
import os

form_class = uic.loadUiType("DantaKing.ui")[0]
data_path = "C:\\CloudStation\\dt_data"
buy_path = data_path + "\\daily_data\\buy_list"
target_path = data_path + "\\daily_data\\target_list"
unit_price = 4500000  # 종목 당 매수금액
today = datetime.datetime.now().date()

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
# Telegram 메세지 전송 함수
def telegram(msg):
    bot.sendMessage(myId, msg)
    return True


def telegram_ch(msg):
    bot.sendMessage("@dantaking_ch", msg)
    return True


def telegram_share(msg):
    bot.sendMessage(ch2, msg)
    return True

################################################
# CpEvent: 실시간 이벤트 수신 클래스
class CpEvent:
    def set_params(self, client, name, caller):
        self.client = client  # CP 실시간 통신 object
        self.name = name  # 서비스가 다른 이벤트를 구분하기 위한 이름
        self.caller = caller  # callback 을 위해 보관

        # 구분값 : 텍스트로 변경하기 위해 딕셔너리 이용
        self.dicflag12 = {'1': '매도', '2': '매수'}
        self.dicflag14 = {'1': '체결', '2': '확인', '3': '거부', '4': '접수'}
        self.dicflag15 = {'00': '현금', '01': '유통융자', '02': '자기융자', '03': '유통대주',
                          '04': '자기대주', '05': '주식담보대출', '07': '채권담보대출',
                          '06': '매입담보대출', '08': '플러스론',
                          '13': '자기대용융자', '15': '유통대용융자'}
        self.dicflag16 = {'1': '정상주문', '2': '정정주문', '3': '취소주문'}
        self.dicflag17 = {'1': '현금', '2': '신용', '3': '선물대용', '4': '공매도'}
        self.dicflag18 = {'01': '보통', '02': '임의', '03': '시장가', '05': '조건부지정가'}
        self.dicflag19 = {'0': '없음', '1': 'IOC', '2': 'FOK'}

    def OnReceived(self):
        # 실시간 처리 - 현재가 주문 체결
        if self.name == 'stockcur':
            code = self.client.GetHeaderValue(0)  # 초
            name = self.client.GetHeaderValue(1)  # 초
            timess = self.client.GetHeaderValue(18)  # 초
            exFlag = self.client.GetHeaderValue(19)  # 예상체결 플래그
            cprice = self.client.GetHeaderValue(13)  # 현재가
            diff = self.client.GetHeaderValue(2)  # 대비
            cVol = self.client.GetHeaderValue(17)  # 순간체결수량
            vol = self.client.GetHeaderValue(9)  # 거래량

            if (exFlag == ord('1')):  # 동시호가 시간 (예상체결)
                self.caller.textBrowser2.append(
                    "실시간(예상체결) %s %s %d 대비 %d 체결량 %d 거래량 %d" % (name, timess, cprice, diff, cVol, vol))
            #    print("실시간(예상체결)", name, timess, "*", cprice, "대비", diff, "체결량", cVol, "거래량", vol)
            elif (exFlag == ord('2')):  # 장중(체결)
                self.caller.textBrowser2.append(
                    "실시간(장중체결) %s %s %d 대비 %d 체결량 %d 거래량 %d" % (name, timess, cprice, diff, cVol, vol))
            #    print("실시간(장중 체결)", name, timess, cprice, "대비", diff, "체결량", cVol, "거래량", vol)

            item = {}
            item['code'] = code
            # rpName = self.objRq.GetDataValue(1, i)  # 종목명
            # rpDiffFlag = self.objRq.GetDataValue(3, i)  # 대비부호
            item['diff'] = diff
            item['cur'] = cprice
            item['vol'] = vol

            # 현재가 업데이트
            self.caller.updateCurPBData(item)

        # 실시간 처리 - 주문체결
        elif self.name == 'conclution':
            # 주문 체결 실시간 업데이트
            conc = {}

            # 체결 플래그
            conc['체결플래그'] = self.dicflag14[self.client.GetHeaderValue(14)]

            conc['주문번호'] = self.client.GetHeaderValue(5)  # 주문번호
            conc['주문수량'] = self.client.GetHeaderValue(3)  # 주문/체결 수량
            conc['주문가격'] = self.client.GetHeaderValue(4)  # 주문/체결 가격
            conc['원주문'] = self.client.GetHeaderValue(6)
            conc['종목코드'] = self.client.GetHeaderValue(9)  # 종목코드
            conc['종목명'] = g_objCodeMgr.CodeToName(conc['종목코드'])

            conc['매수매도'] = self.dicflag12[self.client.GetHeaderValue(12)]

            flag15 = self.client.GetHeaderValue(15)  # 신용대출구분코드
            if (flag15 in self.dicflag15):
                conc['신용대출'] = self.dicflag15[flag15]
            else:
                conc['신용대출'] = '기타'

            conc['정정취소'] = self.dicflag16[self.client.GetHeaderValue(16)]
            conc['현금신용'] = self.dicflag17[self.client.GetHeaderValue(17)]
            conc['주문조건'] = self.dicflag19[self.client.GetHeaderValue(19)]

            conc['체결기준잔고수량'] = self.client.GetHeaderValue(23)
            loandate = self.client.GetHeaderValue(20)
            if (loandate == 0):
                conc['대출일'] = ''
            else:
                conc['대출일'] = str(loandate)
            flag18 = self.client.GetHeaderValue(18)
            if (flag18 in self.dicflag18):
                conc['주문호가구분'] = self.dicflag18[flag18]
            else:
                conc['주문호가구분'] = '기타'

            conc['장부가'] = self.client.GetHeaderValue(21)
            conc['매도가능수량'] = self.client.GetHeaderValue(22)

            info_txt = "%s %s %s %s %s" % (conc['체결플래그'], conc['매수매도'], conc['종목명'], conc['주문수량'], conc['주문가격'])
            self.caller.textBrowser.append(info_txt)
            telegram(info_txt)
            self.caller.updateJangoCont(conc)

            return


################################################
# plus 실시간 수신 base 클래스
class CpPublish:
    def __init__(self, name, serviceID):
        self.name = name
        self.obj = win32com.client.Dispatch(serviceID)
        self.bIsSB = False

    def Subscribe(self, var, caller):
        if self.bIsSB:
            self.Unsubscribe()

        if (len(var) > 0):
            self.obj.SetInputValue(0, var)

        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self.obj, self.name, caller)
        self.obj.Subscribe()
        self.bIsSB = True

    def Unsubscribe(self):
        if self.bIsSB:
            self.obj.Unsubscribe()
        self.bIsSB = False


################################################
# CpPBStockCur: 실시간 현재가 요청 클래스
class CpPBStockCur(CpPublish):
    def __init__(self):
        super().__init__('stockcur', 'DsCbo1.StockCur')


################################################
# CpPBConclusion: 실시간 주문 체결 수신 클래그
class CpPBConclusion(CpPublish):
    def __init__(self):
        super().__init__('conclution', 'DsCbo1.CpConclusion')


################################################
# Cp6033 : 주식 잔고 조회
class Cp6033:
    def __init__(self):
        acc = g_objCpTrade.AccountNumber[0]  # 계좌번호
        accFlag = g_objCpTrade.GoodsList(acc, 1)  # 주식상품 구분
        print(acc, accFlag[0])

        self.objRq = win32com.client.Dispatch("CpTrade.CpTd6033")
        self.objRq.SetInputValue(0, acc)  # 계좌번호
        self.objRq.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objRq.SetInputValue(2, 50)  # 요청 건수(최대 50)
        self.dicflag1 = {ord(' '): '현금',
                         ord('Y'): '융자',
                         ord('D'): '대주',
                         ord('B'): '담보',
                         ord('M'): '매입담보',
                         ord('P'): '플러스론',
                         ord('I'): '자기융자',
                         }

    # 실제적인 6033 통신 처리
    def requestJango(self, caller):
        while True:
            self.objRq.BlockRequest()
            # 통신 및 통신 에러 처리
            rqStatus = self.objRq.GetDibStatus()
            rqRet = self.objRq.GetDibMsg1()
            print("통신상태", rqStatus, rqRet)
            if rqStatus != 0:
                return False

            cnt = self.objRq.GetHeaderValue(7)
            print(cnt)

            for i in range(cnt):
                item = {}
                code = self.objRq.GetDataValue(12, i)  # 종목코드
                item['종목코드'] = code
                item['종목명'] = self.objRq.GetDataValue(0, i)  # 종목명
                item['현금신용'] = self.dicflag1[self.objRq.GetDataValue(1, i)]  # 신용구분
                item['대출일'] = self.objRq.GetDataValue(2, i)  # 대출일
                item['잔고수량'] = self.objRq.GetDataValue(7, i)  # 체결잔고수량
                item['매도가능'] = self.objRq.GetDataValue(15, i)
                item['장부가'] = self.objRq.GetDataValue(17, i)  # 체결장부단가
                # item['평가금액'] = self.objRq.GetDataValue(9, i)  # 평가금액(천원미만은 절사 됨)
                # item['평가손익'] = self.objRq.GetDataValue(11, i)  # 평가손익(천원미만은 절사 됨)
                # 매입금액 = 장부가 * 잔고수량
                item['매입금액'] = item['장부가'] * item['잔고수량']
                item['현재가'] = 0
                item['대비'] = 0
                item['거래량'] = 0

                # 잔고 추가
                #                key = (code, item['현금신용'],item['대출일'] )
                key = code
                caller.jangoData[key] = item

                if len(caller.jangoData) >= 200:  # 최대 200 종목만,
                    break

            if len(caller.jangoData) >= 200:
                break
            if (self.objRq.Continue == False):
                break
        return True


################################################
# 현재가 - 한종목 통신
class CpRPCurrentPrice:
    def __init__(self):
        self.objStockMst = win32com.client.Dispatch('DsCbo1.StockMst')
        return

    def Request(self, code, caller):
        self.objStockMst.SetInputValue(0, code)
        ret = self.objStockMst.BlockRequest()
        if self.objStockMst.GetDibStatus() != 0:
            print('통신상태', self.objStockMst.GetDibStatus(), self.objStockMst.GetDibMsg1())
            return False

        item = {}
        item['code'] = code
        # caller.curData['종목명'] = g_objCodeMgr.CodeToName(code)
        item['cur'] = self.objStockMst.GetHeaderValue(11)  # 현재가
        item['diff'] = self.objStockMst.GetHeaderValue(12)  # 전일대비
        item['vol'] = self.objStockMst.GetHeaderValue(18)  # 거래량

        '''
        caller.curData['기준가'] = self.objStockMst.GetHeaderValue(27)  # 기준가
        caller.curData['예상플래그'] = self.objStockMst.GetHeaderValue(58)  # 예상플래그
        caller.curData['예상체결가'] = self.objStockMst.GetHeaderValue(55)  # 예상체결가
        caller.curData['예상대비'] = self.objStockMst.GetHeaderValue(56)  # 예상체결대비
        '''
        # 10차호가
        for i in range(10):
            key1 = 'offer%d' % (i + 1)
            key2 = 'bid%d' % (i + 1)
            item[key1] = (self.objStockMst.GetDataValue(0, i))  # 매도호가
            item[key2] = (self.objStockMst.GetDataValue(1, i))  # 매수호가

        caller.curDatas[code] = item

        return True


################################################
# CpMarketEye : 복수종목 현재가 통신 서비스
class CpMarketEye:
    def __init__(self):
        # 요청 필드 배열 - 종목코드, 시간, 대비부호 대비, 현재가, 거래량, 종목명
        self.rqField = [0, 1, 2, 3, 4, 10, 17]  # 요청 필드

        # 관심종목 객체 구하기
        self.objRq = win32com.client.Dispatch("CpSysDib.MarketEye")

    def Request(self, codes, caller):
        # 요청 필드 세팅 - 종목코드, 종목명, 시간, 대비부호, 대비, 현재가, 거래량
        self.objRq.SetInputValue(0, self.rqField)  # 요청 필드
        self.objRq.SetInputValue(1, codes)  # 종목코드 or 종목코드 리스트
        self.objRq.BlockRequest()

        # 현재가 통신 및 통신 에러 처리
        rqStatus = self.objRq.GetDibStatus()
        rqRet = self.objRq.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False

        cnt = self.objRq.GetHeaderValue(2)

        for i in range(cnt):
            item = {}
            item['code'] = self.objRq.GetDataValue(0, i)  # 코드
            # rpName = self.objRq.GetDataValue(1, i)  # 종목명
            # rpDiffFlag = self.objRq.GetDataValue(3, i)  # 대비부호
            item['diff'] = self.objRq.GetDataValue(3, i)  # 대비
            item['cur'] = self.objRq.GetDataValue(4, i)  # 현재가
            item['vol'] = self.objRq.GetDataValue(5, i)  # 거래량

            caller.curDatas[item['code']] = item

        return True


################################################
# 주식 주문 처리
class CpRPOrder:
    def __init__(self):
        self.acc = g_objCpTrade.AccountNumber[0]  # 계좌번호
        self.accFlag = g_objCpTrade.GoodsList(self.acc, 1)  # 주식상품 구분
        #print(self.acc, self.accFlag[0])
        self.objOrder = win32com.client.Dispatch("CpTrade.CpTd0311")  # 매수

    def buyOrder(self, code, price, amount):
        # 주식 매수 주문
        #print("신규 매수", code, price, amount)

        self.objOrder.SetInputValue(0, "2")  # 2: 매수
        self.objOrder.SetInputValue(1, self.acc)  # 계좌번호
        self.objOrder.SetInputValue(2, self.accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objOrder.SetInputValue(3, code)  # 종목코드
        self.objOrder.SetInputValue(4, amount)  # 매수수량
        self.objOrder.SetInputValue(5, price)  # 주문단가
        self.objOrder.SetInputValue(7, "0")  # 주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
        self.objOrder.SetInputValue(8, "01")  # 주문호가 구분코드 - 01: 보통

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

    def sellOrder(self, code, price, amount):
        # 주식 매도 주문
        #print("신규 매도", code, price, amount)
        #self.caller.textBrowser.append("매도 %s %d %d" % (code, price, amount))

        self.objOrder.SetInputValue(0, "1")  # 1: 매도
        self.objOrder.SetInputValue(1, self.acc)  # 계좌번호
        self.objOrder.SetInputValue(2, self.accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objOrder.SetInputValue(3, code)  # 종목코드
        self.objOrder.SetInputValue(4, amount)  # 매수수량
        self.objOrder.SetInputValue(5, price)  # 주문단가
        self.objOrder.SetInputValue(7, "0")  # 주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
        self.objOrder.SetInputValue(8, "01")  # 주문호가 구분코드 - 01: 보통

        # 매도 주문 요청
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


################################################
# 테스트를 위한 메인 화면
class MyWindow(QMainWindow, form_class):
    def __init__(self, q):
        super().__init__()
        self.setupUi(self)

        # plus 상태 체크
        if InitPlusCheck() == False:
            exit()

        # 6033 잔고 object
        self.obj6033 = Cp6033()
        self.jangoData = {}

        self.isSB = False
        self.objCur = {}

        # 현재가 정보
        self.curDatas = {}
        self.objRPCur = CpRPCurrentPrice()

        # 주문
        #self.objRPOrder = CpRPOrder()  # 멀티프로세스 방식으로 변경하여 불필요

        # 실시간 주문 체결
        self.objConclusion = CpPBConclusion()

        self.btnStart.clicked.connect(self.StartWatch)
        self.btnStop.clicked.connect(self.StopSubscribe)
        self.btnExit.clicked.connect(self.close)

        # 잔고 요청
        self.obj6033.requestJango(self)
        self.printJango()

        # 타겟 리스트 로드
        self.target_data = self.import_targets()

        # 타이머
        self.timer = QTimer(self)
        self.timer.start(1000)
        self.timer.timeout.connect(self.timeout)

        # 감시시작
        self.StartWatch()

        # 텔레그램 메세지 전송
        telegram("단타킹 프로그램 시작")

        # 주문 queue
        self.q = q

        # 물타기 리스트
        self.multagi = list()

    def closeEvent(self, QCloseEvent):
        print('close')
        self.StopSubscribe()
        self.q.put(None)
        self.deleteLater()
        QCloseEvent.accept()

    def timeout(self):
        current_time = QTime.currentTime()
        text_time = current_time.toString("hh:mm:ss")
        time_msg = "현재시간: " + text_time

        state = g_objCpStatus.IsConnect
        if state == 1:
            state_msg = "서버 연결 중"
        else:
            state_msg = "서버 미 연결 중"

        self.statusbar.showMessage(state_msg + " | " + time_msg)

    def import_targets(self):
        dt = datetime.datetime.now()
        now_day = dt.strftime("%Y%m%d")
        file_name = target_path + "\\target_list_" + now_day[2:] + ".csv"
        data_df = pd.read_csv(file_name)
        data_dict = data_df.to_dict(orient='records')
        data_result = dict()
        for data in data_dict:
            data_result[data['code']] = {'OBJ': data['OBJ'],
                                         'OBJ2': data['OBJ2'],
                                         'name': data['name'],
                                         'scale': data['scale'],
                                         '체결상태': 0,
                                         '주문상태': 0}
            self.textBrowser.append("%s %s %s" %(data['code'], data['name'], data['OBJ']))
            if data['code'] in self.jangoData:
                data_result[data['code']]['주문상태'] = 1
                data_result[data['code']]['체결상태'] = 1
        return data_result

    def StopSubscribe(self):
        if self.isSB:
            for key, obj in self.objCur.items():
                obj.Unsubscribe()
            self.objCur = {}

        self.isSB = False
        self.objConclusion.Unsubscribe()

    def StartWatch(self):
        self.StopSubscribe()

        codes = list(set(self.target_data.keys()) | set(self.jangoData.keys()))

        #objMarkeyeye = CpMarketEye()
        #if (objMarkeyeye.Request(codes, self) == False):
        #    exit()

        self.textBrowser.append("실시간 감시를 시작합니다.")
        # 실시간 현재가  요청
        for code in codes:
            self.objCur[code] = CpPBStockCur()
            self.objCur[code].Subscribe(code, self)
        self.isSB = True

        # 실시간 주문 체결 요청
        self.objConclusion.Subscribe('', self)

    def printJango(self):
        item_count = len(self.jangoData)
        self.tableWidget_jango.setRowCount(item_count)
        for row, value in enumerate(self.jangoData.values()):
            item = QTableWidgetItem(value['종목코드'])
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            self.tableWidget_jango.setItem(row, 0, item)

            item = QTableWidgetItem(value['종목명'])
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            self.tableWidget_jango.setItem(row, 1, item)

            item = QTableWidgetItem(str(value['잔고수량']))
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.tableWidget_jango.setItem(row, 2, item)

            item = QTableWidgetItem(str(value['매도가능']))
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.tableWidget_jango.setItem(row, 3, item)

            item = QTableWidgetItem("{0:.0f}".format(value['매입금액']))
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.tableWidget_jango.setItem(row, 4, item)

            item = QTableWidgetItem("{0:.0f}".format(value['장부가']))
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.tableWidget_jango.setItem(row, 5, item)

        self.tableWidget_jango.resizeRowsToContents()
        self.tableWidget_jango.resizeColumnsToContents()

    # 실시간 주문 체결 처리 로직
    def updateJangoCont(self, pbCont):
        # 잔고 리스트 map 의 key 값
        # key = (pbCont['종목코드'], dicBorrow[pbCont['현금신용']], pbCont['대출일'])
        # key = pbCont['종목코드']
        code = pbCont['종목코드']

        # 주문 체결에서 들어온 신용 구분 값 --> 잔고 구분값으로 치환
        dicBorrow = {
            '현금': ord(' '),
            '유통융자': ord('Y'),
            '자기융자': ord('Y'),
            '주식담보대출': ord('B'),
            '채권담보대출': ord('B'),
            '매입담보대출': ord('M'),
            '플러스론': ord('P'),
            '자기대용융자': ord('I'),
            '유통대용융자': ord('I'),
            '기타': ord('Z')
        }

        # 접수, 거부, 확인 등은 매도 가능 수량만 업데이트 한다.
        if pbCont['체결플래그'] == '접수' or pbCont['체결플래그'] == '거부' or pbCont['체결플래그'] == '확인':
            if (code not in self.jangoData):
                return
            self.jangoData[code]['매도가능'] = pbCont['매도가능수량']
            return

        if (pbCont['체결플래그'] == '체결'):

            if code in self.target_data.keys():
                self.target_data[code]['체결상태'] = 1

            if (code not in self.jangoData):  # 신규 잔고 추가
                if (pbCont['체결기준잔고수량'] == 0):
                    return
                print('신규 잔고 추가', code)
                # 신규 잔고 추가
                item = {}
                item['종목코드'] = pbCont['종목코드']
                item['종목명'] = pbCont['종목명']
                item['현금신용'] = dicBorrow[pbCont['현금신용']]
                item['대출일'] = pbCont['대출일']
                item['잔고수량'] = pbCont['체결기준잔고수량']
                item['매도가능'] = pbCont['매도가능수량']
                item['장부가'] = pbCont['장부가']
                # 매입금액 = 장부가 * 잔고수량
                item['매입금액'] = item['장부가'] * item['잔고수량']

                print('신규 현재가 요청', code)
                #self.objRPCur.Request(code, self)
                self.objCur[code] = CpPBStockCur()
                self.objCur[code].Subscribe(code, self)

                item['현재가'] = self.curDatas[code]['cur']
                item['대비'] = self.curDatas[code]['diff']
                item['거래량'] = self.curDatas[code]['vol']

                self.jangoData[code] = item

            else:
                # 기존 잔고 업데이트
                item = self.jangoData[code]
                item['종목코드'] = pbCont['종목코드']
                item['종목명'] = pbCont['종목명']
                item['현금신용'] = dicBorrow[pbCont['현금신용']]
                item['대출일'] = pbCont['대출일']
                item['잔고수량'] = pbCont['체결기준잔고수량']
                item['매도가능'] = pbCont['매도가능수량']
                item['장부가'] = pbCont['장부가']
                # 매입금액 = 장부가 * 잔고수량
                item['매입금액'] = item['장부가'] * item['잔고수량']

                # 잔고 수량이 0 이면 잔고 제거
                if item['잔고수량'] == 0:
                    del self.jangoData[code]
                    self.objCur[code].Unsubscribe()
                    del self.objCur[code]

        self.printJango()
        return

    # 실시간 현재가 처리 로직
    def updateCurPBData(self, curData):
        code = curData['code']
        self.curDatas[code] = curData
        # print(self.curDatas[code]) -> {'code': 'A006400', 'diff': 1500, 'cur': 235500, 'vol': 115780}

        if code not in self.target_data.keys():
            conc_state = 1
            order_state = 1
        else:
            target_data = self.target_data[code]
            order_state = target_data['주문상태']
            conc_state = target_data['체결상태']
            objPrice = target_data['OBJ']
            adjPrice = target_data['OBJ2']
            tgScale = target_data['scale']

        curPrice = curData['cur']

        current_time = QTime.currentTime()
        text_time = current_time.toString("hh:mm:ss")
        limit_time = QTime(15, 0, 0)

        if order_state == 0:
            if curPrice >= objPrice and current_time < limit_time:
                target_data['주문상태'] = 1
                self.textBrowser.append("목표 매수가 도달 %s @%s" % (code, text_time))
                #telegram("목표 매수가 도달 %s %s @%s" % (code, target_data['name'], text_time))
                telegram_ch("<매수알림>\n%s %s\n매수 기준가 %i원 이하" % (code, target_data['name'], objPrice))
                telegram_share(f"{code},{objPrice},{adjPrice},{text_time}")
                self.objCur[code].Unsubscribe()

                amount = divmod(unit_price * tgScale, adjPrice)[0]
                self.q.put((code, adjPrice, amount))

        elif conc_state == 1 and code in self.jangoData:
            self.upjangoCurData(code)

    def upjangoCurData(self, code):
        # 잔고에 동일 종목을 찾아 업데이트 하자 - 현재가/대비/거래량/평가금액/평가손익
        curData = self.curDatas[code]
        item = self.jangoData[code]
        item['현재가'] = curData['cur']
        item['대비'] = curData['diff']
        item['거래량'] = curData['vol']
        profit = (item['현재가'] * 0.9975 / item['장부가'] - 1) * 100

        # 물타기 추가
        if profit < -5:
            if code not in self.multagi:
                self.multagi.append(code)
                telegram("물타기 알림: %s" % (item['종목명']))

        # 잔고 테이블 현재가, 수익률 업데이트
        searchResult = self.tableWidget_jango.findItems(code, Qt.MatchExactly)
        if searchResult:
            row = searchResult[0].row()
            tableItem = QTableWidgetItem("{0:.0f}".format(item['현재가']))
            tableItem.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.tableWidget_jango.setItem(row, 6, tableItem)
            tableItem = QTableWidgetItem("{0:.2f}".format(profit))
            tableItem.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.tableWidget_jango.setItem(row, 7, tableItem)

        # if profit >= 0.35 and item['매도가능'] > 0:
        #     current_time = QTime.currentTime()
        #     text_time = current_time.toString("hh:mm:ss")
        #
        #     self.textBrowser.append("목표 수익률 달성 %s @%s" % (code, text_time))
        #     # 현재가로 매도주문
        #     self.objRPOrder.sellOrder(code, item['현재가'], self.jangoData[code]['매도가능'])

    def sendBuyOrder(self, code):
        # 1. 현재가 통신
        self.objRPCur.Request(code, self)
        curPrice = self.curDatas[code]['cur']

        # 2. 매수 1호가로 매수 주문
        if self.curDatas[code]['bid1'] > 0:
            amount = divmod(1000000, curPrice)[0]

            price = self.curDatas[code]['bid1']
            self.objRPOrder.buyOrder(code, price, amount)


################################################
# MULTIPROCESS 함수

def run_gui(q):
    # GUI 구동
    app = QApplication(sys.argv)
    myWindow = MyWindow(q)
    myWindow.show()
    app.exec_()


################################################
# 매수내역 출력

class BuyList:
    """
    종가매매를 위해 매수내역 출력
    """
    def __init__(self):
        self.name = buy_path + "\\buy_" + today.strftime("%y%m%d") + ".csv"
        self.file_init()

    def file_init(self):
        if not os.path.exists(self.name):  # 파일이 존재하지 않는 경우에만 초기화 진행
            with open(self.name, 'w') as f:
                f.write("code\n")

    def write(self, msg):
        with open(self.name, 'a') as f:
            f.write(msg + "\n")


def order(q):
    # 주문
    if InitPlusCheck() == False:
        return
    rpOrder = CpRPOrder()
    buy_list = BuyList()

    while True:
        arg = q.get()
        if arg is None:
            print("shutting down")
            return
        print(arg)
        code, price, amount = arg
        while rpOrder.buyOrder(code, price, amount) is False:
            time.sleep(2)
        buy_list.write(code)


if __name__ == "__main__":
    main_q = Queue()
    process_gui = Process(target=run_gui, args=(main_q,))  # GUI 구동 프로세스
    process_order = Process(target=order, args=(main_q,))  # 주문 프로세스
    process_gui.start()
    process_order.start()

    main_q.close()
    main_q.join_thread()

    process_gui.join()
    process_order.join()
    print('end')
