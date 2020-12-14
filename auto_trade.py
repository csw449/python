import win32com.client
import slack
from flask import Flask, Response
from slackeventsapi import SlackEventAdapter
import os
from threading import Thread
from slack import WebClient
import json
import pythoncom

def find_code(name_input): 
 
    # 연결 여부 체크
    objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
    bConnect = objCpCybos.IsConnect
    if (bConnect == 0):
        print("PLUS가 정상적으로 연결되지 않음. ")
        exit()
    
    # 종목코드 리스트 구하기
    objCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
    codeList = objCpCodeMgr.GetStockListByMarket(1) #거래소
    codeList2 = objCpCodeMgr.GetStockListByMarket(2) #코스닥
    
    
    # print("거래소 종목코드", len(codeList))
    for i, code in enumerate(codeList):
        secondCode = objCpCodeMgr.GetStockSectionKind(code)
        name = objCpCodeMgr.CodeToName(code)
        if name == name_input:
            print(code)
            return code
            break
        stdPrice = objCpCodeMgr.GetStockStdPrice(code)
        # print(i, code, secondCode, stdPrice, name)
    
    # print("코스닥 종목코드", len(codeList2))
    for i, code in enumerate(codeList2):
        secondCode = objCpCodeMgr.GetStockSectionKind(code)
        name = objCpCodeMgr.CodeToName(code)
        if name == name_input:
            print(code)
            return code
            break
        stdPrice = objCpCodeMgr.GetStockStdPrice(code)
        # print(i, code, secondCode, stdPrice, name)
    
    # print("거래소 + 코스닥 종목코드 ",len(codeList) + len(codeList2))

def curr_price(code):

    # 연결 여부 체크
    objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
    bConnect = objCpCybos.IsConnect
    if (bConnect == 0):
        print("PLUS가 정상적으로 연결되지 않음. ")
        exit()
    
    # 현재가 객체 구하기
    objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
    try:
        objStockMst.SetInputValue(0, code)   #종목 코드 - 내츄럴 엔도택
    except:
        print("검색되는 종목이 없습니다. 종목명을 확인해주세요.")
        return
    objStockMst.BlockRequest()
    
    # 현재가 통신 및 통신 에러 처리 
    rqStatus = objStockMst.GetDibStatus()
    rqRet = objStockMst.GetDibMsg1()
    print("통신상태", rqStatus, rqRet)
    if rqStatus != 0:
        exit()
    
    # 현재가 정보 조회
    code = objStockMst.GetHeaderValue(0)  #종목코드
    name= objStockMst.GetHeaderValue(1)  # 종목명
    time= objStockMst.GetHeaderValue(4)  # 시간
    cprice= objStockMst.GetHeaderValue(11) # 종가
    diff= objStockMst.GetHeaderValue(12)  # 대비
    open= objStockMst.GetHeaderValue(13)  # 시가
    high= objStockMst.GetHeaderValue(14)  # 고가
    low= objStockMst.GetHeaderValue(15)   # 저가
    offer = objStockMst.GetHeaderValue(16)  #매도호가
    bid = objStockMst.GetHeaderValue(17)   #매수호가
    vol= objStockMst.GetHeaderValue(18)   #거래량
    vol_value= objStockMst.GetHeaderValue(19)  #거래대금
    
    # 예상 체결관련 정보
    exFlag = objStockMst.GetHeaderValue(58) #예상체결가 구분 플래그
    exPrice = objStockMst.GetHeaderValue(55) #예상체결가
    exDiff = objStockMst.GetHeaderValue(56) #예상체결가 전일대비
    exVol = objStockMst.GetHeaderValue(57) #예상체결수량
    
    
    # print("코드", code)
    print("이름", name)
    # print("시간", time)
    print("종가", cprice)
    print("대비", diff)
    print("시가", open)
    print("고가", high)
    print("저가", low)
    print("매도호가", offer)
    print("매수호가", bid)
    print("거래량", vol)
    print("거래대금", vol_value)
    
    
    if (exFlag == ord('0')):
        print("장 구분값: 동시호가와 장중 이외의 시간")
    elif (exFlag == ord('1')) :
        print("장 구분값: 동시호가 시간")
    elif (exFlag == ord('2')):
        print("장 구분값: 장중 또는 장종료")
    
    print("예상체결가 대비 수량")
    print("예상체결가", exPrice)
    print("예상체결가 대비", exDiff)
    print("예상체결수량", exVol)

# This `app` represents your existing Flask app
app = Flask(__name__)

greetings = ["hi", "hello", "hello there", "hey"]
stock = ["주식", "주식 현재가", "현재가", "주가"]

SLACK_SIGNING_SECRET = 'a3c3d6d0e99c348b9892b21bc6f34466'
slack_token = 'xoxb-1557254483319-1578222003460-zxKsKSfzrWqDbUU9VwOrZZOS'
VERIFICATION_TOKEN = 'PWuiMneRbFtBMXtiXqekPwm2'

#instantiating slack client
slack_client = WebClient(slack_token)

# An example of one of your Flask app's routes
@app.route("/")
def event_hook(request):
    json_dict = json.loads(request.body.decode("utf-8"))
    if json_dict["token"] != VERIFICATION_TOKEN:
        return {"status": 403}

    if "type" in json_dict:
        if json_dict["type"] == "url_verification":
            response_dict = {"challenge": json_dict["challenge"]}
            return response_dict
    return {"status": 500}
    return


slack_events_adapter = SlackEventAdapter(
    SLACK_SIGNING_SECRET, "/slack/events", app
)  


@slack_events_adapter.on("app_mention")

def handle_message(event_data):
    # print(event_data)
    def send_reply(value):
        event_data = value
        message = event_data["event"]
        print(message)
        if message.get("subtype") is None:
            command = message.get("text")
            channel_id = message["channel"]
            if any(item in command.lower() for item in greetings):
                message = (
                    "Hello <@%s>! :tada:"
                    % message["user"]  # noqa
                )
                slack_client.chat_postMessage(channel=channel_id, text=message)
            else:
                command_split = str(command).split('>')[1].replace(" ","")
                print(command_split)
                find_code(command_split)

                # result = curr_price(find_code(command_split))
                # message = (
                #     result
                # )
                slack_client.chat_postMessage(channel=channel_id, text=message)
    thread = Thread(target=send_reply, kwargs={"value": event_data})
    thread.start()
    return Response(status=200)


# Start the server on port 3000
# if __name__ == "__main__":
#   app.run(port=3000)



name_input = input("검색하고자 하는 종목명을 입력해 주세요\n")

curr_price(find_code(name_input))

