import requests
import time
import winsound
import os
import sys, json
from PIL import Image
import pyautogui
import pytesseract
import numpy as np
import cv2
import telepot
import colorama
from termcolor import colored, cprint
import traceback, socket
import random, string
from openpyxl import Workbook

pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'

setting = {
    "coordPrice" : [1365, 548, 140, 20],
    "retryDelay" : 0.25,
    "refreshPoint" : (1600, 275),
    "dragLength" : 150,
    "dragDuration" : 0.2,
    "dragDelay" : 1.5,
    "detectionMaxTry": 10
}

if os.path.exists("./data/data.txt"):
    #print("File already existing. Loading...")
    with open("./data/data.txt", 'r') as fil:
        setting = json.load(fil)

setting["coordPrice"] = tuple(setting["coordPrice"])
setting["refreshPoint"] = tuple(setting["refreshPoint"])

def replaceAll(s, lis):
    for l in lis:
        s = s.replace(l,'')
    return s

def isNumber(s):
    try:
        float(s)
        return True
    except:
        return False

def scanSection(cord, showImage = False):
    img = pyautogui.screenshot(region=cord)
    imageFiltered = Image.fromarray(cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY))
    if showImage:
        imageFiltered.show()
    stringified = pytesseract.image_to_string(imageFiltered, lang=None)
    return stringified

def read_screen(coordinate, maxTry = setting["detectionMaxTry"]):
    lis = []
    score = 0
    
    while score < maxTry and len(lis) == 0:
        time.sleep(setting['retryDelay'])
        scanned = scanSection(coordinate).split(" ")[0]
        #print(scanned)
        filtered = replaceAll(scanned,['\x0c',',','\n','S','O','L'])
        filteredP = filtered.replace(".","")
        if isNumber(filtered) or isNumber(filteredP):
            numerified = float(filtered) if isNumber(filtered) else float(filteredP)
            lis.append(numerified)
            break
        score += 1
        
    if len(lis) == 0:
        return None
    else:
        return lis

def checkCoin():
    return replaceAll(scanSection(symbol),['\x0c',',','\n'])

def refresh(coor1 = setting["refreshPoint"], drag = setting["dragLength"], duration = setting["dragDelay"]):
    pyautogui.moveTo(coor1)
    pyautogui.dragTo(coor1[0], coor1[1]+drag, setting["dragDuration"], button='left')
    time.sleep(duration)

def stepNPrice(coordinate = setting["coordPrice"]):
    refresh()
    returned = read_screen(coordinate)
    return float(returned[0]) if returned != None else 10^8

version="May28"

def mkIfNone(path):
    if not os.path.exists(path):
        if path.replace("\\","/").split("/")[-1].replace(".","")!=path.replace("\\","/").split("/")[-1]:
            path="/".join(path.replace("\\","/").split("/")[:-1])
        try:
            os.makedirs(path)
        except:
            pass

helps = {"가격":"최저가 알림 체크/업데이트 \n예: 가격/가격 400000",
         "데이터":"프로그램용 내부 크롤링 데이터 출력 \n예: 데이터",
         "유저데이터":"유저 개인 설정 \n예: 데이터",
         "기본데이터":"유저 개인 설정 \n예: 데이터",
         "기본프린트":"프로그램용 출력 \n예: 기본프린트",
         "프린트":"프로그램용 출력 \n예: 프린트",
         "초대/링크/초대링크":"프로그램 공유용 링크",
         "server":"returns server name",
         "help":"message list",
         "usage":"returns usage \n 예:usage/usage [command]"}
helps_D={"refresh_code":"",
    "code":"",
    "execute":"",
    "kill_execute":"",
    "us":"",
    "whole_data":"",
    "naked_data":"",
    "revive":"",
    "가격d":"환율 체크/업데이트 \n예: 환율/환율 36.3",
    "die":"",
    "helpd":"",
    "usaged":""}

name="StepN"
mkIfNone("./dataxl")

wb = Workbook()
ws1 = wb.active
start = time.strftime('%Y%m%d%H%M%S')
die = False
revive = False
link = 'hantheUltrabot'
Token = "1657206661:AAFero4uNhPawrLQcVtHjgyrOa-LrikxIQA"
trr=""
now="l"
default='default'

bot =telepot.Bot(Token)

ueharaT = "1500311505:AAEIc6gKWzhBswe-KwIfI6kXVaxtO2cdVfk"
uehara =telepot.Bot(ueharaT)

mkIfNone("./data")
mkIfNone("./log")
owner = ["709146795", "518298248"]
developer = "709146795"
if not os.path.exists("./data/StepNids.txt"):
    with open("./data/StepNids.txt",'w') as m:
        m.write("\n".join(["709146795", "518298248"]))
        m.close()

with open("./data/StepNids.txt",'r') as i:
    us = [ID.split("#")[0].replace(' ','') for ID in i.read().split("\n")]

if not os.path.exists("./data/StepNwhite.txt"):
    with open("./data/StepNwhite.txt",'w') as m:
        m.write("\n".join(["518298248","720960187","1673479631","1328593892","1744957536","1545297996"]))
        
with open("./data/StepNwhite.txt",'r') as i:
    whitelist = [idd.split("#")[0].replace(' ','') for idd in i.read().split("\n")]

defaultD = {"default": {
            'price': 500000
            }
         }
global user_data
try:
    with open("./data/StepNdata.txt",'r') as m:
        user_data=eval(m.read())
except:
    try:
        bot.sendMessage(developer,"ERROR")
    except:
        pass
    with open("./data/StepNdata.txt",'w') as m:
        m.write(str(defaultD).replace("]}, '","]}, \n'"))
    user_data = defaultD

for you in us:
    if you not in list(user_data.keys()):
        user_data[you] = user_data['default']
        try:
            bot.sendMessage(you,"Your data was reset")
        except:
            pass

curB = ""
string_pool = string.digits #string.ascii_letters+
global invitation
invitation = ""
for i in range(7):
    invitation += random.choice(string_pool)

with open("./data/StepNinvitation.txt",'w') as i:
    i.write(invitation)

colorama.init()

servers = {"DESKTOP-3MLFMEP":"컴퓨터", "DESKTOP-FEHJQDU":"노트북"}

def upbit(BTC="BTC", cur="KRW"):
    htm = requests.get("https://api.upbit.com/v1/orderbook?markets="+cur+"-"+BTC, headers = {"User-Agent": "Mozilla/5.0"})
    hack = htm.json()
    
    return [hack[0]["orderbook_units"][0]["bid_price"], hack[0]["orderbook_units"][0]["ask_price"], hack[0]["orderbook_units"][0]["bid_size"], hack[0]["orderbook_units"][0]["ask_size"]]

def fo(a):    
    return "{0:.2f}".format(a).ljust(15)

def fo2(a):
    return "{0:.2f}".format(a).ljust(len("{0:.2f}".format(a)))

def fo3(a):    
    return "{0:.6f}".format(a).ljust(15)

def fo4(a):    
    return "{0:.6f}".format(a).ljust(len("{0:.6f}".format(a)))

def oppobool(b):
    if b:
        return False
    else:
        return True

def getcoin(BTC):
    err = 0
    prblm = True
    base = [10^8, 10^8, 10^8, 10^8]
    while err<10 and prblm:
        try:
            base = upbit(BTC, curB) if curB!="" else upbit(BTC)
            prblm = False
        except:
            err+=1
            time.sleep(0.5)
    shoesPrice = stepNPrice()
    
    global data
    global user_data
    global ws1
    data['market'] = {
        "base": {},
        "shoesPrice": 0.001
    }
    try:
        data['market']["base"]["bid"] = float(base[0])
        data['market']["base"]["ask"] = float(base[1])
        data['market']["shoesPrice"] = shoesPrice
        data['market']["shoesPriceKRW"] = float(base[1])*shoesPrice
        ws1.cell(ws1.max_row, 1, shoesPrice)
        ws1.cell(ws1.max_row, 2, float(base[0]))
        ws1.cell(ws1.max_row, 3, float(base[1]))
        ws1.cell(ws1.max_row, 4, float(base[1])*shoesPrice)
    except:
        data['market'] = last['market']

def printer(BTC, you='default'):
    priceColor = colored(fo2(data['market']["shoesPriceKRW"]), 'green' if data['market']["shoesPriceKRW"] < float(user_data[you]['price']) else 'red')
    
    print(f"StepN 최저가:  {data['market']['shoesPrice']} SOL")
    print(f"{BTC} UP가:  {fo(data['market']['base']['bid'])} {fo2(data['market']['base']['ask'])}")
    print(f"원화 가격:  {priceColor}원")
    print("")

def final(you='default'):
    print("현재 시각: "+time.strftime('%Y-%m-%d-%H:%M:%S'))
    
def printerG(BTC, you='default'):
    instantS = ""
    instantS += f"StepN 최저가:  {data['market']['shoesPrice']} SOL"+"\n"
    instantS += f"{BTC} UP가:  {fo(data['market']['base']['bid'])} {fo2(data['market']['base']['ask'])}"+"\n"
    instantS += f"원화 가격:  {fo2(data['market']['shoesPriceKRW'])}원"+"\n"
    instantS += "\n"
    return instantS


def finalG(you='default'):
    instantS = "현재 시각: "+time.strftime('%Y-%m-%d-%H:%M:%S')+"\n"
    return instantS

def number(s):
    try:
        float(s)
        return True
    except:
        return False

def replaceAll(s, lis):
    for l in lis:
        s=s.replace(l,'')
    return s

def inOrNot(dic, key):
    if key in dic:
        return dic[key]
    else:
        return key

def saveFile():
    with open("./data/StepNdata.txt", "w") as m:
        m.write(str(user_data).replace("]}, '","]}, \n'"))

def saveFileIds():
    with open("./data/StepNids.txt", "w") as m:
        m.write("\n".join(us))
    
def dataFormat(jso):
    try:
        inst = f"가격:{jso['price']}\n"
        return inst
    except:
        return ""

colorama.init()

total_coinlist = []
data = {}
last = {}
data['users'] = {}
for you in us+['default']:
    data['users'][you] = {}

recentBanDate = {you:'0' for you in us}

def handle(msg):
    global die
    global revive
    global user_data
    global invitation
    content_type, chat_type, chat_id = telepot.glance(msg)
    chat_id=str(chat_id)
    if content_type == 'text':
        command = msg['text'].split(" ")[0]
        if chat_id in us:
            if len(msg['text'].split(" "))>1:
                content = msg['text'].split(" ")[1].replace("\n","")
                if number(content):
                    if command in ["가격"]:
                        user_data[chat_id]['price'] = float(content)
                        saveFile()
                        bot.sendMessage(chat_id, f"{command}->{content}")
                    
                    elif command.lower() in list(helps_D.keys()):
                        if chat_id in owner:
                            if command in ["가격d"]:
                                user_data['default']['price'] = float(content)
                                saveFile()
                                bot.sendMessage(chat_id, f"{command}->{content}")
                                
                            elif command in ["white_add", "whitelist"]:
                                if content not in whitelist:
                                    whitelist.append(content)
                                    m=open("./data/UpBiwhite.txt",'w')
                                    m.write("\n".join(whitelist))
                                    m.close()
                                    bot.sendMessage(chat_id, f"{content} added")
                            else:
                                bot.sendMessage(chat_id, "None")
                    else:
                        bot.sendMessage(chat_id, "명령어 리스트: "+", ".join(list(helps.keys()))+"\n값 입력은 뒤에 숫자를 넣는다\nex) 환율 40")
            
                elif command.lower()=="usage":
                    if content in helps.keys():
                        bot.sendMessage(chat_id, f"{content}: {helps[content]}")
                    else:
                        bot.sendMessage(chat_id, "\n".join([f"{com}:{helps[com]}" for com in helps]))
                else:
                    bot.sendMessage(chat_id, f"Message not acceptable")
            #출력
            else:
                if command in ["가격"]:
                    content=user_data[chat_id]['price']
                    bot.sendMessage(chat_id, f"{command}={content}")
                elif command =="데이터":
                    bot.sendMessage(chat_id, str(data))
                elif command =="유저데이터":
                    bot.sendMessage(chat_id, dataFormat(user_data[chat_id]))
                elif command.lower()=="기본데이터":
                    bot.sendMessage(chat_id,dataFormat(user_data['default']))
                elif command=="프린트":
                    instant=""
                    instant+=printerG("SOL", chat_id)+"\n"
                    instant+=finalG(chat_id)+"\n"
                    bot.sendMessage(chat_id, instant)
                elif command=="기본프린트":
                    instant=""
                    instant+=printerG("SOL", 'default')+"\n"
                    instant+=finalG('default')+"\n"
                    bot.sendMessage(chat_id, instant)
                elif command in ['링크','초대','초대링크']:
                    bot.sendMessage(chat_id, 't.me/'+link)
                elif command.lower()=="help":
                        bot.sendMessage(chat_id, "command list: "+", ".join(list(helps.keys())))
                elif command.lower()=="usage":
                    bot.sendMessage(chat_id, "\n".join([f"{com}:{helps[com]}" for com in helps]))
                elif command=="server":
                    bot.sendMessage(chat_id, inOrNot(servers, socket.gethostname()))
                elif command.lower() in list(helps_D.keys()):
                    if chat_id in owner:
                        if command.lower()=="refresh_code":                        
                            invitation=""
                            for i in range(7):
                                invitation += random.choice(string_pool)
                            with open("./data/StepNinvitation.txt",'w') as i:
                                i.write(invitation)
                        elif command.lower()=="whole_data":
                            whole_lis=[i+"\n"+dataFormat(user_data[i]) for i in user_data.keys()]
                            try:
                                bot.sendMessage(chat_id,"\n".join(whole_lis))
                            except:
                                for x in range(int(len("\n".join(whole_lis))/4095)+1):
                                    bot.sendMessage(chat_id,"\n".join(whole_lis)[x*4095:(x+1)*4095])
                        elif command.lower()=="naked_data":
                            try:
                                bot.sendMessage(chat_id,str(user_data).replace("]}, '","]}, \n'"))
                            except:
                                for x in range(int(len(str(user_data).replace("]}, '","]}, \n'"))/4095)+1):
                                    bot.sendMessage(chat_id,str(user_data).replace("]}, '","]}, \n'")[x*4095:(x+1)*4095])
                        elif command.lower()=="code":
                            bot.sendMessage(chat_id,invitation)
                        elif command.lower()=="us":
                            bot.sendMessage(chat_id,"\n".join(us))
                        elif command.lower()=="white":
                            bot.sendMessage(chat_id,"\n".join(whitelist))
                        elif command.lower()=="revive":
                            bot.sendMessage(chat_id,"See you soon")
                            revive=True
                        elif command.lower()=="die":
                            bot.sendMessage(chat_id,"See you later")
                            die=True
                        elif command.lower() in ["가격"]:
                            content=user_data['default']['price']
                            bot.sendMessage(chat_id, f"{command}={content}")
                        elif command.lower()=="helpd":
                            bot.sendMessage(chat_id, "command list: "+", ".join(list(helps_D.keys())))
                        elif command.lower()=="usaged":
                            bot.sendMessage(chat_id, "\n".join([f"{com}:{helps_D[com]}" for com in helps_D]))
                        else:
                            bot.sendMessage(chat_id, "None, command list: "+", ".join(list(helps_D.keys())))
                    else:
                        bot.sendMessage(chat_id, "Forbidden access")
                else:
                    bot.sendMessage(chat_id, "명령어 리스트: "+", ".join(list(helps.keys()))+"\n값 입력은 뒤에 숫자를 넣는다\nex) 환율 40")
        elif chat_id in whitelist:
            us.append(chat_id)
            with open("./data/StepNids.txt",'w') as m:
                m.write("\n".join(us))
            user_data[chat_id]=user_data['default']
            recentBanDate[chat_id]='0'
            bot.sendMessage(chat_id, "Welcome")
            uehara.sendMessage(developer, f"{name}: {chat_id} was added")   
        else:
            if command==invitation:
                us.append(chat_id)
                with open("./data/StepNids.txt",'w') as m:
                    m.write("\n".join(us))
                user_data[chat_id]=user_data['default']
                recentBanDate[chat_id]='0'
                bot.sendMessage(chat_id, "Welcome")
                uehara.sendMessage(developer, f"{chat_id} was added")          
                invitation=""
                for i in range(7):
                    invitation += random.choice(string_pool)
                with open("./data/StepNinvitation.txt",'w') as i:
                    i.write(invitation)
            else:
                bot.sendMessage(chat_id, "Unknown access. Contact developer t.me/haninthai for the key")
                bot.sendMessage(developer, f"{chat_id} tried to access: {msg['text']}")

bot.message_loop(handle)
try:
    while not die:
        if now!=time.strftime("%M"):#for kor
            ws1[f'A{ws1.max_row+1}'] = time.strftime('%Y%m%d%H%M%S')
            getcoin("SOL")
            last = data
            
            print("\n\n\n\n\n\n\n\n\n\n\n\n")
            printer("SOL")
            final()
            
            for you in us+[default]:
                data['users'][you]['totalSs'] = printerG("SOL", you) + finalG(you)
            if data['market']['shoesPriceKRW'] < float(user_data['default']['price']):
                winsound.PlaySound("./data/김프의노래.wav", winsound.SND_ALIAS)
            for you in us:
                try:
                    if data['market']['shoesPriceKRW'] < float(user_data[you]['price']):
                        bot.sendMessage(you, data['users'][you]['totalSs'])
                except:
                    if recentBanDate[you] != time.strftime('%Y-%m-%d'):
                        uehara.sendMessage(developer, f"{name}: {fl(customers,you)} blocked this bot")
                        recentBanDate[you] = time.strftime('%Y-%m-%d')
            now=time.strftime("%M")    
        if revive:
            sys.exit()
        time.sleep(3)
    if die:
        print("error "+time.strftime('%Y-%m-%d-%H:%M:%S'))
        wb.save("./dataxl/"+name+start+".xlsx")
        
except:
    trr=traceback.format_exc()
    os.system("start cmd /C "+sys.argv[0])
    winsound.PlaySound("./data/에러.wav", winsound.SND_ALIAS)

if not die:
    print("error "+time.strftime('%Y-%m-%d-%H:%M:%S'))
    wb.save("./dataxl/"+name+start+".xlsx")
    with open("./log/"+sys.argv[0].split("\\")[-1].split(".")[0]+".txt", 'a',encoding='utf-8') as z:
        z.write('시간: '+time.strftime("%Y-%m-%d-%H:%M:%S\n")+"\n"+trr+"\n---------------------------\n\n")