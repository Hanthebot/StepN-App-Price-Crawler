import json
import os
import socket
import sys
import time
import traceback
import winsound

import colorama
import cv2
import numpy as np
from openpyxl import Workbook
from PIL import Image
import pyautogui
import pytesseract
import requests
import telepot
from termcolor import colored

from util import fl, fo2, fo3, fo4, mkIfNone, number, onlyNumber, randString

pytesseract.pytesseract.tesseract_cmd = r"C:/Program Files/Tesseract-OCR/tesseract.exe"

setting = {
    "coordPrice_C": [0.0631, 0.5390, 0.2523, 0.0203],
    "retryDelay": 0.25,
    "refreshPoint_C": [0.4865, 0.2604],
    "dragLength_C": 0.1478,
    "dragDuration": 0.2,
    "dragDelay": 2,
    "detectionMaxTry": 10,
    "saveImg_C": [0.0541, 0.2401, 0.4505, 0.3445],
    "dimension": [1315, 33, 555, 987]
}

print("setting...", end = "")
if os.path.exists("./data/data.json"):
    with open("./data/data.json", "r", encoding = "utf-8") as fil:
        totalSet = json.load(fil)
        if socket.gethostname() in totalSet.keys():
            setting = totalSet[socket.gethostname()]
            print("found")
        else:
            setting = totalSet["default"]
            print("default")
else:
    print("default")

if "place" in setting:
    try:
        cmdXi, cmdYi, cmdX, cmdY = setting["place"]
        os.system(f"cmdow @ /MOV {cmdXi} {cmdYi} /SIZ {cmdX} {cmdY}")
    except OSError:
        pass

xi, yi, x, y = setting["dimension"]
setting["dimension"] = tuple(setting["dimension"])
setting["coordPrice"] = (int(xi + setting["coordPrice_C"][0] * x), int(yi + setting["coordPrice_C"][1] * y), int(setting["coordPrice_C"][2] * x), int(setting["coordPrice_C"][3] * y))
setting["refreshPoint"] = (int(xi + setting["refreshPoint_C"][0] * x), int(yi + setting["refreshPoint_C"][1] * y))
setting["dragLength"] = int(setting["dragLength_C"] * y)
setting["saveImg"] = (int(xi + setting["saveImg_C"][0] * x), int(yi + setting["saveImg_C"][1] * y), int(setting["saveImg_C"][2] * x), int(setting["saveImg_C"][3] * y))

info = {
    "developer": "",
    "developer_id": "",
    "admin": [],
    "bot_token": "",
    "bot_id": "",
    "alert_bot_token": "",
    "subscribers": [],
    "whitelist": [],
    "subscriber_name_dict": {},
    "server_list": {}
}

print("ID information...", end = "")
if os.path.exists("./data/StepNinfo.json"):
    with open("./data/StepNinfo.json", "r", encoding = "utf-8") as fil:
        try:
            info = json.load(fil)
            print("found")
        except json.decoder.JSONDecodeError:
            print("default")
else:
    print("default")

bot = telepot.Bot(info["bot_token"])

alert_bot = telepot.Bot(info["alert_bot_token"])

user_data = {
    "default": {
        "percentage": 0.0,
        "ratioMax": 28,
        "ratioMin": 26
    }
}

print("user data...", end = "")
if os.path.exists("./data/StepNuserData.json"):
    with open("./data/StepNuserData.json", "r", encoding = "utf-8") as fil:
        try:
            user_data = json.load(fil)
            print("found")
        except json.decoder.JSONDecodeError:
            bot.sendMessage(info["developer"], "ERROR")
            print("default")
else:
    print("default")

for you in info["subscribers"]:
    if you not in list(user_data.keys()):
        user_data[you] = user_data["default"]
        try:
            bot.sendMessage(you,"Your data was reset")
        except telepot.exception.TelepotException:
            pass

def saveFile():
    with open("./data/StepNuserData.json", "w", encoding = "utf-8") as fil:
        json.dump(user_data, fil, indent = 4)

def saveInfo():
    with open("./data/StepNinfo.json", "w", encoding = "utf-8") as fil:
        json.dump(info, fil, indent = 4, ensure_ascii = False)

manifest = {
    "name": "StepN",
    "version": "Jun27",
    "curB": "",
    "commands": {
        "퍼센트": "최저가 알림 체크/업데이트 \n예: 퍼센트/퍼센트 1.0",
        "max": "SOL/GST 비율 알림 체크/업데이트 \n예: max/max 27.0",
        "min": "SOL/GST 비율 알림 체크/업데이트 \n예: min/min 27.0",
        "데이터": "프로그램용 내부 크롤링 데이터 출력 \n예: 데이터",
        "유저데이터": "유저 개인 설정 \n예: 데이터",
        "기본데이터": "유저 개인 설정 \n예: 데이터",
        "기본프린트": "프로그램용 출력 \n예: 기본프린트",
        "프린트": "프로그램용 출력 \n예: 프린트",
        "초대/링크/초대링크": "프로그램 공유용 링크",
        "server": "returns server name",
        "help": "message list",
        "usage": "returns usage \n 예:usage/usage [command]"
    },
    "commands_admin": {
        "refresh_code": "",
        "code": "",
        "execute": "",
        "kill_execute": "",
        "us": "",
        "whole_data": "",
        "naked_data": "",
        "revive": "",
        "퍼센트d": "최저가 알림 체크/업데이트 \n예: 퍼센트d/퍼센트d 1.0",
        "maxd": "SOL/GST 비율 알림 체크/업데이트 \n예: maxd/maxd 27.0",
        "mind": "SOL/GST 비율 알림 체크/업데이트 \n예: mind/mind 27.0",
        "die": "",
        "helpd": "",
        "usaged": "",
        "screenshot": "스크린샷 전송. \n예: screenshot OR screenshot 130 400 100 100"
    }
}

data = {
    "market": {},
    "prev_market": {},
    "users": {you: {} for you in info["subscribers"] + ["default"]},
    "start": time.strftime("%Y%m%d%H%M%S"),
    "die": False,
    "revive": False,
    "error": "None",
    "invitation_code": "",
    "execute_on": time.time(),
    "coinlist": ["SOL", "GST", "GMT"],
    "recent_ban_date": {you: "0" for you in info["subscribers"]}
}

data["invitation_code"] = randString(7)

mkIfNone("./dataxl")
mkIfNone("./img")
mkIfNone("./data")
mkIfNone("./log")
wb = Workbook()
ws1 = wb.active
colorama.init()

def upbit(BTC = "BTC", cur = "KRW"):
    htm = requests.get(f"https://api.upbit.com/v1/orderbook?markets={cur}-{BTC}", headers = {"User-Agent": "Mozilla/5.0"})
    hack = htm.json()
    
    return [hack[0]["orderbook_units"][0]["bid_price"], hack[0]["orderbook_units"][0]["ask_price"], hack[0]["orderbook_units"][0]["bid_size"], hack[0]["orderbook_units"][0]["ask_size"]]

def ftx(BTC = "BTC", cur = "USD"):
    htm = requests.get(f"https://ftx.com/api/markets/{BTC}/{cur}/orderbook", headers = {"User-Agent": "Mozilla/5.0"})
    hack = htm.json()
    
    return [hack["result"]["bids"][0][0], hack["result"]["asks"][0][0], hack["result"]["bids"][0][1], hack["result"]["asks"][0][1]]

def scanSection(cord, showImage = False):
    img = pyautogui.screenshot(region=cord)
    imageFiltered = Image.fromarray(cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY))
    if showImage:
        imageFiltered.show()
    stringified = pytesseract.image_to_string(imageFiltered)
    return (stringified, imageFiltered)

def read_screen(coordinate, maxTry = setting["detectionMaxTry"]):
    lis = []
    score = 0
    
    while score < maxTry and len(lis) == 0:
        time.sleep(setting["retryDelay"])
        scanned, _ = scanSection(coordinate)
        scanned = scanned.split(" ")[0]
        filtered = onlyNumber(scanned)
        filteredP = filtered.replace(".", "")
        if number(filtered) or number(filteredP):
            numerified = float(filtered) if number(filtered) else float(filteredP)
            lis.append(numerified)
            break
        score += 1
        
    if len(lis) == 0:
        return None
    return lis

def refresh(coor1 = setting["refreshPoint"], drag = setting["dragLength"], duration = setting["dragDelay"]):
    pos = pyautogui.position()
    pyautogui.moveTo(coor1)
    pyautogui.dragTo(coor1[0], coor1[1] + drag, setting["dragDuration"], button = "left")
    pyautogui.moveTo(pos)
    time.sleep(duration)

def stepNPrice(coordinate = setting["coordPrice"]):
    refresh()
    returned = read_screen(coordinate)
    return float(returned[0]) if returned != None else 10^8

def getcoin():
    base = {BTC: [10**8, 10**8, 10**8, 10**8] for BTC in data["coinlist"]}
    for BTC in data["coinlist"]:
        err = 0
        prblm = True
        while err < 10 and prblm:
            try:
                base[BTC] = ftx(BTC, manifest["curB"]) if manifest["curB"] != "" else ftx(BTC)
                prblm = False
            except KeyError:
                err += 1
                time.sleep(0.5)
    shoesPrice = stepNPrice()
    
    data["market"] = {
        "base": {BTC: {} for BTC in data["coinlist"]},
        "shoesPrice": 0.001
    }
    try:
        for BTC in data["coinlist"]:
            data["market"]["base"][BTC]["bid"] = float(base[BTC][0])
            data["market"]["base"][BTC]["ask"] = float(base[BTC][1])
        data["market"]["shoesPrice"] = shoesPrice
        data["market"]["shoesPriceUSD"] = float(base["SOL"][0]) * shoesPrice
        data["market"]["cost"] = (360 * float(base["GST"][1]) + 40 * float(base["GMT"][1])) * 1.06
        data["market"]["premium"] = (data["market"]["shoesPriceUSD"] / data["market"]["cost"] - 1) * 100
        data["market"]["SOLGSTRatio"] = float(base["SOL"][0]) / float(base["GST"][1])
        ws1.cell(ws1.max_row, 1, shoesPrice)
        ws1.cell(ws1.max_row, 2, data["market"]["shoesPriceUSD"])
        for BTC in data["coinlist"]:
            n = data["coinlist"].index(BTC)
            ws1.cell(ws1.max_row, 2 * n + 3, float(base[BTC][0]))
            ws1.cell(ws1.max_row, 2 * n + 4, float(base[BTC][1]))
        ws1.cell(ws1.max_row, 9, data["market"]["cost"])
        ws1.cell(ws1.max_row, 10, data["market"]["premium"])
        ws1.cell(ws1.max_row, 11, data["market"]["SOLGSTRatio"])
    except KeyError:
        data["market"] = data["prev_data"]

def printer(you = "default"):
    pre_col = colored(fo2(data["market"]["premium"]), "green" if data["market"]["premium"] > float(user_data[you]["percentage"]) else "red")
    rat_c = "red"
    if data["market"]["SOLGSTRatio"] > user_data[you]["ratioMax"]:
        rat_c = "green"
    elif data["market"]["SOLGSTRatio"] < user_data[you]["ratioMin"]:
        rat_c = "blue"
        
    rat_col = colored(fo2(data["market"]["SOLGSTRatio"]), rat_c)
    
    print(f"StepN 최저가:  {data['market']['shoesPrice']} SOL")
    print("\n".join([f"{BTC} FT가:  {fo3(data['market']['base'][BTC]['bid'])} {fo4(data['market']['base'][BTC]['ask'])}" for BTC in data["coinlist"]]))
    print(f"달러 가격:  {fo2(data['market']['shoesPriceUSD'])} USD")
    print(f"StepN 원가:  {fo2(data['market']['cost'])} USD")
    print(f"프리미엄: {pre_col}%")
    print(f"SOL/GST: {rat_col}")

def final():
    print(f"현재 시각: {time.strftime('%Y-%m-%d-%H:%M:%S')}")
    
def printerG():
    instantS = ""
    instantS += f"StepN 최저가:  {data['market']['shoesPrice']} SOL" + "\n"
    instantS += "\n".join([f"{BTC} FT가:  {fo3(data['market']['base'][BTC]['bid'])} {fo4(data['market']['base'][BTC]['ask'])}" for BTC in data["coinlist"]]) + "\n"
    instantS += f"달러 가격:  {fo2(data['market']['shoesPriceUSD'])} USD" + "\n"
    instantS += f"StepN 원가:  {fo2(data['market']['cost'])} USD" + "\n"
    instantS += f"프리미엄: {fo2(data['market']['premium'])}%" + "\n"
    instantS += f"SOL/GST: {fo2(data['market']['SOLGSTRatio'])}" + "\n"
    instantS += "\n"
    return instantS


def finalG():
    instantS = f"현재 시각: {time.strftime('%Y-%m-%d-%H:%M:%S')}\n"
    return instantS
    
def dataFormat(jso):
    try:
        inst = f"가격:{jso['price']}\n"
        return inst
    except KeyError:
        return ""

def handle(msg):
    content_type, _, chat_id = telepot.glance(msg)
    chat_id = str(chat_id)
    if content_type == "text":
        command = msg["text"].split(" ")[0]
        if chat_id in info["subscribers"]:
            if len(msg["text"].split(" ")) > 1:
                content = msg["text"].split(" ")[1].replace("\n","")
                if number(content):
                    if command in ["퍼센트", "max", "min"]:
                        user_data[chat_id][["percentage", "ratioMax", "ratioMin"][["퍼센트", "max", "min"].index(command)]] = float(content)
                        saveFile()
                        bot.sendMessage(chat_id, f"{command}->{content}")
                    
                    elif command.lower() in list(manifest["commands_admin"].keys()):
                        if chat_id in info["admin"]:
                            if command in ["퍼센트d", "maxd", "mind"]:
                                user_data["default"][["percentage", "ratioMax", "ratioMin"][["퍼센트d", "maxd", "mind"].index(command)]] = float(content)
                                saveFile()
                                bot.sendMessage(chat_id, f"{command}->{content}")
                                
                            elif command in ["white_add", "whitelist"]:
                                if content not in info["whitelist"]:
                                    info["whitelist"].append(content)
                                    saveInfo()
                                    bot.sendMessage(chat_id, f"{content} added")
                
                            elif command.lower() == "screenshot":
                                contents = msg["text"].split(" ")
                                dimension = setting["dimension"]
                                try:
                                    if len(contents) >= 5:
                                        dimension = tuple(int(c) for c in contents[1:5])
                                except IndexError:
                                    pass
                                bot.sendMessage(chat_id, "Processing...")
                                img = pyautogui.screenshot(region = dimension)
                                img_name = f"SC_{time.strftime('%Y%m%d%H%M%S')}.png"
                                img.save("./img/" + img_name)
                                with open("./img/" + img_name, "rb") as fil:
                                    bot.sendPhoto(chat_id, fil)
                                os.remove("./img/" + img_name)
                
                            else:
                                bot.sendMessage(chat_id, "None")
                    else:
                        bot.sendMessage(chat_id, "명령어 리스트: " + ", ".join(list(manifest["commands"].keys())) + "\n값 입력은 뒤에 숫자를 넣는다\nex) 환율 40")
            
                elif command.lower() == "usage":
                    if content in manifest["commands"]:
                        bot.sendMessage(chat_id, f"{content}: {manifest['commands'][content]}")
                    else:
                        bot.sendMessage(chat_id, "\n".join([f"{com}:{manifest['commands'][com]}" for com in manifest["commands"]]))
                
                else:
                    bot.sendMessage(chat_id, "Message not acceptable")
            else:
                if command in ["퍼센트", "max", "min"]:
                    content = user_data[chat_id][["percentage", "ratioMax", "ratioMin"][["퍼센트", "max", "min"].index(command)]]
                    bot.sendMessage(chat_id, f"{command}={content}")
                elif command == "데이터":
                    bot.sendMessage(chat_id, str(data))
                elif command == "유저데이터":
                    bot.sendMessage(chat_id, dataFormat(user_data[chat_id]))
                elif command.lower() == "기본데이터":
                    bot.sendMessage(chat_id,dataFormat(user_data["default"]))
                elif command == "프린트":
                    instant = ""
                    instant += printerG()
                    instant += finalG()
                    bot.sendMessage(chat_id, instant)
                elif command == "기본프린트":
                    instant = ""
                    instant += printerG()
                    instant += finalG()
                    bot.sendMessage(chat_id, instant)
                elif command in ["링크", "초대", "초대링크"]:
                    bot.sendMessage(chat_id, "t.me/" + info["bot_id"])
                elif command.lower() == "help":
                    bot.sendMessage(chat_id, "command list: " + ", ".join(list(manifest["commands"].keys())))
                elif command.lower() == "usage":
                    bot.sendMessage(chat_id, "\n".join([f"{com}:{manifest['commands'][com]}" for com in manifest["commands"]]))
                elif command == "server":
                    bot.sendMessage(chat_id, fl(info["server_list"], socket.gethostname()))
                elif command.lower() in list(manifest["commands_admin"].keys()):
                    if chat_id in info["admin"]:
                        if command.lower() == "refresh_code":
                            data["invitation_code"] = randString(7)
                        elif command.lower() == "whole_data":
                            whole_lis=[i + "\n"+dataFormat(user_data[i]) for i in user_data.keys()]
                            try:
                                bot.sendMessage(chat_id, "\n".join(whole_lis))
                            except telepot.exception.TelepotException:
                                for x in range(int(len("\n".join(whole_lis)) / 4095) + 1):
                                    bot.sendMessage(chat_id, "\n".join(whole_lis)[x * 4095:(x + 1) * 4095])
                        elif command.lower() == "naked_data":
                            try:
                                bot.sendMessage(chat_id, json.dumps(user_data, indent = 4))
                            except telepot.exception.TelepotException:
                                for x in range(int(len(json.dumps(user_data, indent = 4))/4095)+1):
                                    bot.sendMessage(chat_id, json.dumps(user_data, indent = 4)[x * 4095:(x + 1) * 4095])
                        elif command.lower() == "code":
                            bot.sendMessage(chat_id, data["invitation_code"])
                        elif command.lower() == "us":
                            bot.sendMessage(chat_id, "\n".join([fl(info["subscriber_name_dict"], you) for you in info["subscribers"]]))
                        elif command.lower() == "white":
                            bot.sendMessage(chat_id, "\n".join(info["whitelist"]))
                        elif command.lower() == "revive":
                            bot.sendMessage(chat_id, "See you soon")
                            data["revive"] = True
                        elif command.lower() == "die":
                            bot.sendMessage(chat_id, "See you later")
                            data["die"] = True
                        elif command.lower() in ["퍼센트d", "maxd", "mind"]:
                            content=user_data["default"][["percentage", "ratioMax", "ratioMin"][["퍼센트d", "maxd", "mind"].index(command)]]
                            bot.sendMessage(chat_id, f"{command}={content}")
                        elif command.lower() == "helpd":
                            bot.sendMessage(chat_id, "command list: " + ", ".join(list(manifest["commands_admin"].keys())))
                        elif command.lower() == "usaged":
                            bot.sendMessage(chat_id, "\n".join([f"{com}:{manifest['commands_admin'][com]}" for com in manifest["commands_admin"]]))
                        elif command.lower() == "screenshot":
                            bot.sendMessage(chat_id, "Processing...")
                            img = pyautogui.screenshot(region = setting["dimension"])
                            img_name = f"S_{time.strftime('%Y%m%d%H%M%S')}.png"
                            img.save("./img/" + img_name)
                            with open("./img/" + img_name, "rb") as fil:
                                bot.sendPhoto(chat_id, fil)
                            os.remove("./img/" + img_name)
                        else:
                            bot.sendMessage(chat_id, "None, command list: " + ", ".join(list(manifest["commands_admin"].keys())))
                    else:
                        bot.sendMessage(chat_id, "Forbidden access")
                else:
                    bot.sendMessage(chat_id, "명령어 리스트: " + ", ".join(list(manifest["commands"].keys())) + "\n값 입력은 뒤에 숫자를 넣는다\nex) 환율 40")
        elif chat_id in info["whitelist"]:
            info["subscribers"].append(chat_id)
            saveInfo()
            user_data[chat_id] = user_data["default"]
            data["recent_ban_date"][chat_id] = "0"
            bot.sendMessage(chat_id, "Welcome")
            alert_bot.sendMessage(info["developer"], f"{manifest['name']}: {chat_id} was added")   
        else:
            if command == data["invitation_code"]:
                info["subscribers"].append(chat_id)
                user_data[chat_id] = user_data["default"]
                data["recent_ban_date"][chat_id] = "0"
                bot.sendMessage(chat_id, "Welcome")
                alert_bot.sendMessage(info["developer"], f"{chat_id} was added")          
                data["invitation_code"] = randString(7)
            else:
                bot.sendMessage(chat_id, f"Unknown access. Contact developer t.me/{info['developer_id']} for the key")
                bot.sendMessage(info["developer"], f"{chat_id} tried to access: {msg['text']}")

bot.message_loop(handle)
try:
    while not data["die"]:
        if data["execute_on"] <= time.time():
            data["execute_on"] = time.time() + 60
            ws1[f"A{ws1.max_row+1}"] = time.strftime("%Y%m%d%H%M%S")
            getcoin()
            data["prev_data"] = data["market"]
            
            print("\n\n\n\n\n\n\n\n\n\n\n\n")
            printer()
            final()
            
            for you in info["subscribers"] + ["default"]:
                data["users"][you]["totalSs"] = printerG() + finalG()
            if data["market"]["premium"] > user_data["default"]["percentage"]:
                pass
            elif data["market"]["SOLGSTRatio"] > user_data["default"]["ratioMax"]:
                pass
            elif data["market"]["SOLGSTRatio"] < user_data["default"]["ratioMin"]:
                pass
            for you in info["subscribers"]:
                try:
                    if data["market"]["premium"] > user_data[you]["percentage"]:
                        bot.sendMessage(you, data["users"][you]["totalSs"])
                    elif data["market"]["SOLGSTRatio"] > user_data[you]["ratioMax"]:
                        bot.sendMessage(you, data["users"][you]["totalSs"])
                    elif data["market"]["SOLGSTRatio"] < user_data[you]["ratioMin"]:
                        bot.sendMessage(you, data["users"][you]["totalSs"])
                except Exception:
                    if data["recent_ban_date"][you] != time.strftime("%Y-%m-%d"):
                        alert_bot.sendMessage(info["developer"], f"{manifest['name']}: {fl(info['subscriber_name_dict'], you)} blocked this bot")
                        data["recent_ban_date"][you] = time.strftime("%Y-%m-%d")
        if data["revive"]:
            sys.exit()
        time.sleep(3)
        
except:
    data["error"] = traceback.format_exc()
    os.system("start cmd /C " + sys.argv[0])
    winsound.PlaySound("./data/에러.wav", winsound.SND_ALIAS)

print(f"error {time.strftime('%Y-%m-%d-%H:%M:%S')}")
wb.save(f"./dataxl/{manifest['name']}{data['start']}.xlsx")
with open("./log/" + sys.argv[0].split('\\')[-1].split('.')[0] + ".txt", "a", encoding = "utf-8") as fil:
    fil.write("시간: {time.strftime('%Y-%m-%d-%H:%M:%S')}\n\n{data['error']}\n---------------------------\n\n")