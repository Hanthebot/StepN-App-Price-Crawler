import requests
import time
import winsound
from getpass import getuser
import os
import sys, json
import traceback
from PIL import Image
import pyautogui,sys,serial
import pytesseract
from serial.tools import list_ports
import numpy as np
import cv2

pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'

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

def filter0(imgArray):
    return imgArray

def filter1(imgArray):
    red, green, blue = imgArray[:,:,0], imgArray[:,:,1], imgArray[:,:,2]
    mask = (green>=90) & (blue<green)
    mask2 = (green<90) & (blue>=green)
    imgArray[:,:,:3][mask] = [0,0,0]
    imgArray[:,:,:3][mask2] = [255,255,255]
    return imgArray

def filter2(imgArray):
    red, green, blue = imgArray[:,:,0], imgArray[:,:,1], imgArray[:,:,2]
    mask = red>=130
    mask2 = red<130
    imgArray[:,:,:3][mask] = [0,0,0]
    imgArray[:,:,:3][mask2] = [255,255,255]
    return imgArray

def filter3(imgArray):
    red, green, blue = imgArray[:,:,0], imgArray[:,:,1], imgArray[:,:,2]
    mask1 = (red < 80) & (green < 80) & (blue < 80)
    mask2 = (red > 230) & (green > 230) & (blue > 230)
    imgArray[:,:,:3][mask1] = [0, 0, 0]
    imgArray[:,:,:3][mask2] = [255,255,255]
    return imgArray
    
maxTry = 1

def scanSection(cord, filt = filter3):
    img = pyautogui.screenshot(region=cord)
    imageFiltered = Image.fromarray(cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY))
    #imageFiltered.show()
    stringified = pytesseract.image_to_string(imageFiltered, lang=None)
    return stringified

def read_screen(coordinate):
    lis = []
    score = 0
    
    while score < maxTry and len(lis) == 0:
        time.sleep(0.24)
        scanned = scanSection(coordinate, filt = filter3)
        print(scanned)
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
    
def read_screen2(coordinate):
    score = 0
    x = True
    lis = []
    scannedBid = scanSection((0, 0, int(coordinate[0]/4), coordinate[1]-1))
    lis.append(scannedBid)
    scannedAsk = scanSection((int(img.size[0]*2/4), 0, int(coordinate[0]*3/4), coordinate[1]-1))
    lis.append(scannedAsk)
    
    return lis

def checkCoin():
    return replaceAll(scanSection(symbol),['\x0c',',','\n'])
    
coordPrice=(1350,563,140,20)
symbol=(1341,85,51,20)

def checkPrice(coordinate = coordPrice):
    g = read_screen(coordinate)
    
    return g

print(checkPrice())