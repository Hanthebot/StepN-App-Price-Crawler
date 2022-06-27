import os
import random
import string

def fl(lis, key):
    if key in lis:
        return lis[key]
    return key

def fo(a):    
    return "{0:.2f}".format(a).ljust(15)

def fo2(a):
    return "{0:.2f}".format(a).ljust(len("{0:.2f}".format(a)))

def fo3(a):    
    return "{0:.6f}".format(a).ljust(15)

def fo4(a):    
    return "{0:.6f}".format(a).ljust(len("{0:.6f}".format(a)))

def mkIfNone(path):
    if not os.path.exists(path):
        if path.replace("\\", "/").split("/")[-1].replace(".", "")!=path.replace("\\", "/").split("/")[-1]:
            path="/".join(path.replace("\\", "/").split("/")[:-1])
        try:
            os.makedirs(path)
        except PermissionError:
            pass

def number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False

def onlyNumber(s, target = "0123456789."):
    return "".join([l for l in s if l in target])

def randString(length = 7):
    rand = ""
    for i in range(length):
        rand += random.choice(string.digits)
    return rand

def replaceAll(s, lis):
    for l in lis:
        s = s.replace(l,"")
    return s
