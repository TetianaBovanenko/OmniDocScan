from omnipage import *
from subprocess import Popen
import os, sys, stat, io
import shutil, platform
import traceback
import inspect

class tcolors:
    reset  = "\u001B[0m"
    red    = "\u001B[91m"
    yellow = "\u001B[33m"
    green  = "\u001B[92m"
    blue   = "\u001B[34m"
    purple = "\u001B[35m"

#region #define constants from sample.h:
YOUR_COMPANY = "YOUR COMPANY"
YOUR_PRODUCT = "YOUR PRODUCT"   # "CSDK-Samples"
SID = 0
PAGE_NUMBER_0 = 0
BUF_SIZE = 1024
INFOMSG = 0
ERRMSG  = 1
APIMSG  = 2
LISTMSG = 3
LOOPMSG = 4
#endregion

#region Other variables
EXIT_SUCCESS = 0
EXIT_FAILURE = 1
EXIT_KNOWN = -1

USE_OEM_LICENSE = False
LICENSE_FILE = ""   # insert name of your License file (lcx or lcxz)
OEM_CODE = ""       # put here the OEM_CODE

PASS = True

outputfolder = ""
#endregion

#region Logging
def InfoMsg(txt):
    print(tcolors.blue + "INF > " + txt + tcolors.reset)

def ErrMsg(strFormat, *params):
    print(tcolors.red + "ERR > " + strFormat.format(*params) + tcolors.reset)
    SetPass(False)

def ApiMsgStr(strMsg):
    print(tcolors.green + "Api > " + strMsg + tcolors.reset)

def ApiMsg(strFormat, *params):
    print(tcolors.green + "Api > " + strFormat.format(*params) + tcolors.reset)

def ListMsg(strFormat, *params):
    print(tcolors.yellow + "Lst > " + strFormat.format(*params) + tcolors.reset)

def LoopMsg(strFormat, *params):
    print("Loop> " + strFormat.format(*params))
#endregion

#region Support functions
def RemoveFile(pathname):
    try:
        if (FileExists(pathname)):
            os.remove(pathname)
        return True
    except:
        print("exception in RemoveFile({})".format(pathname))
        traceback.print_exc()
        return False

def FileExists(pathname):
    return os.path.exists(pathname)

def CopyFile(srcPath, dstPath, makeWritable, deldst = False):
    try:
        if(deldst):
            RemoveFile(dstPath)
        srcP = os.path.realpath(srcPath)
        dstP = os.path.realpath(dstPath)
        if not CheckIsLinux():
            shutil.copyfile(srcP, dstP)
            if (makeWritable):
                os.chmod(dstP, stat.S_IWRITE)
        else:
            p = Popen(['cp', '-p', '--preserve', srcP, dstP])
            p.wait()
            if (makeWritable):
                p = Popen(['chmod', '+w', dstP])
                p.wait()
    except:
        traceback.print_exc()
        return FILE_ACCESS_ERR

    return REC_OK

def CheckPDFPageCountWithRecPDF(fileName):
    InfoMsg("Open PDF file for checking page count with RecPDF -- rPdfOpen()")
    rc, rpd = rPdfOpen(fileName, None)
    if (rc != REC_OK):
        ErrMsg("Error code = {}\n", rc)
        rPdfQuit()
        kRecQuit()
        return False

    InfoMsg("Get page count of PDF file with RecPDF -- rPdfGetPageCount()")
    rc, np = rPdfGetPageCount(rpd)
    if (rc != REC_OK):
        ErrMsg("Error code = {}\n", rc)
        rPdfQuit()
        kRecQuit()
        return False

    InfoMsg("Close PDF file with RecPDF -- rPdfClose()")
    rPdfClose(rpd)

    ApiMsg("Page count = {}", np)
    return True

def CheckIsLinux():
    if "Linux" in platform.system():
        return True
    else:
        return False

def ReadFileToString(filePath, doBinaryMode = False):
    try:
        mode = 'rt'
        if (doBinaryMode):
            mode = 'rb'
        file = io.open(filePath, mode)
        return file.read()
    except:
        print("ReadFileToString(): exception")
        traceback.print_exc()
        return ""

def padLeftSpaces(inputString, padchar, length):
    inputStringLen = len(inputString)
    if (inputStringLen >= length):
        return inputString
    sb = ""
    while (len(sb) < length - inputStringLen):
        sb += padchar
    sb += inputString
    return sb

def padRightSpaces(inputString, padchar, length):
    inputStringLen = len(inputString)
    if (inputStringLen >= length):
        return inputString
    sb = inputString
    while (len(sb) < length):
        sb += padchar
    return sb

def InitEnabledLanguagesArray(langs, enabledLangsList):
    for i in range(LANG_SIZE):
        langs[i] = LANG_DISABLED
    for l in enabledLangsList:
        langs[l] = LANG_ENABLED

def CreateEnabledLanguagesArray(enabledLangsList):
    langs = IntArray(LANG_SIZE)
    InitEnabledLanguagesArray(langs, enabledLangsList)
    return langs

def ListEnabledLanguages():
    rc, langs = kRecGetLanguages(SID)
    if REC_OK == rc:
        for i in range(LANG_SIZE):
            if LANG_ENABLED == langs[i]:
                print("LANGS[{}] is ENABLED".format(i))

def InitEnabledBarTypesArray(barTypes, enabledBarTypesList):
    for i in range(BAR_SIZE):
        barTypes[i] = BAR_DISABLED
    for l in enabledBarTypesList:
        barTypes[l] = BAR_ENABLED

def CreateEnabledBarTypesArray(enabledBarTypesList):
    barTypes = IntArray(BAR_SIZE)
    InitEnabledBarTypesArray(barTypes, enabledBarTypesList)
    return barTypes

def ListEnabledBarTypes():
    rc, barTypes = kRecGetBarTypes(SID)
    if REC_OK == rc:
        for i in range(BAR_SIZE):
            if BAR_ENABLED == barTypes[i]:
                print("BAR_TYPES[{}] is ENABLED".format(i))

def GetPass():
    return PASS

def GetOutputFolder():
    return outputfolder

def SetOutputFolder(value):
    global outputfolder
    outputfolder = value

def CreateDefaultOutputFolder():
    # Windows-only
    envKey = "USERPROFILE"
    if envKey not in os.environ:
        return "."
    out_folder = os.environ[envKey]
    out_folder += os.path.sep + "OmniPage"
    if not os.path.exists(out_folder):
        os.makedirs(out_folder)
    out_folder += os.path.sep + "CSDK22"
    if not os.path.exists(out_folder):
        os.makedirs(out_folder)
    out_folder += os.path.sep + "SamplesOutput"
    if not os.path.exists(out_folder):
        os.makedirs(out_folder)
    out_folder += os.path.sep + "Python"
    if not os.path.exists(out_folder):
        os.makedirs(out_folder)
    return out_folder

def GetSourceCodeFolder(obj):
    src = inspect.getfile(obj)
    return os.path.dirname(os.path.realpath(src))

def SetPass(value):
    global PASS
    PASS = value
#endregion
