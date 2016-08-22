### OPERAUTOMATION v2.0 ###

#VARs
NUMBER_OF_ATTEMPTS = 2
RETRY_TIME = 4
VARS = {}

#Not EDIT
import os
import os.path
import sys
import requests

print("#####   OPERAUTOMATION v2.0   #####\n")
secret = requests.post('http://master-dealer.it/operautomation-activate.php', data={'token': 'Jzjm47t0hgPa4j8gjf565czieY86f02z'})
if secret.status_code == 200:
    if secret.text != 'Token provided is not valid':
        arg = 1
        if len(sys.argv) <= arg:
            key = input('Chiave di accesso:   ')
        else:
            key = sys.argv[arg]
        if key != secret.text:
            input('\nLa chiave di accesso inserita non è valida.\nContattare l\'amministratore per ottenere una nuova chiave')
            sys.exit()
    else:
        input('Il token non è valido. Contattare l\'amministratore del software per risolvere il problema')
        sys.exit()
else:
    input('Errore di connessione con il server (Error ' + str(secret.status_code) + ')')
    sys.exit()

PROGRAM_PATH = "rom"
currProgName = "out"

'''if len(sys.argv) > 1:
    PROGRAM_PATH = sys.argv[1] + "\\" + PROGRAM_PATH
    currProgName = sys.argv[1] + "\\" + currProgName'''

workspace = ""

print("\n\nBenvenuto in Operautomation, il sistema che semplifica il sistema.\n\nScegli il programma da avviare:\n")

stopCycle = 0
argP = arg + 1
while(stopCycle == 0):
    programs = next(os.walk(PROGRAM_PATH))[1]

    n = 1
    while(n <= len(programs)):
        print(str(n) + ". " + programs[n - 1])
        n = n + 1

    if len(sys.argv) <= 2:
        progN = eval(input('\nQuale programma vorresti avviare: '))
    else:
        progN = eval(sys.argv[argP])

    if progN > 0 and progN <= len(programs):
        PROGRAM_PATH += "\\" + programs[progN - 1]
        currProgName += "\\" + programs[progN - 1]
        if os.path.isfile(PROGRAM_PATH + "\program.prog"):
            PROGRAM_PATH += "\program.prog"
            print("\n\nProgramma selezionato: " + PROGRAM_PATH + "\n\n\nSto lavorando per te...")
            stopCycle = 1
        else:
            print("\n\nSono stati trovati sottoprogrammi. Seleziona quale aprire:\n")
            argP = arg + 1
    else:
        stopCycle = 1
        print("\nApplicazione non trovata\n")
        input("Premi Enter per uscire")
    argP += 1

# main execution

import time
import PyPDF2
#from selenium import webdriver
#from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoAlertPresentException
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoSuchFrameException
from selenium.webdriver.support.ui import Select
from os.path import basename
import pandas as pd
import xlrd
import html
from datetime import datetime, timedelta
import calendar
import json
from calendar import monthrange
from PIL import Image
import pytesseract
import subprocess
import re

browser = None
main_window_handle = None
elem = None
excelDF = None
registerT = None
registerX = None
registerY = None
registerZ = None
registerW = None

from selenium import webdriver

# Functions

def gatherInstructions(inData):
    instructions = []
    nChar = 0
    parCounter = 0
    startPos = 0
    endPos = 0
    for c in inData:
        if str(c) == "(":
            parCounter += 1
            if parCounter == 1:
                startPos = nChar
        if str(c) == ")":
            parCounter -= 1
            if parCounter == 0:
                endPos = nChar
        if parCounter == 0 and endPos != 0:
            instructions.append(inData[startPos+1:endPos])
        nChar += 1
    return instructions

def executeInstructions(instructions = []):
    for f in instructions:
        if not f:
            return
        instr = f.split("(", 1)[0]
        if len(f.split("(", 1)[1][:-1]) > 0:
            args = f.split("(", 1)[1][:-1]
            if "(" in args:
                preargs = args[:args.index("(")]
                args = replaceRegisterValues(preargs) + "," + args[args.index("("):]
            else:
                args = replaceRegisterValues(args)
            globals()[instr](args.replace(">>", "").replace("<<", ""))
        else:
            globals()[instr]()

def replaceRegisterValues(inputStr):
    return inputStr.replace("|registerT|", str(registerT)).replace("|registerX|", str(registerX)).replace("|registerY|", str(registerY)).replace("|registerZ|", str(registerZ)).replace("|registerW|", str(registerW))

def prepareWorkspace():
    global workspace
    workspace = os.getcwd() + "\\" + currProgName
    if not os.path.exists(workspace):
        os.makedirs(workspace)
def setWorkspace(path):
    global workspace
    workspace = os.getcwd() + "\\out\\" + path
def restoreWorkspace():
    global workspace
    workspace = os.getcwd() + "\\" + currProgName

def initFirefox(savepath = ""):
    global browser
    global main_window_handle

    #caps = DesiredCapabilities.FIREFOX
    #caps["marionette"] = True
    #caps['acceptSslCerts'] = True
    
    profile = webdriver.FirefoxProfile()
    profile.accept_untrusted_certs = True
    profile.set_preference('browser.download.folderList', 2)
    profile.set_preference('browser.download.manager.showWhenStarting', False)
    profile.set_preference('browser.download.dir', workspace + "\\" + savepath)
    profile.set_preference('browser.helperApps.neverAsk.saveToDisk', ('application/vnd.ms-excel'))
    browser = webdriver.Firefox(firefox_profile=profile)#, capabilities=caps
    while not main_window_handle:
        main_window_handle = browser.current_window_handle

def initIExplorer(savepath = ""):
    global browser
    browser = webdriver.Ie("plugin\\IEDriverServer_x32.exe")

def nav(url):
    browser.get(url)

def wait(seconds):
    time.sleep(float(seconds))

def close():
    browser.close()

def appExit():
    if len(sys.argv) <= 2:
        input("Premi Invio per terminare")
    sys.exit()

def inputF(information):
    global registerT
    registerT = input(information)

def selectFrame(fID = -1):
    if fID == -1:
        browser.switch_to.default_content()
    else:
        i = 0
        while(i < NUMBER_OF_ATTEMPTS):
            f = browser.find_elements_by_tag_name("frame")
            if len(f) > 0:
                if fID.isdigit():
                    browser.switch_to.frame(f[int(fID)])
                    i = NUMBER_OF_ATTEMPTS
                else:
                    for ff in f:
                        if ff.get_attribute("name") == fID:
                            browser.switch_to.frame(ff)
                            i = NUMBER_OF_ATTEMPTS
                            return
                    time.sleep(RETRY_TIME)
            else:
                time.sleep(RETRY_TIME)
            i += 1

def alertText():
    global registerT
    registerT = browser.switch_to_alert().text

def alertAccept(instrunctions):
    instr = gatherInstructions(instrunctions)[0]
    instrFail = gatherInstructions(instrunctions)[1]
    try:
        browser.switch_to_alert().accept()
        executeInstructions(gatherInstructions(instr))
    except NoAlertPresentException as e:
        i = 0
        while(i < NUMBER_OF_ATTEMPTS):
            try:
                browser.switch_to_alert().accept()
                i = NUMBER_OF_ATTEMPTS
            except NoAlertPresentException as e:
                wait(1)
            i += 1
        executeInstructions(gatherInstructions(instrFail))
def alertDismiss(instrunctions):
    instr = gatherInstructions(instrunctions)[0]
    instrFail = gatherInstructions(instrunctions)[1]
    try:
        browser.switch_to_alert().dismiss()
        executeInstructions(gatherInstructions(instr))
    except NoAlertPresentException as e:
        i = 0
        while(i < NUMBER_OF_ATTEMPTS):
            try:
                browser.switch_to_alert().dismiss()
                i = NUMBER_OF_ATTEMPTS
            except NoAlertPresentException as e:
                wait(1)
            i += 1
        executeInstructions(gatherInstructions(instrFail))

def switchToPopup():
    popup_handle = None
    while not popup_handle:
        for handle in browser.window_handles:
            if handle != main_window_handle:
                popup_handle = handle
                break
    browser.switch_to.window(popup_handle)
def switchToMainWindow():
    browser.switch_to.window(main_window_handle)

def selectId(idName):
    global elem
    informations = idName.split(",")
    i = 0
    while(i < NUMBER_OF_ATTEMPTS):
        if i > 0:
            print("Error: id \"" + informations[0] + "\" not found. Retrying...")
            wait(RETRY_TIME)
        try:
            elem = browser.find_element_by_id(informations[0])
            return 1
        except NoSuchElementException:
            pass
        i += 1
    print("Retrying failed. Executing exception...")
    executeInstructions(gatherInstructions(informations[1]))
    return 0

def selectClass(className):
    global elem
    informations = className.split(",")
    i = 0
    while(i < NUMBER_OF_ATTEMPTS):
        if i > 0:
            print("Error: class \"" + informations[0] + "\" not found. Retrying...")
            wait(RETRY_TIME)
        try:
            if informations[1] == "-1":
                elem = []
                elem = browser.find_elements_by_class_name(informations[0])
                return 1
            else:
                elem = None
                elem = browser.find_elements_by_class_name(informations[0])
                if len(elem) > 0:
                    elem = elem[int(informations[1])]
                    return 1
        except NoSuchElementException:
            pass
        i += 1
    print("Retrying failed. Executing exception...")
    executeInstructions(gatherInstructions(informations[2]))
    return 0

def selectName(name):
    global elem
    informations = name.split(",")
    i = 0
    while(i < NUMBER_OF_ATTEMPTS):
        if i > 0:
            print("Error: name \"" + informations[0] + "\" not found. Retrying...")
            wait(RETRY_TIME)
        try:
            if informations[1] == "-1":
                elem = []
                elem = browser.find_elements_by_name(informations[0])
                return 1
            else:
                elem = None
                elem = browser.find_elements_by_name(informations[0])
                if len(elem) > 0:
                    elem = elem[int(informations[1])]
                    return 1
        except NoSuchElementException:
            pass
        i += 1
    print("Retrying failed. Executing exception...")
    executeInstructions(gatherInstructions(informations[2]))
    return 0

def selectByTextInside(text):
    global elem
    informations = text.split(",")
    i = 0
    while(i < NUMBER_OF_ATTEMPTS):
        if i > 0:
            print("Error: text \"" + informations[0] + "\" not found. Retrying...")
            wait(RETRY_TIME)
        try:
            elem = []
            elem = browser.find_elements_by_xpath("//*[contains(text(), '" + informations[0] + "')]")
            if len(elem) > 0:
                if informations[1] != "-1":
                    elem = elem[int(informations[1])]
                return 1
        except NoSuchElementException:
            pass
        i += 1
    print("Retrying failed. Executing exception...")
    executeInstructions(gatherInstructions(informations[2]))
    return 0

def getChildren():
    global elem
    elem = elem.find_elements_by_xpath(".//*")

def selectFromArray(elemId):
    global elem
    elem = elem[int(elemId)]

def selectFromArrayRegister(elemId):
    global registerT
    registerT = registerT[int(elemId)]

def sendEnter():
    elem.send_keys(Keys.RETURN)

def sendClick():
    elem.click()

def selectValue(information):
    select = Select(elem)
    informations = information.split(",")
    if int(informations[0]) == 0:
        select.select_by_value(informations[1])
    elif int(informations[0]) == 1:
        select.select_by_visible_text(informations[1])

def setText(text):
    elem.send_keys(text)

def getText():
    global registerT
    registerT = elem.text

def getAttribute(attrName):
    global registerT
    registerT = elem.get_attribute(attrName)

cycleI = {}
cycleMax = {}
def forFun(forInfos):
    global cycleI
    global cycleMax
    cycleInfos = forInfos.split(",", 1)
    varName = cycleInfos[0].split(";")[0].split("=")[0]
    cycleI[varName] = int(cycleInfos[0].split(";")[0].split("=")[1])
    if cycleInfos[0].split(";")[1].isalpha():
        cycleMax[varName] = VARS[cycleInfos[0].split(";")[1]]
    else:
        cycleMax[varName] = int(cycleInfos[0].split(";")[1])
    # EXECUTE INSTRUCTIONS
    instructions = gatherInstructions(cycleInfos[1])
    while(cycleI[varName] < cycleMax[varName]):
        currI = 0
        while(currI < len(instructions)):
            instructions[currI] = instructions[currI].replace("|i|", ">>" + str(cycleI[varName]) + "<<")
            currI += 1
        executeInstructions(instructions)
        #reset instructions to |i|
        currI = 0
        while(currI < len(instructions)):
            instructions[currI] = instructions[currI].replace(">>" + str(cycleI[varName]) + "<<", "|i|")
            currI += 1
        #reset completed
        cycleI[varName] = cycleI[varName] + 1

cycleIDs = {}
def forEach(forInfos):
    cycleInfos = forInfos.split(",", 2)
    varName = cycleInfos[0]
    cycleIDs[varName] = [1, None, 0]
    instructions = gatherInstructions(cycleInfos[2])
    elems = {}
    if cycleInfos[1] == "0":
        elems = elem
    else:
        elems = VARS[cycleInfos[1]]
    for e, val in elems.items():
        cycleIDs[varName][1] = val
        if cycleIDs[varName][0] > 0:
            executeInstructions(instructions)

def breakF(varName):
    global cycleI
    global cycleMax
    global cycleIDs
    if varName in cycleIDs:
        cycleIDs[varName] = [0, None, 0]
    elif varName in cycleI:
        cycleI[varName] = cycleMax[varName]
  
def ifElemExists(ifStatement):
    # GATHER INSTRUCTIONS
    statements = gatherInstructions(ifStatement.split(",", 1)[1])
    instructions = gatherInstructions(statements[0])
    elseInstructions = gatherInstructions(statements[1])
    # END GATHERING
    ifOrElse = 0
    condition = ifStatement.split(",", 1)[0]
    if condition.startswith("id||"):
        if selectId(ifStatement.split(",", 1)[0][4:] + ",()") == 0:
            ifOrElse = 1
        else:
            ifOrElse = 0
    if condition.startswith("class||"):
        if selectClass(ifStatement.split(",", 1)[0][7:] + ",-1,()") == 0:
            ifOrElse = 1
        else:
            ifOrElse = 0
    if condition.startswith("name||"):
        if selectName(ifStatement.split(",", 1)[0][6:] + ",-1,()") == 0:
            ifOrElse = 1
        else:
            ifOrElse = 0
    if condition.startswith("text||"):
        if selectByTextInside(ifStatement.split(",", 1)[0][6:] + ",-1,()") == 0:
            ifOrElse = 1
        else:
            ifOrElse = 0
    if ifOrElse == 0:
        executeInstructions(instructions)
    elif ifOrElse == 1:
        executeInstructions(elseInstructions)

def ifIsInArrayVAR(ifStatement):
    # GATHER INSTRUCTIONS
    statements = gatherInstructions(ifStatement.split(",", 1)[1])
    instructions = gatherInstructions(statements[0])
    elseInstructions = gatherInstructions(statements[1])
    # END GATHERING
    ifOrElse = 0
    condition = ifStatement.split(",", 1)[0]
    if str(registerT) in VARS[condition]:
        ifOrElse = 0
    else:
        ifOrElse = 1
    if ifOrElse == 0:
        executeInstructions(instructions)
    elif ifOrElse == 1:
        executeInstructions(elseInstructions)

'''def dateCompare(register):   TODO
    global registerT
    global registerX
    global registerY
    global registerZ
    global registerW
    date0 = datetime.strptime(registerT, '%d/%m/%Y')
    date1 = datetime.strptime(globals()[register], '%d/%m/%Y')
    if date0 > date1:
        registerT = 1
    elif date0 == date1:
        registerT = 0
    elif date0 < date1:
        registerT = -1'''

def dateFormatConversion(formatsDatas):
    global registerT
    oldDateFormat = formatsDatas.split(",")[0]
    newDateFormat = formatsDatas.split(",")[1]
    registerT = str(datetime.strptime(registerT, oldDateFormat).strftime(newDateFormat))

def ifLesser(ifStatement):
    # GATHER INSTRUCTIONS
    statements = gatherInstructions(ifStatement.split(",", 1)[1])
    instructions = gatherInstructions(statements[0])
    elseInstructions = gatherInstructions(statements[1])
    # END GATHERING
    if registerT < int(ifStatement.split(",", 1)[0]):
        executeInstructions(instructions)
    else:
        executeInstructions(elseInstructions)

def ifEqual(ifStatement):
    global registerT
    # GATHER INSTRUCTIONS
    statements = gatherInstructions(ifStatement.split(",", 2)[2])
    instructions = gatherInstructions(statements[0])
    elseInstructions = gatherInstructions(statements[1])
    # END GATHERING
    if registerT == None:
        registerT = ""
    condition = ifStatement.split(",", 2)[0]
    if ifStatement.split(",", 2)[1] != "-1":
        if isinstance(VARS[condition], list):
            condition = VARS[condition][int(ifStatement.split(",", 2)[1])]
        else:
            condition = VARS[condition]
    if str(registerT).replace("nan", "") == condition:
        executeInstructions(instructions)
    else:
        executeInstructions(elseInstructions)

def getFiles(path = ""):
    global elem
    elem = []
    for dirname, dirnames, filenames in os.walk(workspace + "\\" + path):
        for filename in filenames:
            elem.append(os.path.join(dirname, filename))

def readPDF(path = ""):
    global registerT
    if "\\" not in path:
        path = cycleIDs[path][1]
    if not path.endswith(".pdf"):
        return
    registerT = ""
    pdfFileObj = open(path, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    currPage = 0
    while(currPage < pdfReader.getNumPages()):
        pageObj = pdfReader.getPage(currPage)
        registerT = registerT + pageObj.extractText()
        currPage = currPage + 1

def getFileName(path = ""):
    global registerT
    if "\\" not in path:
        path = cycleIDs[path][1]
    registerT = basename(path)[:-4]

def getFullFileName(path = ""):
    global registerT
    if "\\" not in path:
        registerT = cycleIDs[path][1]
    else:
        registerT = os.getcwd() + "\\out\\" + path

def extractText(information):
    global registerT
    try:
        informations = information.split(",")
        registerT = registerT[registerT.index(informations[0]) + len(informations[0]):registerT.index(informations[1])]
    except ValueError:
        pass

def replaceText(information):
    global registerT
    informations = information.split(",")
    registerT = registerT.replace(informations[0].replace("\\n", "\x0a"), informations[1].replace("\\n", "\x0a"))

'''def splitText(information):
    global registerT
    sys.exit()
    informations = information.split(",", 1)
    registerT = registerT.split(informations[1])[int(infofrmations[0])]'''

def substr(stringDatas):
    global registerT
    informations = stringDatas.split(":")
    if informations[0] == "":
        start = None
    else:
        start = int(informations[0])
    if informations[1] == "":
        end = None
    else:
        end = int(informations[1])
    registerT = registerT[start:end]

def createArrayFromString(information):
    global VARS
    VARS[information.split(",")[0]] = information.split(",")[1].split(information.split(",")[2])
    #VARS[information.split(",")[0]] = [x for x in VARS[information.split(",")[0]] if x]

def mergeExcel(information):
    arrayExcelFiles = information.split(",", 1)[0]
    exportFilePath = information.split(",", 1)[1]

    wkbk = xlwt.Workbook()
    outsheet = wkbk.add_sheet('Sheet1', cell_overwrite_ok = True)

    dats = VARS[arrayExcelFiles]
    if not isinstance(dats, list):
        dats = VARS[arrayExcelFiles].split(",")
    for f in dats:
        e = f
        if "\\" not in f:
            e = workspace + "\\" + f
        insheet = xlrd.open_workbook(e).sheets()[0]
        for row_idx in iter(range(insheet.nrows)):
            for col_idx in iter(range(insheet.ncols)):
                if insheet.cell_value(row_idx, col_idx) != "":
                    outsheet.write(row_idx, col_idx, insheet.cell_value(row_idx, col_idx))
    wkbk.save(workspace + "\\" + exportFilePath)

def appendExcel(information):
    arrayExcelFiles = information.split(",")[0]
    exportFilePath = information.split(",")[1]
    startingRow = information.split(",")[2]

    wkbk = xlwt.Workbook()
    outsheet = wkbk.add_sheet('Sheet1')

    dats = VARS[arrayExcelFiles]
    if not isinstance(dats, list):
        dats = VARS[arrayExcelFiles].split(",")
    outrow_idx = 0
    for f in dats:
        e = f
        if "\\" not in f:
            e = workspace + "\\" + f
        insheet = xlrd.open_workbook(e).sheets()[0]
        for row_idx in iter(range(insheet.nrows)):
            if row_idx >= int(startingRow) or outrow_idx == 0:
                for col_idx in iter(range(insheet.ncols)):
                    if insheet.cell_value(row_idx, col_idx) != "":
                        outsheet.write(outrow_idx, col_idx, insheet.cell_value(row_idx, col_idx))
                outrow_idx += 1
    wkbk.save(workspace + "\\" + exportFilePath)

def saveToFile(information):
    global registerT
    path = information.split(",")[0]
    convert = information.split(",")[1]
    export = open(workspace + "\\" + path, "w")
    datas = registerT.split("\\n")
    for line in datas:
        if line != "":
            ins = line
            if convert == "1":
                ins += "\n"
            else:
                ins += "\\n"
            export.write(ins)
    export.close()

def loadFromFile(path):
    global registerT
    if not os.path.isfile(workspace + "\\" + path):
        registerT = ""
    else:
        file = open(workspace + "\\" + path, "r")
        registerT = file.read()
        file.close()

def setRegister(newVal):
    global registerT
    registerT = newVal

def getLength():
    global registerT
    registerT = len(registerT)

def getCount():
    global registerT
    registerT = len(elem)

def getCurrDate(dateFormat):
    global registerT
    registerT = str(time.strftime(dateFormat))

def addDays(information):
    global registerT
    dateFormat = information.split(",")[0]
    nDays = int(information.split(",")[1])
    registerT = (datetime.strptime(registerT, dateFormat) - timedelta(days=nDays)).strftime(dateFormat)
def subDays(information):
    global registerT
    dateFormat = information.split(",")[0]
    nDays = int(information.split(",")[1])
    registerT = (datetime.strptime(registerT, dateFormat) - timedelta(days=nDays)).strftime(dateFormat)

def createFolder(directory):
    if not os.path.exists(workspace + "\\" + directory):
        os.makedirs(workspace + "\\" + directory)

def setVar(information):
    global VARS
    values = gatherInstructions(information.split(",", 1)[1])
    result = {}
    if len(values) == 0:
        result = ""
    elif len(values) == 1:
        result = replaceRegisterValues(values[0])
    else:
        for (x, val) in enumerate(values):
            if ":" in val:
                result[val.split(":", 1)[0]] = replaceRegisterValues(val.split(":", 1)[1])
            else:
                result[x] = replaceRegisterValues(val)
    if information.split(",", 1)[0].isdigit():
        VARS[int(information.split(",")[0])] = result
    else:
        VARS[information.split(",")[0]] = result

def unsetVar(varName):
    global VARS
    del VARS[varName]
def getVar(varName):
    global registerT
    registerT = VARS[varName]
def getVars():
    global registerT
    registerT = VARS

def numToMonth(monthNum):
    global registerT
    registerT = calendar.month_name[int(monthNum)]

def getDaysInMounth(yearAndmonth):
    global registerT
    year = int(yearAndmonth.split(",")[0])
    month = int(yearAndmonth.split(",")[1])
    registerT = monthrange(year, month)[1]

def renameFile(information):
    os.rename(workspace + "\\" + information.split(",")[0], workspace + "\\" + information.split(",")[1])

def moveBetweenWorkspaces(information):
    informations = information.split(",")
    currFile = ""
    if informations[2] == "0":
        currFile = os.getcwd() + "\\out\\" + informations[0]
    elif informations[2] == "1":
        currFile = cycleIDs[informations[0]][1]
    newFile = os.getcwd() + "\\out\\" + informations[1]
    if os.path.isfile(newFile):
        os.remove(newFile)
    os.rename(currFile, newFile)

def removeFile(pathToFile):
    if "\\" not in pathToFile:
        os.remove(cycleIDs[pathToFile][1])
    else:
        if pathToFile[0] == "\\":
            pathToFile = pathToFile[1:]
        os.remove(workspace + "\\" + pathToFile)

def zpaddingLeft(information):
    global registerT
    registerT = str(registerT).zfill(int(information))

def mouseClickPosition(information):
    pyautogui.click(4 + int(browser.get_window_position()["x"]) + int(information.split(",")[0]), 4 + int(browser.get_window_position()["y"]) + int(information.split(",")[1]))

def clearElem():
    elem.clear()

def createJson():
    global registerT
    registerT = json.dumps(registerT, sort_keys=True, indent=4)

def sendPOST(url):
    global registerT
    url = information.split(",")[0]
    varName = information.split(",")[1]
    registerT = requests.post(url, data={varName: registerT})

def storX():
    global registerX
    registerX = registerT
def storY():
    global registerY
    registerY = registerT
def storZ():
    global registerZ
    registerZ = registerT
def storW():
    global registerW
    registerW = registerT
def lodX():
    global registerT
    registerT = registerX
def lodY():
    global registerT
    registerT = registerY
def lodZ():
    global registerT
    registerT = registerZ
def lodW():
    global registerT
    registerT = registerW
def addTextToRegistry(text):
    global registerT
    registerT = str(registerT) + text
def addVarFECycleToRegistry(varName):
    global registerT
    if registerT == None:
        registerT = ""
    registerT += cycleIDs[varName][1]
def addValToT(val):
    global registerT
    registerT = str(int(registerT) + int(val))
def subValToT(val):
    global registerT
    registerT = str(int(registerT) - int(val))
def addX():
    global registerT
    global registerX
    if registerT is None:
        registerT = ""
    if registerX is None:
        registerX = ""
    registerT = str(registerT) + str(registerX)
def addValToX(val):
    global registerX
    registerX = str(int(registerX) + int(val))
def addY():
    global registerT
    global registerY
    if registerT is None:
        registerT = ""
    if registerY is None:
        registerY = ""
    registerT = str(registerT) + str(registerY)
def addValToY(val):
    global registerY
    registerY = str(int(registerY) + int(val))
def addZ():
    global registerT
    global registerZ
    if registerT is None:
        registerT = ""
    if registerZ is None:
        registerZ = ""
    registerT = str(registerT) + str(registerZ)
def addValToZ(val):
    global registerZ
    registerZ = str(int(registerZ) + int(val))
def addW():
    global registerT
    global registerW
    if registerT is None:
        registerT = ""
    if registerW is None:
        registerW = ""
    registerT = str(registerT) + str(registerW)
def addValToW(val):
    global registerW
    registerW = str(int(registerW) + int(val))
def clearT(information = ""):
    global registerT
    registerT = None
    if information == "s":
        registerT = ""
def clearX(information = ""):
    global registerX
    registerX = None
    if information == "s":
        registerX = ""
def clearY(information = ""):
    global registerY
    registerY = None
    if information == "s":
        registerY = ""
def clearZ(information = ""):
    global registerZ
    registerZ = None
    if information == "s":
        registerZ = ""
def clearW(information = ""):
    global registerW
    registerW = None
    if information == "s":
        registerW = ""
def clearAll():
    global registerT
    global registerX
    global registerY
    global registerZ
    global registerW
    registerT = None
    registerX = None
    registerY = None
    registerZ = None
    registerW = None
def printT():
    print(registerT)
def printX():
    print(registerX)
def printY():
    print(registerY)
def printZ():
    print(registerZ)
def printW():
    print(registerW)

# EXCEL conversion functions
import xlwt

def convertToExcel(pathToFile):
    if not os.path.isfile(workspace + "\\" + pathToFile):
        wait(15)
    file = open(workspace + "\\" + pathToFile, 'r')
    export_to_xls(html_table_to_excel(file.read()), workspace + "\\" + pathToFile[:-4] + " - converted.xls")

def html_table_to_excel(table):
    data = {}
    table = table[table.index('<tr>') + 4:table.index('</table>')]
    rows = table.replace("\n", "").replace("&nbsp;", " ").replace("&quot;", "").replace("</tr>", "").replace("<th>", "<td>").replace("</th>", "</td>").split('<tr>')
    for (x, row) in enumerate(rows):
        columns = row.split('</td>')
        data[x] = {}
        for (y, col) in enumerate(columns):
            data[x][y] = col.replace('<td>=', '').replace('<td>', '')
    return data

def export_to_xls(data, filename, title='Sheet1'):
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet(title)
    style = xlwt.XFStyle()
    style.num_format_str = "@"
    for x in data.keys():
        for y in data[x].keys():
            worksheet.write(x, y, html.unescape(data[x][y]), style)
    workbook.save(filename)

def getRowCount(pathToFile):
    global registerT
    if not "\\" in pathToFile:
        pathToFile = cycleIDs[pathToFile][1]
    else:
        pathToFile = workspace + "\\" + pathToFile[1:]
    xl = pd.ExcelFile(pathToFile)
    df = xl.parse(xl.sheet_names[0])
    registerT = df.shape[0]

def openExcel(information):
    global excelDF
    pathToFile = information.split(",")[0]
    sheet = information.split(",")[1]
    if not "\\" in pathToFile:
        pathToFile = cycleIDs[pathToFile][1]
    else:
        pathToFile = workspace + "\\" + pathToFile[1:]
    #xl = pd.ExcelFile(pathToFile, keep_default_na=False)
    #excelDF = xl.parse(xl.sheet_names[int(sheet)])
    excelDF = pd.read_excel(open(pathToFile,'rb'), sheetname=int(sheet), keep_default_na=False)

def readExcelCell(cellCoordinates):
    global registerT
    try:
        registerT = excelDF.iloc[int(cellCoordinates.split("-")[0]), int(cellCoordinates.split("-")[1])]
    except:
        registerT = ""

def newExcel(inputData):
    data = {}
    varIndex = inputData.split(",")[0]
    x = int(inputData.split(",")[1])
    y = int(inputData.split(",")[2])
    for (i, row) in enumerate(VARS[varIndex]):
        data[i + y] = {x: row}
    export_to_xls(data, workspace + "\\" + 'temp.xls')
    
def filterExcelByColumnVal(information):
    pathToFile = information.split(",")[0]
    col = information.split(",")[1]
    valComp = information.split(",")[2]
    if "\\" not in pathToFile:
        pathToFile = cycleIDs[pathToFile][1]
    elif pathToFile[0] == "\\":
        pathToFile = workspace + "\\" + pathToFile[1:]
    xl = pd.ExcelFile(pathToFile)
    df = xl.parse(xl.sheet_names[0])
    df = df[df[col] == valComp]
    removeFile("\\" + basename(pathToFile))
    writer = pd.ExcelWriter(pathToFile)
    df.to_excel(writer, 'Sheet1', index = None)
    writer.save()

pytesseract.pytesseract.tesseract_cmd = r'plugin\Tesseract-OCR\tesseract.exe'
def image2Text(pathToImage):
    global registerT
    registerT = pytesseract.image_to_string(Image.open(workspace + "\\" + pathToImage))

def extractToken(token):
    tokImage = Image.open("tokens\\" + token)
    tokImageC = tokImage.crop((689, 459, 689 + 102, 459 + 28))
    tokImageC.save("temp\\" + token)
def token2text(token):
    global registerT
    extractToken(token)
    result = ""
    i = 0
    x = 0
    y = 0
    while(i < 6):
        tokImage = Image.open("temp\\" + token)
        if i == 3:
            x += 6
        tokImageC = tokImage.crop((x, y, x + 15, y + 26))
        tokImageC.save("temp\\" + token[:-4] + "-" + str(i) + ".jpg")
        result += subprocess.Popen(["plugin\\ssocr.exe", "-t", "55", "-d", "-1", "-n", "2", "temp\\" + token[:-4] + "-" + str(i) + ".jpg"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell = True).communicate()[0].decode()[:-1] #"-n", "3", "-i", "1", 
        os.remove("temp\\" + token[:-4] + "-" + str(i) + ".jpg")
        i += 1
        x += 16
    os.remove("temp\\" + token)
    registerT = result.replace(" ", "").replace("_", "")#re.search(r"^-?[0-9]+$", result).group(0)

# End functions

file = open(PROGRAM_PATH, 'r')

lines = file.read().replace("\n", "").split("*")

for line in lines:
    instruction = line.replace("for(", "forFun(").replace("input", "inputF").split("(", 1)[0]
    if instruction == "forEach" or instruction == "forFun" or instruction == "ifEqual":
        inf = line.replace("for(", "forFun(").replace("input", "inputF").split("(", 1)[1][:-1].replace("break", "breakF")
        information = replaceRegisterValues(inf.split(",", 1)[0]) + "," + inf.split(",", 1)[1]
    else:
        information = replaceRegisterValues(line.replace("for(", "forFun(").replace("input", "inputF").split("(", 1)[1][:-1].replace("break", "breakF"))
        
    if instruction == "prepareWorkspace":
        prepareWorkspace()
    elif instruction == "setWorkspace":
        setWorkspace(information)
    elif instruction == "restoreWorkspace":
        restoreWorkspace()
    elif instruction == "initFirefox":
        initFirefox(information)
    elif instruction == "initIExplorer":
        initIExplorer(information)
    elif instruction == "nav":
        nav(information)
    elif instruction == "wait":
        wait(information)
    elif instruction == "close":
        close()
    elif instruction == "appExit":
        appExit()
    elif instruction == "inputF":
        inputF(information)
    elif instruction == "selectFrame":
        if information == "":
            selectFrame()
        else:
            selectFrame(information)
    elif instruction == "alertText":
        alertText()
    elif instruction == "alertAccept":
        alertAccept(information)
    elif instruction == "alertDismiss":
        alertDismiss(information)
    elif instruction == "selectId":
        selectId(information)
    elif instruction == "switchToPopup":
        switchToPopup()
    elif instruction == "switchToMainWindow":
        switchToMainWindow()
    elif instruction == "closeWindow":
        closeWindow()
    elif instruction == "selectClass":
        selectClass(information)
    elif instruction == "selectName":
        selectName(information)
    elif instruction == "selectByTextInside":
        selectByTextInside(information)
    elif instruction == "getChildren":
        getChildren(information)
    elif instruction == "selectFromArray":
        selectFromArray(information)
    elif instruction == "selectFromArrayRegister":
        selectFromArrayRegister(information)
    elif instruction == "sendEnter":
        sendEnter()
    elif instruction == "sendClick":
        sendClick()
    elif instruction == "selectValue":
        selectValue(information)
    elif instruction == "setText":
        setText(information)
    elif instruction == "getText":
        getText()
    elif instruction == "getAttribute":
        getAttribute(information)
    elif instruction == "forFun":
        forFun(information)
    elif instruction == "forEach":
        forEach(information)
    elif instruction == "breakF":
        breakF(information)
    elif instruction == "ifElemExists":
        ifElemExists(information)
    elif instruction == "dateCompare":
        dateCompare(information)
    elif instruction == "ifIsInArrayVAR":
        ifIsInArrayVAR(information)
    elif instruction == "getFiles":
        getFiles(information)
    elif instruction == "ifLesser":
        ifLesser(information)
    elif instruction == "dateFormatConversion":
        dateFormatConversion(information)
    elif instruction == "ifEqual":
        ifEqual(information)
    elif instruction == "readPDF":
        readPDF(information)
    elif instruction == "getFileName":
        getFileName(information)
    elif instruction == "getFullFileName":
        getFullFileName(information)
    elif instruction == "extractText":
        extractText(information)
    elif instruction == "replaceText":
        replaceText(information)
    #elif instruction == "splitText":
    #    splitText(information)
    elif instruction == "substr":
        substr(information)
    elif instruction == "createArrayFromString":
        createArrayFromString(information)
    elif instruction == "mergeExcel":
        mergeExcel(information)
    elif instruction == "appendExcel":
        appendExcel(information)
    elif instruction == "saveToFile":
        saveToFile(information)
    elif instruction == "loadFromFile":
        loadFromFile(information)
    elif instruction == "setRegister":
        setRegister(information)
    elif instruction == "getLength":
        getLength()
    elif instruction == "getCurrDate":
        getCurrDate(information)
    elif instruction == "addDays":
        addDays(information)
    elif instruction == "subDays":
        subDays(information)
    elif instruction == "createFolder":
        createFolder(information)
    elif instruction == "setVar":
        setVar(information)
    elif instruction == "unsetVar":
        unsetVar(information)
    elif instruction == "getVar":
        getVar(information)
    elif instruction == "getVars":
        getVars()
    elif instruction == "numToMonth":
        numToMonth(information)
    elif instruction == "getDaysInMounth":
        getDaysInMounth(information)
    elif instruction == "renameFile":
        renameFile(information)
    elif instruction == "removeFile":
        removeFile(information)
    elif instruction == "storX":
        storX()
    elif instruction == "zpaddingLeft":
        zpaddingLeft(information)
    elif instruction == "mouseClickPosition":
        mouseClickPosition(information)
    elif instruction == "clearElem":
        clearElem(information)
    elif instruction == "createJson":
        createJson()
    elif instruction == "sendPOST":
        sendPOST(information)
    elif instruction == "storY":
        storY()
    elif instruction == "storZ":
        storZ()
    elif instruction == "storW":
        storW()
    elif instruction == "lodX":
        lodX()
    elif instruction == "lodY":
        lodY()
    elif instruction == "lodZ":
        lodZ()
    elif instruction == "lodW":
        lodW()
    elif instruction == "addTextToRegistry":
        addTextToRegistry(information)
    elif instruction == "addVarFECycleToRegistry":
        addVarFECycleToRegistry(information)
    elif instruction == "convertToExcel":
        convertToExcel(information)
    elif instruction == "getRowCount":
        getRowCount(information)
    elif instruction == "openExcel":
        openExcel(information)
    elif instruction == "readExcelCell":
        readExcelCell(information)
    elif instruction == "newExcel":
        newExcel(information)
    elif instruction == "filterExcelByColumnVal":
        filterExcelByColumnVal(information)
    elif instruction == "moveBetweenWorkspaces":
        moveBetweenWorkspaces(information)
    elif instruction == "image2Text":
        image2Text(information)
    elif instruction == "token2text":
        token2text(information)
    elif instruction == "addValToT":
        addValToT(information)
    elif instruction == "subValToT":
        subValToT(information)
    elif instruction == "addX":
        addX()
    elif instruction == "addValToX":
        addValToX(information)
    elif instruction == "addY":
        addY()
    elif instruction == "addValToY":
        addValToY(information)
    elif instruction == "addZ":
        addZ()
    elif instruction == "addValToZ":
        addValToZ(information)
    elif instruction == "addW":
        addW()
    elif instruction == "addValToW":
        addValToW(information)
    elif instruction == "clearT":
        clearT(information)
    elif instruction == "clearX":
        clearX(information)
    elif instruction == "clearY":
        clearY(information)
    elif instruction == "clearZ":
        clearZ(information)
    elif instruction == "clearW":
        clearW(information)
    elif instruction == "clearAll":
        clearAll()
    elif instruction == "printT":
        printT()
    elif instruction == "printX":
        printX()
    elif instruction == "printY":
        printY()
    elif instruction == "printZ":
        printZ()
    elif instruction == "printW":
        printW()

if len(sys.argv) <= 2:
    input("Premi Invio per terminare")

# INSTRUCTIONS REFERENCE

#initFirefox |--| 
#nav |--| urlExample
#selectId |--| idExample
#selectClass |--| classExample <-||-> -1 or other
#selectName |--| nameExample
#selectChild |--| childNum
#sendEnter |--| 
#input |--| inputExample
#forEachElem |--| instruction 1 <-||-> instruction 2 <-||-> ... <-||-> ...
#getText |--| 
#readPDF |--| pathToPDF
#extractTextFromT |--| previousCharsAsString <-||-> charsAfterTextAsString
#replaceTextT |--| oldText <-||-> newText
#closeIfNullT |--| 


# INSTALLATION REQUIREMENTS

# pip install -U selenium
# pip install PyPDF2
