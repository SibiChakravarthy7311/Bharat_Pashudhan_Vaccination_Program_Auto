# Import statements
import webbrowser
import pynput.keyboard
from tkinter import Tk
from PIL import Image
from pynput.mouse import Button

from pynput.keyboard import Key
from urllib import request
import datetime
import keyboard

import time
import os
import json
import pandas as pd
from openpyxl import load_workbook

from win32api import GetKeyState
from win32con import VK_CAPITAL


# UTILITY FUNCTIONS
# Check whether an active internet connection is established between the system and the
# Bharat Pashudhan website
def isConnected(host='https://bharatpashudhan.ndlm.co.in/auth/login'):
    try:
        request.urlopen(host)
        return True
    except:
        return False


# Open the Bharat Pashudhan web application
def openWebApp():
    url = "https://bharatpashudhan.ndlm.co.in/auth/login"
    webbrowser.open(url)
    print("Opened browser")
    time.sleep(10)

    # Close the restore window popup incase it appears on the screen
    closeRestoreWindowPopup(CHROME_RESTORE_CLOSE_POSITION)


# To close the restore window popup in chrome
def closeRestoreWindowPopup(CHROME_RESTORE_CLOSE_POSITION):
    mouse.position = CHROME_RESTORE_CLOSE_POSITION
    mousePressAndRelease(0.25)
    time.sleep(0.5)


# Updates the next file to be worked on in the nextFile.json file path
def updateNextFile(fileIndex, nextFileData, jsonFilePath):
    nextFileData['nextFile'] = fileIndex
    dumpData = json.dumps(nextFileData)
    with open(jsonFilePath, 'w') as file:
        file.write(dumpData)
        file.close()


# Checkpoint function that checks if the Caps Lock is active, saves the current working status of files and
# exits the program
def checkCaps(i, sheet, workbook, data, fileIndex, nextFileIndex, jsonFilePath):
    if GetKeyState(VK_CAPITAL) == 1:
        print("CapsLock on.Procedure Aborted")
        sheet['E2'] = i
        workbook.save(filename=data)
        updateNextFile(fileIndex, nextFileIndex, jsonFilePath)
        exitFunction()


# Close the browser along with the webapp
def closeWebApp():
    os.system("taskkill /f /im chrome.exe")


# Close the browser along with the webapp and reopen it
def rebootApplication():
    # Close the browser
    closeWebApp()

    # Open the webapp
    openWebApp()


# Fill Vaccination For field
def fillVaccinationFor(VACCINATION_FOR_POSITION, VACCINATION_FOR):
    mouse.position = VACCINATION_FOR_POSITION
    mousePressAndRelease(0.25)
    keyboard.write(VACCINATION_FOR)
    time.sleep(1)
    keyboard.send("enter")
    time.sleep(3)


# Fill the Vaccine Type Field
def fillVaccineName(VACCINE_NAME_POSITION, VACCINE_NAME):
    mouse.position = VACCINE_NAME_POSITION
    mousePressAndRelease(0.25)
    keyboard.write(VACCINE_NAME)
    time.sleep(1)
    keyboard.send("enter")
    time.sleep(3)


# Fill the Batch Number
def fillBatchNumber(BATCH_NUMBER_POSITION, BATCH_NUMBER):
    mouse.position = BATCH_NUMBER_POSITION
    mousePressAndRelease(0.25)
    keyboard.write(BATCH_NUMBER)
    time.sleep(1)


# Press and release the left mouse button
def mousePressAndRelease(t=0):
    mouse.press(Button.left)
    time.sleep(t)
    mouse.release(Button.left)


# Select the option at a given position with assigned mouse press and sleep time
def selectOption(POSITION, pressTime=0.25, sleepTime=1):
    mouse.position = POSITION
    mousePressAndRelease(pressTime)
    time.sleep(sleepTime)


# Login procedure using Bharat Pashudhan Username and Password
def login(username, password, USERNAME_POSITION, PASSWORD_POSITION, LOGIN_POSITION, ANIMAL_HEALTH_POSITION, VACCINATION_POSITION, WITHOUT_CAMPAIGN_POSITION):
    mouse.position = USERNAME_POSITION
    mousePressAndRelease(0.25)
    mouse.click(Button.left)
    keyboard.write(username)
    time.sleep(1)

    mouse.position = PASSWORD_POSITION
    mousePressAndRelease(0.25)
    mouse.click(Button.left)
    keyboard.write(password)
    time.sleep(1)

    selectOption(LOGIN_POSITION, 0.25, 10)
    selectOption(ANIMAL_HEALTH_POSITION, 0.25, 2)
    selectOption(VACCINATION_POSITION, 0.25, 2)
    selectOption(WITHOUT_CAMPAIGN_POSITION, 0.25, 2)


# Logout Function
def logout(PROFILE_POSITION, LOGOUT_POSITION):
    mouse.position = PROFILE_POSITION
    time.sleep(0.25)
    mousePressAndRelease(0.25)
    time.sleep(0.5)

    mouse.position = LOGOUT_POSITION
    time.sleep(0.25)
    mousePressAndRelease(0.25)
    time.sleep(2)


# Retrieve the value entered in an input field at a given position
def getEnteredValueAt(mouse, keyboard, POSITION):
    mouse.position = POSITION
    time.sleep(1)
    mouse.click(Button.left)
    mouse.click(Button.left)
    time.sleep(0.5)
    with keyboard.pressed(Key.ctrl.value):
        keyboard.press('c')
        keyboard.release('c')

    data = Tk().clipboard_get()
    time.sleep(0.25)
    return data


# To enter the tag and search
def setTagAndSearch(TAG_NUMBER_POSITION, TAG_SEARCH_POSITION, tagId):
    mouse.position = TAG_NUMBER_POSITION
    mouse.press(Button.left)
    time.sleep(0.25)
    mouse.release(Button.left)
    time.sleep(0.5)
    keyboard.write(tagId)
    time.sleep(0.5)

    mouse.position = TAG_SEARCH_POSITION
    mouse.press(Button.left)
    time.sleep(0.25)
    mouse.release(Button.left)
    time.sleep(3)


# To select the tag result and proceed to the vaccination page
def proceedVaccination(TAG_SELECT_POSITION, VACCINATION_PROCEED_POSITION):
    mouse.position = TAG_SELECT_POSITION
    time.sleep(0.25)
    mousePressAndRelease(0.25)
    time.sleep(0.5)

    mouse.position = VACCINATION_PROCEED_POSITION
    time.sleep(0.25)
    mousePressAndRelease(0.25)
    time.sleep(3)


# Set the vaccination date
def setVaccinationDate(VACCINATION_DATE_DAY_POSITION, VACCINATION_DATE_MONTH_POSITION, VACCINATION_DATE_YEAR_POSITION,
                       vDate, vMonth, vYear):
    mouse.position = VACCINATION_DATE_DAY_POSITION
    time.sleep(0.25)
    mousePressAndRelease(0.25)
    mouse.click(Button.left)
    keyboard.write(vDate)

    mouse.position = VACCINATION_DATE_MONTH_POSITION
    time.sleep(0.25)
    mousePressAndRelease(0.25)
    mouse.click(Button.left)
    keyboard.write(vMonth)

    mouse.position = VACCINATION_DATE_YEAR_POSITION
    time.sleep(0.25)
    mousePressAndRelease(0.25)
    mouse.click(Button.left)
    keyboard.write(vYear)
    time.sleep(0.5)


# Set the vaccination type
def setVaccinationType(VACCINATION_TYPE_POSITION, VACCINATION_REGULAR_POSITION):
    mouse.position = VACCINATION_TYPE_POSITION
    time.sleep(0.25)
    mousePressAndRelease(0.25)
    time.sleep(0.25)

    mouse.position = VACCINATION_REGULAR_POSITION
    time.sleep(0.25)
    mousePressAndRelease(0.25)
    time.sleep(0.5)


# Submit the vaccination Transaction
def submitVaccination(VACCINATION_SUBMIT_POSITION, VACCINATION_DONE_POSITION, VACCINATION_ERROR_POSITION,
                      VACCINATION_CANCEL_POSITION):
    mouse.position = VACCINATION_SUBMIT_POSITION
    time.sleep(0.25)
    mousePressAndRelease(0.25)
    time.sleep(5)

    mouse.position = VACCINATION_DONE_POSITION
    time.sleep(0.25)
    mousePressAndRelease(0.25)
    time.sleep(0.5)

    mouse.position = VACCINATION_ERROR_POSITION
    time.sleep(0.25)
    mousePressAndRelease(0.25)
    time.sleep(0.5)

    mouse.position = VACCINATION_CANCEL_POSITION
    time.sleep(0.25)
    mousePressAndRelease(0.25)
    time.sleep(0.5)


# To clear any previously filled fields in the vaccination page
def resetWindow(CAMPAIGN_POSITION, WITHOUT_CAMPAIGN_POSITION):
    # Switching to Campaign page and returning back to the Without Campaign page seems to do the trick here
    selectOption(CAMPAIGN_POSITION, pressTime=0.25, sleepTime=3)
    selectOption(WITHOUT_CAMPAIGN_POSITION, pressTime=0.25, sleepTime=3)


# Completely quit the program execution
def exitFunction():
    interrupted = Image.open('interrupted.png')
    interrupted.show()
    exit()


# [1] - The sections to be implemented after the base function is completed are marked with [1]

# [2] - The stubs for current working data which are later to be replaced by [1] are marked by [2]

# Initialize/ Assign base working / Data directory
myDir = "D:/EntryData/test/"

# Initialize the constant values representing positions on the screen
CHROME_RESTORE_CLOSE_POSITION = (1507, 110)

USERNAME_POSITION = (545, 490)
PASSWORD_POSITION = (545, 582)
LOGIN_POSITION = (770, 688)
PROFILE_POSITION = (1435, 115)
LOGOUT_POSITION = (1413, 267)

ANIMAL_HEALTH_POSITION = (75, 430)
VACCINATION_POSITION = (75, 470)
CAMPAIGN_POSITION = (273, 274)
WITHOUT_CAMPAIGN_POSITION = (370, 274)

BATCH_NUMBER_POSITION = (290, 483)
VACCINATION_FOR_POSITION = (390, 400)
VACCINE_NAME_POSITION = (690, 400)
TAG_NUMBER_POSITION = (320, 625)
TAG_SEARCH_POSITION = (1440, 625)

TAG_INVALID_POSITION = (915, 510)
TAG_INVALID_TEXT_POSITION = (675, 458)
TAG_SELECT_POSITION = (278, 606)
VACCINATION_PROCEED_POSITION = (1445, 775)

VACCINATION_DATE_DAY_POSITION = (578, 342)
VACCINATION_DATE_MONTH_POSITION = (600, 342)
VACCINATION_DATE_YEAR_POSITION = (629, 342)
VACCINATION_TYPE_POSITION = (1135, 543)

VACCINATION_REGULAR_POSITION = (1135, 585)
VACCINATION_SUBMIT_POSITION = (1326, 717)
VACCINATION_DONE_POSITION = (768, 540)
VACCINATION_ERROR_POSITION = (917, 523)
VACCINATION_CANCEL_POSITION = (1250, 717)

# For this vaccination cycle the VACCINATION_FOR, VACCINE_TYPE and BATCH NUMBER values are constant as below
VACCINATION_FOR = "Foot and Mouth Disease(FMD)"
VACCINE_NAME = "Raksha-Ovac"
BATCH_NUMBER = "01FUT0023"

# Initialize the mouse and keyboard controllers
mouse = pynput.mouse.Controller()
copyKeyboard = pynput.keyboard.Controller()

# Initialize the excel and json files that are to be worked with
# [1]
jsonFilePath = myDir + "nextFile.json"
jsonDataFile = open(jsonFilePath)
x = jsonDataFile.read()
jsonDataFile.close()
nextFileData = json.loads(x)
fileIndex = nextFileData['nextFile']

fileList = os.listdir(myDir)
fileList.remove('nextFile.json')

# MAIN DRIVER
# Open the web application for the first time the program starts running
openWebApp()

# Initialize the program loop
loginAttemptCounter = 5
while fileIndex < len(fileList):
    if not isConnected():
        time.sleep(30)
        continue

    currentFile = fileList[fileIndex]
    # print(currentFile)
    if currentFile[0] == '~':
        fileIndex += 1
        continue
    data = os.path.join(myDir, currentFile)
    workbook = load_workbook(data)
    sheet = workbook.active
    startRow = sheet['E2'].internal_value
    userName = "_".join(currentFile[:-5].split("_")[1:])
    password = "Tirupur#1"

    print(data, userName)

    # print('LoginAttemptCounterStart: {}'.format(loginAttemptCounter))
    # Create the login procedure
    if not loginAttemptCounter:
        time.sleep(120)
        rebootApplication()
        loginAttemptCounter = 5
        continue

    # Calling the login function
    login(userName, password, USERNAME_POSITION, PASSWORD_POSITION, LOGIN_POSITION, ANIMAL_HEALTH_POSITION,
          VACCINATION_POSITION, WITHOUT_CAMPAIGN_POSITION)

    # Check for login failure and Handle the failure and logout cases
    userSub = userName[-11:]
    mouse.position = BATCH_NUMBER_POSITION
    mousePressAndRelease(0.25)
    keyboard.write(userSub)
    time.sleep(1)

    try:
        currentUserName = getEnteredValueAt(mouse, copyKeyboard, BATCH_NUMBER_POSITION)
    except:
        currentUserName = "FalseValue"
    if currentUserName != userSub:
        rebootApplication()
        loginAttemptCounter -= 1
        continue
    else:
        print("Login Successful")

    # Proceed with the program execution after successful login is confirmed
    df = pd.read_excel(data)
    totalRows = len(df)
    # print("Total Rows: {}".format(totalRows))
    flag = True
    currentRow = totalRows

    for i in range(startRow, totalRows):
        if not isConnected():
            currentRow = i
            flag = False
            break

        # Checkpoint by checking the status of Caps Lock key
        checkCaps(i, sheet, workbook, data, fileIndex, nextFileData, jsonFilePath)

        # Reset the batch number field which was previously used for verifying the login status
        resetWindow(CAMPAIGN_POSITION, WITHOUT_CAMPAIGN_POSITION)

        # Fetch the entry data parameters
        tagId = str(df['Tag'][i])
        vDate = str(df['V_Date'][i]).zfill(2)
        vMonth = str(df['V_Month'][i]).zfill(2)
        vYear = str(df['V_Year'][i])

        # Logging code
        dateTimeNow = datetime.datetime.now()
        Today = dateTimeNow.day
        toMonth = dateTimeNow.month

        time_now = dateTimeNow.strftime("   %Y-%m-%d %H:%M:%S")
        logStr = "{} of {} {} {} \n".format(i + 1, totalRows, tagId, time_now)
        print(logStr, end='')

        # Unlike INAPH, the values of Vaccination For, Vaccination name, Vaccine Type, Vaccine subtype, Batch Number -
        # all these details have to be filled prior to searching the animal Tag
        # Fill the Vaccination For
        fillVaccinationFor(VACCINATION_FOR_POSITION, VACCINATION_FOR)

        # Fill the Vaccine Name; By filling the Vaccine Name, Vaccine Type and Vaccine subtype are automatically filled
        fillVaccineName(VACCINE_NAME_POSITION, VACCINE_NAME)

        # Fill the Batch Number
        fillBatchNumber(BATCH_NUMBER_POSITION, BATCH_NUMBER)

        # Code section to enter and search for the tag
        setTagAndSearch(TAG_NUMBER_POSITION, TAG_SEARCH_POSITION, tagId)

        # Check if the tag was invalid
        invalidText = getEnteredValueAt(mouse, copyKeyboard, TAG_INVALID_TEXT_POSITION)
        if invalidText == "Invalid":
            mouse.position = TAG_INVALID_POSITION
            time.sleep(0.25)
            mousePressAndRelease(0.25)
            time.sleep(0.5)
            continue

        # Check if the tag search was successful
        enteredTagId = getEnteredValueAt(mouse, copyKeyboard, TAG_NUMBER_POSITION)

        # If the tag search wasn't successful due to connectivity / some other error and the desired page is not reached
        if enteredTagId != tagId:
            print("Connectivity Error, Retrying Procedure")
            sheet['E2'] = i
            workbook.save(filename=data)
            updateNextFile(fileIndex, nextFileData, jsonFilePath)
            loginAttemptCounter -= 1
            print('LoginAttemptCounter: {}'.format(loginAttemptCounter))
            rebootApplication()
            flag = False
            currentRow = i
            break

        # Reset the login Attempt counter back to 5
        loginAttemptCounter = 5

        # Scroll down if the tag search is successful
        mouse.scroll(0, -6)
        time.sleep(0.5)

        # Checkpoint by checking the status of Caps Lock key
        checkCaps(i, sheet, workbook, data, fileIndex, nextFileData, jsonFilePath)

        # Select the tag row and proceed to make the vaccination entry
        proceedVaccination(TAG_SELECT_POSITION, VACCINATION_PROCEED_POSITION)

        # Set the vaccination date
        setVaccinationDate(VACCINATION_DATE_DAY_POSITION, VACCINATION_DATE_MONTH_POSITION,
                           VACCINATION_DATE_YEAR_POSITION, vDate, vMonth, vYear)

        # Set the vaccination type
        setVaccinationType(VACCINATION_TYPE_POSITION, VACCINATION_REGULAR_POSITION)

        # Save the vaccination transaction
        submitVaccination(VACCINATION_SUBMIT_POSITION, VACCINATION_DONE_POSITION,
                          VACCINATION_ERROR_POSITION, VACCINATION_CANCEL_POSITION)

        # Checkpoint by checking the status of Caps Lock key
        checkCaps(i, sheet, workbook, data, fileIndex, nextFileData, jsonFilePath)

        # Scroll back up after the vaccination is complete
        mouse.scroll(0, 6)

    # Record the changes to the working Excel file
    if flag:
        sheet['E2'] = totalRows
        workbook.save(filename=data)

        # print("Credentials are valid")
        logout(PROFILE_POSITION, LOGOUT_POSITION)
        fileIndex += 1
    else:
        sheet['E2'] = currentRow
        workbook.save(filename=data)

# Update the json file containing the index of the next file to be worked on
updateNextFile(len(fileList), nextFileData, jsonFilePath)

# End the program by displaying the success message
success = Image.open('success.png')
success.show()

