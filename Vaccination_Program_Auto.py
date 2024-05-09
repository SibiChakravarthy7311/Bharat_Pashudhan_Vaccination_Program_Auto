# Import statements
import webbrowser
import pynput.keyboard
from tkinter import Tk
from pynput.mouse import Button
from pynput.keyboard import Key
from urllib import request
import keyboard
import time
import os

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
    # firefox_path = "C:/Program Files/Mozilla Firefox/firefox.exe %s"
    # browser = webbrowser.get(firefox_path)

    url_short = "bharatpashudhan.ndlm.co.in/auth/login"
    # chrome_path = "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s"
    # webbrowser.get(chrome_path).open_new_tab(url_short)

    webbrowser.open(url)
    print("Opened browser")
    time.sleep(5)


# TO BE MODIFIED. THIS IS JUST A TEMPORARY CHECKCAPS FUNCTION TO TERMINATE THE PROGRAM. MORE DETAILS WILL BE
# ADDED TO THIS FUNCTION LATER
def checkCaps(i):
    if GetKeyState(VK_CAPITAL) == 1:
        print("CapsLock on.Procedure Aborted")
        closeWebApp()
        exit()


# Close the browser along with the webapp
def closeWebApp():
    os.system("taskkill /f /im chrome.exe")


# Close the browser along with the webapp and reopen it
def rebootApplication():
    # Close the browser
    closeWebApp()

    # Open the webapp
    openWebApp()


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


# Retrieve the value entered in an input field at a given position
def getEnteredValueAt(mouse, keyboard, POSITION):
    mouse.position = POSITION
    time.sleep(1)
    mouse.click(Button.left)
    mouse.click(Button.left)
    with keyboard.pressed(Key.ctrl.value):
        keyboard.press('c')
        keyboard.release('c')

    data = Tk().clipboard_get()
    return data


# Initialize/ Assign base working / Data directory

# Initialize the constant values representing positions on the screen
USERNAME_POSITION = (550, 474)
PASSWORD_POSITION = (550, 586)
LOGIN_POSITION = (770, 692)

ANIMAL_HEALTH_POSITION = (75, 457)
VACCINATION_POSITION = (75, 498)
WITHOUT_CAMPAIGN_POSITION = (371, 298)
BATCH_NUMBER_POSITION = (290, 505)

# Initialize the mouse and keyboard controllers
mouse = pynput.mouse.Controller()
copyKeyboard = pynput.keyboard.Controller()

# Initialize the excel and json files that are to be worked with
# This is just sample data to temporarily work with
dataIndex = 0
data = [["430104580967", "06-11-2023"], ["430104598092", "06-11-2023"], ["430104600923", "06-11-2023"],
        ["102879683382", "06-11-2023"], ["430104601777", "06-11-2023"]]

#


# MAIN DRIVER
# Open the web application
openWebApp()

# Initialize the program loop
loginAttemptCounter = 5
while dataIndex < len(data):
    print("Reached Here")
    if not isConnected():
        time.sleep(30)
        continue
    currentData = data[dataIndex]
    # print(currentData)
    if currentData[0] == '~':
        dataIndex += 1
        continue
    startRow = 1
    username = "tnnadcptpr_svls1udt_TN"
    password = "Tirupur#1"
    # print(currentData)

    # Create the login procedure
    if not loginAttemptCounter:
        time.sleep(120)
        rebootApplication()
        loginAttemptCounter = 5
        continue

    # Calling the login function
    login(username, password, USERNAME_POSITION, PASSWORD_POSITION, LOGIN_POSITION, ANIMAL_HEALTH_POSITION,
          VACCINATION_POSITION, WITHOUT_CAMPAIGN_POSITION)

    # Handle the proper failure and logout cases
    userSub = username[-11:]
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
    df = data
    totalRows = len(df)
    flag = True
    currentRow = totalRows

    for i in range(startRow, totalRows):
        if not isConnected():
            currentRow = i
            flag = False
            break
        checkCaps(i)


# Initialize the data loop
