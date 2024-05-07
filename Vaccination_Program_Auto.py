# Import statements
import webbrowser
import os

# Utility functions

# Initialize/ Assign base working / Data directory

# Initialize the constant values representing positions on the screen

# Initialize the mouse and keyboard controllers

# Initialize the excel and json files that are to be worked with


#


# MAIN DRIVER
# Initialize the program loop

# Create procedure to open the website
url = "https://bharatpashudhan.ndlm.co.in/auth/login"
chrome_path = "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s"
webbrowser.get(chrome_path).open(url)


# Create the login procedure
#    Handle the proper failure and logout cases

# Close the browser
os.system("taskkill /f /im chrome.exe")

# Initialize the data loop
