#!/usr/bin/env python3
import openpyxl, os, time
from colorama import Fore, Style

global counter
global first_name
global data

counter = 2
first_name = None

# Spreadsheet set up
signin = openpyxl.Workbook()
signin_sheet = signin.get_active_sheet()
signin_sheet.title = "Theta Tau Date Night Sign Ins"
signin_sheet["A1"] = "First Name"
signin_sheet["B1"] = "Last Name"
signin_sheet["C1"] = "Email"
signin_sheet["D1"] = "Major"
signin_sheet["E1"] = "Expected Graduation Date"

with open("logo.txt", "r") as myfile:
    data=myfile.read()
    #print(data)

def printHeader():
    os.system('clear')
    #logo = "  _____ _          _          _____           \n |_   _| |        | |        |_   _|          \n   | | | |__   ___| |_ __ _    | | __ _ _   _ \n   | | | '_ \ / _ \ __/ _` |   | |/ _` | | | |\n   | | | | | |  __/ || (_| |   | | (_| | |_| |\n   \_/ |_| |_|\___|\__\__,_|   \_/\__,_|\__,_|"
    logo = data
    #print("\n")
    print(Fore.RED + "\n".join('{:^170}'.format(s) for s in logo.split("\n")))
    #print((Fore.BLUE + "/------------------------------------------------------------------------------------------------\\").center(525))
    #print((Fore.RED + "|          Welome to date night! The brothers of Theta Tau are excited for you to be here          |").center(525))
    #print((Fore.BLUE + "\\------------------------------------------------------------------------------------------------/").center(525))
    print(Style.RESET_ALL)
def userInput():
    global counter
    global first_name

    first_name = input("(1/5). What's your first name? \n")
    if first_name != "exit":
        last_name = input("\n(2/5). What's your last name? \n")
        email = input("\n(3/5). What's your email? \n")
        major = input("\n(4/5). What's your major? \n")
        grad = input("\n(5/5. What's your expected graduate date? \n")
        signin_sheet['A' + str(counter)] = str(first_name).upper()
        signin_sheet['B' + str(counter)] = str(last_name).upper()
        signin_sheet['C' + str(counter)] = str(email).upper()
        signin_sheet['D' + str(counter)] = str(major).upper()
        signin_sheet['E' + str(counter)] = str(grad).upper()
        counter += 1

def finishScreen():
    os.system('clear')
    print("\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n")
    print((Fore.BLUE + "     /------------------------------------------------------------------------------------------------\\").center(525))
    print(("|                                Enjoy the rest of your evening!                                  |").center(525))
    print(("\\------------------------------------------------------------------------------------------------/").center(525))
    print(Style.RESET_ALL)
while first_name != "exit":
    printHeader()
    userInput()
    finishScreen()
    os.system('echo "\a\a\a\a"')
    time.sleep(2)

signin.save("Orientation-8-4-16.xlsx")
print("Sheet saved")
