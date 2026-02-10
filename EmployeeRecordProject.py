# Project : Client Intake Program
# Student Name : Ichty Te
# Student Number: 200626964
# Date         : Dec 08, 2025
# Instructor   : Wayne Brown
# Course       : COMP1112
# Description  : Program that collects client records, validates them, and saves to Excel.

#import modules
import time
import openpyxl
import os
import re

# This function fixes the client's name
# remove extra spaces and make sure each word starts with a capital letter
def formatClientName(clientName):
    return clientName.strip().title()

#This function checks if the email looks valid using regex
#and that the domain ends with a proper extension
def isValidEmail(clientEmail):
    pattern = r"^[\w\.-]+@[\w\.-]+\.\w+$"
    #if it is match the pattern then return True
    if re.match(pattern, clientEmail):
        return True
    #if not return False
    else:
        return False

#this function checks the phone number
def phoneNumber(clientPhone):
    #replace anything that is not a digit with "" using regex
    #just in case client use - or ()
    phoneDigits = re.sub(r"\D", "", clientPhone)
    # make sure if there are exactly 10 to be valid phone number
    if len(phoneDigits) == 10:
        # first 3 digits are the area code
        areaCode = phoneDigits[:3]
        #also return area code, because it useful for client record
        return phoneDigits, areaCode
    else:
        # if not valid, return None
        return None, None


#this function checks if the postal code is valid
#use regex to match the Canadian format: Letter-Digit-Letter Digit-Letter-Digit
def formatPostalCode(postalCode):
    postalCode = postalCode.strip().upper()
    # Remove spaces so we can check the pattern easily
    postalCodeNoSpace = postalCode.replace(" ", "")
    pattern = r"^[A-Z]\d[A-Z]\d[A-Z]\d$"
    if re.match(pattern, postalCodeNoSpace):
        # If valid, return it in the standard format with a space in the middle
        return postalCodeNoSpace[:3] + " " + postalCodeNoSpace[3:]
    else:
        #if not return None
        return None

#This function writes down any problems into a text file
def logIssue(issueMessage):
    #use try if file is error
    try:
        #add the time,can know when the error happened
        with open("issues.txt", "a") as issueFile:
            issueFile.write(time.ctime() + ": " + issueMessage + "\n")
    #in case the file is error
    except Exception as e:
        print("Error", "logging issue:", e)

#this function checks one client record
#it validates name, email, phone, and postal code
#if something is wrong, the problem will go to logIssue
def processClientRecord(clientName, clientEmail, clientPhone, postalCode):
    #clear the name , no strip, and have capital letter
    formattedName = formatClientName(clientName)

    #isValindEmail(client Email)
    if not isValidEmail(clientEmail):
        # If the email is not valid, isValidEmail returns False
        # log the issue and skip this record by returning None
        logIssue("Invalid email: " + clientEmail)
        return None

    #in PhoneNumber function, return both None, or valid
    phoneDigits, areaCode = phoneNumber(clientPhone)
    #so if phone number is not valid, it will log the issue and write it in issue.txt
    if phoneDigits is None:
        logIssue("Invalid phone: " + clientPhone)
        return None
    #if postalcode invalid, put in Issue.txt
    formattedPostalCode = formatPostalCode(postalCode)
    if formattedPostalCode is None:
        logIssue("Invalid postal: " + postalCode)
        return None

    # If all is valid, return record as a list
    return [formattedName, clientEmail, phoneDigits, formattedPostalCode, areaCode]

# This function saves all valid client records into an Excel file
def saveToExcel(validClientRecords):
    # Use try/except to defend the program from crashing because file cant open or save
    try:
        filePath = "clients_clean.xlsx"
        #if Excel file exists, open it and use the active worksheet
        if os.path.exists(filePath):
            workbook = openpyxl.load_workbook(filePath)
            worksheet = workbook.active
        else:
            #But if the file not exist, create a new
            workbook = openpyxl.Workbook()
            worksheet = workbook.active
            worksheet.title = "Clients"
            # Add a header
            worksheet.append(["Name", "Email", "Phone", "Postal", "AreaCode"])

        # Loop valid client record and then add it as a new row in the worksheet
        for clientRecord in validClientRecords:
            worksheet.append(clientRecord)

        # Save the workbook to the file path
        workbook.save(filePath)
    except Exception as e:
        # If something goes wrong, log the issue and also show the error
        print("Error to saving Excel file: " + str(e))

# This is the main program
# It asks the user to enter client information until they press Enter with no name
# Each record is validated, and only valid ones are saved to Excel
def main():
    #create empty to hold all valid client record
    validClientRecords = []

    print("Client Intake Program")
    print("Enter client records. Leave Name empty to stop.\n")
    #ask for clientName, and remove any extra spaces
    clientName = input("Enter Name (or press Enter to finish): ").strip()
    #use while loop for keep running until user enters a name
    #if the client presses Enter without name, the loop will end
    while clientName != "":
        #ask user for their info, and remove extra space
        clientEmail = input("Enter Email: ").strip()
        clientPhone = input("Enter Phone: ").strip()
        postalCode = input("Enter Postal Code: ").strip()

        #call the process client record function for checking if it is valid or not
        clientRecord = processClientRecord(clientName, clientEmail, clientPhone, postalCode)
        # if it is valid then added to the list
        # if not , it skipped and going to issue.txt
        if clientRecord is not None:
            validClientRecords.append(clientRecord)
        #repeat until client press Enter
        clientName = input("Enter Name (or press Enter to finish): ").strip()
    #if there are valid record then written to excel
    if len(validClientRecords) > 0:
        saveToExcel(validClientRecords)
        print("the record has been saved successfully")
    #if not , warn client than the issue is in issues.txt
    else:
        print("No valid records to save. Check issues.txt for details.")

# call the main function to run the program
main()
