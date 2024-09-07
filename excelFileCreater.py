from openpyxl import Workbook
import random, string
import os

def get_valid_char(forbiddenChars, sequence):
    is_valid_char = False

    if sequence == "file_sequence":
        while not is_valid_char:
            commonFileName = input("File name: ")

            if any(char in forbiddenChars for char in commonFileName):
                forbidden_in_name = [char for char in commonFileName if char in forbiddenChars]
                print(f"Forbidden characters used: {', '.join(forbidden_in_name)}")
                print("Choose another file name.")
            else: is_valid_char = True
        
        return commonFileName
    elif sequence == "separator_sequence":
        while not is_valid_char:
            addSeparator = input("Enter separator: ")

            if any(char in forbiddenChars for char in addSeparator):
                forbidden_in_separator = [char for char in addSeparator if char in forbiddenChars]
                print(f"Forbidden characters used: {', '.join(forbidden_in_separator)}")
                print("Choose another separator.")
            else: is_valid_char = True

        return addSeparator
    else:
        return

wb = Workbook()
ws = wb.active

defaultFileLocation = (os.getcwd())
definedFileLocation = input(f"Default file placement: {defaultFileLocation} \nPlace elsewhere? [y/n]: ")

if definedFileLocation in ("yes", "y"):
    fileLocation = input("Enter location: ")
else: fileLocation = defaultFileLocation

fileAmount = int(input("Number of files: "))
isCustomFileName = input("Custom File Name? [y/n]: ").lower()

forbiddenChars = ["/", "\\", ".", ">", "<", "*", "?", "!", "|", " ", '"', ":" "\\0"]

if isCustomFileName in ("yes", "y"):
    commonFileName = get_valid_char(forbiddenChars, "file_sequence")
else:
    letters = string.ascii_lowercase
    commonFileName = ''.join(random.choice(letters) for i in range(10))
    fileName = print(f"\nRandomized file name: {commonFileName}")

print(f"\n{fileAmount} file(s) wanted.")

if fileAmount >= 1:
    startingAmount = input("Predefined starting iteration? [y/n]: ").lower()
    if startingAmount in ("yes", "y"):
        iteration = int(input("Starting iteration: "))
    else:
        iteration = 1

toAddSeparator = input("Add separator between name and iteration? [y/n]: ").lower()

if toAddSeparator in ("yes", "y"):
    addSeparator = get_valid_char(forbiddenChars, "separator_sequence")
else:
    addSeparator = ""

if commonFileName:
    createFileNames = [f"{commonFileName}{addSeparator}{i}.xlsx" for i in range(iteration, iteration + fileAmount)]
    print(f"Files to be created: {createFileNames}")
    print(f"In directory: {fileLocation}")

confirmation = input("Is this correct? [y/n]: ").lower()

if confirmation in ("yes", "y"):
    for file in createFileNames:
        wb.save(rf"{fileLocation}\{file}")
    print("Excel file(s) created successfully!")
else:
    print("Program cancelled")
