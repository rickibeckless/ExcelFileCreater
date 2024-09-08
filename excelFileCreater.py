from openpyxl import Workbook
import random, string
import os

def check_location(forbiddenDirChars): 
    is_valid_path = False
    fileLocation = ""

    while not is_valid_path:
        directory_path = input("Enter directory location: ")

        if any(char in directory_path for char in forbiddenDirChars):
            forbidden_in_directory = [char for char in directory_path if char in forbiddenDirChars]
            print(f"Forbidden characters used: {', '.join(forbidden_in_directory)}")
            print("Choose another directory location.")
        else:
            if not os.path.exists(directory_path):
                makeDirConfirmation = input(f"The path `{directory_path}` does not exist, do you want to create it? [y/n]: ").lower()

                if makeDirConfirmation in ("yes", "y"):
                    os.makedirs(directory_path)
                    print(f"Directory created: {directory_path}")
                    fileLocation = directory_path
                    is_valid_path = True
                else: 
                    print("Please enter a valid existing directory.")
            else:
                fileLocation = directory_path
                is_valid_path = True
    
    return fileLocation

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

wb = Workbook()
ws = wb.active

defaultFileLocation = (os.getcwd())
definedFileLocation = input(f"Default file placement: {defaultFileLocation} \nPlace elsewhere? [y/n]: ")

forbiddenDirChars = [">", "<", "*", "?", "!", "|", '"', ":" "\\0"]
forbiddenChars = ["/", "\\", ".", ">", "<", "*", "?", "!", "|", " ", '"', ":" "\\0"]

if definedFileLocation in ("yes", "y"):
    fileLocation = check_location(forbiddenDirChars)
else: fileLocation = defaultFileLocation

fileAmount = int(input("Number of files: "))
isCustomFileName = input("Custom File Name? [y/n]: ").lower()

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
