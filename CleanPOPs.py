import pathlib
import os
import csv
import pandas as pd 
from pprint import pprint
import os.path
from os import path
import re
import array
import ntpath

from shutil import copyfile

def longestSubstring(str): 
   
    digit = max(re.findall(r'\d+', str), key = len) 
      
    return digit 
# function that takes a path to a file and returns it's leaf file
def path_leaf(path):
    head, tail = ntpath.split(path)
    return tail or ntpath.basename(head)

#For the given path, get the List of all files in the directory tree 
def getListOfFiles(dirName):
    # create a list of file and sub directories 
    # names in the given directory 
    listOfFile = os.listdir(dirName)
    allFiles = list()
    
    # Iterate over all the entries
    for entry in listOfFile:
        # Create full path
        fullPath = os.path.join(dirName, entry)
        
        # If entry is a directory then get the list of files in this directory 
        if os.path.isdir(fullPath):
            allFiles = allFiles + getListOfFiles(fullPath)
        elif 'SPOP' == fullPath.upper()[-4:]:
            
            allFiles.append(path_leaf(fullPath))             
    return allFiles

# String clearning function to remove all new line characters
def clean_all_newline(lst):
    newlist = []
    for l in lst:
        if '\n' in l:
            newlist.append(l.replace('\n', ''))
        else:
            newlist.append(l)
    return newlist


def AllAlpha(str):
    alpha = False
    digit = False

    for char in fcomponents[0]:
        if char.isalpha():
            alpha = True
        if char.isdigit():
            digit = True

    return alpha and not digit
    


pwd = pathlib.Path().absolute()  

POP_Files = '%s/Unclean POP Files'%(pwd)

CleanFiles = '%s/Clean POP Files'%(pwd)

list_of_files = []

try:    
    list_of_files = getListOfFiles(POP_Files)
except Exception as e:
    print("Error", e)



# copyfile(src, dst)



pprint(list_of_files)
#REMOVE .SPOP
#SPLIT BY -

filenames = []

for file in list_of_files:
    if 'A1' in file:
        continue

    filename = file.upper().split('.SPOP')[0]
    
    fcomponents = filename.split('-')

    alpha = False
    digit = False

    fclean = ""
    if fcomponents[0].isalpha():
        #CHECK IF SECOND COMPONENT IS ALL NUMBERS
        if fcomponents[1].isnumeric():
            fclean = fclean + fcomponents[0] + fcomponents[1]

        if fcomponents[2] == 'V' or fcomponents[2] == 'H':
            fclean = fclean + '-' + fcomponents[2]
        
        if fcomponents[3] == 'PRE' or fcomponents[3] == 'POST':
            fclean = fclean + '-' + fcomponents[3]
        
        if len(fcomponents[4]) == 6 and fcomponents[4].isnumeric():
            fclean = fclean + '-' + fcomponents[4]
        elif len(fcomponents[4]) != 6 and not fcomponents[4].isnumeric():
            workorder = longestSubstring(fcomponents[4])
            fclean = fclean + '-' + str(workorder)


    else:
        
        fclean = fclean + fcomponents[0]

        if fcomponents[1] == 'V' or fcomponents[1] == 'H':
            fclean = fclean + '-' + fcomponents[1]
        
        if fcomponents[2] == 'PRE' or fcomponents[2] == 'POST':
            fclean = fclean + '-' + fcomponents[2]
        
        if len(fcomponents[3]) == 6 and fcomponents[3].isnumeric():
            print("a",fcomponents)
            fclean = fclean + '-' + fcomponents[3]
        elif len(fcomponents[3]) != 6 and not fcomponents[3].isnumeric():
            print("b",fcomponents)
            workorder = longestSubstring(fcomponents[3])
            fclean = fclean + '-' + str(workorder)

    fclean = fclean + '.spop'
    filenames.append(fclean)

    copyfile(POP_Files + '/' + file , CleanFiles + "/" + fclean)
pprint(filenames)

# The anatomy of a satellite is Text + Digits

# if the first component is all string and the second component is all digit we should combine these two to get the satellite

# if the last entry is longer than 6 characters long we need to dig for the site key

