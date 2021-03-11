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
            
            allFiles.append(fullPath)             
    return allFiles   
    
    
    
    
    
def main():
    
    pwd = pathlib.Path().absolute()  

    POP_Files = '%s/Unclean POP Files'%(pwd)

    CleanFiles = '%s/Clean POP Files'%(pwd)

    list_of_files = []
    Cleaned_list = []

    try:    
        list_of_files = getListOfFiles(POP_Files)
    except Exception as e:
        print(e)
       
    for file in list_of_files:                                              # ITERATE THROUGH POP FILES and produce CSV file
        
        metadata = {}
        data = {}

        try:
            metadata['File'] = path_leaf(file)
        except Exception as e:
            print(e)

        POP_data = []                                                       # Temporary list for storing metrics data from second segement of data in POP file

        if os.path.exists(file):
            with open(file, 'r') as f:
                # print('Opening:', file)
             
                try:
                    newline = ''

                    if 'POST' in file.upper() or 'POS' in file.upper() or 'PST' in file.upper():
                        reading_type = 'Post'
                    elif 'PRE' in file.upper():
                        reading_type = 'Pre'
                    else:
                        reading_type = '.'

                    metadata['Reading Type (PRE or POST)'] = reading_type             

                    lineCount = 0

                    while newline != '\n':  # Loop through the first of two segments of data in POP file line by line saving important data to dictionary.
                        
                        # Read in first line
                        newline = f.readline()   
                        lineCount += 1                    
                        # Store Satellite name in dictionary
                        if 'Satellite1' in newline:
                            metadata['Sat'] = newline.split(',')[1].replace('\n', '')
                            metadata['Orbit'] = newline.split(',')[0].replace('\n', '')
                        elif 'Software' in newline:
                            metadata['Software Version'] = newline.split('\t')[1].replace('\n', '')
                        # Store the Filter Model in dicitonary
                        elif 'LNB Model' in newline:                
                            metadata['LNB Model'] = newline.split('\t')[1].replace('\n', '')
                        elif 'LNB Service' in newline:
                            metadata['LNB Service'] = newline.split('\t')[1].replace('\n', '')
                        elif 'LNB System' in newline:
                            metadata['LNB System'] = newline.split('\t')[1].replace('\n', '')
                        # Store Region in dictionary
                        elif 'Region' in newline:                   
                            metadata['Region'] = newline.split('\t')[1].replace('\n', '')
                        # Store Date in dictionary 
                        elif 'Date' in newline:                     
                            metadata['Date'] = newline.split('\t')[1].replace('\n', '')
                        # Store Time in dictionary
                        elif 'Time' in newline:                     
                            metadata['Time'] = newline.split('\t')[1].replace('\n', '')

                except Exception as e:
                    pass
                  
            f.close() 

        print(metadata)
if __name__ == '__main__':
    main()