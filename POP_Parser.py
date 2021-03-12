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
from datetime import date
import datetime
import xlsxwriter
import pdfreader
from pdfreader import PDFDocument, SimplePDFViewer
import sys
import openpyxl

ntpath.basename('a/b/c')

SignalTypes = {
    0:'DSS',
    1:'DVB-S',
    2:'DC2',
    3:'Turbo',
    4:'DVB-S2',
    5:'T-ACM',
    6:'DTVS2',
    7:'DTVS2B',
    8:'Other',
    9:'S2-ACM'
}

# Python program to illustrate union 
# Without repetition  
def Union(lst1, lst2): 
    final_list = list(set(lst1) | set(lst2)) 
    return final_list 

def longestSubstring(str): 

    hasDigit = False

    for char in str:
        if char.isdigit():
            hasDigit = True
            break

    if hasDigit:
        digit = max(re.findall(r'\d+', str), key = len) 
        return digit 
    else:
        return []
# function to trim the last six digits from a string
def get_work_order(file):
    x = re.findall('\d', file)
    work_order = listToString(x)[-6:]                               
    return work_order

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

# String clearning function to remove all new line characters
def clean_all_newline(lst):
    newlist = []
    for l in lst:
        if '\n' in l:
            newlist.append(l.replace('\n', ''))
        else:
            newlist.append(l)
    return newlist

# Function to determine if Reading Type is Post or Pre
# this is necessary because tech's do no always label
# these clearly. 
def CheckReadingType(rowdata):
    if rowdata.upper() == 'POST':
        return 1
    elif rowdata.upper() == 'PST':
        return 1
    elif rowdata.upper() == 'POS':
        return 1
    elif rowdata.upper() == 'PRE':
        return 0
    elif rowdata.upper() == 'AR':
        # print('READING TYPE UNKNOWN')
        return 3


# Function to convert a list to a string
def listToString(s):  
    
    # initialize an empty string 
    str1 = ''  
    
    # traverse in the string   
    for ele in s:  
        str1 += ele   
    
    # return string  
    return str1  

# This function merges two dictionaries
def Merge(dict1, dict2):
    res = {**dict1, **dict2}
    return res


def setup_logfile(pwd):
    LogFilePath = '%s/Log/log.txt'%(pwd)

    if os.path.exists(LogFilePath):
        append_write = 'a' # append if already exists
    else:
        append_write = 'w' # make a new file if not

    # Create Log File
    Log = open(LogFilePath,append_write)
    return Log
# This copy of an ingest file will be stored as either 
# pre-ingest or post-ingest to be stored in a list of dictionaries
def make_ingest_file(dict):
    new_dict = {}
    new_dict['Sat'] = dict['Sat']
    new_dict['Orbit'] = dict['Orbit']
    new_dict['File'] = dict['File']
    new_dict['Site Key'] = dict['Site Key']
    new_dict['Software Version'] = dict['Software Version']
    new_dict['Tran'] = dict['Tran']
    new_dict['Freq'] = dict['Freq']
    new_dict['Polar'] = dict['Polar']
    new_dict['SigType'] = dict['SigType']
    new_dict['CodRate'] = dict['CodRate']
    new_dict['IRD'] = dict['IRD']
    new_dict['C/N'] = dict['C/N']
    new_dict['Es/No'] = dict['Es/No']
    new_dict['Eb/No'] = dict['Eb/No']
    new_dict['Lock'] = dict['Lock']
    new_dict['Limit'] = dict['Limit']
    new_dict['Limit Margin'] = dict['Limit Margin']
    new_dict['Limit Color Code'] = dict['Limit Color Code']
    new_dict['DnLink'] = dict['DnLink']
    new_dict['FreqErr'] = dict['FreqErr']
    new_dict['Baud'] = dict['Baud']
    new_dict['LnbI'] = dict['LnbI']
    new_dict['LnbV'] = dict['LnbV']
    new_dict['Signal Margin'] = dict['Signal Margin']
    new_dict['LNB Model'] = dict['LNB Model']                  
    new_dict['Region'] = dict['Region']        
    new_dict['Date'] = dict['Date']    
    new_dict['Time'] = dict['Time']
    new_dict['LNB Service'] = dict['LNB Service']
    new_dict['LNB System'] = dict['LNB System']
    new_dict['Pre Limit Table'] = dict['Pre Limit Table']
    new_dict['Post Limit Table'] = dict['Post Limit Table']
    return new_dict


def clean_filename(filename, sat):
    filename = filename.upper().replace('SPOP', '')
    filename = filename.upper().replace('POP', '')
    filename = filename.upper().replace(sat.upper(), '')
    return filename

# Create the final ingest file from the pre and post-install ingest files
# This function is broken up into 3 segments
# 1. Compare all pre-ingest dictionaries against all post-ingest dictionaries to find matches
#       - A MATCH means the same Site Key, transponder, Frequency, Polarity, and Orbit
#       - We also calculate Delta C/N and fill in all column values for pre and post metrics (limits, filename, etc.)
# 2. If no matches are found, fill in values for unmatched pre-ingest dictionaries...
# 3. ...fill in values for unmatched post-ingest dictionaries
def make_final_ingest(pre, post):

    # create an array of '0's to track unmatched post-index dictionaries
    # after parsing all files against one another, we are left with a map in the form of an array
    # to signify which dictionary in our list of dictionaries have found a match in the form of a 0 or 1
    # at the end we go through this array and add the remaining post-ingest dictionaries to the final file
    post_index_array = array.array('i',(0 for i in range(0,len(post))))

    # create our final list to be returned and turned into a CSV
    final_list = []

    # begin comparing pre and post ingest dictionaries
    for p in pre:
        # create variable to track pre-post dictionary matches 
        match_found = False

        for t in post:
            if p['Site Key'] == t['Site Key']:
                if p['Tran'] == t['Tran']:
                    if p['Freq'] == t['Freq']:
                        if p['Polar'] == t['Polar']:
                            if p['Orbit'] == t['Orbit']:
                                # A match was found
                                # create empty dictionary to populate with pre and post ingest dictionaries to put into final_list
                                dict_line = {}

                                match_found = True
                                # Change the value of the index of Post ingest dictionary that was matched to 1
                                post_index_array[post.index(t)] = 1
                                
                                # Begin filling in values for final_ingest list
                                dict_line['Site Key'] = p['Site Key']
                                dict_line['Pre Install File'] = p['File']
                                dict_line['Post Install File'] = t['File']
                                dict_line['Software Version'] = p['Software Version']
                                dict_line['Sat'] = t['Sat']
                                dict_line['Tran'] = p['Tran']
                                dict_line['Freq'] = p['Freq']
                                dict_line['Polar'] = p['Polar']

                                if 'R8-' in p['CodRate']:
                                    p['CodRate'] = p['CodRate'].replace('R8-', 'R')
                                elif 'RQ-' in p['SigType']:
                                    p['CodRate'] = p['CodRate'].replace('RQ-', 'R')

                                dict_line['SigType'] = p['SigType']                                  
                                dict_line['CodRate'] = p['CodRate']
                                dict_line['Pre Install C/N'] = p['C/N']
                                dict_line['Pre Install Signal Margin'] = p['Signal Margin']
                                dict_line['Post Install C/N'] = t['C/N']
                                dict_line['Post Install Signal Margin'] = t['Signal Margin']
                                dict_line['Pre Limit Margin'] = p['Limit Margin']
                                dict_line['Pre Limit'] = p['Limit']
                                dict_line['Pre Limit Color Code'] = p['Limit Color Code']
                                dict_line['Post Limit Margin'] = t['Limit Margin']
                                dict_line['Post Limit'] = t['Limit']
                                dict_line['Post Limit Color Code'] = t['Limit Color Code']

                                if t['C/N'] == '.' or p['C/N'] == '.':
                                    dict_line['Delta C/N Pre & Post'] = '.'
                                else:
                                    dict_line['Delta C/N Pre & Post'] = str(round(float(t['C/N']) - float(p['C/N']),2))

                                dict_line['Pre Install Es/No'] = p['Es/No']
                                dict_line['Post Install Es/No'] = t['Es/No']
                                dict_line['Pre Install Eb/No'] = p['Eb/No']
                                dict_line['Post Install Eb/No'] = t['Eb/No']
                                dict_line['Pre Install Lock'] = p['Lock']
                                dict_line['Post Install Lock'] = t['Lock']
                                dict_line['DnLink'] = p['DnLink']
                                dict_line['Pre Install FreqErr'] = p['FreqErr']
                                dict_line['Post Install FreqErr'] = t['FreqErr']
                                dict_line['Baud'] = p['Baud']
                                dict_line['LNB Model'] = p['LNB Model']   
                                dict_line['LNB Service'] = p['LNB Service']
                                dict_line['LNB System'] = p['LNB System']   
                                dict_line['Pre LnbI'] = p['LnbI']
                                dict_line['Pre LnbV'] = p['LnbV'] 
                                dict_line['Post LnbI'] = t['LnbI']
                                dict_line['Post LnbV'] = t['LnbV'] 
                                dict_line['Region'] = p['Region']        
                                dict_line['Pre Install Date'] = p['Date']    
                                dict_line['Post Install Date'] = t['Date'] 
                                dict_line['Pre Install Time'] = p['Time']
                                dict_line['Post Install Time'] = t['Time']
                                dict_line['Limit Table'] = t['Post Limit Table']
                                final_list.append(dict_line)
                                break
                                
        # if no post-ingest files are matched with the pre-ingest file we 
        # are searching under we need to add the pre install file to the final ingest file                      
        if match_found == False:
            # print('P', p['File'])
            dict_line = {}
            dict_line['Site Key'] = p['Site Key']
            dict_line['Pre Install File'] = p['File']
            dict_line['Post Install File'] = '.'
            dict_line['Software Version'] = p['Software Version']
            dict_line['Sat'] = p['Sat']
            dict_line['Tran'] = p['Tran']
            dict_line['Freq'] = p['Freq']
            dict_line['Polar'] = p['Polar']

            if 'R8-' in p['CodRate']:
                p['CodRate'] = p['CodRate'].replace('R8-', 'R')
            elif 'RQ-' in p['SigType']:
                p['CodRate'] = p['CodRate'].replace('RQ-', 'R')

            dict_line['SigType'] = p['SigType']                                   
            dict_line['CodRate'] = p['CodRate']     
            dict_line['Pre Install Signal Margin'] = p['Signal Margin']
            dict_line['Post Install Signal Margin'] = '.'
            dict_line['Pre Install C/N'] = p['C/N']
            dict_line['Post Install C/N'] = '.'
            dict_line['Pre Install Es/No'] = p['Es/No']
            dict_line['Post Install Es/No'] = '.'
            dict_line['Pre Install Eb/No'] = p['Eb/No']
            dict_line['Post Install Eb/No'] = '.'
            dict_line['Pre Install Lock'] = p['Lock']
            dict_line['Post Install Lock'] = '.'
            dict_line['DnLink'] = p['DnLink']
            dict_line['Pre Install FreqErr'] = p['FreqErr']
            dict_line['Post Install FreqErr'] = '.'
            dict_line['Baud'] = p['Baud']
            dict_line['Pre Limit'] = p['Limit']
            dict_line['Pre Limit Color Code'] = p['Limit Color Code']
            dict_line['Pre Limit Margin'] = p['Limit Margin']
            dict_line['Post Limit'] = '.'
            dict_line['Post Limit Color Code'] = '.'
            dict_line['Post Limit Margin'] = '.'
            dict_line['Delta C/N Pre & Post'] = '.'
            dict_line['LNB Model'] = p['LNB Model']       
            dict_line['LNB Service'] = p['LNB Service']
            dict_line['LNB System'] = p['LNB System']
            dict_line['Pre LnbI'] = p['LnbI']
            dict_line['Pre LnbV'] = p['LnbV'] 
            dict_line['Post LnbI'] = '.'
            dict_line['Post LnbV'] = '.'
            dict_line['Region'] = p['Region']     
            dict_line['Pre Install Date'] = p['Date']    
            dict_line['Post Install Date'] = '.'
            dict_line['Pre Install Time'] = p['Time']
            dict_line['Post Install Time'] = '.'
            dict_line['Limit Table'] = p['Pre Limit Table']
            # print(p['Pre Limit Table'])
            final_list.append(dict_line)
        

    # now check which post ingest files were not matched with a pre-ingest file and add them to final ingest file
    for i in range (0, len(post_index_array)):
        if post_index_array[i] == 0:
            dict_line = {}
            # post is a list of dictionaries
            # post[i] is dictionary at position i
            # post[i]['Site Key'] is the Site Key of the dictionary at position i

            dict_line['Site Key'] = post[i]['Site Key']
            dict_line['Pre Install File'] = '.'
            dict_line['Post Install File'] = post[i]['File']
            dict_line['Software Version'] = post[i]['Software Version']
            dict_line['Sat'] = post[i]['Sat']
            dict_line['Tran'] = post[i]['Tran']
            dict_line['Freq'] = post[i]['Freq']
            dict_line['Polar'] = post[i]['Polar']

            if 'R8-' in post[i]['CodRate']:
                post[i]['CodRate'] = post[i]['CodRate'].replace('R8-', 'R')
            elif 'RQ-' in post[i]['SigType']:
                post[i]['CodRate'] = post[i]['CodRate'].replace('RQ-', 'R')

            dict_line['SigType'] = post[i]['SigType']                              
            dict_line['CodRate'] = post[i]['CodRate']
            dict_line['Pre Install Signal Margin'] = '.'
            dict_line['Post Install Signal Margin'] = post[i]['Signal Margin']
            dict_line['Pre Install C/N'] = '.'
            dict_line['Post Install C/N'] = post[i]['C/N']
            dict_line['Pre Install Es/No'] = '.'
            dict_line['Post Install Es/No'] = post[i]['Es/No']
            dict_line['Pre Install Eb/No'] = '.'
            dict_line['Post Install Eb/No'] = post[i]['Eb/No']
            dict_line['Pre Install Lock'] = '.'
            dict_line['Post Install Lock'] = post[i]['Lock']
            dict_line['DnLink'] = post[i]['DnLink']
            dict_line['Pre Install FreqErr'] = '.'
            dict_line['Post Install FreqErr'] = post[i]['FreqErr']
            dict_line['Baud'] = post[i]['Baud']
            dict_line['Pre Limit'] = '.'
            dict_line['Pre Limit Color Code'] = '.'
            dict_line['Pre Limit Margin'] = '.'
            dict_line['Post Limit'] = post[i]['Limit']
            dict_line['Post Limit Color Code'] = post[i]['Limit Color Code']
            dict_line['Post Limit Margin'] = post[i]['Limit Margin']
            dict_line['Delta C/N Pre & Post'] = '.'
            dict_line['LNB Model'] = post[i]['LNB Model']         
            dict_line['LNB Service'] = post[i]['LNB Service']
            dict_line['LNB System'] = post[i]['LNB System']
            dict_line['Pre LnbI'] = '.'
            dict_line['Pre LnbV'] = '.'
            dict_line['Post LnbI'] = post[i]['LnbI']
            dict_line['Post LnbV'] = post[i]['LnbV'] 
            dict_line['Region'] = post[i]['Region']     
            dict_line['Pre Install Date'] = '.'  
            dict_line['Post Install Date'] = post[i]['Date'] 
            dict_line['Pre Install Time'] = '.'
            dict_line['Post Install Time'] = post[i]['Time']
            dict_line['Limit Table'] = post[i]['Post Limit Table']
            final_list.append(dict_line)

    return final_list

def GenerateNewFilename(sat, pol, readtype, sitekey):
    return "%s-%s-%s-%s.spop"%(sat, pol[0],readtype.upper(),sitekey)

def AddBranchToLeaf(file, newfile):
    base = os.path.dirname(file)
    return "%s/%s"%(base, newfile)


def main():

    NoFilterCount = 0
    FilterCount = 0
    AltimeterRader = 0
    filecount = 0
    pop_count = 0
    limitPassCount = 0
    limitFailCount = 0
    LimitWarningCount = 0
    ModCodTableCount = 0
    SLFTableCount = 0
    UnknownReadingTypeCount = 0
    emergencySiteKey = 100000
    now = datetime.datetime.now()
    today = date.today()
    d1 = today.strftime('%d-%m-%Y')
    pre_ingest = []                                                         
    post_ingest = []
    UkknownReadingTypeList = []
    CleanedFiles = []
    

    LocationNames = []

    # Stores PATH to our current directory
    pwd = pathlib.Path().absolute()          

    POP_Files = '%s/POP Files'%(pwd)                                    

    ActiveCarriers = '%s/Excel Files/Active Carrier List.xlsx'%(pwd)
    NoFilterLimits = '%s/SLF_Files/SESCNF.SLF'%(pwd)
    FilterLimits = '%s/SLF_Files/SESC5G.SLF'%(pwd)
    ARFilterLimits = '%s/SLF_Files/SESCAR.SLF'%(pwd)

    ModCodLimitTable = '%s/Excel Files/ModCod Limits Table.csv'%(pwd)
  
    ActiveCarrierDataframeExists = path.exists(ActiveCarriers)
    ActiveCarrierDataframe = pd.read_excel(ActiveCarriers)

    NoFilterLimitDF = pd.read_csv(NoFilterLimits, sep='\t', header=19)
    FilterLimitDF = pd.read_csv(FilterLimits, sep='\t', header=19)
    ARFilterLimitDF = pd.read_csv(ARFilterLimits, sep='\t', header=19)

    ModCodLimitTableDF = pd.read_csv(ModCodLimitTable)
    CSV_FinalFiles = '%s/Output/IngestFile.csv'%(pwd)
    XLSX_FinalFiles = '%s/Output/IngestFile.xlsx'%(pwd)
    # CREATE LOG FILE 
    Log = setup_logfile(pwd)
    
    # ITERATE POP POP FILES AND GETS LIST OF FILENAMES
    list_of_files = getListOfFiles(POP_Files)

    # CHECK THAT WE DID NOT RUN PROGRAM WITHOUT POP FILES
    if len(list_of_files) == 0:
        sys.exit()

    SiteLocation = ""
    # THE MAIN LOOP IN PROCESSING EACH FILE
    for file in list_of_files:                                              

        if 'PRE' in path_leaf(file).upper():
            continue
        
        # if 'A1.SPOP' in file.upper():
        #     os.remove(file)
        #     continue
       
        rowdata = {}
        data = {}
        cleanedfn = ""


        POP_data = []                                                       # Temporary list for storing metrics data from second segement of data in POP file

        if os.path.exists(file):
            with open(file, 'r') as f:
                # print('Opening:', file)
                filecount += 1
                try:
                    newline = ''

                    if 'POST' in file.upper() or 'POS' in file.upper() or 'PST' in file.upper():
                        rowdata['Reading Type (PRE or POST)'] = 'Post'
                    elif 'PRE' in file.upper():
                        rowdata['Reading Type (PRE or POST)'] = 'Pre'
                    else:
                        rowdata['Reading Type (PRE or POST)'] = 'Pre'   

                    lineCount = 0

                    while newline != '\n':  # Loop through the first of two segments of data in POP file line by line saving important data to dictionary.
                        
                        # Read in first line
                        newline = f.readline()   
                        lineCount += 1                    
                        # Store Satellite name in dictionary
                        if 'Satellite1' in newline:
                            rowdata['Sat'] = newline.split(',')[1].replace('\n', '')
                            rowdata['Orbit'] = newline.split(',')[0].replace('\n', '')
                        elif 'Software' in newline:
                            rowdata['Software Version'] = newline.split('\t')[1].replace('\n', '')
                        # Store the Filter Model in dicitonary
                        elif 'LNB Model' in newline:                
                            rowdata['LNB Model'] = newline.split('\t')[1].replace('\n', '')
                        elif 'LNB Service' in newline:
                            rowdata['LNB Service'] = newline.split('\t')[1].replace('\n', '')
                        elif 'LNB System' in newline:
                            rowdata['LNB System'] = newline.split('\t')[1].replace('\n', '')
                        # Store Region in dictionary
                        elif 'Region' in newline:                   
                            rowdata['Region'] = newline.split('\t')[1].replace('\n', '')
                        # Store Date in dictionary 
                        elif 'Date' in newline:                     
                            rowdata['Date'] = newline.split('\t')[1].replace('\n', '')
                        # Store Time in dictionary
                        elif 'Time' in newline:                     
                            rowdata['Time'] = newline.split('\t')[1].replace('\n', '')

                except Exception as e:
                    Log.write('Error: ' + e)
                  
            f.close() 
        
        # ------  EXITING FIRST SEGMENT OF POP FILE  ------ #
        filename = file
        
        fileparts = file.split('/')
        rowdata['Site Key'] = longestSubstring(fileparts[-1])
        # print(rowdata['Site Key'])

        # NewSiteLocation = file.split('/')[7]

        # if len(rowdata['Site Key']) == 6:
        #     print("Pass")
        #     pass
    
        # elif SiteLocation != NewSiteLocation:
            
        #     LocationNames.append({'Location':NewSiteLocation, 'SiteKey': str(emergencySiteKey)})
        #     SiteLocation = NewSiteLocation
        #     rowdata['Site Key'] = str(emergencySiteKey)
        #     emergencySiteKey += 1
        # else:
        #     rowdata['Site Key'] = str(emergencySiteKey)
    

        # ------ Making Pandas Dataframe from metrics Segment of POP File ------ #
      
        POP_Metrics = pd.read_csv(file, sep='\t', header= lineCount - 1)

        # iterate through Dataframe (POP_Metrics) and parse the data
        for index, row in POP_Metrics.iterrows():
                                 
            match = False
            rowdata['Orbit'] = row['Orbit']
            rowdata['Tran'] = row['Tran']
            rowdata['Level'] = row['Level']
            rowdata['Freq'] = row['Freq']
            rowdata['DnLink'] = round(row['DnLink'],2)
            rowdata['C/N'] = row['C/N']
            rowdata['IRD'] = row['IRD']
            rowdata['Eb/No'] = row['Eb/No']
            rowdata['Es/No'] = row['Es/No']
            rowdata['Lock'] = row['Lock']
                      
            if row['Polar'] == 1:                      
                rowdata['Polar'] = 'Horizontal'
            elif row['Polar'] == 2:
                rowdata['Polar'] = 'Vertical'
                
            SigType = SignalTypes[row['SigType']]

            rowdata['CodRate'] = '.'
            
            # if DVB-S is our SigType and coding rate does not start with a Q or 8, we assume it is QPSK
            # else we adopt the QPSK or 8PSK depending on the Q or 8 prefix
            
            # if DVB-S2 is our SigType and coding rate does not start with a Q or an 8, we assume it is 8PSK
            # else we adopt the QPSK or 8PSK depending on the Q or 8 prefix
            
            if '8-' in row['CodRate'] and row['CodRate'].upper() != 'AUTO':
                rowdata['CodRate'] = 'R' + row['CodRate'].replace('8-', '')
                
                SigType = '%s %s'%(SigType, '8PSK')
                
            elif 'Q-' in row['CodRate']and row['CodRate'].upper() != 'AUTO':
                rowdata['CodRate'] = 'R' + row['CodRate'].replace('Q-', '')

                SigType = '%s %s'%(SigType, 'QPSK')

            elif row['CodRate'].upper() != 'AUTO':
                rowdata['CodRate'] = 'R' + row['CodRate']
                
            rowdata['SigType'] = SigType

            # If the SigType is DVB-S we can assume it has a Mod Type of QPSK so we tag that on here
            if rowdata['SigType'] == "DVB-S":
                rowdata['SigType'] = "%s QPSK"%(rowdata['SigType'])

            rowdata['FreqErr'] = round(row['FreqErr'],2)
            rowdata['Baud'] = row['Baud']
            rowdata['LnbV'] = row['LnbV']
            rowdata['LnbI'] = row['LnbI']
            rowdata['LNB'] = row['LNB']
            rowdata['Limit'] = '.'
            rowdata['Limit Margin'] = '.'
            rowdata['Limit Color Code'] = '.'
            rowdata['Pre Limit Table'] = '.'
            rowdata['Post Limit Table'] = '.'

            
            rowdata['File'] = path_leaf(file)

            
            # SET FILENAME TO OUR GENERATED ONE
            
            
            # ------------------------# 
            #  RENAME FILES IN FOLDER #
            # ------------------------#

            # newfile = AddBranchToLeaf(file, rowdata['File'])
            # if newfile != file:
            #     try:
            #         os.rename(file, newfile)
            #     except Exception as e:
            #         print(str(e))

            #-----------------------------------------------------------#
            #                                                           #
            #             ~~     LIMIT PROCESSING      ~~               #
            #                                                           #
            #-----------------------------------------------------------#

            # ----------------------------------------- #
            #           SOFTWARE ABOVE 1.66             #
            # ----------------------------------------- #

            if rowdata['LNB Model'] == 'AR Filter' or rowdata['LNB Model'] == 'Altimeter Radar':
                rowdata['Reading Type (PRE or POST)'] = 'AR'

            rowdata['Reading Type (PRE or POST)'] = 'Pre'

            if float(rowdata['Software Version']) >= 1.66:

                CodRate = rowdata['CodRate']
                SigType = rowdata['SigType']
                
                CodSig = '%s %s'%(SigType, CodRate)

                for i, r in ModCodLimitTableDF.iterrows():
                    if rowdata['Es/No'] == '.':
                        rowdata['Signal Margin'] = '.'
                        
                    elif CodSig == r['Modulation']:
                        rowdata['Signal Margin'] = str(round(float(row['Es/No']) - r['Threshold'],2))
                        break
                    else:
                        rowdata['Signal Margin'] = '.'

                pop_count += 1

                #--------- POST ----------#
                if CheckReadingType(rowdata['Reading Type (PRE or POST)']) == 1:
                    rowdata['Reading Type (PRE or POST)'] = 'Post'
                    rowdata['File'] = GenerateNewFilename(rowdata['Sat'], rowdata['Polar'],rowdata['Reading Type (PRE or POST)'],rowdata['Site Key'])
                    FilterCount += 1

                    for i, r in FilterLimitDF.iterrows():
                        SLF_ModCod = '%s %s %s'%(r['Type'], r['Mod'], r['Code'])
                        ModCod = '%s %s'%(rowdata['SigType'], rowdata['CodRate'])
                        
                        if int(rowdata['Orbit']) == int(r['Orbit']) and rowdata['Tran'] == r['XPR'] and ModCod == SLF_ModCod:
                            
                            match = True
                            SLFTableCount += 1
                            rowdata['Post Limit Table'] = 'SLF'

                            MinLim = float(r['MinLim'])
                            AvgLim = float(r['AvgLim'])

                            if rowdata['Es/No'] == '.':
                                rowdata['Limit'] = '.'
                                rowdata['Limit Color Code'] = '.'
                                rowdata['Limit Margin'] = '.'
                           
                                continue

                            EsNo = float(rowdata['Es/No'])

                            rowdata['Limit Margin'] = round(EsNo - MinLim,2)

                            if EsNo < MinLim:
                                rowdata['Limit'] = 'Fail'
                                rowdata['Limit Color Code'] = 'Red'
                                limitFailCount += 1
                            elif EsNo > MinLim and EsNo < AvgLim:
                                rowdata['Limit'] = 'Pass'
                                rowdata['Limit Color Code'] = 'Yellow'
                                limitPassCount += 1
                                LimitWarningCount += 1
                            elif EsNo > AvgLim:
                                rowdata['Limit'] = 'Pass'
                                rowdata['Limit Color Code'] = 'Green'
                                limitPassCount += 1

                            
                    #--------- NO SLF MATCH ----------#
                    if match == False and rowdata['CodRate'] != '.':
                        CodRate = rowdata['CodRate']
                        SigType = rowdata['SigType']

                        CodSig = '%s %s'%(SigType, CodRate)

                        EsNo = float(rowdata['Es/No'])

                        for i, r in ModCodLimitTableDF.iterrows():

                            if CodSig == r['Modulation']:

                                ModCodTableCount += 1
                            
                                limit = float(r['5G Filter'])
                                
                                rowdata['Limit Margin'] = round(EsNo - limit,2)

                                if EsNo < limit:
                                    rowdata['Limit'] = 'Fail'
                                    rowdata['Limit Color Code'] = 'Red'
                                    limitFailCount += 1
                                # elif CN > MinLim and CN < AvgLim:
                                #     rowdata['Limit'] = 'Pass'
                                #     rowdata['Limit Color Code'] = 'Yellow'
                                elif EsNo > limit:
                                    rowdata['Limit'] = 'Pass'
                                    rowdata['Limit Color Code'] = 'Green'
                                    limitPassCount += 1
                                
                                rowdata['Post Limit Table'] = 'ModCod'
                                break
            

                    post_ingest.append(make_ingest_file(rowdata))

                # ------ ALTIMETER RADAR --------- #
                elif CheckReadingType(rowdata['Reading Type (PRE or POST)']) == 3:

                    
                    rowdata['Reading Type (PRE or POST)'] = 'Pre'
                    rowdata['File'] = GenerateNewFilename(rowdata['Sat'], rowdata['Polar'],rowdata['Reading Type (PRE or POST)'],rowdata['Site Key'])
                    AltimeterRader += 1

                    for i, r in ARFilterLimitDF.iterrows():
                        
                        SLF_ModCod = '%s %s %s'%(r['Type'], r['Mod'], r['Code'])
                        
                        ModCod = '%s %s'%(rowdata['SigType'], rowdata['CodRate'])
                       
               
                        if int(rowdata['Orbit']) == int(r['Orbit']) and rowdata['Tran'] == r['XPR'] and ModCod == SLF_ModCod:
                           
                            match = True
                            SLFTableCount += 1
                            rowdata['Pre Limit Table'] = 'SLF'

                            MinLim = float(r['MinLim'])
                            AvgLim = float(r['AvgLim'])

                            if rowdata['Es/No'] == '.':
                                rowdata['Limit'] = '.'
                                rowdata['Limit Color Code'] = '.'
                                rowdata['Limit Margin'] = '.'
                                continue

                            EsNo = float(rowdata['Es/No'])

                            rowdata['Limit Margin'] = round(EsNo - MinLim,2)
                            
                            if EsNo < MinLim:
                                rowdata['Limit'] = 'Fail'
                                rowdata['Limit Color Code'] = 'Red'
                                limitFailCount += 1
                            elif EsNo > MinLim and EsNo < AvgLim:
                                rowdata['Limit'] = 'Pass'
                                rowdata['Limit Color Code'] = 'Yellow'
                                limitPassCount += 1
                                LimitWarningCount += 1
                            elif EsNo > AvgLim:
                                rowdata['Limit'] = 'Pass'
                                rowdata['Limit Color Code'] = 'Green'
                                limitPassCount += 1

                           
                    #--------- NO SLF MATCH ----------#
                    if match == False and rowdata['CodRate'] != '.':
                        CodRate = rowdata['CodRate']
                        SigType = rowdata['SigType']

                        CodSig = '%s %s'%(SigType, CodRate)
            

                        if rowdata['Es/No'] != '.':
                            EsNo = float(rowdata['Es/No'])

                            for i, r in ModCodLimitTableDF.iterrows():

                                if CodSig == r['Modulation']:
                                    
                                    ModCodTableCount += 1

                                    limit = float(r['AR Filter'])
                                    
                                    rowdata['Limit Margin'] = round(EsNo - limit,2)
                                    
                                    if EsNo < limit:
                                        rowdata['Limit'] = 'Fail'
                                        rowdata['Limit Color Code'] = 'Red'
                                        limitFailCount += 1
                                    # elif CN > MinLim and CN < AvgLim:
                                    #     rowdata['Limit'] = 'Pass'
                                    #     rowdata['Limit Color Code'] = 'Yellow'
                                    elif EsNo > limit:
                                        rowdata['Limit'] = 'Pass'
                                        rowdata['Limit Color Code'] = 'Green'
                                        limitPassCount += 1
                                    
                                    rowdata['Pre Limit Table'] = 'ModCod'
                                    break
                       
                    pre_ingest.append(make_ingest_file(rowdata))

                #--------- PRE ----------#
                elif CheckReadingType(rowdata['Reading Type (PRE or POST)']) == 0:
                    rowdata['Reading Type (PRE or POST)'] = 'Pre'
                    rowdata['File'] = GenerateNewFilename(rowdata['Sat'], rowdata['Polar'],rowdata['Reading Type (PRE or POST)'],rowdata['Site Key'])
                    NoFilterCount += 1

                    
                    for i, r in NoFilterLimitDF.iterrows():
                        SLF_ModCod = '%s %s %s'%(r['Type'], r['Mod'], r['Code'])
                        ModCod = '%s %s'%(rowdata['SigType'], rowdata['CodRate'])
                        
                        
                        if int(rowdata['Orbit']) == int(r['Orbit']) and rowdata['Tran'] == r['XPR'] and ModCod == SLF_ModCod:
                            
                            match = True
                            SLFTableCount += 1
                            rowdata['Pre Limit Table'] = 'SLF'

                            MinLim = float(r['MinLim'])
                            AvgLim = float(r['AvgLim'])

                            if rowdata['Es/No'] == '.':
                                rowdata['Limit'] = '.'
                                rowdata['Limit Color Code'] = '.'
                                rowdata['Limit Margin'] = '.'
                                                 
                                continue

                            EsNo = float(rowdata['Es/No'])

                            rowdata['Limit Margin'] = round(EsNo - MinLim,2)

                            if EsNo < MinLim:
                                rowdata['Limit'] = 'Fail'
                                rowdata['Limit Color Code'] = 'Red'
                                limitFailCount += 1
                            elif EsNo > MinLim and EsNo < AvgLim:
                                rowdata['Limit'] = 'Pass'
                                rowdata['Limit Color Code'] = 'Yellow'
                                limitPassCount += 1
                                LimitWarningCount += 1
                            elif EsNo > AvgLim:
                                rowdata['Limit'] = 'Pass'
                                rowdata['Limit Color Code'] = 'Green'
                                limitPassCount += 1
                          
                            
                    #--------- NO SLF MATCH ----------#
                    if match == False and rowdata['CodRate'] != '.':
                        CodRate = rowdata['CodRate']
                        SigType = rowdata['SigType']

                        CodSig = '%s %s'%(SigType, CodRate)

                        if rowdata['Es/No'] == '.':
                            rowdata['Limit'] = '.'
                            rowdata['Limit Color Code'] = '.'
                            rowdata['Limit Margin'] = '.'
                                                
                            continue

                        EsNo = float(rowdata['Es/No'])

                        for i, r in ModCodLimitTableDF.iterrows():

                            if CodSig == r['Modulation']:
                            
                                ModCodTableCount += 1

                                limit = float(r['No Filter'])
                                
                                rowdata['Limit Margin'] = round(EsNo - limit,2)

                                if EsNo < limit:
                                    rowdata['Limit'] = 'Fail'
                                    rowdata['Limit Color Code'] = 'Red'
                                    limitFailCount += 1
                                # elif CN > MinLim and CN < AvgLim:
                                #     rowdata['Limit'] = 'Pass'
                                #     rowdata['Limit Color Code'] = 'Yellow'
                                elif EsNo > limit:
                                    rowdata['Limit'] = 'Pass'
                                    rowdata['Limit Color Code'] = 'Green'
                                    limitPassCount += 1
                                
                                rowdata['Pre Limit Table'] = 'ModCod'
                                break

                    pre_ingest.append(make_ingest_file(rowdata))
                
            
                
            # ----------------------------------------- #
            #           SOFTWARE BELOW 1.66             #
            # ----------------------------------------- #
            else:
    
                CodRate = rowdata['CodRate']
                SigType = rowdata['SigType']

                CodSig = '%s %s'%(SigType, CodRate)

                for i, r in ModCodLimitTableDF.iterrows():
                    if CodSig == r['Modulation']:                    
                        rowdata['Signal Margin'] = str(round(float(row['Es/No']) - r['Threshold'],2))
                        break
                    else:
                        rowdata['Signal Margin'] = '.'

                pop_count += 1
                
                #--------- PRE ----------#
                if CheckReadingType(rowdata['Reading Type (PRE or POST)']) == 0:
                    rowdata['Reading Type (PRE or POST)'] = 'Pre'
                    rowdata['File'] = GenerateNewFilename(rowdata['Sat'], rowdata['Polar'],rowdata['Reading Type (PRE or POST)'],rowdata['Site Key'])
                    NoFilterCount += 1

                    for i, r in NoFilterLimitDF.iterrows():
                        SLF_ModCod = '%s %s %s'%(r['Type'], r['Mod'], r['Code'])
                        ModCod = '%s %s'%(rowdata['SigType'], rowdata['CodRate'])
                        
                        if int(rowdata['Orbit']) == int(r['Orbit']) and rowdata['Tran'] == r['XPR'] and ModCod == SLF_ModCod:

                            match = True
                            SLFTableCount += 1
                            rowdata['Pre Limit Table'] = 'SLF'

                            MinLim = float(r['MinLim'])
                            AvgLim = float(r['AvgLim'])

                            if rowdata['C/N'] == '.':
                                rowdata['Limit'] = '.'
                                rowdata['Limit Color Code'] = '.'
                                rowdata['Limit Margin'] = '.'
                                                   
                                continue

                            CN = float(rowdata['C/N'])

                            rowdata['Limit Margin'] = round(CN - MinLim,2)

                            if CN < MinLim:
                                rowdata['Limit'] = 'Fail'
                                rowdata['Limit Color Code'] = 'Red'
                                limitFailCount += 1
                            elif CN > MinLim and CN < AvgLim:
                                rowdata['Limit'] = 'Pass'
                                rowdata['Limit Color Code'] = 'Yellow'
                                limitPassCount += 1
                                LimitWarningCount += 1
                            elif CN > AvgLim:
                                rowdata['Limit'] = 'Pass'
                                rowdata['Limit Color Code'] = 'Green'
                                limitPassCount += 1

                            

                    # NO SLF MATCH
                    if match == False and rowdata['CodRate'] != '.':

                        CodRate = rowdata['CodRate']
                        SigType = rowdata['SigType']

                        CodSig = '%s %s'%(SigType, CodRate)

                        CN = float(rowdata['C/N'])

                        for i, r in ModCodLimitTableDF.iterrows():

                            if CodSig == r['Modulation']:

                                ModCodTableCount += 1
                              
                                limit = float(r['No Filter'])
                                
                                rowdata['Limit Margin'] = round(CN - limit,2)

                                if CN < limit:
                                    rowdata['Limit'] = 'Fail'
                                    rowdata['Limit Color Code'] = 'Red'
                                    limitFailCount += 1
                                # elif CN > MinLim and CN < AvgLim:
                                #     rowdata['Limit'] = 'Pass'
                                #     rowdata['Limit Color Code'] = 'Yellow'
                                elif CN > limit:
                                    rowdata['Limit'] = 'Pass'
                                    rowdata['Limit Color Code'] = 'Green'
                                    limitPassCount += 1
                                
                                rowdata['Pre Limit Table'] = 'ModCod'
                                break

                    pre_ingest.append(make_ingest_file(rowdata))   
                
                #--------- ALTIMETER RADAR ----------#
                elif CheckReadingType(rowdata['Reading Type (PRE or POST)']) == 3:
                    rowdata['Reading Type (PRE or POST)'] = 'Pre'
                    rowdata['File'] = GenerateNewFilename(rowdata['Sat'], rowdata['Polar'],rowdata['Reading Type (PRE or POST)'],rowdata['Site Key'])
                    AltimeterRader += 1
             
                    for i, r in ARFilterLimitDF.iterrows():
                        SLF_ModCod = '%s %s %s'%(r['Type'], r['Mod'], r['Code'])
                        ModCod = '%s %s'%(rowdata['SigType'], rowdata['CodRate'])
                        
                        if int(rowdata['Orbit']) == int(r['Orbit']) and rowdata['Tran'] == r['XPR'] and ModCod == SLF_ModCod:

                            match = True
                            SLFTableCount += 1
                            rowdata['Limit Table'] = 'SLF'

                            MinLim = float(r['MinLim'])
                            AvgLim = float(r['AvgLim'])

                            if rowdata['C/N'] == '.':
                                rowdata['Limit'] = '.'
                                rowdata['Limit Color Code'] = '.'
                                rowdata['Limit Margin'] = '.'
                                                   
                                continue

                            CN = float(rowdata['C/N'])

                            rowdata['Pre Limit Margin'] = round(CN - MinLim,2)

                            if CN < MinLim:
                                rowdata['Limit'] = 'Fail'
                                rowdata['Limit Color Code'] = 'Red'
                                limitFailCount += 1
                            elif CN > MinLim and CN < AvgLim:
                                rowdata['Limit'] = 'Pass'
                                rowdata['Limit Color Code'] = 'Yellow'
                                limitPassCount += 1
                                LimitWarningCount += 1
                            elif CN > AvgLim:
                                rowdata['Limit'] = 'Pass'
                                rowdata['Limit Color Code'] = 'Green'
                                limitPassCount += 1

                         
                    #--------- NO SLF MATCH ----------#
                    if match == False and rowdata['CodRate'] != '.':
                        CodRate = rowdata['CodRate']
                        SigType = rowdata['SigType']

                        CodSig = '%s %s'%(SigType, CodRate)

                        CN = float(rowdata['C/N'])

                        for i, r in ModCodLimitTableDF.iterrows():

                            if CodSig == r['Modulation']:
                                
                                ModCodTableCount += 1

                                limit = float(r['AR Filter'])
                                
                                rowdata['Limit Margin'] = round(CN - limit,2)

                                if CN < limit:
                                    rowdata['Limit'] = 'Fail'
                                    rowdata['Limit Color Code'] = 'Red'
                                    limitFailCount += 1
                                # elif CN > MinLim and CN < AvgLim:
                                #     rowdata['Limit'] = 'Pass'
                                #     rowdata['Limit Color Code'] = 'Yellow'
                                elif CN > limit:
                                    rowdata['Limit'] = 'Pass'
                                    rowdata['Limit Color Code'] = 'Green'
                                    limitPassCount += 1
                                
                                rowdata['Pre Limit Table'] = 'ModCod'
                                break

                    pre_ingest.append(make_ingest_file(rowdata))   

                #--------- POST ----------#
                elif CheckReadingType(rowdata['Reading Type (PRE or POST)']) == 1:
                    rowdata['Reading Type (PRE or POST)'] = 'Post'
                    rowdata['File'] = GenerateNewFilename(rowdata['Sat'], rowdata['Polar'],rowdata['Reading Type (PRE or POST)'],rowdata['Site Key'])
                    FilterCount += 1

                    for i, r in FilterLimitDF.iterrows():
                        SLF_ModCod = '%s %s %s'%(r['Type'], r['Mod'], r['Code'])
                        ModCod = '%s %s'%(rowdata['SigType'], rowdata['CodRate'])
                        
                        if int(rowdata['Orbit']) == int(r['Orbit']) and rowdata['Tran'] == r['XPR'] and ModCod == SLF_ModCod:
                            
                            match = True
                            SLFTableCount += 1
                            rowdata['Post Limit Table'] = 'SLF'

                            MinLim = float(r['MinLim'])
                            AvgLim = float(r['AvgLim'])

                            if rowdata['C/N'] == '.':
                                rowdata['Limit'] = '.'
                                rowdata['Limit Color Code'] = '.'
                                rowdata['Limit Margin'] = '.'
                                continue

                            CN = float(rowdata['C/N'])

                            rowdata['Limit Margin'] = round(CN - MinLim,2)

                            if CN < MinLim:
                                rowdata['Limit'] = 'Fail'
                                rowdata['Limit Color Code'] = 'Red'
                                limitFailCount += 1
                            elif CN > MinLim and CN < AvgLim:
                                rowdata['Limit'] = 'Pass'
                                rowdata['Limit Color Code'] = 'Yellow'
                                limitPassCount += 1
                                LimitWarningCount += 1
                            elif CN > AvgLim:
                                rowdata['Limit'] = 'Pass'
                                rowdata['Limit Color Code'] = 'Green'
                                limitPassCount += 1

                            

                   
                    #--------- NO SLF MATCH ----------#
                    if match == False and rowdata['CodRate'] != '.':

                        CodRate = rowdata['CodRate']
                        SigType = rowdata['SigType']

                        CodSig = '%s %s'%(SigType, CodRate)

                        CN = float(rowdata['C/N'])

                        for i, r in ModCodLimitTableDF.iterrows():

                            if CodSig == r['Modulation']:
                                
                                ModCodTableCount += 1

                                limit = float(r['5G Filter'])
                                
                                rowdata['Limit Margin'] = round(CN - limit,2)

                                if CN < limit:
                                    rowdata['Limit'] = 'Fail'
                                    rowdata['Limit Color Code'] = 'Red'
                                    limitFailCount += 1
                
                                elif CN > limit:
                                    rowdata['Limit'] = 'Pass'
                                    rowdata['Limit Color Code'] = 'Green'
                                    limitPassCount += 1
                                
                                rowdata['PostLimit Table'] = 'ModCod'
                                break
                    elif rowdata['CodRate'] != '.':
                        rowdata['Limit Table'] = '.'

                    post_ingest.append(make_ingest_file(rowdata))
                


    #-----------------------------------------------------------#
    #                    CREATE INGEST FILE                     #
    #-----------------------------------------------------------#
    
    ingest_file = make_final_ingest(pre_ingest, post_ingest)                

    #Sort Ingest File
    ingest_file = sorted(ingest_file, key = lambda i: i['Site Key'])

    keys = ingest_file[0].keys()                                           
    with open(CSV_FinalFiles, 'w', newline='')  as output_file:             
        dict_writer = csv.DictWriter(output_file, keys)
        dict_writer.writeheader()
        dict_writer.writerows(ingest_file)

    print('Finished.. Ingest File Created.')

    products = [{'id':46329, 
             'discription':'AD BLeu', 
             'marque':'AZERT',
             'category':'liquid',
             'family': 'ADBLEU', 
             'photos':'D:\\hamzawi\\hamza\\image2py\\46329_1.png'}
            ]




    # create a new workbook
    wb = openpyxl.Workbook()
    ws = wb.active

    ws.append(list(keys))
    # append data
    # iterate `list` of `dict`
    for line in ingest_file:
        # create a `generator` yield product `value`
        # use the fieldnames in desired order as `key`
        values = (line[k] for k in keys)

        # append the `generator values`
        ws.append(values)
  
    wb.save(filename=XLSX_FinalFiles)


    #-----------------------------------------------------------#
    #                       LOG DATA                            #
    #-----------------------------------------------------------#

    # Log statistical data gathered
    Log.write('Parser ran at: ' + str(now) + '\n')
    Log.write('Total files processed: ' + str(filecount) + '\n')
    Log.write('Total lines of data processed: ' + str(pop_count) + '\n')
    Log.write('AR Filter POPs Processed: ' + str(AltimeterRader) + '\n')
    Log.write('No Filter POPs Processed: ' + str(NoFilterCount) + '\n')
    Log.write('5G Filter POPs Processed: ' + str(FilterCount) + '\n')
    Log.write('Carriers found using SLF Limit Tables: ' + str(SLFTableCount) + '\n')
    Log.write('Carriers found using ModCod Limits Table: ' + str(ModCodTableCount) + '\n')
    Log.write('Carriers not identified: ' + str(pop_count - (SLFTableCount + ModCodTableCount)) + '\n')
    Log.write('Limit Passes: ' + str(limitPassCount) + '\n')
    Log.write('Limit Fails: ' + str(limitFailCount) + '\n')
    Log.write('Limit Passes with Low Margins: ' + str(LimitWarningCount) + '\n')
    Log.write('Number of unknown reading types: ' + str(UnknownReadingTypeCount) + '\n')

    for i in UkknownReadingTypeList:
        entry = "\t%s\n"%(i)
        Log.write(entry)

    # We need to read through final ingest list of dictionaries
    # We need to sort by Work Order
    # After sorting, we go through one by one and check post-install limits pass/fail values
    # Each time we get a fail, we add the fails to a list of dictionaries with..
    #   Site Key, Sat, Pol and Transponder
    # If we iterate through ALL the site keys in the final ingest file and they all pass
    # we add that site key to a list of passing site keys

    #-----------------------------------------------------------#
    #                    POST PROCESSSING DATA                  #
    #-----------------------------------------------------------#

    NoPostSiteKeys = []

    SiteKey = ""
    HasPost = False

    for line in ingest_file:
        # if the first line
        if SiteKey == "":
            SiteKey = line['Site Key']
            # check if its post install limit is not equal to '.'
            if line['Post Limit'] != '.':
                HasPost = True  # We have a fail or a Pass
        # if we come upon a new sitekey from the previous 
        elif SiteKey != line['Site Key']:

            if HasPost == False:
                NoPostSiteKeys.append(SiteKey)

            SiteKey = line['Site Key']

            if line['Post Limit'] != '.':
                HasPost = True
            else:
                HasPost = False
        else:
            if line['Post Limit'] != '.':
                HasPost = True
    if HasPost == False:
        NoPostSiteKeys.append(SiteKey)

    PostFailedSiteKeys = []
    PostFailedLines = []
    PostPassedSiteKeys = []
    PostPassedLines = []

    SiteKey = ""
    Pass = True

    for line in ingest_file:
   
        if SiteKey == "":
            SiteKey = line['Site Key']
            # check if its post install limit is not equal to '.'
            if line['Post Limit'] == 'Fail':
                Pass = False  # We have a fail or a Pass
                PostFailedLines.append(line)
            elif line['Post Limit'] == '.':
                Pass = True
            elif line['Post Limit'] == 'Pass':
                PostPassedLines.append(line)
        # if we come upon a new sitekey from the previous 
        elif SiteKey != line['Site Key']:

            if Pass == True and SiteKey not in NoPostSiteKeys:
                PostPassedSiteKeys.append(SiteKey)
            elif Pass == False and SiteKey not in NoPostSiteKeys:
                PostFailedSiteKeys.append(SiteKey)

            SiteKey = line['Site Key']

            if line['Post Limit'] == 'Pass':
                Pass = True
                PostPassedLines.append(line)
            elif line['Post Limit'] == 'Fail':
                Pass = False
                PostFailedLines.append(line)
            else:
                Pass = True
        else:
            if line['Post Limit'] == 'Fail':
                Pass = False
                PostFailedLines.append(line)
            elif line['Post Limit'] == 'Pass':
                PostPassedLines.append(line)

    if Pass == True and SiteKey not in NoPostSiteKeys:
        PostPassedSiteKeys.append(SiteKey)
    elif Pass == False and SiteKey not in NoPostSiteKeys:
        PostFailedSiteKeys.append(SiteKey)
    
    PreFailedSiteKeys = []
    PreFailedLines = []
    PrePassedSiteKeys = []
    PrePassedLines = []

    SiteKey = ""
    Pass = True
  
    for line in ingest_file:
        if SiteKey == "":
            SiteKey = line['Site Key']
            # check if its post install limit is not equal to '.'
            if line['Pre Limit'] == 'Fail':
                Pass = False  # We have a fail or a Pass
                PreFailedLines.append(line)
            elif line['Pre Limit'] == '.':
                Pass = True
            elif line['Pre Limit'] == 'Pass':
                PrePassedLines.append(line)
        # if we come upon a new sitekey from the previous 
        elif SiteKey != line['Site Key']:

            if Pass == True and SiteKey in NoPostSiteKeys:
                PrePassedSiteKeys.append(SiteKey)
            if Pass == False and SiteKey in NoPostSiteKeys:
                PreFailedSiteKeys.append(SiteKey)

            SiteKey = line['Site Key']

            if line['Pre Limit'] == 'Pass':
                Pass = True
                PrePassedLines.append(line)
            elif line['Pre Limit'] == 'Fail':
                Pass = False
                PreFailedLines.append(line)
            else:
                Pass = True
        else:
            if line['Pre Limit'] == 'Fail':
                Pass = False
                PreFailedLines.append(line)
            elif line['Pre Limit'] == 'Pass':
                PrePassedLines.append(line)
    if Pass == True and SiteKey in NoPostSiteKeys:
        PrePassedSiteKeys.append(SiteKey)
    if Pass == False and SiteKey in NoPostSiteKeys:
        PreFailedSiteKeys.append(SiteKey)

    PassingSiteKeys = Union(PostPassedSiteKeys, PrePassedSiteKeys)
    FailingSiteKeys = Union(PostFailedSiteKeys,PreFailedSiteKeys)
    FailedLines = PostFailedLines + PreFailedLines
    PassedLines = PostPassedLines + PrePassedLines

    AllSites = Union(PassingSiteKeys,FailingSiteKeys)
    AllLines = FailedLines + PassedLines



    #-----------------------------------------------------------#
    #                    WRITE PROCESSED DATA                   #
    #-----------------------------------------------------------#



    ReportPass = '%s/Output/Report.xlsx'%(pwd)

    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook(ReportPass)
    worksheet = workbook.add_worksheet()

    cell_format = workbook.add_format()

    cell_format.set_align('center')
    cell_format.set_align('vcenter')

    worksheet.set_column('A:A', 6.5)
    worksheet.set_column('B:B', 11.83)
    worksheet.set_column('C:C', 11.33)

    worksheet.set_column('E:E', 9.5)
    worksheet.set_column('F:F', 11.83)

    worksheet.set_column('G:G', 6.5)
    worksheet.set_column('H:H', 6.5)
    worksheet.set_column('I:I', 5.83)
    worksheet.set_column('J:J', 6.5)
    worksheet.set_column('K:K', 9.5)
    worksheet.set_column('L:L', 9.5)
    worksheet.set_column('M:M', 9.5)
    worksheet.set_column('N:N', 8.0)
   
    worksheet.set_row(0, 30)
    worksheet.set_row(1, 20)

    # Create a format to use in the merged range.
    merge_format = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })

    Bold_Format = workbook.add_format({
        'bold': 1,
        'align': 'center',
        'valign': 'vcenter'
    })

    border_format=workbook.add_format({
        'border':1,
    })

    # Merge series of cells for row one
    worksheet.merge_range('A1:C1', 'All Sites', merge_format)
    worksheet.write(0, 4,'Passed Sites', merge_format)
    worksheet.merge_range('G1:N1', 'Failed Sites', merge_format)
    worksheet.write(0, 15,'Site Key', merge_format)
    worksheet.write(0, 16,'Passed', merge_format)
    worksheet.write(0, 17,'Sat', merge_format)
    worksheet.merge_range('S1:T1', 'Pass', merge_format)
    worksheet.merge_range('U1:V1', 'Fail', merge_format)
    worksheet.merge_range('W1:X1', 'No Lock', merge_format)

    # Writing Row 2 
    RowTwo = ['Site Key', 'Passed Carriers', 'Failed Carriers','', 'Site Key','', 'Site Key', 'Sat', 'Pol', 'Tran', 'Pre Margin', 'Post Margin', 'Limits Used', 'Deg','', '', '', '', 'Pre', 'Post','Pre', 'Post','Pre', 'Post']

    column_num = 0
    for column in RowTwo:
        if column != '':
            worksheet.write(1,column_num, column, Bold_Format)
        column_num += 1

    Log.write("All Sites:\n")
    Log.write("\tSiteKey\t# Passed\t# Failed\n")

    row_num = 2 #for Report File
    for i in AllSites:
        passedCount = 0
        failedCount = 0

        for row in PassedLines:
            if row['Site Key'] == i:
                passedCount += 1
        for row in FailedLines:
            if row['Site Key'] == i:
                failedCount += 1


        worksheet.write(row_num,0, int(i), cell_format)
        worksheet.write(row_num,1, passedCount, cell_format)
        worksheet.write(row_num,2, failedCount, cell_format)

        row_num += 1

        entry = "\t%s\t%s\t\t\t%s\n"%(i, str(passedCount), str(failedCount))
        Log.write(entry)

    dimensions = 'A1:C%s'%(str(row_num))
    worksheet.conditional_format( dimensions, { 'type' : 'no_blanks' , 'format' : border_format} )

    Log.write("Passed Sites:\n")
    Log.write("\tSiteKey\t# Passed\n")
    
    row_num = 2 #for Report File
    for i in PassingSiteKeys:
        # passedCount = 0
        # for row in PassedLines:
        #     if row['Site Key'] == i:
        #         passedCount += 1
            
        worksheet.write(row_num,4, int(i), cell_format)
        # worksheet.write(row_num,5, passedCount, cell_format)

        row_num += 1

        entry = "\t%s\t%s \n"%(i, passedCount)
        Log.write(entry)

    dimensions = 'E1:E%s'%(str(row_num))
    worksheet.conditional_format( dimensions, { 'type' : 'no_blanks' , 'format' : border_format} )

    Log.write("Failed Sites:\n")

    row_num = 2 #for Report File

    for i in FailingSiteKeys:
        Log.write("\t" + i + "\n\t\tSat\t\tPol\t\t\tTran\tPre Margin\tPost Margin\n")
        
        worksheet.write(row_num,6, int(i), cell_format)
        row_num += 1

        for row in ingest_file:
            if i == row['Site Key']:
                entry = "\t\t%s\t%s\t%s\t%s\t\t\t%s\n"%(row['Sat'], row['Polar'], row['Tran'], row['Post Limit Margin'], row['Pre Limit Margin'])
                Log.write(entry)

                worksheet.write(row_num,7, row['Sat'],cell_format)
                worksheet.write(row_num,8, row['Polar'],cell_format)
                worksheet.write(row_num,9, row['Tran'],cell_format)
                worksheet.write(row_num,10, row['Pre Limit Margin'],cell_format)
                worksheet.write(row_num,11, row['Post Limit Margin'],cell_format)

                if row['Limit Table'] == 'ModCod':
                    worksheet.write(row_num,12, "ModCod", cell_format)

                elif row['Limit Table'] == 'SLF':
                    worksheet.write(row_num,12, "SLF", cell_format)

                elif row['Limit Table'] == '.':
                    worksheet.write(row_num,12, ".", cell_format)

                worksheet.write(row_num,13,row['Delta C/N Pre & Post'], cell_format)

                dimension = 'M1:M%s'%(str(row_num))
                worksheet.conditional_format(dimension, {'type': 'data_bar', 'data_bar_2010':True})
                dimension = 'L1:L%s'%(str(row_num))
                worksheet.conditional_format(dimension, {'type': 'data_bar', 'data_bar_2010':True})

                row_num += 1
        dimension = 'M1:M%s'%(str(row_num))
        worksheet.conditional_format(dimension, {'type': 'data_bar', 'data_bar_2010':True})
        dimension = 'L1:L%s'%(str(row_num))
        worksheet.conditional_format(dimension, {'type': 'data_bar', 'data_bar_2010':True})

    dimensions = 'G1:N%s'%(str(row_num))
    worksheet.conditional_format( dimensions, { 'type' : 'no_blanks' , 'format' : border_format} )
    worksheet.conditional_format( dimensions, { 'type' : 'blanks' , 'format' : border_format} )
    row_num = 2 #for Report File

    for i in AllSites:
        passedCount = 0
        failedCount = 0

        worksheet.write(row_num,15, int(i), cell_format)
        if i in PassingSiteKeys:
            worksheet.write(row_num,16, "Yes", cell_format)
        elif i in FailingSiteKeys:
            worksheet.write(row_num,16, "No", cell_format)

        row_num += 1

        Satlist = []

        for row in AllLines:
            if row['Site Key'] == i and row['Sat'] not in Satlist:
                Satlist.append(row['Sat'])

        for sat in Satlist:
            PassPre = 0
            PassPost = 0
            FailPre = 0
            FailPost = 0
            NoLockPre = 0
            NoLockPost = 0
            SLF = 0
            ModCod = 0

            for row in ingest_file:
                if sat == row['Sat'] and row['Pre Limit'] == 'Pass' and i == row['Site Key']:
                    PassPre += 1
                if sat == row['Sat'] and row['Post Limit'] == 'Pass' and i == row['Site Key']:
                    PassPost += 1
                if sat == row['Sat'] and row['Pre Limit'] == 'Fail' and i == row['Site Key']:
                    FailPre += 1
                if sat == row['Sat'] and row['Post Limit'] == 'Fail' and i == row['Site Key']:
                    FailPost += 1
                if sat == row['Sat'] and row['Pre Install Lock'] == 0 and i == row['Site Key']:
                    NoLockPre += 1
                if sat == row['Sat'] and row['Post Install Lock'] == 0 and i == row['Site Key']:
                    NoLockPost += 1
            
            worksheet.write(row_num,17, sat, cell_format)
            worksheet.write(row_num,18, PassPre, cell_format)
            worksheet.write(row_num,19, PassPost, cell_format)
            worksheet.write(row_num,20, FailPre, cell_format)
            worksheet.write(row_num,21, FailPost, cell_format)
            worksheet.write(row_num,22, NoLockPre, cell_format)
            worksheet.write(row_num,23, NoLockPost, cell_format)
            
            row_num += 1

            dimensions = 'P1:X%s'%(str(row_num))
            worksheet.conditional_format( dimensions, { 'type' : 'no_blanks' , 'format' : border_format} )
            worksheet.conditional_format( dimensions, { 'type' : 'blanks' , 'format' : border_format} )
            
            

    Log.write('--------------------------------------------------------------\n')

    workbook.close()

    Log.close()

    print("Created Report")

if __name__ == '__main__':
    main()