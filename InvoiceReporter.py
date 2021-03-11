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

# Python program to illustrate union 
# Without repetition  
def Union(lst1, lst2): 
    final_list = list(set(lst1) | set(lst2)) 
    return final_list 

# Function to convert a list to a string
def listToString(s):  
    
    # initialize an empty string 
    str1 = ''  
    
    # traverse in the string   
    for ele in s:  
        str1 += ele   
    
    # return string  
    return str1  

def GetSiteKeys():
    pwd = pathlib.Path().absolute()   
    # INVOICE PDF WORK STARTS HERE
    invoice_directory = '%s/Invoices'%(pwd)
    invoice_list = os.listdir(invoice_directory)
    
    WorkIDList = ['372856']

    InvoiceInfo = {}

    print(invoice_list)

    for invoice in invoice_list:
        
        pdf = "%s/Invoices/%s"%(pwd,invoice)
        fd = open(pdf, "rb")
        viewer = SimplePDFViewer(fd)
        viewer.render()
        raw_invoice_data = viewer.canvas.strings
        
        xyz = listToString(raw_invoice_data)
        xyz = xyz.split(' ')
        try:
            WorkOrderIndex = xyz.index('Order') + 1
            WorkOrder = xyz[WorkOrderIndex].split('-')[1]
            WorkIDList.append(WorkOrder)
        except:
           pass 
        # print(xyz)
    print("Work Orders", WorkIDList)
    return WorkIDList

def main():
    
    WorkIdList = GetSiteKeys()

    pwd = pathlib.Path().absolute()       

    IngestPath = "%s/CSV_FinalFiles/IngestFile.csv"%(pwd)
    ingestdf = pd.read_csv(IngestPath) 

    
    NoPostSiteKeys = []

  
    SiteKey = ""
    HasPost = False

    for index, line in ingestdf.iterrows():
        # if the first line
        if str(line['Site Key']) in WorkIdList:

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

    for index, line in ingestdf.iterrows():
        if str(line['Site Key']) in WorkIdList:
       
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
  
    for index, line in ingestdf.iterrows():
        if str(line['Site Key']) in WorkIdList:
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
    #           Create Excel Spreadsheet For Ramiro             #
    #-----------------------------------------------------------#
    ReportPass = '%s/Excel Files/Report.xlsx'%(pwd)

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

    worksheet.set_column('H:H', 6.5)
    worksheet.set_column('I:I', 5.83)
    worksheet.set_column('J:J', 8.33)
    worksheet.set_column('K:K', 6.5)
    worksheet.set_column('L:L', 8.83)
    worksheet.set_column('M:M', 9.5)
    worksheet.set_column('N:N', 9.17)
   
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
    worksheet.merge_range('H1:N1', 'Failed Sites', merge_format)
    worksheet.write(0, 15,'Site Key', merge_format)
    worksheet.write(0, 16,'Passed', merge_format)
    worksheet.write(0, 17,'Sat', merge_format)
    worksheet.merge_range('S1:T1', 'Pass', merge_format)
    worksheet.merge_range('U1:V1', 'Fail', merge_format)
    worksheet.merge_range('W1:X1', 'No Lock', merge_format)

    # Writing Row 2 
    RowTwo = ['Site Key', 'Passed Carriers', 'Failed Carriers','', 'Site Key','','', 'Site Key', 'Sat', 'Pol', 'Tran', 'Pre Margin', 'Post Margin', 'Limits Used', '', '', '', '', 'Pre', 'Post','Pre', 'Post','Pre', 'Post']

    column_num = 0
    for column in RowTwo:
        if column != '':
            worksheet.write(1,column_num, column, Bold_Format)
        column_num += 1

   

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


    dimensions = 'A1:C%s'%(str(row_num))
    worksheet.conditional_format( dimensions, { 'type' : 'no_blanks' , 'format' : border_format} )
    
    row_num = 2 #for Report File
    for i in PassingSiteKeys:
        # passedCount = 0
        # for row in PassedLines:
        #     if row['Site Key'] == i:
        #         passedCount += 1
            
        worksheet.write(row_num,4, int(i), cell_format)
        # worksheet.write(row_num,5, passedCount, cell_format)

        row_num += 1


    dimensions = 'E1:E%s'%(str(row_num))
    worksheet.conditional_format( dimensions, { 'type' : 'no_blanks' , 'format' : border_format} )

    row_num = 2 #for Report File

    for i in FailingSiteKeys:
        
        worksheet.write(row_num,7, int(i), cell_format)
        row_num += 1

        for index, row in ingestdf.iterrows():
            if i == row['Site Key']:
                entry = "\t\t%s\t%s\t%s\t%s\t\t\t%s\n"%(row['Sat'], row['Polar'], row['Tran'], row['Post Limit Margin'], row['Pre Limit Margin'])
            
                worksheet.write(row_num,8, row['Sat'],cell_format)
                worksheet.write(row_num,9, row['Polar'],cell_format)
                worksheet.write(row_num,10, row['Tran'],cell_format)
                worksheet.write(row_num,11, row['Post Limit Margin'],cell_format)
                worksheet.write(row_num,12, row['Pre Limit Margin'],cell_format)

                if row['Limit Table'] == 'ModCod':
                    worksheet.write(row_num,13, "ModCod", cell_format)

                elif row['Limit Table'] == 'SLF':
                    worksheet.write(row_num,13, "SLF", cell_format)

                dimension = 'M1:M%s'%(str(row_num))
                worksheet.conditional_format(dimension, {'type': 'data_bar', 'data_bar_2010':True})
                dimension = 'L1:L%s'%(str(row_num))
                worksheet.conditional_format(dimension, {'type': 'data_bar', 'data_bar_2010':True})

                row_num += 1
    
    dimensions = 'H1:N%s'%(str(row_num))
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
            for index, row in ingestdf.iterrows():
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
            
            

  

    workbook.close()

if __name__ == '__main__':
    main()