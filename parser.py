#!/usr/bin/env python

import csv
import datetime
import xml.etree.ElementTree as ET
import xlrd
import logging
#logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s\n%(message)s')
logging.basicConfig(level=logging.DEBUG, format='%(message)s')


#convert CSV to XML
def parse_CSV(comp, colNames, req, of, file):

    #define ET and root
    root = ET.Element("ROOT")
    tree = ET.ElementTree(root)
    base = ET.SubElement(root, "BaseConfig")
    overflow = ET.SubElement(root, "PaymentsOverflow")
    
    with open(file) as csvfile:
        #use i for naming each second level element (under root)
        i = 1
        reader = csv.DictReader(csvfile)
        
        #get \n off of last (or any) fieldname
        fnList = []
        for fn in reader.fieldnames:
            fnList.append((fn.replace('\n', '')))
        
        #redefine reader with updated fieldnames
        reader = csv.DictReader(csvfile, fieldnames=fnList)

        #read each line of the csv file
        #parse into XML tree
        for row in reader:
            rowNum = "row" + str(i)
            baseElm = ET.SubElement(root, rowNum)
            ofElm = ET.SubElement(root, rowNum)  
            count = 1
            for field in colNames:
                #required fields
                if count <= req:
                    tag = field
                    item = ET.SubElement(baseElm, tag)
                    item.text = row[field]
                #overflow
                else:
                    tag = field
                    item = ET.SubElement(ofElm, tag)
                    item.text = row[field]
                count += 1
            i+=1

    date = datetime.datetime.now().strftime("%y_%m_%d")
    #file naming: SUBJECT TO CHANGE
    parsedFileName = "OUTPUTS/"+comp + "_Parsed" + date + ".xml"           

    #write tree to new file
    tree.write(parsedFileName)
    logging.info(parsedFileName + " has been created")
    

#convert excel sheet to CSV
def excel_to_csv(comp, colNames, req, of, excelFile, sheetName):
    wb = xlrd.open_workbook(excelFile)
    sh = wb.sheet_by_name(sheetName)

    path = "OUTPUTS/"+comp+"CSV.csv"

    csv_file = open(path, 'w')

    wr = csv.writer(csv_file)

    for rownum in range(sh.nrows):
        wr.writerow(sh.row_values(rownum))
    
    csv_file.close()

    #convert CSV to XML
    parse_CSV(comp, colNames, req, of, path)


#convert text file to CSV
def txt_to_csv(comp, colNames, req, of, file, delim):
    #read text file
    txtFile = open(file, 'r',  newline='')
    lines = txtFile.readlines()
    txtFile.close()

    #create and open csv
    path = "OUTPUTS/"+comp+"CSV.csv"
    csv_file = open(path, 'w')
    wr = csv.writer(csv_file)

    for line in lines:
        currentLine = line.split(delim)
        wr.writerow(currentLine)

    csv_file.close()

    #convert CSV to XML   
    parse_CSV(comp, colNames, req, of, path)

  
def main():
        
    #***HARD CODE COMPANY NAME***

    #if CheckFileExample.xlsx
    ##check setup file for file type = xlsx!!!!
    #company = "ranger"
    
    #if TabDelimitedCheckBatch.txt
    ##check setup file for file type = txt!!!!
    #company = "ranger"
    #if CSVPaymentFile.csv
    #company = "springs"

    #if _TBAPVirtualCreditCardPaymentExport.txt
    #company = "samet"

    #if JRColeCSVFile_Final.csv
    company = "jrcole"

    #get configuration based on company's SETUP file
    path = "SETUPS/"+company+"_SETUP.txt"
    f = open(path,'r').read().splitlines()

    contents = []

    #read contents of file and place in list
    for i in f:
        contents.append(i)
    
    logging.info("SETUP File Read")

    #get fileType (first item in SETUP file)
    fileType = contents.pop(0)

    logging.info("...Processing " + fileType + "...\n")

    if fileType == "CSV":
        #FILETYPE
        #NUM OF REQUIRED
        #NUM OF OVERFLOW
        #...fields in proper order...
        required = int(contents.pop(0))
        overflow = contents.pop(0)

        '''add extra function to grab files...
        ...from server host an rename before executing parser
        NO SPACES'''
        csvFile = "JRColeCSVFile_Final.csv"
        
        parse_CSV(company, contents, required, overflow, csvFile)
    
    elif fileType == "XLSX":
        #FILETYPE
        #NUM OF REQUIRED
        #NUM OF OVERFLOW
        #SHEET NAME
        #...fields in proper order...
        required = int(contents.pop(0))
        overflow = contents.pop(0)
        sheet = contents.pop(0)

        excelFile = "CheckFileExample.xlsx"

        excel_to_csv(company, contents, required, overflow, excelFile, sheet)
    
    elif fileType == "TXT":
        #FILETYPE
        #NUM OF REQUIRED
        #NUM OF OVERFLOW
        #DELIMITER TYPE
        #...fields in proper order...
        required = int(contents.pop(0))
        overflow = contents.pop(0)
        delimiter = contents.pop(0)

        if delimiter == 'tab':
            delimiter = '\t'
        elif delimiter == 'pipe':
            delimiter = '|'
        elif delimiter == 'comma':
            delimiter = ','

        textFile = '_TBAPVirtualCreditCardPaymentExport.txt'

        txt_to_csv(company, contents, required, overflow, textFile, delimiter)

    else:
        logging.info('Check SETUP file for company. File Type not found.')
    
    logging.info("...DONE.")

main()

