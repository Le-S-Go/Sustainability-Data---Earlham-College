import pdfplumber
import pandas as pd
import calendar
import argparse
import os
import openpyxl

# The key of this dictionary is a meter number, the value is a list with index 0 = account number, 1 = amount, 2 = usage
data_dic = {}

# This function will be called if it's determined that the PDF contains multiple bills
# The function first extracts the account number, amount, and usage from the beginning of the doc
# Then it finds the associated meter number for each account and updates the dictionary

def parse_CombinedPDF(file_name):
    accounts = {}
    with pdfplumber.open(file_name) as pdf: # open pdf file
        text = ''
        for page in pdf.pages: # write pdf file to txt file for ease of access
            text += page.extract_text(keep_blank_chars=True)
        output = open("output1.txt",'w')
        output.write(text)
        output = open('output1.txt', 'r')
        lines = output.readlines()
        finished = False
        for line in enumerate(lines): # iterate through lines of txt file
            if line[1] == 'ACCOUNT ACCOUNT NAME BILLING PERIOD USAGE CURRENT CHARGES\n' and finished == False: # this line appears right before the account information
                finished = True
                Count = 1
                next = lines[line[0]+Count]
                #print(next)
                while 'Total current charges by account' not in next: # this line appears at the end of the account information
                    Count += 1
                    next = lines[line[0]+Count] # iterate through the account information lines
                    if 'THM' in next: # this line contains the account number, amount, and usage information
                        thisline = next.strip().split(' ')
                        if thisline[-1] == 'THM': # alternative logic for messed up bills where THM gets pushed to next line
                            thisline = lines[line[0]+Count-2].strip().split(' ')
                            account_num = thisline[0]
                            amount = thisline[-1].replace(',','')
                            usage = thisline[-2]
                        else: # standard logic
                            if thisline[0].isnumeric() == True and len(thisline[0]) <= 2:
                                account_num = thisline[1]
                            else:
                                account_num = thisline[0]
                            amount = thisline[-1].replace(',','')
                            if '$' in amount: # cleans up amount string to be treated as a number
                                amount = amount[1:]
                            usage = thisline[-3]
                        accounts.update({account_num: [amount, usage]}) # updates account dictionary with account number as key and value[0] = amount, value[1] = usage
            #if 'Billing Period Current Reading' in line[1]: # this line appears at the beginning of the meter section
            if 'Account number Pressure' in line[1]:
                accountline = lines[line[0]+2].strip().split(' ') # line with account number
                account = accountline[0]
                #print('found account ' + account)
                meterline = lines[line[0]-1].strip().split(' ') # line with meter number
                meter = 'N' + str(meterline[-2][6:])
                data_dic.update({meter: [account, accounts.get(account)[0], accounts.get(account)[1]]}) # updates data dictionary with meter number as key and value[0] = account number, value[1] = amount, value[2] = usage
               # print('found meter ' + meter)
            


# This function will be called if it's determined that the PDF contains a single bill
# The function extracts the account number, meter number, amount, and usage from the doc, then updates data_dic
def parse_SinglePDF(file_name):
    with pdfplumber.open(file_name) as pdf: # open pdf file
        text = ''
        for page in pdf.pages: # write pdf file to txt file for ease of access
            text += page.extract_text(keep_blank_chars=True)
        output = open("output2.txt",'w')
        output.write(text)
        output = open('output2.txt', 'r')
        lines = output.readlines()
        for line in enumerate(lines): # iterate through lines of txt file
            if 'CUSTOMER ACCOUNT NUMBER' in line[1]:
                next = lines[line[0]+1].strip().split(' ')
                account = next[2] # account number line
            if 'natural gas to your home or business' in line[1]:
                meter = 'N' + str(line[1].strip().split(' ')[-2][6:]) # meter number line
            if 'Total Current Gas Charges' in line[1]:
                amount = line[1].strip().split(' ')[-1].replace(',','') # amount line
            if 'Demand - Charge for some larger customers' in line[1]:
                usage = line[1].strip().split(' ')[-2] # usage line
        if account and meter and amount and usage:
            data_dic.update({meter: [account, amount, usage]})
         

def parse(list_file_name, root):
    combinedBills = ['CNP-008000098669-Bill.pdf', 'CNP-008000098670-Bill.pdf', 'CNP-008000098470-Bill.pdf', 'CNP-008000098668-Bill.pdf', 'CombinedBill.pdf']
    for file_name in list_file_name:
        path = root + "/" + file_name
        if file_name in combinedBills:
            parse_CombinedPDF(path)
        else:
            parse_SinglePDF(path)


def convert_dic(old_dic):
    updated_dic = { 
    'Amount': [],
    'Meter': [],
    'Usage': [],}
    for key in old_dic.keys():
        updated_dic['Amount'].append(float(old_dic.get(key)[1]))
        updated_dic['Meter'].append(key)
        updated_dic['Usage'].append(float(old_dic.get(key)[2]))
    return(updated_dic)

def update_excel(workbook, new_values, month):
    specific_month = month
    sheet = workbook["Natural Gas"]
    
    for count, cell in enumerate(sheet[4]):
        if cell.value is not None:
            value = cell.value.split(" ")[0]
            if value == specific_month:
                index_amount = count - 1
                index_usage = count 
    
    for i, j in new_values.iterrows():
        for row in sheet.iter_rows(min_row=4):
            if row[4].value == j.iloc[1]:
                row[index_amount].value = j.iloc[0]
                row[index_usage].value = j.iloc[2]

    workbook.save('UpdatedOutput.xlsx')
    
    
if __name__ == "__main__":
    data_dic = {}
    parser = argparse.ArgumentParser()
    parser.add_argument("-f", help="PDF files to be processed")
    parser.add_argument("-m", help="month") 
    parser.add_argument("-e", help='existing excel file')
    args = parser.parse_args()
    files = os.listdir(args.f)
    parse(files, args.f)
    #print(data_dic)
    #print('There are ' + str(len(data_dic)) + ' entries in the data dictionary.')
    updated_dic = convert_dic(data_dic)
    new_values = pd.DataFrame(updated_dic)
    workbook = openpyxl.load_workbook(args.e)
    update_excel(workbook, new_values, args.m)
    #testfile = open('testfile.txt','w')
    #testfile.write('Meter\tAccount\tAmount\tUsage\n')
    #meter_list = data_dic.keys()
    #for key in data_dic.keys():
    #    testfile.write(str(key)+'\t'+str(data_dic.get(key)[0])+'\t'+str(data_dic.get(key)[1])+'\t'+str(data_dic.get(key)[2])+'\n')
    #testfile.write(meter_list)
    #testfile.close()
    print('Successfuly updated the spreadsheet')
