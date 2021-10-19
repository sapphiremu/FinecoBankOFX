# Fineco excel statement to OFX convertor
# Author:  Saffron Murcia
# Date:    2021-10-18
# Updated: 2021-10-19
# Version: 0.4.1
#
# Change Log:
# 0.1 Moved output line creation into functions
# 0.2 Adapted program to detect euro accounts
# 0.3 Program now prompts for filename with a dialog box
# 0.4 Program no longer leaves an empty tkinter window
# 0.4.1 File dialog box filters to XLS files
#
# Known bugs or concerns:
# 1. No testing is done on the layout of the source file.
#    If Fineco change anything, it may produce unintended outcomes.
# 2. The date/time of transactions conversion is "dirty"
#    Improvement is certainly needed.



def fileHeader(currency,accountNumber):
    message="OFXHEADER:100\nDATA:OFXSGML\nVERSION:102\nSECURITY:NONE\nENCODING:USASCII\nCHARSET:1252\nCOMPRESSION:NONE\nOLDFILEUID:NONE\nNEWFILEUID:NONE\n"
    message=message+("<OFX>\n")
    message=message+("<SIGNONMSGSRSV1>\n")
    message=message+("<SONRS>\n")
    message=message+("<STATUS>\n")
    message=message+("<CODE>0</CODE>\n")
    message=message+("<SEVERITY>INFO</SEVERITY>\n")
    message=message+("</STATUS>\n")
    message=message+("<DTSERVER></DTSERVER>\n")
    message=message+("<LANGUAGE>ENG</LANGUAGE>\n")
    message=message+("<INTU.BID></INTU.BID>\n")
    message=message+("</SONRS>\n")
    message=message+("</SIGNONMSGSRSV1>\n")
    message=message+("<BANKMSGSRSV1>\n")
    message=message+("<STMTTRNRS>\n")
    message=message+("<TRNUID>1</TRNUID>\n")
    message=message+("<STATUS>\n")
    message=message+("<CODE>0</CODE>\n")
    message=message+("<SEVERITY>INFO</SEVERITY>\n")
    message=message+("</STATUS>\n")
    message=message+("<STMTRS>\n")
    message=message+("<CURDEF>"+currency+"</CURDEF>\n")
    message=message+("<BANKACCTFROM>\n")
    message=message+("<BANKID></BANKID>\n")
    message=message+("<ACCTID>"+accountNumber+"</ACCTID>\n")
    message=message+("<ACCTTYPE>SAVINGS</ACCTTYPE>\n")
    message=message+("</BANKACCTFROM>\n")
    return(message)

def transaction(trnDate,trnAmount,trnDesc):
    uniqueID = str(int(hashlib.sha1((trnDate+trnAmount+trnDesc).encode("utf-8")).hexdigest(), 16)%(10 ** 31))
    message="\n"
    message=message+("<STMTTRN>\n")
    message=message+("<TRNTYPE>OTHER</TRNTYPE>\n")

    message=message+("<DTPOSTED>"+trnDate+"000000</DTPOSTED>\n")
    message=message+("<TRNAMT>"+trnAmount+"</TRNAMT>\n")
    message=message+("<FITID>"+uniqueID+"</FITID>\n")

    message=message+("<NAME>"+trnDesc+"</NAME>\n")
    message=message+("<MEMO></MEMO>\n")
    message=message+("</STMTTRN>\n")
    return (message)

def transactionPrep():
    message=("<BANKTRANLIST>\n<DTSTART></DTSTART>\n<DTEND></DTEND>\n")
    return (message)

def fileFooter():
    message=("</BANKTRANLIST>\n<LEDGERBAL>\n<BALAMT></BALAMT>\n<DTASOF></DTASOF>\n</LEDGERBAL>\n</STMTRS>\n</STMTTRNRS>\n</BANKMSGSRSV1>\n</OFX>\n")
    return (message)


# main program

import tempfile
import pandas as pd
import hashlib
from tkinter import filedialog as fd
from tkinter import Tk
import subprocess

root = Tk()
root.withdraw()
fileName = fd.askopenfilename(filetypes=[("Fineco Excel Statement", "*.xls")])
if len(fileName)!=0:

    outputFile = tempfile.gettempdir()+"\working.ofx"

    of = open(outputFile,"w")


    df = pd.read_excel(fileName,header=None)
    myArray=df.to_numpy()
    if "EUR" in myArray[0][0]:
        print("Found EURO account")
        currency="EUR"
    else:
        print("No or unknwon currency specified in input file.")
        print("*** USE EXTREME CAUTION AND CHECK OUTPUT FOR SANITY BEFORE IMPORTING ***")
        currency=input("Enter currency symbol: ")


    if currency=="EUR":
        accountNumber = str(myArray[0][0][-12:])
    elif currency=="GBP":
        accountNumber = str(myArray[0][0][-7:])
    else:
        exit(1)
        print("The program should have terminated with an error.")
        print("If you are reading this, then we are on unknown grounds")
        print("with the file format, layout or currency symbol.")
        print("\nIT IS HIGHLY RECOMMENDED TO TERMINATE NOW")
        print("BAD THINGS MAY HAPPEN IF YOU TRY TO USE THE GENERATED FILE!")
        input()
    of.write(fileHeader(currency,accountNumber))




    of.write(transactionPrep())


    transArray = myArray[6:]
    for x in transArray:

        if x[2]>0:
            amount=str(x[2])
        else:
            amount="-"+str(x[3])
        if currency=="EUR":
            hackedDate = (x[0][-4:]+x[0][3:5]+x[0][0:2])
        elif currency=="GBP":
            b=str(x[0])
            hackedDate = b[0:4]+b[5:7]+b[8:10]
        desc = x[4]
        
        of.write(transaction(hackedDate,amount,desc))

    of.write (fileFooter())
    of.close()
    subprocess.call(['explorer.exe',outputFile])
else:
    print("Nothing to do or program error.")
root.destroy()
