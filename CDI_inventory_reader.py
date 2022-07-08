
import PyPDF2
import re
from csv import DictWriter
import csv
import os


##! NEED TO PULL LICENSE AS WELL 
""" Need to know if it is the whole lot or not .If it is not the whole lot, then they (juan) needs to go into the lot and get the individual licenses and put the pallets not on the order on hold, so they dont get pulled. If it is the whole """

#! This is the directory where the inventory file is stored 
inventory_directory = "C:\\Users\\Chris R\\Desktop\\Python Projects\\Mini_Projects\\cdi_reader"  #! Use this at home
##inventory_directory = "C:\\Users\\crodea\\Desktop\\\Mini_Projects\\cdi_reader"  #! Use this at work 
#inventory_directory = "C:\\Users\\crodea\\Desktop\\OneDrive\\OneDrive - US Cold Storage\\Python\\CDI_allocations_project" #! One drive folder
#print(file_directory)

#!This is the directory that holds the allocation PDF's 
#path = "C:\\Users\\crodea\\Desktop\\test_allocations"  #! This is the path for work desktop
path = "C:\\Users\\Chris R\\Desktop\\Python Projects\\Mini_Projects\\cdi_reader\\test_allocations"   #! This is the path at home 


#! Open Up file and get the contents of the first page 
# def get_text(cdi_pdf):
#     pdfFileObj = open('CDI1.pdf', 'rb')
#     pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
#     pageObj = pdfReader.getPage(0)
#     text = pageObj.extractText()
#     return text

#! This will get the total number of pages for each pdf. 
def get_page_cnt(pdf): #! example of possible input ==> ['CDI0.pdf', 'CDI1.pdf', 'CDI2.pdf']
    invoice = pdf
    pdfFileObj = open(f'{invoice}', 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    pages = pdfReader.getNumPages()
    return pages


#! This function will find the order Number located in the PDF and store it in a variable to be used later as the output csv file name
def getOrderNumber(pdf):
    pattern = re.compile("SO000\\d\\d\\d\\d\\d\\d", re.IGNORECASE)
    match = pattern.findall(pdf)
    return match[0] 

#! Gets the pallet ID from the CDI allocation PDF file that is opened 
def getPalletId(cdi_pallets):
    pattern = re.compile("\\d\\d\\d\\d-\\d\\d\\d-\\d\\d-[a-zA-Z]\\d\\d-[a-zA-Z]", re.IGNORECASE)
    match = pattern.findall(cdi_pallets)
    return match #! Returns a list of pallet ID's ==> ['3822-153-08-S04-B', '6622-153-13-S02-F', '6622-153-13-S03-C', '6622-153-13-S04-B']

#! Gets the quantity that is requested on the CDI order manifest. *NOT* the quantity in inventory
def getQTY(input):
    pattern = re.compile("(\d+\.0000\s+CS)", re.MULTILINE)
    match = pattern.findall(input) #! Returns ['84.0000\n\n\nCS', '84.0000\n\n\nCS', '84.0000\nCS']
    
    qty = []
    for items in match:
        strip_new_line = items.split("\n") #! Returns ['84.0000', '', '', 'CS']
        add_CS = strip_new_line[0] + " CS" #! Returns 84.0000 CS
        
        qty.append(add_CS)
    
    return qty

#! Renames all the cdi pdfs in the folder 'test_allocations' so that we can loop through them 
def rename_files():
    import os

    #! Finds the path to the folder that holds the allocation PDF's 
    #path = "C:\\Users\\crodea\\Desktop\\test_allocations"  #! This is the path at work 
    #path = "C:\\Users\\Chris R\\Desktop\\Python Projects\\Mini_Projects\\cdi_reader\\test_allocations"   #! This is the path at home 

    #! Changes the directory to the path with the files 
    os.chdir(path)

    #! Iterates through files and renames them CDI0,CDI1,CDI2,etc. 
    for count, f in enumerate(os.listdir()):
        f_name, f_ext = os.path.splitext(f)
        f_name = "CDI" + str(count)

        new_name = f'{f_name}{f_ext}'
        os.rename(f, new_name)

    #! Create a combined list of file names that we can interate through
    file_list = []
    files = os.listdir()    
    for items in files:
        #print(items)
        file_list.append(items)

    return file_list

#! Opens the inventory excel sheet (8-3-1 report in WMS). looks up each pallet ID that is in the PID list
#! found in the getPalletId() function
def read_from_inventory_csv(number, qty):
    #csv_file = csv.reader(open("inventory.csv", "r"), delimiter=",") #! need to open the file within the for loop, or else it will not restart the search from top of file
    row_info = []
    os.chdir(inventory_directory) #! Change back from the pdf folder.
    
    for items in range(len(number)):
        csv_file = csv.reader(open("inventory.csv", "r"), delimiter=",")
        palletID = number[items]
        qty_requested = qty[items]
        #print(number)
        add_comma = "'" + str(palletID) #! Turns 3822-150-05-S01-D into '3822-150-05-S01-D . Because the inventory csv adds a comma in front of PID
        #print(add_comma)
        for row in csv_file:
            #print(number[items])
            if add_comma == row[13]:        
                #print(row)
                info = {
                    "Product Code": " ".join(row[2].split()),
                    "Lot Number": " ".join(row[1].split()), #! Strips the extra spaces 
                    "Quantity Requested From CDI": qty_requested, 
                    "Quantity in INV": " ".join(row[8].split()),
                    "Batch ID": palletID,
                    "Pallet ID Matched": " ".join(row[13].split()),
                    "License Number": " ".join(row[12].split())
                }

                row_info.append(info) #! Adds to the row_info list, which the function will return later. 
                break
            else:
                """Need to create a template for product that doesn't have a match """
                #print(str(palletID) + " No match " + str(row[13]))
                continue
        continue

    return row_info




def main():
    pdfs = rename_files()
    print(pdfs)

    for pdf in pdfs:
        os.chdir(path) #! Need to change back to the folder where the pdf's are because the 'read_from_inventory_csv function changes the path to 
        print(pdf)
        invoice = pdf
        pdfFileObj = open(f'{invoice}', 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
        pages = pdfReader.getNumPages()
        pageObj = pdfReader.getPage(0)
        text = pageObj.extractText()    

        orderNumber = getOrderNumber(text)

        #print(pages)
        myPalletList = [] #! Stores all the pallet ID's from each page of CDI pdf. 

        for page in range(pages):
            pageObj = pdfReader.getPage(page)
            text = pageObj.extractText()
            PID = getPalletId(text) #! Gets all the Pallet ID on Manifest
            QTY_REQUESTED = getQTY(text) #! Gets all the qty requested on Manifest
            if len(PID) == 0: #! If any page is empty or has no pallet ID, program will move on to next page
                continue
            else:
                for palletID in PID: #! Adds the pallet id's 1 by 1 instead of as a whole so multiple pages, wont build a list of lists
                    myPalletList.append(palletID)
                #print(myPalletList)
        
        matched_pallets = read_from_inventory_csv(myPalletList, QTY_REQUESTED)
        
        with open(f'{orderNumber}.csv','w', newline='') as outfile:
                writer = DictWriter(outfile, ('Product Code','Lot Number','Quantity Requested From CDI','Quantity in INV', 'Batch ID', 'Pallet ID Matched', 'License Number'))
                writer.writeheader()
                writer.writerows(matched_pallets)
            
    

main()
