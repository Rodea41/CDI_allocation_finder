
import PyPDF2
import re
from csv import DictWriter
import csv
import os

#! This is the directory where the inventory file is stored 
inventory_directory = "C:\\Users\\Chris R\\Desktop\\Python Projects\\Mini_Projects\\cdi_reader"
#print(file_directory)

#!This is the directory that holds the allocation PDF's 
#path = "C:\\Users\\crodea\\Desktop\\test_allocations"  #! This is the path for work desktop
path = "C:\\Users\\Chris R\\Desktop\\Python Projects\\Mini_Projects\\cdi_reader\\test_allocations"   #! This is the path at home 


#! Open Up file and get the contents of the first page 
def get_text(cdi_pdf):
    pdfFileObj = open('CDI1.pdf', 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    pageObj = pdfReader.getPage(0)
    text = pageObj.extractText()
    return text

#! This will get the total number of pages for each pdf. 
def get_page_cnt(pdf): #! example of possible input ==> ['CDI0.pdf', 'CDI1.pdf', 'CDI2.pdf']
    invoice = pdf
    pdfFileObj = open(f'{invoice}', 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    pages = pdfReader.getNumPages()
    return pages


#! This function will find the order Number located in the PDF and store it in a variable to be used later as the output csv file name
def getOrderNumber():
    pass

#! Gets the pallet ID from the CDI allocation PDF file that is opened 
def getPalletId(cdi_pallets):
    pattern = re.compile("\\d\\d\\d\\d-\\d\\d\\d-\\d\\d-[a-zA-Z]\\d\\d-[a-zA-Z]", re.IGNORECASE)
    match = pattern.findall(cdi_pallets)
    return match #! Returns a list of pallet ID's ==> ['3822-153-08-S04-B', '6622-153-13-S02-F', '6622-153-13-S03-C', '6622-153-13-S04-B']

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


#PID = getPalletId(text)
#! Opens the inventory excel sheet (8-3-1 report in WMS). looks up each pallet ID that is in the PID list
#! found in the getPalletId() function
def read_from_inventory_csv(number):
    #csv_file = csv.reader(open("inventory.csv", "r"), delimiter=",") #! need to open the file within the for loop, or else it will not restart the search from top of file
    row_info = []
    os.chdir(inventory_directory) #! Change back from the pdf folder.
    
    for items in range(len(number)):
        csv_file = csv.reader(open("inventory.csv", "r"), delimiter=",")
        palletID = number[items]
        add_comma = "'" + str(palletID) #! Turns 3822-150-05-S01-D into '3822-150-05-S01-D . Because the inventory csv adds a comma in front of PID
        #print(add_comma)
        for row in csv_file:
            #print(number[items])
            if add_comma == row[13]:        
                #print(row)
                info = {
                    "Product Code": " ".join(row[2].split()),
                    "Lot Number": " ".join(row[1].split()), #! Strips the extra spaces 
                    "Quantity": " ".join(row[8].split()),
                    "Batch ID": palletID,
                    "Pallet ID Matched": " ".join(row[13].split()),
                }

                row_info.append(info) #! Adds to the row_info list, which the function will return later. 
                break
            else:
                #print(str(palletID) + " No match " + str(row[13]))
                continue
        continue

    #rm_extra_spaces = " ".join(palletInfo.split())
    #convert_to_list = rm_extra_spaces.split(" ")
    #print(row_info)
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
        #print(pages)
        myPalletList = [] #! Stores all the pallet ID's from each page of CDI pdf. 

        for page in range(pages):
            pageObj = pdfReader.getPage(page)
            text = pageObj.extractText()
            PID = getPalletId(text)
            if len(PID) == 0: #! If any page is empty or has no pallet ID, program will move on to next page
                continue
            else:
                for palletID in PID: #! Adds the pallet id's 1 by 1 instead of as a whole so multiple pages, wont build a list of lists
                    myPalletList.append(palletID)
                #print(myPalletList)
        
        matched_pallets = read_from_inventory_csv(myPalletList)
        
        with open(f'{pdf}.csv','w', newline='') as outfile:
                writer = DictWriter(outfile, ('Product Code','Lot Number','Quantity', 'Batch ID', 'Pallet ID Matched'))
                writer.writeheader()
                writer.writerows(matched_pallets)
            
        #print(myPalletList)
        #print(pdf)
                
        #!!!!! THIS IS WHERE YOU LEFT OFF 
        #read_from_inventory_csv(myPalletList)    
            
            # print(pdf)
            # print("Page: " + str(page))
            # print(PID)

            
    # myPalletList = []
    # #print(PID)
    # #print(inventory)
    # print(pdfs[1])
    # #print(os.getcwd()){pdfs[1]}

    # with open(f'{pdfs[1]}.csv','w') as outfile:
    #     writer = DictWriter(outfile, ('Product Code','Lot Number','Quantity', 'Batch ID', 'Pallet ID Matched'))
    #     writer.writeheader()
    #     writer.writerows(inventory)
        
    

main()
