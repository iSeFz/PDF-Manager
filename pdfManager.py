# Description : Program to manage pdf files (merge, split and extract)
# Author : Seif Yahia
# Last Modified Date : 2 Feb. 2023

import PyPDF2
import os


# Merge multiple pdf files into one file
def merge():
    mergedObject = PyPDF2.PdfFileMerger()
    nFiles = 1
    while(True):
        # Take input files from user
        pdf_file = input(f"Enter the path of pdf #{nFiles} (Press ENTER To Stop!): ").replace("\"", "")
        if(pdf_file == ""):
            # If the user pressed ENTER in the first time
            # Display the menu again then exit the program
            if(nFiles == 1):
                run()
                exit()
            break
        # Merge pdf files via the PdfFileMerger object
        mergedObject.append(PyPDF2.PdfFileReader(pdf_file, 'rb'))
        nFiles += 1
    outfile = input("Enter a NAME to the output merged PDF: ").replace(".pdf", "")
    o_path = input("Enter a PATH to the output merged PDF: ").replace("\"", "")
    # Change the directory of the running code
    # To the specified directory by the user
    os.chdir(o_path)
    mergedObject.write(outfile + ".pdf")
    mergedObject.close()
    print("\nFiles are merged successfully!!")
    print(f"Go check {outfile}.pdf at {os.getcwd()}")


# Extract a certain page from pdf and put it in a new pdf
def extract():
    global head
    # Take inputs needed from the user
    filepath = input("Enter the PATH of the input PDF: ").replace("\"", "")
    print("\tExtracted page will be named as \"fileName-pageNumber.pdf\"")
    filename = input("Enter the file NAME: ").replace(".pdf", "")
    # Get the number of pages in the entered file
    reader_pdf = PyPDF2.PdfFileReader(filepath)
    n_pages = reader_pdf.getNumPages()
    # Check if the input of page number is correct or not
    while(True):
        numpage = input("Enter the page number to extract: ")
        if(numpage.isdigit()):
            # Convert the number into an integer
            numpage = int(numpage)
            if(numpage == 0):
                print("\tPaginating starts from 1, for the first page enter 1.")
            # If the page number is valid
            elif(numpage <= n_pages):
                break
            else:
                print("\tINVALID page number!!")
        # If the input is not a number
        else:
            print("\tINVALID INPUT!! Enter ONLY numbers.")
    # Create the pdf and add the page to the writer object
    writer_pdf = PyPDF2.PdfFileWriter()
    writer_pdf.addPage(reader_pdf.getPage(numpage - 1))
    # Change the directory to the same file directory
    # By removing the filename itself
    head, tail = os.path.split(filepath)
    os.chdir(filepath.replace(tail, ""))
    # Write the selected page to a new file named 'filename-numpage.pdf'
    with open(f'{filename}-{numpage}.pdf', 'wb') as f:
        writer_pdf.write(f)
    print("\nPage extracted from PDF successfully!!")
    print(f"Go check {filename}-{numpage}.pdf at {os.getcwd()}")


# Split the selected pdf in to separate files
# Each with one page from the original pdf
def separate():
    # Take input names from the user
    filepath = input("Enter the PATH of the input PDF: ").replace("\"", "")
    print("\tSeparated pages will be named as \"templateName-pageNumber.pdf\"")
    filename = input("Enter the template NAME: ").replace(".pdf", "")
    reader_pdf = PyPDF2.PdfFileReader(filepath)
    # Get number of pages of the input file
    page_num = reader_pdf.getNumPages()
    # Create a new folder/directory to write files
    newPath = os.path.join(filepath.replace(".pdf", ""), "")
    os.mkdir(newPath)
    os.chdir(newPath)
    # Loop over all pages in the selected file
    for page in range(page_num):
        writer_pdf = PyPDF2.PdfFileWriter()
        writer_pdf.addPage(reader_pdf.getPage(page))
        # Write every page in a separate pdf
        # named 'filename-current_page_number.pdf'
        with open(f'{filename}-{page + 1}.pdf', 'wb') as f:
            writer_pdf.write(f)
    print("\nPDF separated into pages successfully!!")
    print(f"Go check the files at {newPath}")


# Main function to run the program
def run():
    while True:
        print('''\nThe program has the following features:
(1) Merge multiple pdf files
(2) Extract a single page from pdf
(3) Split pdf into separate pages
(4) Exit''')
        choice = input("Choose one of the options above (1-4) >> ")
        if(choice == '1'):
            merge()
        elif(choice == '2'):
            extract()
        elif(choice == '3'):
            separate()
        elif(choice == '4'):
            print("\tThank you for using the PDF manager!")
            break
        else:
            print("\n\t\tINVALID INPUT!!")


# Greeting message then the program starts
print("\tWelcome to the PDF manager program!", end = '')
run()
