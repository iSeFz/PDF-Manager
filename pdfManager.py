# Description : Program to manage pdf files (merge, split and extract)
# Author : Seif Yahia
# Last Modified Date : 3 Feb. 2023
# Version : 1.5

import PyPDF2
import os


# Merge multiple pdf files into one file
def merge():
    mergedObject = PyPDF2.PdfMerger()
    nFiles = 1
    while(True):
        # Take input files from user
        pdf_file = input(f"Enter the path of pdf #{nFiles} (Press ENTER To Stop!): ").replace("\"", "")
        # Check if the file exist or not
        if(((os.path.exists(pdf_file) == False) or (os.path.isdir(pdf_file) == True)) and (pdf_file != "")):
            print("\t\tFile does NOT exist!")
            continue
        if(pdf_file == ""):
            # If the user pressed ENTER in the first time
            # Display the menu again then exit the program
            if(nFiles == 1):
                run()
                exit()
            break
        # Merge pdf files via the PdfMerger object
        mergedObject.append(PyPDF2.PdfReader(pdf_file))
        nFiles += 1
    outfile = input("Enter a NAME to the output merged PDF: ").replace(".pdf", "")
    # Check if the path entered is a valid directory or not
    while(True):
        outpath = input("Enter a PATH to the output merged PDF: ").replace("\"", "")
        if(os.path.isdir(outpath) == True):
            break
        else:
            print("\t\tPath does NOT exist!")
            continue
    # Change the directory of the running code
    # To the specified directory by the user
    os.chdir(outpath)
    mergedObject.write(outfile + ".pdf")
    mergedObject.close()
    print("\nFiles are merged successfully!!")
    print(f"Go check {outfile}.pdf at {os.getcwd()}")


# Extract a certain page from pdf and put it in a new pdf
def extract():
    global head
    # Take inputs needed from the user
    while(True):
        filepath = input("Enter the PATH of the input PDF: ").replace("\"", "")
        # Check if the entered path is a valid file but NOT a directory
        if((os.path.exists(filepath) == True) and (os.path.isdir(filepath)) == False):
            break
        else:
            print("\t\tFile does NOT exist!")
            continue
    print("\tExtracted page will be named as \"fileName-pageNumber.pdf\"")
    filename = input("Enter the file NAME: ").replace(".pdf", "")
    # Get the number of pages in the entered file
    reader_pdf = PyPDF2.PdfReader(filepath)
    n_pages = len(reader_pdf.pages)
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
    writer_pdf = PyPDF2.PdfWriter()
    writer_pdf.add_page(reader_pdf.pages[numpage - 1])
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
    # Take inputs needed from the user
    while(True):
        filepath = input("Enter the PATH of the input PDF: ").replace("\"", "")
        # Check if the entered path is a valid file but NOT a directory
        if((os.path.exists(filepath) == True) and (os.path.isdir(filepath)) == False):
            break
        else:
            print("\t\tFile does NOT exist!")
    print("\tSeparated pages will be named as \"templateName-pageNumber.pdf\"")
    filename = input("Enter the template NAME: ").replace(".pdf", "")
    reader_pdf = PyPDF2.PdfReader(filepath)
    # Check if the new path already exists
    newPath = os.path.join(filepath.replace(".pdf", ""), "")
    if(os.path.exists(newPath) == True):
        print(f"\tDirectory {newPath} already exists!")
        # Take another new directory from the user
        while(True):
            newPath = input("Enter a NEW path for the output directory: ").replace("\"", "")
            # Check if the entered path is a valid NEW directory
            if(os.path.exists(newPath) == False):
                break
            else:
                print(f"\tDirectory {newPath} already exists!")
    # Create a folder with the new directory to write files
    os.mkdir(newPath)
    os.chdir(newPath)
    # Loop over all pages in the selected file
    for page in range(len(reader_pdf.pages)):
        writer_pdf = PyPDF2.PdfWriter()
        writer_pdf.add_page(reader_pdf.pages[page])
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
