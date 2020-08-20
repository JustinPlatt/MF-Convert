import os
from pathlib import Path
from pdfreader import PDFDocument, SimplePDFViewer
import PyPDF2
import sys
import time


def is_valid(choice, file_ct):
    """Check if the report choice is a valid one"""
    x = False
    if choice.isnumeric():
        if int(choice) in range(file_ct):
            x = True
    return x

def main():
    """Docstring will go here"""
    print('*'*40 +'\nERT PDF TO TXT CONVERTER\n')
    print('Showing PDF files in ' + os.getcwd())
    files = []

    f_count = 0
    for file in os.listdir():
        if file.endswith(".pdf") or file.endswith(".PDF"):
            f_name = file.rsplit('.', maxsplit=1)[0]
            print('(' + str(f_count) + ')  ' + file)
            files.append(f_name) #take the file name w/o the .pdf
            f_count +=1

    prompt = 'Enter the number corresponding to the target pdf, or q to quit: '
    choice = input(prompt)
    while is_valid(choice,f_count) == False:
        if choice == 'q':
            print('Quitting.')
            sys.exit()
        else:
            print('Invalid choice - try again.')
            choice = input(prompt)
    if choice !='q':
        pdf_to_open = str(files[int(choice)]) + '.pdf'
        print(pdf_to_open)

    #open the protected pdf and remove the password
    print('Converting to unlocked PDF')
    p_pdf = open(pdf_to_open, 'rb')  #this will change
    pdfReader = PyPDF2.PdfFileReader(p_pdf)
    pdfReader.decrypt('')
    pdfWriter = PyPDF2.PdfFileWriter()

    #write all the pages in the unlocked file to a new pdf
    for pageNum in range(pdfReader.numPages):
        pageObj = pdfReader.getPage(pageNum)
        pdfWriter.addPage(pageObj)

    u_name = input('Output file name? (excluding ".txt"):  ')  #prompt for the name of the new unlocked pdf
    #print(Path(u_name + '.txt').is_file())
    #sys.exit()

    pdfOutputFile = open('temp_pdf.pdf', 'wb')
    pdfWriter.write(pdfOutputFile)

    pdfOutputFile.close()
    p_pdf.close()

    #open the unlocked pdf.  We'll use PDFDocument and SimplePDFViewer to pull all the text
    u_pdf = open('temp_pdf.pdf', "rb")
    doc = PDFDocument(u_pdf)
    reader = SimplePDFViewer(u_pdf)
    start_time = time.time()
    pgs = [p for p in doc.pages()]  #count number of pages
    page_ct = len(pgs)
    print('Writing ' + str(page_ct) + ' pages to ' + u_name + '.txt ...')

    with open(u_name + '.txt', 'w') as g:
        for pg in range(page_ct): #cycle through pages
            reader.navigate(pg+1)
            reader.render()
            if (pg+1)%10 == 0:
                print('Processing page ' + str(pg+1) + ' of ' + str(page_ct))
            st = reader.canvas.strings #list with 1 line per element
            for l in range(len(st)):
                ln = st[l].encode('ascii', 'replace') #turn unknown chars into ?
                ln = ln.decode('ascii', 'strict')
                g.write(ln+'\n')

    u_pdf.close()
    os.remove('temp_pdf.pdf')

    print('Saved as ' + u_name + '.txt')
    print("This took %s seconds." % round((time.time() - start_time),2))

if __name__ == "__main__":
	main()
