## MF-Convert
Mainframe-generated PDF to TXT converter.

A first attempt at Git / writing a Python script from scratch

### Purpose:
I work with reports that are generated from a mainframe system that writes to a PDF by default.  Previous attempts at using the PDF data via copy/paste were burdensome because consecutive spaces would convert to a single space.  This makes it much more difficult to import/manipulate the data.

There are certainly easier and more elegant ways to approach this task, and any feedback is welcome.

### Pseudocode:
* Find all pdf files in the CWD and assign them an index 0, 1, ... n-1
* Prompt user to choose a file to convert to text, or to quit
* (PyPDF2) Open file, remove the blank password, then write all the pages to a new pdf file and save it as temp_pdf.pdf
* Prompt user for name of txt output file
* (pdfreader) Cycle through each page of temp_pdf and write each line to the txt output file.  Each line is encoded ('ascii', 'replace') and decoded ('ascii', 'strict') before being written to the file to address unknown characters.
    * Every 10 pages, print a message stating which page is currently being processed.
* Delete temp_pdf.pdf, close all other files.
* Print a save confirmation and the amount of time the conversion took.

### To do:
* Find a better way to deal with unknown characters
* Add option to convert to txt, csv (comma or pipe delimited)
* Bulk conversion
* Check that the txt file won't overwrite an existing one
* After converting a file, check if user wants to process another one, instead of quitting
* Look into command prompt options
