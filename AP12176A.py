# We import NumPy into Python
import numpy as np
import pandas as pd
import datetime
import re
import codecs
import pdfreader
from pdfreader import PDFDocument, SimplePDFViewer
import PyPDF2


def is_inv(z):
    """Check if a string is an invoice"""
    return re.search('^\s\d{2}\/\d{2}\/\d{2}\s+\S+\s+\d{2}\/\d{2}\/\d{2}.+$',str(z))

#open the protected pdf and remove the password
def main():
    print('Opening pdf and writing to decrypted copy')
    p_pdf = open('AP12176A_20200701_142127.pdf', 'rb')  #this will change
    pdfReader = PyPDF2.PdfFileReader(p_pdf)
    pdfWriter = PyPDF2.PdfFileWriter()
    pdfReader.decrypt('')

    #write all the pages in the unlocked file to a new pdf
    print('Writing to decrypted copy.')
    for pageNum in range(pdfReader.numPages):
        pageObj = pdfReader.getPage(pageNum)
        pdfWriter.addPage(pageObj)

    u_name = input('File name for decrypted copy? (excluding ".pdf"):  ')  #prompt for the name of the .txt output file
    u_file = u_name + '.pdf'

    #pdfOutputFile = open(u_file, 'wb')
    print('Finishing writing to decrypted copy')
    pdfOutputFile = open('temp_pdf.pdf', 'wb')
    pdfWriter.write(pdfOutputFile) #write to the temporary unlocked pdf

    pdfOutputFile.close()
    p_pdf.close()

    #open the unlocked pdf.  We'll use PDFDocument and SimplePDFViewer to pull all the text
    #_pdf = open(u_file, "rb")
    u_pdf = open('temp_pdf.pdf', "rb")
    doc = PDFDocument(u_pdf)
    reader = SimplePDFViewer(u_pdf)

    print('counting pages in pdf')
    pgs = [p for p in doc.pages()]  #count number of pages
    page_ct = len(pgs)

    print('cycling through pages')
    with open(u_name + '.txt', 'w') as g:
        for pg in range(page_ct): #cycle through pages
            reader.navigate(pg+1)
            reader.render()
            if (pg+1)%10 == 0:
                print('processing page ' + str(pg+1) + ' of ' + str(page_ct))
            st = reader.canvas.strings #list with 1 line per element
            for l in range(len(st)):
                ln = st[l].encode('ascii', 'replace').decode('utf-8') #turn unknown chars into ?
                g.write(ln+'\n')

    group_exp = '^\s(.{8})\s(.{17})\s\s(.{8})\s\s(.{10})\s\s(.{21})\s(.{10})\s(.{18})\s(.{12})(.+)?$' #regex for grouping an invoice..
    group_inv = re.compile(group_exp) #set as a regex
    acct_exp ='ACCOUNT.+(\d{6}).+DEPARTMENT\s+(\w{8})' #find acct and center
    find_acct = re.compile(acct_exp)
    vend_exp = '^([\w]{10})\s+(\S.+\S)\s+$' #find 10 'word' chars, then a space, then everything up to the line break
    find_vendor = re.compile(vend_exp)
    corp_exp = 'COMPANY\s(\w{4})\s+DATE' #find the 4 characters between COMPANY and DATE
    find_corp = re.compile(corp_exp)

    corp = 'XXXX'
    center = 'XXXXXXXX'
    account = 'XXXXXX'
    v_short = 'XXXX'
    v_long = 'XXXXXXXX'

    data_tmp=[]

    ct = 0

    with open(u_name + '.txt', 'r') as h:
        for line in h:
            ln = str(line) #remove leading/trailing spaces and newline chars
            if find_corp.search(ln):  #look for a new corp
                if corp != find_corp.search(ln).group(1):
                    corp = find_corp.search(ln).group(1)
            elif find_vendor.search(ln): #look for a new vendor
                if v_short != str(find_vendor.search(ln).group(1)):
                    v_short = str(find_vendor.search(ln).group(1))
                    v_long = str(find_vendor.search(ln).group(2))
            elif find_acct.search(ln):  #look for a new acct/center
                if center != find_acct.search(ln).group(2) or account != find_acct.search(ln).group(1):
                    center = find_acct.search(ln).group(2)
                    account = find_acct.search(ln).group(1)
            elif is_inv(ln):  #look for an invoice
                tmp = group_inv.search(ln) #print(is_inv(ln).groups())
                gl_eff = tmp.group(1).strip()
                inv_num = tmp.group(2).strip()
                inv_date = tmp.group(3).strip()
                po = tmp.group(4).strip()
                desc = tmp.group(5).strip()
                q = tmp.group(6).strip()
                if q == '':
                    qty = 0
                else: #qty
                    qty = q.replace(',','')
                prod_id = tmp.group(7).strip()
                if len(tmp.group(8).strip()) != 0 : #expense
                    exp = tmp.group(8).strip()
                    exp = float(str(exp).replace(',',''))
                else:
                    exp = 0
                if len(tmp.group(9).strip()) != 0 : #expense
                    cred = tmp.group(9).strip()
                    cred = float(str(cred).replace(',',''))
                else:
                    cred = 0
                if int(qty) == 0:
                    per_unit = 0
                else:
                    per_unit = round(float(exp)/float(qty),2)
                new_row = [corp, center, account, v_short, v_long, gl_eff, inv_num, inv_date, po, desc, qty, prod_id, exp, cred, per_unit]
                data_tmp.append(new_row)
                ct += 1
                if ct % 1000 == 0:
                    print('Finished adding row ' + str(ct))

    data_cols = ['Company', 'Center', 'Account', 'Vendor_Short', 'Vendor_Long',	'GL_Effective_Date', 'Inv_Number', 'Inv_Date', 'PO', 'Description',	'Qty', 'ProdID', 'Expense', 'Credit', 'Per_Unit_Cost']
    col_widths=[14,12,12,18,35,22,20,13,12,31,10,20,15,15,18]
    data_inv=pd.DataFrame(data=data_tmp, columns = data_cols)
    data_inv['GL_Effective_Date']=pd.to_datetime(data_inv['GL_Effective_Date'])
    data_inv['Inv_Date']=pd.to_datetime(data_inv['Inv_Date'])
    data_inv['Account']=data_inv['Account'].astype('int64')
    data_inv['Qty']=data_inv['Qty'].astype('int64')
    i_rows = data_inv['Company'].size

    with pd.ExcelWriter(u_name + '.xlsx', engine = 'xlsxwriter', datetime_format = 'm/d/yyyy') as writer:
        data_inv.to_excel(writer, sheet_name='DATA', index = False )
        workbook  = writer.book
        worksheet = writer.sheets['DATA']
        curr_format = workbook.add_format({'num_format': '$#,##0.00;[Red]($#,##0.00)'})
        worksheet.set_column(12, 12, 13, curr_format) #first col, last col, width, format
        worksheet.set_column(13, 13, 13, curr_format)
        worksheet.autofilter('A1:O' + str(i_rows+1))
        worksheet.freeze_panes(1,0) #freeze 1st row
        for a in range(len(col_widths)):
            worksheet.set_column(a,a,col_widths[a])

    print('DONE - pulled ' + str(i_rows) + ' lines into ' + u_name + '.xlsx')



if __name__ == "__main__":
	main()
