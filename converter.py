# -*- coding: utf-8 -*-
"""
Created on Mon Apr 16 10:23:52 2018
@author: Kamal Zakieldin
"""
from subprocess import Popen, PIPE
import docx
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
import io
import sys
import os
import comtypes.client
import os.path
    

def convert_pdf_to_txt(path,fname):
    rsrcmgr = PDFResourceManager()
    retstr = io.StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = open(fname, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos = set()

    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password, caching=caching,check_extractable=True):
        interpreter.process_page(page)

    text = retstr.getvalue()

    fp.close()
    device.close()
    retstr.close()
    print(text)
    updated_string, updated_list = process_on_text(text)
    print(updated_string)
    print(updated_list)
    return updated_string

def convert_docx_to_Text(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        txt = para.text.encode('ascii', 'ignore')
        newtxt = txt.decode('unicode_escape')
        fullText.append(newtxt)
    mytxt = '\n'.join(fullText)
    print(mytxt)
    updated_string, updated_list = process_on_text(mytxt)
    print(updated_string)
    print(updated_list)
    return updated_string

def process_on_text(fulltext):
    print(fulltext)
    newlist = []
    newstring = ""
    newline = ""
    for newline in fulltext.splitlines():
            #print(newtxt)
        if newline == '':
            newline =  "\n"   # seperate lines
        newlist.append(newline)
        newstring += newline
    return newstring,newlist

def get_alldocs_to_text(filename, file_path):
    #normal doc / docx / odt / pdf / rtf
    if filename[-5:] == ".docx":
        return convert_docx_to_Text(filename)
    
    elif filename[-4:] == ".pdf":
        return convert_pdf_to_txt(file_path,filename)
    
    elif filename[-4:] == ".rtf":
        pdf_file =  rtf_file_to_pdf(filename,file_path)
        if pdf_file == None :
            print("could not read a .rtf file")
        return convert_pdf_to_txt(filepath,pdf_file)
    
    else:                   # odt  or doc files
        newconvertedPDF = odt_file_to_pdf(filename,file_path)
        if newconvertedPDF == None :
            print("file type: ", filename[-4:] ,"could not read a that file.")
        return convert_pdf_to_txt(filepath,newconvertedPDF)

def odt_file_to_pdf(filename,fpath): #works fine with odt files
    wdFormatPDF = 17
    final_file = None
    output_file = filename.split('.')[0]
    in_file = os.path.join(fpath, filename)
    out_file = os.path.join(fpath,output_file + '.pdf' )
    from win32com.client import gencache, constants, Dispatch
    # that's the magic part
    gencache.EnsureModule('{00020905-0000-0000-C000-000000000046}', 0, 8, 3)
    app = Dispatch("Word.Application.8")
    # open a document
    doc = app.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    app.Quit()
    final_file = output_file + ".pdf"
    return final_file

def rtf_file_to_pdf(filename,fpath):
   
    wdFormatPDF = 17
    input_dir = fpath
    #output_dir = fpath
    final_file = None
    in_file = os.path.join(input_dir, filename)
    output_file = filename.split('.')[0]
    if ( filename.split('.')[1] == "rtf"):
        out_file = os.path.join(input_dir,output_file + '.pdf' )
        word = comtypes.client.CreateObject('Word.Application')
        doc = word.Documents.Open(in_file)
        doc.SaveAs(out_file, FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()
        final_file = output_file + ".pdf"
    return final_file

if __name__ == '__main__':
    filepath = "your path to that file"
    filename = "the file name"
    txt = get_alldocs_to_text(filename, filepath)
