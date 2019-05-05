import sys
import os
import comtypes.client

wdFormatPDF = 17

current_dir = os.getcwd()

for i in os.listdir(current_dir):
    if i.endswith('.docx'):
        word = comtypes.client.CreateObject('Word.Application')
        output_file_name = i[:-5] # removes .docx extension and selects just the file name for saving
            
        out_file = os.path.abspath('{}'.format(output_file_name))

        doc = word.Documents.Open(current_dir  + "/{}".format(i))
        doc.SaveAs(out_file, FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()