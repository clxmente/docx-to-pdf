import sys
import os
import comtypes.client

wdFormatPDF = 17

current_dir = os.getcwd()

for i in os.listdir(current_dir):
    if i.endswith('.docx'):
        word = comtypes.client.CreateObject('Word.Application')
        output_file_name = i[:-5] # removes .docx extension and selects just the file name for saving
        if output_file_name.startswith("/"):
            final_out = output_file_name[1:]  # removes the "/" at the beginning of the file name that i think is needed for the program to work
        else:
            final_out = output_file_name
            
        out_file = os.path.abspath('{}'.format(final_out))

        doc = word.Documents.Open(current_dir  + i)
        doc.SaveAs(out_file, FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()