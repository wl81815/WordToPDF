# -*- coding:utf-8 -*-
# author:Pifu
# place:WF
# time:2019-11-18
import os
import comtypes.client

wdFormatPDF = 17

path = r'D:\公司\1111'

for parent,dirnames,filenames in os.walk(path):
    for filename in filenames:
        if not (filename.endswith('.doc') or filename.endswith('docx')):
            continue
        in_file = os.path.join(parent,filename)
        pdf_name = filename.split('.')[0]+'.pdf'
        out_file = os.path.join(parent,pdf_name)
        print(out_file)
        word = comtypes.client.CreateObject('Word.Application')
        doc = word.Documents.Open(in_file)
        doc.SaveAs(out_file,FileFormat = wdFormatPDF)
        doc.Close()
        word.Quit()
