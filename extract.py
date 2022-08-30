from PyPDF2 import PdfFileReader
import re
import csv
import pandas as pd

# Brings in PDF file
pdf = PdfFileReader('visa-merchant-data-standards-manual.pdf')

''' Regular expression to identify MCCs'''
new_mcc_re = re.compile(r'^\d{4}  [A-Z].*') # get all starting with 4 digits + 2 spaces + Caps + everything after

''' main extraction function to iterate through pages, sentences '''
mcc_dict = {}
for page in range(24, 106):
  page_object = pdf.getPage(page) # zero-indexed page 24 - 106
  page_text = page_object.extractText()
  sentence_lst = page_text.split('\n')

  for sentence in sentence_lst:
    if new_mcc_re.match(sentence):
      mcc_num, *mcc_desc = sentence.split()
      mcc_desc = ' '.join(mcc_desc)
      mcc_dict[mcc_num] = mcc_desc

'''Exports to .csv file'''
a_file = open("MCCs.csv", "w")
writer = csv.writer(a_file)
for key, value in mcc_dict.items():
    writer.writerow([key, value])

''' Exports to .xlxs file'''
df = pd.DataFrame(data=mcc_dict, index=[0])
#convert into excel
df = (df.T)
df.set_axis(['MCCs Descriptions'], axis=1, inplace=True)
df.to_excel("MCCs.xlsx")
