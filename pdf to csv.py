import re
import pandas
import pdfplumber
import pandas as pd
from collections import namedtuple
import os

# Read a PDF File
os.chdir('/Users/jsun/Desktop/xxxxx/')
cwd = os.getcwd()
#print(cwd)

Line = namedtuple('Line', 'Product')

# Find a pattern
line_re = re.compile(r'\d+ \w+\d+\w+\d+ \w+') # this might need to manually put in pattern
line2_re = re.compile(r'\d+ \w+\d+\w+ \w+') # this might need to manually put in pattern
line3_re = re.compile(r'\d+ \w{7} \w{3}') # this might need to manually put in pattern
line4_re = re.compile(r'\d+ \w{6} \w{3}')
line5_re = re.compile(r'\d+ \w+\d{1}/\d{1}\w+ \w{3}')
line6_re = re.compile(r'\d+ \w{5} \w{3}')
line7_re = re.compile(r'\d+ \w{4} \w{3}')
line8_re = re.compile(re.compile(r'\d+ \w{3}\d \w{3}'))
line9_re = re.compile(r'\d (.*) \w+ \d.\d{2} (.*) (.*)') # and this too
line10_re = re.compile(r'\w+ \d{6} \w+')
line11_re = re.compile(r'\d{2}/\d{2}/\d{4} \d{8}-\d{2} \d')

file = "xxxxx.pdf" # This need to change based on unique pdf name

# Looping
lines = []
lines2 = []
with pdfplumber.open(file) as pdf:
    pages = pdf.pages
    for page in pdf.pages:
        text = page.extract_text()
        for line in text.split('\n')[16:]:
            if line_re.search(line) or line2_re.search(line) or line3_re.search(line) or line4_re.search(line) or \
                    line5_re.search(line) or line6_re.search(line) or line7_re.search(line) or line8_re.search(line):
                items = line.split()[1]
                if items == "ON":
                    continue
                lines.append(Line(items))
            elif line9_re.search(line):
                items2 = line.split()[1]
                lines2.append(items2)

with pdfplumber.open(file) as pdf:
    pages = pdf.pages
    for page in pdf.pages:
        text = page.extract_text()
        for line in text.split('\n')[2:3]:
            if line11_re.search(line):
                quote_no = line.split()[1]
    #print(quote_no)

with pdfplumber.open(file) as pdf:
    pages = pdf.pages
    for page in pdf.pages:
        text = page.extract_text()
        for line in text.split('\n')[4:5]:
            if line10_re.search(line):
                cust_no = line.split()[3]
    #print(cust_no)

# Create a new Excel sheet
df = pd.DataFrame(lines)
df2 = pd.DataFrame({"PDSC Price": lines2})
df3 = pd.DataFrame({"PDSC Qty Max": []})
df4 = pd.DataFrame({"PDSR Price": []})
df5 = pd.DataFrame({"PDSR Qty Max": []})
df6 = pandas.concat([df, df2, df3, df4, df5], axis = 1)

df.to_csv('Product.csv', index = False)
df2.to_csv('PDSC Price.csv', index = False)
df6.to_excel('{}_{}.xlsx'.format(cust_no, quote_no), index = False)
