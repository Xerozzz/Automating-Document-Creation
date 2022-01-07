from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
from zipfile import ZipFile
import csv
import logging
import os
import yaml


def read_yaml(file_path):
    with open(file_path, "r") as f:
        return yaml.safe_load(f)


# Create Mailmerge Document
# Change this value if you use another name besides Template.docx
template = "Template.docx"
document = MailMerge(template, 'r')
fields = document.get_merge_fields()

# Open File
# Change this value if you use another name besides data.csv
file = open('data.csv', 'r')
document = MailMerge(template, 'r')
csvreader = csv.reader(file)
rows = []

for row in csvreader:
    rows.append(row)
file.close()

# Creation of documents
count = 1
for i in rows:
    document.merge(  # CHANGE THE FIELDS HERE ACCORDING TO YOUR CSV DATA AND TEMPLATE MERGEFIELDS
        Name=i[0],
        Index=i[1],
        Phone=i[2],
        Email=i[3],
    )
    document.write(f'{i[0]}.docx')
    count += 1

# Create zip package
zipObj = ZipFile('../storage/data.zip', 'w')
for i in rows:
    filename = f'{i[0]}.docx'
    zipObj.write(filename)
    os.remove(filename)
zipObj.close()
