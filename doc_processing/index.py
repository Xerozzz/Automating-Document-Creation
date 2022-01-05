from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
from zipfile import ZipFile
import csv
from botocore.exceptions import ClientError
import logging
import os
import boto3
import yaml


def read_yaml(file_path):
    with open(file_path, "r") as f:
        return yaml.safe_load(f)


# Read Config Data
config = read_yaml("config.yml")
BUCKET = config['BUCKET']

# Create Mailmerge Document
template = "Test.docx"
document = MailMerge(template, 'r')
fields = document.get_merge_fields()

# Open File
file = open('data.csv', 'r')
csvreader = csv.reader(file)
rows = []

for row in csvreader:
    rows.append(row)

file.close()

# Input data
count = 1
for i in rows:
    document.merge(
        Name=i[0],
        Index=i[1],
        Phone=i[2],
        Email=i[3],
    )
    document.write(f'{i[0]}.docx')
    count += 1

# Create zip package
zipObj = ZipFile('./data.zip', 'w')
for i in rows:
    filename = f'{i[0]}.docx'
    zipObj.write(filename)
    os.remove(filename)
zipObj.close()

# Create Boto3 and Upload File
