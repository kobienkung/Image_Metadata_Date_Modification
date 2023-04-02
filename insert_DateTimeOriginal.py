from datetime import datetime
from exif import Image
import os
import subprocess
import pandas as pd
from openpyxl import load_workbook
import re

# Get all file information
input_folder_path = r'C:\Users\kobienkung\OneDrive - KMITL\REMINISENCE'
file_name = []
file_extension = []
file_path = []
for path, directories, files in os.walk(input_folder_path):
    for file in files:
        file_path.append(os.path.join(path, file))
        split_names = os.path.splitext(file)
        file_name.append(split_names[0])
        file_extension.append(split_names[1])
df_name = pd.DataFrame({'file_name':file_name, 'file_extension':file_extension, 'file_path':file_path})

# Export file information to .xlsx file
FilePath = "file_name.xlsx"
try:
    ExcelWorkbook = load_workbook(FilePath)
    with pd.ExcelWriter(FilePath, engine = 'openpyxl') as writer:
        writer.book = ExcelWorkbook
        df_name.to_excel(writer, sheet_name='file_path')
except Exception as exc:
    if type(exc) == FileNotFoundError: 
        df_name.to_excel(writer, sheet_name='Python')
    else:
        print(exc)











def unix_to_CE(unix_num):
    if len(unix_num) == 10:
        date_from_unix = datetime.utcfromtimestamp(int(unix_num)).strftime("%Y:%m:%d %H:%M:%S")
    elif len(unix_num) == 13:
        date_from_unix = datetime.utcfromtimestamp(int(unix_num)/1000).strftime("%Y:%m:%d %H:%M:%S")
    else:
        pass
    return date_from_unix

withDatetimeOriginalFile = {'.jpeg','.jpg','.mp4'}
withCreationTimeFile = {'.png'}
validExtension = withDatetimeOriginalFile.union(withCreationTimeFile)

FilePath = "Exiftool_output10621.xlsx"
df = pd.read_excel(FilePath, sheet_name='Sheet1', dtype=str, keep_default_na=False)
if 'output' not in df.columns:
    df['output'] = ''

start_ind = 10000
stop_ind = len(df)
start_time = datetime.now()
for i in range(start_ind,stop_ind):
    print(i)
    file_extension = df['file_extension'][i]
    if file_extension not in validExtension: 
        df['output'][i] = 'invalid'# file_extension'
        continue
    img_path = df['file_path'][i]
    img_attr = subprocess.check_output(['exiftool', img_path])
    img_attr = str(img_attr)
    if 'Creation Time' in img_attr or 'Date/Time Original' in img_attr: 
        df['output'][i] = 'already'# have date taken'
        continue
    
    all_dates = []
    img_attr_split = str(img_attr).split('\\r\\n')
    all_date_attr = [a for a in img_attr_split if 'date/time' in a.lower()]
    all_dates = [a[a.find(': ')+2:] for a in all_date_attr]

    if df['remark'][i]:
        if df['remark'][i] == 'ready_date':
            date_from_file_name = df['to_use_number'][i]
            all_dates.append(date_from_file_name)
        elif df['remark'][i] == 'unix' and len(df['to_use_number'][i]) in [10,13]:
            date_from_file_name = unix_to_CE(df['to_use_number'][i])
            all_dates.append(date_from_file_name)

    min_date = min(all_dates)
    if not min_date: 
        df['output'][i] = 'non date'
        continue

    if file_extension.lower() in withDatetimeOriginalFile:
        exiftool_command = ['exiftool', '-overwrite_original', f'-DatetimeOriginal="{min_date}"', img_path]
    elif file_extension.lower() in withCreationTimeFile:
        exiftool_command = ['exiftool', '-overwrite_original', f'-PNG:CreationTime="{min_date}"', img_path]
    try:
        df['output'][i] = subprocess.check_output(exiftool_command)
    except:
        df['output'][i] = 'error'

stop_time = datetime.now()
print(f'done {str(i-start_ind+1)} images in {stop_time - start_time}')

df.to_excel(f'Exiftool_output{str(i+1)}.xlsx')



# Extract date from the name
# Set dateTime for video, pdf 
# Check if it's png file
# Does it already have -DateTimeOriginal / -CreationTime
# Collect output status results

#unix_name = '1587310027591'
#date_from_unix = 'fake'
# get unix date from file name
# demo




df = pd.read_excel('file_name.xlsx', sheet_name='file_path', keep_default_na=False)
for i in range(len(df)):
    df['remark'][i] = unix_to_CE(df['to_use_number'][i])

file = '2014-05-18'
def try_parsing_date(text):
    for fmt in ('%Y-%m-%d', '%d.%m.%Y', '%d/%m/%Y'):
        try:
            print(datetime.strptime(text, fmt))
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    #raise ValueError('no valid date format found')
date_from_file_name = try_parsing_date(file)
if date_from_file_name:  all_dates.append(date_from_file_name)

def try_parsing_date_with_time(text):
    for fmt in ("%Y:%m:%d %H:%M:%S+%Z"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    #raise ValueError('no valid date format found')
test = [a for a in all_dates if min_date in a]
for a in test:
    complete_min_date =  try_parsing_date_with_time(a)

#---------------------------------------------------------------
output_folder_path = 'C:/Users/kobienkung/Pictures/Exiftool_output'

if not os.path.exists(output_folder_path): 
    os.mkdir(output_folder_path)
    
img_file_name = 'Screenshot (143).png'
img_path = input_folder_path + '/' + img_file_name

with open(img_path, 'rb') as img_file:
    img = Image(img_file)

# TO DO
# if img.has_exif and img.get(datetime_original):

[print(a + ' ' + img.get(a)) for a in img.list_all() if 'datetime' in a]

first_datetime = min([img.get(a) for a in img.list_all() if 'datetime' in a])

img.set('datetime_original', img.datetime_original)

new_img_path = output_folder_path + '/' + img_file_name

with open(new_img_path, 'wb') as new_img_file:
    new_img_file.write(img.get_file())
