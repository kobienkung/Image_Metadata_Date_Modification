from datetime import datetime
from exif import Image
import os
import subprocess
import pandas as pd
from openpyxl import load_workbook
import re

def get_file_information(input_folder_path, excel_file_name):
    '''Read all files in a given folder then export file_name and file_extension as an excel'''
    # Get all file information
    file_name = []
    file_extension = []
    file_path = []
    remark = []
    to_use_date = []
    for path, directories, files in os.walk(input_folder_path):
        for file in files:
            file_path.append(os.path.join(path, file))
            split_names = os.path.splitext(file)
            file_name.append(split_names[0])
            file_extension.append(split_names[1])
    df_name = pd.DataFrame({'file_name':file_name, 'file_extension':file_extension, 'file_path':file_path})

    # Export file information to .xlsx file
    FilePath = excel_file_name
    try:
        ExcelWorkbook = load_workbook(FilePath)
        with pd.ExcelWriter(FilePath, engine = 'openpyxl') as writer:
            writer.book = ExcelWorkbook
            df_name.to_excel(writer)
    except Exception as exc:
        if type(exc) == FileNotFoundError: 
            df_name.to_excel(FilePath)
        else:
            print(exc)


input_folder_path = r'C:\Users\kobienkung\OneDrive - KMITL\REMINISENCE\Marriage ceremony'
excel_file_name = 'wedding.xlsx'
get_file_information(input_folder_path, excel_file_name)
# open exel file and manually get date_taken/DateTimeOriginal from file_name
# if dates are found in file_name, put 'ready_date' or 'unix' in remark column
# then put unix_number(eg., '1646396683562') for 'unix' remark in to_use_date column
# or put YYYY:mm:DD HH:MM:SS:fff (eg., '2021:10:31 18:04:40:940') for 'ready_date' remark in to_use_date column




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

FilePath = excel_file_name
df = pd.read_excel(FilePath, sheet_name='Sheet1', dtype=str, keep_default_na=False)
if 'output' not in df.columns:
    df['output'] = ''

start_ind = 0
stop_ind = len(df)
start_time = datetime.now()
print(start_time)
for i in range(start_ind,stop_ind):
    print(i)
    file_extension = df['file_extension'][i]
    if file_extension not in validExtension: 
        df['output'][i] = 'invalid'# file_extension'
        continue
    img_path = df['file_path'][i]
    img_attr = subprocess.check_output(['exiftool', img_path])
    img_attr = str(img_attr)
    # to implement as an option if to skip files with existing date_taken
    if 'Creation Time' in img_attr or 'Date/Time Original' in img_attr: 
        df['output'][i] = 'already'# have date taken'
        continue
    
    all_dates = []
    img_attr_split = str(img_attr).split('\\r\\n')
    all_date_attr = [a for a in img_attr_split if 'date/time' in a.lower()]
    all_dates = [a[a.find(': ')+2:] for a in all_date_attr]

    if df['remark'][i]:
        if df['remark'][i] == 'ready_date':
            date_from_file_name = df['to_use_date'][i]
            all_dates.append(date_from_file_name)
        elif df['remark'][i] == 'unix' and len(df['to_use_date'][i]) in [10,13]:
            date_from_file_name = unix_to_CE(df['to_use_date'][i])
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
print(stop_time)
print(f'done {str(i-start_ind+1)} images in {stop_time - start_time}')

df.to_excel(f'Exiftool_output{str(i+1)}.xlsx')