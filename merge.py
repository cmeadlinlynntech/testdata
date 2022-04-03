### Import modules
import sys

from os import listdir
from os.path import isfile, join
from openpyxl import Workbook
import pandas as pd


### Create data path
DATA_PATH = '.'


### Create a list of data files to read them
def list_files():
    files = [f for f in listdir(DATA_PATH) if isfile(join(DATA_PATH, f))]
    files = [f for f in files if f.endswith(".Absorbance")]
    files.sort()
    return files

data_files = list_files()


### Read the first data file
file_path = join(DATA_PATH, data_files[0])
def read_data_file(file_path):
    df = pd.read_csv(file_path, skiprows = 19, skipfooter = 2, engine = 'python', delimiter = '\t', header = None)
    return df


### Get the first column and create a dictionary for dataframe
first_data_file = data_files[0]
first_data_file_data = read_data_file(first_data_file)
data_dict = {}
data_dict['x'] = first_data_file_data[0].tolist()
    
    
### Append the rest of the files to dataframe
rest = data_files[0:]
for f in rest:
    header_title = f
    header_title_list = header_title.split(".")
    data_dict[header_title_list[0]] = read_data_file(f)[1].tolist()
    
df = pd.DataFrame(data = data_dict)


### Import data file to Excel
df.to_excel('PVP Spectral Reading.xlsx')
