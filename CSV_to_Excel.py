# importing the required modules
import pandas as pd
import glob
from pathlib import Path

# specifying the folder with the csv files
path = "(folder name)"

# getting the csv files from the folder
files = glob.glob(path + "/*.csv")

# reading each csv file and exporting in xlsx format
for filename in files:
    
    # reading the content of the csv file
    df = pd.read_csv(filename)
    
    # storing the file name without the csv extension
    newname = Path(filename).stem
    
    # exporting the csv file in xlsx format
    pd.DataFrame(df).to_excel(newname + '.xlsx')
