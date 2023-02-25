# importing the required modules
import pandas as pd
import glob
from pathlib import Path

# specifying the folder with the csv files
path = "Divvy_2015"

# getting the csv files from the folder
files = glob.glob(path + "/*.csv")

# defining an empty list to store the content
#data_frame = pd.DataFrame()
#content = []

# reading each csv file and exporting in xlsx format
for filename in files:
    
    # reading the content of the csv file
    df = pd.read_csv(filename)
    
    # storing the file name without the csv extension
    newname = Path(filename).stem
    
    # exporting the csv file in xlsx format
    pd.DataFrame(df).to_excel(newname + '.xlsx')
    #content.append(df)

# converting content to data frame
#data_frame = pd.concat(content)
#pd.DataFrame(data_frame).to_excel('2022.xlsx')