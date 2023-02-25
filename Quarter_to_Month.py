# importing the required modules
import pandas as pd

# reading the csv file
df = pd.read_csv('Divvy_2013/Divvy_Trips_2013.csv')

# change the column to datetime
df['starttime'] = pd.to_datetime(df['starttime'])

# separate the quarter dataframe into months
jan = df[df['starttime'].dt.month == 1]
feb = df[df['starttime'].dt.month == 2]
mar = df[df['starttime'].dt.month == 3]
apr = df[df['starttime'].dt.month == 4]
may = df[df['starttime'].dt.month == 5]
jun = df[df['starttime'].dt.month == 6]
jul = df[df['starttime'].dt.month == 7]
aug = df[df['starttime'].dt.month == 8]
sep = df[df['starttime'].dt.month == 9]
oct = df[df['starttime'].dt.month == 10]
nov = df[df['starttime'].dt.month == 11]
dec = df[df['starttime'].dt.month == 12]

# exporting the data into excel files
pd.DataFrame(jan).to_excel('201301-divvy-tripdata.xlsx')
pd.DataFrame(feb).to_excel('201302-divvy-tripdata.xlsx')
pd.DataFrame(mar).to_excel('201303-divvy-tripdata.xlsx')
pd.DataFrame(apr).to_excel('201304-divvy-tripdata.xlsx')
pd.DataFrame(may).to_excel('201305-divvy-tripdata.xlsx')
pd.DataFrame(jun).to_excel('201306-divvy-tripdata.xlsx')
pd.DataFrame(jul).to_excel('201307-divvy-tripdata.xlsx')
pd.DataFrame(aug).to_excel('201308-divvy-tripdata.xlsx')
pd.DataFrame(sep).to_excel('201309-divvy-tripdata.xlsx')
pd.DataFrame(oct).to_excel('201310-divvy-tripdata.xlsx')
pd.DataFrame(nov).to_excel('201311-divvy-tripdata.xlsx')
pd.DataFrame(dec).to_excel('201312-divvy-tripdata.xlsx')