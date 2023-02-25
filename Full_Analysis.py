# import the required modules
import glob
import pandas as pd
import numpy as np
from datetime import datetime
import calendar
import time
start_time = time.time()

# specifying the folder with the csv files
path = "Analysis/2022"

# getting the csv files from the folder
files = glob.glob(path + "/*.xlsx")

# defining an empty list to store the content
df = pd.DataFrame()
content = []

# defining the list of columns the we have interest in
columns = ['rideable_type', 'started_at', 'ended_at', 'member_casual']

# reading each xlsx file in the specified path
for filename in files:
    
    # reading the content of the xlsx file and appending it to the list we created
    df = pd.read_excel(filename,
                    usecols = columns,
                    parse_dates = ['started_at', 'ended_at'])
    content.append(df)


#df = pd.read_excel('Analysis/2022/202201-divvy-tripdata.xlsx',
#                    usecols = columns,
#                    parse_dates = ['started_at', 'ended_at'])


# converting content to data frame
df = pd.concat(content)

# create a new column with the length of each ride in minutes
df['ride_length'] = round((df['ended_at'] - df['started_at']).dt.total_seconds() / 60, 2)

# are there any null values, duplicated rows and ride lengths not positive?
# null = df.isnull().sum()
# duplicate = df.duplicated().sum()
# not_positive = (df['ride_length'] <= 0).sum()

# replacing blanck rows for 'NaN' to then drop them -- not used for 2022
#df = df.replace('', np.nan, inplace = True)
#df = df.dropna()

# removing duplicate rows -- OK
df = df.drop_duplicates()

# removing rows with the length rides below 1min and higher than 24h
df = df[(df['ride_length'] > 1) & (df['ride_length'] < 1440)]

# replace bikes labeled as "docked" with "classic"
df['rideable_type'] = df['rideable_type'].replace(['docked_bike'], 'classic_bike')

# get columns with the hour of the day, day of the week, day of the month and month of the year
df['hour'] = df['started_at'].dt.hour
df['day'] = df['started_at'].dt.day
df['week_day'] = df['started_at'].dt.weekday
df['month'] = df['started_at'].dt.month

# get a column with the time of the day for the 'hd' pivot table
df['time_of_day'] = np.nan
df.loc[df['hour'] <= 5, 'time_of_day'] = 'Night'
df.loc[(df['hour'] > 5) & (df['hour'] <= 11), 'time_of_day'] = 'Morning'
df.loc[(df['hour'] > 11) & (df['hour'] <= 17), 'time_of_day'] = 'Afternoon'
df.loc[df['hour'] > 17, 'time_of_day'] = 'Evening'

# get a column with the season of the year for the 'my' pivot table
df['season'] = np.nan
df.loc[(df['month'] == 12) | (df['month'] <= 2), 'season'] = 'Winter'
df.loc[(df['month'] >= 3) & (df['month'] <= 5), 'season'] = 'Spring'
df.loc[(df['month'] >= 6) & (df['month'] <= 8), 'season'] = 'Summer'
df.loc[(df['month'] >= 9) & (df['month'] <= 11), 'season'] = 'Fall'

# replace the month number for the month name
#df['month'] = df['started_at'].dt.strftime('%B')

# sort the data by ascending order of date
#df.sort_values(by=['started_at'], inplace = True)

# creating two workbooks to fill with the final outputs
#workbook_tot = pd.ExcelWriter('Output_Total.xlsx', engine = 'xlsxwriter')
workbook_c_vs_a = pd.ExcelWriter('Output_Casual_vs_Annual_2.xlsx', engine = 'xlsxwriter')

## calculating the metrics for the following conditions
# all data
#total_all = (df.describe()).to_excel(workbook_tot, sheet_name = 'Total')
total_casual_vs_annual = (df.groupby('member_casual').describe()).to_excel(workbook_c_vs_a, sheet_name = 'Total')

# type of bike ('tb')
#tb_all = (df.groupby('rideable_type').describe()).to_excel(workbook_tot, sheet_name = 'TB')
tb_casual_vs_annual = (df.pivot_table(index=["rideable_type", "member_casual"], 
                                    values='ride_length', 
                                    aggfunc={np.count_nonzero, np.average, np.median})
                                    ).reset_index().to_excel(workbook_c_vs_a, sheet_name = 'TB', index = False)

# hour of the day ('hd')
#hd_all = (df.groupby('hour').describe()).to_excel(workbook_tot, sheet_name = 'HD')
hd_casual_vs_annual = (df.pivot_table(index=["hour", "member_casual", 'time_of_day'], 
                                    values='ride_length', 
                                    aggfunc={np.count_nonzero, np.average, np.median})
                                    ).reset_index().to_excel(workbook_c_vs_a, sheet_name = 'HD', index = False)

# week day ('wd')
#wd_all = (df.groupby('week_day').describe()).to_excel(workbook_tot, sheet_name = 'WD')
wd_casual_vs_annual = (df.pivot_table(index=["week_day", "member_casual"], 
                                    values='ride_length', 
                                    aggfunc={np.count_nonzero, np.average, np.median})
                                    ).reset_index().to_excel(workbook_c_vs_a, sheet_name = 'WD', index = False)

# day of the month ('dm')
#dm_all = (df.groupby('day').describe()).to_excel(workbook_tot, sheet_name = 'DM')
dm_casual_vs_annual = (df.pivot_table(index=["day", "member_casual"], 
                                    values='ride_length', 
                                    aggfunc={np.count_nonzero, np.average, np.median})
                                    ).reset_index().to_excel(workbook_c_vs_a, sheet_name = 'DM', index = False)

# month of the year ('my')
#my_all = (df.groupby('month').describe()).to_excel(workbook_tot, sheet_name = 'MY')
my_casual_vs_annual = (df.pivot_table(index=["month", "member_casual", 'season'], 
                                    values='ride_length', 
                                    aggfunc={np.count_nonzero, np.average, np.median})
                                    ).reset_index().to_excel(workbook_c_vs_a, sheet_name = 'MY', index = False)

# time of the day ('td')
#td_all = (df.groupby('time_of_day').describe()).to_excel(workbook_tot, sheet_name = 'TD')
#td_casual_vs_annual = (df.pivot_table(index=["time_of_day", "member_casual"], 
#                                    values='ride_length', 
#                                    aggfunc={np.count_nonzero, np.average, np.median})
#                                    ).to_excel(workbook_c_vs_a, sheet_name = 'TD')

# season of the year ('sy')
#sy_all = (df.groupby('season').describe()).to_excel(workbook_tot, sheet_name = 'SY')
#sy_casual_vs_annual = (df.pivot_table(index=["season", "member_casual"], 
#                                    values='ride_length', 
#                                    aggfunc={np.count_nonzero, np.average, np.median})
#                                    ).to_excel(workbook_c_vs_a, sheet_name = 'SY')

# saving the data to the workbooks
#workbook_tot.save()
workbook_c_vs_a.save()
print("--- %s seconds ---" % (time.time() - start_time))