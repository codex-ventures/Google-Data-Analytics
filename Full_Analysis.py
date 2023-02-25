# import the required modules
import glob
import pandas as pd
import numpy as np
from datetime import datetime
import calendar

# specifying the folder with the csv files
path = "(folder path)"

# getting the csv files from the folder
files = glob.glob(path + "/*.xlsx")

# defining an empty list to store the content
df = pd.DataFrame()
content = []

# defining the list of columns we are interested in
columns = ['rideable_type', 'started_at', 'ended_at', 'member_casual']

# reading each xlsx file in the specified path
for filename in files:
    
    # reading the content of the xlsx file and appending it to the list we created
    df = pd.read_excel(filename, usecols = columns, parse_dates = ['started_at', 'ended_at'])
    content.append(df)

# converting content to data frame
df = pd.concat(content)

# create a new column with the length of each ride in minutes
df['ride_length'] = round((df['ended_at'] - df['started_at']).dt.total_seconds() / 60, 2)

# removing duplicate rows
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

# get a column with the time of the day
df['time_of_day'] = np.nan
df.loc[df['hour'] <= 5, 'time_of_day'] = 'Night'
df.loc[(df['hour'] > 5) & (df['hour'] <= 11), 'time_of_day'] = 'Morning'
df.loc[(df['hour'] > 11) & (df['hour'] <= 17), 'time_of_day'] = 'Afternoon'
df.loc[df['hour'] > 17, 'time_of_day'] = 'Evening'

# get a column with the season of the year
df['season'] = np.nan
df.loc[(df['month'] == 12) | (df['month'] <= 2), 'season'] = 'Winter'
df.loc[(df['month'] >= 3) & (df['month'] <= 5), 'season'] = 'Spring'
df.loc[(df['month'] >= 6) & (df['month'] <= 8), 'season'] = 'Summer'
df.loc[(df['month'] >= 9) & (df['month'] <= 11), 'season'] = 'Fall'

# creating a workbook to fill with the final output
final_workbook = pd.ExcelWriter('Final_Output.xlsx', engine = 'xlsxwriter')

## calculating the metrics for the following conditions
# all data
total = (df.groupby('member_casual').describe()).to_excel(final_workbook, sheet_name = 'Total')

# type of bike
type_of_bike = (df.pivot_table(index=["rideable_type", "member_casual"], 
                                values='ride_length', 
                                aggfunc={np.count_nonzero, np.average, np.median})
                                ).reset_index().to_excel(final_workbook, sheet_name = 'TB', index = False)

# hour of the day
hour_of_day = (df.pivot_table(index=["hour", "member_casual", 'time_of_day'], 
                                values='ride_length', 
                                aggfunc={np.count_nonzero, np.average, np.median})
                                ).reset_index().to_excel(final_workbook, sheet_name = 'HD', index = False)

# week day
weekday = (df.pivot_table(index=["week_day", "member_casual"], 
                            values='ride_length', 
                            aggfunc={np.count_nonzero, np.average, np.median})
                            ).reset_index().to_excel(final_workbook, sheet_name = 'WD', index = False)

# day of the month
day_of_month = (df.pivot_table(index=["day", "member_casual"], 
                                values='ride_length', 
                                aggfunc={np.count_nonzero, np.average, np.median})
                                ).reset_index().to_excel(final_workbook, sheet_name = 'DM', index = False)

# month of the year
month = (df.pivot_table(index=["month", "member_casual", 'season'], 
                        values='ride_length', 
                        aggfunc={np.count_nonzero, np.average, np.median})
                        ).reset_index().to_excel(final_workbook, sheet_name = 'MY', index = False)

final_workbook.save()
