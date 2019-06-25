import requests
import xlrd
import csv
import pandas as pd
from datetime import datetime
import os

def pull_desired_files_and_store_as_csvs(dates):
    """
    This function web scrapes the target website, pulls all of the xlsx files based on the selected dates using the xlrd function,
    converts to csv files, and reads in all csv's into a master pandas dataframe that will be used for all subsequent calculations.
    Arguments:
        dates: list. List of all of the dates (month/year, formatted based on website xlsx formatting), which is looped through
    Outputs:
        master_df: Pandas dataframe. 
        csv's of all of the files that were scraped, saved in the Documents folder.
    """
    #Define the base url we will be pulling off of
    base_url = "https://www.dmr.nd.gov/oilgas/mpr/" 
    #Create a master pandas dataframe that we will return that holds all of the data that we want
    master_df=pd.DataFrame()
    #loop through all of the dates in the list
    for date in dates:        
        desired_url=base_url+date+'.xlsx'
        r = requests.get(desired_url)  # make an HTTP request
        #Open the contents of the xlsx file as an xlrd workbook    
        workbook = xlrd.open_workbook(file_contents=r.content)  
        #Take first worksheet in the workbook, as it's the one we'll be using
        worksheet = workbook.sheet_by_index(0) 
        #Obtain the year/month date data from the worksheet, and convert the ReportDate column from float to datetime,
        #using xlrd datemode functionality
        for i in range(1, worksheet.nrows):
            wrongValue = worksheet.cell_value(i,0)
            workbook_datemode = workbook.datemode
            year, month, day, hour, minute, second = xlrd.xldate_as_tuple(wrongValue, workbook_datemode)
            worksheet._cell_values[i][0]=datetime(year, month, 1).strftime("%m/%Y")
        #Generate a csv name to save under
        file_name='C:/Bakken/'+date+'.csv'
        #Save as a csv
        csv_file = open(file_name, 'w',newline='')
        #Create writer to csv file
        wr = csv.writer(csv_file)
        #Loop through all the rows and write to csv file
        for rownum in range(worksheet.nrows):
            wr.writerow(worksheet.row_values(rownum))
        #Close the csv file
        csv_file.close()
        #Read in csv as pandas dataframe
        dataframe=pd.read_csv(file_name)
        #Append to the master dataframe
        master_df=master_df.append(dataframe)
    #Return the final master dataframes
    return master_df

def main():
    #Create a new folder called 'Bakken' to drop all of the files in
    newpath = r'C:\Bakken' 
    if not os.path.exists(newpath):
        os.makedirs(newpath)
    #Pull the months that we want to process--December 2016 to January 2019
    dates=['2016_12', '2017_01', '2017_02', '2017_03', '2017_04', '2017_05', '2017_06', 
           '2017_07', '2017_08', '2017_09', '2017_10', '2017_11', '2017_12',
           '2018_01', '2018_02', '2018_03', '2018_04', '2018_05', '2018_06', 
           '2018_07', '2018_08', '2018_09', '2018_10', '2018_11', '2018_12', '2019_01']
    #Run through the web scraper and save the desired csv files. Create a master dataframe with all of the months' data
    master_dataframe_production=pull_desired_files_and_store_as_csvs(dates)
    #Declare the ReportDate column as a pandas datetime object
    master_dataframe_production['ReportDate']=pd.to_datetime(master_dataframe_production['ReportDate'])
    #Write the master dataframe to a master csv 
    master_dataframe_production.to_csv(newpath+'\master_dataframe_production.csv')

if __name__== "__main__":
    main()

