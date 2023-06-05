""" *****************************************************************************
-*- coding: utf-8 -*-
    \file           2.Merging_Script
    \author         Hans (You Yang) ONG
    \co-author      A. L.
    \co-author      J. W. F.
    \co-author      W. C. W.

    \creation date  230523
    \last updated   050623

    \brief          This script extracts, merges, and processes data from multiple Excel files.
                    It reads Temperature, Concentration, Blaze Statistics, Blaze LW Distribution,
                    and Blaze CW Distribution data from the Excel files, merges them into a single
                    DataFrame, and adds summary data to the merged DataFrame.
                    The script saves the merged DataFrame as a new Excel file.
                    The script also handles nearest neighbor data resampling and addresses the
                    behavior when Blaze data is smaller.
***************************************************************************** """
import pandas as pd         
from openpyxl import load_workbook
                                        # libraries
filenames = ["C1R2"]
# -----------------------------------------------------------------------------
# -----------------------------------------------------------------------------
# \brief Extract Temperature and Concentration data from an Excel file.
#
# This function reads Temperature and Concentration data from an Excel file and returns a pandas DataFrame.
#
# \param file_name:     The name of the Excel file to be read.
# \param sheet_name:    The name of the sheet containing the Temperature and Concentration data.
# \param skiprows:      The number of rows to skip while reading the Excel file.
#
# \return:              A pandas DataFrame containing the Temperature and Concentration data.
# -----------------------------------------------------------------------------
def read_tc_data(file_name, sheet_name, skiprows):
    df_TC = pd.read_excel(file_name, sheet_name=sheet_name, skiprows=skiprows)
    df_TC = df_TC.iloc[:, 0:7]
    df_TC['Time (sec)'] = df_TC['Time (sec)'].astype(float)
    df_TC.fillna(method="ffill", inplace=True)
    return df_TC

# -----------------------------------------------------------------------------
# \brief Read Blaze Statistics from an Excel file.
#
# This function reads Blaze Statistics from an Excel file and returns a pandas DataFrame.
#
# \param file_name:     The name of the Excel file to be read.
# \param sheet_name:    The name of the sheet containing the Blaze Statistics.
# \param skiprows:      The number of rows to skip while reading the Excel file.
#
# \return:              A pandas DataFrame containing the Blaze Statistics.
# -----------------------------------------------------------------------------
def read_blaze_stats(file_name, sheet_name, skiprows):
    header_df = pd.read_excel(file_name, sheet_name=sheet_name, header=None, usecols="E:P", nrows=6)
    new_column_titles = header_df.apply(lambda row: ' '.join(row.values.astype(str)), axis=0)
    df_Blaze_Stats = pd.read_excel(file_name, sheet_name=sheet_name, skiprows=skiprows)
    df_Blaze_Stats.columns.values[4:] = new_column_titles.values
    df_Blaze_Stats.set_index(['Local Time'], inplace=True)
    df_Blaze_Stats = df_Blaze_Stats.resample("1s").mean().interpolate('linear')
    df_Blaze_Stats.reset_index(inplace=True)
    return df_Blaze_Stats

# -----------------------------------------------------------------------------
# \brief Read Blaze LW Distribution from an Excel file.
#
# This function reads Blaze LW Distribution data from an Excel file and returns a pandas DataFrame.
#
# \param file_name:     The name of the Excel file to be read.
# \param sheet_name:    The name of the sheet containing the Blaze LW Distribution.
# \param skiprows:      The number of rows to skip while reading the Excel file.
#
# \return:              A pandas DataFrame containing the Blaze LW Distribution.
# -----------------------------------------------------------------------------
def read_blaze_LW_dist(file_name, sheet_name, skiprows):
    df_Blaze_LW_Dist = pd.read_excel(file_name, sheet_name=sheet_name, skiprows=skiprows)
    df_Blaze_LW_Dist.set_index(['Local\nTime'], inplace=True)
    df_Blaze_LW_Dist = df_Blaze_LW_Dist.resample("1s").mean().interpolate('linear')
    df_Blaze_LW_Dist.reset_index(inplace=True)
    return df_Blaze_LW_Dist

# -----------------------------------------------------------------------------
# \brief Read Blaze CW Distribution from an Excel file.
#
# This function reads Blaze CW Distribution data from an Excel file and returns a pandas DataFrame.
#
# \param file_name:     The name of the Excel file to be read.
# \param sheet_name:    The name of the sheet containing the Blaze CW Distribution.
# \param skiprows:      The number of rows to skip while reading the Excel file.
#
# \return:              A pandas DataFrame containing the Blaze CW Distribution.
# -----------------------------------------------------------------------------
def read_blaze_CW_dist(file_name, sheet_name, skiprows):
    df_Blaze_CW_Dist = pd.read_excel(file_name, sheet_name=sheet_name, skiprows=skiprows)
    df_Blaze_CW_Dist.set_index(['Local\nTime'], inplace=True)
    df_Blaze_CW_Dist = df_Blaze_CW_Dist.resample("1s").mean().interpolate('linear')
    df_Blaze_CW_Dist.reset_index(inplace=True)
    return df_Blaze_CW_Dist


# -----------------------------------------------------------------------------
# \brief Merge DataFrames and handle common columns.
#
# This function merges the given DataFrames and checks for equality in common columns.
# If common columns are equal, it drops the redundant ones.
#
# \param df1:    First DataFrame.
# \param df2:    Second DataFrame.
# \param df3:    Third DataFrame.
# \param df4:    Fourth DataFrame.
#
# \return:       A merged DataFrame.
# -----------------------------------------------------------------------------
def merge_df(df1, df2, df3, df4, filename):
    df_merged = pd.merge_asof(right=df2, left=df1, right_on="Experimental time (sec)", left_on="Time (sec)", suffixes=('', '_Blaze_Stats'))
    df_merged = pd.merge_asof(right=df3, left=df_merged, right_on="Experimental time (sec)", left_on="Time (sec)", suffixes=('', '_Blaze_LW_Dist'))
    df_merged = pd.merge_asof(right=df4, left=df_merged, right_on="Experimental time (sec)", left_on="Time (sec)", suffixes=('', '_Blaze_CW_Dist'))
    
    # Check if 'Local Time' is in datetime format
    if df_merged['Local Time'].dtype != 'datetime64[ns]':
        df_merged['Local Time'] = pd.to_datetime(df_merged['Local Time'], origin='1900-01-01', unit='D')
        
    df_merged.set_index('Local Time', inplace=True)
    df_merged = df_merged.resample("60s", origin='start').mean().interpolate('linear')

    df_merged.index = df_merged.index.strftime('%Y-%m-%d %H:%M:%S')  # format 'Local Time'

    df_merged.to_excel('Merged/' + filename + '_Merged.xlsx', engine='openpyxl') 
    if len(df_merged) > 0:
        print("----- Merge Successful for file:", filename, "-----")
    else:
        print("----- Merge Unsuccessful for file:", filename, "-----")
    return df_merged



#@TODO the summary data is begin read either as formula, or "raw" value where all the formulas will show up as NaN.
# find a way to read as value, 
def read_summary_data(file_name, sheet_name):                                          
    wb = load_workbook(file_name, read_only=True, data_only=True) 
    df_summary = pd.read_excel(wb, sheet_name=sheet_name, index_col=None, header=None, usecols="A:B", engine='openpyxl')
    df_summary.columns = ["Parameter", "Value"]
    #display(df_summary)
    return df_summary


def add_summary_data(df_merged, df_summary):
    #display(df_summary)
    paramtype = ''
    for i in range(df_summary.shape[0]):
        param_name = df_summary.iloc[i, 0]
        
        param_value = df_summary.iloc[i, 1]
        #print(param_name)
          
        df_merged[str(param_name) + paramtype] = param_value # e.g [PCM] in summary table value == NaN, the df_merged's og data is changed to NaN
        
        if param_name == 'STARTING CONDITIONS':
            paramtype = '_static_start'
        if param_name == "EXPERIMENTAL RESULTS":
            paramtype = '_static_exp'                               
    return df_merged

filepath = "C_EXP_FORMATTED/"
extension = ".xlsx"
filenames = ["C1R2_FORMATTED"]
for i in range(2, 31): 
    filenames.append("C" + str(i) + "_FORMATTED")
filenames.remove("C4_FORMATTED")  # C4 is broken don't use

for filename in filenames:
    df_summary = read_summary_data(filepath + filename + extension, 'Summary')
    df_TC = read_tc_data(filepath + filename + extension, 'Temp and Conc', 1)
    df_Blaze_Stats = read_blaze_stats(filepath + filename + extension, 'Blaze Statistics', 7)
    df_Blaze_LW_Dist = read_blaze_LW_dist(filepath + filename + extension, 'Blaze LW Distribution', 2)
    df_Blaze_CW_Dist = read_blaze_CW_dist(filepath + filename + extension, 'Blaze CW Distribution', 2)

    df_merged = merge_df(df_TC, df_Blaze_Stats, df_Blaze_LW_Dist, df_Blaze_CW_Dist, filename)

    # Add summary data to the merged DataFrame
    df_merged = add_summary_data(df_merged, df_summary)
    df_merged.to_excel('Merged/' + filename + '_Merged.xlsx')


# nearest neighbour more than 1 second.
# check behaviour for merge when blaze is smaller

