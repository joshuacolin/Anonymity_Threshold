import pandas as pd
import time
import datetime
import tkinter as tk
import webbrowser
from tkinter import filedialog,messagebox, ttk
import os
import sys

########################################## Execute the script ##########################################
def LB(ExportUnits_file_path, Response_file_path, Participants_file_path, ColumnNames, anonymity_threshold_value):
    start_time = time.time()
    
    #Get level columns
    level_Columns_Spaces  = ColumnNames
    level_Columns = [x.strip() for x in level_Columns_Spaces]
    Units = list(level_Columns)
    Units.insert(0, 'UnitName')

    # checking missing IDs column names
    if len(level_Columns) < 1:
        messagebox.showerror("Missing Column Names", "Missing Level Column Names.")
        return
    
    #Employee ID
    EmpID = 'Unique Identifier'
    ResponseEmployeeID = 'Participant Unique Identifier'

    # Unique Identifier + level columns
    Participants_Columns = list(level_Columns)
    Participants_Columns.insert(0, EmpID)
    Participants_Columns.append('Respondent')
    
    ########################### read files ############################################################
    # processing Participant File
    Participants = pd.read_csv(f'{Participants_file_path}', dtype='str')

    # processing Export Unit file
    ExportUnitsDf = pd.read_csv(f'{ExportUnits_file_path}', dtype='str')

    # processing D&A file
    Responses = pd.read_csv(f'{Response_file_path}', dtype='str')


    ############################ Participants File ################################################################
    # checking if column names exist in the Participant File
    missing_Columns = []
    for colName in level_Columns:
        if colName not in Participants:
            missing_Columns.append(colName)
    missing_ColumnsDF = pd.DataFrame(list(missing_Columns))

    # renaming the "Missing Required Columns" column
    missing_ColumnsDF.rename(columns={0: 'Missing Columns'}, inplace=True)

    if len(missing_ColumnsDF) > 0:
        messagebox.showerror("Error", "Missing Columns: \n" + '\n'.join(missing_Columns))
        #sys.exit()

    # obtaining the employee and manager IDs from the Participants file
    Participants_IDs = Participants[Participants_Columns]

    # removing blank employee ID participants
    Participants_NoNulls = Participants_IDs.dropna(subset=[EmpID, level_Columns[0]])

    # filtering non respondent participants
    Participants_IDs_Filtered = Participants_NoNulls.loc[Participants_NoNulls['Respondent'] == 'true']

    # setting Employee ID and Manager ID to be strings
    Participants_InvitedCount = Participants_IDs_Filtered[Participants_Columns].astype(str)
    Participants_Path = Participants_NoNulls[Participants_Columns].astype(str)

############################ Export Units File ################################################################
    exportUnits_IDS = ExportUnitsDf[Units]

# removing blank employee ID participants
    exportUnits_NoNulls = exportUnits_IDS.dropna(subset=level_Columns[0])

# setting Manager ID to be strings
    exportUnits_Final = exportUnits_NoNulls[Units].astype(str)


############################ Response File ################################################################
    response_IDS = Responses[[ResponseEmployeeID]].iloc[2:]

# removing blank employee id responses
    response_NoNulls = response_IDS.dropna(subset=[ResponseEmployeeID])

#replacing ".0" with nothing
    response_NoDecimals = response_NoNulls.replace('.0', '')

# converting employee ID to be string
    response_Final = response_NoDecimals[[ResponseEmployeeID]].astype(str)

# adding response column
    response_Final['Response'] = 'Yes'

    ############################ Path generation ################################################################
    # adding the column Path to the Participants DF
    Participants_Path['Path'] = ''
    exportUnits_Final['Path'] = ''
    # concatenating the level paths into one string
    for levelNum in range(len(level_Columns)):
        if levelNum == 0:
            Participants_Path['Path'] = Participants_Path['Path'] + Participants_Path[level_Columns[levelNum]]
            exportUnits_Final['Path'] = exportUnits_Final['Path'] + exportUnits_Final[level_Columns[levelNum]]
        else:
            Participants_Path['Path'] = Participants_Path['Path'] + '/' + Participants_Path[level_Columns[levelNum]]
            exportUnits_Final['Path'] = exportUnits_Final['Path'] + '/' + exportUnits_Final[level_Columns[levelNum]]

    # removing nan values from path - Participants
    Participants_Path['Path'] = Participants_Path['Path'].str.replace('/nan', '')
    Participants_PathFinal = Participants_Path.replace('nan', '')
    Participants_PathFinal = Participants_PathFinal.drop(level_Columns, axis=1)
    Participants_PathFinal = Participants_PathFinal.drop('Respondent', axis=1)

    # removing nan values from path - ExportUnits
    exportUnits_Final['Path'] = exportUnits_Final['Path'].str.replace('/nan', '')
    exportUnits_PathFinal = exportUnits_Final.replace('nan','')

    # obtaining the level number
    exportUnits_PathFinal['LevelNum'] = exportUnits_PathFinal['Path'].str.count('/') + 1

    # adding path column to the participants table
    ParticipantsWithPath = pd.merge(Participants_InvitedCount, Participants_PathFinal, on='Unique Identifier', how='left')

    # adding response column to the participants table
    ParticipantsWithResponses = pd.merge(ParticipantsWithPath, response_Final, left_on=EmpID, right_on=ResponseEmployeeID, how='left').drop(ResponseEmployeeID, axis=1)

    # filling non respondents with No
    ParticipantsWithResponses['Response'] = ParticipantsWithResponses[['Response']].fillna('No')
    
    ############################ Calculating the Invited and Response Counts ################################################################
    # calculating the expected count - No Response
    # converting the path columns into a list
    ParticipantsPathNoReponse = ParticipantsWithResponses.loc[ParticipantsWithResponses['Response'] == 'No']
    ParticipantsPathNRList = ParticipantsPathNoReponse[['Path']].sort_values('Path').values.tolist()
    # converting the list into a string
    ParticipantsPathStrNR = ' '.join([str(elem) for elem in ParticipantsPathNRList])

    # filtering participants who responded the survey
    ParticipantsPathR = ParticipantsWithResponses.loc[ParticipantsWithResponses['Response'] == 'Yes']
    # converting the path columns into a list
    ParticipantsPathRList = ParticipantsPathR[['Path']].values.tolist()

    # converting the list into a string
    ParticipantsPathStrR = ' '.join([str(elem) for elem in ParticipantsPathRList])

    expected_count_NR_dic = {}
    response_count_dic = {}
    count = 0

    for index, row in exportUnits_PathFinal.iterrows():
        if row['LevelNum'] == len(level_Columns):
            correctedPath = row['Path']
        else:
            correctedPath = row['Path'] + '/'
        # calculating the expected count
        expectedCountNR = ParticipantsPathStrNR.count(correctedPath)
        expected_count_NR_dic[row['Path']] = expectedCountNR
        # calculating the response count
        responseCount = ParticipantsPathStrR.count(correctedPath)
        response_count_dic[row['Path']] = responseCount

        print(count, expectedCountNR, responseCount, correctedPath)
        count += 1

    # assigning the paths to a data frame
    expectedCountDf = pd.DataFrame(list(expected_count_NR_dic.items()))
    responseCountDf = pd.DataFrame(list(response_count_dic.items()))

    # renaming the "ExpectedCount" column
    expectedCountDf.rename(columns={1: 'InvitedCountNR'}, inplace=True)
    responseCountDf.rename(columns={1: 'ResponseCount'}, inplace=True)

    # joining the path df with the participants df
    ExportUnitsExpectedCountNR = pd.merge(exportUnits_PathFinal, expectedCountDf, left_on='Path', right_on=0, how= 'left').drop(0, axis=1)
    dfAnonymityThreshold = pd.merge(ExportUnitsExpectedCountNR, responseCountDf, left_on='Path', right_on=0, how= 'left').drop(0, axis=1)

    # adding the No response values with the response count, to get the invited count
    dfAnonymityThreshold['InvitedCount'] = dfAnonymityThreshold['InvitedCountNR'] + dfAnonymityThreshold['ResponseCount']

    #dfAnonymityThreshold.sort_values(ManID).to_csv('PathReport.csv')

    # deleting invited count for non respondents, levelNum and Path columns
    del dfAnonymityThreshold['InvitedCountNR']
    del dfAnonymityThreshold['LevelNum']
    del dfAnonymityThreshold['Path']


    # calculating the anonymity threshold
    dfAnonymityThreshold['Anonymity Threshold'] = [True if x >= anonymity_threshold_value else False for x in dfAnonymityThreshold['ResponseCount']]

    '''SAVE DOCUMENT'''

    # Create Folder to save the report
    new_filename_to_save = Participants_file_path.split('/')
    
    # extract the file name to delete ".csv" string
    final_filename = new_filename_to_save[-1]

    # extract the number of characters to be deleted from the original path and assign to the new folder
    len_user_file = len(final_filename) + 1

    # delete the length of the user file to create the original path
    new_path_to_save = Participants_file_path[0:-len_user_file]

    # create folder to save the report
    reportPath = new_path_to_save + '/Anonymity_Report'
    if not os.path.exists(reportPath):
        os.makedirs(reportPath)

    #Save new DF as CSV
    def save_doc():
        today = datetime.datetime.today()
        timestamp_str = today.strftime("%Y%m%d_%H%M")
        # reordering the columns
        columnNamesList = dfAnonymityThreshold.columns.values.tolist()
        #print(columnNamesList)
        a, b = columnNamesList.index('ResponseCount'), columnNamesList.index('InvitedCount')
        columnNamesList[b], columnNamesList[a] = columnNamesList[a], columnNamesList[b]
        #print(columnNamesList)
        dfAnonymityThresholdReport = dfAnonymityThreshold.reindex(columns=columnNamesList)
        dfAnonymityThresholdReport.sort_values(level_Columns[0]).to_csv(f'{new_path_to_save}/Anonymity_Report/AnonymityReport{timestamp_str}.csv',index=False)
        return True
    
    #Call Save_doc function
    save_status = save_doc()
    end_time = time.time()        
    print("Execution time in seconds:", round(end_time-start_time,2), "seconds")  
    
    if save_status:
        messagebox.showerror("Success", "Doc was saved.")
    else:
        messagebox.showerror("Failed", "Doc was not saved due to an error.") 