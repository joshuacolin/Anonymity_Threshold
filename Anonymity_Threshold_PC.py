import pandas as pd
import time
import datetime
import tkinter as tk
import webbrowser
from tkinter import filedialog,messagebox, ttk
import os

########################################## Execute the script ##########################################
def PC(ExportUnits_file_path, Response_file_path, Participants_file_path, Empl_Man_IDs, anonymity_threshold_value):
    start_time = time.time()
    
    # checking missing IDs column names
    if len(Empl_Man_IDs) < 2:
        messagebox.showerror("Missing Column Names", "Employee ID and Manager ID column names are required to perform the analysis.")
        return
    
    #Employee ID
    EmpID = Empl_Man_IDs[0]

    #response file - Employee ID
    if EmpID == 'Unique Identifier':
        ResponseEmployeeID = 'Participant Unique Identifier'
    else:
        ResponseEmployeeID = EmpID

    #Manager ID
    ManID_Spaces = Empl_Man_IDs[1]
    ManID = ManID_Spaces.strip()

    ########################### read files ############################################################
    # processing Participant File
    Participants = pd.read_csv(f'{Participants_file_path}', dtype='str')

    # processing Export Unit file
    ExportUnitsDf = pd.read_csv(f'{ExportUnits_file_path}', dtype='str')

    # processing D&A file
    Responses = pd.read_csv(f'{Response_file_path}', dtype='str')


    ############################ Participants File ################################################################
    # checking if column names exist in the Participant File
    if EmpID not in Participants:
        messagebox.showerror("KeyError", "Please check the Employee ID column provided.")
        return
    
    if ManID not in Participants:
        messagebox.showerror("KeyError", "Please check the Manager ID column provided.")
        return

    # obtaining the employee and manager IDs from the Participants file
    print(Participants)
    Participants_IDs = Participants[[EmpID, ManID, 'Respondent']]
    Participants_Metadata = Participants[['First Name', 'Last Name', 'Email', EmpID]]

    # removing blank employee ID participants
    Participants_NoNulls = Participants_IDs.dropna(subset=[EmpID])

    # replacing '' with 0 for participants without Manager ID
    #Participants_ManagerID_NotNull = Participants_NoNulls.fillna(0)

    # filtering non respondent participants
    Participants_IDs_Filtered = Participants_NoNulls.loc[Participants_NoNulls['Respondent'] == 'true']

    # setting Employee ID and Manager ID to be strings
    Participants_InvitedCount = Participants_IDs_Filtered[[EmpID, ManID]].astype(str)
    Participants_Path = Participants_NoNulls[[EmpID, ManID]].astype(str)

############################ Export Units File ################################################################
    exportUnits_IDS = ExportUnitsDf[[ManID]]

# removing blank employee ID participants
    exportUnits_NoNulls = exportUnits_IDS.dropna(subset=ManID)

# removing "No Manager" rows from Manager ID column
    exportUnits_IDs_Final = exportUnits_NoNulls[~exportUnits_NoNulls[ManID].str.contains('NoManager')]

# setting Manager ID to be strings
    exportUnits_Final = exportUnits_IDs_Final[[ManID]].astype(str)

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
    # assigning values of Employee ID and Manager ID columns into a list
    hierarchy = Participants_Path[[EmpID, ManID]].values.tolist()

    # obtaining the paths for all the employees
    manager = set()
    employee = {}

    for e,m in hierarchy:
        manager.add(m)
        employee[e] = m

    # recursively determine parents until child has no parent, appending parents at the beginning of the list
    def ancestors(p):
        return (ancestors(employee[p]) if p in employee else []) + [p]

    # creating the path for each employee
    path = {}
    for k in (set(employee.keys())):
        # creating one path for each employee
        pathstr = '/'.join(ancestors(k))
        # removing 0/ from the paths
        #pathStrClean = pathstr.replace(pathstr[:2], '',1)
        path[k] = pathstr

    # assigning the paths to a data frame
    pathDf = pd.DataFrame(list(path.items()))

    # renaming the "Path" column
    pathDf.rename(columns={1: 'Path'}, inplace=True)

    # joining the path df with the participants df
    ParticipantsWithPath = pd.merge(Participants_InvitedCount, pathDf, left_on=EmpID, right_on=0, how= 'left').drop(0, axis=1)

    # joining the response df with the participants df
    ParticipantsWithResponses = pd.merge(ParticipantsWithPath, response_Final, left_on=EmpID, right_on=ResponseEmployeeID, how= 'left').drop(ResponseEmployeeID, axis=1)
    ParticipantsWithResponses['Response'] = ParticipantsWithResponses[['Response']].fillna('No')

    # joining the path df with the Export Units df
    ExportUnitsWithPath = pd.merge(exportUnits_Final, pathDf, left_on=ManID, right_on=0, how= 'left').drop(0, axis=1)

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
    for index, row in ExportUnitsWithPath.iterrows():
        correctedPath = str(row['Path']) + '/'
        # calculating the expected count
        expectedCountNR = ParticipantsPathStrNR.count(correctedPath)
        expected_count_NR_dic[row[ManID]] = expectedCountNR
        # calculating the response count
        responseCount = ParticipantsPathStrR.count(correctedPath)
        response_count_dic[row[ManID]] = responseCount

        print(count, expectedCountNR, responseCount, correctedPath)
        count += 1

    # assigning the paths to a data frame
    expectedCountDf = pd.DataFrame(list(expected_count_NR_dic.items()))
    responseCountDf = pd.DataFrame(list(response_count_dic.items()))

    # renaming the "ExpectedCount" column
    expectedCountDf.rename(columns={1: 'InvitedCountNR'}, inplace=True)
    responseCountDf.rename(columns={1: 'ResponseCount'}, inplace=True)

    # joining the path df with the participants df
    ExportUnitsExpectedCountNR = pd.merge(ExportUnitsWithPath, expectedCountDf, left_on=ManID, right_on=0, how= 'left').drop(0, axis=1)
    dfAnonymityThreshold = pd.merge(ExportUnitsExpectedCountNR, responseCountDf, left_on=ManID, right_on=0, how= 'left').drop(0, axis=1)

    # adding the No response values with the response count, to get the invited count
    dfAnonymityThreshold['InvitedCount'] = dfAnonymityThreshold['InvitedCountNR'] + dfAnonymityThreshold['ResponseCount']

    #dfAnonymityThreshold.sort_values(ManID).to_csv('PathReport.csv')

    # deleting invited count for non respondents
    del dfAnonymityThreshold['InvitedCountNR']

    # calculating the anonymity threshold
    dfAnonymityThreshold['Anonymity Threshold'] = [True if x >= anonymity_threshold_value else False for x in dfAnonymityThreshold['ResponseCount']]

    # removing "nan/" from Path
    dfAnonymityThreshold['Path'] = dfAnonymityThreshold['Path'].str[4:]

    # adding the metadata fields to the report
    dfAnonymityThresholdReport = pd.merge(dfAnonymityThreshold, Participants_Metadata, left_on=ManID, right_on=EmpID, how= 'left').drop(EmpID, axis=1)
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
        dfAnonymityThresholdReport[['First Name', 'Last Name', 'Email', ManID, 'Path', 'InvitedCount', 'ResponseCount', 'Anonymity Threshold']].sort_values(ManID).to_csv(f'{new_path_to_save}/Anonymity_Report/AnonymityReport{timestamp_str}.csv',index=False)
        return True
    
    #Call Save_doc function
    save_status = save_doc()
    end_time = time.time()        
    print("Execution time in seconds:", round(end_time-start_time,2), "seconds")  
    
    if save_status:
        messagebox.showerror("Success", "Doc was saved.")
    else:
        messagebox.showerror("Failed", "Doc was not saved due to an error.") 