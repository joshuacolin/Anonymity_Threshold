import pandas as pd
import time
import datetime
import tkinter as tk
import webbrowser
from tkinter import filedialog,messagebox, ttk
import os
import sys

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
    ManID = ManID_Spaces.lstrip()
    
    ########################### read files ############################################################
    # processing Participant File
    print("Reading files.....")
    Participants = pd.read_csv(f'{Participants_file_path}', dtype='str')

    # processing Export Unit file
    ExportUnitsDf = pd.read_csv(f'{ExportUnits_file_path}', dtype='str')

    # processing D&A file
    Responses = pd.read_csv(f'{Response_file_path}', dtype='str')

    ############################ Export Units File ################################################################
    print('')
    print("################################################################################################")
    print("Processing Export Units file...")
    exportUnits_IDS = ExportUnitsDf[['UnitName', 'UnitID', 'ParentUnitID', ManID]]

# removing blank employee ID participants
    exportUnits_NoNulls = exportUnits_IDS.dropna(subset=ManID)

# identifying Placeholders from Manager ID column
    print("Checking for Placeholders in the hierarchy...")
    Placeholders = exportUnits_NoNulls[exportUnits_NoNulls[ManID].str.contains('NoManager')]
    if len(Placeholders) > 0:
        print("Placeholders identified:")
        print(Placeholders[[ManID]])

        # converting placeholders to list
        PlaceholdersList = Placeholders[[ManID]].values.tolist()

        # converting EmployeeID from Participant File to list
        EmployeeIDList = Participants[[EmpID]].values.tolist()

        # checking if the placeholder information is in the Participant file
        for p in PlaceholdersList:
            if p in EmployeeIDList:
                print("Placeholder", p, " exists in Participant file" )
            else:
                # obtaining the row that matches the placeholder ID
                ExportUnitsFiltered = exportUnits_NoNulls.loc[exportUnits_NoNulls[ManID].isin(p)]
                # obtaining the parent unit id
                PlaceholderParentUnitID = ExportUnitsFiltered['ParentUnitID'].iloc[0]
                # converting the parent unit id into a list
                PlaceholderParentUnitIDList = PlaceholderParentUnitID.split()
                # looking for that parent unit id into the export units file, if it doesn't exist that placeholder is the top level
                ExportUnitsFilteredtParentUnit = exportUnits_NoNulls.loc[exportUnits_NoNulls['UnitID'].isin(PlaceholderParentUnitIDList)]
                # obtaining placeholder information
                # Name
                PlaceholderName = ExportUnitsFiltered['UnitName'].iloc[0]
                # Employee ID
                PlaceholderEmployeeID = ExportUnitsFiltered[ManID].iloc[0]
                # Email
                PlaceholderEmail = ''
                # Manager ID
                if (len(ExportUnitsFilteredtParentUnit) > 0):
                    PlaceholderManagerID = ExportUnitsFilteredtParentUnit[ManID].iloc[0]
                else:
                    PlaceholderManagerID = 'nan'

                # adding placeholder information into the participant file
                new_row = {'First Name': PlaceholderName, 'Last Name': PlaceholderName, 'Email': PlaceholderEmail, 'Respondent': 'false', EmpID: PlaceholderEmployeeID, ManID: PlaceholderManagerID}
                print(new_row)
                Participants = pd.concat([Participants, pd.DataFrame([new_row])], ignore_index=True)

                print("Information for Placeholder", p, " added to the Participant file")
    else: 
        print("No Placeholders identified")

# setting Manager ID to be strings
    exportUnits_Final = exportUnits_NoNulls[[ManID]].astype(str)

    print("All good with Export Units file")


    ############################ Participants File ################################################################
    # checking if column names exist in the Participant File
    print('')
    print("################################################################################################")
    print("Processing participant file...")
    if EmpID not in Participants:
        messagebox.showerror("KeyError", "Please check the Employee ID column provided.")
        return
    
    if ManID not in Participants:
        messagebox.showerror("KeyError", "Please check the Manager ID column provided.")
        return
    
    # removing blank employee ID participants
    Participants_NoNulls = Participants.dropna(subset=[EmpID])

    # obtaining the employee and manager IDs from the Participants file
    Participants_IDs = Participants_NoNulls[[EmpID, ManID, 'Respondent']]
    Participants_Metadata = Participants_NoNulls[['First Name', 'Last Name', 'Email', EmpID]]

    # filtering non respondent participants
    Participants_IDs_Filtered = Participants_IDs.loc[Participants_IDs['Respondent'] == 'true']

    # setting Employee ID and Manager ID to be strings
    Participants_InvitedCount = Participants_IDs_Filtered[[EmpID, ManID]].astype(str)
    Participants_Path = Participants_NoNulls[[EmpID, ManID]].astype(str)

    ################################################################################################################################
    # checking if all managers are in the participant file
    print('')
    print("################################################################################################")
    print("Checking for missing managers in the participant file")
    # removing blank managers
    managersListNoNulls = Participants_Path[ManID].loc[Participants_Path[ManID] != 'nan']

    # removing duplicates from managers list
    managersList = managersListNoNulls.drop_duplicates()

    # obtaining the list of employees
    employeesList = Participants_Path[[EmpID]]

    # joining managers and employees lists
    ManagerEmployeeMerge = pd.merge(managersList, employeesList, how='left', left_on=ManID, right_on=EmpID)

    # identifying missing managers
    MissingManagers = ManagerEmployeeMerge.loc[ManagerEmployeeMerge[EmpID].isnull()]

    # if there are missing managers, display an error message
    if len(MissingManagers) > 0:
        print("########################################################################################")
        print("Missing Managers:")
        print(MissingManagers)
        messagebox.showerror("Missing Managers", "There are managers missing in the Participants List."+'\nPlease review the console.')
        # finish the script
        sys.exit()

    print("All good with Participants file")

############################ Response File ################################################################
    print('')
    print("################################################################################################")
    print("Processing response file...")
# filtering partial responses or Finished = False, so only Finished = true is included in the subset
    response_Submitted = Responses.loc[(Responses['Finished'] == 'True') | (Responses['Finished'] == '1')]

    response_IDS = response_Submitted[[ResponseEmployeeID]]#.iloc[2:]

# removing blank employee id responses
    response_NoNulls = response_IDS.dropna(subset=[ResponseEmployeeID])

#replacing ".0" with nothing
    response_NoDecimals = response_NoNulls.replace('.0', '')

# converting employee ID to be string
    response_Final = response_NoDecimals[[ResponseEmployeeID]].astype(str)

# adding response column
    response_Final['Response'] = 'Yes'

    print("All good with Response file")

    ############################ Path generation ################################################################
    # assigning values of Employee ID and Manager ID columns into a list
    print('')
    print("################################################################################################")
    print("Creating the hierarchy...")
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
    #print(ParticipantsWithPath)

    # joining the response df with the participants df
    ParticipantsWithResponses = pd.merge(ParticipantsWithPath, response_Final, left_on=EmpID, right_on=ResponseEmployeeID, how= 'left')#.drop(ResponseEmployeeID, axis=1)
    ParticipantsWithResponses['Response'] = ParticipantsWithResponses[['Response']].fillna('No')

    # joining the path df with the Export Units df
    ExportUnitsWithPath = pd.merge(exportUnits_Final, pathDf, left_on=ManID, right_on=0, how= 'left').drop(0, axis=1)
    ParticipantsRenamed = Participants_Path
    ParticipantsRenamed.rename(columns={ManID: 'ManagerID_2'}, inplace=True)
    ExportUnitsFullInfo = pd.merge(ExportUnitsWithPath, ParticipantsRenamed, left_on=ManID, right_on=EmpID, how='left').drop(EmpID, axis=1)

    ############################ Calculating the Invited and Response Counts ################################################################
    # calculating the expected count - No Response
    print('')
    print("################################################################################################")
    print("Calculating the expected count and response count for direct reports...")
    # filtering participants who didn't respond the survey
    ParticipantsPathNoReponse = ParticipantsWithResponses.loc[ParticipantsWithResponses['Response'] == 'No']
    # filtering participants who responded the survey
    ParticipantsPathR = ParticipantsWithResponses.loc[ParticipantsWithResponses['Response'] == 'Yes']

    # obtaining the direct reports response count
    DirectReportsResponseCount = ParticipantsPathR.groupby(ManID).size().reset_index(name='DR_ResponseCount')

    # obtaining the direct reports partial invited count
    DirectReportsPartialInvitedCount = ParticipantsPathNoReponse.groupby(ManID).size().reset_index(name='DR_PartialInvitedCount')

    #obtaining the direct reports invited count
    DirectReportsDF = pd.merge(DirectReportsPartialInvitedCount, DirectReportsResponseCount, on= ManID, how='outer')
    # replacing nan values with 0
    DirectReportsDF = DirectReportsDF.fillna(0)
    DirectReportsDF['DR_InvitedCount'] = DirectReportsDF['DR_PartialInvitedCount'] + DirectReportsDF['DR_ResponseCount']

    # deleting invited count for non respondents
    del DirectReportsDF['DR_PartialInvitedCount']
    # merging the direct report information with the ExportUnitsDF
    ExportUnitsWithPathDirectReports = pd.merge(ExportUnitsFullInfo, DirectReportsDF, on=ManID, how='left')
    ExportUnitsWithPathDirectReports = ExportUnitsWithPathDirectReports.fillna(0)

    #ExportUnitsWithPathDirectReports.to_csv('ExportUnitsInfo.csv')

    # creating a DF to use the group by feature
    ExportUnitsToReduce = ExportUnitsWithPathDirectReports

    # creating a while loop to get all invited and response count for each manager
    print('')
    print("################################################################################################")
    print("Calculating the response and invited count for all reports...")
    count = 1
    while(len(ExportUnitsToReduce) > 1):
        # group by manager
        if 'ResponseCount' in ExportUnitsWithPathDirectReports.columns:
            ExportUnitsGrouped = ExportUnitsToReduce.groupby('ManagerID_2')['ResponseCount'].agg('sum').reset_index(name='ResponseCountSum')
            ExportUnitsGroupedInvited = ExportUnitsToReduce.groupby('ManagerID_2')['InvitedCount'].agg('sum').reset_index(name='InvitedCountSum')
        else:
            ExportUnitsGrouped = ExportUnitsToReduce.groupby('ManagerID_2')['DR_ResponseCount'].agg('sum').reset_index(name='ResponseCountSum')
            ExportUnitsGroupedInvited = ExportUnitsToReduce.groupby('ManagerID_2')['DR_InvitedCount'].agg('sum').reset_index(name='InvitedCountSum')
        # rename column
        ExportUnitsGrouped.rename(columns={'ManagerID_2': ManID}, inplace=True)
        ExportUnitsGroupedInvited.rename(columns={'ManagerID_2': ManID}, inplace=True)
        # replacing nan with 0
        ExportUnitsGrouped = ExportUnitsGrouped.fillna(0)
        ExportUnitsGroupedInvited = ExportUnitsGroupedInvited.fillna(0)
        # getting the reduced DF
        ExportUnitsReducedMid = pd.merge(ExportUnitsToReduce, ExportUnitsGrouped, on=ManID, how='right')
        ExportUnitsReduced = pd.merge(ExportUnitsReducedMid, ExportUnitsGroupedInvited, on=ManID, how='left')
        #print(ExportUnitsReduced)
        #removing nan rows
        #ExportUnitsReduced.dropna(subset=[ManID], inplace=True)
        #print(ExportUnitsReduced)
        
        ExportUnitsSum = ExportUnitsReduced[[ManID, 'ResponseCountSum']]
        ExportUnitsSumInvited = ExportUnitsReduced[[ManID, 'InvitedCountSum']]
        #print(ExportUnitsSum)
        # joining the responsecountsum to the responsecount column
        ExportUnitsSumUpdatedMid = pd.merge(ExportUnitsWithPathDirectReports, ExportUnitsSum, on=ManID, how='left')
        ExportUnitsSumUpdated = pd.merge(ExportUnitsSumUpdatedMid, ExportUnitsSumInvited, on=ManID, how='left')
        # filling nan values
        ExportUnitsSumUpdated = ExportUnitsSumUpdated.fillna(0)
        #print(ExportUnitsSumUpdated)
        # adding responsecount
        if 'ResponseCount' in ExportUnitsWithPathDirectReports.columns:
            ExportUnitsSumUpdated['ResponseCountFinal'] = ExportUnitsSumUpdated['ResponseCount'] + ExportUnitsSumUpdated['ResponseCountSum']
            ExportUnitsSumUpdated['InvitedCountFinal'] = ExportUnitsSumUpdated['InvitedCount'] + ExportUnitsSumUpdated['InvitedCountSum']
            # deleting old responsecount column
            del ExportUnitsSumUpdated['ResponseCount']
            del ExportUnitsSumUpdated['InvitedCount']
            ExportUnitsSumUpdated.rename(columns={'ResponseCountFinal': 'ResponseCount'}, inplace=True)
            ExportUnitsSumUpdated.rename(columns={'InvitedCountFinal': 'InvitedCount'}, inplace=True)
        else:
            ExportUnitsSumUpdated['ResponseCount'] =  ExportUnitsSumUpdated['DR_ResponseCount'] + ExportUnitsSumUpdated['ResponseCountSum']
            ExportUnitsSumUpdated['InvitedCount'] =  ExportUnitsSumUpdated['DR_InvitedCount'] + ExportUnitsSumUpdated['InvitedCountSum']
        # deleting responsecountsum columns
        del ExportUnitsSumUpdated['ResponseCountSum']
        del ExportUnitsSumUpdated['InvitedCountSum']
        # transfering updated information to exportunits original
        ExportUnitsWithPathDirectReports = ExportUnitsSumUpdated

        # updating dataframe to reduce
        ExportUnitsToReduce = pd.DataFrame()
        ExportUnitsToReduce = ExportUnitsReduced[[ManID, 'ManagerID_2', 'ResponseCountSum', 'InvitedCountSum']]
        ExportUnitsToReduce.rename(columns={'ResponseCountSum': 'ResponseCount'}, inplace=True)
        ExportUnitsToReduce.rename(columns={'InvitedCountSum': 'InvitedCount'}, inplace=True)
        #print(ExportUnitsWithPathDirectReports)
        print('-----------------------------------------------------------')
        print(count, len(ExportUnitsToReduce))
        count +=1

    print('')
    print("################################################################################################") 
    print("Printing Anonymity Report:")  
    print(ExportUnitsWithPathDirectReports.sort_values(ManID))
    
    dfAnonymityThreshold = ExportUnitsWithPathDirectReports

    # deleting partial count columns
    #del dfAnonymityThreshold['InvitedCountPartial']
    #del dfAnonymityThreshold['ResponseCountPartial']

    # calculating the anonymity threshold
    dfAnonymityThreshold['Anonymity Threshold'] = [True if x >= anonymity_threshold_value else False for x in dfAnonymityThreshold['ResponseCount']]

    # removing "nan/" from Path
    dfAnonymityThreshold['Path'] = dfAnonymityThreshold['Path'].str[4:]

    # adding the metadata fields to the report
    dfAnonymityThresholdReport = pd.merge(dfAnonymityThreshold, Participants_Metadata, left_on=ManID, right_on=EmpID, how= 'left').drop(EmpID, axis=1)
    #SAVE DOCUMENT

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
        dfAnonymityThresholdReport[['First Name', 'Last Name', 'Email', ManID, 'Path', 'DR_InvitedCount', 'DR_ResponseCount', 'InvitedCount', 'ResponseCount', 'Anonymity Threshold']].sort_values(ManID).to_csv(f'{new_path_to_save}/Anonymity_Report/AnonymityReport{timestamp_str}.csv',index=False)
        return True
    
    #Call Save_doc function
    save_status = save_doc()
    end_time = time.time()        
    print("Execution time in seconds:", round(end_time-start_time,2), "seconds")  
    
    if save_status:
        messagebox.showerror("Success", "Doc was saved.")
    else:
        messagebox.showerror("Failed", "Doc was not saved due to an error.") 