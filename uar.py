# -*- coding: utf-8 -*-
"""
Created on Wed Feb 15 18:08:39 2023

@author: GINFA
"""

from win32com import client
import pyodbc
import pandas as pd
from datetime import datetime
import os
import sys
import base64
import xlsxwriter
import requests
import json
import warnings
warnings.filterwarnings("ignore")
import time
import pandas as pd
import pyodbc
import os
import glob
from datetime import datetime
############################################################################################################
################################################# SETTINGS #################################################
############################################################################################################
'''
The settings file will contain information on the server connections, sql queries, Power Automate and 
Power BI connections.
'''
try:
    settings = json.loads(open("settings.json").read())
except FileNotFoundError:
    print("Settings file not found. Please contact Global IT.")
    sys.exit(1)

############################################################################################################
############################################################################################################

class Connection:
    '''Class to execute a connection to the server and query back data.'''
    
    def __init__(self, serverSettings, authentification = "Windows"):
        
        self.settings = serverSettings
        
        assert (authentification == "Windows" or authentification == "Password"),\
            "Wrong authentification method, please choose between 'Windows' or 'Password"
        
        if authentification == "Windows":
        # If authentification method is "Windows", SQL engine will connect through Windows Authentification    
            auth = ";Trusted_Connection=yes;"
        
        # If authentification method is "Password", SQL engine will connect through AD Password
        elif authentification == "Password":
            
            auth = ";UID=" + self.settings["username"] + ";PWD=" + self.settings["password"]
            
        else:
            
            print("Cannot connect: wrong authentification.")
            sys.exit(1) 
            
        # Sql server connection details
        self.cnxn = pyodbc.connect("DRIVER={ODBC Driver 17 for SQL Server};SERVER=" + self.settings["server"] + ";DATABASE=" 
                              + self.settings["database"] + auth)
        
    
    def query(self):
        '''Execute a specific query given a query string and return a pandas DataFrame.'''
        
        return pd.read_sql(self.settings["query"], self.cnxn)
    
    
class Logger:
    '''Class to log and store a copy of the queried data for audit.
    Logger will log all the load data in individual files (CSV copies of the SQL queries); and it 
    will create a summary of all this data, registering as well if any approval is sent to an email.
    '''
    
    def __init__(self, timestamp, filepath, dfs: dict):
        
        # Get data and timestamp of data
        self.timestamp = timestamp
        self.filepath = filepath
        self.dfs = dfs # Dictionary of load datasets, with {name : dataframe}
        
        # SUMMARY
        # It will contain a summary of manager IDs with the # of employees and # of rows for each dataset 
        self.summary = pd.DataFrame() 

        # For each load dataset
        for name, df in self.dfs.items():
            
            # If the summary is empty (it means this is the first dataset)
            if self.summary.shape[0] == 0:
                self.summary = self.__summarize(df, name)
            # If the summary is not empty, then merge the with the new dataset
            else:
                self.summary = self.summary.merge(self.__summarize(df, name), on="Manager ID", how="outer")
     
        # We remove NANs and cast numbers to int
        #self.summary = self.summary.fillna(0)
        self.summary = self.summary.astype(int, errors="ignore")
        # Set default values
        self.summary["Sent Approval"] = "No" # Log if managers have been sent their report - By default NO
        self.summary["Email"] = ""
        self.summary["Date"] = ""
        self.summary["Status code"] = ""


    def __summarize(self, df, name):
        '''Private method to summarize a dataset.
        Takes a dataframe and name from the dictionary of datasets.
        Returns the summary of employees and rows for each manager.
        '''
        
        df_ = df.groupby("Manager ID").size().reset_index() # Nr of rows per manager
        df_.columns = ["Manager ID", f"{name} - Nr Rows"]
        df_[f"{name} - # Emp."] = df.groupby("Manager ID")["User ID"].nunique().reset_index(drop=True) # Nr of employees per manager
        
        return df_
        
        
    def logEmail(self, manager, email, date, status):
        '''Method to indicate an approval has been sent to an email.
        Takes as arguments the manager ID, email, date and status of the response.
        '''
        # Modify summary to account for email sent
        self.summary.loc[self.summary["Manager ID"] == manager, "Sent Approval"] = "Yes"
        self.summary.loc[self.summary["Manager ID"] == manager, "Email"] = email
        self.summary.loc[self.summary["Manager ID"] == manager, "Date"] = date
        self.summary.loc[self.summary["Manager ID"] == manager, "Status code"] = int(status)

        
    def dumpLogs(self):
        '''Method to dump logs to files.'''
        # Log all load datasets
        for name, df in self.dfs.items():
            # Log all queried data
            df.to_csv(self.filepath + f'/log-{name}-{int(self.timestamp)}.csv', index=False, sep=";", encoding="utf-8")

        # Log summary
        self.summary.to_csv(self.filepath + f"/summary-{int(self.timestamp)}.csv", index=False, sep=";", encoding="utf-8") 
        

        
def FileParser(filepath):
    '''Function to read and parse correctly the "User Permissions" text file from Axapta 2.5'''
            
    # We create an empty list to collect all rows
    rows = []
    # We get the name of columns
    cols = ["User ID", "Access"] + [f"Menu L{i}" for i in range(1, 9)]

    # Open file
    with open(filepath, encoding="utf-8", errors='ignore') as infile:
        
        # Read lines
        for line in infile.readlines():
            
            # If line is new user (the line starts with 'User permissions: @user')
            if "User permissions:" in line:        
                # Determine user and extract initials (the user initials are between parenthesis)
                usr = line[line.index("User permissions:") + len("User permissions:"):].split(",")[0]
                openPar = usr.find("(")
                fromRight = usr[openPar + 1:]
                closePar = fromRight.find(")")
                init = fromRight[:closePar].upper()
                
                # Determine a first set of empty permissions
                menus = [""] * 8
                
            # Access menu ignored
            elif line.startswith("Access\tMenu"):
                pass
                
            # Then add data
            else:
                # Split Level and Menu
                level, menu = line.split("\t")
                
                # Extract menu
                menu_formatted = menu.lstrip()
                
                # Determine the padding to assign column index
                pad = len(menu) - len(menu_formatted)
                ix = pad // 7
                
                # Put the menu in the right column and set all other sub-levels to empty
                menus[ix] = menu_formatted.strip()
                menus[ix+1:] = ["" for i in range(ix+1, 8)]
                
                # Determine the row and append to rows
                r = [init, level] + menus
                rows.append(r)
                
    data = pd.DataFrame(rows, columns=cols)
    data = data[data["User ID"] != ""] # Exclude the empty user
    data["User ID"] = data["User ID"].str.upper()
    
    return data
        

class ReportCreator:
    '''Class to create the managers reports invidivually or batch-wise'''  
    
    def __init__(self, responsible, settings, authentification="Windows", testing=False):
        '''Constructor. Takes as arguments:
            - responsible: (str) the email account from the person running the process. This will be the requestor of any potential
            Approval request.
            - settings: (dict) the settings dictionary previously load.
            - authentification: (str) either 'Windows' or 'Password'. Serves to authentificate to the server.
            - testing: (bool) indicates if the process is running in 'test mode' or not.
        '''
        self.responsible = responsible
        self.settings = settings
        self.authentification = authentification
        self.testing = testing
        
        # Get timestamp at creation time
        self.timestamp = datetime.timestamp(datetime.now())
        self.date = f"{datetime.fromtimestamp(self.timestamp).year}-{datetime.fromtimestamp(self.timestamp).month}"
        self.__internalLog(f"Instance created (Timestamp: {self.timestamp}) / Responsible: {responsible} / Authentification: {authentification} / Testing: {testing}")
        
        try:
            self.__loadData()
            self.__internalLog("Data loaded")
        except:
            self.__internalLog("Data loading failed.")
        
        # Create Directory to store reports
        self.filepath = settings["path"] + f"Reports/{self.date}"
        if not os.path.exists(self.filepath):
            os.makedirs(self.filepath,exist_ok=True)
        
        # Create log
        self.logger = Logger(timestamp = self.timestamp, filepath = self.filepath, dfs = self.masterData)      
        
        
    def __internalLog(self, message):
        '''Logs internal Python execution steps.'''
        
        if self.testing == True:
            print(message)
        
        if not os.path.exists(settings["path"] + "data/logs/"):
            os.makedirs(settings["path"] + "data/logs/",exist_ok=True)
        
        with open(settings["path"] + f"data/logs/{datetime.fromtimestamp(self.timestamp).strftime('%Y-%m-%d-%H-%M')}.log", "a+") as log:
            log.write(f"{datetime.now()}\t{message}\n")
        
        
    def __loadData(self):
        '''Loads the data for the report creation. All sources are listed in the Settings file along with the needed queries.'''
        
        # Get data from MANAGERS table
        mgrs = Connection(settings["server"]["MANAGERS"]).query()
        self.__internalLog("DATA-->Managers database read successfully.")
        # Retrieve all manager's emails
        self.mgr_emails = mgrs[["User ID", "Email"]]
        df2 = pd.DataFrame(mgrs)

        
 


        # Get data from USERS tables
        ## From Axapta 2009
        self.access2009 = Connection(settings["server"]["USERS-Ax2009"], authentification = self.authentification).query()
        self.__internalLog("DATA-->Axapta 2009 database read successfully.")
        ## From Axapta 2.5
        self.access25 = Connection(settings["server"]["USERS-Ax2.5"], authentification = self.authentification).query()
        self.__internalLog("DATA-->Axapta 2.5 database read successfully.")
        
        ### TEMP CUSTOM ADDING RKK ###
        if os.path.exists(settings["path"] + "data/RKK_UserInfoCompanyAccess.xlsx"):
            _ = pd.read_excel(settings["path"] + "data/RKK_UserInfoCompanyAccess.xlsx", sheet_name="UserCompanyAccess")
            self.access25 = pd.concat([self.access25, _])
            self.__internalLog("DATA-->Custom RKK file read successfully.")
        ###########################
        
        ## From permissions files
        self.permissions = pd.concat([FileParser(settings["server"]["USERS-Ax2.5"]["permissions1"]), FileParser(settings["server"]["USERS-Ax2.5"]["permissions2"])]) 
        self.__internalLog("DATA-->Permissions files read successfully.")
        
        # Merge on native User ID with permissions file
        users_ = self.access25.drop("Company Access", axis=1).drop_duplicates() # We don't want company in the permissions
        self.permissions = users_.merge(self.permissions, on="User ID", how="left")
        
        # Merge users with managers
        self.access2009 = self.access2009.merge(mgrs, on="User ID", how="left")
        self.access25 = self.access25.merge(mgrs, on="User ID", how="left")
        self.permissions = self.permissions.merge(mgrs, on="User ID", how="left")

        del mgrs
        self.__internalLog("DATA-->All merges successful.")
        
        ### TEMP CUSTOM MAPPING ###
        if os.path.exists(settings["path"] + "data/custom_map.csv"):
            for line in open(settings["path"] + "data/custom_map.csv").readlines():
                usr_id, mgr_id = line.strip().split(",")
                self.access2009.loc[self.access2009["User ID"] == usr_id.upper(), "Manager ID"] = mgr_id.upper()
                self.access25.loc[self.access25["User ID"] == usr_id.upper(), "Manager ID"] = mgr_id.upper()
                self.permissions.loc[self.permissions["User ID"] == usr_id.upper(), "Manager ID"] = mgr_id.upper()
        self.__internalLog("DATA-->Custom map file read successfully.")
        ###########################
        
        self.masterData = {"Ax2009": self.access2009, "Ax2.5": self.access25, "Permissions": self.permissions}
        
        # List of all managers IDs
        self.allManagerIDs = pd.concat([self.access2009["Manager ID"], self.access25["Manager ID"]]).sort_values().dropna().unique()
       # allManagerIDs.to_csv('C:/Users/svccomplianceid001/Desktop/file2.csv')
        
        return
            
            
    def createReports(self, managerString, Excel=False, send_approval=False):
        '''
        Create reports for a list of managers.
        managerString: (str) "all" to process all managers. Otherwise a comma-separated list of manager initials to process.
        Excel: (bool) indicates if an Excel version of the report should be created.
        send_approval: (bool) indicates if the Approval request should be sent to the manager.
        '''
        
        self.Excel = Excel
        
        # Parse the manager list string list of IDs
        managerList, notFound = self.__getMgrList(managerString)
        
        for nF in notFound:
            print(nF)
            self.__internalLog(f"Manager {nF} not in database.")
            
        self.__internalLog(f"Create reports - Options: Excel {Excel} / Approval {send_approval}")
        self.__internalLog(f"Managers: {','.join(managerList)}")
        
        # Get number of IDs to process
        df3 = pd.DataFrame(notFound)
        df3.to_csv('C:/Users/svccomplianceid001/Desktop/file3.csv')

        n = len(managerList)
        i = 0
        df4 = pd.DataFrame(managerList)
        df4.to_csv('C:/Users/svccomplianceid001/Desktop/file4.csv')
        
        # For each manager ID
        for manager in managerList:
              
            i += 1
            try:               
                email = self.__getMgrEmail(manager)
                attachments = self.__getAttachments(manager) # If last manager, include logs in attachments
    
                self.__internalLog(f"{i/n:.2%}" + (f" - Sending to {email}" if send_approval else ""))
                    
                # If we want to send the approval flow and it is not a test
                if send_approval:
                    status = self.__sendEmail(email, attachments)
                    self.logger.logEmail(manager, email, str(datetime.now()), status)
            
            except:
                self.__internalLog(f"Report creation failed at {manager}.")
         
        if (n != 0) and (not self.testing):
            # When all managers have been processed, update dashboards        
            self.__triggerDashboardUpdate()
                
        # Create log files
        self.logger.dumpLogs()
        
        # Log to internal log
        self.__internalLog("Process finished.")
        return

    
    def __getMgrList(self, mgrString):
        '''Private method to parse dot-separated manager IDs from string into list.'''
        
        # If managerList = "all" we retrieve all manager IDs
        if mgrString == "ALL":
            managerList = self.allManagerIDs

        # we split the manager IDs to a list
        else:
            managerList = mgrString.split(",")
            managerList = [x.strip() for x in managerList]
            
        # Filter out managers not in database
        notFound = list(set(managerList).difference(set(self.allManagerIDs)))
        managerList = list(set(managerList).intersection(set(self.allManagerIDs)))
            
        return managerList, notFound
    
    
    def __getAttachments(self, manager):
        '''Private method to retrieve user access reports from manager
            Attachments have a strict indexing:
                1st file is for Axapta-2009-Access
                2nd file is for Axapta-2.5-Access
                3rd file is for Axapta-2.5-Permissions
            '''
            
        # Start with empty list of attachments
        attachments = [""] * 3
        
        # If manager has employees with access to Axapta 2009, we create a report and assign it to attachments (1st file)
        if manager in list(self.access2009["Manager ID"]):
            attachments[0] = self.__individualReport(manager, axapta="2009")[0]

        # If manager has employees with access to Axapta 2.5, we create a report and assign it to attachments (2nd file: Access, 3rd file: Permissions)
        if manager in list(self.access25["Manager ID"]):
            attachments[1], attachments[2] = self.__individualReport(manager, axapta="2.5")

        return attachments
    
    
    def __getMgrEmail(self, manager):
        '''Private method to retrieve a manager's email address.'''
        
        # If manager ID has an email associated and it is not a test, we fetch the manager email to send Approval
        if self.mgr_emails["User ID"].eq(manager).any() and (not self.testing):
            email = self.mgr_emails[self.mgr_emails["User ID"] == manager].iloc[0]["Email"]
        # Else if the manager has no ID associated or it is a test, we send approval to the process responsible
        else:
            email = self.responsible
            
        return email
        
        
    def __individualReport(self, manager, axapta = "2009"):
        '''Private method to generate an Excel and PDF report for a manager.'''
        
        attachments = []
        self.axapta = axapta
        
        ############## CREATE XLSX REPORT ##############
        # Report filename to save        
        # Get the table for the specific manager, along with the nr of unique employees
        if self.axapta == "2009":
            reportFilename = self.settings["path"] + f"Reports/{self.date}/{manager}-{self.axapta}-{self.date}-{int(self.timestamp)}"
            dfs = {reportFilename : self.access2009}
            
        elif self.axapta == "2.5":
            reportFilename = self.settings["path"] + f"Reports/{self.date}/{manager}-{self.axapta}-{self.date}-{int(self.timestamp)}"
            reportFilename2 = self.settings["path"] + f"Reports/{self.date}/{manager}-permissions-{self.date}-{int(self.timestamp)}"
            dfs = {reportFilename: self.access25, reportFilename2: self.permissions}
            
        for reportFilename, df in dfs.items():
            table, nr_employees = self.__createReport(manager, df) 
            table.drop_duplicates(inplace=True)
            
            # Create writer
            writer = pd.ExcelWriter(f'{reportFilename}.xlsx', engine='xlsxwriter')
            
            # Write table
            table.to_excel(writer, sheet_name='Report', startrow=1, index=False, header=False)
            headers = table.columns
            
            # Apply Format
            self.__formatReport(writer, manager, nr_employees, headers)
            
            # Close File
            writer.close()
            
            ## CREATE PDF
            # Open Microsoft Excel
            excel = client.Dispatch("Excel.Application")
              
            # Read Excel File
            sheets = excel.Workbooks.Open(reportFilename + '.xlsx')
            work_sheets = sheets.Worksheets[0]
              
            # Convert into PDF File
            work_sheets.ExportAsFixedFormat(0, reportFilename + '.pdf') # 0 for TypePDF, 1 for TypeXPS
            excel.Workbooks.Close()
            
            attachments.append(reportFilename + '.pdf')
            
            self.__internalLog(f"{reportFilename} for {manager} created.")
            
            if not self.Excel:           
                os.remove(reportFilename + '.xlsx')
            
        return attachments
    
    
    def __sendEmail(self, email, attachments):
        '''Private method to trigger an Approval request for a given email address and attachments.'''
        
        self.__internalLog(f"Attempting to send Approval request to {email}.")
        
        # Encoded files
        encoded_files = []
               
        # For each PDF file in the attachments list
        for file in attachments[:3]:
            
            # If the file is not empty
            if file != "":
                
                # Read the byte-level data
                with open(file, 'rb') as f:
                    binary_data = f.read()
                    f.close()
                # Encode and append to list
                encoded_files.append(base64.b64encode(binary_data).decode())
                
            # Else the file is empty, so we pass an empty string
            else:
                encoded_files.append("")
        
        # Unpack the files
        file1, file2, file3 = encoded_files

        obj = {'emailAddress': email, 'requestor': self.responsible, 'ax2009': file1, 'ax25': file2, 'permissions': file3}
        
        return self.__safeRequestor(self.settings["approvals"]["url"], headers = self.settings["approvals"]["header"], json = obj, request_name = f"Approval for {email}")
        self.process_and_insert_data()
    def process_and_insert_data():
    # Database connection details
        server_name = 'sql-inventorymanagement-001.database.windows.net'
        db_name = 'GlobalRPA'
        user = 'SVCRPAPROD0022ID03'
        password = 'rd9X$bVp/Jcx#d*p3/9ED*6i'
    
        conn_str = f"DRIVER={{SQL Server}};Server={server_name};DATABASE={db_name};UID={user};PWD={password}"
    
        # Find the latest Excel files
        folder_path = '//dkrmed095.dkrmed.radiometer.rmg/citrixsandbox/RMED/SOX_AX_userreview/Automation/Reports/2024-11'
        file_2009 = max(glob.glob(os.path.join(folder_path, '*-2009-*.xlsx')), key=os.path.getmtime)
        file_permissions = max(glob.glob(os.path.join(folder_path, '*permissions*.xlsx')), key=os.path.getmtime)
    
        # Read data from Excel files
        df_2009 = pd.read_excel(file_2009)
        df_2005 = pd.read_excel(file_permissions)
    
        # Process data (add columns, clean data, etc.)
        df_2009["manager_email"] = "your_manager_email"  # Replace with actual manager email
        df_2009["job_date"] = datetime.now()
        df_2009.columns = df_2009.columns.str.replace(' ', '_')
    
        df_2005["manager_email"] = "your_manager_email"
        df_2005["job_date"] = datetime.now()
        df_2005.columns = df_2005.columns.str.replace(' ', '_')
        df_2005['User_Name'] = df_2005['User_Name'].str.replace(',', '_')
        df_2005 = df_2005.fillna('NA')
    
        # Connect to SQL Server
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
    
        # Insert data into SQL Server
        columns_2009 = ', '.join(df_2009.columns)
        insert_query_2009 = f"INSERT INTO dbo.uar_ax2009 ({columns_2009}) VALUES ({', '.join(['?' for _ in df_2009.columns])})"
        cursor.executemany(insert_query_2009, df_2009.values.tolist())
    
        columns_permissions = ', '.join(df_2005.columns)
        insert_query_permissions = f"INSERT INTO dbo.uar_ax2005 ({columns_permissions}) VALUES ({', '.join(['?' for _ in df_2005.columns])})"
        cursor.executemany(insert_query_permissions, df_2005.values.tolist())
    
        conn.commit()
        conn.close()
        
   
    
    
    def __triggerDashboardUpdate(self):
        '''Private method to trigger the Power BI dashboards update.
        '''
        
        self.__internalLog("Attempting to trigger Dashboard Update.")
                            
        
        # Logs to pass to Power Automate trigger
        logAttachments = {}
        
        # for each log, we append the latest-update timestamp and pass the buffer to the json
        for n, df in self.masterData.items():
            
            # We need to pass the latest-update datetime to the dashboard
            df["latest-update"] = 0
            df.loc[0,"latest-update"] = int(self.timestamp) # We only need one row
            
            # Include in the atachments the buffer of the CSV file
            logAttachments[n] = df.to_csv(encoding="utf-8", sep=";", index=False)
                
        obj = {'requestor': self.responsible, 'log2009': logAttachments["Ax2009"], 'log25': logAttachments["Ax2.5"], 'log25-permissions': logAttachments["Permissions"]}
        
            
        return self.__safeRequestor(self.settings["PBI-update"]["url"], headers = self.settings["PBI-update"]["header"], json = obj, request_name="Dashboard Update")
    
    
    def __safeRequestor(self, url, headers, json, attempt_limit = 10, request_name = ""):
        '''Private method to send an HTTP request attempt.'''
        
        attempts = 0
        status_code = 0
        
        while ((attempts < attempt_limit) and (status_code != 202)):
            
            time.sleep(2**attempts) # Time pause before attempting
            
            try:
                x = requests.post(url, headers = headers, json = json)
                status_code = x.status_code
                
                if status_code == 202:
                    self.__internalLog(f"Request {request_name} success.")
                    return status_code
            
            except:
                pass
            
            attempts += 1
            self.__internalLog(f"Request {request_name} failed (status code {status_code}). Remaining attempts: {attempt_limit - attempts}.")
            
        return status_code
                        
            
    def __createReport(self, manager, df):
        '''Private method to summarize the report information for a manager.
        Arguments:
            -manager: the manager ID
            -df: the dataframe to summarize
        '''
        
        if manager == "BLANK":
            # Filter the data for the manager and the relevant fields
            df = df[df["Manager ID"].isna()]
            
        else:        
            # Filter the data for the manager and the relevant fields
            df = df[df["Manager ID"] == manager]
            
        # Get the number of employees for the manager
        nr_employees = df['User ID'].nunique(dropna=True)
            
        df.sort_values(list(df.columns), inplace=True)
        df = df.iloc[:,:-2]
        
        self.__internalLog(f"Report for {manager} - Employees: {nr_employees} / Rows: {df.shape[0]}.")
        if  nr_employees == 0:
            # Filter the data for the manager and the relevant fields
            self.__internalLog(f"Report nor generated for {manager} - Employees count is : {nr_employees}")
            logger.info("No employees found for " + {manager}+ " with AX2009 or AX2.5 access.");
            logger.info(f"Report will not be generated for :{manager}.");
        
        return df, nr_employees

    
    
    def __formatReport(self, writer, manager, nr_employees, headers):
        '''Private method to format the report correctly.'''
        
        # Get book and sheet
        workbook = writer.book
        worksheet = writer.sheets['Report']
        
        # Set paper and layout options
        worksheet.set_paper(9) # A4
        worksheet.set_landscape()
        worksheet.set_margins(top=0.75, bottom=0.3, left=0.3, right=0.3)
        
        ### Set header and footer
        # For header we want: Left: Title - Center: Manager - Right: Radiometer Logo
        worksheet.set_header(f'&L&"Tahoma,Bold"&12User Access Review\nAxapta {self.axapta}&C&12Manager: {manager}&R&G', {'image_right': 'RM logo CMYK.jpg'})
        # For footer we want: Left: Creation timestamp - Center: # of Employees for the manager - Right: Page number and total pages
        worksheet.set_footer(f'&LCreated at: {datetime.fromtimestamp(self.timestamp)}&C# of Employees: {nr_employees}&RPage &P of &N')
        
        # Table header format
        header_format = workbook.add_format({'bold': True, 'font_color': 'black', 'font_size': 12,'text_wrap':'true',
                                             'align':'center', 'valign':'vcenter', 'top': 5, 'bottom': 5})
        
        # Cell formats - Center
        cell_format_c = workbook.add_format({'bold': False, 'font_color': 'black', 'left': 1, 'right': 1})
        cell_format_c.set_align('center')
        cell_format_c.set_align('vcenter')
        cell_format_c.set_shrink()
        
        # Cell formats - Left aligned
        cell_format_l = workbook.add_format({'bold': False, 'font_color': 'black', 'left': 1, 'right': 1})
        cell_format_l.set_align('vcenter')
        cell_format_l.set_shrink()
        
        # Cell formats - Left aligned - Text wrap
        cell_format_lw = workbook.add_format({'bold': False, 'font_size': 10, 'font_color': 'black', 'left': 1, 'right': 1})
        cell_format_lw.set_align('vcenter')
        cell_format_lw.set_text_wrap()
        
        # Cell formats - Amounts
        cell_format_a = workbook.add_format({'bold': False, 'font_color': 'black', 'left': 1, 'right': 1, 'num_format': '#,##0.00'})
        cell_format_a.set_align('center')
        cell_format_a.set_align('vcenter')
        
        # Set column widths
        worksheet.set_column('A:A', 8, cell_format_c) # ID
        if self.axapta == "2009":
            worksheet.set_column('B:B', 40, cell_format_l) # Name
            worksheet.set_column('C:C', 11, cell_format_c) # Group ID
            worksheet.set_column('D:D', 40, cell_format_l) # Group Name
            worksheet.set_column('E:E', 9, cell_format_c) # Company Access
            worksheet.set_column('F:F', 12, cell_format_a) # Purchase
            worksheet.set_column('G:G', 12, cell_format_a) # Inventory
        elif self.axapta == "2.5":
            cell_format_c.set_bottom(4)
            cell_format_l.set_bottom(4)
            cell_format_lw.set_bottom(4)
            worksheet.set_column('B:B', 25, cell_format_l) # Name
            worksheet.set_column('C:C', 9, cell_format_c) # Company Access
            worksheet.set_column('D:D', 10, cell_format_c) # Access
            worksheet.set_column('E:E', 10, cell_format_lw) # M1
            worksheet.set_column('F:F', 10, cell_format_lw) # M2
            worksheet.set_column('G:G', 10, cell_format_lw) # M3
            worksheet.set_column('H:H', 10, cell_format_lw) # M4
            worksheet.set_column('I:I', 10, cell_format_lw) # M5
            worksheet.set_column('J:J', 10, cell_format_lw) # M6
            worksheet.set_column('K:K', 10, cell_format_lw) # M7
            worksheet.set_column('L:L', 10, cell_format_lw) # M8
        
        # Write table headers
        for i in range(len(headers)):
            worksheet.write(0, i, headers[i], header_format)  # Cell is bold and italic.

    
#############################################################################################        
####################### COMMAND LINE INTERFACE ##############################################
def CLI():
    
       
    print("### WELCOME TO USER DATA MODEL REPORT CREATOR ###\n")
    
    responsible = ""
    
    while(responsible == ""):
        
        responsible = input("Please enter your email: ")
        
    opts = ["Create all reports", "Create a specific report"]
    
    for x, o in enumerate(opts):
        
        print(f"{x+1} - {o}")
    
    print("(Press Ctrl+C to Cancel)")
    
    option = False
    
    while(not option):
        report_type = input("\nSelect option: ")
        if (report_type == "1") or (report_type == "2"):
            option = True
    
    excel = input("Create Excel? (y/n) ").lower() == "y"
    approval = input("Send approval request? (y/n) ").lower() == "y"
    test = input("Is it a test? (y/n) ").lower() == "y"
    mgr = None
    
    if report_type == "1":
        mgr = "all"
    
    elif report_type == "2":
        mgr = ""
        while (mgr == ""):
            mgr = input("Enter the manager's initials: ")
            
    return responsible, excel, approval, test, mgr
#############################################################################################
#############################################################################################

# Parse commandline arguments
argc = len(sys.argv)
 
# If too many arguments, exit with error 
if argc > 6:
    print("WRONG USAGE: Too many arguments.")
    sys.exit(1)
    
# If all arguments are passed, parse them accordingly
elif (argc > 1):

    test = False if sys.argv[1] == "false" else True
    responsible = sys.argv[2]
    excel = True if sys.argv[3] == "true" else False
    approval = True if sys.argv[4] == "true" else False
    mgr = sys.argv[5]
    
# If no arguments are passed, get them from Command Line Interface
else:
    responsible, excel, approval, test, mgr = CLI()

# Create instance
rc = ReportCreator(responsible = responsible, settings = settings, testing = test)
# Create reports given the arguments
rc.createReports(mgr.upper(), Excel=excel, send_approval=approval)
sys.exit(0)
    

if test:
    print(f"Argc: {argc} | Responsible: {responsible} | Excel: {excel} | Approval: {approval} | Manager: {mgr} | Test: {test}")
    input("\n -- PAUSED -- \n(press any key to continue)")