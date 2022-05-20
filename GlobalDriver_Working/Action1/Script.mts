Dim intEventReturn, intSummaryCount    
Dim intMasterCount, strInputFilePath
'Dim intLogOffStartCount, intLogOffEndCount
Dim strTestCaseStatus, strExecutionConditions
Dim strTestCaseName, intPassCount
Dim intStartRowNumber, intEndRowNumber
Dim intTestcaseMoveCount
Dim intFirstLoopStartCount, intFirstLoopEndCount
Dim intReportLoopStartCount, intReportLoopEndCount
Dim intEndLoopStartCount, intEndLoopEndCount
Dim bln_Flag 
Dim intTestStepCount
Dim Flg_LogVerify
Dim bln_Passfail
Dim strEnvi 
Dim intCallIterPass
Dim intCallIterFail 
Dim strServer
Dim strUsername

intTestStepCount = ""
Public strHTMLFileName,strFileURL,htmlresult,fso,intflag,intPass,intFail

intPass=0
intFail=0

SystemUtil.CloseProcessByName("EXCEL.exe")

Dim strInputFileName, strTestcase_Name, intTestcaseOrderNo

strUsername = Environment.value("UserName")
'msgbox strUsername

'strTestcase_Name = Environment.value("TestName")
'Capturing Testscript name from ALM
'Set strTestcase_Name = QCUtil.CurrentTest
'strTestcase_Name = strTestcase_Name.Name
'Environment.value("TestCaseName")=strTestcase_Name


'Fetch working environment and server to run,Either from test lab or from user environment vqariables
'Set strTestcase_Envi = QCUtil.CurrentTestSet
Set strTestcase_Envi = Nothing

If strTestcase_Envi is Nothing Then

    Set Obj=CreateObject("wscript.shell")
    Set objEnv=Obj.Environment("User")
    strServer=objEnv.Item ("Server")
    strEnvi=objEnv.Item ("Site")
    strApplication=objEnv.Item ("Application")
    strIEVersion=objEnv.Item ("IEVersion")
    strEnvironment=objEnv.Item ("Environment")
    If strServer="" or strEnvi ="" or strApplication="" Then
        Reporter.reportevent micFail ,"Select Server & Site","Server name\Working environment\Application are not choosen.Exiting the test "
        Exittest
    End If
    
    strServersplit=Split(strServer,";")
    If UBOUND(strServersplit)>1 OR UBOUND(strServersplit)<1 Then
        Reporter.reportevent micFail ,"Select Server","Server name and Working environment are not Matching.Exiting the test "
    Exittest 
    ElseIf UBOUND(strServersplit)=1 Then    
    Environment.Value ("Server")=strServersplit(0)
    Environment.Value("WorkingEnvironment")=strServersplit(1)
    End If
    Environment.Value ("Site")=strEnvi
    Environment.Value("Application")=strApplication
    Environment.Value ("IEVersion")=strIEVersion
    Environment.Value("Environment")=strEnvironment
else

    strServer = strTestcase_Envi.Field("CY_User_05")
    strApplication=strTestcase_Envi.Field("CY_User_06")
    strIEVersion=strTestcase_Envi.Field("CY_User_08")
    strEnvironment=strTestcase_Envi.Field("CY_User_07")
    strEnvi = strTestcase_Envi.Field("CY_User_04")
    If strServer="" or strEnvi ="" or strApplication="" Then
        Reporter.reportevent micFail ,"Select Server & Site","Server name\Working environment\Application fields are not choosen.Exiting the test "
        Exittest
    End If
    strServersplit=Split(strServer,";")
    If UBOUND(strServersplit)>1 OR UBOUND(strServersplit)<1 Then
        Reporter.reportevent micFail ,"Select Server","Server name and Working environment are not Matching.Exiting the test "
    Exittest 
    ElseIf UBOUND(strServersplit)=1 Then    
    Environment.Value ("Server")=strServersplit(0)
    Environment.Value("WorkingEnvironment")=strServersplit(1)
    End If
    Environment.Value ("Site")=strEnvi
    Environment.Value("Application")=strApplication
    Environment.Value("IEVersion")=strIEVersion
    Environment.Value("Environment")=strEnvironment
    
End If

bln_Flag = False

GetAUTRootfolder = Pathfinder.Locate("..\..\..\")
'msgbox GetAUTRootfolder
If INSTR(GetAUTRootfolder, "Scripts\") Then
	RelPath = Split(GetAUTRootfolder,"Scripts\")
	GetAUTRootfolder = RelPath(0)
	StrSettingspath = "..\..\..\..\Settings\"
 Else 
     
    StrSettingspath = "..\..\..\Settings\"
End If


'msgbox GetAUTRootfolder

'StrSettingspath = "..\..\..\Settings\"
'Loading Initialization  Settings
Environment.LoadFromFile StrSettingspath&"Env.xml"

Environment.Value("TOP_LEVEL") = GetAUTRootfolder
Environment.Value("SETTINGS_LEVEL") = GetAUTRootfolder&"Settings\"
Environment.Value("VBS_LEVEL") = GetAUTRootfolder&"Libraries\HPALibrary\"
Environment.Value("OR_LEVEL") = GetAUTRootfolder&"Object Repositories\"
Environment.Value("INI_LEVEL") = GetAUTRootfolder&"Initilization\"
Environment.Value("ENVIRONMENT_LEVEL") = GetAUTRootfolder&"Environment\"
Environment.Value("LOOP") = GetAUTRootfolder&"Loop\"
Environment.Value("MACRO") = GetAUTRootfolder&"Macro\"

Loadfunctionlibrary  Environment.Value("INI_LEVEL")&"InitializationSettings.vbs"
pathEnvironmentsheet=Environment.Value("ENVIRONMENT_LEVEL")&"EnvironmentCredentials.xls"
Loadfunctionlibrary  Environment.Value("VBS_LEVEL") & "GenericFunctionLibrary.qfl"
Loadfunctionlibrary  Environment.Value("VBS_LEVEL") & "ObjectVerification.vbs"

'Creating Local environment
localpath=Environment.Value ("LOCAL_LEVEL")
'Set qcc=qcutil.QCConnection 
username= strUsername 

strInputFilePath=CreateLocalEnvironment(localpath,PARAMETER("datasheetname"),username,Environment.Value ("Server"))
Environment.value("EnvInputfilepath")=strInputFilePath
Reporter.ReportEvent micDone ,"Test Information","datasheetused:"&PARAMETER("datasheetname") &"Server Used:"&Environment.Value ("Server")

'createtestdatanotepad()
strReportFilePath = Environment.Value("REPORT_LEVEL")
StartTime = Timer
stTime = Time
Environment.Value("StartTime") = StartTime
Environment.Value("StTime") = stTime
SysCurrentdate = Day(now) & "/" & month(now)  & "/" & year(now) & " " & Hour(Now)&":"&Minute(Now) 
Environment.Value("CurrentDate") = SysCurrentdate
                   OpenFile()
                   

'Loading input sheet

strInputFilePath = Environment.Value("LOCAL_LEVEL")&"Working\"&username&"\"&Environment.Value ("Server")&"\"& PARAMETER("datasheetname") 


'Verify the Input File and run the excel macro
DataTable.AddSheet("Summary")
DataTable.AddSheet("DDT")
DataTable.AddSheet("Loops_Summary")
'DataTable.AddSheet("Common")
DataTable.ImportSheet strInputFilePath,"Summary","Summary"
'DataTable.ImportSheet strInputFilePath,"Common","Common"
DataTable.ImportSheet strInputFilePath,Environment.Value("IEVersion"),"DDT"
DataTable.ImportSheet Environment.Value ("LOOP")&"Loop.xls","Sheet1","Loops_Summary"

Environment.value("valcolumn")=ValColumn ' fetching the G column value to fetch values on 
ValColumnPos Environment.value("valcolumn"),"DDT"
intSummaryCount = Datatable.GetSheet("Summary").GetRowCount

DataTable.GetSheet("Summary").SetCurrentRow(1)         


For intMasterCount = 1 To intSummaryCount    

    DataTable.GetSheet("Summary").SetCurrentRow(intMasterCount) 
    'DataTable.GetSheet("Summary").SetCurrentRow(intMasterCount) 
'	strTestCaseStatus = UCase(Trim(DataTable.Value("Status", "Summary")))
	strExecutionCondition = UCase(Trim(DataTable.Value("Execution_Condition", "Summary")))
    
'    	If strTestCaseStatus = "NO RUN" OR strTestCaseStatus = "FAIL" Then 
		If strExecutionCondition = "" OR strExecutionCondition = "Y" Then	

			'Summary Sheet values
			DataTable.GetSheet("Summary").SetCurrentRow(intMasterCount) 		
			strTestCaseName = Trim(DataTable.Value("TestCase_Name", "Summary"))
			    
			    strModuleName = Trim(DataTable.Value("Module_Name", "Summary"))
			    strTestcase_Name=strTestcaseName
			    Environment("TS_Name") = strTestCaseName
			    Environment.value("Module_Name") = strModuleName
			    createtestdatanotepad()
			    flag=1
			    
	    ElseIf strExecutionCondition = "N" Then
	    flag=0
			End if
			
If flag=0 AND intMasterCount=intSummaryCount  Then
setURL()
	CloseFile()
            
               Exit For 
End If

    'strTestcaseName = UCase(Trim(DataTable.Value("TestCase_Name", "Summary")))
If flag=1 Then

    'Test case start row no & Test case end row no
            DDTSheetRowCount=Datatable.GetSheet ("DDT").GetRowCount
            DDTSheetRowCount=DDTSheetRowCount+1 ' to match the no of rows in Excel sheet
       intStartRowNumber= GetRowNo(strTestcaseName,strInputFilePath,DDTSheetRowCount)
       intEndRowNumber=GetRowNo(strTestcaseName&"_END",strInputFilePath,DDTSheetRowCount)
        intLogOffStartCount=GetRowNo("LOGOFF",strInputFilePath,DDTSheetRowCount)
        intLogOffEndCount=GetRowNo("LOGOFF_END",strInputFilePath,DDTSheetRowCount)
           

  
  
            'Summary Sheet values
            DataTable.GetSheet("Summary").SetCurrentRow(intMasterCount) 
            ' intStartRowNumber = DataTable.Value("Start_Row_Number", "Summary")        
            'intEndRowNumber = DataTable.Value("End_Row_Number", "Summary")
          '  intLogOffStartCount = Trim(DataTable.Value("Logoff_Start", "Summary"))    
            ' intLogOffEndCount = Trim(DataTable.Value("Logoff_End", "Summary")) 
            intlaunchstartRowno=0
            intlaunchEndRowno=intLogOffStartCount -1


            'Assoicate Object repositories with test
            datatable.SetCurrentRow intMasterCount
            strORs = Trim(DataTable.Value("Required_OR", "Summary"))    
            initsetup = 0
            initsetup = Fn_Gen_Initialization(strORs)

            If initsetup = 1 Then
               intEventReturn = 0
               

            'Testcase execution loop
            For intTestcaseMoveCount = 1 To 1
    
                strSheetName = "DDT"
                DataTable.GetSheet(strSheetName).SetCurrentRow(intTestcaseMoveCount) 
                Datatable.AddSheet ("Environment_Summary")
                Datatable.ImportSheet pathEnvironmentsheet,strEnvi,"Environment_Summary"
               
           
    '<<<<<<<<< TestCase Part >>>>>>>>>>>
                intlaunchLoopStartCount =intlaunchstartRowno
                intlaunchLoopEndCount=intlaunchEndRowno
                intReportLoopStartCount = intStartRowNumber 
                intReportLoopEndCount = intEndRowNumber
                intEndLoopStartCount = intLogOffStartCount             
                intEndLoopEndCount = intLogOffEndCount 
                'launch part
                intEventReturn = Fn_Gen_ScriptEngine (intlaunchLoopStartCount+1, intlaunchLoopEndCount, strSheetName, strTestCaseName,strInputFilePath)
                'testcasepart
                intEventReturn = Fn_Gen_ScriptEngine (intReportLoopStartCount+1, intReportLoopEndCount-1, strSheetName, strTestCaseName,strInputFilePath)
                'End part
                 intEventReturn = Fn_Gen_ScriptEngine (intEndLoopStartCount+1, intEndLoopEndCount-1, strSheetName, strTestCaseName,strInputFilePath)
                If strTestcase_Envi is Nothing Then
                Reporter.ReportEvent micDone ,"ScreenshotPath","Path to screenshot on the local drive:"&Environment.Value("SCREENSHOT_PATH")
                
                 If intflag = 1 Then
  	intCallIterPass = 0
  Else
  intCallIterPass = 1
  End If
  'intCallIterFail = intFailflag
              
             
               If intCallIterPass = 1 Then
               intflag=0
                  	'strpassfail = "Pass"
               	         Environment.Value("PASSFAIL") = "PASS"
               	         intPass=intPass+1
               	         Environment.Value("PASSCOUNT")= intPass
               	        
                         strTCID = Environment.Value("TS_Name")
                         strFilename = Environment.Value("REPORTFILENAME")
                         TestResult()

                
               	Else if intCallIterPass = 0 Then
               	'strpassfail = "Fail"
               	 Environment.Value("PASSFAIL") = "FAIL"
               	 intFail=intFail+1
               	         Environment.Value("FAILCOUNT")= intFail
                         strTCID = Environment.Value("TS_Name")
                         strFilename = Environment.Value("REPORTFILENAME")
                         intflag=0
                         TestResult()

        
               Else 
                        Environment.Value("PASSFAIL") = "PASS"
                         intPass=intPass+1
               	         Environment.Value("PASSCOUNT")= intPass
                         strTCID = Environment.Value("TS_Name")
                         strFilename = Environment.Value("REPORTFILENAME")
                         intflag=0
                         TestResult()
              
               	
               
               End If

End if
               
                
                else
                    Set fso = CreateObject("Scripting.FileSystemObject")
                     If fso.FileExists (Environment.Value("ALMDocPath")) Then
                        Call UploadScreenshotstoALM(Environment.Value("ALMDocPath"))
                        Call UploadtestdatatoALM(Environment.value("TestDataFilename"))
                        
                     Else
                     '''''Reporter.ReportEvent micFail ,"Upload Screenshot to ALM","File with screenshot doesnt exist on the local drive:"&Environment.Value("ALMDocPath")
                    
                    End If
                End If
        

            Next
            else
                Reporter.ReportEvent micFail,"Initialization","Having issue in Initialization Settting"
            End if

            bln_Flag = True
            If intMasterCount=intSummaryCount Then
               Exit For    	
        End if
       ' Exit For
    End If
    
Next

If bln_Flag = False Then
    Reporter.ReportEvent micFail, "Test Script Name Mismatch", "Given Test Script Name Mismatch Error"
End If


Function ValColumn()

If Environment.Value("Application")="MHP" AND Environment.Value ("Environment")="SIT" then
    ValColumn="Value_SIT"
    
ElseIf Environment.Value("Application")="MHP" AND Environment.Value ("Environment")="UAT" Then
    ValColumn="Value_UAT"
ElseIf Environment.Value ("Application")="MHP" AND Environment.value("Environment")="ICD10" Then
    ValColumn="Value_ICD10"
ElseIf Environment.value("Application")="MHP" AND Environment.value("Environment")="TRNG" Then
    ValColumn="Value_TRNG"
ElseIf Environment.value("Application") ="MHP" AND Environment.Value ("Environment")="DEV" Then
    ValColumn="Value_DEV"
 ElseIf Environment.value("Application") ="QNXT" AND Environment.Value ("Environment")="SDC_TEST" Then
  ValColumn="Value"
 ElseIf Environment.value("Application") ="QNXT" AND Environment.Value ("Environment")="HPA_TEST" Then
  ValColumn="Value"
 
else
     ValColumn="Value"
End If

    
End Function

Function ValColumnPos(ValColumn,strsheet)
    
    Col_Count=Datatable.GetSheet (strsheet).GetParameterCount 
    For i = 1 To Col_Count Step 1
        strColumnname=datatable.GetSheet(strsheet).GetParameter(i).Name
        If strColumnname=ValColumn Then
            Environment.value("DynColumnNo")=i
            Valcolumnflag=1
            Exit for
        End If
        
    Next
    
    If ValColumnflag<>1 Then
    Reporter.ReportEvent micFail ,"Find Value Column","Unable to find environment column name:"&ValColumn&"in the Datasheet"
        exittest
        
    End If
    
    
    
End Function

Function GetRowNo(strTestcasename,strDatasheetPath,intRowcount)

Set excel = createobject("excel.application")
Set workbook=excel.workbooks.open(strDatasheetPath)
	
With workbook.worksheets(Environment.Value("IEVersion")).Range("A1:"&"A"&intRowCount)
Set findobj=.find(strTestcasename)
If not findobj is nothing  Then
	
y=findobj.value
y=ucase(y)
If y<>strTestcasename Then
	do	
		Set findobj = .FindNext(findobj)
		z=findobj.value
		rowno=findobj.row
		z=ucase(z)
 		If z=ucase(strTestcasename) then 
 			flag=1
 			Exit do
 		End if 

	Loop 
Elseif y=ucase(strTestcasename) then
	rowno =findobj.row
	flag=1
End If
else
reporter.ReportEvent micFail ,"Find Row No","Failed to find the row no of the test case:"&strTestcasename
End If
End With
workbook.Save 
workbook.Close 
Set excel =nothing

If flag=1 Then
	
	Reporter.ReportEvent micPass ,"Find Row No","Row No of"&strTestcasename&":"&rowno
	GetRowNo=rowno-1
Else 
	Reporter.ReportEvent  micFail ,"Finding Row No of test case:","Unable to find the  row no of :"&strTestcasename
	exittest
End If	
		
	
End Function


Function setURL()

flag=0
    datatable.SetCurrentRow 1
    EnvsheetRowCount=datatable.GetSheet("Environment_Summary").GetRowCount
    For i = 1 To EnvsheetRowCount Step 1
            datatable.SetCurrentRow i
            GetServername=Datatable.value("ServerName","Environment_Summary")
            GetApplicationName=Datatable.Value ("Application","Environment_Summary")
                If Trim(UCASE(GetServername))=Trim(UCASE(strServer)) Then
                    If Trim(UCASE(GetApplicationName))=Trim(UCASE(strApplication)) Then
                    getUrl=Datatable.value("URL","Environment_Summary")
                    Environment.Value("URL") = getUrl
                    'flag=1
                    Exit for        
                    End If
                
                else
                End If
    Next

End Function


''______________________________________________________________________________
'
' Function: Fn_Gen_ScriptEngine
' Author: Vadi
' Creation Date: 03/29/2011
' Description: 
' Input Parameters: intStartCount - initial loop limit
'                    intEndCount - loop end limit
'                    strSheetName - Data Table Sheet Name
'                    strReportName - Report type name
'                    strTestCaseName - Test case name
' Output Parameters: N/A
'
'_________________________________________________________________________________

Function Fn_Gen_ScriptEngine ( intStartCount, intEndCount, strSheetName, strTestCaseName, strInputFilePath)
	
    Fn_Gen_ScriptEngine = 0
    intCallReturn = 0

    For intTestStepCount= intStartCount To intEndCount step 1
        
        DataTable.GetSheet(strSheetName).SetCurrentRow(intTestStepCount)
        Environment.value("CurrentRow")=intTestStepCount        
    

        'DataTable values assigning
        A = DataTable.Value("Browser_Name",strSheetName)
        B = DataTable.Value("Page_Name", strSheetName)
        C = DataTable.Value("Frame_Name", strSheetName)
        D = Trim(DataTable.Value("Object_Name", strSheetName))
        E = Trim(UCase(DataTable.Value("Object_Type", strSheetName)))        
        F = DataTable.Value("Event", strSheetName)
        
        If Environment.Value ("Application")="MHP" Then
            G=    cstr(Trim(DataTable.Value(Environment.value("valcolumn"), strSheetName)))
            
        elseif environment.value("Application")="TPA" Then
            G=    cstr(Trim(DataTable.Value(Environment.value("valcolumn"), strSheetName)))
            
         elseif environment.value("Application")="QNXT" AND Environment.Value ("Environment")="HPA_TEST" Then
            G=    cstr(Trim(DataTable.Value(Environment.value("valcolumn"), strSheetName)))  
            
         elseif environment.value("Application")="QNXT" AND Environment.Value ("Environment")="SDC_Test" Then
            G=    cstr(Trim(DataTable.Value(Environment.value("valcolumn"), strSheetName)))     
            
        else
        
        G = cstr(Trim(DataTable.Value("Value", strSheetName)))
        End If
        
        
    
        H = DataTable.Value("Report_Flag", strSheetName)    
        'ReportingMessage= Datatable.Value ("ReportingMessage",strSheetName)    
        If instr(D,"_")<>0 Then
            strOcc=instr(D,"_")
               strlen=len(D)
               strreqlen=strlen-strOcc
               ReportingMessage=Right(D,strreqlen)
                           
            
            else
            ReportingMessage=D
            
    
            
        End If
        
        'Report Flag verification
        If H = 0 Then
        
            reporter.ReportEvent micDone ,"Skip row","You have successfully skipped the row no "&intStartCount
        else
        
        If G <> "NA" Then
            Select Case E
        
        '# Generic Actions            
                Case "LAUNCH"    
				    strServer = Environment.Value ("Server")
				    strApplication = Environment.Value("Application")
        
					intCallReturn=LaunchURLfromEnvSheet(strServer,strApplication)
              

Case "LAUNCH_EDGE"    
				    strServer = Environment.Value ("Server")
				    strApplication = Environment.Value("Application")
        
					intCallReturn=LaunchURLandLoginfromEnvSheet_Edge(strServer,strApplication)

					
					
					
               Case "LAUNCH_VUE360"    
				    strServer = Environment.Value ("Server")
				    strApplication = Environment.Value("Application")
        
					intCallReturn=LaunchURLandLoginfromEnvSheet_VUE360(strServer,strApplication)
					
					
               Case "LAUNCH_VUE360URL"    
				    strServer = Environment.Value ("Server")
				    strApplication = Environment.Value("Application")
        
					intCallReturn=LaunchURLfromEnvSheet_VUE360(strServer,strApplication)

                Case "SUBLAUNCH"                        
                    intCallReturn = FN_Gen_WebApplicationStartUp(G)
                    
                                    Case "RUNQUERYDEV"
                    intCallReturn = Fn_Gen_MultiQuery_SQLDEV (conn,G,strTestCaseName,F,strInputFilePath,D)
                    
                                    Case "VERIFYDOWNLOAD"                        
                    intCallReturn = Fn_Gen_DownloadVerification(G)

                    Case "VERIFYEXCELDOWNLOAD"                        
                    intCallReturn = Fn_Gen_DownloadExcelVerification(G)

                    
                    Case "RANDOM9"
				    strLength = Right(E,1)
					intRowValue = intTestStepCount 
					intCallReturn =FN_Gen_GetRandomNumDDT(strLength,strInputFilePath,"DDT",intRowValue,Environment.value("DynColumnNo") )	
                    
                    Case "RANDOMNUM9"
				    strLength = Right(E,1)
					intRowValue = intTestStepCount 
					intCallReturn =FN_Gen_GetRandomNum(strLength,strInputFilePath,"DDT",intRowValue,Environment.value("DynColumnNo") )	

                    
				Case "SENDKEYSEDIT"
				     intCallReturn = Fn_Gen_WebEditSet(A,B,C,D,F,G)
                

			Case "OPTIONAL"						
					intCallReturn = Fn_Gen_OptionalObject(A,B,C,D,E,F,G)				
					
				
                Case "SENDKEYS"
                    intCallReturn = Fn_Gen_SendKeys(A,B,C,D,F,G)
                    
                Case "SELECTUMNODE"
                	intCallReturn = FN_Gen_SelectUMNode(A,B,C,D,E,G)

				Case "INSIGHTOBJECT" 
					intCallReturn = Fn_Gen_InSightObject(A,D,F,G)   	
					
                Case "APPLICATIONSYNC" 
                    intCallReturn = Fn_Gen_ApplicationSync(A)   

                Case "ACTIVATE" 
                    intCallReturn = Fn_Gen_Activate(A)                          
				Case "PARSEX12_837P"
					intCallReturn = ParseX12_837P(G)
                    
                Case "WAITANDVERIFY"            
                    intCallReturn = FN_Gen_WaitAndVerify (A,B,C,D,F)
                    
                 Case "BROWSER"            
                    intCallReturn = Fn_Gen_Browser (A,B,C,D,F)
                    
                 Case "NOTEXIST"            
                    intCallReturn = Fn_Gen_ObjectNotExist (A,B,C,D,F)
                                
				Case "WEBTABLEEDIT"	
					intCallReturn = Fn_Gen_WebTableEditSelection(A,B,C,D,F,G)			

             
				Case "PARSEX12_837I"
					intCallReturn = ParseX12_837I(G)

				Case "PARSEX12_837D"
					intCallReturn = ParseX12_837D(G)

				Case "PARSEX12_278"
					intCallReturn = Fn_Gen_ParseX12_278(G)

				Case "GETVALUEFROMLOOP"
				 	intRowValue = intTestStepCount 
					intCallReturn = GetValuefromLoop(intRowValue,G,strInputFilePath)		

				Case "GETVALUEFROMLOOPSERVICELINE"
				 	intRowValue = intTestStepCount 
					intCallReturn = GetValuefromLoopServiceLine(intRowValue,G,strInputFilePath)					

				Case "WAITANDVERIFY"			
					intCallReturn = FN_Gen_WaitAndVerify (A,B,C,D,F)


   				Case "GETVALUE"            
                     intRowValue = intTestStepCount 
                    intCallReturn = Fn_Gen_GetAppValue(intRowValue, Environment.value("DynColumnNo"), "DDT",A,B,C,D,F,strInputFilePath,G) 

				Case "GETVALUECOMMON"            
                     intRowValue = intTestStepCount 
                    intCallReturn = Fn_Gen_GetAppValueCommon(intRowValue, Environment.value("DynColumnNo"), "DDT",A,B,C,D,F,strInputFilePath,G)
                  

 				Case "GETVALUEWEBTABLE"            
                     intRowValue = intTestStepCount 
                    intCallReturn = Fn_Gen_GetValueWebTable(intRowValue, Environment.value("DynColumnNo"), "DDT",A,B,C,D,F,strInputFilePath,G) 

               Case "GETPOSITION"			
					 intRowValue = intTestStepCount 
					intCallReturn = Fn_Gen_GetAppPosition(intRowValue, 8, "DDT",A,B,C,D,F,strInputFilePath)     
					
					Case "GETUSER"			
					 intRowValue = intTestStepCount 
					intCallReturn = Fn_Gen_GetAppPosition(intRowValue, 8, "DDT",A,B,C,D,F,strInputFilePath)     

				Case "EXCELSPLIT"			
					 intRowValue = intTestStepCount 
					intCallReturn = Fn_Gen_EXCELSplit(intRowValue, 8, "DDT",F,strInputFilePath,G) 					
    
				Case "AUTOINCREMENT"			
					 intRowValue = intTestStepCount 
					intCallReturn = Fn_Gen_AutoIncrement(intRowValue, 8, "DDT", F, strInputFilePath)	
				
				Case "DATEVERIFICATION"			
					 intCallReturn = Fn_Gen_DATEFORMAT(F,G)					
	
                Case "WAITPROPERTY"
                    intCallReturn = Fn_Gen_WaitProperty(A,B,C,D,F,G)

                Case "CLOSE"                                
                    intCallReturn = FN_Gen_WebApplicationClose(A,B,C)  

                Case "BROWSERCLOSE"
                    intCallReturn = FN_Gen_WebApplicationBrowserClose()

                Case "SCREENSHOTS"
                    intCallReturn =    Fn_Gen_Screenshot_Capture(strTestcase_name)
                    
                Case "SCREENSHOT"
                
                    intCallReturn =    Fn_Gen_Screenshot_Word(A,G)

                Case "SCREENSHOT01"
                    intCallReturn =    Fn_Gen_Screenshot_Capture01(G)

                Case "WAIT"
                    intCallReturn = Fn_Gen_Wait (G)
                    
                    
                Case "VALIDATEPDF"
                    intCallReturn = Fn_Gen_ValidateReportPDF()
                    
                    
                Case "VALIDATEPROVPDF"
                    intCallReturn = Fn_Gen_ValidateProvReportPDF()
                    
                Case "VALIDATEENROLMENTPROVPDF"
                    intCallReturn = Fn_Gen_ValidateEnrolmentProvReportPDF(G)
                    
                
                Case "VALIDATEAUTHPDF"
                    intCallReturn = Fn_Gen_ValidateAuthReportPDF()

                    
                    
                Case "VALIDATEMEMPDF"
                    intCallReturn = Fn_Gen_ValidateMemReportPDF()
                    
                    
                Case "VALIDATEWCPDF"
                    intCallReturn = Fn_Gen_ValidateWCReportPDF(G)
                    
				 Case "DELETEEXISTREPORTPDF"
				     intCallReturn = Fn_Gen_DeleteLocalPDF(G)

                Case "LAUNCH FIREFOX"                        
                    intCallReturn = FN_Gen_LaunchFireFox(G)

                Case "Provider"    
                    intCallReturn = FN_Gen_Lnk_Provider(A,B,G)

                Case "WEBLIST"                               
                    intCallReturn = Fn_Gen_WebList(A,B,C,D,F,G)                
        
        
                Case "WEBLISTTEXT"                               
                    intCallReturn = Fn_Gen_WebListEnterText(A,B,C,D,F,G)
        
                Case "LINK"                                
                    intCallReturn = Fn_Gen_Link(A,B,C,D,F,G)    
    
                Case "WEBELEMENT"            
                    intCallReturn = Fn_Gen_WebElement(A,B,C,D,F,G)
                    
                Case "WEBELEMENTVALUE"
                	intCallReturn = Fn_Gen_WebElementValue (A,B,C,D,F,G,strInputFilePath)

                Case "WEBELEMENTVERIFICATION"            
                    intCallReturn =  Fn_Gen_WebElement_Verification(A,B,C,D,F)
                    
                Case "CUSTOMWEBELEMENT"            
                    intCallReturn =  Fn_Gen_CustomWebElement(A,B,C,F,G)
    
                Case "IMAGE"            
                    intCallReturn = Fn_Gen_Image(A,B,C,D,F,G)    
    
               Case "STATIC"            
                    intCallReturn = Fn_Gen_Static(A,B,C,D,F,G)
                    
                Case "WEBBUTTON"            
                    intCallReturn = Fn_Gen_WebButton(A,B,C,D,F,G)  
                    
                Case "WEBBUTTONEDGE"            
                    intCallReturn = Fn_Gen_WebButton(A,B,C,D,F,G) 
                    
                Case "INSIGHTOBJ"            
                    intCallReturn = Fn_Gen_Insight(A,B,C,D,F,G) 

Case "WEBRADIOINDEX"            
                    intCallReturn = Fn_Gen_WebRadioIndexSelect(A,B,C,D,F,G)  
                    
                    Case "WEBTABLERADIO"            
                    intCallReturn = Fn_Gen_WebTableRadioSelect(A,B,C,D,F,G)  


                Case "WINBUTTON"            
                    intCallReturn = Fn_Gen_WinButton(A,B,C,D,F,G)                    
    
                Case "WEBFILE"
                    intCallReturn = Fn_Gen_WebFile_Upd(A,B,C,D,F,G)
    
                Case "WEBEDIT"            
                    intCallReturn = Fn_Gen_WebEdit(A,B,C,D,F,G)
                    
				Case "SERVICELINEEDIT"
					intCallReturn = Fn_Gen_ServiceLineWebEdit(A,B,C,D,F,G)

                Case "WEBEDITITERATE"
                     intCallReturn = VerifyWOStatus(A,B,C,D,F,G)          

                Case "WEBEDITVERIFICATION"            
                    intCallReturn = Fn_Gen_WebEdit_Verification(A,B,C,D,F,G)

                Case "WINEDIT"    
                    'G = G & strTestCaseName & ".PDF"
                    intCallReturn = Fn_Gen_WinEdit(A,B,C,D,F,G)
                    
                    Case "WINEDITCLICK"    
                    'G = G & strTestCaseName & ".PDF"
                    intCallReturn = Fn_Gen_WinEditClick(G)
                    
                    
                    Case "WINEDITCLICKEDGE"    
                    'G = G & strTestCaseName & ".PDF"
                    intCallReturn = Fn_Gen_WinEditClickEdge(G)
                    
                    
                    Case "WINEDITSAVEASCLICK"    
                    'G = G & strTestCaseName & ".PDF"
                    intCallReturn = Fn_Gen_WinEditSaveAsClick(G)
                    
                    
                    Case "WINEDITSAVEASCLICKEDGE"    
                    'G = G & strTestCaseName & ".PDF"
                    intCallReturn = Fn_Gen_WinEditSaveAsClickEdge(G)
                    
                    
                    Case "WINEDITSAVEAS"    
                    'G = G & strTestCaseName & ".PDF"
                    intCallReturn = Fn_Gen_WinEditSaveAsPDF(A,B,C,D,F,G)
                    
                Case "PDFPAGEEXIST"
                    intCallReturn = Fn_Gen_PDFPageExist(A,B,C)
    
                Case "WEBCHECKBOX"
                    If InStr(D, "WtblChk") > 0 Then
                        intCallReturn = FN_Gen_WebTableCheckBox(A,B,C,D,F,G)  
                    Else
                        intCallReturn = FN_Gen_WebCheckBox(A,B,C,D,F,G)  
                    End If

				Case "WEBTABLEEDIT"	
					intCallReturn = Fn_Gen_WebTableSetEdit(A,B,C,D,F,G) 

                Case "WEBRADIOGROUP"                    
                    intCallReturn = FN_Gen_WebTableRadioButton(A,B,C,D,F,G) 

                Case "TABLEIMAGECLICK"    
                    intCallReturn = FN_Gen_PortalImageClick(A,B,C,D,F,G)

                Case "MEMBER TROUBLESHOOTER"
                    intCallReturn = FN_Gen_TroubleShooterRadioButton(A,B,C,D,F,G)  

                Case "PROVIDER TROUBLESHOOTER"
                    intCallReturn = FN_Gen_TroubleShooterRadioButton(A,B,C,D,F,G)  

                Case "PAYMENT TROUBLESHOOTER"
                    intCallReturn = FN_Gen_TroubleShooterRadioButton(A,B,C,D,F,G) 

                Case "WEBRADIOBUTTON"                    
                    intCallReturn = FN_Gen_WebRadioButton(A,B,C,D,F,G)                      

                Case "WEBTABLEFORM"            
                    intCallReturn = Fn_Gen_WebTableClaimsCodeEntry (A,B,C,D,F,G)                      
            
                Case "WEBGRIDFORM"            
                    intCallReturn = Fn_Gen_WebGridClaimsCodeEntry(A,B,C,D,F,G) 
                    
                Case "WBFLISTSELECT"
                	intCallReturn = Fn_Gen_Wbfgrid1(A,B,C,D,F,G)

                Case "WEBTABLEVERIFY"
                    intCallReturn = Fn_Gen_WebTable(A,B,C,D,F,G)      
                
                Case "WBFGRIDVERIFY"    
                    intCallReturn =Fn_Gen_Wbfgrid(A,B,C,D,F,G)
                    
                Case "WBFGRID"
                	intCallReturn =Fn_Gen_WebTable (A,B,C,D,F,G)

                Case "WEBTABLECHECKBOX"
                    intCallReturn = FN_Gen_WebCheckBoxTable(A,B,C,D,F,G)    

                Case "WBFGRIDCHECKBOX"
                    intCallReturn = FN_Gen_WbfGridCheckBox(A,B,C,D,F,G)        
                    
                Case "WBFGRIDRADIOBUTTON"
                    intCallReturn = FN_Gen_WbfGridRadioSelect(A,B,C,D,F,G)    

 				Case "WBFGRIDVALUE"
                    intCallReturn = FN_Gen_WbfGridValueSelect(A,B,C,D,F,G)   
                
                Case "WEBTABLECHECK"
                    intCallReturn = FN_Gen_WebCheckTable(A,B,C,D,F,G)   

                Case "UNCHECK"
                    intCallReturn = Fn_Gen_UnCheck(A,B,C,D,F,G)

                Case "SELECTRADIOBUTTON"
                    intCallReturn = Fn_Gen_WebRadioSelect(A,B,C,D,F,G)

                Case "WEBTABLEVALUEVERIFY"                    
                    intCallReturn = FN_Gen_WebTableValueCheck(A,B,C,D,F,G)
                    


				Case "WEBTABLEVALUECHECK"					
					intCallReturn = FN_Gen_WebTableValueCheck(A,B,C,D,F,G)
				
				Case "NEWWEBTABLEVALUEVERIFY"					
					intCallReturn = FN_Gen_WebTableValueVerify(A,B,C,D,F,G)
					
				Case "WBFGRIDVALUECHECK"					
					intCallReturn = FN_Gen_WbfGridValueCheck1(A,B,C,D,F,G)	
				
                Case "WBFGRIDVALUEVERIFY"                    
                    intCallReturn = FN_Gen_WbfGridValueCheck(A,B,C,D,F,G)                    

                Case "WEBTABLELINK"
                    intCallReturn = Fn_Gen_WebTableLink(A,B,C,D,F,G)  

                Case "WEBTABLELIST"
                    intCallReturn = Fn_Gen_WebTableList(A,B,C,D,F,G)  

                Case "WEBTABLEEDIT"
                    intCallReturn = Fn_Gen_WebTableEdit(A,B,C,D,F,G)
                    
                Case "DUPLICATESINWEBTABLE"
                    intCallReturn = Fn_Gen_DuplicateinWebTable(A,B,C,D,F,G)

                Case "CHECKWEBTABLE"
                    intCallReturn = Fn_Gen_CheckWebTable(A,B,C,D,F,G)

                Case "WEBTABLEELEMENT"
                    intCallReturn = Fn_Gen_WebTableElement(A,B,C,D,F,G)  

                Case "SENDKEYS"
                    intCallReturn = Fn_Gen_SendKeys(A,B,C,D,F,G)

                Case "APPLICATIONSYNC" 
                    intCallReturn = Fn_Gen_ApplicationSync(A)   

                Case "LONGAPPLICATIONSYNC" 
                    intCallReturn = Fn_Gen_LongApplicationSync(A)

                Case "WAITFORBROWSERPAGETOSYNC" 
                    intCallReturn = Fn_Gen_WaitForBrowserPageToSync(A,B)   

                Case "SIGNINUSER" 
                    intCallReturn = Fn_Gen_SignInUser(Environment.Value ("Server"),Environment.Value("Application"))   

				Case "SIGNINPMUSER" 
					intCallReturn = Fn_Gen_SignInPMUser(Environment.Value ("Server"),Environment.Value("Application"))  

                Case "SIGNOUTUSER" 
                    intCallReturn = Fn_Gen_SignOutUser(A,B)   

                Case "MEMBERSIGNIN" 
                    intCallReturn = Fn_Gen_MemberSignIn(Environment.Value ("Server"),Environment.Value("Application"))   

                Case "RADCOMBOBOX"            
                    intCallReturn = Fn_Gen_RadComboBox(A,B,C,D,F,G)
                Case "CLICKBROWSERDIALOGBUTTON" 
                    intCallReturn = Fn_Gen_ClickBrowserDialogButton(A,B,C,D,F,G) 


                Case "WEBEDITSELECT"	
					 intCallReturn = Fn_Gen_TelObjSelect(A,B,C,D,F,G)


                Case "WEBEDIT"            
                    intCallReturn = Fn_Gen_WebEdit(A,B,C,D,F,G)
'#Window Controls
                Case "WINEDIT"    
'                    G = G & strTestCaseName & ".PDF"
                    intCallReturn = Fn_Gen_WinEdit(A,B,C,D,F,G)

                Case "WINBUTTON"            
                    intCallReturn = Fn_Gen_WinButton(A,B,C,D,F,G)        
 
            '#Table Actions
            
                Case "WEBTABLE"
                    intCallReturn = Fn_Gen_WebTable(A,B,C,D,F,G)

				Case "WBFGRID"
					intCallReturn = Fn_Gen_WbfGrid(A,B,C,D,F,G)

                Case "WEBRADIOGROUP"                    
                    intCallReturn = FN_Gen_WebTableRadioButton(A,B,C,D,F,G)  
                    Flg_LogVerify = intCallReturn
                    
               Case "WEBRADIOGROUPSELECT"                    
                    intCallReturn = FN_Gen_Webradiogroupselect(A,B,C,D,F,G)  

                Case "WBFGRIDCHECKBOX"
                    intCallReturn = FN_Gen_WbfGridCheckBox(A,B,C,D,F,G)    

                Case "WEBTABLECHECKBOX"
                    intCallReturn = FN_Gen_WebCheckBoxTable(A,B,C,D,F,G)                    

                Case "WEBTABLEFORM"            
                    intCallReturn = Fn_Gen_WebTableClaimsCodeEntry(A,B,C,D,F,G)  
            
                    Case "WEBTABLEVERIFY"
                    intCallReturn = Fn_Gen_WebTable(A,B,C,D,F,G)                      

                Case "CHECKWEBTABLE"
                    intCallReturn = Fn_Gen_CheckWebTable(A,B,C,D,F,G)

                Case "WEBTABLECHECK"
                    intCallReturn = FN_Gen_WebCheckTable(A,B,C,D,F,G)        
                Case "WEBTABLEVALUEVERIFY"                    
                    intCallReturn = FN_Gen_WebTableValueCheck(A,B,C,D,F,G)    

                Case "WEBTABLEROWCLICK"
                    intCallReturn = Fn_Gen_WebTableRowClick(A,B,C,D,F,G)

                Case "WEBTABLECELLVERIFY"
                    intCallReturn = Fn_Gen_WebTableCellVerify(A,B,C,D,F,G)

                Case "WEBTABLEELEMENT"
                    intCallReturn = Fn_Gen_WebTableElement(A,B,C,D,F,G)  

                Case "WEBTABLELINK"
                    intCallReturn = Fn_Gen_WebTableLink(A,B,C,D,F,G)  

                Case "AFFILIATION"
                    intCallReturn = FN_Gen_AffiliationsSelection(A,B,C,G)

                Case "ACTIVATE" 
                    intCallReturn = Fn_Gen_Activate(A) 
                    
                Case "ACTIVATEBROWSER"
                    intCallReturn = Fn_Gen_ActivateBrowser(A,B)
                Case "BROWSERCOUNT"
                    If UCaseTrim((A)) = "CALL TRACKING"  Then
                        Set odesc = Description.Create()
                        odesc("micclass").Value = "Browser"
                        odesc("title").Value = "QNXT - Call Tracking Management.*"
                        set obj = Desktop.ChildObjects(odesc) 
                        ItemCount = obj.Count()
                        If ItemCount <> 0 Then
                            intCallReturn = 1
                            Reporter.ReportEvent micPass,"Browser Count Verification: ", ItemCount & "  Browser exists in desktop"
                        Else
                            intCallReturn = 0
                            Reporter.ReportEvent micFail,"Browser Count Verification: ", "No Browser exists in desktop"             
                        End If
                    End If
                
                
                Case "WAITANDVERIFY"            
                    intCallReturn = FN_Gen_WaitAndVerify (A,B,C,D,F)                    

                Case "CLOSE"                                
                    intCallReturn = FN_Gen_WebApplicationClose(A,B,C)  

                Case "SCREENSHOT"
                    intCallReturn =    Fn_Gen_Screenshot_Word(A,G)

                Case "SCREENSHOTS"
                    intCallReturn =    Fn_Gen_Screenshot_Capture()
                    
                Case "WEBTABLEROWCLICK"
                    intCallReturn = Fn_Gen_WebTableRowClick(A,B,C,D,F,G)

                Case "WEBTABLECELLVERIFY"
                    intCallReturn = Fn_Gen_WebTableCellVerify(A,B,C,D,F,G)

                Case "CHECKDATE"
                    intCallReturn = FN_Gen_WebTableChkOverrideDate (A,B,C,D,F,G)

                Case "TEMPLATECLICK"            
                    intCallReturn = Fn_Gen_UM_TemplateTree(G)

                Case "VERIFICATION"                    
                    intCallReturn = Fn_Gen_ObjectVerification(A,B,C,D,F,G)    

                Case "VALUEVERIFICATION"                    
                    intCallReturn = Fn_Gen_WebTableVerification(A,B,C,D,F,G)        
                    
                Case "WEBTABLEVERIFICATION"                    
                    intCallReturn = Fn_Gen_Wbfgrid_WebtableVerification(A,B,C,D,F,G)     
                    
                Case "GETACTIVATIONLINK"
                    intCallReturn = FN_Gen_GetActivationLink(G)
                    
				Case "WBFGRIDSORT"
					intCallReturn = Fn_Gen_WbfGridSort(A,B,C,D,F,G)
				
				Case "SORT"
					intCallReturn = Fn_Gen_Sort(A,B,C,D,F,G)				
				
                Case "WBFGRIDSORTORDER"
                    intCallReturn = Fn_Gen_WbfGridSortOrder(A,B,C,D,F,G)

                Case "DYNAMICCHECKBOX"
                    intCallReturn = FN_Gen_DynamicWebTableCheckBox(A,B,C,D,F,G)    

                Case "DYNAMICMEMBER"
                    intCallReturn = Fn_Dynamic_MemberDataSetup(A,B,C,D,F,G)    

                Case "DYNAMICPROVIDER"
                    intCallReturn = Fn_Dynamic_MemberDataSetup(A,B,C,D,F,G)    

                Case "DYNAMICWEBLIST"
                    intCallReturn = Fn_Gen_DynamicWebList(A,B,C,D,F,G)    

                Case "LINKCLICK"                    
                    intCallReturn = Fn_Gen_LinkVerification(A,B,C,D,F,G)                    


                Case "DYNAMICRADIOBUTTON"
                    intCallReturn = Fn_Gen_DynamicWebTableRadioButton(A,B,C,D,F,G)
                    
                 Case "DYNAMICWEBTABLELINK"
                    intCallReturn = Fn_Gen_DynamicWebTableLink(A,B,C,D,F,G)

                Case "PROCESSLOGCHECKBOX"
                    intCallReturn = Fn_Gen_ProcessLogWebTableCheckBox(A,B,C,D,F,G)

                Case "TABLECOLUMNPOSITION"
                    intCallReturn = Fn_Gen_TableColumnPosition(A,B,C,D,F,G)

                Case "WAITPROPERTY"
                    intCallReturn = Fn_Gen_WaitProperty(A,B,C,D,F,G)

                Case "WEBTABLE"
                    intCallReturn = Fn_Gen_WebTable(A,B,C,D,F,G)

                Case "CONTENTVERIFY"                    
                    intCallReturn = Fn_Gen_ContentVerification(A,B,C,D,F,G)

                Case "COLUMNALIGNMENT"
                    intCallReturn = Fn_Gen_ColumnAlignment(A,B,C,D,F,G)
                    
                Case "SORT"
                    intCallReturn = Fn_Gen_Sorting(A,B,C,D,F,G)

                Case "LOADING"
                    intCallReturn = Fn_Gen_WebTableLoading(A,B,C,D,F,G)

                Case "URGENTCASE"
                    intCallReturn = Fn_Gen_UrgentCaseVerify(A,B,C,D,F,G)

                Case "LETTERVERIFY"
                    intCallReturn = Fn_Gen_PDFLetterVerification(G)

                Case "PDFCLOSE"
                    intCallReturn = Fn_Gen_PDFClose()

                Case "EDITSSELECTION"
                    intCallReturn = Fn_Gen_Claim_EditSelection (A,B,C,D,F,G)

                Case "CALLTRACKINGUSERSELECTION"
                    intCallReturn = Fn_Gen_CallTracking_UserSelection (A,B,C,D,F,G)

                Case "DISTINCTPROVIDER"
                    intCallReturn = Fn_Gen_ProvDistinct (A,B,C,D,F,G)

                Case "PROVIDERSORT"
                    intCallReturn = Fn_Gen_ProviderSort (A,B,C,D,F,G)

                Case "SORTING"
                    intCallReturn = Fn_Gen_ResultSort (A,B,C,D,F,G)

                Case "MEMDESCSORT"
                    intCallReturn = Fn_Gen_Mem_Desc_Sort(A,B,C,D,F,G)

                Case "SERVHISTCODESORT"
                    intCallReturn = Fn_Gen_Mem_Desc_Sort(A,B,C,D,F,G)

                Case "WEBRADIOSELECT"
                    intCallReturn = Fn_Gen_Second_Radio_Select (A,B,C,D,F,G)

                Case "PROVPAGENUMBER"
                    intCallReturn = Fn_Gen_Prov_Page_Number ()

                Case "ROLESCHECKBOX"
                    intCallReturn = Fn_Gen_RolesCheckBox_Sel (A,B,C,D,F,G)

                Case "WAIT"
                    intCallReturn = Fn_Gen_Wait (G)
                Case "RANDOMDATE"
                    intRowValue = intTestStepCount 
                    intCallReturn =FN_Gen_GetDateBetweenDates(strInputFilePath,"DDT",intRowValue,Environment.value("DynColumnNo") )

			
                Case "LOGVERIFY"
'                    Flg_LogVerify = intCallReturn
'                    intCallReturn = Fn_Gen_LogVerificaitonFile(G,Flg_LogVerify)
                    'intCallReturn = Fn_Gen_LogVerificaitonFile(G,Flg_LogVerify,strTestCaseName, strInputFilePath)

                Case "RUN QUERY"
                    intCallReturn = Fn_Gen_Query_Execution (conn,G,strTestCaseName,F,D)

                Case "RUN MULTIQUERY"
                    intCallReturn = Fn_Gen_MultiQuery_Execution (conn,G,strTestCaseName,F,strInputFilePath,D)
                    
                Case "RUN MULTIQUERYCOMMON"
                    intCallReturn = Fn_Gen_MultiQuery_Execution_Common (conn,G,strTestCaseName,F,strInputFilePath,D)
                    
                Case "DBCONNECT"
                    intCallReturn = Fn_Gen_DB_CONNECT (conn,G)

                Case "WEBTABLECHECK"
                    intCallReturn = FN_Gen_WebCheckTable(A,B,C,D,F,G)

                Case "CHECKBOXVERIFY"          
                     intRowValue = intTestStepCount                             
                    intCallReturn =     Fn_Gen_CheckBox_Sel_Verify (A,B,C,D,F,G,intRowValue,strInputFilePath)

                 Case "RADIOGROUPVERIFY"          
                     intRowValue = intTestStepCount                             
                    intCallReturn =     Fn_Gen_RadioGroup_Sel_Verify (A,B,C,D,F,G,intRowValue,strInputFilePath)

                 Case "TABLEVERIFYVALUE"
                    intCallReturn = FN_Gen_WebtableServiceVerifcation (A,B,C,D,F,G)

                Case "TABLEELEMENTVERIFY"
                    intCallReturn = FN_Gen_WebtableElementVerifcation (A,B,C,D,F,G)
'                    intCallReturn = FN_Gen_WebtableElementVerifcation (A,B,C,D,F,G)

                Case "DBVALUEVERIFY"
                    intCallReturn = FN_Gen_DBValueVerify (F,strInputFilePath)

                Case "TABLEVERIFICATION"
                    intCallReturn = Fn_Gen_TableVerification(A,B,C,D,F,G)
                    
                Case "RADCOMBOBOX"        
                    intCallReturn = Fn_Gen_RadComboBox (A,B,C,D,F,G)
                    
                Case "MAILSIGNIN"
                    intCallReturn = FN_Gen_MailSignIn(G)

                Case "MAILSIGNOUT"
                    intCallReturn = FN_Gen_MailSignOut(G)
                Case "MAILSUBJECTVERIFICATION"
                    intCallReturn = FN_Gen_MailSubjectVerification(F,G)
                
                Case "BORDERCOLORVERIFY"
                    intCallReturn = Fn_Gen_BorderColorVerify (A,B,C,D,F,G)
                    
                Case "WAITFORBROWSERPAGETOSYNC" 
                    intCallReturn = Fn_Gen_WaitForBrowserPageToSync(A,B)
                    
                Case "WEBFILE"
                    intCallReturn = Fn_Gen_WebFile (A,B,C,D,F,G)
                    
                Case "EDI FILE UPDATE"
                    intCallReturn = FN_Gen_EdiFileUpdateusingDB(G,strInputFilePath)

                Case "EDI FILE NOUPDATE"
                    intCallReturn = FN_Gen_EdiFileNoUpdate(G,strInputFilePath)
    
                Case "X12ACKNOWLEDGECHECK"
                    intCallReturn = Fn_Gen_X12AcknowledgeCheck(A,B,C,D,F,G)
                    
                Case "X12ARCHIVE"
                    intCallReturn = Fn_Gen_X12Archive(A,B,C,D,F,G)
                    
                Case "X12TABLELINK"
                    intCallReturn = Fn_Gen_X12FileDownload(A,B,C,D,F,G)
                


                Case "WEBELEMENTVERIFICATION"   
                     intCallReturn =  Fn_Gen_WebElement_Verification(A,B,C,D,F)
                    Flg_LogVerify = intCallReturn


                Case "GETVALUE"            
                     intRowValue = intTestStepCount 
                    intCallReturn = Fn_Gen_GetAppValue(intRowValue, 8, "DDT",A,B,C,D,F,strInputFilePath)     

                Case "BATCHFILE"
                    intCallReturn = Fn_Gen_BatchFile(G)

                Case "GRIDSORTVERIFY"
                    intCallReturn = Fn_Gen_TableSortVerify (A,B,C,D,F,G)
                    
                    Case "TABLESORTVERIFY"
                    intCallReturn = Fn_Gen_TableSortVerify1 (A,B,C,D,F,G)

                Case "WEBTABLEEDIT"
                    intCallReturn = Fn_Gen_WebTableEdit(A,B,C,D,F,G)

                Case "WEBTABLELIST"
                    intCallReturn = Fn_Gen_WebTableList(A,B,C,D,F,G)

                Case "INCREMENTVALUE"            
                    intRowValue = intTestStepCount 
                    intCallReturn = Fn_Gen_IncrementValue(intRowValue, Environment.value("DynColumnNo"), "DDT",A,B,C,D,F,G,strInputFilePath)   
				
				Case "AUTOINCREMENT"			
					 intRowValue = intTestStepCount 
					intCallReturn = Fn_Gen_AutoIncrement(intRowValue, 8, "DDT", F, strInputFilePath)	

                Case "TABLEVERIFICATION"
                    intCallReturn = Fn_Gen_TableVerification(A,B,C,D,F,G)

                Case "GENERALWEBTABLECHECKBOX"
                    intCallReturn = Fn_Gen_GeneralWebTableCheckBox(A,B,C,D,F,G)

'###Smoke $$##
                'Case "DYNAMICCHECKBOX"
                '    intCallReturn = FN_Gen_DynamicWebTableCheckBox(A,B,C,D,F,G)    

                Case "DYNAMICMEMBER"
                    intCallReturn = Fn_Dynamic_MemberDataSetup(A,B,C,D,F,G)    

                Case "DYNAMICPROVIDER"
                    intCallReturn = Fn_Dynamic_MemberDataSetup(A,B,C,D,F,G)    

                Case "DYNAMICWEBLIST"
                    intCallReturn = Fn_Gen_DynamicWebList(A,B,C,D,F,G)    

            '    Case "DYNAMICRADIOBUTTON"
            '        intCallReturn = Fn_Gen_DynamicWebTableRadioButton(A,B,C,D,F,G)

                Case "PROCESSLOGCHECKBOX"
                    intCallReturn = Fn_Gen_ProcessLogWebTableCheckBox(A,B,C,D,F,G)

                Case "MEMBER TROUBLESHOOTER"
                    intCallReturn = FN_Gen_TroubleShooterRadioButton(A,B,C,D,F,G)  
                    Flg_LogVerify = intCallReturn
                Case "PROVIDER TROUBLESHOOTER"
                    intCallReturn = FN_Gen_TroubleShooterRadioButton(A,B,C,D,F,G)
                    Flg_LogVerify = intCallReturn

                Case "PAYMENT TROUBLESHOOTER"
                    intCallReturn = FN_Gen_TroubleShooterRadioButton(A,B,C,D,F,G)
                    Flg_LogVerify = intCallReturn

                Case "USERSELECTION"
                    intCallReturn = Fn_Gen_Call_UsersSelection(A,B,C,D,F,G)

                Case "SETDYNAMICVALUE"
                    intCallReturn = Fn_QNXT_Gen_ExternalEnrollmentCOBTemplate(A,B,C,D,E)
            

                Case "TEXTFILEVALIDATION"
                    intCallReturn = Fn_Gen_TextFile(G)
                    
                Case "VALUETOEXCEL"
                    intCallReturn = Fn_Gen_ValuetoExcel(A,B,C,D,F,G,strInputFilePath)
                    
                Case "FILESAVE"
                    intCallReturn = Fn_Gen_FileSaveAs(A)
                
                Case "XLDBCOMPARE"
                    intCallReturn = Fn_Gen_XLDBCountCompare (conn,G)
                
                Case "BROWSERBACK"
                    intCallReturn = Fn_Gen_BrowserBack(A)
                    
                    Case "FULLSCREEN"
                    intCallReturn = Fn_Gen_BrowserMaximize(A)
                    
                Case "SETTINGWEBPACKAGE"
                    intCallReturn = Fn_Gen_SettingWebPackage (F,G)
                
                Case "VANEDIUPDATE"
                    intCallReturn = Fn_Gen_VAN_FileUpdate(G,strInputFilePath)
                    
                Case "VANFILEDROP"
                    intCallReturn = Fn_Gen_VANFileDrop(G)
                


   				Case "X12FILEDROP"
					intCallReturn = Fn_Gen_X12FileDrop(G)
				
                Case "VANRESPONSECHECKTA1"
                    intCallReturn = Fn_Gen_VANResponseCheck_TA1(G,strInputFilePath)
                    
                Case "VANRESPONSECHECK999"
                    intCallReturn = Fn_Gen_VANResponseCheck_999(G,strInputFilePath)
                
                Case "UPDATEAUTOADJUST"
                    intcallReturn = Fn_Gen_Update_AutoAudjustment(strInputFilePath,G)
                    
              '  Case "OPTIONAL" COMMENTED ON 0929
               '     intCallReturn = Fn_Gen_OptionalObject(A,B,C,D,E,F,G)
                    
                Case "DBVERIFICATION"
                    'intCallReturn = Fn_Gen_DBVerification(G)
                    intCallReturn = Fn_Gen_DBVerification(strInputFilePath)    
                    
                Case "MAILCONTENTVERIFY"
'                    Msgbox strInputFilePath
                    intCallReturn = FN_Gen_MailContentVerification(strInputFilePath,G,F)
                    
                'Case "MAILCONTENTVERIFY"
'                    Msgbox strInputFilePath
'                    intCallReturn = FN_Gen_MailContentVerification(strInputFilePath,G)

                Case "WIN_CLOSE"                                
                    intCallReturn = FN_Gen_WinApplicationClose(A,B,C)
                    
                Case "WEBELEMENTINDEX"    
                    intCallReturn = Fn_Gen_WebElementIndex(A,B,C,D,F,G)
                    
                Case "WEBRADIOGROUPCONTINUE"                    
                    intCallReturn = FN_Gen_WebTableContinueRadioButton(A,B,C,D,F,G) 
                    
                Case "WIN_PDFREPORT"
                    intCallReturn = Fn_Gen_Win_PDFSaveInitiation()
                    
                Case "DYNAMIC MULTIQUERY"
                    intCallReturn = Fn_Gen_DynamicQuery_Execution (conn,G,strTestCaseName,F,strInputFilePath,D)
                    
                Case "COMPAREVALUES"
                    intCallReturn = Fn_CompareTwoValues(G)
                    
                 Case "GETCELLVALUE"
                    intRowValue = intTestStepCount
                    intCallReturn = Fn_GetCellValue(intRowValue, Environment.value("DynColumnNo"), "DDT",A,B,C,D,F,G,strInputFilePath)  
                    
                Case "COMPARETOOLTIP"
                    intCallReturn = fn_CompareToolTip(A,B,C,D,F,G)
                    
                Case "SPLITVALUES"
                     intCallReturn = Fn_SplitValues(intTestStepCount+1,G,strInputFilePath)
                   
                Case "WIN_PDFPAGEEXIST"
                    intCallReturn = Fn_Gen_Win_PDFPageExist(A,B,C)

				Case "WORDEXIST"
					intCallReturn = Fn_Gen_WordExist(A,B,C)
					
				Case "WINMENULIST"								
					intCallReturn = Fn_Gen_WinMenuList(A,B,C,D,F,G)  
					 
				Case "INSIGHTOBJECT"
					intCallReturn = Fn_Gen_InSightObjectButtonClick(D)
				
				Case "FIREFOXDIALOG"	
					intCallReturn = Fn_Gen_FirefoxDialogHandle(A,F)
					
				Case "RADCOMBOBOX_SELECT"		
					intCallReturn = Fn_Gen_RadComboBox_Select (A,B,C,D,F)
					
				Case "EXECUTEBATFILE" 
					intCallReturn = Fn_Execute_BatFile (strTestCaseName,G)
					
				Case "PARSEX12_837P"
					intCallReturn = ParseX12_837P(G)
					
				Case "PARSEX12_837I"
					intCallReturn = ParseX12_837I(G)
					
				Case "PARSEX12_837D"
					intCallReturn = ParseX12_837D(G)
					
				Case "PARSEX12_278"
					intCallReturn = Fn_Gen_ParseX12_278(G)
					
				Case "GETVALUEFROMLOOP"
				 	intRowValue = intTestStepCount 
					intCallReturn = GetValuefromLoop(intRowValue,G,strInputFilePath)		
						
				Case "GETVALUEFROMLOOPSERVICELINE"
				 	intRowValue = intTestStepCount 
					intCallReturn = GetValuefromLoopServiceLine(intRowValue,G,strInputFilePath)
				
                Case "GENERATEX12"
                	intRowValue = intTestStepCount 
                	intCallReturn=Fn_Gen_GenerateX12 (intRowValue,G,strInputFilePath)
                
                
                Case "FIREFOXDIALOG"    
                    intCallReturn = Fn_Gen_FirefoxDialogHandle(A,F)
                    
                Case "RADCOMBOBOX_SELECT"        
                    intCallReturn = Fn_Gen_RadComboBox_Select (A,B,C,D,F)
                    
				Case "GETPOSITION"			
					 intRowValue = intTestStepCount 
					intCallReturn = Fn_Gen_GetAppPosition(intRowValue, 8, "DDT",A,B,C,D,F,strInputFilePath) 

				Case "EXCELVALUEVERIFY" 
					intCallReturn = FN_Gen_ExcelValueVerify(F,G)
				
				Case "EXCELVALUEVERIFYMULTIPLE" 
					intCallReturn = FN_Gen_ExcelValueVerifyMultiple(F,G)				
           
                    
                Case "EXECUTEBATFILE" 
                    intCallReturn = Fn_Execute_BatFile (strTestCaseName,G)
                    
                 
					
				Case "EXCELVALUEVERIFYNAME" 
					intCallReturn = FN_Gen_ExcelValueVerify1(F,G)
					
				'Case "OPTIONAL"  '' commented on 12/4
					'intCallReturn = Fn_Gen_OptionalObject(A,B,C,D,E,F,G)

  				Case "DYNAMICLINK"
					intCallReturn = Fn_Gen_DynamicLink(A,B,C,D,F,G)
					
				Case "UMDYNAMICLINK"
					intCallReturn = Fn_Gen_UMDynamicLink(A,B,C,D,F,G)	
				
                Case "DYNAMICVALUECAPTURE"
                    intRowValue = intTestStepCount
					intCallReturn = Fn_Gen_DynamicValueCapture(intRowValue,8, "DDT",A,B,C,D,F,strInputFilePath,G)	

			    Case "WEBTABLEENDLINK"
                    intCallReturn = Fn_Gen_WebTableEndLink(A,B,C,D,F,G)  			
                    
                Case "CALLFUNCTION"                    
                    intStartRowNumber = intTestStepCount+1
                    intCallReturn = FN_Gen_CallFunction(F,G,strInputFilePath)
                    
                Case "CALLLOGOFFFUNCTION"                    
                    intStartRowNumber = intTestStepCount+1
                    intCallReturn = FN_Gen_CalllogoffFunction(F,G,strInputFilePath)    
                    
'                    If intCallReturn = 1 Then                         
'                        strTestCaseName = Trim(DataTable.Value("TestCase_Name", "Summary"))
'                        intEndRowNumber = DataTable.Value("End_Row_Number", "Summary") 
'                        If Cint(intStartRowNumber) < Cint(intEndRowNumber) Then
'                            Call Fn_Gen_ScriptEngine (intStartRowNumber, intEndRowNumber, "DDT", strTestCaseName, strInputFilePath)                                                
'                            intCallReturn = 1
'                        End If             
'                    End If
'
                Case "COMPARECOLUMNVALUES"
                	intCallReturn=Fn_Gen_CompareColumnValues(A,B,C,D,F,G,strInputFilePath)
                	
                Case "XMLVERIFICATION"
					intCallReturn =Fn_Gen_XMLVerification(A,B,C,D,F,G)	
					

				Case "WBFVALUECHECK"
					intCallReturn = FN_Gen_WbfgridValueCheck(A,B,C,D,F,G) 

				Case "ACTIVEWINDOWCLOSE"
					intCallReturn=FN_Gen_CloseWindow()
				Case "TABOUT"
					intCallReturn=FN_Gen_TabOut()
				Case "GETROWVALUE"
					intRowValue = intTestStepCount 
                    intCallReturn=FN_Gen_GetRowValue(intRowValue,"DDT",A,B,C,D,F,strInputFilePath,G)
					
				Case "EMPTYTABLE"
					intCallReturn = FN_Gen_EmptyTable(A,B,C,D,F,G)	
				
				Case "COLUMNCHECK"
					intCallReturn = FN_Gen_WebTableColumnCheck(A,B,C,D,F,G)				
						    
                Case Else
				  'exittest 
				  Reporter.ReportEvent micFail ,"Keyword Not found","Keyword:"&E&" "&"Not Found"
				  intCallReturn = 0
				  'exittest
                  
            End Select
        End If
        
        If G<>"" and D<>"NA" Then
            call Writetestdata(ReportingMessage,G)
        End If
        'writing test data to notepad
        
            
        
        'Check test step FAIL status - skip current iteration            
        If intCallReturn = 0 Then    
            If G="" Then
                Reporter.ReportEvent micFail,ReportingMessage, "Failed at Row No:" &  intTestStepCount 
                Environment.Value("ReportMsg") = ReportingMessage&" : "&D&" : " &E &" Failed at Row No:" &  intTestStepCount
                
                else
                Reporter.ReportEvent micFail,ReportingMessage, "Failed at Row No:" &  intTestStepCount &",Value used:"&G
                  Environment.Value("ReportMsg") = ReportingMessage&" : "&D&" : "&E &" Failed at Row No:" &  intTestStepCount &",Value used:"&G
                  
            End If
            
        bln_Passfail=True
          If bln_Passfail=True Then
          	intflag=1
          End If
            
     
            Fn_Gen_ScriptEngine = 0
			'Call Fn_Gen_CloseBrowser			
            Exit Function
        End If        


    If intCallReturn = 1 Then            
        Fn_Gen_ScriptEngine = 1        
        If G="" Then
            Reporter.ReportEvent micDone, ReportingMessage , "Succeded at Row No:    " & intTestStepCount
            else
            Reporter.ReportEvent micDone, ReportingMessage , "Succeded at Row No:    " & intTestStepCount &",Value used:"&G
        End If
        
    End if

        End If                            

            Next 

End Function

