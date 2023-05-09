Const Version = "3.0.4"
'-----------------------------------------------------------------------------------------------
'The script copies the workflow definitions for the selected “Markets” of the selected “Environment” and its “Cluster” from the folder named “Workflows” in the current directory.
'---------------- 
'Requirements:
'1. Clusters are a subset of systems within an Amazon region which i`s sub folder of environment folder.
'2. All configuration properties need to be predefined inside the cluster folder for each region.
'3. The name of the config file should be like "WF ID - SYS Name" and if it is common for all systems, then just make it like "WF ID - ALL".
'4. For the config file for the standard imports, start the file name with "import"
'For workflows in workflows folder, do not update schedule
'Script contents
'Imports - Deviations
'Revisions
'1.0.0 APN 10/2/2015
'1.0.1 Dirghayu 10/12/2015 Update all other import wfs with std WF code.  Remove weird characters that magically appear in the start of some WF defn.  Update VBscript title as it is important to do in order to track version #.
'1.0.2 APN 12/8/2015 Remove special treatment of MOVE ORDERS file
'1.0.3 APN 2/23/2016 Dev#1570 Make all Imports non-system
'1.0.3 DN 04/12/2016 Dev#1789 We will not update the WF property PasswordEncryption, if exists.
'1.0.4 APN 04/22/2016 Dev#1460 Make all Imports non-system.  Not sure what happenned in 1570 above.
'1.0.4 DN  04/25/2016 Dev#1801 Standard Edge Interface Release Process - Do not overwrite Move file workflow file name settings
'1.0.5 APN 05/11/2016 Dev#???? Check for extra LFs or blanks in "Folder To Object"
'1.0.6 APN 05/16/2016 Dev#1873 Single version of WF property update not getting updated while bundled was.
'1.0.7 APN 05/17/2016 If no Import XML exists, do not include Import folder.  Also do not show message at end about updating Import workflows.
'1.0.7 DN 05/19/2016 Updated the script to read the special charactres and remove the unwanted characters(ï»¿)
'1.0.7 DN 06/27/2016 Do not update the WF code if existing WF code version # is higher than the release’s version
'1.0.8 DN 02/16/2017 Avoid updating mandatory flag update
'1.0.8 DN 03/06/2017 Update the workflow property "Configuration Property" as we are updating "Export Configuration"
'1.0.8 DN 03/10/2017 Update the workflow property "Workflows Monitored" as we are updating "Export Configuration"
'1.0.8 DN 03/16/2017 Added code to load touch event
'1.0.8 KR 17/03/2017 Added Delete Code for WorkFlow Property.(Defect #3302 ,#3308)
'1.0.8 KR 21/03/2017 Added code to Update the Workflow Property Value.(defect #3332)
'1.0.8 DN 04/06/2017 Fix for generating duplicate bundle WF.
'1.0.8 DN 04/20/2017 Code to escape '-'
'1.0.8 DN 29/06/2017 Do not update Costum Imports with std code
'1.0.9 ddave 14/07/2017 Added code to update wfProperty("Workflows Monitored") in email alert
'2.0.0 KR - Update the Code for standard Release.
'2.0.1 KR - Added code to log an error message if there is a workflow different name but same process id.
'2.0.2 APN 12/13/2018 - added WFs to list of exceptions to updateImportWFs
'2.0.3 'SGHUNCHALA & KRAVAL - 20190207 - Update code for standard order export wfs and to get wfs by it's processID , Change the position of functions files execution , Update the code to use current version of processrelease in system event log
'3.0.0 'SGHUNCHALA 20200529 - Legacy Update
'3.0.1 PPatel: Update script to import view and veiwGroup 
'3.0.2 PPatel:Update script to import ReportTemplates , PushReports , SummaryTemplates and Update script for add or update all classes members
'3.0.3 PPatel : Update script to import DynamicMemberDefinitions
'3.0.4 Maulik Makadiya [IN:124893] Improve Release process logging 09-May-2023
'-----------------------------------------------------------------------------------------------

'Constants
'subfolderName where Independent WF Processes are
Const IMPORT_INDEPENDENT_WF_PATH = "WF_ONEPERFILE"
'WF XML definition for the import process
Const IMPORT_WF_XML_DEF = "Import XML.txt"
'Default Property Name
Const DEFAULT_WF_PROPERTY_NAME = "Configuration Property"
'Title, ID And XML definition of the WF process which moves files from FTP folder to system import folder
Const WF_TITLE_FILES_FROM_FTP = "Import - MOVE FILES FROM FTP FOLDER"
Const WF_ID_FILES_FROM_FTP = "IMP_FILES_FROM_FTP"
Const WF_XMLFILE_FILES_FROM_FTP = "WF To Move Files From FTP Folder.xml"
'Title, ID And XML definition of the WF process which moves error files from system import folder to FTP error folder
Const WF_TITLE_ERROR_FILES_TO_FTP = "Import - MOVE ERROR FILES TO FTP FOLDER"
Const WF_ID_ERROR_FILES_TO_FTP = "IMP_ERR_FILES_TO_FTP"
Const WF_XMLFILE_ERROR_FILES_TO_FTP = "WF To Move Error Files to FTP Error Folder.xml"
'Title, ID And XML definition of the WF process which moves success files from system import folder to FTP success folder
Const WF_TITLE_SUCCESS_FILES_TO_FTP = "Import - MOVE SUCCESS FILES TO FTP FOLDER"
Const WF_ID_SUCCESS_FILES_TO_FTP = "IMP_SUC_FILES_TO_FTP"
Const WF_XMLFILE_SUCCESS_FILES_TO_FTP = "WF To Move Success Files to FTP Success Folder.xml"
'Definition of some WF properties
Const WF_IMPORT_ACTIVE = 0
Const WF_FTP_ACTIVE = 0
Const WF_SUSPENDED = 0
Const WF_ENABLELOGGING = 1
Const WF_NONSYSTEM = 0
Const WF_PRIORITY = 3
Const ENVMESSAGE =  "Error: Not a valid environment."
Const HKEY_LOCAL_MACHINE = &H80000002
Const strBaseKeyPath = "SOFTWARE\WOW6432Node\StayinFront\Active Elk\Systems"




Dim objOutErrFile,strdataline,blnUserDefinedKey,strKeyValue,strcommonfolder,blndeleterequired,strReleaseDetails,strReleaseID
blndeleterequired =False

Dim strDefaultWFXMLDef
ReleaseTitle = "Release Work-flow Processes (ver. " & Version & ")"

'We create an array with the subfolders under each import subfolder
lstImportSubfolders = Array("inbox","error","completed")

'DNimbark 20170629 - Do not update Costume Imports with std code
Dim strLstCustomeImports
strLstCustomeImports = ","

Dim strEnvironment

'AAA 2014/12/1. We use an array to indicate the settings to keep in WF properties
lstWFPropertiesToKeep = Array("AddPrime","FullDataImport","FullDataImportUpdatedField","FullDataImportFieldValue","FullDataImportRemoveField","FullDataImportFilterFields","FullDataImportFilterExpToKeepItems")

'Vable where we store the default settings from config_file
Dim General_Settings
Dim objSystemLog
Dim intTotalTouchEvents, intAddedTouchEvent, intErroredTouchEvent
intTotalTouchEvents  = 0
intAddedTouchEvent   = 0
intErroredTouchEvent = 0

Dim intTotalTouchDBParts, intAddedTouchDBParts, intErroredTouchDBParts
intTotalTouchDBParts  = 0
intAddedTouchDBParts   = 0
intErroredTouchDBParts = 0

Dim intTotalTouchDBTemplates, intAddedTouchDBTemplates, intErroredTouchDBTemplates
intTotalTouchDBTemplates  = 0
intAddedTouchDBTemplates   = 0
intErroredTouchDBTemplates = 0

Dim intTotalTouchDBTemplatePos, intAddedTouchDBTemplatePos, intErroredTouchDBTemplatePos
intTotalTouchDBTemplatePos  = 0
intAddedTouchDBTemplatePos   = 0
intErroredTouchDBTemplatePos = 0

Dim intTotalTouchDB, intAddedTouchDB, intErroredTouchDB
intTotalTouchDB  = 0
intAddedTouchDB   = 0
intErroredTouchDB = 0

Dim intTotalKPIS, intAddedKPIS, intErroredKPIS
intTotalKPIS  = 0
intAddedKPIS   = 0
intErroredKPIS = 0

Dim intTotalViews, intAddedViews, intErroredViews
intTotalViews  = 0
intAddedViews   = 0
intErroredViews = 0

Dim intTotalPushReports, intAddedPushReports, intErroredPushReports
intTotalPushReports  = 0
intAddedPushReports   = 0
intErroredPushReports = 0

Dim intTotalViewGroups, intAddedViewGroups, intErroredViewGroups
intTotalViewGroups  = 0
intAddedViewGroups   = 0
intErroredViewGroups = 0

Dim intTotalSummaryTemplates, intAddedSummaryTemplates, intErroredSummaryTemplates
intTotalSummaryTemplates  = 0
intAddedSummaryTemplates   = 0
intErroredSummaryTemplates = 0

Dim intTotalDynamicMemberDefinitions, intAddedDynamicMemberDefinitions, intErroredDynamicMemberDefinitions
intTotalDynamicMemberDefinitions  = 0
intAddedDynamicMemberDefinitions   = 0
intErroredDynamicMemberDefinitions = 0

Dim intTotalReportTemplates, intAddedReportTemplates, intErroredReportTemplates
intTotalReportTemplates  = 0
intAddedReportTemplates   = 0
intErroredReportTemplates = 0

Dim intTotalAnalytics, intAddedAnalytics, intErroredAnalytics
intTotalAnalytics  = 0
intAddedAnalytics   = 0
intErroredAnalytics = 0

'We get the current folder
sn = Wscript.ScriptName 
fn = Wscript.ScriptFullName 
fp = Replace(fn, sn, "")
strFolder = fp

'We create a IE object in order to show the progress
Set objIE = WScript.createobject("internetexplorer.application", "IE_")
Sub IE_onQuit()
  Wscript.Quit
End Sub
'in code, the colon acts As a line feed
objIE.navigate2 "about:blank" : objIE.width = 600 : objIE.height = 650 : objIE.toolbar = False : objIE.menubar = False : objIE.statusbar = False : objIE.visible = True
With objIE.Document.parentWindow.screen
	objIE.Left = (.availWidth  - objIE.Width ) \ 2
	objIE.Top  = (.availHeight - objIE.Height) \ 2
End With
'Create a function to show a wait message
'We put a title to the new popup window
objIE.document.title = "Processes Release" & ReleaseTitle
objIE.document.write "<!DOCTYPE html>"
objIE.document.write "<script>function ShowWaitMessage() {var element = document.getElementById('waitMessage');if (element!=null){if (element.innerHTML.length>15)element.innerHTML = 'wait .';else element.innerHTML += '.';}};</script>"
objIE.document.write "<style> label {width:170px;display: inline-block;} h3 {margin:0px;padding:0px}</style>"
'Write a header
objIE.document.write "<h3> StayinFront app for Processes Release</h3>"
WriteLineSeparator
'open fso 
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(strFolder)

'validate folders are exist.
Set objMainFolders = CreateObject("Scripting.Dictionary")
Set objDirectory = objFSO.GetFolder(objFolder)
'We go over each subfolder
For Each SubFld in objDirectory.SubFolders
	objMainFolders.Add SubFld.name, SubFld.path
Next

Dim blnExitScript
blnExitScript = False

If objMainFolders.Exists("Config") Then
	strValidateReleaseConfig = ValidateReleaseConfig(objFSO) 
	If strValidateReleaseConfig <> "" Then 
		WriteErrorLine(strValidateReleaseConfig)
		blnExitScript = True
	End If
End If	

'select environment from list and disabled it once it selected
strEnvironmentPath = ""
If Not blnExitScript  AND objMainFolders.Exists("Config") Then
	objIE.document.write "<fieldset ><legend><p>1. Select an Environment and Cluster: </p></legend>"
	strEnvironmentPath = GetSelectedSubFolderPath(objMainFolders.Item("Config") , "Environment")
	strEnvironment = split(strEnvironmentPath,"Config\")
	If objFSO.FolderExists(strEnvironmentPath) Then
		strEnvironmentPath = AddEndSlashIfNecessary(strEnvironmentPath)
		strClusterPath = GetSelectedSubFolderPath(strEnvironmentPath , "Cluster")
	End if
	objIE.document.write "</fieldset>"
End If

strValidateEnvironment = ValidateEnvironment(strEnvironmentPath)
if Not blnExitScript And strValidateEnvironment <> "" Then 
	WriteErrorLine(strValidateEnvironment)
	blnExitScript = True
End If	

'select Cluster from list and disabled it once it selected
If Not blnExitScript And objFSO.FolderExists(strClusterPath) Then
	strSubfolder = AddEndSlashIfNecessary(strClusterPath)
	Set objDirectory = Nothing
	Set objDirectory = objFSO.GetFolder(strSubfolder)
'praval 20150612 #1574 - Deactivate SKUs Workflow - deliver globally
	Set objExpTurn = CreateObject("Scripting.Dictionary")
	Set objMoveOrdertoFtp = CreateObject("Scripting.Dictionary")
	Set objOrderExport = CreateObject("Scripting.Dictionary")
	Set objImport = CreateObject("Scripting.Dictionary")
	Set objAllWF = CreateObject("Scripting.Dictionary")
	Set objMarkets = CreateObject("Scripting.Dictionary")
	
	For each objFile in objDirectory.Files
		parts = split(objFile.Name, ".")
		If UCase(parts(0))="MARKETS" Then
			Const ForReading = 1
			Dim arrMarkets()
			Dim iUpperBound
			Set ts = objFSO.OpenTextFile(objFile.path, 1)

			iUpperBound = 0
			While Not ts.AtEndOfStream
				ReDim Preserve arrMarkets(iUpperBound)
				arrMarkets(UBound(arrMarkets)) = ts.ReadLine
				iUpperBound = iUpperBound + 1
			Wend
			ts.Close
		Else
			StrWFP = ""
			StrWFP = split(parts(0),Chr(45))
			if (ubound(StrWFP) <> 1) then
				StrWFP = split(parts(0),Chr(150))
				if (ubound(StrWFP) <> 1) then
					StrWFP = split(parts(0),Chr(151))
				End if
			end if
	'praval 20150612 #1574 - Deactivate SKUs Workflow - deliver globally
			If (ubound(StrWFP) >= 1) then
				Select Case trim(UCase(StrWFP(0)))
				  Case "EXP_TURN_IN":
					If Not objExpTurn.Exists(trim(UCase(StrWFP(1)))) Then
						objExpTurn.Add trim(UCase(StrWFP(1))), objFile.path
					Else
						objExpTurn.Item(trim(UCase(StrWFP(1)))) = objFile.path
					End If
				  Case "ORDER_EXPORT_STD":
					If Not objOrderExport.Exists(trim(UCase(StrWFP(1)))) Then
						objOrderExport.Add trim(UCase(StrWFP(1))), objFile.path
					Else
						objOrderExport.Item(trim(UCase(StrWFP(1)))) = objFile.path
					End If
				  Case "IMPORT":
					If Not objImport.Exists(trim(UCase(StrWFP(1)))) Then
						objImport.Add (trim(UCase(StrWFP(1)))), objFile.path
					Else
						objImport.Item(trim(UCase(StrWFP(1)))) = objFile.path
					End If
					
				  Case else
					if trim(UCase(StrWFP(1))) = "ALL" Then
						If Not objAllWF.Exists(trim(UCase(StrWFP(0)))) Then
							objAllWF.Add (trim(UCase(StrWFP(0)))), objFile.path
						Else
							objAllWF.Item(trim(UCase(StrWFP(0)))) = objFile.path
						End if
					Else
						tmpdicwfid = ""
						tmpdicwfid = trim(StrWFP(0))&"-"&trim(StrWFP(1))
						tmpdicwfid = trim(UCase(tmpdicwfid))
						If Not (objAllWF.Exists(tmpdicwfid)) Then
							objAllWF.Add (tmpdicwfid), objFile.path
						Else
							objAllWF.Item(tmpdicwfid) = objFile.path
						End if
					End if
					
				  End Select
			Else
				objIE.document.write "File name is incorrect. File name must contain hyphen which separates the workflow name and process code. :'" & parts(0) & "'."
				objIE.document.write "File :'" & parts(0) & "' Not processed."
			End If
		End if
	Next
End If

Dim arrMarketsSelected()

if not blnExitScript then	
	ReDim Preserve arrMarketsSelected(UBound(arrMarkets))
	If	UBound(arrMarkets) >= 0 Then
		'We get the Market
		objIE.document.write "<fieldset ><legend><p>2. Select a Market : </p></legend>"
		arrMarketsSelected(0) = GetselectedMarket(arrMarkets,objIE)
		objIE.document.write "</fieldset>"
		if (arrMarketsSelected(0)(0) = "Exit") Then
			blnExitScript = True
		End if
	Else
		WriteCommentLine("Market file dose not have any market.")
		blnExitScript = True
	End if
End If

if not blnExitScript then	
	If UBound(arrMarketsSelected(0)) >= 0 Then
		'We get the username and password
		objIE.document.write "<fieldset ><legend><p>3. UserName and Password : </p></legend>"
		objIE.document.write "<b>USERNAME :</b> <INPUT id=txtUser name=UserName style='WIDTH: 150px; HEIGHT: 15px' size=15></INPUT></P><br>"
		objIE.document.write "<b>PASSWORD :</b> <INPUT id=txtPass name=Password Type=Password style='WIDTH: 150px; HEIGHT: 15px' size=15></INPUT></P><br>"
		objIE.document.getElementById("txtUser").disabled = true
		objIE.document.getElementById("txtPass").disabled = true
		
		blnOkToProcess = ""
		blnOkToProcess = GetUserNamePass(objIE)
		sOKtoProcess = ""
		sUser = ""
		sPass = ""
		While sOKtoProcess = ""
			If (blnOkToProcess = 6) then
					objIE.document.write "<INPUT type='button' id=ClickOk value='OK'/>"
					sOKtoProcess = objIE.document.getElementById("ClickOk").value
					objIE.document.getElementById("ClickOk").disabled = true
			Elseif(blnOkToProcess = 7) then
				blnOkToProcess = GetUserNamePass(objIE)
			Elseif(blnOkToProcess = 2) Then
				WriteErrorLine("ERROR: You press cancel to exit script running ")
				objIE.document.getElementById("txtUser").Value = ""
				objIE.document.getElementById("txtPass").Value = ""
				sOKtoProcess = "cancel"
				blnExitScript = True
			End if
		Wend
		objIE.document.getElementById("txtUser").disabled = true
		objIE.document.getElementById("txtPass").disabled = true
		
		sUser = objIE.document.getElementById("txtUser").value
		sPass = objIE.document.getElementById("txtPass").value
		objIE.document.write "</fieldset>"
	Else
		WriteCommentLine("Market not selected.")
		blnExitScript = True
	End if
End if



IF not blnExitScript Then
	'We show the wait message
	ShowWaitMessage
	WriteLineSeparator
	
	For Each strQtyMember In arrMarketsSelected(0)
		'sServer = "https://retailservicesedgeqa.stayinfrontcreative.com;timeout=180;proxy=auto;cookies=1"
		sServer = ""
		sSystem = strQtyMember
		'We connect to CRM system
		Set System = CreateObject("ActivElk.System")
		
		If sUser = "" Then
			If system.connectdlg then
				msgbox "Cannot connect to visual elk."
				wscript.quit(1)
			End If
		Else
			System.Connect sServer, sSystem, sUser, sPass
			HideWaitMessage
			If Err <> 0 Then
				objIE.document.write "<br><span style='color:red'>Connection Error :</span> ' " & Err.Description & "'. System : " & sSystem & ". Not processed. </br>"
			Else
				objIE.document.write "<br><span style='color:green'>Connected To : </span>' " & sSystem & "'.</br>"
			'praval 20150626 Update the Export Configuration 
				
				'SGHUNCHALA & KRAVAL - 20190207 - Update code to use current version of process release in system event log 
				set objSystemLog  = System.LG_LogJobStarted("Release " & strReleaseID)
				objSystemLog.LogNormalDetail "The process started at " & Now() & "."
				Dim strFilter
				Dim arrWfproperty			'To store workflow property 
				arrWfproperty = Array("Gateway Location","Gateway Archive Location","Order Export File Name","Order Export File Path","Export Configuration")
				strFilter = "workflowproperties.exists(wfProperty='"&arrWfproperty(0)&"') and "
				strFilter = strFilter  & "workflowproperties.exists(wfProperty='"&arrWfproperty(1)&"') and "
				strFilter = strFilter  & "workflowproperties.exists(wfProperty='"&arrWfproperty(2)&"') and "
				strFilter = strFilter  & "workflowproperties.exists(wfProperty='"&arrWfproperty(3)&"') and "
				strFilter = strFilter  & "workflowproperties.exists(wfProperty='"&arrWfproperty(4)&"') and "
				'SGHUNCHALA & KRAVAL - 20190207 - Update filter to exclude MDLZ russia's customize order export wfs
				strFilter = strFilter  & "((WF_ProcID <> 'RUS_ORD_SAP') and (WF_ProcID <> 'ORDER_EXPORT_STD_RU') or IsNull(WF_ProcID))"
				For each objInstance in system.folders.wf_processes.scan(strFilter)
					ChangeProperty objInstance, strQtyMember
				Next

				If objMainFolders.Exists("imports") Then
					objSystemLog.LogNormalDetail "The ""imports"" folder is executing."
					strConfigFile = objImport.Item(trim(UCase(strQtyMember)))
					strImportFolder = objMainFolders.Item("imports")
					strImportFolder = AddEndSlashIfNecessary(strImportFolder)
					Call ImportProcessWorkflow(System,strImportFolder, strConfigFile)
					objSystemLog.LogNormalDetail "The ""imports"" folder is executed successfully."
				End If
				
				
				strEnvironmentPath = ""
				If objMainFolders.Exists("workflows") Then
					objSystemLog.LogNormalDetail "The ""workflows"" folder is executing."
					strEnvironmentPath = GetSubFolderPath(objMainFolders.Item("workflows"),"global")
					If objFSO.FolderExists(strEnvironmentPath) Then
						strEnvironmentPath = AddEndSlashIfNecessary(strEnvironmentPath)
						strWorkFlowFolder = GetSubFolderPath(strEnvironmentPath , "workflows")
									
						Set objWFDirectory = Nothing
						Set objWFDirectory = objFSO.GetFolder(strWorkFlowFolder)
						'We go over each subfolder
						For Each WFSubFld in objWFDirectory.SubFolders
							'SGHUNCHALA & KRAVAL - 20190207 - Update code for standard order export workflows
							If UCase(WFSubFld.name) = "STANDARD EXPORT" Then
								For Each objInstance In system.folders.wf_processes.scan(strFilter)
									ADDWFFolder System, objOrderExport, objAllWF, WFSubFld, strQtyMember , objInstance
									objIE.document.write "<br>Adding or Updating for the Workflow/' " & objInstance.Title & "'  -- <span style='color:green'>Done.</span></br>"
								Next
							Else
								set objInstance = nothing
								ADDWFFolder System, objOrderExport, objAllWF, WFSubFld, strQtyMember , objInstance
								objIE.document.write "<br>Adding or Updating for the Workflow/' " & WFSubFld.name & "'  -- <span style='color:green'>Done.</span></br>"
							End IF							
						Next
						Set objWFDirectory = Nothing
					End if
					objSystemLog.LogNormalDetail "The ""workflows"" folder is executed successfully."
				End if
				
				'SGhunchala 20200529 - Add Dashboard part functionality
				If objMainFolders.Exists("dashboardparts") Then
					objSystemLog.LogNormalDetail "The ""dashboardparts"" folder is executing."
					WriteLineSeparator
					strLogFile = loadTouchDBParts(System,objMainFolders.Item("dashboardparts"))
					objIE.document.write "Adding or Updating Touch Dashboard Parts</br>"
					objIE.document.write "Total Touch DBParts found : " & intTotalTouchDBParts & "</br>"
					objIE.document.write "Total Updated/Added Touch DBParts : " & intAddedTouchDBParts & "</br>"
					objIE.document.write "Total Touch DBParts errored to Add\Update : " & intErroredTouchDBParts & "</br>"
					objSystemLog.LogNormalDetail "The ""dashboardparts"" folder is executed successfully."
				End If
				
				'SGhunchala 20200529 - Add Dashboard template functionality
				If objMainFolders.Exists("dashboardtemplates") Then
					objSystemLog.LogNormalDetail "The ""dashboardtemplates"" folder is executing."
					WriteLineSeparator
					strLogFile = loadTouchDBTemplates(System,objMainFolders.Item("dashboardtemplates"))
					objIE.document.write "Adding or Updating Touch Dashboard templates</br>"
					objIE.document.write "Total Touch DB templates found : " & intTotalTouchDBTemplates & "</br>"
					objIE.document.write "Total Updated/Added Touch DB templates : " & intAddedTouchDBTemplates & "</br>"
					objIE.document.write "Total Touch DB templates errored to Add\Update : " & intErroredTouchDBTemplates & "</br>"
					objSystemLog.LogNormalDetail "The ""dashboardtemplates"" folder is executed successfully."
				End If
				
				'SGhunchala 20200529 - Add Dashboard template Position functionality
				If objMainFolders.Exists("dashboardtemplatespos") Then
					objSystemLog.LogNormalDetail "The ""dashboardtemplatespos"" folder is executing."
					WriteLineSeparator
					strLogFile = loadTouchDBTemplatePos(System,objMainFolders.Item("dashboardtemplatespos"))
					objIE.document.write "Adding or Updating Touch Dashboard template Position</br>"
					objIE.document.write "Total Touch DB template Position found : " & intTotalTouchDBTemplatePos & "</br>"
					objIE.document.write "Total Updated/Added Touch DB template Position : " & intAddedTouchDBTemplatePos & "</br>"
					objIE.document.write "Total Touch DB template Position errored to Add\Update : " & intErroredTouchDBTemplatePos & "</br>"
					objSystemLog.LogNormalDetail "The ""dashboardtemplatespos"" folder is executed successfully."
				End If
				
				'SGhunchala 20200529 - Add touch Dashboards
				If objMainFolders.Exists("dashboards") Then
					objSystemLog.LogNormalDetail "The ""dashboards"" folder is executing."
					WriteLineSeparator
					strLogFile = loadTouchDB(System,objMainFolders.Item("dashboards"))
					objIE.document.write "Adding or Updating Touch Dashboards</br>"
					objIE.document.write "Total Touch Dashboard found : " & intTotalTouchDB & "</br>"
					objIE.document.write "Total Updated/Added Touch Dashboard : " & intAddedTouchDB & "</br>"
					objIE.document.write "Total Touch Dashboard errored to Add\Update : " & intErroredTouchDB & "</br>"
					objSystemLog.LogNormalDetail "The ""dashboards"" folder is executed successfully."
				End If
				
				'SGhunchala 20200529 - Add KPIs
				If objMainFolders.Exists("KPIs") Then
					objSystemLog.LogNormalDetail "The ""KPIs"" folder is executing."
					WriteLineSeparator
					strLogFile = loadKPIS(System,objMainFolders.Item("KPIs"))
					objIE.document.write "Adding or Updating KPIs</br>"
					objIE.document.write "Total KPIs found : " & intTotalKPIS & "</br>"
					objIE.document.write "Total Updated/Added KPIs : " & intAddedKPIS & "</br>"
					objIE.document.write "Total KPIs errored to Add\Update : " & intErroredKPIS & "</br>"
					objSystemLog.LogNormalDetail "The ""KPIs"" folder is executed successfully."
				End If
				
				'PPatel 20200528 - Add Views 
				If objMainFolders.Exists("Views") Then
					objSystemLog.LogNormalDetail "The ""Views"" folder is executing."
					WriteLineSeparator
					strLogFile = loadViews(System,objMainFolders.Item("Views"))
					objIE.document.write "Adding or Updating Views</br>"
					objIE.document.write "Total Views found : " & intTotalViews & "</br>"
					objIE.document.write "Total Updated/Added Views : " & intAddedViews & "</br>"
					objIE.document.write "Total Views errored to Add\Update : " & intErroredViews & "</br>"
					objSystemLog.LogNormalDetail "The ""Views"" folder is executed successfully."
				End If
				
				'PPatel 20200528 - Add ViewsGroup
				If objMainFolders.Exists("ViewGroups") Then
					objSystemLog.LogNormalDetail "The ""ViewGroups"" folder is executing."
					WriteLineSeparator
					strLogFile = loadViewGroups(System,objMainFolders.Item("ViewGroups"))
					objIE.document.write "Adding or Updating ViewGroups </br>"
					objIE.document.write "Total ViewGroups  found : " & intTotalViewGroups & "</br>"
					objIE.document.write "Total Updated/Added ViewGroups : " & intAddedViewGroups & "</br>"
					objIE.document.write "Total ViewGroups errored to Add\Update : " & intErroredViewGroups & "</br>"
					objSystemLog.LogNormalDetail "The ""ViewGroups"" folder is executed successfully."
				End If
				
				'PPatel 20200528 - Add PushReports
				If objMainFolders.Exists("PushReports") Then
					objSystemLog.LogNormalDetail "The ""PushReports"" folder is executing."
					WriteLineSeparator
					strLogFile = loadPushReports(System,objMainFolders.Item("PushReports"))
					objIE.document.write "Adding or Updating PushReports </br>"
					objIE.document.write "Total PushReports  found : " & intTotalPushReports & "</br>"
					objIE.document.write "Total Updated/Added PushReports : " & intAddedPushReports & "</br>"
					objIE.document.write "Total PushReports errored to Add\Update : " & intErroredPushReports & "</br>"
					objSystemLog.LogNormalDetail "The ""PushReports"" folder is executed successfully."
				End If
				
				'PPatel 20200528 - Add ReportTemplates
				If objMainFolders.Exists("ReportTemplates") Then
					objSystemLog.LogNormalDetail "The ""ReportTemplates"" folder is executing."
					WriteLineSeparator
					strLogFile = loadReportTemplates(System,objMainFolders.Item("ReportTemplates"))
					objIE.document.write "Adding or Updating ReportTemplates </br>"
					objIE.document.write "Total ReportTemplates  found : " & intTotalReportTemplates & "</br>"
					objIE.document.write "Total Updated/Added ReportTemplates : " & intAddedReportTemplates & "</br>"
					objIE.document.write "Total ReportTemplates errored to Add\Update : " & intErroredReportTemplates & "</br>"
					objSystemLog.LogNormalDetail "The ""ReportTemplates"" folder is executed successfully."
				End If
				
				'PPatel 20200528 - Add SummaryTemplates
				If objMainFolders.Exists("SummaryTemplates") Then
					objSystemLog.LogNormalDetail "The ""SummaryTemplates"" folder is executing."
					WriteLineSeparator
					strLogFile = loadSummaryTemplates(System,objMainFolders.Item("SummaryTemplates"))
					objIE.document.write "Adding or Updating SummaryTemplates </br>"
					objIE.document.write "Total SummaryTemplates  found : " & intTotalSummaryTemplates & "</br>"
					objIE.document.write "Total Updated/Added SummaryTemplates : " & intAddedSummaryTemplates & "</br>"
					objIE.document.write "Total SummaryTemplates errored to Add\Update : " & intErroredSummaryTemplates & "</br>"
					objSystemLog.LogNormalDetail "The ""SummaryTemplates"" folder is executed successfully."
				End If
				
				'PPatel 20200528 - Add DynamicMemberDefinitions
				If objMainFolders.Exists("DynamicMemberDefinitions") Then
					objSystemLog.LogNormalDetail "The ""DynamicMemberDefinitions"" folder is executing."
					WriteLineSeparator
					strLogFile = loadDynamicMemberDefinitions(System,objMainFolders.Item("DynamicMemberDefinitions"))
					objIE.document.write "Adding or Updating DynamicMemberDefinitions </br>"
					objIE.document.write "Total DynamicMemberDefinitions  found : " & intTotalDynamicMemberDefinitions & "</br>"
					objIE.document.write "Total Updated/Added DynamicMemberDefinitions : " & intAddedDynamicMemberDefinitions & "</br>"
					objIE.document.write "Total DynamicMemberDefinitions errored to Add\Update : " & intErroredDynamicMemberDefinitions & "</br>"
					objSystemLog.LogNormalDetail "The ""DynamicMemberDefinitions"" folder is executed successfully."
				End If
				
				'SGhunchala 20200529 - Add analytics
				If objMainFolders.Exists("analytics") Then
					objSystemLog.LogNormalDetail "The ""Analytics"" folder is executing."
					WriteLineSeparator
					strLogFile = LoadAnalyticViews(System,objMainFolders.Item("analytics"))
					objIE.document.write "Adding or Updating Analytics</br>"
					objIE.document.write "Total Analytics found : " & intTotalAnalytics & "</br>"
					objIE.document.write "Total Updated/Added Analytics : " & intAddedAnalytics & "</br>"
					objIE.document.write "Total Analytics errored to Add\Update : " & intErroredAnalytics & "</br>"
					objSystemLog.LogNormalDetail "The ""Analytics"" folder is executed successfully."
				End If
				
				'ppatel 20200529 - Add importprocess
				If objMainFolders.Exists("importCSVs") Then
					WriteLineSeparator
					objIE.document.write "The importCSVs folder is executing.</br>"
					objSystemLog.LogNormalDetail "The ""importCSVs"" folder is executing."
					On Error Resume Next
					strImportFolder = objMainFolders.Item("importCSVs")
					strImportFolder = AddEndSlashIfNecessary(strImportFolder)
					strcommonfolder =strImportFolder & "\" & "ALL\"
					If objFSO.FolderExists(strcommonfolder) Then
						call Convertexceltocsv(strcommonfolder & "Inbox\" )
						call Loadimports(System,strcommonfolder,objSystemLog)
						
						blndeleterequired = True
					end if
					strclientFolder = strImportFolder & "\" & sSystem & "\"
					If objFSO.FolderExists(strclientFolder) Then
						call Convertexceltocsv(strclientFolder & "Inbox\" )
						call Loadimports(System,strclientFolder,objSystemLog)
						Set folder = objFSO.GetFolder(strclientFolder& "Inbox\")

						' delete all files in root folder
						for each f in folder.Files
							if LCase(objFSO.GetExtensionName(f.Name)) = "csv" then
							   On Error Resume Next
							   name = f.name
							   f.Delete True
							   If Err Then
								 WScript.Echo "Error deleting:" & Name & " - " & Err.Description
									objOutErrFile.WriteLine("there are some error while Error deleting:" & Name & " - " & Err.Description)
								
									objSystemLog.LogNormalDetail "See error log for error details "
									objSystemLog.LogErrorDetail "","", "there are some error while Error deleting:" & Name & " - " & Err.Description
							   End If
							   On Error GoTo 0
							end if
						Next
						
					end if
					If Err <> 0 Then
						If Trim(Err.Description) = "" Then
							objOutErrFile.WriteLine("there are some error while process import file  " )
							
							objSystemLog.LogErrorDetail "there are some error while process import file  "
						Else
							objOutErrFile.WriteLine("there are some error while process import file  " & Err.Description)
							
							objSystemLog.LogNormalDetail "See error log for error details "
							objSystemLog.LogErrorDetail "","", "there are some error while process import file  " & Err.Description
						End If
					end if
					On Error goto 0
					objSystemLog.LogNormalDetail "The ""importCSVs"" folder is executed successfully."
					objIE.document.write "The importCSVs folder is executed successfully.</br>"
				End If
				
				'SGhunchala 20200529 - Update Workflow level members - (workflows - SystemName)
				Call UpdateWorkflowMembers(System)
				
				'SGHUNCHALA & KRAVAL - 20190207 - Change the position of function file exception
				
				'======= KRaval and SGhunchala - Functions will be call from here =======

					'Call functions
					
				'==================================================
				
				'DNimbark 20160608 - function to load the touch event
				If objMainFolders.Exists("touchevents") Then
					objSystemLog.LogNormalDetail "The ""touchevents"" folder is executing."
					WriteLineSeparator
					strLogFile = loadTouchEvents(System,objMainFolders.Item("touchevents"))
						objIE.document.write "Adding or Updating Touch Events</br>"
						objIE.document.write "Total Touch Events found : " & intTotalTouchEvents & "</br>"
						objIE.document.write "Total Updated/Added Touch Events : " & intAddedTouchEvent & "</br>"
						objIE.document.write "Total Touch Events errored to Add\Update : " & intErroredTouchEvent & "</br>"
					'objIE.document.write "Touch Event Add\Update -- <span style='color:green'>Done.</span> Please check the generated log file '" & strLogFile & "' for more detail.</br>"
					objSystemLog.LogNormalDetail "The ""touchevents"" folder is executed successfully."
				End If
				
				'KRaval - Added code to Enable touch event.
				Call EnableTouchEvent(System)
			objSystemLog.LogJobComplete "The process completed at " & Now() & "."
			
				'KWatwani 20170927 - Defect #4500 - The processrelease.vbs is not updating all workflows on the first pass
				'DNimbark 12/10/2015 To update the Workflows
				'DNimbark 20160511 - Removed the condition for checking count of the dictonary
				Dim intCount
				intCount = 0
				intCount = updateImportWFs()
				If intCount>0 Then
					objIE.document.write "<br>Updating the remaining import workflows at the system</br>"
					objIE.document.write "<br>" & intCount & " workflows successfully updated</br>Completed!"
				End If
			End if
		End If
		WriteLineSeparator	
	Next
	
	if blndeleterequired = True then
		' delete all files in root folder
		Set folder = objFSO.GetFolder(strcommonfolder& "Inbox\")
		for each f in folder.Files
			if LCase(objFSO.GetExtensionName(f.Name)) = "csv" then
			   On Error Resume Next
			   name = f.name
			   f.Delete True
			   If Err Then
				 WScript.Echo "Error deleting:" & Name & " - " & Err.Description
					objOutErrFile.WriteLine("there are some error while Error deleting:" & Name & " - " & Err.Description)
				
					objSystemLog.LogNormalDetail "See error log for error details "
					objSystemLog.LogErrorDetail "","", "there are some error while Error deleting:" & Name & " - " & Err.Description
			   End If
			   On Error GoTo 0
			end if
		Next
	end if
Else
	WriteCommentLine("Close the Internet explorer and again run the script to import this workflow processes.")
End if

Set system = Nothing

objIE.document.write "</html>"
'We remove the wait message
HideWaitMessage
'Completed message box
msgbox "Completed.",0,ReleaseTitle


'2015-03-30, Jhirpara Function to get market from list of array.
'parameters:
'	arrMarkets: List of market
'	objIE:		internet explorer object to print status
Function GetselectedMarket(arrMarkets,objIE)
	Dim arrUpdatedMarkets()
	
	blnExitselectedMarket = False
	strListMarket = ""
	strListMarket = " Please Add Market Number separated by : (,)" & vbNewLine
	strListMarket = strListMarket & vbNewLine
	
	iMarket = 0
	strListMarket = strListMarket & iMarket & " - ALL" & vbNewLine
	blnValidSystems = True
	
	For Each strmarket in arrMarkets
		iMarket = iMarket + 1
		If CheckRegistrySystemKeyExists(strmarket) Then
			strListMarket = strListMarket & iMarket & " - " & strmarket & vbNewLine
		Else
			strtempMsg = strmarket & " System is not found in this server."
			WriteCommentLine ("Error: " & strtempMsg)
			'MsgBox strtempMsg, vbCritical, "Error Message"
			blnValidSystems = False
		End If	
	Next
	
	If blnValidSystems Then
		strNoListMarket = ""
		Do While strNoListMarket = ""
			strNoListMarket = InputBox(strListMarket, ReleaseTitle, "")
			
			If TypeName(strNoListMarket) = "Empty" Then
				MsgBox "User has cancelled"
				blnExitselectedMarket = True
				Exit Do
			End If
		Loop
	Else
		blnExitselectedMarket = True
	End If	
	
	strNoListMarket  = Trim(strNoListMarket)
	if not (strNoListMarket = "") then
		if ((strNoListMarket = 0) and (not InStr(strNoListMarket,",0,")) and (not instr(strNoListMarket,",0")))  Then
			iSelect = 0
			For Each strSelMar in arrMarkets
				ReDim Preserve arrUpdatedMarkets(iSelect)
				arrUpdatedMarkets(UBound(arrUpdatedMarkets)) = strSelMar
				iSelect = iSelect + 1
			Next
		Else
			'We get all selected market array ID
			lstArrayIDm = Split(strNoListMarket,",")
			Dim iListMarket
			iListMarket = 0
			
			For Each strSelMar in lstArrayIDm
				strmarketIndex = ""
					strmarketIndex = arrMarkets(strSelMar - 1)
				
				If Err <> 0 Then
					objIE.document.write "<br><span style='color:red'>Invalid market selected :</span> ' " &  strSelMar & "'. Error : " & Err.Description & ". Not processed. </br>"
				Else
					if strmarketIndex <> "" Then
						ReDim Preserve arrUpdatedMarkets(iListMarket)
						arrUpdatedMarkets(UBound(arrUpdatedMarkets)) = strmarketIndex
						iListMarket = iListMarket + 1
					Else
						strErrorMsg = ""
						strErrorMsg = "Sorry! Market No. : " & strSelMar & ". Is out of range."
						MsgBox strErrorMsg,0,ReleaseTitle
					End if
				End if
			Next 
		End if
	End If
	
	If blnExitselectedMarket then
		Dim arrExitselectedMarket(0)
		arrExitselectedMarket(0) = "Exit"
		GetselectedMarket = arrExitselectedMarket
	Else
		if UBound(arrUpdatedMarkets) >= 0 then
		
			strHtmlCHKOptions = ""
			strHtmlCHKOptions = "<br>"
			For Each strmarket in arrMarkets
				strHtmlCHKOptions = strHtmlCHKOptions & "<input type='checkbox' name='market' id='" & strmarket & "' value='" & strmarket & "'>" & strmarket & "<br>"
			Next
			strHtmlCHK = ""
			strHtmlCHK = "<label>Select Market : </label>" & strHtmlCHKOptions & "" & _
				"<br><input type='hidden' id='cheked' value='' /><button type='button' id='PressOk' onclick='assignValue()'>Press OK:</button><br>"
			strScript = ""
			strScript = "<script> function assignValue() { document.getElementById('cheked').value = 'OK';}</script>"
			objIE.document.write strScript
			objIE.document.write strHtmlCHK
			
			For Each strCHKmarket in arrUpdatedMarkets
				objIE.document.getElementById(strCHKmarket).checked = True 
			Next
			objIE.document.getElementById("PressOk").disabled = true
		End if
		GetselectedMarket = arrUpdatedMarkets
	End if	
	
	
End Function


'2015-03-30, Jhirpara Function to get username and password from user.
'parameters:
'	objIE:	internet explorer object to print status
Function GetUserNamePass(objIE)
	UserName = InputBox("Enter UserName :", ReleaseTitle, "")
	PassWord = InputBox("Enter PassWord :", ReleaseTitle, "")
	
	objIE.document.getElementById("txtUser").Value = UserName
	objIE.document.getElementById("txtPass").Value = PassWord
	
	strMessagetitle = ""
	strMessagetitle = "Verify the UserName and Password" & vbNewLine
	strMessagetitle = strMessagetitle & "UserName : " & UserName & vbNewLine
	strMessagetitle = strMessagetitle & "PassWord : "& PassWord & vbNewLine
	
	GetUserNamePass = MsgBox(strMessagetitle,3,ReleaseTitle)
End Function
'2018-03-05 KRaval
Function EnableTouchEvent(System)
	set fso = CreateObject("Scripting.FileSystemObject")
	strEventFile = "touchevents - " & System.Name & ".txt"
	strDirectory = strClusterPath & "\" & strEventFile
	If fso.FileExists(strDirectory) Then
		objSystemLog.LogNormalDetail "Started execution to enable the touch event."
		Set objCodeFile = fso.OpenTextFile(strDirectory)
		If not objCodeFile.AtEndOfStream Then
			Do Until objCodeFile.AtEndOfStream
				strCode = objCodeFile.ReadLine
				If Len(strCode) > 0 Then
					strSplitLine = Split(strCode, "|")
					strRoleCode = strSplitLine(0)
					strEventNames = strSplitLine(1)
					blnEnable = CBool(strSplitLine(2))
					If System.Folders.Roles.count("Code = '"& strRoleCode &"'")>0 then
						For each objRole in System.Folders.Roles.scan("Code = '"& strRoleCode &"'")
							'strEvent = Split(strEventNames,",")
							For Each strEvent In Split(strEventNames,",")
								If Len(strEvent) > 0 then						
									set objTouchEvent = System.Folders.TCG_TouchEvent.First("Description.Primary='"& strEvent &"'")
									set objRoleTouchEvent = objRole.Folders.TCG_TouchEvents.First("TCG_TouchEvent.Description.Primary='"& strEvent &"'")
									If Not objTouchEvent.IsNull Then									
										On Error Resume Next
											Set objTransaction = System.BeginTransaction
											If blnEnable Then
												If objRoleTouchEvent.IsNull Then
													set objRoleTouchEvent = objRole.Folders.TCG_TouchEvents.CreateNewInstance()
													objRoleTouchEvent.TCG_TouchEvent = objTouchEvent
													If Err.Number <> 0 Then
														objSystemLog.LogErrorDetail "","","Error : " & Err.Description
													End If
												End If	
												objRoleTouchEvent.selected = 1
												objRoleTouchEvent.Save objTransaction
												set objval = objTransaction.validate		
												if objval.status <> 3 then
													objTransaction.Commit
													objSystemLog.LogNormalDetail """"& strEvent &""" touch event has been enabled for the role " & strRoleCode & "."
												else
													objSystemLog.LogErrorDetail "","","'" & objval.Result.Message & "'" 
												end if
											Else
												If Not objRoleTouchEvent.IsNull Then
													objRoleTouchEvent.delete objTransaction
													set objval = objTransaction.validate		
													if objval.status <> 3 then
														objTransaction.Commit
														objSystemLog.LogNormalDetail """"& strEvent &""" touch event has been disabled for the role " & strRoleCode & "."
													else
														objSystemLog.LogErrorDetail "","","There were some error in disabling the touch event : '" & objval.Result.Message & "'" 
													end if
												Else
													objSystemLog.LogNormalDetail """"& strEvent &""" touch event is already disabled for the role " & strRoleCode & "."
												End If
											End If
											SET_NOTHING Array(objTransaction,objval)
										On Error Goto 0
									Else
										objSystemLog.LogNormalDetail "The " & strEvent & " touch event was not found."
									End If								
									set objTouchEvent = nothing
									set objRoleTouchEvent = nothing								
								End If
							Next
						Next
					Else 
						objSystemLog.LogNormalDetail "The " & ObjRoleCode & "role code was not found."
					End If
				End If
			Loop
		End If
		objSystemLog.LogNormalDetail "Completed execution to enable the touch event."
	End If
End Function

'2015-03-30, Jhirpara Function to get username and password from user.
'parameters:
'	objWF:			Workflow object
'	constantName	property name
'	PropertyValue	property value
Function updateCreateWFProperty(objWF, constantName, PropertyValue)
	'RegExp used to read the constant value
	'Pattern to match with the constant we want to read
		'Check If that property exists
'praval 20150612 #1574 - Deactivate SKUs Workflow - deliver globally
		Set objScan = objWF.Folders.WorkflowProperties.First("UCase(wfProperty)='" & UCase(constantName) & "'")
		If objScan.IsNull Then
			'We need to create a new property
			Set objScan = objWF.Folders.WorkflowProperties.CreateNewInstance()
			if isempty(constantName) then
				objScan.wfProperty = "Configuration Property"
				objScan.wfValue = PropertyValue
			else
				objScan.wfProperty = constantName
				objScan.wfValue = PropertyValue
			end if
		'ddave 14/07/2017 Added code to update wfProperty("Workflows Monitored") in email alert	
		ElseIf constantName = "Export Configuration" OR constantName = "Workflows Monitored" then
			objScan.wfValue = PropertyValue
		Else
			objScan.wfValue = PropertyValue
		End If
		'Save the property
		objScan.Save		'We put the new value
		updateCreateWFProperty = True
End Function 
'praval 20150626 Update the Export Configuration
Function updateExportWFProperty(objWF, constantName, PropertyValue)
        'ddave 14/07/2017 Added code to update wfProperty("Workflows Monitored") in email alert	
		If constantName = "Export Configuration" OR constantName = "Workflows Monitored" Then
		Set objScan = objWF.Folders.WorkflowProperties.First("UCase(wfProperty)='" & UCase(constantName) & "'")
			objScan.wfValue = PropertyValue
		End If
		'Save the property
		objScan.Save		'We put the new value
		updateExportWFProperty = True
End Function 


'2015-03-30, Jhirpara Function to add workflow with property or without properties.
'parameters:
'	objIE:			internet explorer object to print status
'	objCxExport : 	route import dictionary object which contain list of system to import route
'	System:			Import system object
'	objAllWF : 		Dictionary object for all remain workflow
'	WFSubFld:		workflow folder which contain xml of workflow.
'	strQtyMember : 	its system name/ market name
'SGHUNCHALA & KRAVAL - 20190207 - Update function for standard order export workflows
Sub ADDWFFolder(System, objOrderExport, objAllWF, WFSubFld, strQtyMember,objInstance)
	Dim strAction, strActionstart, strActionIndex, strEntry, strEntryIndex
	
	Set objWFFDirectory = Nothing
	Set objWFFDirectory = objFSO.GetFolder(WFSubFld)
	For Each objWFFile In objWFFDirectory.Files	
		If Not objWFFile Is Nothing Then
			parts = split(objWFFile.Name, ".")
			If UCase(parts(1))="XML" Then
				intFileCount = intFileCount + 1
				'Fetch the processid
				strProcID  = parts(0)
				
				If UCase(strProcID) = "STANDARD EXPORT" Then
					strWFtitle = objInstance.Title
				Else
					StrWFP = ""
					StrWFP = split(parts(0),Chr(45))
					if (ubound(StrWFP) <> 1) then
						StrWFP = split(parts(0),Chr(150))
						if (ubound(StrWFP) <> 1) then
							StrWFP = split(parts(0),Chr(151))
						End if
					end if	

					If (ubound(StrWFP) >= 1) then
						strProcID  = Trim(StrWFP(1))
						
						'DNimbark 04/20/2017 Code to escape '-'
						strWFtitle = Replace(StrWFP(0),"#DASH#","-")
					Else
						'objIE.document.write "File name is incorrect. File name must contain hyphen separated to workflow name and process code. :'" & parts(0) & "'."
						blnWFProcess = False
					End If
	   		     End If
				'DNimbark 20160519 - Added the code to read the special characters.
				'Read contents of file
				intLineSepratores = GetLineSeparator(objWFFile.Path)
				Set objStream = CreateObject("ADODB.Stream")
				objStream.CharSet = "utf-8"
				objStream.LineSeparator = intLineSepratores
				objStream.Open
				objStream.LoadFromFile(objWFFile.Path)				
				
				If objStream.EOS Then
					strContent = ""
				Else
					strContent = objStream.ReadText()
				End If
				
				'Fetch instance and update Defn
				If Len(strContent)>0 Then
					On Error Resume Next
						'SGHUNCHALA & KRAVAL - 20190207 - Update code to get workflow using processID excepting standard order export workflow
						If UCase(strProcID) <> "STANDARD EXPORT" Then
							Set objInstance = Nothing
							Set objInstance = System.Folders.WF_Processes.Scan("UCase(WF_ProcID)='" & UCase(strProcID) & "'").Fetch
						End If
					On Error Goto 0
					
					If objInstance Is Nothing Then
						Set objInstance = System.Folders.WF_Processes.CreateNewInstance()
						intAdd = intAdd + 1
						'iActive = InputBox( objWFFile.Name, "Enter Active Status for file ", "Enter 1 for Active or 0 for Inactive.")
						objInstance.WF_ProcID = strProcID
						objInstance.Title = strWFtitle
						objInstance.Active = True
					Else
						intUpdate = intUpdate + 1
						If not isnull(objInstance.Defn_diff.value) then	
							System.ConsolidateDefn(objInstance)
							objInstance.Save
						End if
						objInstance.Title = strWFtitle
					End If
					objInstance.ReExecution.Condition.Value = 3 'Abandoned or Errored
					objInstance.ReExecution.Method.Value = 1 'Restart
					objInstance.Version.Version = "1.0"
					strConfigFile = ""
					If strProcID = "EXP_TURN_IN" Then
						strConfigFile = objExpTurn.Item(UCase(strQtyMember))
						'DNimbark 20160615 - check for the config property path for Market marked as 'ALL'
						If strConfigFile = "" Then
							strConfigFile = objExpTurn.Item("ALL")
						End If
					ElseIf strProcID = "MOVE_ORDERS_TO_FTP" Then
						strConfigFile = objMoveOrdertoFtp.Item(UCase(strQtyMember))
						'DNimbark 20160615 - check for the config property path for Market marked as 'ALL'
						If strConfigFile = "" Then
							strConfigFile = objMoveOrdertoFtp.Item("ALL")
						End If
					ElseIf strProcID = "ORDER_EXPORT_STD" Then
						strConfigFile = objOrderExport.Item(UCase(strQtyMember))
						'DNimbark 20160615 - check for the config property path for Market marked as 'ALL'
						If strConfigFile = "" Then
							strConfigFile = objOrderExport.Item("ALL")
						End If
					Else
						If(objAllWF.Exists(UCase(strProcID))) Then
							strConfigFile = objAllWF.Item(UCase(strProcID))
						Else
							If(objAllWF.Exists(trim(UCase(strProcID) &"-"& UCase(strQtyMember)))) Then
								strConfigFile = objAllWF.Item(trim(UCase(strProcID) &"-"& UCase(strQtyMember)))
							End If
						End If
					End if
					
					if strConfigFile <> "" Then
						Dim objTextCFGStream
						'Read contents of file
						Set objTextCFGStream = objFSO.OpenTextFile(strConfigFile)
						'DNimbark 20170120 - Function to copy XSD file to the path provided in WF property
						If strProcID = "VAN_AL_IMP" OR strProcID = "VAN_AL_EXP" Then
							If strProcID = "VAN_AL_EXP" Then
								strXSDPropName = "SchemaPath:~"
							Else
								strXSDPropName = "XSDPath:~"
							End If
							strContent = CopyXSDFile(objInstance,objTextCFGStream.ReadAll,objWFFDirectory,strXSDPropName,strContent)
						Else
						 strobjcfgfile = ReadCFGInfoFromFile(objInstance, objTextCFGStream, objWFFDirectory, strContent)
						End If
					     objTextCFGStream.close
					End if
					
					'Do not update schedule
					' If Not objInstance.IsNew AND objInstance.WF_ProcID <> "GMDELSTOREVISIT" Then
						' If(InStr(objInstance.Defn,"<entryevents>")) <> 0 Then
							' strEntryIndex= InStr(objInstance.Defn,"</entryevents>")
							' strEntry = Left(objInstance.Defn,strEntryIndex+ 14) 
						' End If
						
						' strActionstart= InStr(strContent,"</entryevents>")
						' strAction = Mid(strContent,strActionStart + 14,Len(strContent))
						' strContent = strEntry + strAction
					' End If
					
					'DNimbark 20160519 - Empty the Def_diff if not null
					If not isnull(objInstance.Defn_diff.value) then	
						objSystem.ConsolidateDefn(objInstance)
						objInstance.Save
					End if
					
					'DNimbark 20160519 - Removed the unwanted characters from the script
					'DNimbark 20160627 - Compare the WF code version #
					If checkVersion(objInstance.Defn,Replace(strContent,"ï»¿","")) Or UCase(objInstance.WF_ProcID) = "IMP_PLANOGRAM" Then
						objInstance.Defn = Replace(strContent,"ï»¿","")
					End If
					'KRaval 11/05/2018 Added code to add the System Event Log.
					On error resume next
						objInstance.Save
						If Err.Number <> 0 Then
							objSystemLog.LogNormalDetail "Error occurred while saving Workflow. Please check the Error Log."
							ObjSystemLog.LogErrorDetail "","","Error : " & Err.Description
						Else
							objSystemLog.LogNormalDetail "Added or Updated for the Workflow/" & objInstance.Title
						End If
					On error goto 0
				End IF
				'Close stream
				'objTextStream.Close
				If Not objStream Is Nothing Then
					If objStream.State <> 0 Then
						On Error Resume Next
							objStream.Close
						On Error Goto 0
					End If
				End If
			end if
		End If
	Next
End sub	
'praval 20150626 Update the Export Configuration 
Sub ChangeProperty(objInstance, strQtyMember)
		strConfigFile = ""
		strProcID = "ORDER_EXPORT_STD"
			If strProcID = "ORDER_EXPORT_STD" Then
				strConfigFile = objOrderExport.Item(UCase(strQtyMember))
			End if
			If strConfigFile <> "" Then
				Dim objTextCFGStream
				'Read contents of file
				Set objTextCFGStream = objFSO.OpenTextFile(strConfigFile)
				call ExportCFGInfoFromFile(objInstance, objTextCFGStream)
				objTextCFGStream.close
			End IF	
End sub	
'2015-03-30, Jhirpara Function to readconfigration file and check if it contain schema path or not.
'if schema path is present then it will create schema path and transfer xsd file on that path.
'parameters:
'	objInstance:		Workflow instance
'	objTextStream:		configuration property in stream format
'	objWFFDirectory : 	its directory which is having xsd file as well as xml file
'	strContent: 		Xml file content
Function ReadCFGInfoFromFile(objInstance, objTextStream, objWFFDirectory, strContent)
		strLine = ""
		dim i, s
		dim blnEmpty
		dim WM : WM = False
		While Not objTextStream.AtEndOfStream
		i = 0
		s = 1
		blnEmpty = False
			strLine = Trim(objTextStream.ReadLine)
			If Len(strLine)>0 Then
				IF InStr(strLine, ":~") Then
					Value1 = Split(strLine,":~")
					if (value1(0) = "SchemaPath") then
						For Each objXSDFile In objWFFDirectory.Files	
							If Not objXSDFile Is Nothing Then
								xsdparts = split(objXSDFile.Name, ".")
								If UCase(xsdparts(1))="XSD" Then
									blnSchema = false
									strTempPath = ""
									strTempPath = objWFFDirectory.Path
									
									strXSDPath = value1(1)
									If Right(strXSDPath,1)<> "\" Then
										strXSDPath = strXSDPath & "\"
									End If
									if strXSDPath <> "" then
										If createFolderIfNotExist(strXSDPath) Then
											objFSO.CopyFile objXSDFile, strXSDPath
											strXSDFile = strXSDPath & objXSDFile.Name
										End If
									
										strTemp = ""
										strTemp = strXSDFile
										
										if (InStr(strContent,objXSDFile.Name)) > 1 then
											strContent = Replace(strContent,(MID(strContent,(((Instr(strContent,objXSDFile.Name)) - 3)),Len(objXSDFile.Name)+3)),strTemp)
										end if
										blnSchema = True
									End if
								end if
							End If
						Next
						s = 0
					Elseif (value1(1) = "") then
						Do while not objTextStream.AtEndOfStream
							strtemp = objTextStream.ReadLine
							'ddave 14/07/2017 Added code to update wfProperty("Workflows Monitored") in email alert	
							If value1(0) = "Configuration Property" OR value1(0) = "Workflows Monitored" Then
								If Value1(1) = "" Then
									Value1(1) = strtemp
								Else
									Value1(1) = Value1(1) & vbNewLine & strtemp
								End If
								WM = True
							ElseIf ((strtemp = "</folders>") Or (strtemp = "</xsl:stylesheet>")) then
								strLine = strLine & VBNewLine & strtemp
								Value1 = Split(strLine,":~")
								exit do
							else
								strLine = strLine & VBNewLine & strtemp
							end if	
						Loop
					end if
					i = 1
				'KRaval 02/05/2018 Added the code to update the WF properties having blank value.	
				ElseIf InStr(strLine, "::") Then 
					Value1 = Split(strLine,"::")
					i = 1
					blnEmpty = True
				Else
					Value1 = Value1 & strLine
				End if
						
				if (i=1 and s=1) then
					If WM Then 
						updateCreateWFProperty objInstance, Value1(0), Value1(1)
					Else
						If Instr(strLine,vbNewLine) > 0 Then
							updateCreateWFProperty objInstance, Value1(0), split(strLine,":~"&vbnewline)(1)
						Else
							If Not blnEmpty Then
								updateCreateWFProperty objInstance, Value1(0), split(strLine,":~")(1)
							Else
								updateCreateWFProperty objInstance, Value1(0), ""
							End If
						End If
					End If
					Value1 = ""
				end if
			End If
		Wend
		ReadCFGInfoFromFile = strContent
End Function		
'praval 20150626 Update the Export Configuration 
Function ExportCFGInfoFromFile(objInstance, objTextStream)
		strLine = ""
		dim i, s
		While Not objTextStream.AtEndOfStream
		i = 0
		s = 0
			strLine = Trim(objTextStream.ReadLine)
			If Len(strLine)>0 Then
				IF InStr(strLine, ":~") Then
					Value1 = Split(strLine,":~")
					'ddave 14/07/2017 Added code to update wfProperty("Workflows Monitored") in email alert	
					if (value1(0) = "Export Configuration" OR value1(0) = "Workflows Monitored") then
						if (value1(1) = "") then
							Do while not objTextStream.AtEndOfStream
								strtemp = objTextStream.ReadLine
								If ((strtemp = "</folders>") Or (strtemp = "</xsl:stylesheet>")) then
									strLine = strLine & VBNewLine & strtemp
									Value1 = Split(strLine,":~")
									exit do
								else
									strLine = strLine & VBNewLine & strtemp
								end if	
							Loop
						End If
						s = 1
					End if					
					i = 1
				Else
					Value1 = Value1 & strLine
				End IF
					if (i=1 and s=1) then
						call updateExportWFProperty (objInstance, Value1(0), Value1(1))
						Value1 = ""
					end if
			End If
		Wend
End Function
'2015-03-30, Jhirpara Function to Import standard import workflow with all the property.
'Parameters:
'	System:				System assigned
'	strImportFolder : 	import folder which contain list of folders
'	strConfigFile : 	configuration file which is use to import
Function ImportProcessWorkflow(System,strImportFolder,strConfigFile)
	'We remove the wait message
	HideWaitMessage
	Dim objTextStream
	'Read contents of file
	Set objTextStream = objFSO.OpenTextFile(strConfigFile)
	'Create the object with default info
	Set System_Settings = new SystemSettings
	'We read until :GeneralSettings
	System_Settings.CrmSystemName = System.Name
	System_Settings.ReadInfoFromFile(objTextStream)
	'Initialize vables
	intUpdate = 0
	intAdd = 0
	intError = 0
	intComplete = 0
	'We read the default import XML definition if it exists
	'DNimbark 12/10/2015 removed the unwanted content.  APN - Not working for std imports so switched around some lines.
	strDefaultWFImportXMLDef = ReadFileInfo(strImportFolder & IMPORT_WF_XML_DEF)
	strDefaultWFImportXMLDef = Replace(strDefaultWFImportXMLDef,"ï»¿","")
	strDefaultWFXMLDef = strDefaultWFImportXMLDef
	'We check If we have all necessary info to create the WF in the system
	strEnoughInfo = System_Settings.EnoughInfoToCreateWFNoUser
	If Len(strEnoughInfo)>0 Then
		'We do not have enough info on Config file
		WriteErrorLine("ERROR: " & strEnoughInfo)
		intError = intError + 1
	Else
		'We need to create the WF process in the system
		'We were able to log into the system
		'We create the Import folder structure
		'Jhirpara 20142612 - defect, folder with the market code isn’t getting created
		'WriteCommentLine("Creating Import folder structure")
		IF createFolderIfNotExist(objFSO.getparentfoldername(System_Settings.SystemImportFolder)) then
			If CreateImportFolderStructure(System_Settings.SystemImportFolder, System_Settings.ObjectToFolderInfo) Then
				'WriteCommentLine("<span style='color:green'>OK</span>")
			End If
		End if
		'We create the FTP folder structure
		'WriteCommentLine("Creating FTP folder path")
		'If createFolderIfNotExist(System_Settings.FtpFolderPath) Then
		If Not objFSO.FolderExists(System_Settings.FtpFolderPath) Then
			WriteCommentLine("You need to create the folder '" & System_Settings.FtpFolderPath & "'")
			'WriteCommentLine("<span style='color:green'>OK</span>")
		End If
		'WriteCommentLine("Creating FTP Error folder path")
		'If createFolderIfNotExist(System_Settings.FtpErrorFolderPath) Then
		If Not objFSO.FolderExists(System_Settings.FtpErrorFolderPath) Then
			WriteCommentLine("You need to create the folder '" & System_Settings.FtpErrorFolderPath & "'")
			'WriteCommentLine("<span style='color:green'>OK</span>")
		End If
		'We create/update the WF process to move files from FTP folder to system import folder
		blnResult = createUpdateFTPWFprocess(System, strImportFolder & WF_XMLFILE_FILES_FROM_FTP, WF_TITLE_FILES_FROM_FTP, WF_ID_FILES_FROM_FTP, _
			Array(Array("WFPROPERTY_FTP_FOLDER_PATH", System_Settings.FtpFolderPath), _
				Array("WFPROPERTY_CURRENT_MARKET", System_Settings.SystemMarket), _
				Array("WFPROPERTY_CURRENT_ENVIRONMENT", System_Settings.SystemEnvironment), _
				Array("WFPROPERTY_FOLDER_TO_OBJECT", System_Settings.ObjectToFolderInfo), _
				Array("WFPROPERTY_IMPORT_FOLDER", System_Settings.SystemImportFolder)))
		WriteCommentLine("Creating/updating WF import processes to move the Files to FTP Folder in this system  -- <span style='color:green'>Done.</span>")
		'We create/update the WF process to move error files from system import folder to Error FTP folder
		blnResult = blnResult And createUpdateFTPWFprocess(System, strImportFolder & WF_XMLFILE_ERROR_FILES_TO_FTP, WF_TITLE_ERROR_FILES_TO_FTP, WF_ID_ERROR_FILES_TO_FTP, _
			Array(Array("WFPROPERTY_FTP_ERROR_FOLDER_PATH", System_Settings.FtpErrorFolderPath), _
				Array("WFPROPERTY_IMPORT_FOLDER", System_Settings.SystemImportFolder)))
		WriteCommentLine("Creating/updating WF process to move Error files to FTP Folder  -- <span style='color:green'>Done.</span>")
		
		'We create/update the WF process to move success files from system import folder to Success FTP folder
		blnResult = blnResult And createUpdateFTPWFprocess(System, strImportFolder & WF_XMLFILE_SUCCESS_FILES_TO_FTP, WF_TITLE_SUCCESS_FILES_TO_FTP, WF_ID_SUCCESS_FILES_TO_FTP, _
		Array())
		WriteCommentLine("Creating/updating WF import processes to move Success file to FTP folder  -- <span style='color:green'>Done.</span>")	
		
		'We are going to create/update all necessary WF processes
		'We show the wait message
		ShowWaitMessage
		intSysAdd = 0
		intSysUp  = 0
		
							
		'We call a function to create all Import WF processes
		strResult = CreateUpdateWFImportProcesses(System,strImportFolder, strDefaultWFImportXMLDef, System_Settings.SystemImportFolder,intSysAdd, intSysUp)
		'We remove the wait message
		HideWaitMessage
		intAdd = intAdd + intSysAdd
		intUpdate = intUpdate + intSysUp
		If Len(strResult)=0 Then
			'WriteCommentLine("<span style='color:green'>OK</span>  added:" & intSysAdd & " updated:" & intSysUp)
		Else
			intError = intError + 1
			WriteErrorLine(strResult)
			blnResult = false
		End If
		If blnResult Then
			intComplete = intComplete + 1
		End If
	End If
	
	'WriteLineSeparator
	'Completed adding processes
	objIE.document.write "<font size=3 color=black>"	
	WriteStandardLine("imports WF updated: <b>" & intUpdate & "</b>, imports WF added:<b>" & intAdd & "</b>")
End Function

'Functions to use dIfferent styles to show the info on window log
Function WriteStandardLine(strLine)
	objIE.document.write "<span style='color:black;font-size:medium'>" & strLine & "<span><br>"
End Function
Function WriteCommentLine(strLine)
	objIE.document.write "<span style='color:black;font-size:small'>" & strLine & "<span><br>"
End Function
Function WriteErrorLine(strLine)
	objIE.document.write "<span style='color:red;font-size:small'>" & strLine & "<span><br>"
End Function
Function WriteLineSeparator()
	objIE.document.write "<hr>"
End Function
Function ShowWaitMessage()
	objIE.document.write "<span id='waitMessage'>wait .</span><script>ShowWaitMessage(); var waitMessagefn = setInterval(function(){ShowWaitMessage()},100);</script>"
End Function
Function HideWaitMessage()
	objIE.document.write "<script>clearInterval(waitMessagefn); var elem=document.getElementById('waitMessage');elem.parentNode.removeChild(elem);</script>"
End Function

'2015-03-30, Jhirpara Function to select specific subfolder from list and return its path value.
'Parameters:
'	objFolder:		Folder which contain list of folders
'	strListFor : 	Environment
Function GetSelectedSubFolderPath(objFolder,strListFor)
	strSubfolder = AddEndSlashIfNecessary(objFolder)
	Set objDirectory = objFSO.GetFolder(strSubfolder)
	'We go over each subfolder
	strHtmlSelectOptions = ""
	For Each SubFld in objDirectory.SubFolders
		strHtmlSelectOptions = strHtmlSelectOptions & "<option value='" & SubFld.Name & "'>" & SubFld.Name & "</option>"
	Next
	
	strHtmlSelect = "<label>Select " & strListFor & " : </label><select style='width:400'  id='select_" & strListFor &"'><option value=''>Select " & strListFor & " </option>" & strHtmlSelectOptions & "</select>" & _
		"<br>"
	objIE.document.write strHtmlSelect
	strSelectedFile = ""
	While strSelectedFile = ""
			strSelectedFile = objIE.document.getElementById("select_" & strListFor &"").value
			Wscript.sleep(500)

		If Len(strSelectedFile)>0 Then
			If Not objFSO.FolderExists(AddEndSlashIfNecessary(objDirectory.path) & strSelectedFile) Then
				MsgBox "Sorry " & AddEndSlashIfNecessary(objDirectory.path) & strSelectedFile & " ." & strListFor & " does not exist, enter another " & strListFor & "."
				strSelectedFile = ""
				objIE.document.getElementById("select_" & strListFor &"").value = ""
			Else
				objIE.document.getElementById("select_" & strListFor &"").disabled = true
			End if
		End if
	Wend
	GetSelectedSubFolderPath = AddEndSlashIfNecessary(objDirectory.path) & strSelectedFile
End Function
'praval 20150612 #1574 - Deactivate SKUs Workflow - deliver globally
Function GetSubFolderPath(objFolder,strListFor)
	strSubfolder = AddEndSlashIfNecessary(objFolder)
	Set objDirectory = objFSO.GetFolder(strSubfolder)
		For Each SubFld in objDirectory.SubFolders
		If SubFld.Name = strListFor Then
		GetSubFolderPath = AddEndSlashIfNecessary(objDirectory.path) & strListFor
		End If
	Next	
End Function

'Function to read the file context
'Parameters:
'	strFileName:	Name of the file to read
'	return the content of the file or an empty string
Function ReadFileInfo(strFileName)
	FileContent = ""
	If objFSO.FileExists(strFileName) Then
		Set objFile = objFSO.GetFile(strFileName)
		Set objWFXMLStream = objFile.OpenAsTextStream(1,0)
		FileContent = objWFXMLStream.ReadAll
	End If
	ReadFileInfo = FileContent
End Function


'Function to create the system import folder structure If it does not exist
'Parameters:
'	strMainFolder: 		system import folder
'	strSubfolderList:	the object ID list where we have the subfolder name per each object ID
'	return true If all folders exist in the system
Function CreateImportFolderStructure(strMainFolder, strSubfolderList)
	strMainFolder = AddEndSlashIfNecessary(strMainFolder)
	blnError = False
	'check If system import folder exists If not we try to create it
	If createFolderIfNotExist(strMainFolder) Then
		'We get all lines of object ID list
		lstObjectIDs = Split(strSubfolderList,vbNewLine)
		For Each strLine in lstObjectIDs
			strLine = UCase(Trim(strLine))
			If strLine = "" Then
				'Skip empty lines
			Else
				Select Case Left(strLine, 1)
					Case "#":
						'Do nothing comment line
					Case Else
						'We read the subfolderName
						i = InStr(strLine, ",")
						If i > 0 Then
							strFolderName = Trim(Mid(strLine, i + 1))
							If Len(strFoldername)>0 Then
								strMainSubFolder = AddEndSlashIfNecessary(strMainFolder & strFolderName)
								'Create the new subfolder
								If createFolderIfNotExist(strMainSubFolder) Then
									'Create the necessary subfolders per each object ID import folder
									For Each item in lstImportSubfolders
										blnError = blnError Or Not createFolderIfNotExist(strMainSubFolder & item)
									Next
								Else
									blnError = True
								End If
							End If
						End If
				End Select
			End If
		Next
	Else
		blnError = True
	End If
	CreateImportFolderStructure = Not blnError
End Function

'Function to create If it does not exists a folder
'Parameters:
'	strFolderName: the folder name
'	return false If the folder does not exist or it could not be created
Function createFolderIfNotExist(strFolderName)
	createFolderIfNotExist = True
	If Not objFSO.FolderExists(strFolderName) Then
			objFSO.CreateFolder(strFolderName)
			blnError = Err <> 0
		If blnError Then
			WriteErrorLine("Error: We could not create the folder '" & strFolderName & "'")
		End If
		createFolderIfNotExist = Not blnError
	End If
End Function


'Function to create/update a WF process and update/create its properties
'parameters:
'	System: The CRM system object
'	WFxmlDefFileName: filename where we have the WF xml definition
'	WFTitle: the WF title
'	WFID: The WF Id
'	lstProperties: An array of arrays, the inner array has two elements, the first one is the WF property name and the second one the WF property value
'	return true If everything was OK
Function createUpdateFTPWFprocess(objSystem, WFxmlDefFileName, WFTitle, WFId,lstProperties)
	createUpdateFTPWFprocess = false
	'We check If xml definition file exists
	If objFSO.FileExists(WFxmlDefFileName) Then
		'We get or create the WF process
		Set objWF = getFTPWF(objSystem, WFTitle, WFId, WFxmlDefFileName)
		'We check If we could get/create the WF process
		If Not objWF is Nothing Then
			'The WF process was created/updated
			'WriteCommentLine("<span style='color:green'>OK</span>")
			'We need to create/update the WF properties, we need to check the property name into WF definition and put the value
			blnError = False
			For Each item in lstProperties
				blnError = blnError Or Not updateCreateWFFTPProperty(objWF, item(0),item(1))
			Next
			'Return If everything was OK
			createUpdateFTPWFprocess = Not blnError
		Else
			intError = intError + 1
			WriteErrorLine("Error creating WF process")
		End If
	End If
End Function 

'Function to get/create a FTP WF process
'Parameters
'	System: The CRM system object
'	WFTitle: the WF title
'	WFID: The WF Id
'	WFXmlFileName: filename where we have the WF xml definition
'	Return the WF process
Function getFTPWF(objSystem,WFTitle,WFID,WFXmlFileName)
	'We do a scan to look for the WF process
	Set objScan = objSystem.Folders.WF_Processes.Scan("UCase(WF_ProcID)='" & WFID & "'",,1)
	'We read the xml definition file
	ImportXMLFile = ReadFileInfo(WFXmlFileName)
	If objScan.EndOfScan Then
		'it does not exist we create it
		intAdd = intAdd + 1
		Set objWF = objSystem.Folders.WF_Processes.CreateNewInstance()
		objWF.WF_ProcID = WFID
		'We put all necessary info into the WF process
		objWF.Active = WF_FTP_ACTIVE
		objWF.Suspended = WF_SUSPENDED
		objWF.EnableLogging = WF_ENABLELOGGING
		objWF.NonSystem = WF_NONSYSTEM
		objWF.Priority = WF_PRIORITY
	Else
		'it exists so we update it
		intUpdate = intUpdate + 1
		Set objWF = objScan.Fetch
		'AAA 2014/03/06. We keep the current WF settings and the schedule info
		'AAA 2014/08/01. We update the code to just modify the vbscript action info
		Set re = New RegExp
		re.Global  = False
		re.IgnoreCase = True
		re.Pattern = "([\s|\S]*<Action\b.*AeWorkFlow\.VBScript\.Action[^>]*>)([\s|\S]*</Action>)"
		Set MatchesCurrentWF = re.Execute(objWF.Defn)   ' Execute search.
		Set MatchesNewWF = re.Execute(ImportXMLFile)   ' Execute search.
		strResult = ""
		If MatchesCurrentWF.Count>0 And MatchesNewWF.Count>0 Then
			If MatchesCurrentWF(0).SubMatches.Count>1 And MatchesNewWF(0).SubMatches.Count>1 Then
				ImportXMLFile= MatchesCurrentWF(0).SubMatches(0) & re.Replace(objWF.Defn,MatchesNewWF(0).SubMatches(1))
			End if
		End if	
	End If
	objWF.Title = WFTitle
	
	'Jhirpara 2015/05/19. #1579, Report files are not being diverted to the proper country
	'Consolidate defn with defn_diff and update with the latest XML
	If not isnull(objWF.Defn_diff.value) then	
		objSystem.ConsolidateDefn(objWF)
		objWF.Save
	End if
	
	objWF.NonSystem = WF_NONSYSTEM
	'DNimbark 20160627 - Compare the WF code version #
	If checkVersion(objWF.Defn,ImportXMLFile) Then
		objWF.Defn = ImportXMLFile
	End If
	'We save the WF process
	objWF.Save
	Set getFTPWF = objWF
End Function

'This function is used to create/update a property into an Import WF process.
'The property name is read from the value of the constant that it is passed as parameter
'We read this constant from the definition info of the WF process
'Parameters:
'	objWF: 			The WF process
'	constantName:	Name of the constant where we are going to read the property name to update
'	PropertyValue:	Value to put in the WF property
'	Return true If property was updated/created
Function updateCreateWFFTPProperty(objWF, constantName, PropertyValue)
	'RegExp used to read the constant value
	Set re = New RegExp
	re.Global  = False
	re.IgnoreCase = True
	'Pattern to match with the constant we want to read
	re.Pattern = "\s*Const\s+" & constantName & "\s*=\s*""([^""]*)"""
	Set Matches = re.Execute(objWF.Defn)   ' Execute search.
	strResult = ""
	If Matches.Count>0 Then
		If Matches(0).SubMatches.Count>0 Then
			'We get the value of the constant
			strResult = Trim(Matches(0).SubMatches(0))
		End If
	End If
	'Check If we read the constant value
	If Len(strResult)>0 Then
		'Check If that property exists
		Set objScan = objWF.Folders.WorkflowProperties.Scan("wfProperty='" & strResult & "'",,1)
		If objScan.EndOfScan Then
			'We need to create a new property
			Set objWFProperty = objWF.Folders.WorkflowProperties.CreateNewInstance()
			objWFProperty.wfProperty = strResult
		Else
			Set objWFProperty = objScan.Fetch
		End If
		'We put the new value
		objWFProperty.wfValue = PropertyValue
		'Save the property
		objWFProperty.Save
		updateCreateWFFTPProperty = True
	Else
		WriteErrorLine("Error constant WF property " & constantName & " not found in WF definition")
		updateCreateWFFTPProperty = False
	End If
End Function 

'This function is to create all necessary WF import processes
'We read a subfolder where we have all config files and the WF XML Definition
'Parameters:
'	objSystem:	CRM system
'	strSubFolder:	The subfoldername where the necessary files to create the import WF processes are
'	defaultWFImportXMLdef: The default WF XML def that we have in main folder
'	strImportFolderPath:	The subfolder where import files are for this system
'	intAdd:	We will return the number of WF processes created
'	intUpdate: We will return the number of WF processes updated
'	Return a string with the error, If return an empty string everything went OK
Function CreateUpdateWFImportProcesses(objSystem,strSubfolder, defaultWFImportXMLdef, strImportFolderPath,byRef intAdd, byRef intUpdate)
	strSubfolder = AddEndSlashIfNecessary(strSubfolder)
	Set objDirectory = objFSO.GetFolder(strSubfolder)
	'We go over each subfolder
	For Each SubFld in objDirectory.SubFolders
		'First we read the xml def for the WF process. If it does not exist we use the default one in main folder
		strWFXMLDef = ReadFileInfo(AddEndSlashIfNecessary(SubFld.Path) & IMPORT_WF_XML_DEF)
		'AAA 2014/04/08. We get the number of files in the subfolder
		intFilesInSubFolder = SubFld.Files.Count
		'DNimbark 20170629 - Do not update Costume Imports with std code
		blnCostumeWF = False
		If Len(strWFXMLDef)=0 Then
			strWFXMLDef=defaultWFImportXMLdef
		Else
			'AAA 2014/04/08. We have the xml def file in subfolder, so it is not a WF property
			intFilesInSubFolder = intFilesInSubFolder - 1
			'DNimbark 20170629 - Do not update Costume Imports with std code
			blnCostumeWF = True
		End If
		'Check if we have a xml definition
		If Len(strWFXMLDef)>0 Then
			'We go over each file inside of the subfolder
			For Each objFile In SubFld.Files	
				If Not objFile Is Nothing Then
					parts = split(objFile.Name, ".")
					'AAA 2014/08/01. We check if we only have on "." in the file name
					If Ubound(parts)=1 Then
						'We check if it is a txt and not is "Import XML" one
						If UCase(parts(1))="TXT" And UCase(parts(0)) <> "IMPORT XML"  Then
							'we create a new WF property or a new WF process
							intFileCount = intFileCount + 1
							'Fetch the processid
							strProcID  = parts(0)
							StrWFP = split(parts(0),"- ") 
							strProcID  = Trim(StrWFP(0))
							'DNimbark 20170629 - Do not update Costume Imports with std code
							If blnCostumeWF Then
								strLstCustomeImports = strLstCustomeImports & strProcID & ","
							End If
							
							'Read contents of file
							Set objWFXMLStream = objFile.OpenAsTextStream(1,0)
							If objWFXMLStream.AtEndOfStream Then
								strPropertyValue = ""
							Else
								strPropertyValue = objWFXMLStream.ReadAll
								'We replace the some info in the property with the info on config file
								strPropertyValue = Replace(strPropertyValue,"<CLIENT>",objSystem.Name)
								strPropertyValue = Replace(strPropertyValue,"<_PUT_IMPORT_PATH_>",strImportFolderPath)
							End If		
							If Len(strPropertyValue)>0 Then
								'We check if current subfolder is the special one where each txt file must be a WF process
								'AAA 2014/02/20. Adam asked to create a WF process per each config file although it belongs to a WF process with a group of properties
								'so we create a new WF process per each config file and then we check if we need to create a group of import files
								
								'We need to create/update a new WF process for this file
								'We get the property name. We remove the first digits If they exist
								Set re = New RegExp
								re.Global  = False
								re.IgnoreCase = True
								re.Pattern = "\s*\d*(.*)"
								strPropertyNameNoIndex = StrWFP(1)
								Set matches = re.Execute(strPropertyNameNoIndex)
								If matches.Count > 0 Then
									strPropertyNameNoIndex = matches(0).Submatches(0)
								End If
							'We call a function to create/update the new process with default property
							'Jhirpara 2015/03/23. WF names add 2 spaces in between the hyphen and the import name.update those with 2 spaces to one space.
							If UCase(SubFld.Name) = IMPORT_INDEPENDENT_WF_PATH OR intFilesInSubFolder>=1 Then 
								CreateImportWFProces objSystem, Left(strProcID,20),Left("Import - " & trim(strPropertyNameNoIndex),50),strWFXMLDef,DEFAULT_WF_PROPERTY_NAME,strPropertyValue,intAdd,intUpdate,False
							End If
							If UCase(SubFld.Name) <> IMPORT_INDEPENDENT_WF_PATH Then
								'We need to create/update a new property for this file. Subfolder name will be the WF Title
								CreateImportWFProces objSystem, Left("IMP_" & UCase(SubFld.Name),20),Left("Import - " & UCase(SubFld.Name),50),strWFXMLDef,StrWFP(1),strPropertyValue,intAdd,intUpdate,True
							End If
						Else
							'file does not have any info we do not need to create a WF Process
						End If
						End If
					End If
				End If
			Next
		Else
			CreateUpdateWFImportProcesses = "ERROR we do not have an Import WF Definition for '" & SubFld.Name & "'"
		End if
	Next
End Function

'This function creates the WF import process with the info that is passed by parameters
'We can create/update a WF process or create/update a property depending on the propertyName
'Parameters:
'	objSystem: CRM system
'	strProcID:	WF Process ID
'	strWFTilte:	WF Title
'	strWFXMLDef:	WF XML definition
'	strPropertyName:	Property Name to create/update
'	strPropertyValue:	Property Value
'	intAdd:	We will return the number of WF processes created
'	intUpdate: We will return the number of WF processes updated
'   blnIsBundel : True if the WF is bundle
Sub CreateImportWFProces(objSystem, strProcID,strWFTitle,strWFXMLDef,strPropertyName,strPropertyValue,byRef intAdd, byRef intUpdate, blnIsBundel)
	Set objImportWFProcess = Nothing
	Set objWFProperty = Nothing
	On Error Resume Next
	Set objImportWFProcess = objSystem.Folders.WF_Processes.Scan("UCase(Trim(WF_ProcID))='" & UCase(strProcID) & "'").Fetch
	'20170406 Dnimbark - Fix for generating duplicate bundle WF.
	If (objImportWFProcess IS Nothing OR objImportWFProcess.IsNull) AND blnIsBundel Then
		Set objImportWFProcess = objSystem.Folders.WF_Processes.Scan("UCase(Trim(Title))='" & UCase(strWFTitle) & "'").Fetch
	End If
	On Error goto 0
	Dim strOriginalTitle,strNewTitle
	If objImportWFProcess Is Nothing Then
		Set objImportWFProcess = objSystem.Folders.WF_Processes.CreateNewInstance()
		intAdd = intAdd + 1
		objImportWFProcess.WF_ProcID = strProcID
		objImportWFProcess.Title = strWFTitle
		objImportWFProcess.Active = WF_IMPORT_ACTIVE
		objImportWFProcess.Suspended = WF_SUSPENDED
		objImportWFProcess.EnableLogging = WF_ENABLELOGGING
		objImportWFProcess.NonSystem = WF_NONSYSTEM
		objImportWFProcess.Priority = WF_PRIORITY
	Else
		'AAA 2014/03/06. We keep the current WF settings and the schedule info
		'AAA 2014/08/01. We update the code to update just the vbscript action
		'Jhirpara 2015/03/23. WF names add 2 spaces in between the hyphen and the import name.update those with 2 spaces to one space.
		objImportWFProcess.Title = strWFTitle
		objImportWFProcess.NonSystem = WF_NONSYSTEM
		Set re = New RegExp
		re.Global  = False
		re.IgnoreCase = True
		re.Pattern = "([\s|\S]*<Action\b.*AeWorkFlow\.VBScript\.Action[^>]*>)([\s|\S]*</Action>)"

		'Jhirpara 2015/02/18. #1345, Standard Import Release Script needs fixing
		'Jhirpara 2015/02/26. #1367, Standard Import Workflow Release script is changing existing schedules to 00:00
		'Consolidate defn with defn_diff and update with the latest XML
		If not isnull(objImportWFProcess.Defn_diff.value) then	
			objSystem.ConsolidateDefn(objImportWFProcess)
			objImportWFProcess.Save
		End if
		
		'DNimbark 13/10/2015 To update the title of the XML
		strOriginalTitle = "title=""" &  Split(Split(objImportWFProcess.Defn.value,"title=""")(1),"""")(0) & """"
		strNewTitle = "title=""" &  Split(Split(strWFXMLDef,"title=""")(1),"""")(0) & """"
		
		Set MatchesCurrentWF = re.Execute(objImportWFProcess.Defn)   ' Execute search.
		Set MatchesNewWF = re.Execute(strWFXMLDef)   ' Execute search.
		strResult = ""
		If MatchesCurrentWF.Count>0 And MatchesNewWF.Count>0 Then
			If MatchesCurrentWF(0).SubMatches.Count>1 And MatchesNewWF(0).SubMatches.Count>1 Then
				strWFXMLDef= MatchesCurrentWF(0).SubMatches(0) & re.Replace(objImportWFProcess.Defn,MatchesNewWF(0).SubMatches(1))
			End if
		End if
		intUpdate = intUpdate + 1
	End If
	'DNimbark 13/10/2015 To update the title of the XML
	strWFXMLDef = Replace(strWFXMLDef,strOriginalTitle,strNewTitle)
	'DNimbark 20160627 - Compare the WF code version #
	If checkVersion(objImportWFProcess.Defn,strWFXMLDef) Then
		objImportWFProcess.Defn = strWFXMLDef
	End If
	objImportWFProcess.Save
	
	'We check the property name to see if it is a WF process with only one property or not
	If DEFAULT_WF_PROPERTY_NAME = strPropertyName Then
		' Set The Workflow Property From Text File.
		On Error Resume Next
		Set objWFProperty = objImportWFProcess.Folders.WorkflowProperties.Scan("UCase(wfProperty)='" & UCase(strPropertyName) & "'").Fetch
		On Error goto 0
	Else
		'Many properties per each WF process so we check if the property exists but we do not take into account the first digits
		Set re = New RegExp
		re.Global  = False
		re.IgnoreCase = True
		re.Pattern = "\s*\d*(.*)"
		strPropertyNameNoIndex = strPropertyName
		Set matches = re.Execute(strPropertyNameNoIndex)
		If matches.Count > 0 Then
			strPropertyNameNoIndex = matches(0).Submatches(0)
		End If
		'As the WF property name can begin with a index value or not and that index value can be modified to put a new order it is not easy to
		'find a filter expression to be sure that we select the right property, so we go over each property to know if it has the same name without
		'index info
		For Each objTempWFProperty In objImportWFProcess.Folders.WorkflowProperties.Scan(,"wfProperty;wfValue")
			strTmpPropertyName = objTempWFProperty.wfProperty
			'We get the property name without the first digits
			Set matches = re.Execute(strTmpPropertyName)
			If matches.Count > 0 Then
				strTmpPropertyName = matches(0).Submatches(0)
			End If 
			If UCase(Trim(strTmpPropertyName))=UCase(Trim(strPropertyNameNoIndex)) Then
				'It is the same WF property
				Set objWFProperty = objTempWFProperty
				Exit For
			End if
		Next	
	End if
	
	If objWFProperty Is Nothing Then
		Set objWFProperty = objImportWFProcess.Folders.WorkflowProperties.CreateNewInstance()
	Else
		'DN 2017/02/16. Avoid updating mandatory flag update
		'strPropertyValue = keepMFlagUpdateOnlyUponInsert(objWFProperty.wfValue,strPropertyValue)
		'AAA 2014/03/06. We avoid updating the flag update only upon insert in the WF property if it is put to active
		'strPropertyValue = keepFlagUpdateOnlyUponInsert(objWFProperty.wfValue,strPropertyValue,False)
		'DN 2016/06/29. We avoid updating the default flag 
		'strPropertyValue = keepFlagUpdateOnlyUponInsert(objWFProperty.wfValue,strPropertyValue,True)
		'AAA 2014/12/1. We call a function to check the settings to keep in the WF property
		'strPropertyValue =keepWFProperties(objWFProperty.wfValue,strPropertyValue,lstWFPropertiesToKeep)
		
		'DNimbark 20160412 - We will not update the WF property PasswordEncryption, if exists.
		strPropName = "PasswordEncryption"
		If InStr(objWFProperty.wfValue,strPropName) > 0 Then
			strPropertyValue = keepPasswordEncryption(objWFProperty.wfValue,strPropertyValue,strPropName)
		End If
		
	End If
	objWFProperty.wfProperty = strPropertyName
	'KRAVAL - To Avoid Bad Config.
	If Len(objWFProperty.wfValue)>0 then
		objWFProperty.wfValue = ""
	End If
	objWFProperty.wfValue = strPropertyValue
	objWFProperty.Save
	
	If UCase(objImportWFProcess.WF_ProcID) = "IMP_ORDER" Then
					strResult = "Call Update Product Pack to Pricelist Product"
				
					'Check If that property exists
					Set objScan = objImportWFProcess.Folders.WorkflowProperties.Scan("wfProperty='" & strResult & "'",,1)
					If objScan.EndOfScan Then
						'We need to create a new property
						Set objWFProperty = objImportWFProcess.Folders.WorkflowProperties.CreateNewInstance()
						objWFProperty.wfProperty = strResult
					Else
						Set objWFProperty = objScan.Fetch
					End If
					'We put the new value
					objWFProperty.wfValue = "No"
					'Save the property
					objWFProperty.Save			
	End If
	
	Set objImportWFProcess = Nothing
	set objWFProperty = Nothing
End Sub

'This function is used to add "\" at the end of the folder If it does not exist
Function AddEndSlashIfNecessary(strFolderPath)
	strFolderPath = Trim(strFolderPath)
	If InStr(strFolderPath,"/") Then
		strSlash = "/"
	Else
		strSlash = "\"
	End If
	If (Right(strFolderPath,1)<>strSlash) Then
		strFolderPath = strFolderPath & strSlash
	End If
	AddEndSlashIfNecessary = strFolderPath
End Function

'AAA 2014/12/01
'Function to keep the property info if it already exists in the config property
Function keepWFProperties(strCurrentProperty,strNewProperty,lstProperties)
	Set re = New RegExp
	re.Global  = False
	re.IgnoreCase = True
	For Each strPropertyName in lstProperties
		re.Pattern = "\n\s*(" & strPropertyName & "\s*,[^\r|^\n]*)"
	Set Matches = re.Execute(strCurrentProperty)
		strPreviousInfo = ""
	If Matches.Count>0 Then
			strPreviousInfo = Matches(0).SubMatches(0)
			If Len(strPreviousInfo & "")>0 Then
			re.Pattern = "\n:importprop[^\r|^\n]*"
			Set Matches = re.Execute(strNewProperty)
			If Matches.Count>0 Then
					strNewProperty = re.Replace(strNewProperty,Matches(0) & vbNewLine & strPreviousInfo)
			End If
		End If		
	End If
	Next
	keepWFProperties = strNewProperty
End Function



'AAA 2014/03/06
'Function to keep the flag update only upon insert in the WF property
Function keepFlagUpdateOnlyUponInsert(strCurrentProperty,strNewProperty,blnIsDefaultFlag)
	'First we get the members where we have 0 in the flag "update only upon insert" it is the 4 paramemter in DataMapping
	Set lstFieldsOnlyUpdateUponInsert = CreateObject("scripting.dictionary")
	Set lstKeepDefaultValueFlag = CreateObject("scripting.dictionary")
	'We are going to use regular expression to get those members
	Set re = New RegExp
	re.Global  = True
	re.IgnoreCase = True
	re.Multiline = True
	'First we get all dataMapping info
	re.Pattern = "\n:DataMapping([\s,\S]*.*)"
	Set Matches = re.Execute(strCurrentProperty)   ' Execute search.
	strResult = ""
	If Matches.Count>0 Then
		If Matches(0).SubMatches.Count>0 Then
			strResult = Trim(Matches(0).SubMatches(0))
		End if
	End if
	'we get each line in datamapping where we have 4 properties, we check 3 "|" after the comma
	'DN 2016/06/29. We avoid updating the default flag , if blnIsDefaultFlag is True
	If blnIsDefaultFlag Then
		re.Pattern = "(.*,.*\|.*\|.*)"
	Else
		re.Pattern = "(.*,.*\|.*\|.*\|.*)"
	End If
	Set Matches = re.Execute(strResult)   ' Execute search.
	strResult = ""
	intAux = 0
	While (Matches.Count>intAux)
		If Matches(intAux).SubMatches.Count>0 Then
			'we get the member name and the property value
			strLine = Trim(Matches(intAux).SubMatches(0))
			'DN 2016/06/29. We avoid updating the default flag , if blnIsDefaultFlag is True
			If blnIsDefaultFlag Then
				re.Pattern = "(.*),.*\|.*\|(.*).*"
			Else
				re.Pattern = "(.*),.*\|.*\|.*\|\s*(\d).*"
			End If
			Set InfoInLine = re.Execute(strLine)
			If InfoInLine.Count>0 Then
				If InfoInLine(0).SubMatches.Count>0 Then
					strValue = Trim(InfoInLine(0).SubMatches(1))
					'if property value is different than 0 we add the member into the dictionary object
					If strValue<>"0" Then
						If blnIsDefaultFlag Then
							If Right(strLine,1) = Chr(13) Then
								strLine = Mid(strLine,1,len(strLine)-1)
							End If
							If Right(strLine,1) = "|" Then
								strLine = Mid(strLine,1,len(strLine)-1)
							End If
							
							arrPipeValues  = Split(strLine,"|")
							strPipecount   = UBound(arrPipeValues)
							strValue       = arrPipeValues(strPipecount)
							intUBound      = strPipecount
							strUpdateValue = ""
							
							If (strPipecount > 3) Then
								strValue       = arrPipeValues(strPipecount-1)
								strUpdateValue = arrPipeValues(strPipecount)
							ElseIf strPipecount = 3 Then
								If Not isExpressionValid(strCurrentProperty,arrPipeValues) Then
									strValue 	   = arrPipeValues(strPipecount-1)
									strUpdateValue = arrPipeValues(strPipecount)
								End If
							End If
							lstKeepDefaultValueFlag.add Trim(InfoInLine(0).SubMatches(0)), Replace(Replace(strValue,vbNewLine,""),Chr(13),"") & "-" & intUBound & "-" & Replace(Replace(strUpdateValue,vbNewLine,""),Chr(13),"")
						Else
							lstFieldsOnlyUpdateUponInsert.add Trim(InfoInLine(0).SubMatches(0)), strValue
						End If
					End if
				End if
			End if
		End if
		intAux = intAux + 1
	Wend
	'We check if we have some item in the dictionary
	If lstFieldsOnlyUpdateUponInsert.Count>0 OR lstKeepDefaultValueFlag.Count>0 Then
		'we get over each line of the property
		lstLines = Split(strNewProperty,vbNewLine)
		strUpdatedProperty = ""
		blnDataMappingSection = false
		For Each strLine in lstLines
			strNewLine = strLine
			if blnDataMappingSection Then
				'We are under datamappingsection
				intCommaPos = Instr(strLine,",")
				If IntCommaPos>0 Then
					strValue = Trim(Left(strLine,intCommaPos-1))
					'We check if the member is in the dictionary object
					If blnIsDefaultFlag Then
						If lstKeepDefaultValueFlag.Exists(strValue) Then
							'We split the line per properties
							lstParameters = Split(strLine,"|")
							'We need at least 4 items because we are going to update the third item
							intAux        	   = 0
							strFlagValue 	   = lstKeepDefaultValueFlag.Item(strValue)
							intUBound    	   = CInt(Trim(Split(strFlagValue,"-")(1)))
							strUpdateFlagValue = Trim(Split(strFlagValue,"-")(2))
							strFlagValue  	   = Trim(Split(strFlagValue,"-")(0))
							
							While Ubound(lstParameters)< intUBound
								Redim preserve lstParameters(ubound(lstParameters)+1)
							Wend
							'We update the parameter
							If strUpdateFlagValue <> "" Then
								lstParameters(intUBound - 1) = strFlagValue
								lstParameters(intUBound) = strUpdateFlagValue
							Else 
								lstParameters(intUBound) = strFlagValue
							End IF
							'We create the new line again
							strNewLine = Join(lstParameters,"|")
							If Right(strNewLine,1) = "|" Then
								strNewLine = Mid(strNewLine,1,len(strNewLine)-1)
							End If
						End if
					Else
						If lstFieldsOnlyUpdateUponInsert.Exists(strValue) Then
							'We split the line per properties
							lstParameters = Split(strLine,"|")
							'We need at least 4 items because we are going to update the third item
							intAux = 0
							While Ubound(lstParameters)<3
								Redim preserve lstParameters(ubound(lstParameters)+1)
							Wend
							'We update the parameter
							lstParameters(3) = lstFieldsOnlyUpdateUponInsert(strValue)
							'We create the new line again
							strNewLine = Join(lstParameters,"|")
						End if
					End if
				End if
			Else
				'We check if we are in the datamapping section
				blnDataMappingSection = Instr(UCase(strLine),":DATAMAPPING")>0
			End if
			'We create the new property info
			strUpdatedProperty = strUpdatedProperty & strNewLine & vbNewLine
		Next
	Else
		'we do not have any member with that property so we can leave it as it is
		strUpdatedProperty = strNewProperty
	End if
	'We update the strNewProperty with the strUpdatedProperty
	keepFlagUpdateOnlyUponInsert = strUpdatedProperty
End Function

'DN 2017/02/16. Avoid updating mandatory flag update
Function keepMFlagUpdateOnlyUponInsert(strCurrentProperty,strNewProperty)
	Dim strConfigPropLine, strSearchKeyWord, strReplaceKeyWord
	For each strConfigPropLine in Split(strCurrentProperty,vbNewLine)
		strSearchKeyWord  = ""
		strReplaceKeyWord = ""
		If UCase(Right(strConfigPropLine,2)) = "|M" Then
			arrConfigPropLine = Split(strConfigPropLine,",")
			strSearchKeyWord  = arrConfigPropLine(0) & ","
			If InStr(strNewProperty,strSearchKeyWord) <> 0 Then
				For Each strMatches in Split(strNewProperty,strSearchKeyWord)
					If InStr(arrConfigPropLine(1),Split(strMatches,vbNewLine)(0)) = 1 AND strReplaceKeyWord = "" Then
						strReplaceKeyWord = strSearchKeyWord & Split(strMatches,vbNewLine)(0)
					End If
				Next
			End If
			If strReplaceKeyWord <> "" Then
				strNewProperty = Replace(strNewProperty,strReplaceKeyWord,strConfigPropLine)
			End If
		End If
	Next
	keepMFlagUpdateOnlyUponInsert = strNewProperty
End Function

'This is a class to keep the necessary info to create the WF processes into a CRM system
'It is used to store the info of each system from config file
Class SystemSettings
	'Main vables. We dont use getters & setters to simplIfy the code
	Public FtpFolderPath 
	Public FtpErrorFolderPath
	Public CrmServerName
	Public CrmSystemName
	Public CrmUserName
	Public CrmUserPwd
	Public SystemEnvironment
	Public SystemMarket
	Public SystemImportFolder
	Public ObjectToFolderInfo
	'Initialize the vables when a new object is created
	Public Sub Class_Initialize()
        FtpFolderPath = ""
		FtpErrorFolderPath = ""
		CrmServerName = ""
		CrmSystemName = ""
         CrmUserName = ""
		CrmUserPwd = ""
		SystemEnvironment = ""
		SystemMarket = ""
		SystemImportFolder = ""
		ObjectToFolderInfo = ""
     End Sub
	 'We use this function to initialize an object with the info of the other object of the same class
	 Public Function Init(objGeneralSettings)
        FtpFolderPath = objGeneralSettings.FtpFolderPath
		FtpErrorFolderPath = objGeneralSettings.FtpErrorFolderPath
		CrmServerName = objGeneralSettings.CrmServerName
		CrmSystemName = objGeneralSettings.CrmSystemName
        CrmUserName = objGeneralSettings.CrmUserName
		CrmUserPwd = objGeneralSettings.CrmUserPwd
		SystemEnvironment = objGeneralSettings.SystemEnvironment
		SystemMarket = objGeneralSettings.SystemMarket
		SystemImportFolder = objGeneralSettings.SystemImportFolder
		ObjectToFolderInfo = objGeneralSettings.ObjectToFolderInfo
		Set Init = Me
     End Function

	 'This method read the config file and populate the vables of this class with the info in the file
	 Public Sub ReadInfoFromFile(objTextStream)
		strLine = ""
		 While Not objTextStream.AtEndOfStream
			strLine = Trim(objTextStream.ReadLine)
			If Len(strLine)>0 Then
				Select Case Left(strLine, 1)
					Case "#":
						'Do nothing comment line
					Case Else
						i = InStr(strLine, ",")
						If i > 0 Then
							Select Case  UCase(Trim(Left(strLine, i - 1)))
								Case "GENERAL_FTP_FOLDER_PATH":
									FtpFolderPath = Trim(Mid(strLine, i + 1))
								Case "GENERAL_FTP_ERROR_FOLDER_PATH":
									FtpErrorFolderPath = Trim(Mid(strLine, i + 1))
								Case "SERVER":
									CrmServerName = Trim(Mid(strLine, i + 1))
								Case "SYSTEM":
									CrmSystemName = Trim(Mid(strLine, i + 1))
								Case "USER":
									CrmUserName = Trim(Mid(strLine, i + 1))
								Case "PASSWORD":
									CrmUserPwd = Trim(Mid(strLine, i + 1))
								Case "ENVIRONMENT":
									SystemEnvironment = Trim(Mid(strLine, i + 1))
								Case "MARKET":
									SystemMarket = Trim(Mid(strLine, i + 1))
								Case "IMPORTFOLDER":
									SystemImportFolder = Trim(Mid(strLine, i + 1))
								Case "OBJECTTOFOLDERFILENAME":
									ObjectToFolderFilename = Trim(Mid(strLine, i + 1))
									'DNimbark 20160425 - Standard Edge Interface Release Process - Do not overwrite Move file workflow file name settings
									If objFSO.FileExists(strFolder & "imports\" & ObjectToFolderFilename) Then
										Set objFile = objFSO.GetFile(strFolder & "imports\" & ObjectToFolderFilename)
										ObjectToFolderInfo = getExtraFolderFileName(objFile.OpenAsTextStream(1,0).ReadAll)
									Else
										ObjectToFolderInfo = getExtraFolderFileName("")
									End If
							End Select
						End If
				End Select
			End If
		Wend
	 End Sub
	 
	 'Function to check If we have enough info to create a WF process with the info in this class
	 'We do not take into accout the user info
	 Public Function EnoughInfoToCreateWFNoUser()
		strResult = ""
		If Len(FtpFolderPath)=0  Then
			strResult = " FTP folder Path"
		End If
		If Len(FtpErrorFolderPath)=0  Then
			If Len(strResult) Then strResult = strResult & ","
			strResult = strResult & " FTP Error folder Path"
		End If
		If Len(CrmSystemName)=0  Then
			If Len(strResult) Then strResult = strResult & ","
			strResult = strResult & " CRM System"
		End If
		If Len(SystemEnvironment)=0 Then
			If Len(strResult) Then strResult = strResult & ","
			strResult = strResult & " Environment"
		End If
		If Len(SystemMarket)=0  Then
			If Len(strResult) Then strResult = strResult & ","
			strResult = strResult & " Market"
		End If
		If Len(SystemImportFolder)=0  Then
			If Len(strResult) Then strResult = strResult & ","
			strResult = strResult & " Import Folder"
		End If
		If Len(strResult) Then
			strResult = "No enter the following mandatory info: " & strResult
		End If
		EnoughInfoToCreateWFNoUser = strResult
	 End Function
	 'Function to check If we have enough info to create a WF process with the info in this class
	 Public Function EnoughInfoToCreateWF()
		strResult = EnoughInfoToCreateWFNoUser
		If Len(CrmUserName)=0 Then
			If Len(strResult)>0 Then 
				strResult = strResult & ", User"
			Else
				strResult = "No enter the following mandatory info: User"
			End If
		End If
		EnoughInfoToCreateWF = strResult
	 End Function
End Class

'DNimbark 12/10/2015 Function to update import workflows that are not included in release
'APN 12/13/2018 - added WFs to list of exceptions
Public function updateImportWFs()
	Dim objWF,intUpdateWFs
	
	Dim strTempWFXMLDef,strOldSchedule,strDefaultSchedule
	strTempWFXMLDef = ""
	intUpdateWFs = 0
	If Len(strDefaultWFXMLDef) > 0 Then
		strDefaultSchedule = "schedule=""" & Split(Split(strDefaultWFXMLDef,"schedule=""")(1),"""")(0) & """"
		
		For each objWF in System.Folders("WF_Processes").Scan("WF_ProcID @ 'IMP_' AND WF_ProcID <> 'IMP_ERR_FILES_TO_FTP' AND WF_ProcID <> 'IMP_FILES_FROM_FTP' AND WF_ProcID <> 'IMP_TERRTOUSERDEL' AND WF_ProcID <> 'IMP_SUC_FILES_TO_FTP' AND WF_ProcID <> 'IMP_DOCUMENTS' AND WF_ProcID <> 'IMP_PLANOGRAM' AND WF_ProcID <> 'IMP_KPIDATA'")
			'DNimbark 20160511 - Removed the condition for checking in the dictonary
			'DNimbark 20170629 - Do not update Costume Imports with std code
			If InStr(strLstCustomeImports,"," & objWF.WF_ProcID & ",") = 0 Then
				strTempWFXMLDef = strDefaultWFXMLDef
				objWF.NonSystem = WF_NONSYSTEM
				If not isnull(objWF.Defn_diff) Then
					System.ConsolidateDefn(objWF)
					objWF.save
				End IF			
				If UBound(Split(objWF.Defn.value,"schedule=""")) > 0 Then
					strOldSchedule ="schedule=""" &  Split(Split(objWF.Defn.value,"schedule=""")(1),"""")(0) & """"
					strTempWFXMLDef = Replace(strTempWFXMLDef,strDefaultSchedule,strOldSchedule)
					
					If not isnull(objWF.Defn_diff) Then
						System.ConsolidateDefn(objWF)
						objWF.save
					End IF
					'DNimbark 20160627 - Compare the WF code version #
					If checkVersion(objWF.Defn.Value,strTempWFXMLDef) Then
						objWF.Defn.Value = strTempWFXMLDef
					End If				
					objWF.save
					intUpdateWFs = intUpdateWFs + 1
				End If
			End If
		Next
	End If
	updateImportWFs = intUpdateWFs
End Function

'DNimbark 20160412 - We will not update the WF property PasswordEncryption, if exists.
Public Function keepPasswordEncryption(oldWFProperty,newWFProperty,strPropName)
	Dim oldEncryptionValue
	Dim newEncryptionValue
	
	oldEncryptionValue = strPropName & Split((Split(oldWFProperty,strPropName)(1)),vbNewLine)(0)
	newEncryptionValue = strPropName & Split((Split(newWFProperty,strPropName)(1)),vbNewLine)(0)
	
	keepPasswordEncryption = Replace(newWFProperty,newEncryptionValue,oldEncryptionValue)
End Function

'DNimbark 20160425 - Standard Edge Interface Release Process - Do not overwrite Move file workflow file name settings
Public Function getExtraFolderFileName(strProperty)
	Dim strOldWFProperty 
	Dim strTempProperties
	Dim arrprops
	strTempProperties = ""
	Set objWF = System.Folders.WF_Processes.Scan("UCase(Title)='" & UCase("Import - MOVE FILES FROM FTP FOLDER") & "'").Fetch
	If Not objWF Is Nothing Then
		Set objWFProp = objWF.Folders("WorkflowProperties").First("UCase(wfProperty)='Folder To Object'")
		If Not objWFProp Is Nothing Then
			strOldWFProperty = objWFProp.wfValue
		End If
	End if
'msgbox "strOldWFProperty: " & strOldWFProperty
'msgbox "splitting newline: " & ubound(Split(strOldWFProperty,vbNewLine))

	If strOldWFProperty <> "" Then
		For Each arrProps in Split(strOldWFProperty,vbNewLine)
			'msgbox "splitting arrProps: " & ubound(Split(arrProps,","))
			if uBound(Split(arrProps,","))>=0 then
			If Split(arrProps,",")(0) >= 900 Then
				If strTempProperties <> "" Then
					strTempProperties = strTempProperties & vbNewLine & arrProps
				Else
					strTempProperties = arrProps
				End If
			End If
			End If
		Next
	End If
	getExtraFolderFileName = strProperty & strTempProperties
End Function

'Jhirpara 20150223 Defect #1324 Won't allow just an LF as end-of-line where it was accepting it in the past with v 1.0.13 of the std import.
'Function to get Line Separator

Function GetLineSeparator(strFileName)            
 
	Dim objStream, strData, intLineSeparator, intInfoFirstLine
	intLineSeparator = ""
	intInfoFirstLine = 10000
   
	For Each intOption In Array(-1,10,13)
   
		Set objStream = CreateObject("ADODB.Stream")
		objStream.CharSet = "utf-8"
		objStream.LineSeparator = intOption
		objStream.Open
		objStream.LoadFromFile(strFileName)
		strData = objStream.ReadText(-2)
		If Len(strData) < intInfoFirstLine Then
			intInfoFirstLine = Len(strData)
			intLineSeparator = intOption
		End If
		objStream.Close
	Next
   
	GetLineSeparator = intLineSeparator
               
End Function

'DNimbark 20160627 - Function to compare the version of the WF definition and the version of the release, 
'return False if WF definition is higher than the release's version 
'else return true
Public Function checkVersion(strWFDefn,strRelDefn)
	strWFVersion  = ""
	strRelVersion = ""
	
	If InStr(UCase(strWFdefn),vbNewLine & "VERSION") > 0 Then
		strWFVersion = Split(Split(UCase(strWFdefn),vbNewLine & "VERSION")(1),"""")(1)
	End If
	
	If InStr(UCase(strRelDefn),vbNewLine & "VERSION") > 0 Then
		strRelVersion = Split(Split(UCase(strRelDefn),vbNewLine & "VERSION")(1),"""")(1)
	End If
	
	If Len(strWFVersion) > 0 AND Len(strRelVersion) > 0 Then
		blnResult = compareFloat(strWFVersion,strRelVersion)
	ElseIf Len(strWFVersion) = 0 Then
		blnResult = True
	ElseIf Len(strRelVersion) = 0 Then
		blnResult = False
	End If
	checkVersion = blnResult
End Function

'DNimbark 20160627 - Function to compare WF version
Function compareFloat(fltWFVersion,fltReleaseVersion)
	aryWFversion = Split(fltWFVersion,".")
	aryReleaseVersion = Split(fltReleaseVersion,".")
	compareFloat = True
	for i=0 to UBound(aryWFversion)
		If UBound(aryReleaseVersion) > UBound(aryReleaseVersion) AND i > UBound(aryReleaseVersion) Then
			compareFloat = False
			Exit For
		End IF
		If aryWFversion(i) > aryReleaseVersion(i) Then
			compareFloat = False
			Exit For
		ElseIf aryWFversion(i) < aryReleaseVersion(i) Then
			compareFloat = True
			Exit For
		End If
	Next
End Function

'DNimbark 20160701 - Function to check the expression is valid or not
Function isExpressionValid(strCurrentProperty,arrPipeValues)
	strClass  = ""
	blnStatus = True
	strFilter = Trim(Split(arrPipeValues(0),",")(1)) & "." & arrPipeValues(1)
	On Error Resume Next
		strClass      = Trim(Split(Split(Split(strCurrentProperty,"#Import Class Configuration")(1),vbNewLine)(1),",")(1))
		Set expFilter = CreateObject("ActivElk.Filter")
		If Not expFilter.Parse(strFilter, System.Classes(strClass), strFilterErr) Then
			blnStatus = False
		End If
	On Error GoTo 0
	
	isExpressionValid = True
	If strClass = "" Then
		isExpressionValid = False
	Else
		If Not blnStatus Then
			isExpressionValid = False
		End If
	End If
End Function


Sub addJS(strFilePathAndNameJS, strDescription, strEventType, strExternalID, objLogFile)
	strEventType      = LTrim(RTrim(strEventType))
	strDescription    = LTrim(RTrim(strDescription))
	objLogFile.WriteLine("Maintaining " & LTrim(RTrim(strEventType)) & " TouchEvent record (" & strDescription & ").")

	Set objFSO        = CreateObject("Scripting.FileSystemObject")
	Set objFile       = objFSO.GetFile(strFilePathAndNameJS)
	Set objTextStream = objFile.OpenAsTextStream(1,0)
	strContent        = objTextStream.ReadAll
	If blnDebug Then MsgBox strContent End If

	Set objTouchEvent = Nothing
	On Error Resume Next
	If blnDebug Then 
		MsgBox "Description.primary ='" & strDescription & "'"
		MsgBox "Type ='" & strEventType & "'"
	End If

	If (strEventType <> "OE" and strEventType <> "DEX" and strEventType <> "CX" and strEventType <> "VISIT") then
		Set objTouchEvent = System.Folders("TCG_TouchEvent").Scan("Description.primary ='" & strDescription & "'").Fetch
		strContent = doNotUpdatedPromptMSG(strContent,objTouchEvent.JavascriptBlock)
	Else
		Set objTouchEvent = System.Folders("TCG_TouchEventFunction").Scan("Description.primary ='" & strDescription & "'").Fetch
	End If
	On Error goto 0

	If Not objTouchEvent Is Nothing Then
		Set objTransaction = System.BeginTransaction

		objLogFile.WriteLine(" Found existing " & LTrim(RTrim(strEventType)) & " TouchEvent record titled " & strDescription & ".")
		objLogFile.WriteLine(String(100,"-"))
		objLogFile.WriteLine(" Old JavaScript content")
		objLogFile.WriteLine(String(100,"-"))
		objLogFile.WriteLine(objTouchEvent.JavascriptBlock.Value)
		objLogFile.WriteLine(String(100,"-"))

		'DNimbark 20160916 - Need to accept a new naming convention for Touch Event loading (Touch Event Description~Touch Event Type~Touch Event External ID).  The "key" of the Touch Event remains Description.
		If strExternalID <> "" AND Not IsEmpty(strExternalID) Then
			objTouchEvent.EX_External.Id = strExternalID
		End IF
		objTouchEvent.JavascriptBlock = strContent
		objTouchEvent.EventType.Value = LTrim(RTrim(strEventType))
		objTouchEvent.Save objTransaction

		objLogFile.WriteLine(" New JavaScript content")
		objLogFile.WriteLine(String(100,"-"))
		objLogFile.WriteLine(objTouchEvent.JavascriptBlock.Value)
		objLogFile.WriteLine(String(100,"-"))

		Set objVal = objTransaction.Validate

		If objVal.status <> 3 Then
			objTransaction.Commit
			If blnDebug Then MsgBox "objTransaction.Commit" End If
			objLogFile.WriteLine("Saved TouchEvent record.")
		Else
			strErrMessage = objTransaction.Result & " (" & objTransaction.Status  & ")"
			MsgBox "Error on Save: " & strErrMessage,16,"" & strProcessNameForLog & " Error"
			objLogFile.WriteLine("Error on Save: " & strErrMessage & ".")
		End If

		If blnDebug Then MsgBox "after update - " & objTouchEvent.JavascriptBlock.Value End If
	Else
		'Add code to create a touch event
		If blnDebug Then MsgBox "Touch Event not found, creating one..." End If
		objLogFile.WriteLine(" Creating new " & LTrim(RTrim(strEventType)) & " TouchEvent record titled " & strDescription & ".")
		objLogFile.WriteLine(" JAVASCRIPT CONTENT")
		objLogFile.WriteLine(String(100," -"))
		objLogFile.WriteLine(strContent)
		objLogFile.WriteLine(String(100," -"))

		Set objTransaction              = System.BeginTransaction
	
		If (strEventType <> "OE" and strEventType <> "DEX" and strEventType <> "CX" and strEventType <> "VISIT") Then
			Set objInstance = System.Folders.TCG_TouchEvent.CreateNewInstance()
			objInstance.EventType.Value = LTrim(RTrim(strEventType))
		Else
			Set objInstance = System.Folders.TCG_TouchEventFunction.CreateNewInstance()
			objInstance.FunctionType.Value = "OE"
			objInstance.FunctionName = strDescription
		End If
		objInstance.Description.Primary = strDescription
		'DNimbark 20160916 - Need to accept a new naming convention for Touch Event loading (Touch Event Description~Touch Event Type~Touch Event External ID).  The "key" of the Touch Event remains Description.
		If strExternalID <> "" AND Not IsEmpty(strExternalID) Then
			objInstance.EX_External.Id	= strExternalID
		End If
		objInstance.JavascriptBlock     = strContent
		objInstance.Save objTransaction

		Set objVal = objTransaction.Validate
		
		If objVal.status <> 3 Then
			objTransaction.Commit
			If blnDebug Then MsgBox "objTransaction.Commit" End If
			objLogFile.WriteLine("Saved TouchEvent record.")
			intAddedTouchEvent   = intAddedTouchEvent + 1 
		Else
			strErrMessage = objTransaction.Result & " (" & objTransaction.Status  & ")"
			MsgBox "Error on Save: " & strErrMessage,16,"" & strProcessNameForLog & " Error"
			objLogFile.WriteLine("Error on Save: " & strErrMessage & ".")
			intErroredTouchEvent = intErroredTouchEvent + 1
		End If

	End If

End Sub

'DNimbark 20170120 - Function to copy XSD file to the path provided in WF property
Function CopyXSDFile(objInstance,strCFGContent,objWFFDirectory,strXSDPropName,strContent)
	On Error Resume Next
		strXSDFullPath = Split(Split(strCFGContent,strXSDPropName)(1),vbNewLine)(0)
		If strXSDPropName = "SchemaPath:~" Then
			strXSDFullPath = AddEndSlashIfNecessary(strXSDFullPath) & "AdditionalLoadExport.xsd"
		End If
		strXSDFileName = Split(strXSDFullPath,"\")(UBound(Split(strXSDFullPath,"\")))
		strXSDPath     = Replace(strXSDFullPath,strXSDFileName,"")
		
		For Each objLine in Split(strCFGContent,vbNewLine)
			strWFProp = Split(objLine,":~")
			strPropName = strWFProp(0)
			strPropValue = strWFProp(1)
			
			If strPropName <> "" Then
				updateCreateWFProperty objInstance, strPropName, strPropValue
			End If
		next
	
		Set objFSO 	   = CreateObject("Scripting.FileSystemObject")
		set objXSDFile = objFSO.GetFile(AddEndSlashIfNecessary(objWFFDirectory.Path) & strXSDFileName)
		objFSO.CopyFile objXSDFile, strXSDPath
		strTemp = strXSDPath & strXSDFileName
		If strXSDPropName = "SchemaPath:~" Then
			oldXSDPath = Trim(Split(Split(strContent,"&lt;Schema&gt;")(1),"&lt;/Schema&gt;")(0))
			strContent = Replace(strContent,oldXSDPath,strTemp)
	End If
		
		CopyXSDFile = strContent
	On Error Goto 0
End Function

Sub addRoles(strFilePathAndNameCSV, objLogFile)
	If blnDebug Then MsgBox "Enter addRoles, csv is " & strFilePathAndNameCSV End If

	'Find .csv file
	Set objFSO        = CreateObject("Scripting.FileSystemObject")
	Set objFile       = objFSO.GetFile(strFilePathAndNameCSV)
	Set objTextStream = objFile.OpenAsTextStream(1,0)
	'arrHeader         = Split(objTextStream.ReadLine,",") 'There is no header!

	Do Until objTextStream.AtEndOfStream
		'Array that data needs to be read into
		arrTextStreamLines = Split(objTextStream.ReadLine,",")
		Set objRole        = Nothing
		Set objRoleEvent   = Nothing
		Set objRole        = System.Folders.Roles.First("Role.Code='" & arrTextStreamLines(0) & "'")
		If blnDebug Then MsgBox "System.Folders.Roles.Scan(Role.Code='" & arrTextStreamLines(0) & "').Fetch" End If

		objLogFile.WriteLine("Adding RoleTouchEvent join record. Role is " & arrTextStreamLines(0) & ", TouchEvent is " & arrTextStreamLines(1) & ".")

		
		strFilter_TCG_RoleTouchEvents = "Role.Code='" & arrTextStreamLines(0) & "' And TCG_TouchEvent.Description.Primary = '" & arrTextStreamLines(1) & "'"
		objLogFile.WriteLine(" Scanning for RoleTouchEvent record using filter " & strFilter_TCG_RoleTouchEvents & ".")

		If Not System.Folders.TCG_RoleTouchEvents.Scan(strFilter_TCG_RoleTouchEvents).EndOfScan Then
			Set objRoleEvent = System.Folders.TCG_RoleTouchEvents.First(strFilter_TCG_RoleTouchEvents)
		End If

		If Not objRole Is Nothing AND Not objRole.IsNull Then
			Set objTransaction = System.BeginTransaction

			If Not objRoleEvent Is Nothing Then
				If blnDebug Then MsgBox "objRoleEvent already exists!" End If
				objLogFile.WriteLine(" RoleTouchEvent record already exists.")

				With objRoleEvent
					If blnDebug Then MsgBox ".Selected = " & arrTextStreamLines(2) End If
					objLogFile.WriteLine(" Updating RoleTouchEvent record to " & arrTextStreamLines(2) & ".")
					.Selected = arrTextStreamLines(2)
					.Save objTransaction
				End With

			Else
				If blnDebug Then MsgBox "System.Folders.TCG_TouchEvent.First(Description.Primary = '" & arrTextStreamLines(1) & "')" End If
				Set objTouchEvent               = System.Folders.TCG_TouchEvent.First("Description.Primary = '" & arrTextStreamLines(1) & "'")
				Set objRoleEvent                = objRole.Folders.TCG_TouchEvents.CreateNewInstance()
				Set objRoleEvent.TCG_TouchEvent = objTouchEvent

				If Not objRoleEvent Is Nothing Then

					With objRoleEvent
						If blnDebug Then MsgBox ".Selected = " & arrTextStreamLines(2) End If
						objLogFile.WriteLine(" Setting RoleTouchEvent record to " & arrTextStreamLines(2) & ".")
						.Selected = arrTextStreamLines(2)
						.Save objTransaction
					End With

				End If

			End If

			Set objVal = objTransaction.Validate

			If objVal.Status <> 3 Then
				objTransaction.Commit
				If blnDebug Then MsgBox "objTransaction.Commit" End If
				objLogFile.WriteLine("Saved RoleTouchEvent record.")
			Else
				strErrMessage = objTransaction.Result & " (" & objTransaction.Status  & ")"
				MsgBox "Error on Save: " & strErrMessage,16,"" & strProcessNameForLog & " Error"
				objLogFile.WriteLine("Error on Save: " & strErrMessage & ".")
			End If

		Else
			MsgBox "Role " & arrTextStreamLines(0) & " was not found."
		End If

	Loop

End Sub

'SGhunchala 20200529 - Automation Release Script Update
Function UpdateWorkflowMembers(System)
	set fso = CreateObject("Scripting.FileSystemObject")
	strWorkflowFile = "workflows - " & System.Name & ".txt"
	strDirectory = strClusterPath & "\" & strWorkflowFile
	If fso.FileExists(strDirectory) Then
		objSystemLog.LogNormalDetail "Started execution to update the workflow members."
		Set objCodeFile = fso.OpenTextFile(strDirectory)
		If not objCodeFile.AtEndOfStream Then
			Do Until objCodeFile.AtEndOfStream
				strCode = objCodeFile.ReadLine
				If Len(strCode) > 0 Then
					If Left(strCode,1) <> "#" Then
						strSplitLine = Split(strCode, "|")
						strFolderScanFilter = strSplitLine(0)
						strFields = strSplitLine(1)
						If strFolderScanFilter <> "" and strFields <> ""  Then
						Else
							objSystemLog.LogErrorDetail "Invalid line : " & strCode
						End If
						If System.Folders.WF_Processes.Exists(strFolderScanFilter) then
							For each objWF in System.Folders.WF_Processes.scan(strFolderScanFilter)
								strWFPrompt = objWF.Prompt
								Set objTransaction = System.BeginTransaction
								For Each strFieldSet in Split(strFields,",")
									execute "objWF." & Split(strFieldSet,"~")(0) & "=" & Split(strFieldSet,"~")(1)
								Next
								objWF.save objTransaction
								set objval = objTransaction.validate		
								if objval.status <> 3 then
									objTransaction.Commit
									objSystemLog.LogNormalDetail "The '"& strWFPrompt &"' workflow has been updated successfully."
								else
									objSystemLog.LogErrorDetail "","","There were some error in updating the '"& strWFPrompt &"' workflow members : '" & objval.Result.Message & "'" 
								end if
								SET_NOTHING Array(objTransaction,objval)
							Next
						Else 
							objSystemLog.LogNormalDetail "The workflow does not exists : " & strFolderScanFilter
						End If
					End If
				End If
			Loop
		End If
		objSystemLog.LogNormalDetail "Completed execution to update the workflow members."
	End If
End Function

'DB_PART_TITLE~(Campaign|CX_Route|CX_RouteGroup|at_General|CI_Issue|BannerGroup|DC_Group|Store|LocationVisit|SurveyActivity|AS_Territory|PromoNormal|stdUser])~(FULL|JAVASCRIPT|HTML)~(HTML,KPI,ANALYTICS,TILEGROUP)~KEY
'Last part is optional
function loadTouchDBParts(System,strDirectoryPath)

	Dim objFSO                 : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objFolder              : Set objFolder = objFSO.GetFolder(strDirectoryPath)
	objSystemLog.LogNormalDetail "The touch dashboard part loading process started."
	strDirectoryPath = AddEndSlashIfNecessary(strDirectoryPath)
	
	For Each objFile In objFolder.Files
		blnProcess = True
		strFileName = objFile.Name
		arrTokens   = Split(strFileName,".")

		If (UCase(arrTokens(1)) = "TXT") Then
			intTotalTouchDBParts = intTotalTouchDBParts + 1
			strCodeFilePath = objFile.Path			
			If objFSO.FileExists(strCodeFilePath) Then
				Set objCodeFile = objFSO.OpenTextFile(strCodeFilePath, 1)
				If objCodeFile.AtEndOfStream Then
					objSystemLog.LogNormalDetail "File is empty : " & objFile.Name
					intErroredTouchDBParts = intErroredTouchDBParts + 1
				Else
					'strCode = Replace(Replace(Replace(objCodeFile.ReadAll,"&","&amp;"),"<","&lt;"),">","&gt;")
					strCode = objCodeFile.ReadAll
					If Len(Trim(strCode) & "") > 0 Then
						arrcfg = parseConfigs(Array("members","permitgroup","keymapping"),strCode)
						arrFileParts = Split(arrTokens(0), "~")
						if UBound(arrFileParts)>0 Then
							strFilePathAndNameJS = strDirectoryPath + "\" + objFile.Name
							
							strDBPartKey = arrFileParts(0)  'Dashboard part Key
							strDBPartCode = arrFileParts(1)  'Code definition part
							'strDBPartClass = arrFileParts(1) 'Class Name
							'strDBPartType = ""
							'strNewKey = ""
							
							If UCase(strDBPartCode)="FULL" or UCase(strDBPartCode)="HTML" or UCase(strDBPartCode)="JAVASCRIPT" Then
								strFilter = "Key='"& strDBPartKey &"'"
								If arrcfg(2).Count>0 Then
									strFilter = arrcfg(2).Item("Expression")
								End If
								If System.Folders.TDB_DashboardPart.Exists(strFilter) Then
									set objDBPart = System.Folders.TDB_DashboardPart.First(strFilter)
								Else
									If UCase(strDBPartCode) = "FULL" Then
										set objDBPart = System.Folders.TDB_DashboardPart.CreateNewInstance(,,strDBPartKey)
									Else
										blnProcess = False
										objSystemLog.LogNormalDetail "Touch Dashboard Part Not Exists in system. Wrong File Name Format to create new dashboard part : " & strFileName
									End IF
								End If
								IF blnProcess Then
									If Not objDBPart.IsNull Then
										set objTransaction = System.BeginTransaction
										strDefinition = ""
										If objFSO.FileExists(strDirectoryPath & "DEFINITION\" & strFileName) Then
											set objFileToRead = objFSO.OpenTextFile(strDirectoryPath & "DEFINITION\" & strFileName,1)
											strFileText = objFileToRead.ReadAll()
											if Len(strFileText & "")>0 Then
												strDefinition = Trim(strFileText)
											End If
										End If
										If strDefinition <> "" then
											If UCase(strDBPartCode) = "FULL" Then
													objDBPart.Definition.Value = strDefinition
											Else
												If UCase(strDBPartCode) = "JAVASCRIPT" Then
													strStartString = "<JSScript>"
													strEndString = "</JSScript>"
												ElseIf UCase(strDBPartCode) = "HTML" Then
													strStartString = "<HTML>"
													strEndString = "</HTML>"
												End If
												strDefinition = Replace(Replace(Replace(strDefinition,"&","&amp;"),"<","&lt;"),">","&gt;")
												strOldDefinition = objDBPart.Definition.Value
												intCodeStartPoint = InStr(strOldDefinition,strStartString)
												intCodeEndPoint = InStr(strOldDefinition,strEndString)
												intCodeLength = intCodeEndPoint - intCodeStartPoint
												
												If intCodeStartPoint = 0 OR intCodeEndPoint = 0 Then
													strOldCode = ""
												Else
													strOldCode = Replace(Mid(strOldDefinition,intCodeStartPoint,intCodeLength),strStartString,"")
												End If
												strDefinition = Replace(strOldDefinition,strOldCode,strDefinition)
												objDBPart.Definition.Value = strDefinition
											End If
										End If
										
										For Each strKey in arrcfg(0).Keys
											if UCase(strKey) = "DASHBOARDPARTTYPE" Then
												objDBPart.DashboardPartType = System.Folders.TDB_DashboardPartTypes.first("key='"& arrcfg(0).Item(strKey) &"'")
											Else
												Execute "objDBPart." & strKey & " = """ & arrcfg(0).Item(strKey) & """"
											End IF
										Next
										
										'Permit groups
										AddUserGroups objDBPart,arrcfg(1),objTransaction
										
										objDBPart.save objTransaction
										set objval = objTransaction.validate		
										if objval.status <> 3 then
											intAddedTouchDBParts = intAddedTouchDBParts + 1
											objTransaction.Commit
											objSystemLog.LogNormalDetail "The '"& strDBPartTitle &"' touch dashboard part has been created/updated."
										else
											intErroredTouchDBParts = intErroredTouchDBParts + 1
											objSystemLog.LogErrorDetail "","","There were some error in updating/creating the '"& strDBPartTitle &"' touch dashboard part : '" & objval.Result.Message & "'" 
										end if
									Else
										intErroredTouchDBParts = intErroredTouchDBParts + 1
										objSystemLog.LogErrorDetail "Something Wrong : " & strFileName
									End If
								Else
									intErroredTouchDBParts = intErroredTouchDBParts + 1
								End IF
							Else
								intErroredTouchDBParts = intErroredTouchDBParts + 1
								objSystemLog.LogErrorDetail "Wrong File Name Format : " & strFileName
							End If
						Else
							intErroredTouchDBParts = intErroredTouchDBParts + 1
							objSystemLog.LogErrorDetail "Wrong File Name Format : " & strFileName
						End IF			
					Else
						objSystemLog.LogNormalDetail "File is empty : " & objFile.Name
						intErroredTouchDBParts = intErroredTouchDBParts + 1
					End If
				End If
			End If
		End If
	Next
	objSystemLog.LogNormalDetail "The touch dashboard part loading process completed."
	SET_NOTHING Array(objFSO,objFolder,objTransaction,objVal,objDBPart)
End function

Function SET_NOTHING(arrOfObject)
    For Each obj In arrOfObject
        set obj = Nothing
    Next
End function

function loadTouchDBTemplates(System,strDirectoryPath)
	Dim objFSO                 : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objFolder              : Set objFolder = objFSO.GetFolder(strDirectoryPath)
	objSystemLog.LogNormalDetail "The touch dashboard template loading process started."
	strDirectoryPath = AddEndSlashIfNecessary(strDirectoryPath)
	
	
	For Each objFile In objFolder.Files
		blnProcess = True
		strFileName = objFile.Name
		arrTokens   = Split(strFileName,".")

		If (UCase(arrTokens(1)) = "TXT") Then
			intTotalTouchDBTemplates = intTotalTouchDBTemplates + 1
			strCodeFilePath = objFile.Path			
			If objFSO.FileExists(strCodeFilePath) Then
				Set objCodeFile = objFSO.OpenTextFile(strCodeFilePath, 1)
				If objCodeFile.AtEndOfStream Then
					objSystemLog.LogNormalDetail "File is empty : " & objFile.Name
					intErroredTouchDBTemplates = intErroredTouchDBTemplates + 1
				Else
					strCode = objCodeFile.ReadAll
					If Len(Trim(strCode) & "") > 0 Then
						arrcfg = parseConfigs(Array("members","keymapping"),strCode)
						strDBTemplateKey = arrTokens(0)
						strFilter = "Key='"& strDBTemplateKey &"'"
						
						if arrcfg(1).Count>0 Then
							strFilter = arrcfg(1).Item("Expression")
						End If
						
						set objDBTemplate = System.Folders.TDB_Template.first(strFilter)
						If objDBTemplate.IsNull Then
							set objDBTemplate = System.Folders.TDB_Template.CreateNewInstance(,,strDBTemplateKey)
						End If
						
						If Not objDBTemplate.IsNull then
							set objTransaction = System.BeginTransaction
						
							If objFSO.FileExists(strDirectoryPath & "TEMPLATE\" & strFileName) Then
								set objFileToRead = objFSO.OpenTextFile(strDirectoryPath & "TEMPLATE\" & strFileName,1)
								strFileText = objFileToRead.ReadAll()
								if Len(strFileText & "")>0 Then
									objDBTemplate.TemplateHTML = Trim(strFileText)
								End If
							End If
							
							For each strKey in arrcfg(0).Keys
								Execute "objDBTemplate." & strKey & " = """ & arrcfg(0).Item(strKey) & """"
							Next
							
							objDBTemplate.save objTransaction
							set objval = objTransaction.validate		
							if objval.status <> 3 then
								intAddedTouchDBTemplates = intAddedTouchDBTemplates + 1
								objTransaction.Commit
								objSystemLog.LogNormalDetail "The '"& strDBTemplateKey &"' touch dashboard template has been created/updated."
							else
								intErroredTouchDBTemplates = intErroredTouchDBTemplates + 1
								objSystemLog.LogErrorDetail "","","There were some error in updating/creating the '"& strDBTemplateKey &"' touch dashboard template : '" & objval.Result.Message & "'" 
							end if
						Else
							intErroredTouchDBTemplates = intErroredTouchDBTemplates + 1
							objSystemLog.LogErrorDetail "Something Wrong : " & strDBTemplateKey
						End If
						
					Else
						intErroredTouchDBTemplates = intErroredTouchDBTemplates + 1
						objSystemLog.LogNormalDetail "File is empty : " & objFile.Name
					End If
				End If
			Else
				intErroredTouchDBTemplates = intErroredTouchDBTemplates + 1
				objSystemLog.LogErrorDetail "File doesn't exists : " & strCodeFilePath
			End If
		End IF
	Next
	
	objSystemLog.LogNormalDetail "The touch dashboard template loading process completed."
End function

Function loadTouchDBTemplatePos(System,strDirectoryPath)
	Dim objFSO                 : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objFolder              : Set objFolder = objFSO.GetFolder(strDirectoryPath)
	objSystemLog.LogNormalDetail "The touch dashboard template Position loading process started."
	
	
	For Each objFile In objFolder.Files
		blnProcess = True
		strFileName = objFile.Name
		arrTokens   = Split(strFileName,".")

		If (UCase(arrTokens(1)) = "TXT") Then
			intTotalTouchDBTemplatePos = intTotalTouchDBTemplatePos + 1
			strCodeFilePath = objFile.Path			
			If objFSO.FileExists(strCodeFilePath) Then
				Set objCodeFile = objFSO.OpenTextFile(strCodeFilePath, 1)
				If objCodeFile.AtEndOfStream Then
					objSystemLog.LogNormalDetail "File is empty : " & objFile.Name
					intErroredTouchDBTemplatePos = intErroredTouchDBTemplatePos + 1
				Else
					strCode = objCodeFile.ReadAll
					arrcfg = parseConfigs(Array("keymapping","positions"),strCode)
					strFilter = "Key='"& arrTokens(0) &"'"
					if arrcfg(0).Count>0 Then
						strFilter = arrcfg(0).Item("Expression")
					End If
					If Len(Trim(strCode) & "") > 0 Then
						set objDBTemplate = System.Folders.TDB_Template.first(strFilter)
						If Not objDBTemplate.IsNull Then
							set objTransaction = System.BeginTransaction
							arrLines = Split(strCode,vbCrlf)
							For Each strLine in arrcfg(1).Keys
								arrParts = Split(Trim(strLine),"|")
								If UBound(arrParts) > 1 Then
									strTag = arrParts(0)
									strTitle = arrParts(1)
									intListOrder = CInt(arrParts(2))
									
									If UBound(arrParts) > 2 Then
										strNewKey = arrParts(3)
									Else
										strNewKey = ""
									End If
									
									set objDBTemplatePos  = objDBTemplate.folders.TDB_TemplatePosns.first("Tag='"& strTag &"'")
									If objDBTemplatePos.IsNull Then
										if strNewKey = "" Then
											set objDBTemplatePos  = objDBTemplate.folders.TDB_TemplatePosns.CreateNewInstance
										Else
											set objDBTemplatePos  = objDBTemplate.folders.TDB_TemplatePosns.CreateNewInstance(,,strNewKey)
										End If
										objDBTemplatePos.Members("Tag").Value = strTag
									End If
									If Not objDBTemplatePos.IsNull Then
										objDBTemplatePos.Title = strTitle
										objDBTemplatePos.ListOrder = intListOrder
										objDBTemplatePos.save objTransaction
									Else
										objSystemLog.LogErrorDetail "Something Wrong : Template - '"& arrTokens(0) &"', Position - '"& strTag &"'"
									End If
								End If
							Next
							set objval = objTransaction.validate		
							if objval.status <> 3 then
								intAddedTouchDBTemplatePos = intAddedTouchDBTemplatePos + 1
								objTransaction.Commit
								objSystemLog.LogNormalDetail "The '"& arrTokens(0) &"' touch dashboard template's position(s) have been created/updated."
							else
								intErroredTouchDBTemplatePos = intErroredTouchDBTemplatePos + 1
								objSystemLog.LogErrorDetail "","","There were some error in updating/creating the '"& arrTokens(0) &"' touch dashboard template's Position(s) : '" & objval.Result.Message & "'" 
							end if
						Else
							objSystemLog.LogErrorDetail "Touch Dashboard Template doesn't exist in system : " & arrTokens(0)
							intErroredTouchDBTemplatePos = intErroredTouchDBTemplatePos + 1
						End If
					Else
						objSystemLog.LogErrorDetail "File is empty : " & objFile.Name
						intErroredTouchDBTemplatePos = intErroredTouchDBTemplatePos + 1
					End If
				End If
			Else
				intErroredTouchDBTemplatePos = intErroredTouchDBTemplatePos + 1
				objSystemLog.LogErrorDetail "File doesn't exists : " & strCodeFilePath
			End If
		End If
	Next
						
	objSystemLog.LogNormalDetail "The touch dashboard template Position loading process Completed."
End Function

function loadTouchDB(System,strDirectoryPath)
	Dim objFSO                 : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objFolder              : Set objFolder = objFSO.GetFolder(strDirectoryPath)
	objSystemLog.LogNormalDetail "The touch dashboard loading process started."
	strDirectoryPath = AddEndSlashIfNecessary(strDirectoryPath)
	
	For Each objFile In objFolder.Files
		blnProcess = True
		strFileName = objFile.Name
		arrTokens   = Split(strFileName,".")
		strDBKey = arrTokens(0)

		If (UCase(arrTokens(1)) = "TXT") Then
			intTotalTouchDB = intTotalTouchDB + 1
			strCodeFilePath = objFile.Path			
			If objFSO.FileExists(strCodeFilePath) Then
				Set objCodeFile = objFSO.OpenTextFile(strCodeFilePath, 1)
				If objCodeFile.AtEndOfStream Then
					objSystemLog.LogNormalDetail "File is empty : " & objFile.Name
					intErroredTouchDB = intErroredTouchDB + 1
				Else
					strCode = objCodeFile.ReadAll
					If Len(Trim(strCode) & "") > 0 Then
						arrcfg = parseConfigs(Array("dashboardmembers","positions","permitgroup","keymapping"),strCode)
						strFilter = "Key='"& strDBKey &"'"
						If arrcfg(3).Count > 0  Then
							strFilter = arrcfg(3).Item("Expression")
						End If
						blnProcess = True
						set objTransaction = System.BeginTransaction
						If system.Folders.TDB_Dashboards.Exists(strFilter) Then
							set objDB = system.Folders.TDB_Dashboards.first(strFilter)
						Else
							set objDB = system.Folders.TDB_Dashboards.CreateNewInstance(,,strDBKey)
						End If
						If objDB.IsNew Then
							if arrcfg(0).Count > 1 Then
								if arrcfg(0).Exists("TDB_Template") and arrcfg(0).Exists("ClassNameEnum") Then
									strTemplateName = arrcfg(0).Item("TDB_Template")
									If System.folders.TDB_Template.Exists("Title='"& strTemplateName &"'") Then
										set objDB.TDB_Template = System.folders.TDB_Template.first("Title='"& strTemplateName &"'")
										objDB.ClassNameEnum.value = arrcfg(0).Item("ClassNameEnum")
										blnProcess = True
									Else
										blnProcess = False
										objSystemLog.LogNormalDetail "The '"& strTemplateName &"' touch dashboard template does not exists in system."
									End If
								Else
									blnProcess = False
									objSystemLog.LogNormalDetail "The '"& strDBKey &"' touch dashboard doesn't exist in system and config file to create new touch dashboard doesn't contain sufficient information."
								End If
							Else
								blnProcess = False
								objSystemLog.LogNormalDetail "The '"& strDBKey &"' touch dashboard doesn't exist in system and config file to create new touch dashboard doesn't contain sufficient information."
							End If
						End If
						
						if blnProcess Then
							'Dashboard level members
							For Each strKey in arrcfg(0).Keys
								If UCase(strKey) <> "TDB_TEMPLATE" and UCase(strKey) <> "KEY" and UCase(strKey) <> "CLASSNAMEENUM" Then
									If UCase(strKey) = "STARTFUNCTIONNAME" Then
										If objFSO.FileExists(strDirectoryPath & "JAVASCRIPT\" & strFileName) Then
											set objFileToRead = objFSO.OpenTextFile(strDirectoryPath & "JAVASCRIPT\" & strFileName,1)
											strFileText = objFileToRead.ReadAll()
											if Len(strFileText & "")>0 Then
												execute "objDB." & strKey & " = """ &  arrcfg(0).Item(strKey) & """"
												objDB.JavaScriptBlock = Trim(strFileText)
											End If
										End If
									Else
										execute "objDB." & strKey & " = """ &  arrcfg(0).Item(strKey) & """"
									End If
								End If
							Next
							
							'Position - Dashboard part mapping
							For Each strKey in arrcfg(1).Keys
								strTag = strKey
								arrParts = Split(arrcfg(1).Item(strKey),"|")
								set objDBPartPos = Nothing
								If System.folders.TDB_DashboardPart.Exists("Title.Primary='"& arrParts(0) &"'") Then
									If objDB.Folders.TDB_TemplatedParts.Exists("TDB_TemplatePosn.Tag='"& strTag &"'") Then
										set objDBPartPos = objDB.Folders.TDB_TemplatedParts.first("TDB_TemplatePosn.Tag='"& strTag &"'")
									Else
										If objDB.TDB_Template.folders.TDB_TemplatePosns.Exists("Tag='"& strTag &"'") Then
											if UBound(arrParts)>1 Then
												set objDBPartPos = objDB.Folders.TDB_TemplatedParts.CreateNewInstance(,,arrParts(2))
											Else
												set objDBPartPos = objDB.Folders.TDB_TemplatedParts.CreateNewInstance
											End If
											objDBPartPos.TDB_TemplatePosn = objDB.TDB_Template.folders.TDB_TemplatePosns.First("Tag='"& strTag &"'")
										End If
									End If
									
									If Not objDBPartPos Is Nothing Then
										If Not objDBPartPos.IsNull Then
											objDBPartPos.TDB_DashboardPart = System.folders.TDB_DashboardPart.First("Title.Primary='"& arrParts(0) &"'")
											If UBound(arrParts)>0 Then
												objDBPartPos.Title = arrParts(1)
											End If
											objDBPartPos.save objTransaction
										End If
									End If									
								Else
									objSystemLog.LogNormalDetail "The touch dashboard part doesn't exists in system : " & arrParts(0) 
								End If
							Next
							
							'Permit groups
							AddUserGroups objDB,arrcfg(2),objTransaction

							If blnProcess Then
								objDB.save objTransaction
								set objval = objTransaction.validate		
								if objval.status <> 3 then
									intAddedTouchDB = intAddedTouchDB + 1
									objTransaction.Commit
									objSystemLog.LogNormalDetail "The '"& arrTokens(0) &"' touch dashboard has been created/updated."
								else
									intErroredTouchDB = intErroredTouchDB + 1
									objSystemLog.LogErrorDetail "","","There were some error in updating/creating the '"& arrTokens(0) &"' touch dashboard : '" & objval.Result.Message & "'" 
								end if								
							Else
								intErroredTouchDB = intErroredTouchDB + 1
								objSystemLog.LogNormalDetail "Something wrong : " & strDBKey
							End If
						Else
							intErroredTouchDB = intErroredTouchDB + 1
							objSystemLog.LogNormalDetail "Something wrong : " & strDBKey
						End If 
					Else
						objSystemLog.LogNormalDetail "File is empty : " & objFile.Name
						intErroredTouchDB = intErroredTouchDB + 1
					End IF
				End IF
			Else
				intErroredTouchDB = intErroredTouchDB + 1
				objSystemLog.LogNormalDetail "File doesn't exists : " & strCodeFilePath
			End IF
		End IF
	Next
	
	objSystemLog.LogNormalDetail "The touch dashboard loading process completed."
End function

function loadKPIS(System,strDirectoryPath)
	Dim objFSO                 : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objFolder              : Set objFolder = objFSO.GetFolder(strDirectoryPath)
	objSystemLog.LogNormalDetail "The KPIs loading process started."
	strDirectoryPath = AddEndSlashIfNecessary(strDirectoryPath)
	
	For Each objFile In objFolder.Files
		blnProcess = True
		strFileName = objFile.Name
		arrTokens   = Split(strFileName,".")
		strKPIKey = arrTokens(0)

		If (UCase(arrTokens(1)) = "TXT") Then
			intTotalKPIS = intTotalKPIS + 1
			strCodeFilePath = objFile.Path			
			If objFSO.FileExists(strCodeFilePath) Then
				Set objCodeFile = objFSO.OpenTextFile(strCodeFilePath, 1)
				If objCodeFile.AtEndOfStream Then
					objSystemLog.LogNormalDetail "File is empty : " & objFile.Name
					intErroredKPIS = intErroredKPIS + 1
				Else
					strCode = objCodeFile.ReadAll
					If Len(Trim(strCode) & "") > 0 Then
						arrcfg = parseConfigs(Array("KPIMEMBERS","PERMITGROUP","KPITRIGGERS","KEYMAPPING"),strCode)
						strFilter = "Key='"& strKPIKey &"'"
						If arrcfg(3).Count>0 Then
							strFilter = arrcfg(3).Item("Expression")
						End If
						blnProcess = True
						If System.Folders.KPIs.exists(strFilter) Then
							set objKPI = System.Folders.KPIs.first(strFilter)
						Else
							set objKPI = System.Folders.KPIs.CreateNewInstance(,,strKPIKey)
							if arrcfg(0).Exists("FolderName") Then
								objKPI.FolderName = arrcfg(0).Item("FolderName")
							Else
								blnProcess = False
							End If
						End If
						
						If blnProcess Then
							set objTransaction = System.BeginTransaction
							If objFSO.FileExists(strDirectoryPath & "DEFINITION\" & strFileName) Then
								set objFileToRead = objFSO.OpenTextFile(strDirectoryPath & "DEFINITION\" & strFileName,1)
								strFileText = objFileToRead.ReadAll()
								if Len(strFileText & "")>0 Then
									objKPI.Definition = Trim(strFileText)
								End If
							End If
							
							For Each strKey in arrcfg(0).Keys
								If UCase(strKey) <> "FOLDERNAME" and UCase(strKey) <> "KEY" Then
									If UCase(strKey) <> "KPI_GROUP" Then
										execute "objKPI." & strKey & " = """ & arrcfg(0).Item(strKey) & """"
									Else 
										If System.Folders.kpi_Groups.Exists("Title.Primary='"& arrcfg(0).Item("kpi_Group") &"'") Then
											objKPI.Kpi_Group = System.Folders.kpi_Groups.first("Title.Primary='"& arrcfg(0).Item("kpi_Group") &"'")
										End If
									End If
								End If
							Next
							
							
							AddUserGroups objKPI,arrcfg(1),objTransaction
														
							objKPI.save objTransaction
							set objVal = objTransaction.Validate
							if objval.status <> 3 then
								intAddedKPIS = intAddedKPIS + 1
								objTransaction.Commit
								objSystemLog.LogNormalDetail "The '"& strKPIKey &"' KPI has been created/updated."
							else
								objSystemLog.LogErrorDetail "","","There were some error in updating/creating the '"& strKPIKey &"' KPI : '" & objval.Result.Message & "'" 
								intErroredKPIS = intErroredKPIS + 1
							end if		
						Else
							intErroredKPIS = intErroredKPIS + 1
						End If
					End If
				End If
			End If
		End If
	Next

	objSystemLog.LogNormalDetail "The KPIs loading process Completed."
End function

'PPatel - 20200528 - Views Deployment
function loadViews(System,strDirectoryPath)
	Dim objFSO                 : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objFolder              : Set objFolder = objFSO.GetFolder(strDirectoryPath)
	objSystemLog.LogNormalDetail "The Views loading process started."
	strDirectoryPath = AddEndSlashIfNecessary(strDirectoryPath)
	
	For Each objFile In objFolder.Files
		blnProcess = True
		strFileName = objFile.Name
		arrTokens   = Split(strFileName,".")
		strViewKey = arrTokens(0)

		If (UCase(arrTokens(1)) = "TXT") Then
			intTotalViews = intTotalViews + 1
			strCodeFilePath = objFile.Path			
			If objFSO.FileExists(strCodeFilePath) Then
				Set objCodeFile = objFSO.OpenTextFile(strCodeFilePath, 1)
				If objCodeFile.AtEndOfStream Then
					objSystemLog.LogNormalDetail "File is empty : " & objFile.Name
					intErroredViews = intErroredViews + 1
				Else
					strCode = objCodeFile.ReadAll
					If Len(Trim(strCode) & "") > 0 Then
						arrcfg = parseConfigs(Array("viewmembers","permitgroup","keymapping"),strCode)
						strFilter = "Key='"& strViewKey &"'"
						If arrcfg(2).Count>0 Then
							strFilter = arrcfg(2).Item("Expression")
						End If
						blnProcess = True
						If System.Folders.Views.exists(strFilter) Then
							set objView = System.Folders.Views.first(strFilter)
						Else
							set objView = System.Folders.Views.CreateNewInstance(,,strViewKey)
						End If
						
						If blnProcess Then
							set objTransaction = System.BeginTransaction
							If objFSO.FileExists(strDirectoryPath & "DEFINITION\" & strFileName) Then
								set objFileToRead = objFSO.OpenTextFile(strDirectoryPath & "DEFINITION\" & strFileName,1)
								strFileText = objFileToRead.ReadAll()
								if Len(strFileText & "")>0 Then
									objView.State = Trim(strFileText)
								End If
							End If
							
							For Each strKey in arrcfg(0).Keys
								If UCase(strKey) <> "FOLDERNAME" and UCase(strKey) <> "KEY" Then
									If UCase(strKey) <> "VIEWGROUP" Then
										execute "objView." & strKey & " = """ & arrcfg(0).Item(strKey) & """"
									Else 
										If System.Folders.ViewGroups.Exists("Title.Primary='"& arrcfg(0).Item("ViewGroup") &"'") Then
											set objView.ViewGroup = System.Folders.ViewGroups.first("Title.Primary='"& arrcfg(0).Item("ViewGroup") &"'")
										End If
									End If
								End If
							Next
							
							
							AddUserGroups objView,arrcfg(1),objTransaction
														
							objView.save objTransaction
							set objVal = objTransaction.Validate
							if objval.status <> 3 then
								intAddedViews = intAddedViews + 1
								objTransaction.Commit
								objSystemLog.LogNormalDetail "The '"& strViewKey &"' View has been created/updated."
							else
								objSystemLog.LogErrorDetail "","","There were some error in updating/creating the '"& strViewKey &"' View : '" & objval.Result.Message & "'" 
								intErroredViews = intErroredViews + 1
							end if		
						Else
							intErroredViews = intErroredViews + 1
						End If
					End If
				End If
			End If
		End If
	Next

	objSystemLog.LogNormalDetail "The Views loading process Completed."
End function

'PPatel - 20200528 - Views Group Deployment
function loadViewGroups(System,strDirectoryPath)
	Dim objFSO                 : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objFolder              : Set objFolder = objFSO.GetFolder(strDirectoryPath)
	objSystemLog.LogNormalDetail "The Views loading process started."
	strDirectoryPath = AddEndSlashIfNecessary(strDirectoryPath)
	
	For Each objFile In objFolder.Files
		blnProcess = True
		strFileName = objFile.Name
		arrTokens   = Split(strFileName,".")
		strViewGroupKey = arrTokens(0)

		If (UCase(arrTokens(1)) = "TXT") Then
			intTotalViewGroups = intTotalViewGroups + 1
			strCodeFilePath = objFile.Path			
			If objFSO.FileExists(strCodeFilePath) Then
				Set objCodeFile = objFSO.OpenTextFile(strCodeFilePath, 1)
				If objCodeFile.AtEndOfStream Then
					objSystemLog.LogNormalDetail "File is empty : " & objFile.Name
					intErroredViewGroups = intErroredViewGroups + 1
				Else
					strCode = objCodeFile.ReadAll
					If Len(Trim(strCode) & "") > 0 Then
						arrcfg = parseConfigs(Array("ViewGroupmembers","permitgroup","keymapping"),strCode)
						strFilter = "Key='"& strViewGroupKey &"'"
						If arrcfg(2).Count>0 Then
							strFilter = arrcfg(2).Item("Expression")
						End If
						blnProcess = True
						If System.Folders.ViewGroups.exists(strFilter) Then
							set objViewGroup = System.Folders.ViewGroups.first(strFilter)
						Else
							set objViewGroup = System.Folders.ViewGroups.CreateNewInstance(,,strViewGroupKey)
						End If
						
						If blnProcess Then
							set objTransaction = System.BeginTransaction
							
							For Each strKey in arrcfg(0).Keys
								If UCase(strKey) <> "FOLDERNAME" and UCase(strKey) <> "KEY" Then
									execute "objViewGroup." & strKey & " = """ & arrcfg(0).Item(strKey) & """"
								end if
							Next
							
							
							AddUserGroups objViewGroup,arrcfg(1),objTransaction
														
							objViewGroup.save objTransaction
							set objVal = objTransaction.Validate
							if objval.status <> 3 then
								intAddedViewGroups = intAddedViewGroups + 1
								objTransaction.Commit
								objSystemLog.LogNormalDetail "The '"& strViewGroupKey &"' ViewGroup has been created/updated."
							else
								objSystemLog.LogErrorDetail "","","There were some error in updating/creating the '"& strViewGroupKey &"' ViewGroup : '" & objval.Result.Message & "'" 
								intErroredViewGroups = intErroredViewGroups + 1
							end if		
						Else
							intErroredViewGroups = intErroredViewGroups + 1
						End If
					End If
				End If
			End If
		End If
	Next

	objSystemLog.LogNormalDetail "The ViewGroups loading process Completed."
End function

'PPatel - 20200528 - Report Template
function loadReportTemplates(System,strDirectoryPath)
	Dim objFSO                 : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objFolder              : Set objFolder = objFSO.GetFolder(strDirectoryPath)
	objSystemLog.LogNormalDetail "The Report Template loading process started."
	strDirectoryPath = AddEndSlashIfNecessary(strDirectoryPath)
	
	For Each objFile In objFolder.Files
		blnProcess = True
		strFileName = objFile.Name
		arrTokens   = Split(strFileName,".")
		strReportTemplateskey = arrTokens(0)

		If (UCase(arrTokens(1)) = "TXT") Then
			intTotalReportTemplates = intTotalReportTemplates + 1
			strCodeFilePath = objFile.Path			
			If objFSO.FileExists(strCodeFilePath) Then
				Set objCodeFile = objFSO.OpenTextFile(strCodeFilePath, 1)
				If objCodeFile.AtEndOfStream Then
					objSystemLog.LogNormalDetail "File is empty : " & objFile.Name
					intErroredReportTemplates = intErroredReportTemplates + 1
				Else
					strCode = objCodeFile.ReadAll
					If Len(Trim(strCode) & "") > 0 Then
						arrcfg = parseConfigs(Array("keymapping","members"),strCode)
						strFilter = "Key='"& strReportTemplateskey &"'"
						If arrcfg(0).Count > 0 Then
							strFilter = arrcfg(0).Item("Expression")
						End If
							
						if System.Folders.PZ_TouchReportTemplates.Exists(strFilter) Then
							set objInst = System.Folders.PZ_TouchReportTemplates.first(strFilter)
						Else
							set objInst = System.Folders.PZ_TouchReportTemplates.CreateNewInstance(,,strReportTemplateskey)
						End If
						
						If Not objInst.IsNull Then
							set objTransaction = System.BeginTransaction
							For Each strKey in arrcfg(1).Keys
								if UCase(strKey) <> "KEY" Then
									If UCase(strKey) <> "PZ_PRINTER" Then
										execute "objInst." & strKey & " = """ & arrcfg(1).Item(strKey) & """"
									Else 
										If System.Folders.PZ_Printers.Exists("Description.Primary='"& arrcfg(1).Item("PZ_PRINTER") &"'") Then
											set objInst.PZ_Printer = System.Folders.PZ_Printers.first("Description.Primary='"& arrcfg(0).Item("PZ_PRINTER") &"'")
										End If
									End If
								End If
							Next
							
							'Update Templates block
							strTemplateFilePath = strDirectoryPath & "Templates\" 
							Set objTemplateFolder = objFSO.GetFolder(strTemplateFilePath)
							For Each objTemplateFile in objTemplateFolder.Files
								i = InStrRev(objTemplateFile.Name, ".")
								If Mid(objTemplateFile.Name, 1, i - 1) = strReportTemplateskey Then
									Set objWFXMLStream = objTemplateFile.OpenAsTextStream(1,0)
									objInst.ReportTemplate  = objWFXMLStream.ReadAll
								End If
							Next	
							
							objInst.save objTransaction
							set objVal = objTransaction.Validate
							if objval.status <> 3 then
								intAddedReportTemplates = intAddedReportTemplates + 1
								objTransaction.Commit
								objSystemLog.LogNormalDetail "The '"& strReportTemplateskey &" has been created/updated."
							else
								intErroredReportTemplates = intErroredReportTemplates + 1
								objSystemLog.LogErrorDetail "","","There were some error in updating/creating the '"& strReportTemplateskey &" : '" & objval.Result.Message & "'" 
							end if
						Else
							intErroredReportTemplates = intErroredReportTemplates + 1
						End If
					Else
						intErroredReportTemplates = intErroredReportTemplates + 1
					End If
				End If
			Else
				intErroredReportTemplates = intErroredReportTemplates + 1
			End If
		End If
	Next
End function

'PPatel - 20200528 - Report Template
function loadSummaryTemplates(System,strDirectoryPath)
	Dim objFSO                 : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objFolder              : Set objFolder = objFSO.GetFolder(strDirectoryPath)
	objSystemLog.LogNormalDetail "The Summary Template loading process started."
	strDirectoryPath = AddEndSlashIfNecessary(strDirectoryPath)
	
	For Each objFile In objFolder.Files
		blnProcess = True
		strFileName = objFile.Name
		arrTokens   = Split(strFileName,".")
		strSummaryTemplateskey = arrTokens(0)

		If (UCase(arrTokens(1)) = "TXT") Then
			intTotalSummaryTemplates = intTotalSummaryTemplates + 1
			strCodeFilePath = objFile.Path			
			If objFSO.FileExists(strCodeFilePath) Then
				Set objCodeFile = objFSO.OpenTextFile(strCodeFilePath, 1)
				If objCodeFile.AtEndOfStream Then
					objSystemLog.LogNormalDetail "File is empty : " & objFile.Name
					intErroredSummaryTemplates = intErroredSummaryTemplates + 1
				Else
					strCode = objCodeFile.ReadAll
					If Len(Trim(strCode) & "") > 0 Then
						arrcfg = parseConfigs(Array("keymapping","members"),strCode)
						strFilter = "Key='"& strSummaryTemplateskey &"'"
						If arrcfg(0).Count > 0 Then
							strFilter = arrcfg(0).Item("Expression")
						End If
						if System.Folders.SummaryTemplate.Exists(strFilter) Then
							set objInst = System.Folders.SummaryTemplate.first(strFilter)
						Else
							set objInst = System.Folders.SummaryTemplate.CreateNewInstance(,,strSummaryTemplateskey)
						End If
						
						If Not objInst.IsNull Then
							set objTransaction = System.BeginTransaction
							For Each strKey in arrcfg(1).Keys
								if UCase(strKey) <> "KEY" Then
									execute "objInst." & strKey & " = """ & arrcfg(1).Item(strKey) & """"
								End If
							Next
							
							'Update Templates block
							strTemplateFilePath = strDirectoryPath & "Templates\" 
							Set objTemplateFolder = objFSO.GetFolder(strTemplateFilePath)
							For Each objTemplateFile In objTemplateFolder.Files
								i = InStrRev(objTemplateFile.Name, ".")
								If Mid(objTemplateFile.Name, 1, i - 1) = strSummaryTemplateskey Then
									Set objWFXMLStream = objTemplateFile.OpenAsTextStream(1,0)
									objInst.SummaryTemplate  = objWFXMLStream.ReadAll
								End If
							Next	
							
							objInst.save objTransaction
							set objVal = objTransaction.Validate
							if objval.status <> 3 then
								intAddedSummaryTemplates = intAddedSummaryTemplates + 1
								objTransaction.Commit
								objSystemLog.LogNormalDetail "The '"& strSummaryTemplateskey &" has been created/updated."
							else
								intErroredSummaryTemplates = intErroredSummaryTemplates + 1
								objSystemLog.LogErrorDetail "","","There were some error in updating/creating the '"& strSummaryTemplateskey &" : '" & objval.Result.Message & "'" 
							end if
						Else
							intErroredSummaryTemplates = intErroredSummaryTemplates + 1
						End If
					Else
						intErroredSummaryTemplates = intErroredSummaryTemplates + 1
					End If
				End If
			Else
				intErroredSummaryTemplates = intErroredSummaryTemplates + 1
			End If
		End If
	Next
End function

function loadDynamicMemberDefinitions(System,strDirectoryPath)
	Dim objFSO                 : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objFolder              : Set objFolder = objFSO.GetFolder(strDirectoryPath)
	objSystemLog.LogNormalDetail "The Summary Template loading process started."
	strDirectoryPath = AddEndSlashIfNecessary(strDirectoryPath)
	
	For Each objFile In objFolder.Files
		blnProcess = True
		strFileName = objFile.Name
		arrTokens   = Split(strFileName,".")
		strDynamicMemberDefinitionskey = arrTokens(0)

		If (UCase(arrTokens(1)) = "TXT") Then
			intTotalDynamicMemberDefinitions = intTotalDynamicMemberDefinitions + 1
			strCodeFilePath = objFile.Path			
			If objFSO.FileExists(strCodeFilePath) Then
				Set objCodeFile = objFSO.OpenTextFile(strCodeFilePath, 1)
				If objCodeFile.AtEndOfStream Then
					objSystemLog.LogNormalDetail "File is empty : " & objFile.Name
					intErroredDynamicMemberDefinitions = intErroredDynamicMemberDefinitions + 1
				Else
					strCode = objCodeFile.ReadAll
					If Len(Trim(strCode) & "") > 0 Then
						arrcfg = parseConfigs(Array("keymapping","members","permitgroup"),strCode)
						strFilter = "Key='"& strDynamicMemberDefinitionskey &"'"
						If arrcfg(0).Count > 0 Then
							strFilter = arrcfg(0).Item("Expression")
						End If
						strFilter = strFilter & "and DefinitionStatus.value <> 'D'"
						if System.Folders.DynamicMemberDefinitions.Exists(strFilter) Then
							set objInst = System.Folders.DynamicMemberDefinitions.first(strFilter)
						Else
							set objInst = System.Folders.DynamicMemberDefinitions.CreateNewInstance(,,strDynamicMemberDefinitionskey)
						End If
						
						If Not objInst.IsNull Then
							Dim strSEClassName : strSEClassName = ""
							set objTransaction = System.BeginTransaction
							strMemberName =arrcfg(1).Item("MemberName")
							strMemberTitle =arrcfg(1).Item("MemberTitle")
							strTitle =arrcfg(1).Item("Title")
							strClassName =arrcfg(1).Item("ClassName")
							strSEClassName =arrcfg(1).Item("StorageExpressionClassName")
							strlength =arrcfg(1).Item("length")
							strScale =arrcfg(1).Item("Scale")
							strDataType =arrcfg(1).Item("DataType")
							strOrderOnCRMForm =arrcfg(1).Item("OrderOnCRMForm")
							strLinesOnCRMForm =arrcfg(1).Item("LinesOnCRMForm")
							strVisibleOnForm =arrcfg(1).Item("VisibleOnForm")
							strMustNotBeNull =arrcfg(1).Item("MustNotBeNull")
							strCRMSearch =arrcfg(1).Item("CRMSearch")
							strMultiSelect =arrcfg(1).Item("MultiSelect")
							strFolderName =arrcfg(1).Item("FolderName")
							strFolderTitle =arrcfg(1).Item("FolderTitle")
							strTouchCanAccessExpression =arrcfg(1).Item("TouchCanAccessExpression")
							strTouchCanModifyExpression =arrcfg(1).Item("TouchCanModifyExpression")
							strTouchFormOrder =arrcfg(1).Item("TouchFormOrder")
							strCanModifyExpression =arrcfg(1).Item("CanModifyExpression")
							strCanAccessExpression =arrcfg(1).Item("CanAccessExpression")
							strDefaultValue = arrcfg(1).Item("DefaultValue")
							strCanViewRetiredValue = arrcfg(1).Item("CanViewRetiredValue")
							strTouchFormPositioning = arrcfg(1).Item("TouchFormPositioning")
							strTouchFormElement = arrcfg(1).Item("TouchFormElement")
							strDefinitionStatus = arrcfg(1).Item("DefinitionStatus")
							strEnumerationName = arrcfg(1).Item("EnumerationName")
							cfgExpression = arrcfg(1).Item("Expression")
							cfgStorageExpression = arrcfg(1).Item("StorageExpression")
							'blnUpdatePicklistValue = arrcfg(1).Item("UpdatePicklistValue")
							
							'Update Definitions block
							if UCase(strDataType)= "ENUM" then
								If objFSO.FileExists(strDirectoryPath & "DEFINITION\" & strFileName) Then
									set objFileToRead = objFSO.OpenTextFile(strDirectoryPath & "DEFINITION\" & strFileName,1)
									strFileText = objFileToRead.ReadAll()
									if Len(strFileText & "")>0 Then
										strDefinition = Trim(strFileText)
									End If
								End If
							End If
							
							' if UCase(strDataType)= "ENUM" And blnUpdatePicklistValue then
								' objInst.UpdateMemberType System.Classes("DynamicEnumMember")
								' objInst.Definition.EnumerationDefinition = strDefinition
							' End IF
							
							'If Not blnUpdatePicklistValue Then
								call UpdateDM(System,objInst,strMemberName,strMemberTitle,strTitle,strClassName,strSEClassName,strlength,strScale,strDataType,strOrderOnCRMForm,strLinesOnCRMForm,strVisibleOnForm,strCRMSearch,strDefaultValue,strMustNotBeNull,strMultiSelect,strFolderName,strFolderTitle,strTouchCanAccessExpression,strTouchCanModifyExpression,strTouchFormOrder,strCanModifyExpression,strCanAccessExpression,strCanViewRetiredValue,strTouchFormPositioning,strTouchFormElement,strDefinitionStatus,strEnumerationName,strDefinition,arrcfg(2),objTransaction,cfgExpression,cfgStorageExpression)
							'End If
							objInst.save objTransaction
							set objVal = objTransaction.Validate
							if objval.status <> 3 then
								intAddedDynamicMemberDefinitions = intAddedDynamicMemberDefinitions + 1
								objTransaction.Commit
								objSystemLog.LogNormalDetail "The '"& strDynamicMemberDefinitionskey &" has been created/updated."
							else
								intErroredDynamicMemberDefinitions = intErroredDynamicMemberDefinitions + 1
								objSystemLog.LogErrorDetail "","","There were some error in updating/creating the '"& strDynamicMemberDefinitionskey &" : '" & objval.Result.Message & "'" 
							end if
						Else
							intErroredDynamicMemberDefinitions = intErroredDynamicMemberDefinitions + 1
						End If
					Else
						intErroredDynamicMemberDefinitions = intErroredDynamicMemberDefinitions + 1
					End If
				End If
			Else
				intErroredDynamicMemberDefinitions = intErroredDynamicMemberDefinitions + 1
			End If
		End If
	Next
End function

function UpdateDM(System,oDynamicMemberDef,strMemberName,strMemberTitle,strTitle,strClassName,strSEClassName,strlength,strScale,strDataType,strOrderOnCRMForm,strLinesOnCRMForm,strVisibleOnForm,strCRMSearch,strDefaultValue,strMustNotBeNull,strMultiSelect,strFolderName,strFolderTitle,strTouchCanAccessExpression,strTouchCanModifyExpression,strTouchFormOrder,strCanModifyExpression,strCanAccessExpression,strCanViewRetiredValue,strTouchFormPositioning,strTouchFormElement,strDefinitionStatus,strEnumerationName,strDefinition,strUserGroups,objTransaction,cfgExpression,cfgStorageExpression)
	If Not oDynamicMemberDef Is Nothing Then
		if oDynamicMemberDef.isnew then
			oDynamicMemberDef.ClassName = strClassName
			oDynamicMemberDef.MemberName = strMemberName
		else
			if oDynamicMemberDef.DefinitionStatus.value <> "U" Then
				oDynamicMemberDef.DefinitionStatus.value = "U"				
			end if
		end if
		guidDMDKey = Replace(Replace(Lcase(oDynamicMemberDef.Key),"{",""),"}","")
		
		If UCase(strDataType) = UCase("DATE") Then
			intDataType = 10
			If strClassName = "CU_Display" Then
				strStorageExpression = cfgStorageExpression
				strExpression = cfgExpression
			Else
				strStorageExpression = strClassName & "_DM_Ext1.Data_DM_Date"
				strExpression = strClassName & "_DM_Ext1_Folder.First(Filter_Name_DM='"& guidDMDKey &"').Data_DM_Date"
			End If
		ElseIf UCase(strDataType) = UCase("Date and time") Then
			intDataType = 12
			strStorageExpression = strClassName & "_DM_Ext1.Data_DM_Date"
			strExpression = strClassName & "_DM_Ext1_Folder.First(Filter_Name_DM='"& guidDMDKey &"').Data_DM_Date"
		ElseIf UCase(strDataType) = UCase("True or false") Then
			intDataType = 13
			If strClassName = "CU_Display" or strClassName = "PromoNormal" Then
				strStorageExpression = cfgStorageExpression
				strExpression = cfgExpression
			Else
				strStorageExpression = strClassName & "_DM_Ext1.Data_DM_Bit"
				strExpression = strClassName & "_DM_Ext1_Folder.First(Filter_Name_DM='"& guidDMDKey &"').Data_DM_Bit"
			End If
		ElseIf UCase(strDataType) = UCase("ENUM") Then
			If strClassName = "stdUser" or strClassName =  "PromoNormal" or strClassName = "CU_Display" Then
				strStorageExpression = cfgStorageExpression
				strExpression = cfgExpression
			Else
				If strSEClassName <> "" Then
					strStorageExpression = strSEClassName & "_DM_Ext1.Data_DM"
					strExpression = strSEClassName & "_DM_Ext1s.First(Filter_Name_DM='"& guidDMDKey &"').Data_DM"
				Else
					strStorageExpression = strClassName & "_DM_Ext1.Data_DM"
					strExpression = strClassName & "_DM_Ext1_Folder.First(Filter_Name_DM='"& guidDMDKey &"').Data_DM"
				End If
			End IF
		ElseIf UCase(strDataType) = UCase("Number") Then
			intDataType = 16
			If strClassName = "CU_Display" Or strClassName = "VT_ScheduleEntry" Then
				strStorageExpression = cfgStorageExpression
				strExpression = cfgExpression
			Else
				If strSEClassName <> "" Then
					strStorageExpression = strSEClassName & "_DM_Ext1.Data_DM_Num"
					strExpression = strSEClassName & "_DM_Ext1s.First(Filter_Name_DM='"& guidDMDKey &"').Data_DM_Num"
				Else
					strStorageExpression = strClassName & "_DM_Ext1.Data_DM_Num"
					strExpression = strClassName & "_DM_Ext1_Folder.First(Filter_Name_DM='"& guidDMDKey &"').Data_DM_Num"
				End If
			End If
		ElseIf UCase(strDataType) = UCase("Number with decimal") Then
			intDataType = 20
			If strClassName = "stdUser" Or strClassName = "CU_Display" Then
				strStorageExpression = cfgStorageExpression
				strExpression = cfgExpression
			Else
				strStorageExpression = strClassName & "_DM_Ext1.Data_DM_Num"
				strExpression = strClassName & "_DM_Ext1_Folder.First(Filter_Name_DM='"& guidDMDKey &"').Data_DM_Num"
			End If
		ElseIf UCase(strDataType) = UCase("Text") Then
			intDataType = 5
			If strClassName = "stdUser" Or strClassName = "VT_ScheduleEntry" Or strClassName = "CU_Display" Or strClassName = "AS_Territory" or strClassName = "PromoNormal" Then
				strStorageExpression = cfgStorageExpression
				strExpression = cfgExpression
			Else
				If strSEClassName <> "" Then
					strStorageExpression = strSEClassName & "_DM_Ext1.Data_DM"
					strExpression = strSEClassName & "_DM_Ext1_Folder.First(Filter_Name_DM='"& guidDMDKey &"').Data_DM"
				Else
					strStorageExpression = strClassName & "_DM_Ext1.Data_DM"
					strExpression = strClassName & "_DM_Ext1_Folder.First(Filter_Name_DM='"& guidDMDKey &"').Data_DM"
				End If
			End If
		End If
		
		oDynamicMemberDef.StorageExpression = strStorageExpression
		oDynamicMemberDef.Expression = strExpression
		if UCase(strDataType)= "ENUM" then
			oDynamicMemberDef.UpdateMemberType System.Classes("DynamicEnumMember")
			oDynamicMemberDef.Definition.EnumerationDefinition = strDefinition
			if strMultiSelect<> "" then
				oDynamicMemberDef.Definition.MultiSelect = CBool(strMultiSelect)
			end if
			if strFolderName <> "" then
				oDynamicMemberDef.Definition.FolderName = strFolderName
			end if
			if strFolderTitle <> "" then
				oDynamicMemberDef.Definition.FolderTitle.Primary = strFolderTitle
			end if
			if strCanViewRetiredValue <> "" then
				oDynamicMemberDef.Definition.CanViewRetiredValue = strCanViewRetiredValue 
			end if
			if strEnumerationName<> ""then
				oDynamicMemberDef.Definition.EnumerationName = strEnumerationName 
			end if
		else
			oDynamicMemberDef.Definition.DataType.value = intDataType
			oDynamicMemberDef.Definition.length = strlength
			oDynamicMemberDef.Definition.Scale = strScale
		end if
		oDynamicMemberDef.Definition.DefaultValue = strDefaultValue
		oDynamicMemberDef.MemberTitle.Primary = strMemberTitle
		oDynamicMemberDef.Title.Primary = strTitle
		
		if strOrderOnCRMForm <> "" Then oDynamicMemberDef.Definition.OrderOnCRMForm = CInt(strOrderOnCRMForm)
		if strLinesOnCRMForm <> "" Then oDynamicMemberDef.Definition.LinesOnCRMForm = CInt(strLinesOnCRMForm)
		if strVisibleOnForm <> "" Then oDynamicMemberDef.Definition.VisibleOnForm = CBool(strVisibleOnForm)
		if strCRMSearch <> "" Then oDynamicMemberDef.Definition.CRMSearch = CBool(strCRMSearch)
		if strMustNotBeNull <> "" Then oDynamicMemberDef.Definition.MustNotBeNull = CBool(strMustNotBeNull)
		
		oDynamicMemberDef.TouchCanAccessExpression = strTouchCanAccessExpression
		oDynamicMemberDef.TouchCanModifyExpression = strTouchCanModifyExpression
		if strTouchFormOrder <> "" Then oDynamicMemberDef.TouchFormOrder = CInt(strTouchFormOrder)
		if StrTouchFormPositioning<> "" then
			oDynamicMemberDef.TouchFormPositioning.value  = StrTouchFormPositioning 
		end if
		if strTouchFormElement<> "" then
			if system.folders.TouchFormRegion.Exists("Description.Primary = '" & strTouchFormElement &"'") then
				set oDynamicMemberDef.TouchFormElement  = system.folders.TouchFormRegion.first("Description.Primary = '" & strTouchFormElement &"'") 
			end if
		end if
		
		oDynamicMemberDef.CanModifyExpression = strCanModifyExpression
		oDynamicMemberDef.CanAccessExpression = strCanAccessExpression
		
		If strClassName = "DC_Question" Then
			Dim objDMDParentInUse : Set objDMDParentInUse = Nothing
			arrQuestionClass = Array("DC_Question","DC_Q_Checklist","DC_Q_Date","DC_Q_DateTime","DC_Q_Decimal","DC_Q_DMReferencePhoto","DC_Q_Duration","DC_Q_Integer","DC_Q_Photo","DC_Q_Picklist"," DC_Q_ProdCount","DC_Q_Signature","DC_Q_Text","DC_Q_YNNA")
			For I = 0 To ubound(arrQuestionClass)
				strKey = arrQuestionClass(i) & "." & cfgStorageExpression
				set objDMDInUse = oDynamicMemberDef.Folders.DynamicInUseMembers.First("Ucase(Key) = '"& Ucase(strKey) &"'")
				If objDMDInUse.IsNull then
					set objDMDInUse = oDynamicMemberDef.Folders.DynamicInUseMembers.CreateNewInstance(,,strKey)
					If arrQuestionClass(i) = "DC_Question" Then
						Set objDMDParentInUse = objDMDInUse
					Else
						If Not objDMDParentInUse Is Nothing Then Set objDMDInUse.Parent = objDMDParentInUse
					End If
					objDMDInUse.save objTransaction
				End IF
			Next 
		ElseIf strClassName = "CU_Display" Then
			strKey = strClassName & "." & strStorageExpression
			set objDMDInUse = oDynamicMemberDef.Folders.DynamicInUseMembers.First("Ucase(Key) = '"& Ucase(strKey) &"'")
			If objDMDInUse.IsNull then
				set objDMDInUse = oDynamicMemberDef.Folders.DynamicInUseMembers.CreateNewInstance(,,strKey)
				objDMDInUse.save objTransaction
			End IF
		ElseIf strClassName = "VT_ScheduleEntry" Then
			strKey = strClassName & "." & strStorageExpression
			set objDMDInUse = oDynamicMemberDef.Folders.DynamicInUseMembers.First("Ucase(Key) = '"& Ucase(strKey) &"'")
			If objDMDInUse.IsNull then
				set objDMDInUse = oDynamicMemberDef.Folders.DynamicInUseMembers.CreateNewInstance(,,strKey)
				objDMDInUse.save objTransaction
			End IF
		ElseIf strClassName = "stdUser" Then
			set objDMDInUse = oDynamicMemberDef.Folders.DynamicInUseMembers.First
			If objDMDInUse.IsNull then
				set objDMDInUse = oDynamicMemberDef.Folders.DynamicInUseMembers.CreateNewInstance(,,"stdUser." & strStorageExpression)
				objDMDInUse.save objTransaction
			End IF
		Else
			set objDMDInUse = oDynamicMemberDef.Folders.DynamicInUseMembers.First
			If objDMDInUse.IsNull then
				set objDMDInUse = oDynamicMemberDef.Folders.DynamicInUseMembers.CreateNewInstance(,,guidDMDKey & "." & strStorageExpression)
				objDMDInUse.save objTransaction
			End IF
		End If
		
		call AddUserGroups (oDynamicMemberDef,strUserGroups,objTransaction)
		oDynamicMemberDef.DefinitionStatus.value = strDefinitionStatus
	End IF		
End function
'PPatel - 20200528 - Report Template
function loadPushReports(System,strDirectoryPath)
	Dim objFSO                 : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objFolder              : Set objFolder = objFSO.GetFolder(strDirectoryPath)
	objSystemLog.LogNormalDetail "The Push Reports loading process started."
	strDirectoryPath = AddEndSlashIfNecessary(strDirectoryPath)
	
	For Each objFile In objFolder.Files
		blnProcess = True
		strFileName = objFile.Name
		arrTokens   = Split(strFileName,".")
		strPushReportskey = arrTokens(0)

		If (UCase(arrTokens(1)) = "TXT") Then
			intTotalPushReports = intTotalPushReports + 1
			strCodeFilePath = objFile.Path			
			If objFSO.FileExists(strCodeFilePath) Then
				Set objCodeFile = objFSO.OpenTextFile(strCodeFilePath, 1)
				If objCodeFile.AtEndOfStream Then
					objSystemLog.LogNormalDetail "File is empty : " & objFile.Name
					intErroredPushReports = intErroredPushReports + 1
				Else
					strCode = objCodeFile.ReadAll
					If Len(Trim(strCode) & "") > 0 Then
						arrcfg = parseConfigs(Array("keymapping","members","distribution","teams"),strCode)
						strFilter = "Key='"& strPushReportskey &"'"
						If arrcfg(0).Count > 0 Then
							strFilter = arrcfg(0).Item("Expression")
						End If
						if System.Folders.PR_Reports.Exists(strFilter) Then
							set objInst = System.Folders.PR_Reports.first(strFilter)
						Else
							set objInst = System.Folders.PR_Reports.CreateNewInstance(,,strPushReportskey)
						End If
						
						If Not objInst.IsNull Then
							set objTransaction = System.BeginTransaction
							For Each strKey in arrcfg(1).Keys
								if UCase(strKey) <> "KEY" Then
									select case UCase(strKey)
									case "PR_REPORTDEF"
										If System.Folders.PR_ReportDef.Exists("Title='"& arrcfg(1).Item(strKey) &"'") Then
											set objInst.PR_ReportDef = System.Folders.PR_ReportDef.first("Title='"& arrcfg(1).Item(strKey) &"'")
										End If
									case "TCG_DX_SYNCRULE"
										If System.Folders.TCG_DX_SyncRule.Exists("EX_External.ID='"& arrcfg(1).Item(strKey) &"'") Then
											set objInst.TCG_DX_SyncRule = System.Folders.TCG_DX_SyncRule.first("EX_External.ID='"& arrcfg(1).Item(strKey) &"'")
										End If
									case "START","FINISH"
										execute "objInst." & strKey & " ="& cdate(arrcfg(1).Item(strKey) )
									case else
										execute "objInst." & strKey & " = """ & arrcfg(1).Item(strKey) & """"
									end select
								End If
							Next
							
							'Update DEFINITION block
							strDefinitionFilePath = strDirectoryPath & "DEFINITION\" 
							Set objDefinitionFolder = objFSO.GetFolder(strDefinitionFilePath)
							For Each objDefinitionFile In objDefinitionFolder.Files
								i = InStrRev(objDefinitionFile.Name, ".")
								If Mid(objDefinitionFile.Name, 1, i - 1) = strPushReportskey and (lcase(objFSO.GetExtensionName(objDefinitionFile)) ="html" or lcase(objFSO.GetExtensionName(objDefinitionFile))="htm") Then
									strFileName =objDefinitionFile.path
									blnSummaryLoaded = false
									If Not System.IsUsingStorageAspect() Then
										blnSummaryLoaded = objInst.HS_SavedSummaryWithImages.LoadHTMLDocument(strFileName)
									Else
										blnSummaryLoaded = objInst.HS_SavedSummaryWithImages.LoadHTMLDocument2(strFileName, objTransaction)
									End If

									If not blnSummaryLoaded Then	
										' Check Properties for error message
										
										If objInst.HS_SavedSummaryWithImages.properties.Exists("LoadHTMLDocument_Error") Then
											strMessage = objInst.HS_SavedSummaryWithImages.Properties("LoadHTMLDocument_Error").value
										Else
											strMessage = System.LoadString("HS_LOADHTMLFAILED")
										End If
										objSystemLog.LogErrorDetail "","","There were some error in Update Summary '" &  strMessage & "' in PushReport '"& strPushReportskey &"'"
									End If
								End If
							Next	
							call AddDistributions(objInst,arrcfg(2),objTransaction)
							call AddTeams(objInst,arrcfg(3),objTransaction)
							objInst.save objTransaction
							set objVal = objTransaction.Validate
							if objval.status <> 3 then
								intAddedPushReports = intAddedPushReports + 1
								objTransaction.Commit
								objSystemLog.LogNormalDetail "The '"& strPushReportskey &" has been created/updated."
							else
								intErroredPushReports = intErroredPushReports + 1
								objSystemLog.LogErrorDetail "","","There were some error in updating/creating the '"& strPushReportskey &" : '" & objval.Result.Message & "'" 
							end if
						Else
							intErroredPushReports = intErroredPushReports + 1
						End If
					Else
						intErroredPushReports = intErroredPushReports + 1
					End If
				End If
			Else
				intErroredPushReports = intErroredPushReports + 1
			End If
		End If
	Next
End function
'SGHUNCHALA - 20200529 - Touoch Event and Touch Event Function Deployment
function LoadTouchEvents(System,strDirectoryPath)
	Dim objFSO                 : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objFolder              : Set objFolder = objFSO.GetFolder(strDirectoryPath)
	objSystemLog.LogNormalDetail "The touch events loading process started."
	strDirectoryPath = AddEndSlashIfNecessary(strDirectoryPath)
	
	For Each objFile In objFolder.Files
		blnProcess = True
		strFileName = objFile.Name
		arrTokens   = Split(strFileName,".")
		strTouchEvent = arrTokens(0)
		arrParts = Split(strTouchEvent,"~")

		If (UCase(arrTokens(1)) = "TXT") Then
			intTotalTouchEvents = intTotalTouchEvents + 1
			strCodeFilePath = objFile.Path			
			If objFSO.FileExists(strCodeFilePath) Then
				Set objCodeFile = objFSO.OpenTextFile(strCodeFilePath, 1)
				If objCodeFile.AtEndOfStream Then
					objSystemLog.LogNormalDetail "File is empty : " & objFile.Name
					intErroredTouchEvent = intErroredTouchEvent + 1
				Else
					strCode = objCodeFile.ReadAll
					If Len(Trim(strCode) & "") > 0 Then
						arrcfg = parseConfigs(Array("keymapping","members"),strCode)
						strFilter = "Key='"& arrParts(0) &"'"
						If arrcfg(0).Count > 0 Then
							strFilter = arrcfg(0).Item("Expression")
						End If
						if System.Folders(arrParts(1)).Exists(strFilter) Then
							set objInst = System.Folders(arrParts(1)).first(strFilter)
						Else
							set objInst = System.Folders(arrParts(1)).CreateNewInstance(,,arrParts(0))
						End If
						
						If Not objInst.IsNull Then
							set objTransaction = System.BeginTransaction
							For Each strKey in arrcfg(1).Keys
								if UCase(strKey) <> "KEY" Then
									execute "objInst." & strKey & " = """ & arrcfg(1).Item(strKey) & """" 
								End If
							Next
							
							'Update javascript block
							If objFSO.FileExists(strDirectoryPath & "JAVASCRIPT\" & strFileName) Then
								set objFileToRead = objFSO.OpenTextFile(strDirectoryPath & "JAVASCRIPT\" & strFileName,1)
								strFileText = objFileToRead.ReadAll()
								if Len(strFileText & "")>0 Then
									objInst.JavaScriptBlock = Trim(strFileText)
								End If
							End If
							objInst.save objTransaction
							set objVal = objTransaction.Validate
							if objval.status <> 3 then
								intAddedTouchEvent = intAddedTouchEvent + 1
								objTransaction.Commit
								objSystemLog.LogNormalDetail "The '"& arrParts(0) &"' "& arrParts(1) &" has been created/updated."
							else
								intErroredTouchEvent = intErroredTouchEvent + 1
								objSystemLog.LogErrorDetail "","","There were some error in updating/creating the '"& arrParts(0) &"' "& arrParts(1) &" : '" & objval.Result.Message & "'" 
							end if
						Else
							intErroredTouchEvent = intErroredTouchEvent + 1
						End If
					Else
						intErroredTouchEvent = intErroredTouchEvent + 1
					End If
				End If
			Else
				intErroredTouchEvent = intErroredTouchEvent + 1
			End If
		End If
	Next
End function
'SGHUNCHALA - 20200529 AnalyticViews Deployment
function LoadAnalyticViews(System,strDirectoryPath)
	Dim objFSO                 : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objFolder              : Set objFolder = objFSO.GetFolder(strDirectoryPath)
	objSystemLog.LogNormalDetail "The Analytic Views loading process started."
	strDirectoryPath = AddEndSlashIfNecessary(strDirectoryPath)
	
	For Each objFile In objFolder.Files
		blnProcess = True
		strFileName = objFile.Name
		arrTokens   = Split(strFileName,".")
		strAnalyticKey = arrTokens(0)

		If (UCase(arrTokens(1)) = "TXT") Then
			intTotalAnalytics = intTotalAnalytics + 1
			strCodeFilePath = objFile.Path			
			If objFSO.FileExists(strCodeFilePath) Then
				Set objCodeFile = objFSO.OpenTextFile(strCodeFilePath, 1)
				If objCodeFile.AtEndOfStream Then
					objSystemLog.LogNormalDetail "File is empty : " & objFile.Name
					intErroredAnalytics = intErroredAnalytics + 1
				Else
					strCode = objCodeFile.ReadAll
					If Len(Trim(strCode) & "") > 0 Then
						arrcfg = parseConfigs(Array("analyticviewmembers","attributes","drillhierarchies","permitgroup","keymapping"),strCode)
						strFilter = "Key='"& strAnalyticKey &"'"
						If arrcfg(4).Count>0 Then
							strFilter = arrcfg(4).Item("Expression")
						End If
						If System.folders.AnalyticViews.Exists(strFilter) Then
							set objAnalytic = System.folders.AnalyticViews.first(strFilter)
						Else
							set objAnalytic = System.folders.AnalyticViews.CreateNewInstance(,,strAnalyticKey)
						End If
						
						If Not objAnalytic.IsNull Then
							set objTransaction = System.BeginTransaction
							For each strKey in arrcfg(0).Keys
								if UCase(strKey) <> "KEY" Then
									IF UCase(strKey) <> "ANALYTICVIEWGROUP" Then
										execute "objAnalytic." & strKey & " = """ & arrcfg(0).Item(strKey) & """"
									Else
										If System.folders.AnalyticViewGroups.Exists("Title.Primary='"& arrcfg(0).Item("AnalyticViewGroup") &"'") Then
											objAnalytic.AnalyticViewGroup = System.folders.AnalyticViewGroups.first("Title.Primary='"& arrcfg(0).Item("AnalyticViewGroup") &"'")
										End If
									End If
								End If
							Next
							
							For each strKey in arrcfg(1).Keys
								If System.folders.AnalyticViewAttributes.Exists("ID='"& strKey &"'") Then
									If Not objAnalytic.folders.Attributes.Exists("AnalyticViewAttribute.ID='"& strKey &"'") Then
										if arrcfg(1).Item(strKey) <> "" Then
											set objAttribute = objAnalytic.folders.Attributes.CreateNewInstance(,,arrcfg(1).Item(strKey))
										Else
											set objAttribute = objAnalytic.folders.Attributes.CreateNewInstance
										End IF
										objAttribute.AnalyticViewAttribute = System.folders.AnalyticViewAttributes.first("ID='"& strKey &"'")
										objAttribute.save objTransaction
									End If							
								End If
							Next
							
							For each strKey in arrcfg(2).Keys
								If System.Folders.DrillHierarchies.Exists("Title.Primary='"& strKey &"'") Then
									If Not objAnalytic.folders.DrillHierarchies.Exists("DrillHierarchy.Title.Primary='"& strKey &"'") Then
										if arrcfg(1).Item(strKey) <> "" Then
											set objAttribute = objAnalytic.folders.DrillHierarchies.CreateNewInstance(,,arrcfg(1).Item(strKey))
										Else
											set objAttribute = objAnalytic.folders.DrillHierarchies.CreateNewInstance
										End IF
										objAttribute.DrillHierarchy = System.Folders.DrillHierarchies.first("Title.Primary='"& strKey &"'")
										objAttribute.save objTransaction
									End If
								End If
							Next
							AddUserGroups objAnalytic,arrcfg(3),objTransaction
							
							If objFSO.FileExists(strDirectoryPath & "DEFINITION\" & strFileName) Then
								set objFileToRead = objFSO.OpenTextFile(strDirectoryPath & "DEFINITION\" & strFileName,1)
								strFileText = objFileToRead.ReadAll()
								if Len(strFileText & "")>0 Then
									objAnalytic.Definition = Trim(strFileText)
									objAnalytic.DefinitionW = Trim(strFileText)
								End If
							End If
							
							'If this is an analytic view template then check for publish
							Dim IsAnalyticTemplate : IsAnalyticTemplate = System.Nvl(objAnalytic.AnalyticViewTemplate,False)
							Dim isChildVersionHigher : isChildVersionHigher = False
							If IsAnalyticTemplate Then
								Dim newVersion: newVersion = 0
								isChildVersionHigher = checkAnalyticChildVersion(objAnalytic, newVersion)
								If isChildVersionHigher Then
									objAnalytic.Published.Version = newVersion
								End If
							End If
							
							objAnalytic.save objTransaction
							set objVal = objTransaction.Validate
							if objval.status <> 3 then
								intAddedAnalytics = intAddedAnalytics + 1
								objTransaction.Commit
								objSystemLog.LogNormalDetail "The '"& strAnalyticKey &"' analytic view has been created/updated."
								If IsAnalyticTemplate Then
									objAnalytic.Publish()
								End If
							else
								intErroredAnalytics = intErroredAnalytics + 1
								objSystemLog.LogErrorDetail "","","There were some error in updating/creating the '"& strAnalyticKey &"' analytic view : '" & objval.Result.Message & "'" 
							end if
						End If
					Else
						intErroredAnalytics = intErroredAnalytics + 1
					End If
				End If
			Else
				intErroredAnalytics = intErroredAnalytics + 1
			End If
		End If
	Next
	objSystemLog.LogNormalDetail "The Analytic Views loading process completed."
End Function

'ppatel add or update all class members
Function Loadimports(System,scriptdir,objSystemLog)
		strFolder = scriptdir & "Inbox"
		call Createrequiredfolders(scriptdir)
		if objFSO.FolderExists(strFolder) then
			Set objDirectory = Nothing
			Set objDirectory = objFSO.GetFolder(strFolder)
			For Each objFile In objDirectory.Files
				blnUserDefinedKey = false
				If Not objDirectory Is Nothing Then			
					If Not objFile Is Nothing Then
						If LCase(objFSO.GetExtensionName(objFile.Name)) = "csv" Then
							dim TLFname,BlnCancreatenew,KeyMappingMembers,SubFolder,ParentKeyMapping,CastDownclassname
							'TLFname = left(objFile.Name,len(objFile.Name) - 4)
							objSystemLog.LogNormalDetail "Process started for file '" & objFile.Name & "'."
							strErrLogFileName = scriptdir & "Error\"&  left(objFile.Name,len(objFile.Name) - 4) &"-ErrorLogFile-" & System.name &".txt"
							If objFSO.FileExists(strErrLogFileName) Then
								objFSO.DeleteFile(strErrLogFileName)
							End If
							Set objOutErrFile = objFSO.CreateTextFile(strErrLogFileName,True)
						
							'TLFname
							set strData1 = objFile.OpenAsTextStream(1,0)
							strData = strData1.ReadAll
							strData1.close
							arrLines = Split(strData,vbCrLf)
							LineCount = UBound(arrLines)
							Dim Lineno
							Lineno = 0
							BlnSubFolder = False
							for i=0 to LineCount
								if Left(arrLines(i), 1) = "#" then
									Lineno = Lineno + 1
									arrayconfig = Split(arrLines(i),":~")
									if UBound(arrayconfig) = 1 then											
										select Case Ucase(arrayconfig(0))
											Case UCase("#ToplevelFolder")
												TLFname = arrayconfig(1)
											Case UCase("#CanCreateNew")
												BlnCancreatenew = arrayconfig(1)
											Case UCase("#KeyMapping")
												KeyMappingMembers = arrayconfig(1)
											Case Ucase("#SubFolder")
												SubFolder = arrayconfig(1)
												if SubFolder<>"" then
													BlnSubFolder = True
												end if
											Case Ucase("#ParentKeyMapping")
												ParentKeyMapping = arrayconfig(1)
											Case Ucase("#CastDownclassname")
												CastDownclassname = arrayconfig(1)
										End Select
									end if
								else
									Exit For
								end if
							next
							if TLFname<> "" and BlnCancreatenew<>"" and KeyMappingMembers<>"" then
								
								Set objFolder = System.Folders(TLFname)
								If Not objFolder Is Nothing Then
									if CastDownclassname="" then
										Set objCreateClass = objFolder.TargetClass
									else
										set objCreateClass = System.Classes(CastDownclassname)
									end if
									if LineCount> Lineno then
										Dim blnStatus
										blnStatus = True
										if BlnSubFolder = True then
											if ParentKeyMapping<>"" then
												ArrayParentKeyMapping = readLine(ParentKeyMapping)
												if not  ValidateMembers(ArrayParentKeyMapping,objCreateClass) then
													blnStatus = False
												end if
											else
												'objOutErrFile.WriteLine("TLF '"& TLFname &"' dose not Exist in system")
												objSystemLog.LogNormalDetail "TLF '"& TLFname &"' dose not Exist in system."
												blnStatus = False
											end if
										end if
										if blnStatus = True then
											StrFirstline = arrLines(Lineno)
											ArrayKeyMappingMembers = readLine(KeyMappingMembers)
											ArrayDataMembers = readLine(StrFirstline)
											
											if BlnSubFolder = True then
												if ValidateMembers(ArrayParentKeyMapping,objCreateClass) then
													ArrayParentKeyMappingMemberswithindex = fieldIndex(ArrayParentKeyMapping,ArrayDataMembers,blnStatus,"ParentKeyMapping")
												else
													blnStatus = False
												end if
											end if	
											if blnStatus = True then 
												Lineno = Lineno + 1
												dim BlnValidateDataMembers,BlnValidatekeyMembers
												for i=Lineno to LineCount
													strdataline = arrLines(i)
													if strdataline<>"" then
														ArrayMemberValues = readLine(strdataline)
														if UBound(ArrayDataMembers) = UBound(ArrayMemberValues) then
															if BlnSubFolder = True then
																strFetchParentFilter = ""
																strFetchParentFilter = GetInstanceFilter(ArrayParentKeyMappingMemberswithindex, ArrayMemberValues ,BlnStatus)
																if BlnStatus = True and strFetchParentFilter <> "" then
																	On Error Resume next
																	Dim objInst
																	set objInst = Nothing 
																	'Set tmpFS = objFolder.Scan(strFetchParentFilter, , 1)
																	Set tmpFS = objFolder.First(strFetchParentFilter)
																	If Err <> 0 Then
																		If Trim(Err.Description) = "" Then
																			'objOutErrFile.WriteLine(strdataline &"Parent Key Filter error in expression: " & strFetchParentFilter)
																			
																			objSystemLog.LogNormalDetail strdataline &"Parent Key Filter error in expression: " & strFetchParentFilter
																		Else
																			'objOutErrFile.WriteLine(strdataline &"Parent Key Filter " & Err.Description)
																			
																			objSystemLog.LogNormalDetail "See error log for data line '" & strdataline & "'"
																			objSystemLog.LogErrorDetail "","",strdataline & "Parent Key Filter " & Err.Description
																		End If
																		Exit Function
																	Else
																		On Error Goto 0
																		'If Not tmpFS.EndOfScan Then
																		If Not tmpFS is nothing and not tmpFS.isnull Then
																			Set objInst = tmpFS
																		End If
																	End If
																	On Error Goto 0
																	If not objInst Is Nothing Then
																		set objsubfolder = Nothing
																		set objsubfolder = objInst.Folders(SubFolder)
																		if not objsubfolder is nothing then
																			set objsubfolderClass = objsubfolder.TargetClass
																			if isempty(BlnValidateDataMembers) or BlnValidateDataMembers="" then
																				if ValidateMembers(ArrayDataMembers,objsubfolderClass) then
																					BlnValidateDataMembers = True
																				else
																					BlnValidateDataMembers = False
																				end if
																			end if
																			if isempty(BlnValidatekeyMembers) or BlnValidatekeyMembers="" then
																				if ValidateMembers(ArrayKeyMappingMembers,objsubfolderClass) then
																					ArrayKeyMappingMemberswithindex = fieldIndex(ArrayKeyMappingMembers,ArrayDataMembers,blnStatus,"KeyMapping")
																					BlnValidatekeyMembers=True
																				else
																					BlnValidatekeyMembers=False
																				end if
																			end if
																			if blnStatus = True then
																				if BlnValidatekeyMembers=True and BlnValidateDataMembers=True then
																					call ImportCSVs(ArrayKeyMappingMemberswithindex,ArrayDataMembers,ArrayMemberValues,objsubfolder,BlnCancreatenew,CastDownclassname,blnStatus)
																					Erase ArrayKeyMappingMemberswithindex
																				end if
																			else
																				'objOutErrFile.WriteLine("Not all Key Mapping Members Exist in found in File")
																				objSystemLog.LogNormalDetail "Not all Key Mapping Members Exist in found in File"
																			end if
																		else
																			'objOutErrFile.WriteLine(strdataline &"No SubFolder with name " &  SubFolder &" found in TLF "& TLFname)
																			objSystemLog.LogNormalDetail strdataline &" No SubFolder with name " &  SubFolder &" found in TLF "& TLFname
																		end if
																	Else
																		'objOutErrFile.WriteLine(strdataline &"No Record found in TLF "& TLFname &" Parent Key Filter " & Err.Description)
																		objSystemLog.LogNormalDetail strdataline &" No Record found in TLF "& TLFname &" Parent Key Filter " & Err.Description
																	end if
																end if 
															else
																if isempty(BlnValidateDataMembers) or BlnValidateDataMembers="" then
																	if ValidateMembers(ArrayDataMembers,objCreateClass) then
																		BlnValidateDataMembers = True
																	else
																		BlnValidateDataMembers = False
																	end if
																end if
																if isempty(BlnValidatekeyMembers) or BlnValidatekeyMembers="" then
																	if ValidateMembers(ArrayKeyMappingMembers,objCreateClass) then
																		ArrayKeyMappingMemberswithindex = fieldIndex(ArrayKeyMappingMembers,ArrayDataMembers,blnStatus,"KeyMapping")
																		BlnValidatekeyMembers=True
																	else
																		BlnValidatekeyMembers=False
																	end if
																end if
																if blnStatus = True then
																	if BlnValidatekeyMembers=True and BlnValidateDataMembers=True then
																		call ImportCSVs(ArrayKeyMappingMemberswithindex,ArrayDataMembers,ArrayMemberValues,objFolder,BlnCancreatenew,CastDownclassname,blnStatus)
																		Erase ArrayKeyMappingMemberswithindex
																	end if
																else
																	'objOutErrFile.WriteLine("Not all Key Mapping Members Exist in found in File")
																	objSystemLog.LogNormalDetail "Not all Key Mapping Members Exist in found in File"
																end if
															end if
															BlnValidateDataMembers = ""
															BlnValidatekeyMembers = ""
														else
															'objOutErrFile.WriteLine("Not all members values are available in line" &strdataline)
															objSystemLog.LogNormalDetail "Not all members values are available in line" &strdataline
														end if
														Erase ArrayMemberValues
														strdataline = ""
													end if
												next
											else
												'objOutErrFile.WriteLine("Not all Parent Key Mapping Members Exist in found in File")
												objSystemLog.LogNormalDetail "Not all Parent Key Mapping Members Exist in found in File"
											End if
											
											Erase ArrayDataMembers
											Erase ArrayKeyMappingMembers
										end if
									else
										'objOutErrFile.WriteLine("No Data Exist to import in File")
										objSystemLog.LogNormalDetail "No Data Exist to import in File"
									end if
									set objFolder = Nothing
								else
									'objOutErrFile.WriteLine("TLF '"& TLFname &"' dose not Exist in system")
									objSystemLog.LogNormalDetail "TLF '"& TLFname &"' dose not Exist in system"
								end if
								TLFname = "" 
								BlnCancreatenew = "" 
								KeyMappingMembers = "" 
								SubFolder = ""
								ParentKeyMapping = ""
								CastDownclassname =""
							else
								objOutErrFile.WriteLine("Please make sure file must have Config columns 'ToplevelFolder','CanCreateNew','KeyMapping'")
								objSystemLog.LogNormalDetail "Please make sure file must have Config columns 'ToplevelFolder','CanCreateNew','KeyMapping'"
							end if
							objOutErrFile.close
							objSystemLog.LogNormalDetail "Process completed for file '" & objFile.Name & "'."
							set oOutErrFile= objFSO.GetFile(strErrLogFileName)
							
							If not oOutErrFile.Size > 0 Then
								oOutErrFile.delete
								strcompletefilepath = scriptdir &"Completed\"
								strfilepath = strcompletefilepath   &  left(objFile.Name,len(objFile.Name) - 4) & "-" & System.name& "." & objFSO.GetExtensionName(objFile.Name)
								If objFSO.FileExists(strfilepath) Then
									objFSO.DeleteFile strfilepath
								end if
								objFSO.CopyFile  objFile.Path ,strfilepath
								strcompletefilepath = ""
								strfilepath = ""
							else
								strerrorfilepath = scriptdir &"Error\"
								strfilepath = strerrorfilepath  & objFile.Name
								If objFSO.FileExists(strfilepath) Then
									objFSO.DeleteFile strfilepath
								end if
								objFSO.CopyFile objFile.Path ,strfilepath
								strerrorfilepath = ""
								strfilepath = ""
							end if
							set oOutErrFile = Nothing
							set objOutErrFile = Nothing
							set objFolder = Nothing
							BlnCancreatenew = ""
						end if
					end if
				end if
			next
		else
			objSystemLog.LogNormalDetail "Folder with name " & strFolder &" not Exits."
		end if
	
End Function

Function readLine(ByRef strLine)
		Dim strTmpLine
		Dim aryCols
		Dim aryOutCols
		Dim strResult
		Dim blnHasQuote
		Dim i
		Dim iIdx	:iIdx = 0
		mstrDelimiter = "|"
		mreadCount = mreadCount + Len(strLine)
		aryCols = Split(strLine, mstrDelimiter)

		ReDim aryOutCols(Ubound(aryCols))

		i = 0
		while i <= Ubound(aryCols)
			If Mid(aryCols(i), 1, 1) = """" Or blnHasQuote Then
				If Not blnHasQuote Then
					strResult = Mid(aryCols(i), 2)
					blnHasQuote = True
				Else
					strResult = strResult & aryCols(i)
				 End If
				If Ucase(Mid(strResult, 1, Len(strResult) - 1)) = UCase(NULL_VALUE) Then
					aryOutCols(iIdx) = Null
				Else
					aryOutCols(iIdx) = Replace(trim(Mid(strResult, 1, Len(strResult) - 1)),vbTab,"")
				End If
				iIdx = iIdx + 1
				blnHasQuote = False
			Else
				If UCase(aryCols(i)) = UCase(NULL_VALUE) Then
					aryOutCols(iIdx) = Null
				Else
					aryOutCols(iIdx) = Replace(trim(aryCols(i)),vbTab,"")
				End If
				iIdx = iIdx + 1
			End If
			i = i + 1
		Wend
		If UBound(aryOutCols) > iIdx - 1 Then
			ReDim Preserve aryOutCols(iIdx - 1)
		End If
		readLine = aryOutCols
End Function

Function fieldIndex(ArrayKeyMappingMembers,ArrayDataMembers,ByRef blnStatus,StrKeyMappingtype)
	dim ArrayKeyMappingMemberswithindex
	ReDim ArrayKeyMappingMemberswithindex(Ubound(ArrayKeyMappingMembers),1)
	Dim i,j
	'fieldIndex = -1
	For i = 0 To Ubound(ArrayKeyMappingMembers)
		if blnStatus = True then
			blnStatus = False
			if Right(ArrayKeyMappingMembers(i), Len(ArrayKeyMappingMembers(i)) - (Len(ArrayKeyMappingMembers(i))-1)) = "*" then
				strKeyMappingMember = Left(ArrayKeyMappingMembers(i),Len(ArrayKeyMappingMembers(i))-1)
			else
				strKeyMappingMember =  Trim(UCase(ArrayKeyMappingMembers(i)))
			end if
			For j = 0 To Ubound(ArrayDataMembers)
				If Right(ArrayDataMembers(j), Len(ArrayDataMembers(j)) - (Len(ArrayDataMembers(j))-1)) = "*" and StrKeyMappingtype = "ParentKeyMapping" Then
					strArrayDataMember = Left(ArrayDataMembers(j),Len(ArrayDataMembers(j))-1)
				Else
					strArrayDataMember = ArrayDataMembers(j)
				End IF
				
				If Trim(UCase(strArrayDataMember)) = Trim(UCase(strKeyMappingMember)) Then
					ArrayKeyMappingMemberswithindex(i,0)=ArrayKeyMappingMembers(i)
					ArrayKeyMappingMemberswithindex(i,1) = j
					blnStatus = True
					Exit for
				End If
			next
		else
			Exit For
		End if
		strKeyMappingMember = ""
	next
	fieldIndex = ArrayKeyMappingMemberswithindex
End Function

Function ImportCSVs(ArrayKeyMappingMemberswithindex,ArrayDataMembers,ArrayMemberValues,objFolder,BlnCancreatenew,strCastDownclassname,ByRef blnStatus)
	Dim strFetchFilter,objInst
	set objInst = Nothing
	dim blndataerror
	blndataerror =False
	strFetchFilter = GetInstanceFilter(ArrayKeyMappingMemberswithindex, ArrayMemberValues ,BlnStatus)
	set objCastDownclassname = System.Classes(strCastDownclassname)
	if BlnStatus = True and strFetchFilter <> "" then
		'objFolder.E
		On Error Resume next
		if objFolder.count(strFetchFilter) > 1 then 
			objOutErrFile.WriteLine(strdataline &"there are more then one record in system with filter " & strFetchFilter &"please update KeyMapping to get unique record")
			objSystemLog.LogNormalDetail "See error log for data line '" & strdataline & "'"
			objSystemLog.LogErrorDetail "","",strdataline &"there are more then one record in system with filter " & strFetchFilter &"please update KeyMapping to get unique record"
			Exit Function
		End If
		
		Set tmpFS = objFolder.Scan(strFetchFilter, , 1)
		If Err <> 0 Then
			If Trim(Err.Description) = "" Then
				'objOutErrFile.WriteLine(strdataline & " Key Filter error in expression: " & strFetchFilter)
				
				objSystemLog.LogNormalDetail strdataline & " Key Filter error in expression: " & strFetchFilter
			Else
				'objOutErrFile.WriteLine(strdataline &"Key Filter " & Err.Description)
				
				objSystemLog.LogNormalDetail "See error log for data line '" & strdataline & "'"
				objSystemLog.LogErrorDetail "","",strdataline &" Key Filter " & Err.Description
			End If
			Exit Function
		Else
			On Error Goto 0
			If Not tmpFS.EndOfScan Then
				Set objInst = tmpFS.Fetch
			End If
		End If
		On Error Goto 0
		Set objTxn = System.BeginTransaction
		If Trim(Ucase(ArrayDataMembers(Ubound(ArrayDataMembers)))) = "ISDELETED" and  Trim(Ucase(ArrayMemberValues(Ubound(ArrayMemberValues)))) = "TRUE"Then
			if not objInst Is Nothing then
				objInst.delete objTxn
				call CommitTransaction(objTxn,System)
			end if
			Exit Function
		end if
		If objInst Is Nothing Then
			if BlnCancreatenew then
				if objFolder.CanCreateNewInstance then
					blnNewInstance = true
					if blnUserDefinedKey and strKeyValue <> "" Then
						if strCastDownclassname <> "" then
							set objInst = objFolder.CreateNewInstance(objCastDownclassname,,strKeyValue)
						else
							set objInst = objFolder.CreateNewInstance(,,strKeyValue)
						end if
						
					Else
						if strCastDownclassname <> "" then
							set objInst = objFolder.CreateNewInstance(objCastDownclassname)
						else
							set objInst = objFolder.CreateNewInstance()
						end if
					End If
				else
					'objOutErrFile.WriteLine("System not allow to create new instance in TLF: " & objFolder.name)
					objSystemLog.LogNormalDetail "System not allow to create new instance in TLF: " & objFolder.name
				end if
			else
				'objOutErrFile.WriteLine("Configuration Error. No suitable to create new instance in TLF: " & objFolder.name)
				objSystemLog.LogNormalDetail "Configuration Error. No suitable to create new instance in TLF: " & objFolder.name
				exit Function
			end if
		end if
		For i = 0 To Ubound(ArrayDataMembers)
			if BlnStatus = True and blndataerror= False then
				if Trim(Ucase(ArrayDataMembers(i))) <> "ISDELETED" AND Trim(Ucase(ArrayDataMembers(i))) <> "KEY" and ArrayMemberValues(i)<>"" Then
					if blnNewInstance or not checkvalueexistsinKeyMapping(ArrayKeyMappingMemberswithindex,ArrayDataMembers(i)) then
						call Updatemembervalue(objInst,ArrayDataMembers(i),ArrayMemberValues(i),BlnStatus,blndataerror)
					end if
				end if	
			else
				Exit function
			end if
		next
		if BlnStatus = True and blndataerror= False then
			objInst.save objTxn
		end if
		call CommitTransaction(objTxn,System)
	end if
End Function

Function Updatemembervalue(objInst,StrMembers,StrMemberValues,ByRef BlnStatus,ByRef blndataerror)
	on error resume next
	if StrMemberValues<>"" and not Right(StrMembers, Len(StrMembers) - (Len(StrMembers)-1)) = "*" then
		if ucase(right(StrMembers , 5) )= "VALUE" Then
			StrMembers = left(StrMembers , InStrRev (StrMembers, ".")-1) 
		end if
		i = InStr(StrMembers, ".")
		if i > 0 then
			Set objMember = objInst.members(Trim(Left(StrMembers, i - 1)))
			if objMember.definition.membertype = 3 then  
				StrFilter =  Mid(StrMembers, i + 1) &" = '" &StrMemberValues &"'"
				set objjoin = objMember.targetfolder.first(StrFilter)
				if not objjoin.isnull then
					'call Updatevalue (objMember,objjoin)
					set objInst.members(Trim(Left(StrMembers, i - 1))).value = objjoin
				else
					objOutErrFile.WriteLine(strdataline & " for set join member " & objMember.definition.title.value &" Not able to find the record  " &StrFilter & "in Folder " & objMember.targetfolder.title.value)	
					objSystemLog.LogNormalDetail strdataline & " for set join member " & objMember.definition.title.value &" Not able to find the record  " &StrFilter & "in Folder " & objMember.targetfolder.title.value
					blndataerror = True
				end if
				set objjoin = nothing
				Exit Function
			elseif objMember.definition.membertype = 2 and instr(i+1,StrMembers,".") > 0  then  
				set objMember2 = nothing
				Set objMember2 = objMember.members(mid(StrMembers,i+1,instr(i+1,StrMembers,".") - i-1))
				if not objMember2 is nothing then
					if objMember2.definition.membertype = 3 then  
						StrFilter =  Mid(StrMembers, instr(i+1,StrMembers,".") + 1) &" = '" &StrMemberValues &"'"
					set objjoin = objMember2.targetfolder.first(StrFilter)
					if not objjoin.isnull then
						'call Updatevalue (objMember,objjoin)
						set objInst.members(left(StrMembers,instr(i+1,StrMembers,".")-1)).value = objjoin
					else
						objOutErrFile.WriteLine(strdataline & " for set join member " & objMember2.definition.title.value &" Not able to find the record  " &StrFilter & "in Folder " & objMember2.targetfolder.title.value)	
						objSystemLog.LogNormalDetail strdataline & " for set join member " & objMember2.definition.title.value &" Not able to find the record  " &StrFilter & "in Folder " & objMember2.targetfolder.title.value
						blndataerror = True
					end if
					set objjoin = nothing
					Exit Function
					end if
				end if
			end if
		end if
	
		Set objMember = objInst.members(StrMembers)
		Select Case objMember.definition.membertype
			Case 0  ' Data
				objInst.members(StrMembers) = StrMemberValues
			Case 1 ' Enummeration
				on error resume next
					objInst.members(StrMembers).value = StrMemberValues
					If Err <> 0 Then
						objSystemLog.LogNormalDetail "There were some error occurred while setting the value to the member '" & StrMembers & "'."
						objSystemLog.LogErrorDetail "","",StrMembers & " - Error: " & Err.Description
						objOutErrFile.WriteLine("There were some error occurred while setting the value to the member '" & StrMembers & "'." & " - Error: " & Err.Description)
						blndataerror = True
					End IF
				on error goto 0
			Case 2  ' Class
				objOutErrFile.WriteLine("Please Provide class member's 'Members' value Not able to set class members ")
				BlnStatus = False
			Case 3  ' join
				objOutErrFile.WriteLine("Please Provide join member 'Key Mapping Members' value for set join member ")
				objSystemLog.LogNormalDetail "Please Provide join member 'Key Mapping Members' value for set join member "
				BlnStatus = False	
		End Select
	end if
	If Err <> 0 Then
		objSystemLog.LogNormalDetail "There were some error occurred while setting the value to the member '" & StrMembers & "'."
		objSystemLog.LogErrorDetail "","",StrMembers & " - Error: " & Err.Description
		'objOutErrFile.WriteLine("There were some error occurred while setting the value to the member '" & StrMembers & "'." & " - Error: " & Err.Description)
	End IF
	on error goto 0
End Function

Function GetInstanceFilter(ArrayKeyMappingMemberswithindex, ArrayMemberValues ,ByRef BlnStatus)
	Dim strFetchFilter,Blnoptional
	strKeyValue = ""
	For i = 0 To Ubound(ArrayKeyMappingMemberswithindex)
		Blnoptional = False
		If (UCase(ArrayKeyMappingMemberswithindex(i,0)) = "KEY") Then
			blnUserDefinedKey = true
			strKeyValue = ArrayMemberValues(ArrayKeyMappingMemberswithindex(i,1))
		End If
		strMemberValues = ArrayMemberValues(ArrayKeyMappingMemberswithindex(i,1))	
		if Right(ArrayKeyMappingMemberswithindex(i,0), Len(ArrayKeyMappingMemberswithindex(i,0)) - (Len(ArrayKeyMappingMemberswithindex(i,0))-1)) = "*" then
			strKeyMappingMember = Left(ArrayKeyMappingMemberswithindex(i,0),Len(ArrayKeyMappingMemberswithindex(i,0))-1)
			Blnoptional = True
		else
			strKeyMappingMember = ArrayKeyMappingMemberswithindex(i,0)
		end if
		if strMemberValues <> "" then
			StrFilter = strKeyMappingMember & "=""" & Replace(strMemberValues, """", """""") & """"
			If strFetchFilter <> "" Then
				strFetchFilter = strFetchFilter & " and "
			End If
			strFetchFilter = strFetchFilter & StrFilter
		else
			if Blnoptional = False then
				BlnStatus = False
				objOutErrFile.WriteLine(strdataline & """Key Mapping Members value must not be null""")
				objSystemLog.LogNormalDetail strdataline & """Key Mapping Members value must not be null"""
			end if
		end if
		strKeyMappingMember = ""
	next
	GetInstanceFilter = strFetchFilter
End Function

Function ValidateMembers(arrMembers, objClassDef)
	Dim i, objFilter, objMember, objTarget,BlnValid
	BlnValid = True
	For i = 0 To Ubound(arrMembers)
		if not isnull(arrMembers(i)) then
			If arrMembers(i) <> "" and Ucase(trim(arrMembers(i))) <> "ISDELETED" and not Right(arrMembers(i), Len(arrMembers(i)) - (Len(arrMembers(i))-1)) = "*" Then 
				'Set objMember = GetMemberDef(objClassDef, arrMembers(i), objTarget)
				'If objMember Is Nothing Then
					'BlnValid = False
					'objOutErrFile.WriteLine("Member with name  " & arrMembers(i) & " not found")
				'End If
				strErrors = ""
				Set objExpression = CreateObject("ActivElk.Filter")
				If Not objExpression.Parse(arrMembers(i),objClassDef, strErrors) Then		
					BlnValid = False
					'objOutErrFile.WriteLine("Error while Evaluate Member " & arrMembers(i) & " Expression Error: " & strErrors)
					objSystemLog.LogErrorDetail  "","", "Error while Evaluate Member " & arrMembers(i) & " Expression Error: " & strErrors
				End If
				set objExpression = Nothing
			End If
		End If
	Next
	
	ValidateMembers = BlnValid
End Function

Function CommitTransaction(objTxn,objSystem)
	Dim objval
	Set objval = objTxn.validate
	If objval.status <> 3 then
		objTxn.Commit
	else
		'StrSystemNameError = StrSystemNameError & vbNewLine & "Found error to disable IC_VERIFYPHOTOS workflow - " & objval.Result.Message & " In " & objSystem.name
		objOutErrFile.WriteLine(strdataline & " Error : " & objval.Result.Message)
		objSystemLog.LogNormalDetail "See error log for data line '" & strdataline & "'"
		objSystemLog.LogErrorDetail "","",strdataline & " Error : " & objval.Result.Message
	End if
	set objTxn = Nothing
End Function

Function Createrequiredfolders(scriptdir)
	StrCompleted = scriptdir & "Completed"
	StrError = scriptdir &"Error"
	If Not objFSO.FolderExists(StrCompleted) Then
		objFSO.CreateFolder (StrCompleted)
	End If
	If Not objFSO.FolderExists(StrError) Then
		objFSO.CreateFolder (StrError)
	End If
End Function

Function AddEndSlashIfNecessary(strFolderPath)
	strFolderPath = Trim(strFolderPath)
	If InStr(strFolderPath,"/") Then
		strSlash = "/"
	Else
		strSlash = "\"
	End If
	If (Right(strFolderPath,1)<>strSlash) Then
		strFolderPath = strFolderPath & strSlash
	End If
	AddEndSlashIfNecessary = strFolderPath
End Function


Function AddDistributions(objInst,objDict,objTransaction)
	For Each strKey in objDict.Keys
		strvalues = objDict.Item(strKey)
		arrParts = Split(strvalues,"|")
		strKeyAccount = arrParts(0)
		strOrganization = arrParts(1)
		strRole = arrParts(2)
		strSalesTeam = arrParts(3)
		strTerritory = arrParts(4)
		strfilter = ""
		if strKeyAccount <> "" then
			strfilter = "KeyAccount.Name = '" & strKeyAccount &"'"
		end if
		if strOrganization <> "" then
			if strfilter = "" then
				strfilter = "Organization.Name = '" & strOrganization &"'"
			else
				strfilter = strfilter &"and Organization.Name = '" & strOrganization &"'"
			end if
		end if
		if strRole <> "" then
			if strfilter = "" then
				strfilter = "Role.Code = '" & strRole &"'"
			else
				strfilter = strfilter & "and Role.Code = '" & strRole &"'"
			end if
		end if
		if strSalesTeam <> "" then
			if strfilter = "" then
				strfilter = "SalesTeam.Name = '" & strSalesTeam &"'"
			else
				strfilter = strfilter &"and SalesTeam.Name = '" & strSalesTeam &"'"
			end if
		end if
		if strTerritory <> "" then
			if strfilter = "" then
				strfilter = "Territory.TerritoryName = '" & strTerritory &"'"
			else
				strfilter =strfilter & "and Territory.TerritoryName = '" & strTerritory &"'"
			end if
		end if
		if strfilter<>"" then
			strcount = objInst.Folders("AppliesTo").count(strfilter)
			if strcount = 1 then
				set objDistribution =  objInst.Folders("AppliesTo").first(strfilter)
			elseif strcount<= 0 then
				set objDistribution =  objInst.Folders("AppliesTo").CreateNewInstance()
			end if
			If Not objDistribution.IsNull Then
				if strKeyAccount <> "" then
					if System.folders.BannerGroups.Exists("UCase(name) ='"&  UCase(strKeyAccount) &"'") then
						set objDistribution.KeyAccount = System.folders.BannerGroups.First("UCase(name) ='"&  UCase(strKeyAccount) &"'")
					end if
				end if
				if strOrganization <> "" then
					if System.folders.Stores.Exists("UCase(name) = '"&  UCase(strOrganization) &"'") then
						set objDistribution.Organization = System.folders.Stores.First("UCase(name) = '"&  UCase(strOrganization) &"'")
					end if
				end if
				if strRole <> "" then
					if System.folders.Roles.Exists("UCase(Code) = '"&  UCase(strRole) &"'") then
						set objDistribution.Role = System.folders.Roles.First("UCase(Code) = '"&  UCase(strRole) &"'")
					end if
				end if
				if strSalesTeam <> "" then
					if System.folders.SalesTeams.Exists("UCase(name) = '"&  UCase(strSalesTeam) &"'") then
						set objDistribution.SalesTeam = System.folders.SalesTeams.First("UCase(name) = '"&  UCase(strSalesTeam) &"'")
					end if
				end if
				if strTerritory <> "" then
					if System.folders.Territories.Exists("UCase(TerritoryName) = '"&  UCase(strTerritory) &"'") then
						set objDistribution.Territory = System.folders.Territories.First("UCase(TerritoryName) = '"&  UCase(strTerritory) &"'")
					end if
				end if
				objDistribution.Save objTransaction
			End If
		End If
	Next
End Function

Function AddTeams(objInst,objDict,objTransaction)
	For Each strKey in objDict.Keys
		strEnabled = objDict.Item(strKey)
		arrParts = Split(strEnabled,"|")
		Set oAccess = objInst.Folders("Teams").First("Team.name='" & strKey &"'")
		if Ucase(arrParts(0)) = "TRUE" then
			If oAccess.IsNull Then
				Set ObjTeam = System.Folders("Teams").First("name='"& strKey &"'")
				If UBound(arrParts)>0 Then
					Set oAccess = objInst.Folders("Teams").CreateNewInstance(,,arrParts(1))
				Else
					Set oAccess = objInst.Folders("Teams").CreateNewInstance()
				End If
				Set oAccess.Team = ObjTeam
				Set oAccess.PR_Report = objInst
			End If
			If Not oAccess.IsNull Then
				oAccess.IsEnabled = strEnabled
				oAccess.Save objTransaction
			End If
		else
			If Not oAccess.IsNull Then 
				oAccess.IsEnabled = strEnabled
				oAccess.Save objTransaction
			end if
		end if
	Next
End Function

Function AddUserGroups(objInst,objDict,objTransaction)
	For Each strKey in objDict.Keys
		strPermission = objDict.Item(strKey)
		arrParts = Split(strPermission,"|")
		Set oAccess = objInst.Folders("Groups").First("Group.GroupCode='" & strKey &"'")
		If oAccess.IsNull Then
			Set ObjUserGroup = System.Folders("UserGroups").First("GroupCode='"& strKey &"'")
			If UBound(arrParts)>0 Then
		 		Set oAccess = objInst.Folders("Groups").CreateNewInstance(,,arrParts(1))
			Else
				Set oAccess = objInst.Folders("Groups").CreateNewInstance()
			End If
			Set oAccess.Group = ObjUserGroup
		End If
		If Not oAccess.IsNull Then
			oAccess.Permission.Value = arrParts(0)
			oAccess.Save objTransaction
		End If
	Next
End Function

Function Convertexceltocsv(strImportFolder)
	StrArchive = strImportFolder & "Archive\"
	If Not objFSO.FolderExists(StrArchive) Then
		objFSO.CreateFolder (StrArchive)
	End If
	For Each oFile In objFSO.GetFolder(strImportFolder).Files
		If UCase(objFSO.GetExtensionName(oFile.Name)) = "XLS" Or  UCase(objFSO.GetExtensionName(oFile.Name)) = "XLSX" Then
			strFileContent = ""
			ReadExcelFile(oFile.Path)
			
			strfilepath = StrArchive  & oFile.Name
			If objFSO.FileExists(strfilepath) Then
				objFSO.DeleteFile strfilepath
			end if
			objFSO.MoveFile oFile.Path ,StrArchive
			strfilepath = ""
		End If
	Next
End Function

Function ReadExcelFile(ByVal strFile)
  ' Local variable declarations
  Dim objExcel, objSheet, objCells
  Dim nUsedRows, nUsedCols, nTop, nLeft, nRow, nCol
  

  ' Default return value
  ReadExcelFile = Null

  ' Create the Excel object
  On Error Resume Next
  Set objExcel = CreateObject("Excel.Application")
  If (Err.Number <> 0) Then
    Exit Function
  End If

  ' Don't display any alert messages
  objExcel.DisplayAlerts = 0  

  ' Open the document as read-only
  On Error Resume Next
  Call objExcel.Workbooks.Open(strFile, False, True)
  If (Err.Number <> 0) Then
    Exit Function
  End If

  ' If you wanted to read all sheets, you could call
  ' objExcel.Worksheets.Count to get the number of sheets
  ' and the loop through each one. But in this example, we
  ' will just read the first sheet.
	 ' Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)
	for each objSheet in objExcel.Worksheets
		 ' Dim arrSheet()
		  ' Get the number of used rows
		  nUsedRows = objSheet.UsedRange.Rows.Count

		  ' Get the number of used columns
		  nUsedCols = objSheet.UsedRange.Columns.Count

		  ' Get the topmost row that has data
		  nTop = objSheet.UsedRange.Row

		  ' Get leftmost column that has data
		  nLeft = objSheet.UsedRange.Column

		  ' Get the used cells
		  Set objCells = objSheet.Cells

		  ' Dimension the sheet array
		  'ReDim arrSheet(nUsedRows - 1, nUsedCols - 1)
			

		  ' Loop through each row
		  For nRow = 0 To (nUsedRows - 1)
			' Loop through each column
			For nCol = 0 To (nUsedCols - 1)
				' Add the cell value to the sheet array
				if nCol = 0 then
					strFileContent = strFileContent  & objCells(nRow + nTop, nCol + nLeft).Value
				else
					if Left(objCells(nRow + nTop, 1).Value, 1) = "#"  and isempty(objCells(nRow + nTop, nCol + nLeft).Value) then
						exit for
					end if
					strFileContent = strFileContent  & "|" & objCells(nRow + nTop, nCol + nLeft).Value 
				end if
			Next
			strFileContent = strFileContent & vbNewLine
		  Next
		' For i=0 to UBound(arrSheet)
			' For j=0 to 2
				' strFileContent = strFileContent  & arrSheet(i,j) 
				' if Left(arrSheet(i,j), 1) = "#" then
					' exit for
				' end if
				' if j <> 2 Then
					' strFileContent = strFileContent & "|"
				' End IF
			' Next
			' if i <> UBound(arrSheet) Then
				' strFileContent = strFileContent & vbNewLine
			' End IF
			
		' Next		
		set  a = objFSO.CreateTextFile(left(strFile,instrrev(strFile,"\")) & objSheet.name&".csv", true)
		a.WriteLine(strFileContent)
		a.Close()
		strFileContent = ""
		'Erase arrSheet
		  ' Close the workbook without saving
		  Call objSheet.Close(False)
	next
  ' Quit Excel
  objExcel.Application.Quit
End Function

Function checkvalueexistsinKeyMapping(ArrayKeyMappingMemberswithindex,strvalue)
	For lposition = 0 To UBound(ArrayKeyMappingMemberswithindex) 
		if ArrayKeyMappingMemberswithindex(lposition,0)= strvalue then
			checkvalueexistsinKeyMapping=true
		end if
		lposition = lposition+1
	Next  
End Function

function parseConfigs(arrItems,strText)
	Dim aryConfig()
	ReDim aryConfig(UBound(arrItems))
	For i=0 to UBound(arrItems)
		Set aryConfig(i) = CreateObject("Scripting.Dictionary")
		aryConfig(i).CompareMode = 1
	Next
	
	Dim oSection
	Set oSection = Nothing
	
	For Each strLine in Split(strText,vbCrlf)
		strLine = Trim(strLine)
		If strLine <> "" Then
			Select Case Left(strLine, 1)
				Case "#":
					'Do nothing comment line
				Case ":"
					for i=0 to UBound(arrItems)
						if UCase(":" & arrItems(i)) = UCase(strLine) Then
							Set oSection = aryConfig(i)
						End If
					Next
				Case Else
					If Not oSection Is Nothing Then
						i = InStr(strLine, ",")
						If i > 0 Then
							oSection.Add Trim(Left(strLine, i - 1)), Trim(Mid(strLine, i + 1))
						Else
							oSection.Add strLine, ""
						End If
					End If
			End Select
		End If
	Next
	parseConfigs = aryConfig
End function

Function checkAnalyticChildVersion(viewObject, ByRef newVersion)
	Dim flag : flag = false
	If Not viewObject.Published.IsNull Then
		Dim publishedVersion : publishedVersion = viewObject.Published.Version.Value
		For Each childView in viewObject.Folders.ChildViews.Scan(,,,"ParentVersion Desc")
			If childView.ParentVersion.Value > publishedVersion Then
				flag = true
				newVersion = childView.ParentVersion.Value
				Exit For
			End If
		Next
	End If
checkAnalyticChildVersion = flag
End Function


Function GetFolderFromPath(strFullPath)
  ' Check for null or empty path
  If IsNull(strFullPath) Or strFullPath = "" Then
    GetFolderFromPath = ""
  Else
    ' Split the path into an array
    arrPath = Split(strFullPath, "\")

    ' Check if the path ends with a backslash
    If UBound(arrPath) = 0 And arrPath(0) = "" Then
      GetFolderFromPath = ""
    Else
      ' Get the last folder name
      strLastFolder = arrPath(UBound(arrPath)-1)

      ' Check if the last folder name is empty
      If strLastFolder = "" Then
        GetFolderFromPath = ""
      Else
        ' Check if the keyword exists in the host name
        strHostName = CreateObject("WScript.Network").ComputerName
        If InStr(strHostName, strLastFolder) > 0 Then
          GetFolderFromPath = strLastFolder
        Else
          GetFolderFromPath = ""
        End If
      End If
    End If
  End If
End Function


Function ValidateEnvironment(strEnvironmentPath)

  strKeyword = GetFolderFromPath(strEnvironmentPath)

  If strKeyword = "" Then
    ValidateEnvironment = ENVMESSAGE
  Else
    ' Check if the keyword exists in the host name
    strHostName = CreateObject("WScript.Network").ComputerName

    If InStr(strHostName, strKeyword) > 0 Then
      'That means all good
    Else
      ValidateEnvironment = ENVMESSAGE
    End If
  End If
End Function


Sub LogErrorDetail(strLogMessage)
	If strLogMessage  <> "" Then 
		ObjSystemLog.LogErrorDetail "","", strLogMessage
	End If
End Sub

Sub LogNormalDetail(strLogMessage)
	If strLogMessage  <> "" Then 
		objSystemLog.LogNormalDetail strLogMessage
	End If
End Sub

Function CheckRegistrySystemKeyExists(strSystemNameKey)
	Set objRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
	objRegistry.EnumKey HKEY_LOCAL_MACHINE, strBaseKeyPath, arrSubKeys

	For Each strSubKey In arrSubKeys
		If strSubKey = strSystemNameKey Then
		  CheckRegistrySystemKeyExists = True
		  Exit Function
		End If
	Next

	CheckRegistrySystemKeyExists = False
End Function


Function ValidateReleaseConfig(l_objFSO)

	'Function to validate ReleaseInfo.config file
	'Parameter(s):
	'strReleaseConfigPath - Path of the ReleaseInfo.config file to be validated
	
	
	strErrorMessageInValidConfig = " ReleaseInfo.config should be configured like this. <br> Release~PEPSICOEUXXXX <br> Spira Ticket numbers~INXXXX <br> Developer~Rickey A."
									 
	' Get the current directory
	strCurDir = l_objFSO.GetAbsolutePathName(".")

	' Set the file path
	strReleaseConfigPath = strCurDir & "\config\ReleaseInfo.config"

	ValidateReleaseConfig = ""
	
	'Check if the ReleaseInfo.config file exists or not
	If Not l_objFSO.FileExists(strReleaseConfigPath) Then
		ValidateReleaseConfig = "Error: ReleaseInfo.config file is missing or not accessible"
		Exit Function
	End If

	' Read the content of ReleaseInfo.config file
	Set objConfigFile = l_objFSO.OpenTextFile(strReleaseConfigPath)
	strReleaseDetails = objConfigFile.ReadAll()
	objConfigFile.Close()

	' Split the ReleaseInfo.config file content by line breaks
	arrReleaseLines = Split(strReleaseDetails, vbCrLf)
	
	' Check if ReleaseInfo.config file has minimum 3 required fields
	If UBound(arrReleaseLines) < 2 Then
		ValidateReleaseConfig = "Error: ReleaseInfo.config file does not contain required fields"
	End If

	Dim arrReleaseFields(2)
	For i = 0 To 2
		arrFields = Split(arrReleaseLines(i), "~")
			' Check if ReleaseInfo.config file has required fields and values

		If UBound(arrFields) < 1 Or arrFields(0) <> Array("Release", "Spira Ticket numbers", "Developer")(i) Then
			ValidateReleaseConfig =  "Error: ReleaseInfo.config file does not contain required field - " & Array("Release", "Spira Ticket numbers", "Developer")(i)
			Exit Function
		End If
		arrReleaseFields(i) = arrFields(1)
	Next

	If ValidateReleaseConfig <> "" Then 
		ValidateReleaseConfig = ValidateReleaseConfig & vbCrLf & strErrorMessageInValidConfig
	End If
	
	If ValidateReleaseConfig <> "" Then
		MsgBox Replace(ValidateReleaseConfig , "<br>", vbCrLf) , vbCritical, "Configuration Error"
	End If
	
	
	strReleaseID = arrReleaseFields(0)



End Function