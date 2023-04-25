'*************************************************
'V.1.0.4 
'WF to export internal accounts, the code is quite similar to all external accounts but not 100% the same
'diff from internal export are
'		Internal export WF has to export info from a period of time every 15 mins
'		In Internal export UpdateDODynamicMembers we do not check PVB territories
'		In Internal export we put a max of 3000 order to export per process
'Change History 
'V.1.0.1
'20220921. AA. Pepsi Iberia - 6615157 - No territory assigned to Digital orders in export (when territory change)
'			Update store.Territories subfolder by store.TerritoriesNoFilter to take into account all store assignments
'V.1.0.2
'20220917  MM US 4587883 Validate XML to ensure its not currupted. Skip the order if XML is not valid

'V.1.0.3
'20221113  MM US 7067765 "Delivery instructions" order field XML validation (Spira#118425)
'			New code will help to avoid unicode issue in XML. It will allow to use only given char in text fields

'V.1.0.4
'20221126 MM [IN:119124] Add Max record to export property to Internal Order Export WF


'************************************************
Const BYTES_PER_SECOND = 4000
Const MIN_SEC_TO_WAIT_TO_UPLOAD = 5

'Specific WF properties for internal order export
Const BeginExportTime = "Begin Export Time"
Const EndExportTime = "End Export Time"
Const UTCOffset = "UTC Offset"

Const SysParamFTP = "SIF FTP"
Const SysParamSIF = "SIF Path"
Const archiveFolderPropertyName = "Archive Folder"
Const exportFilenamePropertyName = "Export File name"
Const fixedArchivePath = "Export\Archive\"
Const fromSifFolder = "From_Sif\"
Const errorSifFolder = "Errors\"
Const strMainXmlTagName = "Orders"
Const tempFolder = "Temp\"
Const tempValidateFolder = "Temp2\"

'Names of the values to have in PropertiesIn Object
Const   strReexecutionCount = "WF_ReExecutionCount"
'to read previous log
Const	strLog_PropertyName = "LogInfo"

'Const and variables to manage the replace of strange chars in text fields on XML fields
Const REGEX = "Regex allow special characters"
Const REGEX_DEFAULT = "[^a-z0-9 ~!_áÁéÉíÍóÓúÚñÑ¡¿ª@#$%^&*()-+=|{}':;.,<>/?]"
Const ORDERTEXTFIELDS_DEFAULT = "deliveryinstructions"
Const ORDERTEXTFIELDS = "Order Text Fields"
Dim strRegularExp : strRegularExp = ""
Dim lstTextFields : lstTextFields = ""

'Order complete status
Dim objCompleteStatus
'Order readony status
Dim objReadOnlyStatus
Dim maxFileSize
Dim intMaxSizePerFile: intMaxSizePerFile=0
Dim strExportFileNameTemplate
Dim strEnvironment,strExportPath,strArchivePath,strFTPFolderPath,strUserName,strPassword,strFlag
Dim objFSO
Dim strWFLogMessage : strWFLogMessage = ""
Dim strErrorInfo: strErrorInfo = ""
Dim wFStatus: wFStatus = False
Dim  objWFLog
'IN:104979 IELK#2767318 WKFL - EA - PepsiCo IRE-EU3 Prod - Iberia L2 WF Alert - PepsiCo Export Internal Accounts
'AAA 20211116 getting the max number of records to export per process
Dim intMaxRecordsToExport

Dim objProgressUI
Dim objPropertiesOut
	
Sub DoIt(System,PendingWork,PropertiesIn,PropertiesOut)
	Dim strXML, arrXML, strKey, strPath, strFileName 
	Dim  tmpfile, LoadStatus, xmlParseErr, strFilter 
	Dim Root, Node, pi
	
	Dim intcount,strRoutID
	Set objPropertiesOut = PropertiesOut
	
	Set objProgressUI = Nothing
	If Not PropertiesIn Is Nothing Then
		Set objProgressUI = PropertiesIn.Item("ProgressUI")
		If Not objProgressUI Is Nothing Then
				objProgressUI.Initialize 100,System.Title       
		End If
	End If
	
	'checking if we have a propertiesIn with previous log. If it exists means that the job has been reexecuted
	If PropertiesIn.Exists(strLog_PropertyName) Then
		strWFLogMessage = PropertiesIn(strLog_PropertyName) & vbNewLine
		strWFLogMessage = strWFLogMessage & vbNewLine & vbNewLine & " ********** RUNNING THE JOB AGAIN " & Now & vbNewLine
		strWFLogMessage = strWFLogMessage & "It is the " & PropertiesIn(strReexecutionCount) & " time that Pending Work is reexecuted" & vbNewLine
	End If
	
 	Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objWFLog = System.WF_LogEntry(PendingWork.WF_Process.Title, 8)
    objWFLog.PendingWork.Value = PendingWork 
	objWFLog.WF_Process.Value = PendingWork.WF_Process.Title.Value   
	objWFLog.WorkflowProcess.Value = PendingWork.WF_Process.Value
	
	strWFLogMessage = strWFLogMessage & "Workflow: " & PendingWork.WF_Process.Title & vbNewLine
	strWFLogMessage = strWFLogMessage & "Started: " & Now & vbNewLine
	strWFLogMessage = strWFLogMessage & "Properties : " & vbNewLine
 	For Each oWFProperty In PendingWork.WF_Process.Folders.WorkflowProperties.Scan(,"wfProperty;wfValue",,"wfProperty Asc")
		strWFLogMessage = strWFLogMessage & oWFProperty.wfProperty & " : " & Mid(oWFProperty.wfValue,1,1000) & vbNewLine
 	Next
	
	strArchiveFolderName = Trim(PendingWork.getWorkflowProperty(archiveFolderPropertyName, ""))
 	
 	If Not IsNull(pendingWork.WhoCreated) Then
		strUserId = pendingWork.WhoCreated
	Else
		strUserId = pendingWork.RecordStamp.WhoAdd
	End If
	strWFLogMessage = strWFLogMessage & "logged user : " & strUserID & vbNewLine
 	
 	'This code is only for internal export WF
	'AAA If the job was created by the WF server we need to check if the curent time is a valid time for the export
 	'we have WF properties to indicate the start/end time and the offset from UTC of the timezone
 	If strUserId="wrkflow" Then
 		strStartTime = PendingWork.getWorkflowProperty(BeginExportTime,"")
		strEndTime = PendingWork.getWorkflowProperty(EndExportTime,"")
		If Len(strStartTime)<>4 Or Len(strEndTime)<>4 Then
			'We expect to put the start/finish time on format HHmm, if not we create an error
			PopulateEventLog PendingWork, strWFLogMessage, strErrorInfo
			Err.Raise 1000, PendingWork.WF_Process.Title,"*** ERROR : Wrong info on begin and end export time, the expected info is HHmm" 
	 		Exit Sub
		Else
			'We check the current time in Spain from UTC time
	 		intHourOffset = PendingWork.getWorkflowProperty(UTCOffset,1)
	 		If (DaylightTime(Now)) Then
	 			intHourOffset = intHourOffset + 1
	 		End If
	 		dtTime = DateAdd("h",intHourOffset,System.GetUTCNow)
	 		strTimeText = Right("0" & Hour(dtTIme),2) & Right("0" & Minute(dtTime),2)
	 		'If current time is not in valid period we do not do anything
	 		If Not (strTimeText>=strStartTime And strTimeText<=strEndTime) Then
	 			strWFLogMessage = strWFLogMessage & vbNewLine & vbNewLIne & vbNewLIne & "Current Spanish time  is " & strTimeText & " which is out of " & strStartTime & " and " & strEndTime & " so the export is not done"
	 			PopulateEventLog PendingWork, strWFLogMessage, strErrorInfo
	 			Exit Sub
	 		End If
			strWFLogMessage = strWFLogMessage & vbNewLine & "Current Spanish time  is " & strTimeText & vbNewLine
	 	End If
 	End If
 	
 	strMaxSizePerFile = Trim(PendingWork.getWorkflowProperty("MaxSizePerFile", ""))
	If Len(strMaxSizePerFile)>0 Then
		On Error Resume Next
			intMaxSizePerFile = Int(strMaxSizePerFile)
		On Error Goto 0
	End If
	
	'vables to update special chars in xml text fields
	strRegularExp = Trim(PendingWork.getWorkflowProperty(REGEX, REGEX_DEFAULT))
	lstTextFields = Split(Trim(PendingWork.getWorkflowProperty(ORDERTEXTFIELDS, ORDERTEXTFIELDS_DEFAULT)),",")	
	
	strArchivePath = AddSlashIfNotExists(strArchiveFolderName)
	strExportFileNameTemplate = Trim(PendingWork.getWorkflowProperty(exportFilenamePropertyName, ""))
	strSystemFolderPath = getSystemParameterPath(SysParamSIF)
	strFTPFolderPath = getSystemParameterPath(SysParamFTP)
	strXMLdef = PendingWork.getWorkflowProperty("XML mapping", "")	
	strFilter = PendingWork.getWorkflowProperty("ExportFilter","1=0")


	'AAA 20200719. We check if user is admin or not. If not we need to add only the orders for valid territories
	Set objUserIsAdmin = System.Folders.Users.First("ID='" & strUserId & "' and InList(Role.code,'PepsiIB_ADM','SYSADMIN')")
	If objUserIsAdmin.IsNull Then
		' If Len(strFilter)>0 Then
			' strFilter = "( " & strFilter & " ) AND "
		' End If
		' strFilter = strFilter & " (LocationVisit.Territory.Reps.Exists((isNull(validFrom) or ValidFrom&lt;=Date) and (isNull(validTo) or ValidTo&gt;=Date) and rep.ID=""" & strUserId & """) Or LocationVisit.Territory.SalesTeam.Manager.ID=""" & strUserID & """)"
		
		'rdash - 2022-03-28 - IN: 110069 - Adding this to the filter currently because a long running SQL query that can potentially bring down the server.  We replace this code with below logic
		Dim sInlistFilter : sInlistFilter = ""
		For Each oSalesTeam In System.Folders("SalesTeams").Scan("Manager.ID='" & strUserID & "'")
			If InStr(sInlistFilter,"'" & oSalesTeam.Key & "'") = 0 Then
				sInlistFilter = sInlistFilter & ",'" & oSalesTeam.Key & "'"
			End If
		Next
		For Each oSalesTeam In System.Folders("SalesTeams").Scan()
			For Each oTerritories In oSalesTeam.Folders("AS_Territories").Scan("Folders.Reps.Exists((isNull(validFrom) or ValidFrom<=Date) and (isNull(validTo) or ValidTo>=Date) and rep.ID='" & strUserId & "')")
				If InStr(sInlistFilter,"'" & oSalesTeam.Key & "'") = 0 Then
					sInlistFilter = sInlistFilter & ",'" & oSalesTeam.Key & "'"
				End If
			Next
		Next

		If Len(sInlistFilter) Then
			strFilter =  strFilter & " And InList(LocationVisit.Territory.SalesTeam.Key" & sInlistFilter & ")"
		End If
	End If
	'replacing ' into " on filter to avoid issues on xml formatting
	strFilter = Replace(strFilter,"'","""")
	
	strStatusFilter = Trim(PendingWork.getWorkflowProperty("Export Order Status", "Key=""S"""))
	Set objCompleteStatus = System.Folders.OE_OrderStatuses.First(strStatusFilter)
	Set objReadOnlyStatus = System.Folders.OE_OrderStatuses.First("Key='R'")
	
	strWFLogMessage = strWFLogMessage & _
		"Archive folder : " & strArchivePath & vbNewLine & _
		exportFilenamePropertyName & " : " & strExportFileNameTemplate & vbNewLine & _
		"System path : " & strSystemFolderPath & vbNewLine & _
		"Ftp Path : " & strFTPFolderPath & fromSifFolder & vbNewLine & _
		"order filter: " & strFilter & vbNewLine
	PopulateEventLog PendingWork, strWFLogMessage, strErrorInfo
	
	If Not objCompleteStatus.IsNull And Len(strXMLdef)>0 And Len(strArchivePath)>0 And Len(strExportFileNameTemplate)>0 And Len(strSystemFolderPath)>0 And Len(strFTPFolderPath)>0  And _ 
		createFolderIfNotExist(objFSO, strSystemFolderPath & fixedArchivePath) And createFolderIfNotExist(objFSO,strSystemFolderPath & fixedArchivePath & strArchivePath) And _
		createFolderIfNotExist(objFSO,strSystemFolderPath & fixedArchivePath & strArchivePath & tempFolder)  And _
		createFolderIfNotExist(objFSO,strSystemFolderPath & fixedArchivePath & strArchivePath & tempValidateFolder) Then
		
		strArchivePath = strSystemFolderPath & fixedArchivePath & strArchivePath
		strWFLogMessage = strWFLogMessage & "Archive folder : " & strArchivePath & vbNewLine
		PopulateEventLog PendingWork, strWFLogMessage, strErrorInfo
		
		CopyMoveExistingFiles PropertiesIn, objFSO, strArchivePath, strFTPFolderPath & fromSifFolder, strExportFileNameTemplate, strWFLogMessage
		
		'AA 2021/05/07. Bug 1078363. We move the orders to be exported to readyTosend status (not editable) to avoid editing the order when it is on export process
		strWFLogMessage = strWFLogMessage & vbNewLine & "--Putting order from delivery pending to ready to send status. Begins at "  & Now  &  vbNewLine
		PopulateEventLog PendingWork, strWFLogMessage, strErrorInfo
		intNumSavedItems = 0
		'filter to get all orders from editable status to move them to R status (readonly orders)
		strFilterOrdersToMove = Replace(Replace("( " & strFilter & " ) AND Status.CanEdit","&lt;","<"),"&gt;",">")
		'filter added to manage the Digital orders
		strFilterOrdersToUpdate = Replace(Replace("( " & strFilter & " ) AND pikOrderType.EX_External.ID=""DO""","&lt;","<"),"&gt;",">")
		strWFLogMessage = strWFLogMessage & "filters to use: " & vbNewLine & _
			"Move to delivery pending : " & strFilterOrdersToMove & vbNewLine & _
			"Update DO order : " & strFilterOrdersToUpdate & vbNewLine
		
		totalRSRecords = System.Folders.OE_OrderHeaders.Count(strFilterOrdersToMove)
		'this is only on internal orders
		'IN:104979 IELK#2767318 WKFL - EA - PepsiCo IRE-EU3 Prod - Iberia L2 WF Alert - PepsiCo Export Internal Accounts
		'AAA 20211116 getting the max number of records to export per process
		intMaxRecordsToExport = PendingWork.getWorkflowProperty("Max Records to Export", 3000)	
		'getting the max number of records to export per process
		If intMaxRecordsToExport<totalRSRecords Then
			totalRSRecords = intMaxRecordsToExport
		End If
		totalDORecords = System.Folders.OE_OrderHeaders.Count(strFilterOrdersToUpdate)
		totalRecords = 2*(totalRSRecords+totalDORecords)
		strWFLogMessage = strWFLogMessage & "Updating status on #orders:" & totalRecords/2 & vbNewLine
		PopulateEventLog PendingWork, strWFLogMessage, strErrorInfo

		'We run the same code twice one to move the orders from ready to send to send (not editable to editable)
		'2nd time to move the orders from Digital
		For intLoop=1 To 2
			If IntLoop=1 Then
				strFilterExpression = strFilterOrdersToMove
			Else
				strFilterExpression = strFilterOrdersToUpdate
				'RMK IN:107288 Digital - Simplify some digital order export expressions
				strWFLogMessage = strWFLogMessage & vbNewLine & "--Setting digital order dynamic members. Begins at "  & Now  &  vbNewLine
				strWFLogMessage = strWFLogMessage & "Updating " & totalDORecords & " Digital Orders" & vbNewLine
				PopulateEventLog PendingWork, strWFLogMessage, strErrorInfo
			End If
		
			For Each objOrder In System.Folders.OE_OrderHeaders.Scan(strFilterExpression,"Status;OrderDate;DeliveryDate;OrderNo",intMaxRecordsToExport)
				intPercentage = Int(100 * intNumSavedItems  /totalRecords)
				If PendingWork.ProgressPercentComplete <> intPercentage Then
					PendingWork.ProgressPercentComplete = intPercentage
					PendingWork.StepDesc = intNumSavedItems & " of " & (totalRSRecords/2) & "    objects processed "
					PopulateEventLog PendingWork, strWFLogMessage, strErrorInfo
				End If
				If intLoop = 2 Then
					UpdateDODynamicMembers objOrder, strWFLogMessage
				End If
				objOrder.Status = objReadOnlyStatus
				'AAA 20220822 Calling a method to save the order where code does any possible change to fix a validation issue and code tries to save the order
				'if there is a validation issue the order will not be saved and a message will be added to the log
				SavingTheObject Nothing, objOrder, strWFLogMessage
				'TODO, pending to update the code to be able to check the number of order processed, number of orders saved, number of failed orders
				Set objTrans = System.BeginTransaction
				objOrder.Save objTrans
				strValidationError = checkValidation(objTrans)
				If Len(strValidationError)>0 Then
						strWFLogMessage = strWFLogMessage & "***** ERROR **** :Order # " & objOrder.OrderNo & " cannot be saved for " & strValidationError & vbNewLine
					Else
						'Saving the order
						objTrans.Commit
						intNumSavedItems = intNumSavedItems + 1
				End If
			Next
			If IntLoop=1 Then
				strWFLogMessage = strWFLogMessage & "--Finish moving order to ready to send status. Ends at " & Now &  vbNewLine 
			Else
				strWFLogMessage = strWFLogMessage & "--Finish setting digital order dynamic members. Ends at " & Now &  vbNewLine 
			End If
		Next
		

		'IN:101059 Digital - 1308889 - SiF_Export Digital orders to ERP
		'We only export orders that are not editable, i.e. in ready to send status Or DO status
		strFilter = "( " & Replace(strFilter,"'","""") & " ) AND (Not Status.CanEdit)"
		
		'IN:108772 Digital - 2478907 - Exclude digital orders that have not been updated to ready to send
		strFilter = strFilter & " And Status.Key&lt;&gt;""DO"" "
		
		strExpConfig = Replace(strXMLdef,"{0}",strFilter)
		Set objExport = System.GetExportToXML
		If Not objExport.ParseConfig(strExpConfig) Then
			Set objXML = CreateObject("Msxml2.DOMDocument")
			objXML.LoadXML(strExpConfig)
			strWFLogMessage = strWFLogMessage & "Error with export configuration." & objXML.parseError.reason &  vbNewLine
			strErrorInfo = strErrorInfo & "Error with export configuration." & objXML.parseError.reason &  vbNewLine
			wFStatus = False
			PopulateEventLog PendingWork, strWFLogMessage, strErrorInfo
			Err.Raise vbObjectError + 513, PendingWork.WF_process.title,strErrorInfo
		Else
			'Set CallBack properties
			Set objCallBack = New UpdateOrderCallback
			Set objCallBack.system = system
			intNumOfExportedOrders = System.Folders.OE_OrderHeaders.Count(Replace(Replace(strFilter,"&lt;","<"),"&gt;",">"))
			objCallBack.totalRecords = intNumSavedItems + intNumOfExportedOrders
			objCallBack.processedRecords = intNumSavedItems
			Set objCallBack.PendingWork = pendingWork
			Set objCallBack.objProgressUI = objProgressUI
			Set objCallBack.PropertiesOut = PropertiesOut
			'End setting CallBack properties
			Set objDictCallBack = CreateObject("Scripting.Dictionary")
			objDictCallBack.add "UpdateOrderCallback", objCallBack
			Set objExport.objCallBackDict = objDictCallBack 
			strWFLogMessage = strWFLogMessage & "Num of orders to export: " & intNumOfExportedOrders & vbNewLine
			PopulateEventLog PendingWork, strWFLogMessage, strErrorInfo
			'Set new export xml for each export order process
			If Not objExport.CreateExport(strMainXmlTagName) Then
				Set objXML = CreateObject("Msxml2.DOMDocument")
				objXML.LoadXML(strExpConfig)
				strWFLogMessage = strWFLogMessage & "Error with export configuration 2." & objXML.parseError.reason &  vbNewLine
				strErrorInfo = strErrorInfo & "Error with export configuration 2." & objXML.parseError.reason &  vbNewLine
				wFStatus = False
				PopulateEventLog PendingWork, strWFLogMessage, strErrorInfo
				Err.Raise vbObjectError + 513, PendingWork.WF_process.title,strErrorInfo
			Else
				Set objCallBack.objXMLExport = objExport
				strExportFileName = GetTimeStampFileName(strExportFileNameTemplate, strArchivePath)
				strFName = strArchivePath & tempFolder & strExportFileName
				strValidateFName = strArchivePath & tempValidateFolder & strExportFileName
				objCallBack.strOutputPath = strFName
				objCallBack.strOutputValidatePath = strValidateFName
				objCallBack.strLog = ""
				strErr = objExport.DoExport(Nothing)
				strWFLogMessage = strWFLogMessage & objCallBack.strLog
				If Len(objExport.objXMLExportNode.xml)>Len(strMainXmlTagName) + 5 Then
					
					'We update the XML to a pretty version. We want to save this to temp2 as a backup so if its not good then we can move temp original file
					objExport.FormatXML(objExport.objXMLExportNode)
					SaveXMLWithReTry objExport, objCallBack.strOutputValidatePath
					
					'Just for testing to crash the file. Use this for debug only
					'Call MashupXMLFileToCrash(Nothing, objCallBack.strOutputValidatePath, objCallBack.strLog, true) 
					
					If Not ValidateXMLFile(objCallBack.strOutputValidatePath,objCallBack.strLog, Nothing) Then 
						objCallBack.strLog = objCallBack.strLog & "Formating the file is causing an issue, so moving a file without formatting" & objCallBack.strOutputPath & vbNewLine
						copyMoveFile objFSO, objCallBack.strOutputPath, strArchivePath, strFTPFolderPath & fromSifFolder, strWFLogMessage
						DeleteFileWithReTry objFSO, objCallBack.strOutputValidatePath 'Delete the file from temp2 which is not valid
					Else
						copyMoveFile objFSO, objCallBack.strOutputValidatePath, strArchivePath, strFTPFolderPath & fromSifFolder, strWFLogMessage
						DeleteFileWithReTry objFSO, objCallBack.strOutputPath  'Delete the file from temp which is not valid
					End If

					strWFLogMessage = strWFLogMessage & "File saved to " & objCallBack.strOutputPath & vbNewLine
				Else
					strWFLogMessage = strWFLogMessage & " --- No Orders to export --- " & vbNewLine
				End If
				intExported = intExported + objCallBack.intTotalExported
				intErroredExport = intErroredExport + objCallBack.intTotalUpdateFailedCount
				PopulateEventLog PendingWork, strWFLogMessage, strErrorInfo
			End If
		End If	
	Else
		wFStatus = False
		strWFLogMessage = strWFLogMessage & "Error. Archive path, Ftp Path & fileName are mandatories" & vbNewLine
		strErrorInfo = strErrorInfo & "Error. Archive path, Ftp Path & fileName are mandatories" & vbNewLine
		PopulateEventLog PendingWork, strWFLogMessage, strErrorInfo
		Err.Raise vbObjectError + 513, PendingWork.WF_process.title,strErrorInfo
	End If
	
 	strWFLogMessage = strWFLogMessage & "Total exported records: " & intExported & "  Total failed records: " & intErroredExport & vbNewLine
 	strWFLogMessage = strWFLogMessage & "Finished: " & Now & vbNewLine
	strWFLogMessage = strWFLogMessage & "-------------------------------------------------------------------------------------------------" & vbNewLine & vbNewLine
	
	objWFLog.WF_LogEventType.Value = 1 'Process Flow
	PopulateEventLog PendingWork, strWFLogMessage, strErrorInfo
		 	
	Set objWFLog = Nothing
	Set objDOM = Nothing
	Set objFSO = Nothing
 	
End Sub
'20220822 AAA function to call for saving the order, it checks any possible validation error
'objTrans parameter must be nothing if it was not created and the transacitonal object if the object is saved with other items
'it returns true if the object was saved
Function SavingTheObject(ByRef objTrans,ByRef objInstance, ByRef strMessage)
    blnObjectsaved = False
    'vable to know if transactional object was sent or not to the function
    blnTransCreatedInthisFunction = False
    '20220822 AAA Calling a method to update the order to fix any possible validation
    UpdateOrderToFixPossibleValidations objInstance, strMessage
    'creating transaction if not exists
    If objTrans Is Nothing Then
        Set objTrans = System.BeginTransaction
        blnTransCreatedInthisFunction = True
    End If
	'20220825 AVMishra IN:114995 - Checking if order has a valid orderNo
    'Remove any previous possible validation messages
    objInstance.ExtractedDetails = ""
	objInstance.DM_OEValidationError = 0 
    If (Len(objInstance.orderNo & "")=12 And InStr(objInstance.OrderNo,"null")<=0) Then
        'save the instance
        objInstance.Save objTrans
        'getting the validation messages if any
        strValidationError = checkValidation(objTrans)
    Else
        strValidationError = "Order Number has not a valid value"
    End If
    If Len(strValidationError)>0 Then
            'show validation error
            strMessage = strMessage & "***** ERROR **** :Order # " & objInstance.OrderNo & " cannot be saved for " & strValidationError & vbNewLine
			'20220825 AVMishra IN:114995 - Adding error into Order field, we use SQL because order cannot be saved
            System.ODBC_Execute("update OE_OrderHeader set ExtractedDetails = '" & Left(strValidationError,250) & "', DM_OrderHeaderAutoData_DM_Bit_0=1 where OE_OrderHeader = '" & objInstance.key & "'")
        Else
            If blnTransCreatedInthisFunction Then
                'Commit the transaction
                objTrans.Commit
            End If
            blnObjectsaved = True
    End If
    SavingTheObject = blnObjectsaved
End Function

'20220822 AAA function to call before saving an order to fix any possible validation issue
Sub UpdateOrderToFixPossibleValidations(ByRef objOrder, ByRef strMessage)
	'Rdash 20220727 [IN:112794] - Updated code to check Delivery date by checking validation error in Transaction and update the delivery date if deliverydate<=orderdate
	If (objOrder.DeliveryDate <= objOrder.OrderDate) Then
		deliverydate = objOrder.OrderDate + 1
		If (Weekday(deliverydate) = 7) Then
			deliverydate = deliverydate + 2
		ElseIf (Weekday(deliverydate) = 1) Then 
			deliverydate = deliverydate + 1
		End If
		strMessage = strMessage & "The delivery date is same as order date or same as todays date so we are updating the delivery date of OrderNo. " & objOrder.OrderNo & " from " & objOrder.DeliveryDate & " to " & deliverydate &  vbNewLine
		objOrder.DeliveryDate = deliverydate
	End If
End Sub

'20220822 AAA Function to check the validation message
Function checkValidation(objTrans)
	Set oValidate = objTrans.Validate
	strErrMessage = ""
	If oValidate.Count>0 Then
		For i=0 To oValidate.count-1
			strErrMessage = strErrMessage & oValidate(i).Instance.prompt & "  -  " & oValidate(i).Message & vbNewLine
		Next
	End If
	checkValidation = strErrMessage
End Function

Function DeleteFileWithReTry(objFSO, strSourceFile)

	DeleteFileWithReTry = False
	
	
	If objFSO.FileExists(strSourceFile) Then
		
		'We try to delete the file in one minute in case that file is locked for previous action
		intAttempts = 0
		
		Do
			intAttempts = intAttempts + 1
			On Error Resume Next
			objFSO.DeleteFile strSourceFile
			blnError = (Err.Number <> 0) 
			strError = ". Error: " & err.Number & ". " & err.Description
			On Error Goto 0
			If blnError Then
				'Wait less than 2 secs
				dteStart = Now
				While DateDiff("s",dteStart,Now)<2
				Wend
			End If
			
		Loop Until (Not blnError) Or intAttempts>30
		
		If blnError Then
			strWFLogMessage = strWFLogMessage & Now & " *****ERROR File could not be deleted from path " & strSourceFile & " Error." & err.Number & " > " & err.Description & vbNewLine
		Else
			DeleteFileWithReTry = True
		End If
	Else
		strWFLogMessage = strWFLogMessage & Now & " *****File does not exists: " & strSourceFile &  vbNewLine
	End If	
		
End Function

Sub OnlyMoveFile(objFSO, strSourceFile,strTargetFile)
	
	'We try to move the file in one minute in case that file is locked for previous action
	intAttempts = 0
	
	Do
		intAttempts = intAttempts + 1
		On Error Resume Next
		objFSO.CopyFile strSourceFile, strTargetFile, True
		blnError = (Err.Number <> 0) 
		strError = ". Error: " & err.Number & ". " & err.Description
		On Error Goto 0
		If blnError Then
			'Wait less than 2 secs
			dteStart = Now
			While DateDiff("s",dteStart,Now)<2
			Wend
		Else
			'We will disable this log because of the fact that it will scan 3k orders and this log will mashup other useful info. So we will use this only during investigation
			'strWFLogMessage = strWFLogMessage & "File moved to temp folder " & strTargetFile & " on attempt: " & intAttempts & vbNewLine
			'Delete the file with 30 re-try
			DeleteFileWithReTry objFSO, strSourceFile
		End If
		
	Loop Until (Not blnError) Or intAttempts>30
	
	If blnError Then
		strWFLogMessage = strWFLogMessage & Now & " *****ERROR File could not be moved to path " & strTargetFile & " Error." & err.Number & " > " & err.Description & vbNewLine
	End If
		
End Sub

'Function to read a input property
Function getPropertyInInfo(PropertiesIn, strPropertyName)
	If PropertiesIn.Exists(strPropertyName) Then
		getPropertyInInfo = PropertiesIn(strPropertyName)
	Else
		getPropertyInInfo = ""
	End If
End Function

Sub copyMoveFile(objFSO, strFileName,strArchivePath,strFTPSIFPath, ByRef strWFLogMessage)
	On Error Resume next
		' save file for archive
		objFSO.CopyFile strFileName, strFTPSIFPath 
		blnError = Err.Number <> 0
		strError = ". Error: " & err.Number & ". " & err.Description
	On Error Goto 0
		If blnError Then 
		strWFLogMessage = strWFLogMessage & "*****ERROR File could not be copied to path " & strFTPSIFPath & strError & vbNewLine
		Else
		strWFLogMessage = strWFLogMessage & "File copied to FTP folder " & strFTPSIFPath & vbNewLine
		'We try to move the file in one minute in case that file is locked for previous action
		intAttempts = 0
		Do
			intAttempts = intAttempts + 1
			On Error Resume Next
			objFSO.MoveFile strFileName,strArchivePath
				blnError = (Err.Number <> 0) 
				strError = ". Error: " & err.Number & ". " & err.Description
			On Error Goto 0
			If blnError Then
				strWFLogMessage = strWFLogMessage & Now & " *****ERROR File could not be moved to path " & strArchivePath & strError  & " " & Now & vbNewLine
				'Wait less than 2 secs
				dteStart = Now
				While DateDiff("s",dteStart,Now)<2
				Wend
			Else
				strWFLogMessage = strWFLogMessage & "File copied to archive folder " & strArchivePath & " on attempt: " & intAttempts & vbNewLine
			End If
		Loop Until (Not blnError) Or intAttempts>30
		If blnError Then
			strWFLogMessage = strWFLogMessage & Now & " *****ERROR File could not be moved to path " & strArchivePath & " Error." & err.Number & " > " & err.Description & vbNewLine
		End If
	End If
End Sub


Sub CopyMoveExistingFiles(PropertiesIn, objFSO, strArchivePath,strFTPSIFPath, strExportFileNameTemplate, ByRef strWFLogMessage)
	If Not PropertiesIn Is Nothing Then
		If PropertiesIn.Exists("PreviousLogInfo") Then
		  strWFLogMessage = strWFLogMessage & vbNewLine & _
			"------------- Exported info before stopping the WF Job ------------------- " & vbNewLine & _ 
			propertiesin.item("PreviousLogInfo") & vbNewLine & _
			"-------------End of the log of job before stopping the WF job -------------" & vbNewLine & vbNewLine
		End If
	End If
	strWFLogMessage = strWFLogMessage & "*** Checking if there is a non moved file that has to be moved to archivoe and FTP from a previous errored export ***" & vbNewLine
	Set objMainFolder = objFSO.GetFolder(strArchivePath & tempFolder )
	'Initialize the regular expression
	Set re = New RegExp
	re.Global  = False
	re.IgnoreCase = True
	intDot = InStrRev(strExportFileNameTemplate, ".")
	If intDot > 0 Then
		strFileNameSearch = "^" & Left(strExportFileNameTemplate, intDot - 1) & ".*" & Mid(strExportFileNameTemplate, intDot)
	Else
		strFileNameSearch = "^" & strExportFileNameTemplate & ".*"
	End If
	re.Pattern = strFileNameSearch
	For Each objFile In objMainFolder.Files
		If re.Test(objFile.Name) Then
			strWFLogMessage = strWFLogMessage & "Moving previuosly created file " & objFile.Name & vbNewLine
			copyMoveFile objFSO, objFile.path,strArchivePath,strFTPSIFPath, strWFLogMessage
		Else
			strWFLogMessage = strWFLogMessage & "File does not match the filename template, not moved " & objFile.Name & vbNewLine
		End If
	Next
	strWFLogMessage = strWFLogMessage & "*** Finish checking previous non sent files ***" & vbNewLine
End Sub

'Function to populate different DMs on Digital orders
'In Internal orders the territory will be PVM first, if not TV. Different than for external where first check is PVB
Sub UpdateDODynamicMembers (objOrderHeader, ByRef strWFLogMessage)

	strOrderDate = System.GetFormatDate(objOrderHeader.OrderDate)
	strFilter = ""
	strFilter = strFilter & "(isNull(validFrom) or validFrom<=" & strOrderDate & ") and "
	strFilter = strFilter & "(IsNull(validTo) Or ValidTo>=" & strOrderDate & ") And "
	strFilter = strFilter & "Territory.DM_RouteType<>'REP' And "
	strFilter = strFilter & "Territory.Team.Ex_External.ID='<TEAM>' "

	'20220921 AAA Pepsi Iberia - 6615157 - No territory assigned to Digital orders in export (when territory change)
	'We use TerritoriesNoFilter subfolder to have access to all territory assignments of the store
	'IMPORTANT *** Below code is not necessary in Internal Export orders only on External Export Order
	'Set objTerritory = objOrderHeader.Store.Folders.TerritoriesNoFilter.First(Replace(strFilter, "<TEAM>", "PVB"),"IIF(IsNull(ValidFrom),DateSerial(2000,1,1),ValidFrom) Asc;IIF(IsNull(ValidTo),DateSerial(3000,1,1),validTo) Desc")
	'If objTerritory.IsNull Then
		Set objTerritory = objOrderHeader.Store.Folders.TerritoriesNoFilter.First(Replace(strFilter, "<TEAM>", "PVM"),"IIF(IsNull(ValidFrom),DateSerial(2000,1,1),ValidFrom) Asc;IIF(IsNull(ValidTo),DateSerial(3000,1,1),validTo) Desc")
		If objTerritory.IsNull Then
			Set objTerritory = objOrderHeader.Store.Folders.TerritoriesNoFilter.First(Replace(strFilter, "<TEAM>", "TV"),"IIF(IsNull(ValidFrom),DateSerial(2000,1,1),ValidFrom) Asc;IIF(IsNull(ValidTo),DateSerial(3000,1,1),validTo) Desc")
		End If
	'End IF		
	
	If Not objTerritory.IsNull Then
		objOrderHeader.DM_DOTeamID = system.nvl(objTerritory.Territory.Team.EX_External.ID.Value, "")
		objOrderHeader.DM_DOTerritoryID = system.nvl(objTerritory.Territory.EX_External.ID.Value, "")
		objOrderHeader.DM_DOTerritoryName = system.nvl(objTerritory.Territory.TerritoryName.Value, "")
	Else
		strWFLogMessage = strWFLogMessage & "--No Territory to set DMs with for Order " & objOrderHeader.OrderNo & vbNewLine & vbNewLine
	End If

End Sub

Sub PopulateEventLog(PendingWork, strWFLogMessage, strErrorInfo)
	PendingWork.ProgressDescr.Value = strWFLogMessage
	 
	objWFLog.Notes = strWFLogMessage
	'AAA 2014/03/24. We put the error info in the error description
	If Len(strErrorInfo)>0 Then
		PendingWork.ErrorInfo.Value = "Following errors have been found: " & vbNewLine & strErrorInfo
		objWFLog.WF_LogEventType.Value = 2 'Error
	End If	
	objWFLog.Save
	
	'adding the info on propertiesOUt to read in case of reexecution
	objPropertiesOut.Add strLog_PropertyName, PendingWork.ProgressDescr.Value
	PendingWork.Save
End Sub

Function createFolderIfNotExist(objFSO,strFolderName)
	createFolderIfNotExist = True
	If Not objFSO.FolderExists(strFolderName) Then
			objFSO.CreateFolder(strFolderName)
	End If
	If  Not Right(strFolderName, 1) = "\"  Then
	    	strFolderName = strFolderName &"\"
	End If
End Function

Function MoveToFTPFolder(oShell, objFTP, objFile)
	MoveToFTPFolder = ""
	If Not FileExists(objFTP,objFile.Name) Then
		strParent = objFile.ParentFolder
		Set objFolder = oShell.Namespace(strParent)
		Set objNewItem = objFolder.ParseName(objFile.Name)
		objFTP.MoveHere objNewItem,  16+8+4 '1024 No show UI 16 yes to all, 4 no show progress bar
		blnMoved = False
		'We calculate the max number we are going to wait to upload the file
		intAttempts = -1
		intMaxAttemps = (objFile.Size\BYTES_PER_SECOND)+1
		strFileName = objFile.Name
		dteBeginSending = Now
		While Not blnMoved And (DateDiff("s",dteBeginSending,Now)<MIN_SEC_TO_WAIT_TO_UPLOAD Or intAttempts<intMaxAttemps)
			If FileExists(objFTP,strFileName) Then
				blnMoved = True
			Else
				'We wait less than 2 secs
				dteStart = Now
				While DateDiff("s",dteStart,Now)<2
				Wend
				intAttempts = intAttempts + 1
			End If
		Wend
		If Not blnMoved Then
			MoveToFTPFolder = "File could not be moved"
		End If
	Else
		MoveToFTPFolder = "File already exists in FTP Folder"
	End If
End Function



' ****************************************
Function prettyXml(ByVal sDirty)
' ****************************************
' Put whitespace between tags. (Required for XSL transformation.)
' ****************************************
'12/13/2019 PK IN 74829 - USDSD : New Line Character  in Settlement XML File			
'sDirty = Replace(sDirty, "><", ">" & vbCrLf & "<")
' ****************************************
' Create an XSL stylesheet for transformation.
' ****************************************
  Dim objXSL : Set objXSL = CreateObject("Msxml2.DOMDocument")
  objXSL.loadXML  "<xsl:stylesheet version=""1.0"" xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"">" & _
                  "<xsl:output method=""xml"" indent=""yes""/>" & _
                  "<xsl:template match=""/"">" & _
                  "<xsl:copy-of select="".""/>" & _
                  "</xsl:template>" & _
                  "</xsl:stylesheet>"
' ****************************************
' Transform the XML.
' ****************************************

  Dim objXML : Set objXML = CreateObject("Msxml2.DOMDocument")
  objXML.loadXml sDirty
  objXML.transformNode objXSL
  prettyXml = objXML.xml
End Function


Function GetTimeStampFileName(strFileName, strArchivePath)
	Dim intDot, strStamp
	Do
		strStamp = Year(Now()) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2) & Right("0" & Hour(Now()), 2) & Right("0" & Minute(Now()), 2) & Right("0" & Second(Now()), 2)
		intDot = InStrRev(strFileName, ".")
		If intDot > 0 Then
			strNewFileName = Left(strFileName, intDot - 1) & strStamp & Mid(strFileName, intDot)
		Else
			strNewFileName = strFileName& strStamp
		End If
	Loop While objFSO.FileExists(strArchivePath & strNewFileName)
	GetTimeStampFileName = strNewFileName
End Function

Function getSystemParameterPath(strSysParamName)
	strInfo = ""
	Set objSysParam = System.Folders.stdSystemParameters.First("Name='" & strSysParamName & "'")
	If Not objSysParam.IsNull Then
		strInfo = AddSlashIfNotExists(Replace(Replace(objSysParam.Expression,"\\","\"),"""",""))
	End If
	getSystemParameterPath = strInfo
End Function

Function AddSlashIfNotExists(strPath)
	If Right(Trim(strPath),1) <> "\" and Len(Trim(strPath))>0 Then
			strPath = Trim(strPath) & "\"
	End If
	AddSlashIfNotExists = strPath
End Function

'-------------- Update Order -------------'
Class UpdateOrderCallback
	Public System

	Public expExtractDate			'Expression of what Extracted Date should be

	Public strOutputPath			'Output file path & name
	Public strOutputValidatePath	'Output file path & name for Validation file
	Public strLog					'Output logging string
	Private strErrMessage			'Error message used in validation

	Public intTotalExported        	'Total Orders Exported
	Public intTotalUpdatedCount    	'Total Orders Updated
	Public intTotalUpdateFailedCount'Number of orders where update failed

	Public PendingWork				'WF pendingjob, it is used to update the progress bar
	Public totalRecords				'Num of records to export
	Public processedRecords			'Num of exported records until now
	Public objProgressUI
	Public PropertiesOut

	Public objXMLExport				'XML output
	Public objStatus				'Status to be updated to
	Private objTrans


	'DESCRIPTION:
	'			Updates OrderHeader with values specified in class (expExtractDate, objStatus) 
	'PARAMS:
	'			objStore - Store object to update   
	'           objXMLClassTag
	'RETURNS:	True if no errors encountered when updating order
	'MODS:      JM 2012-03-14 SR21-66871 Added new parameter objXMLClassTag to function Run
	Function run(objItem, objXMLClassTag)
		processedRecords = processedRecords+1
		intPercentage = Int(100 * processedRecords  /totalRecords)
		If PendingWork.ProgressPercentComplete <> intPercentage Then
			PendingWork.ProgressPercentComplete = intPercentage
			PendingWork.StepDesc = processedRecords & " of " & totalRecords & "    objects processed "
			PopulateEventLog PendingWork, strWFLogMessage, strErrorInfo
		End If
	
		run = False
		If me.System Is Nothing Then
			strLog = strLog & "System could not be found in the call back." & Chr(13) & Chr(10)
			run = False
			Exit Function
		ElseIf Not TypeName(me.System) = "IAeSystem" Then
			strLog = strLog & "System could not be found in the call back." & Chr(13) & Chr(10)
			run = False
			Exit Function		
		End If
		
		'We move the order to processing status
		objItem.Status = objCompleteStatus
		'We save the date and time of the export
		objItem.ExtractedDate = Now
		'save filename in details
		objItem.ExtractedDetails = strOutputPath
		'created a transaction to save the order header and order details chnages
		Set objTrans = me.System.BeginTransaction
		'AAA 20210730 bug 1347998. Wrong quantities in Delivery for order 000771101783
		'Copy some info into order detail DM
		For Each objOrderDetail In objItem.Folders.OE_OrderDetails.Scan("1=1","OE_OrderData.DM_GreenTax; OE_OrderData.DM_SugarTax; OE_OrderData.DM_UnitsPerCase; OE_OrderData.DM_Tax1Percentage; OE_OrderData.DM_Tax2Percentage; OE_OrderData.DM_Tax3Percentage; OE_OrderData.DM_NPIM; OE_OrderData.DM_PNP; OE_OrderData.DM_PTU; OE_OrderData.DM_QIM")
			If Len(objOrderDetail.OE_OrderData.DiscountDetail3 & "")>0 Then
				'20220406 Adding a new code to update DM info on order detail from info on OE_OrderData.DiscountDetail3, the info in that field is updated on scripted discount rule
				'and the info on DM cannot be updated in that code because the updated DM info on order detail in Touch is not saved into Edge
				lstDMFields = Array("DM_GreenTax","DM_SugarTax","DM_UnitsPerCase","DM_Tax1Percentage","DM_Tax2Percentage","DM_Tax3Percentage","DM_NPIM","DM_PNP","DM_PTU","DM_QIM")
				lstInfoFromTouch = Split(objOrderDetail.OE_OrderData.DiscountDetail3,"|")
				If (Ubound(lstInfoFromTouch)>=Ubound(lstDMFields)) Then
					For intNumDms=0 To Ubound(lstDMFields)-1
						objOrderDetail.Members("OE_OrderData." & lstDMFields(intNumDms)) = lstInfoFromTouch(intNumDms)
					Next
					objOrderDetail.Save objTrans
				End If
			Else
				'Info is not created on scripted discount rule
				If Not objOrderDetail.TaxRate.IsNull Then
					'Copying tax rates
					objOrderDetail.OE_OrderData.DM_Tax1Percentage = System.NVL(objOrderDetail.TaxRate.Rate1,0)
					objOrderDetail.OE_OrderData.DM_Tax2Percentage = System.NVL(objOrderDetail.TaxRate.Rate2,0)
					objOrderDetail.OE_OrderData.DM_Tax3Percentage = System.NVL(objOrderDetail.TaxRate.Rate3,0)
				End If
				'Copying packs per case
				'objOrderDetail.OE_OrderData.DM_UnitsPerCase = objOrderDetail.Product.PacksPerCase
				objOrderDetail.OE_OrderData.DM_OrderDetail.DM_Numeric4 = objOrderDetail.Product.PacksPerCase
				'Copying Sugar and Green tax info
				objOrderDetail.OE_OrderData.DM_SugarTax = Round(System.NVL(objOrderDetail.Product.dm_productpack2.DM_Num1,0),3)
				objOrderDetail.OE_OrderData.DM_GreenTax = Round(System.NVL(objOrderDetail.Product.dm_productpack2.DM_Num2,0),3)
				objOrderDetail.Save objTrans
			End If
		Next
		'AAA 20220822 Calling a method to save the order where code does any possible change to fix a validation issue and code tries to save the order
		'	if there is a validation issue the order will not be saved and a message will be added to the log
		blnObjectSaved = SavingTheObject(objTrans, objItem, strLog)
		If blnObjectSaved Then 
				Set xmlOrderHeader = objXMLClassTag
				For Each xmlBonus In xmlOrderHeader.selectNodes("./bonusallocations")
					strOrderLineKey = xmlbonus.GetElementsByTagName("orderLineKey")(0).text
					strBonusDesc = xmlbonus.GetElementsByTagName("bonusdiscdesc")(0).text & ";" & xmlbonus.GetElementsByTagName("U1_BonusAllocation")(0).text & ";" & xmlbonus.GetElementsByTagName("U2_BonusAllocation")(0).text & ";" & xmlbonus.GetElementsByTagName("U3_BonusAllocation")(0).text
					For Each xmlOrderDetail In objXMLExport.objXMLExportNode.selectNodes("//order_line[orderLineKey='" & strOrderLineKey & "' and U1BonusEntitlementQty>0]")
						If Len(xmlOrderDetail.GetElementsByTagName("U1_BonusDiscountDesc")(0).text)>0 Then
							strPreviousDiscount = xmlOrderDetail.GetElementsByTagName("U1_BonusDiscountDesc")(0).text & "|" 
						Else
							strPreviousDiscount = ""
						End If
						xmlOrderDetail.GetElementsByTagName("U1_BonusDiscountDesc")(0).text=strPreviousDiscount & strBonusDesc 
					Next
					For Each xmlOrderDetail In objXMLExport.objXMLExportNode.selectNodes("//order_line[orderLineKey='" & strOrderLineKey & "' and U2BonusEntitlementQty>0]")
						If Len(xmlOrderDetail.GetElementsByTagName("U2_BonusDiscountDesc")(0).text)>0 Then
							strPreviousDiscount = xmlOrderDetail.GetElementsByTagName("U2_BonusDiscountDesc")(0).text & "|" 
						Else
							strPreviousDiscount = ""
						End If
						xmlOrderDetail.GetElementsByTagName("U2_BonusDiscountDesc")(0).text=strPreviousDiscount & strBonusDesc 
					Next
					For Each xmlOrderDetail In objXMLExport.objXMLExportNode.selectNodes("//order_line[orderLineKey='" & strOrderLineKey & "' and U3BonusEntitlementQty>0]")
						If Len(xmlOrderDetail.GetElementsByTagName("U3_BonusDiscountDesc")(0).text)>0 Then
							strPreviousDiscount = xmlOrderDetail.GetElementsByTagName("U3_BonusDiscountDesc")(0).text & "|" 
						Else
							strPreviousDiscount = ""
						End If
						xmlOrderDetail.GetElementsByTagName("U3_BonusDiscountDesc")(0).text=strPreviousDiscount & strBonusDesc 
					Next
					xmlOrderHeader.removeChild xmlBonus
				Next
			lstConfig = Array(Array("order_lines","order_line",Array()))
				For Each oSection In lstConfig
					Set xmlSection = xmlOrderHeader.selectSingleNode("./" & oSection(0))
					If xmlSection Is Nothing Then
						Set xmlSection = xmlOrderHeader.AppendChild(objXMLExport.objXMLExportNode.OwnerDocument.createElement(oSection(0)))
					End If
					Set xmlItems = xmlOrderHeader.selectNodes("./"& oSection(1))
					For intItem=0 To xmlItems.length-1
						Set newXmlItem = xmlItems(intItem).cloneNode(True)
						If Ubound(oSection(2))>0 Then
							Set xmlsubSection = xmlOrderHeader.selectSingleNode("./" & oSection(2)(0))
							If xmlsubSection Is Nothing Then
								Set xmlsubSection = newXmlItem.AppendChild(objXMLExport.objXMLExportNode.OwnerDocument.createElement(oSection(2)(0)))
							End If
							Set xmlsubItems = newXmlItem.selectNodes("./"& oSection(2)(1))
							For intsubItem=0 To xmlsubItems.length-1
								Set newSubXmlItem = xmlsubItems(intsubItem).cloneNode(True)
								xmlsubSection.AppendChild newSubXmlItem
								xmlsubItems(intsubItem).parentNode.removeChild xmlsubItems(intsubItem)
							Next
						End If
						xmlSection.AppendChild newXmlItem
						xmlItems(intItem).parentNode.removeChild xmlItems(intItem)
					Next
				Next
			
			'Code to update the user if it is a telesale territory.
			'User should be the preseller mix or if there is not any food product the preseller bevarege if exists
			'if there is not any preseller the user will be the teleseller
			'We put a new code to populate the delivery territory, it seems that the expression does not work fine due to use DM
			'US 853166 Sequency delivery from SiF to HH
			'Add a new code to manage the sequence on delivery territory
				Set objTerritoryTeamID = xmlOrderHeader.GetElementsByTagName("TerritoryTeamID")
				strTerritoryTeamID = ""
				If objTerritoryTeamID.Length>0 Then
					strTerritoryTeamID = UCase(objTerritoryTeamID(0).text)
				End If
				Set objDeliveryTerritoryID = xmlOrderHeader.GetElementsByTagName("deliveryTerritoryID")
				strDeliveryTerritoryID = ""
				If objDeliveryTerritoryID.Length>0 Then
					strDeliveryTerritoryID = UCase(Trim(objDeliveryTerritoryID(0).text))
				End If
				Set objDeliveryTerritorySeq = xmlOrderHeader.GetElementsByTagName("repsec")
				strDeliveryTerritorySeq = ""
				If objDeliveryTerritorySeq.Length>0 Then
					strDeliveryTerritorySeq = UCase(Trim(objDeliveryTerritorySeq(0).text))
				End If
				If strTerritoryTeamID="TV" Or Len(strDeliveryTerritoryID)=0 Or Len(strDeliveryTerritorySeq)=0 Then
					Set objOrderHeaderKey = xmlOrderHeader.GetElementsByTagName("OrderHeaderKey")
					strOrderHeaderKey = ""
					If objOrderHeaderKey.Length>0 Then
						strOrderHeaderKey = objOrderHeaderKey(0).text
					End If
					If Len(strOrderHeaderKey)>0 Then
						Set ObjOrderHeader = System.OE_OrderHeaders.First("Key='" & strOrderHeaderKey & "'")
						If Not objOrderHeader.IsNull Then
							If Len(strDeliveryTerritoryID)=0  Or Len(strDeliveryTerritorySeq)=0 Then
								Set objDelivTerritory = objOrderHeader.Store.Folders.Territories.First("(isNull(validFrom) or validFrom<=date()) and (isNull(validTo) or ValidTo>=date()) and Territory.DM_RouteType=""REP""")
								If Not objDelivTerritory.IsNull Then
									If Len(strDeliveryTerritoryID)=0 And Not objDelivTerritory.Territory.IsNull And objDeliveryTerritoryID.Length>0 Then
										objDeliveryTerritoryID(0).text = objDelivTerritory.Territory.TerritoryName
									End If
									If Len(strDeliveryTerritorySeq)=0 And objDeliveryTerritorySeq.Length>0 Then
										objDeliveryTerritorySeq(0).text = Right("000" & System.NVL(objDelivTerritory.Cycle.Week1.Mon,""),3)
									End If
								End If
							End If
						    If Not objOrderHeader.Account.IsNull And (strTerritoryTeamID="TV" Or ObjOrderHeader.pikOrderType.EX_External.ID = "DO") Then
								If Not IsNull(objOrderHeader.OrderDate) And (Left(objOrderHeader.Account.accountNo,3)="NEX" Or Left(objOrderHeader.Account.accountNo,3)="EXT") Then
									strOrdeDate = System.GetFormatDate(objOrderHeader.OrderDate)
									strTeamId = "PVB"
									If ObjOrderHeader.Folders.OE_OrderDetails.Exists("Product.ProductCategory.EX_External.ID!='PCB'") Then
										strTeamId = "PVM"
									End If
								'20220921 AAA Pepsi Iberia - 6615157 - No territory assigned to Digital orders in export (when territory change)
								'We use TerritoriesNoFilter subfolder to have access to all territory assignments of the store
								Set objTerritory = objOrderHeader.Store.Folders.TerritoriesNoFilter.First("(isNull(validFrom) or validFrom<=" & strOrdeDate & ") and (isNull(validTo) or ValidTo>=" & strOrdeDate & ") and Territory.Team.Ex_External.ID='" & strTeamID & "' and Territory.Reps.Exists((isNull(validFrom) or validFrom<=" & strOrdeDate & ") and (isNull(validTo) or ValidTo>=" & strOrdeDate & ") and (Rep.Role.Code='PepsiIB_PS' or Rep.Role.Code='PepsiIB_PSH'))","IIF(IsNull(ValidFrom),DateSerial(2000,1,1),ValidFrom) Asc;IIF(IsNull(ValidTo),DateSerial(3000,1,1),validTo) Desc")
									If objTerritory.IsNull And strTeamId = "PVB" Then
										strTeamId = "PVM"
										Set objTerritory = objOrderHeader.Store.Folders.TerritoriesNoFilter.First("(isNull(validFrom) or validFrom<=" & strOrdeDate & ") and (isNull(validTo) or ValidTo>=" & strOrdeDate & ") and Territory.Team.Ex_External.ID='" & strTeamID & "' and Territory.Reps.Exists((isNull(validFrom) or validFrom<=" & strOrdeDate & ") and (isNull(validTo) or ValidTo>=" & strOrdeDate & ") and (Rep.Role.Code='PepsiIB_PS' or Rep.Role.Code='PepsiIB_PSH'))","IIF(IsNull(ValidFrom),DateSerial(2000,1,1),ValidFrom) Asc;IIF(IsNull(ValidTo),DateSerial(3000,1,1),validTo) Desc")
									End If 
									If Not objTerritory.IsNull Then
										Set objTerritory = objTerritory.Territory.Value
										Set objUser = objTerritory.Folders.Reps.First("(isNull(validFrom) or validFrom<=" & strOrdeDate & ") and (isNull(validTo) or ValidTo>=" & strOrdeDate & ") and (Rep.Role.Code='PepsiIB_PS' or Rep.Role.Code='PepsiIB_PSH')","IIF(IsNull(ValidFrom),DateSerial(2000,1,1),ValidFrom) ascending;IIF(IsNull(ValidTo),DateSerial(3000,1,1),validTo) descending")
										If Not objUser.IsNull Then
											Set objUser = objUser.Rep.Value
											Set objOrderHeaderRepID = xmlOrderHeader.GetElementsByTagName("RepID")
											If objOrderHeaderRepID.Length>0 Then
												objOrderHeaderRepID(0).text = objUser.ID
											End If
											Set objOrderHeaderRepLegacyID = xmlOrderHeader.GetElementsByTagName("RepLegacyID")
											If objOrderHeaderRepLegacyID.Length>0 Then
												If IsNull(objUser.DM_LegacyCode) Then
													objOrderHeaderRepLegacyID(0).text =  ""
												Else
													objOrderHeaderRepLegacyID(0).text = objUser.DM_LegacyCode
												End If
											End If
											
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			
			'Fetch config from workflow property
			For Each txtField In lstTextFields
                txtField = Trim(txtField)
                'Remove invalid char/unicode from text field
				If txtField <> "" Then
					Set objTextFieldInOrder = xmlOrderHeader.GetElementsByTagName(txtField)
					If objTextFieldInOrder.Length>0 Then
						objTextFieldInOrder(0).text = RemoveInValidCharFromString(objTextFieldInOrder(0).text)
					End If
				End If	
            Next
			
			
			blnValidateFileFlag = False
			blnSizeLimitOver = False
			If SaveXMLWithReTry (objXMLExport, strOutputValidatePath) Then 'Save XML in another temporary folder to validate it.
				'Just for testing to crash the file. Use this for debug only
				'Call MashupXMLFileToCrash(objItem, strOutputValidatePath, strLog, false) 
				If ValidateXMLFile(strOutputValidatePath,strLog, objItem) Then  'Validate if the temp2 file is OK?
					blnValidateFileFlag = True
				End If
			End If
			
			If Not blnValidateFileFlag Then 
				DeleteFileWithReTry objFSO, strOutputValidatePath 'Delete the file from temp2 if its not valid/currupted
			Else
				
				If intMaxSizePerFile>0 And Len(objXMLExport.objXMLExportNode.xml)>intMaxSizePerFile Then 'Now check if file size if overflow or not
					
					blnSizeLimitOver = True
					objXMLExport.FormatXML(objXMLExport.objXMLExportNode)
					SaveXMLWithReTry objXMLExport, strOutputPath
					
					
					'Just for testing to crash the file. Use this for debug only
					'Call MashupXMLFileToCrash(objItem, strOutputPath, strLog, true) 
				
					'After formating temp file after file size limit, We check if the format is ok? If not ok then move temp2 file to temp folder without any format. (this file is already tested before) 
					'Please note that we can not RE-FORMAT this again and again because FormatXML can be fired only once and after that it will reset XML. So we can format only once. 
					If Not ValidateXMLFile(strOutputPath,strLog, objItem) Then 
						strLog = strLog & "Formating the file is causing an issue, so moving a file without formatting" & strOutputPath & vbNewLine
						DeleteFileWithReTry objFSO, strOutputPath 'Delete the file from temp which is not valid
						OnlyMoveFile objFSO, strOutputValidatePath,strOutputPath  'Move last saved temp2 file to temp file
					End If
					
				Else
				
					'If temp2 file is okay and size is under limit then just move it to temp folder
					OnlyMoveFile objFSO, strOutputValidatePath,strOutputPath
				objTrans.Commit
				End If
			End If
			
			If blnValidateFileFlag Then 
				If blnSizeLimitOver Then 
					strLog = strLog & "File saved to " & strOutputPath & vbNewLine
					'Delete the temp2 file now as it has been completed
					DeleteFileWithReTry objFSO, strOutputValidatePath  'Delete the file from temp which is not valid
					copyMoveFile objFSO, strOutputPath, strArchivePath, strFTPFolderPath & fromSifFolder, strLog
					strExportFileName = GetTimeStampFileName(strExportFileNameTemplate, strArchivePath)
					strOutputPath = strArchivePath  & tempFolder & strExportFileName
					strOutputValidatePath = strArchivePath  & tempValidateFolder & strExportFileName
					objXMLExport.CreateExport(strMainXmlTagName)
					strLog = strLog & "New file created " & strOutputPath & vbNewLine
				End If
				
				objTrans.Commit 'Finally commit the order changes
				strLog = strLog & "SUCCESS: Updated - order: " & objItem.orderNo & " " & objItem.Prompt("FullPrompt") & vbNewLine
				intTotalUpdatedCount = intTotalUpdatedCount + 1
				intTotalExported = intTotalExported + 1
				run = True
			Else
				intTotalUpdateFailedCount = intTotalUpdateFailedCount + 1
				strLog = strLog & "File Not Saved. " & strOutputPath & vbNewLine
				strLog = strLog & "Not saved Order. XML is not valid. : " &  objItem.orderNo & " " & objItem.Prompt("FullPrompt") & vbNewLine
			
				'Remove current order object XML from the node. This order will be in R status so next time it will be covered.
				objXMLClassTag.parentNode.RemoveChild(objXMLClassTag)
				run = False
			End If
		Else
			'JM 2012-03-14 SR21-66871 added below code so if validation fail, order should not be exported.
			objXMLClassTag.parentNode.RemoveChild(objXMLClassTag)
			intTotalUpdateFailedCount = intTotalUpdateFailedCount + 1
			run = False
		End If
		'code to manage if the WF server is restarted
		If Not objProgressUI Is Nothing Then
			If objProgressUI.Halted Then
				PropertiesOut.Add "PreviousLogInfo",strLog
				PropertiesOut.Add "WF_ActionOutcome",4
			End If  
		End If
		Set objTrans = Nothing
	End Function
End Class

'This function checks if the date is in daylight saving time or not
'In Spain is from last Sunday of March until Last Sunday of October
Function DaylightTime(dtDate)
	blnDaylight = False
	dtTempDate = dtDate
	If (Month(dtDate)>3 And Month(dtDate)<10) Then
		blnDaylight=True
	ElseIf Month(dtDate)=3 Then
		blnDaylight=True
		Do
			dtTempDate = DateAdd("d",1,dtTempDate)
			If DatePart("w", dtTempDate)=1 And Month(dtTempDate)=3 Then
				blnDaylight=False
			End If
		Loop While blnDaylight And Month(dtTempDate)=3 
	ElseIf Month(dtDate)=10 Then
		blnDaylight = False
		Do
			dtTempDate = DateAdd("d",1,dtTempDate)
			If DatePart("w", dtTempDate)=1 And Month(dtTempDate)=10 Then
				blnDaylight=True
			End If
		Loop While Not blnDaylight And Month(dtTempDate)=10
	Else
		blnDaylight = False
	End If
	DaylightTime = blnDaylight
End Function


Sub MashupXMLFileToCrash(objItem, spath, strLog, blnsizelimit)
	'This is just for testing purpose
	
	
	If objItem Is Nothing Then 
		blnCrashFlag =  True
	ElseIf objItem.Key = "4FAB67A5-7642-4FE0-A81C-E85927F71853"  Or objItem.Key = "{4FAB67A5-7642-4FE0-A81C-E85927F71853}" Or blnsizelimit Then 
		blnCrashFlag =  True
	Else
		blnCrashFlag =  False
	End If	
	
	If blnCrashFlag Then 
		strLog = strLog & "CRASHING : " & spath &  vbNewLine
		Set objLogTX = objFSO.OpenTextFile(spath, 8, True, 0)
		objLogTX.WriteLine "HELOOOO"
		objLogTX.Close 
		For i = 1 To 10000
			'wait for sometime
		Next
	End If
				
End Sub

Function ValidateXMLFile(spath, strLog, objItem)
	'This function will validate XML data to prevent unicode char related issue
	ValidateXMLFile = True

	Dim objXSL : Set objXSL = CreateObject("Msxml2.DOMDocument")
	objXSL.loadXML  "<xsl:stylesheet version=""1.0"" xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"">" & _
	"<xsl:output method=""xml"" indent=""yes""/>" & _
	"<xsl:template match=""/"">" & _
	"<xsl:copy-of select="".""/>" & _
	"</xsl:template>" & _
	"</xsl:stylesheet>"
	
	Dim objXML : Set objXML = CreateObject("Msxml2.DOMDocument")
	objXML.Load spath
	objXML.transformNode objXSL
	
	If objXML.parseError.errorCode <> 0 Then
		'Just keeping below statement for future purpose so if required we can use this error.
		'ValidateXMLFile = "Parse Error line " & objXML.parseError.line & ", character " &   objXML.parseError.linePos  
		ValidateXMLFile = False
	End If
	
	Set objXSL = Nothing
	Set objXML = Nothing
	
End Function 


Function SaveXMLWithReTry(objXMLExport, strPath)
	SaveXMLWithReTry = False
		
	'We try to save the XML in one minute in case that file is locked for previous action
	intAttempts = 0
	
	Do
		intAttempts = intAttempts + 1
		On Error Resume Next
		
		If objXMLExport.Save(strPath) Then 'Save XML file
	SaveXMLWithReTry = True
		Else		
			SaveXMLWithReTry = False
		End If
	
		blnError = (Err.Number <> 0) 
		strError = ". Error: " & err.Number & ". " & err.Description
		On Error Goto 0
		If blnError Or Not SaveXMLWithReTry Then
			'Wait less than 2 secs
			dteStart = Now
			While DateDiff("s",dteStart,Now)<2
			Wend
		End If
		
	Loop Until (Not blnError) Or intAttempts>30
	
	If blnError Then
		strWFLogMessage = strWFLogMessage & Now & " *****ERROR XML could not be saved to path " & strPath & " Error." & err.Number & " > " & err.Description & vbNewLine
	Else
		SaveXMLWithReTry = True
	End If
	
End Function

Function RemoveInValidCharFromString(Str)
	'This function is used to remove unicode char which are creating an issue in XML.
	Str = System.NVL(Str,"")'Just ensure that its not null.
	RemoveInValidCharFromString = Str
	Set regExp = New RegExp
	regExp.IgnoreCase = True
	regExp.Global = True
	regExp.Pattern = strRegularExp 'Add here every character you don't consider as special character. This is configured in workflow property
	RemoveInValidCharFromString = regExp.Replace(Str, "")
End Function
