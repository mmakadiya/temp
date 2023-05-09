		Line  409: 				objSystemLog.LogNormalDetail "The process started at " & Now() & "."
	Line  425: 					objSystemLog.LogNormalDetail "The ""imports"" folder is executing."
	Line  430: 					objSystemLog.LogNormalDetail "The ""imports"" folder is executed successfully."
	Line  436: 					objSystemLog.LogNormalDetail "The ""workflows"" folder is executing."
	Line  460: 					objSystemLog.LogNormalDetail "The ""workflows"" folder is executed successfully."
	Line  465: 					objSystemLog.LogNormalDetail "The ""dashboardparts"" folder is executing."
	Line  472: 					objSystemLog.LogNormalDetail "The ""dashboardparts"" folder is executed successfully."
	Line  477: 					objSystemLog.LogNormalDetail "The ""dashboardtemplates"" folder is executing."
	Line  484: 					objSystemLog.LogNormalDetail "The ""dashboardtemplates"" folder is executed successfully."
	Line  489: 					objSystemLog.LogNormalDetail "The ""dashboardtemplatespos"" folder is executing."
	Line  496: 					objSystemLog.LogNormalDetail "The ""dashboardtemplatespos"" folder is executed successfully."
	Line  501: 					objSystemLog.LogNormalDetail "The ""dashboards"" folder is executing."
	Line  508: 					objSystemLog.LogNormalDetail "The ""dashboards"" folder is executed successfully."
	Line  513: 					objSystemLog.LogNormalDetail "The ""KPIs"" folder is executing."
	Line  520: 					objSystemLog.LogNormalDetail "The ""KPIs"" folder is executed successfully."
	Line  525: 					objSystemLog.LogNormalDetail "The ""Views"" folder is executing."
	Line  532: 					objSystemLog.LogNormalDetail "The ""Views"" folder is executed successfully."
	Line  537: 					objSystemLog.LogNormalDetail "The ""ViewGroups"" folder is executing."
	Line  544: 					objSystemLog.LogNormalDetail "The ""ViewGroups"" folder is executed successfully."
	Line  549: 					objSystemLog.LogNormalDetail "The ""PushReports"" folder is executing."
	Line  556: 					objSystemLog.LogNormalDetail "The ""PushReports"" folder is executed successfully."
	Line  561: 					objSystemLog.LogNormalDetail "The ""ReportTemplates"" folder is executing."
	Line  568: 					objSystemLog.LogNormalDetail "The ""ReportTemplates"" folder is executed successfully."
	Line  573: 					objSystemLog.LogNormalDetail "The ""SummaryTemplates"" folder is executing."
	Line  580: 					objSystemLog.LogNormalDetail "The ""SummaryTemplates"" folder is executed successfully."
	Line  585: 					objSystemLog.LogNormalDetail "The ""DynamicMemberDefinitions"" folder is executing."
	Line  592: 					objSystemLog.LogNormalDetail "The ""DynamicMemberDefinitions"" folder is executed successfully."
	Line  597: 					objSystemLog.LogNormalDetail "The ""Analytics"" folder is executing."
	Line  604: 					objSystemLog.LogNormalDetail "The ""Analytics"" folder is executed successfully."
	Line  611: 					objSystemLog.LogNormalDetail "The ""importCSVs"" folder is executing."
	Line  638: 									objSystemLog.LogNormalDetail "See error log for error details "
	Line  639: 									objSystemLog.LogErrorDetail "","", "there are some error while Error deleting:" & Name & " - " & Err.Description
	Line  650: 							objSystemLog.LogErrorDetail "","", "there are some error while process import file  "
	Line  654: 							objSystemLog.LogNormalDetail "See error log for error details "
	Line  655: 							objSystemLog.LogErrorDetail "","", "there are some error while process import file  " & Err.Description
	Line  659: 					objSystemLog.LogNormalDetail "The ""importCSVs"" folder is executed successfully."
	Line  676: 					objSystemLog.LogNormalDetail "The ""touchevents"" folder is executing."
	Line  684: 					objSystemLog.LogNormalDetail "The ""touchevents"" folder is executed successfully."
	Line  689: 			objSystemLog.LogJobComplete "The process completed at " & Now() & "."
	Line  718: 					objSystemLog.LogNormalDetail "See error log for error details "
	Line  719: 					objSystemLog.LogErrorDetail "","", "there are some error while Error deleting:" & Name & " - " & Err.Description
	Line  872: 		objSystemLog.LogNormalDetail "Started execution to enable the touch event."
	Line  897: 														objSystemLog.LogErrorDetail "","","Error : " & Err.Description
	Line  905: 													objSystemLog.LogNormalDetail """"& strEvent &""" touch event has been enabled for the role " & strRoleCode & "."
	Line  907: 													objSystemLog.LogErrorDetail "","","'" & objval.Result.Message & "'" 
	Line  915: 														objSystemLog.LogNormalDetail """"& strEvent &""" touch event has been disabled for the role " & strRoleCode & "."
	Line  917: 														objSystemLog.LogErrorDetail "","","There were some error in disabling the touch event : '" & objval.Result.Message & "'" 
	Line  920: 													objSystemLog.LogNormalDetail """"& strEvent &""" touch event is already disabled for the role " & strRoleCode & "."
	Line  926: 										objSystemLog.LogErrorDetail "","", "The " & strEvent & " touch event was not found."
	Line  934: 						objSystemLog.LogErrorDetail "","", "The " & ObjRoleCode & "role code was not found."
	Line  939: 		objSystemLog.LogNormalDetail "Completed execution to enable the touch event."
	Line 1148: 							objSystemLog.LogNormalDetail "Error occurred while saving Workflow. Please check the Error Log."
	Line 1149: 							ObjSystemLog.LogErrorDetail "","","Error : " & Err.Description
	Line 1151: 							objSystemLog.LogNormalDetail "Added or Updated for the Workflow/" & objInstance.Title
	Line 2724: 		objSystemLog.LogNormalDetail "Started execution to update the workflow members."
	Line 2736: 							objSystemLog.LogErrorDetail "","", "Invalid line : " & strCode
	Line 2749: 									objSystemLog.LogNormalDetail "The '"& strWFPrompt &"' workflow has been updated successfully."
	Line 2751: 									objSystemLog.LogErrorDetail "","","There were some error in updating the '"& strWFPrompt &"' workflow members : '" & objval.Result.Message & "'" 
	Line 2756: 							objSystemLog.LogErrorDetail "","", "The workflow does not exists : " & strFolderScanFilter
	Line 2762: 		objSystemLog.LogNormalDetail "Completed execution to update the workflow members."
	Line 2772: 	objSystemLog.LogNormalDetail "The touch dashboard part loading process started."
	Line 2786: 					objSystemLog.LogErrorDetail "","", "File is empty : " & objFile.Name
	Line 2815: 										objSystemLog.LogErrorDetail "","", "Touch Dashboard Part Not Exists in system. Wrong File Name Format to create new dashboard part : " & strFileName
	Line 2872: 											objSystemLog.LogNormalDetail "The '"& strDBPartTitle &"' touch dashboard part has been created/updated."
	Line 2875: 											objSystemLog.LogErrorDetail "","","There were some error in updating/creating the '"& strDBPartTitle &"' touch dashboard part : '" & objval.Result.Message & "'" 
	Line 2879: 										objSystemLog.LogErrorDetail "","", "Something Wrong : " & strFileName
	Line 2886: 								objSystemLog.LogErrorDetail "","", "Wrong File Name Format : " & strFileName
	Line 2890: 							objSystemLog.LogErrorDetail "","", "Wrong File Name Format : " & strFileName
	Line 2893: 						objSystemLog.LogErrorDetail "","", "File is empty : " & objFile.Name
	Line 2900: 	objSystemLog.LogNormalDetail "The touch dashboard part loading process completed."
	Line 2913: 	objSystemLog.LogNormalDetail "The touch dashboard template loading process started."
	Line 2928: 					objSystemLog.LogErrorDetail "","", "File is empty : " & objFile.Name
	Line 2966: 								objSystemLog.LogNormalDetail "The '"& strDBTemplateKey &"' touch dashboard template has been created/updated."
	Line 2969: 								objSystemLog.LogErrorDetail "","","There were some error in updating/creating the '"& strDBTemplateKey &"' touch dashboard template : '" & objval.Result.Message & "'" 
	Line 2973: 							objSystemLog.LogErrorDetail "","", "Something Wrong : " & strDBTemplateKey
	Line 2978: 						objSystemLog.LogErrorDetail "","", "File is empty : " & objFile.Name
	Line 2983: 				objSystemLog.LogErrorDetail "","",  "File doesn't exists : " & strCodeFilePath
	Line 2988: 	objSystemLog.LogNormalDetail "The touch dashboard template loading process completed."
	Line 2994: 	objSystemLog.LogNormalDetail "The touch dashboard template Position loading process started."
	Line 3008: 					objSystemLog.LogErrorDetail "","", "File is empty : " & objFile.Name
	Line 3049: 										objSystemLog.LogErrorDetail "","", "Something Wrong : Template - '"& arrTokens(0) &"', Position - '"& strTag &"'"
	Line 3057: 								objSystemLog.LogNormalDetail "The '"& arrTokens(0) &"' touch dashboard template's position(s) have been created/updated."
	Line 3060: 								objSystemLog.LogErrorDetail "","","There were some error in updating/creating the '"& arrTokens(0) &"' touch dashboard template's Position(s) : '" & objval.Result.Message & "'" 
	Line 3063: 							objSystemLog.LogErrorDetail "","", "Touch Dashboard Template doesn't exist in system : " & arrTokens(0)
	Line 3067: 						objSystemLog.LogErrorDetail "","", "File is empty : " & objFile.Name
	Line 3073: 				objSystemLog.LogErrorDetail "","", "File doesn't exists : " & strCodeFilePath
	Line 3078: 	objSystemLog.LogNormalDetail "The touch dashboard template Position loading process Completed."
	Line 3084: 	objSystemLog.LogNormalDetail "The touch dashboard loading process started."
	Line 3099: 					objSystemLog.LogErrorDetail "","", "File is empty : " & objFile.Name
	Line 3126: 										objSystemLog.LogErrorDetail "","",  "The '"& strTemplateName &"' touch dashboard template does not exists in system."
	Line 3130: 									objSystemLog.LogErrorDetail "","",  "The '"& strDBKey &"' touch dashboard doesn't exist in system and config file to create new touch dashboard doesn't contain sufficient information."
	Line 3134: 								objSystemLog.LogErrorDetail "","",   "The '"& strDBKey &"' touch dashboard doesn't exist in system and config file to create new touch dashboard doesn't contain sufficient information."
	Line 3186: 									objSystemLog.LogNormalDetail "The touch dashboard part doesn't exists in system : " & arrParts(0) 
	Line 3199: 									objSystemLog.LogNormalDetail "The '"& arrTokens(0) &"' touch dashboard has been created/updated."
	Line 3202: 									objSystemLog.LogErrorDetail "","","There were some error in updating/creating the '"& arrTokens(0) &"' touch dashboard : '" & objval.Result.Message & "'" 
	Line 3206: 								objSystemLog.LogErrorDetail "","", "Something wrong : " & strDBKey
	Line 3210: 							objSystemLog.LogErrorDetail "","", "Something wrong : " & strDBKey
	Line 3213: 						objSystemLog.LogErrorDetail "","",  "File is empty : " & objFile.Name
	Line 3219: 				objSystemLog.LogErrorDetail "","", "File doesn't exists : " & strCodeFilePath
	Line 3224: 	objSystemLog.LogNormalDetail "The touch dashboard loading process completed."
	Line 3230: 	objSystemLog.LogNormalDetail "The KPIs loading process started."
	Line 3245: 					objSystemLog.LogErrorDetail "","",  "File is empty : " & objFile.Name
	Line 3297: 								objSystemLog.LogNormalDetail "The '"& strKPIKey &"' KPI has been created/updated."
	Line 3299: 								objSystemLog.LogErrorDetail "","","There were some error in updating/creating the '"& strKPIKey &"' KPI : '" & objval.Result.Message & "'" 
	Line 3311: 	objSystemLog.LogNormalDetail "The KPIs loading process Completed."
	Line 3318: 	objSystemLog.LogNormalDetail "The Views loading process started."
	Line 3333: 					objSystemLog.LogErrorDetail "","",  "File is empty : " & objFile.Name
	Line 3380: 								objSystemLog.LogNormalDetail "The '"& strViewKey &"' View has been created/updated."
	Line 3382: 								objSystemLog.LogErrorDetail "","","There were some error in updating/creating the '"& strViewKey &"' View : '" & objval.Result.Message & "'" 
	Line 3394: 	objSystemLog.LogNormalDetail "The Views loading process Completed."
	Line 3401: 	objSystemLog.LogNormalDetail "The Views loading process started."
	Line 3416: 					objSystemLog.LogErrorDetail "","",  "File is empty : " & objFile.Name
	Line 3450: 								objSystemLog.LogNormalDetail "The '"& strViewGroupKey &"' ViewGroup has been created/updated."
	Line 3452: 								objSystemLog.LogErrorDetail "","","There were some error in updating/creating the '"& strViewGroupKey &"' ViewGroup : '" & objval.Result.Message & "'" 
	Line 3464: 	objSystemLog.LogNormalDetail "The ViewGroups loading process Completed."
	Line 3471: 	objSystemLog.LogNormalDetail "The Report Template loading process started."
	Line 3486: 					objSystemLog.LogErrorDetail "","", "File is empty : " & objFile.Name
	Line 3533: 								objSystemLog.LogNormalDetail "The '"& strReportTemplateskey &" has been created/updated."
	Line 3536: 								objSystemLog.LogErrorDetail "","","There were some error in updating/creating the '"& strReportTemplateskey &" : '" & objval.Result.Message & "'" 
	Line 3556: 	objSystemLog.LogNormalDetail "The Summary Template loading process started."
	Line 3571: 					objSystemLog.LogErrorDetail "","", "File is empty : " & objFile.Name
	Line 3611: 								objSystemLog.LogNormalDetail "The '"& strSummaryTemplateskey &" has been created/updated."
	Line 3614: 								objSystemLog.LogErrorDetail "","","There were some error in updating/creating the '"& strSummaryTemplateskey &" : '" & objval.Result.Message & "'" 
	Line 3633: 	objSystemLog.LogNormalDetail "The Summary Template loading process started."
	Line 3648: 					objSystemLog.LogErrorDetail "","", "File is empty : " & objFile.Name
	Line 3723: 								objSystemLog.LogNormalDetail "The '"& strDynamicMemberDefinitionskey &" has been created/updated."
	Line 3726: 								objSystemLog.LogErrorDetail "","","There were some error in updating/creating the '"& strDynamicMemberDefinitionskey &" : '" & objval.Result.Message & "'" 
	Line 3930: 	objSystemLog.LogNormalDetail "The Push Reports loading process started."
	Line 3945: 					objSystemLog.LogErrorDetail "","", "File is empty : " & objFile.Name
	Line 4004: 										objSystemLog.LogErrorDetail "","","There were some error in Update Summary '" &  strMessage & "' in PushReport '"& strPushReportskey &"'"
	Line 4015: 								objSystemLog.LogNormalDetail "The '"& strPushReportskey &" has been created/updated."
	Line 4018: 								objSystemLog.LogErrorDetail "","","There were some error in updating/creating the '"& strPushReportskey &" : '" & objval.Result.Message & "'" 
	Line 4037: 	objSystemLog.LogNormalDetail "The touch events loading process started."
	Line 4053: 					objSystemLog.LogErrorDetail "","", "File is empty : " & objFile.Name
	Line 4090: 								objSystemLog.LogNormalDetail "The '"& arrParts(0) &"' "& arrParts(1) &" has been created/updated."
	Line 4093: 								objSystemLog.LogErrorDetail "","","There were some error in updating/creating the '"& arrParts(0) &"' "& arrParts(1) &" : '" & objval.Result.Message & "'" 
	Line 4112: 	objSystemLog.LogNormalDetail "The Analytic Views loading process started."
	Line 4127: 					objSystemLog.LogErrorDetail "","", "File is empty : " & objFile.Name
	Line 4211: 								objSystemLog.LogNormalDetail "The '"& strAnalyticKey &"' analytic view has been created/updated."
	Line 4217: 								objSystemLog.LogErrorDetail "","","There were some error in updating/creating the '"& strAnalyticKey &"' analytic view : '" & objval.Result.Message & "'" 
	Line 4229: 	objSystemLog.LogNormalDetail "The Analytic Views loading process completed."
	Line 4246: 							objSystemLog.LogNormalDetail "Process started for file '" & objFile.Name & "'."
	Line 4309: 												objSystemLog.LogErrorDetail "","", "TLF '"& TLFname &"' dose not Exist in system."
	Line 4346: 																			objSystemLog.LogErrorDetail "","", strdataline &"Parent Key Filter error in expression: " & strFetchParentFilter
	Line 4350: 																			objSystemLog.LogNormalDetail "See error log for data line '" & strdataline & "'"
	Line 4351: 																			objSystemLog.LogErrorDetail "","",strdataline & "Parent Key Filter " & Err.Description
	Line 4389: 																				objSystemLog.LogErrorDetail "","", "Not all Key Mapping Members Exist in found in File"
	Line 4393: 																			objSystemLog.LogErrorDetail "","",  strdataline &" No SubFolder with name " &  SubFolder &" found in TLF "& TLFname
	Line 4397: 																		objSystemLog.LogErrorDetail "","",  strdataline &" No Record found in TLF "& TLFname &" Parent Key Filter " & Err.Description
	Line 4423: 																	objSystemLog.LogErrorDetail "","", "Not all Key Mapping Members Exist in found in File"
	Line 4430: 															objSystemLog.LogErrorDetail "","", "Not all members values are available in line" &strdataline
	Line 4438: 												objSystemLog.LogErrorDetail "","", "Not all Parent Key Mapping Members Exist in found in File"
	Line 4446: 										objSystemLog.LogNormalDetail "No Data Exist to import in File"
	Line 4451: 									objSystemLog.LogErrorDetail "","", "TLF '"& TLFname &"' dose not Exist in system"
	Line 4461: 								objSystemLog.LogErrorDetail "","", "Please make sure file must have Config columns 'ToplevelFolder','CanCreateNew','KeyMapping'"
	Line 4464: 							objSystemLog.LogNormalDetail "Process completed for file '" & objFile.Name & "'."
	Line 4496: 			objSystemLog.LogErrorDetail "","", "Folder with name " & strFolder &" not Exits."
	Line 4594: 			objSystemLog.LogNormalDetail "See error log for data line '" & strdataline & "'"
	Line 4595: 			objSystemLog.LogErrorDetail "","",strdataline &"there are more then one record in system with filter " & strFetchFilter &"please update KeyMapping to get unique record"
	Line 4604: 				objSystemLog.LogErrorDetail "","", strdataline & " Key Filter error in expression: " & strFetchFilter
	Line 4608: 				objSystemLog.LogNormalDetail "See error log for data line '" & strdataline & "'"
	Line 4609: 				objSystemLog.LogErrorDetail "","",strdataline &" Key Filter " & Err.Description
	Line 4647: 					objSystemLog.LogErrorDetail "","", "System not allow to create new instance in TLF: " & objFolder.name
	Line 4651: 				objSystemLog.LogErrorDetail "","", "Configuration Error. No suitable to create new instance in TLF: " & objFolder.name
	Line 4690: 					objSystemLog.LogErrorDetail "","", strdataline & " for set join member " & objMember.definition.title.value &" Not able to find the record  " &StrFilter & "in Folder " & objMember.targetfolder.title.value
	Line 4707: 						objSystemLog.LogErrorDetail "","", strdataline & " for set join member " & objMember2.definition.title.value &" Not able to find the record  " &StrFilter & "in Folder " & objMember2.targetfolder.title.value
	Line 4725: 						objSystemLog.LogErrorDetail "","", "There were some error occurred while setting the value to the member '" & StrMembers & "'."
	Line 4726: 						objSystemLog.LogErrorDetail "","",StrMembers & " - Error: " & Err.Description
	Line 4736: 				objSystemLog.LogErrorDetail "","", "Please Provide join member 'Key Mapping Members' value for set join member "
	Line 4741: 		objSystemLog.LogErrorDetail "","", "There were some error occurred while setting the value to the member '" & StrMembers & "'."
	Line 4742: 		objSystemLog.LogErrorDetail "","",StrMembers & " - Error: " & Err.Description
	Line 4774: 				objSystemLog.LogErrorDetail "","", strdataline & """Key Mapping Members value must not be null"""
	Line 4798: 					objSystemLog.LogErrorDetail  "","", "Error while Evaluate Member " & arrMembers(i) & " Expression Error: " & strErrors
	Line 4816: 		objSystemLog.LogNormalDetail "See error log for data line '" & strdataline & "'"
	Line 4817: 		objSystemLog.LogErrorDetail "","",strdataline & " Error : " & objval.Result.Message
	Line 5208: 		ObjSystemLog.LogErrorDetail "","", strLogMessage
	Line 5214: 		objSystemLog.LogNormalDetail strLogMessage