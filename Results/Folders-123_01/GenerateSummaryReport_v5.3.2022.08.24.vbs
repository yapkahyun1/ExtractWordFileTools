Dim giRow
Dim flagFolderMove
Dim WshShell, strCurDir, strParentDir, strTempDir, ObjFSO, ObjExcel, ObjWorkbook, ObjSheet, strTimeStamp, strTimeStamp1, strWorkBookFileName, strSummaryDir, strPassDir, strFailDir, strStopDir
ReDim arrSourceFolder(-1)
ReDim arrStatus(-1)
Set WshShell = CreateObject("WScript.Shell")
Set ObjFSO = CreateObject("Scripting.FileSystemObject")
Set ObjExcel = CreateObject("Excel.Application")
giRow = 2
flagFolderMove = "false"

Call Main()

Function Main()
	strTimeStamp = getDateTimeFormat
	ObjExcel.Visible = False
	Set ObjWorkbook = ObjExcel.Workbooks.Add
	Set ObjSheet = ObjWorkbook.Sheets.Item(1)
	strCurDir    = WshShell.CurrentDirectory
	strTempDir = Split(strCurDir,"\")
	strParentDir = strTempDir(0)
	For i = 1 To UBound(strTempDir) - 1
		strParentDir = strParentDir & "\" & strTempDir(i)
	Next
	If flagFolderMove = "true" Then
		strSummaryDir = createFolder(strParentDir, strTimeStamp)
		strPassDir = createFolder(strSummaryDir, "PASSED")
		strFailDir = createFolder(strSummaryDir, "FAILED")
		strStopDir = createFolder(strSummaryDir, "STOPPED")
	End If

	ObjSheet.Rows(1).Font.Size = 12
	ObjSheet.Rows(1).Font.Bold = true
	ObjSheet.Rows(1).Font.Color = RGB(255,255,255)
	For i=1 to 27
		ObjSheet.Cells(1,i).Interior.Color = RGB(139,0,139)
	Next
	
	ObjSheet.Cells(1,1) = "Module"
	ObjSheet.Cells(1,2) = "Test Case ID"
	ObjSheet.Cells(1,3) = "Mode"
	ObjSheet.Cells(1,4) = "Project Name"
	ObjSheet.Cells(1,5) = "Platform"
	ObjSheet.Cells(1,6) = "Execution Status"
	ObjSheet.Cells(1,7) = "Retest Status"
	ObjSheet.Cells(1,8) = "Execution Date"
	ObjSheet.Cells(1,9) = "Start Time"
	ObjSheet.Cells(1,10) = "End Time"
	ObjSheet.Cells(1,11) = "Duration"
	ObjSheet.Cells(1,12) = "HH"
	ObjSheet.Cells(1,13) = "MM"
	ObjSheet.Cells(1,14) = "SS"
	ObjSheet.Cells(1,15) = "Total Secs"
	ObjSheet.Cells(1,16) = "Total Failed Steps"
	ObjSheet.Cells(1,17) = "Total Passed Steps"
	ObjSheet.Cells(1,18) = "Total Step"
	ObjSheet.Cells(1,19) = "Secs/Step"
	ObjSheet.Cells(1,20) = "Category"
	ObjSheet.Cells(1,21) = "Environment"
	ObjSheet.Cells(1,22) = "Failed Description"
	ObjSheet.Cells(1,23) = "Action Item"
	ObjSheet.Cells(1,24) = "Previous Cycle Result"
	ObjSheet.Cells(1,25) = "PIC"
	ObjSheet.Cells(1,26) = "Status"
	ObjSheet.Cells(1,27) = "Tracker"
	
    
	Call ShowSubFolders(strCurDir, ObjFSO, ObjSheet, strPassDir, strFailDir)
	If flagFolderMove = "true" Then
		strWorkBookFileName = strSummaryDir & "\Summary Report " & strTimeStamp & ".xlsx"
	Else
		strWorkBookFileName = strCurDir & "\" & strTimeStamp & ".xlsx"

	End If
	
	ObjSheet.Cells(giRow + 2,1).Font.Bold = true
	ObjSheet.Cells(giRow + 2,1) = "Total Passed"
	ObjSheet.Cells(giRow + 3,1).Font.Bold = true
	ObjSheet.Cells(giRow + 3,1) = "Total Failed"
	ObjSheet.Cells(giRow + 4,1).Font.Bold = true
	ObjSheet.Cells(giRow + 4,1) = "Total Stopped"
	ObjSheet.Cells(giRow + 5,1).Font.Bold = true
	ObjSheet.Cells(giRow + 5,1) = "Total Executed"
	
	pCount = 0
	fCount = 0
	sCount = 0
	For i = 2 to giRow
		If NOT(ObjSheet.Cells(i,1).value ="Pre-Requisite") Then
			If ObjSheet.Cells(i,6) = "Passed" Then
				pCount = pCount + 1
			ElseIf ObjSheet.Cells(i,6) = "Failed" Then
				fCount = fCount + 1
			ElseIf ObjSheet.Cells(i,6) = "Stopped" Then
				sCount = sCount + 1
			End If
		End If
	Next
	ObjSheet.Cells(giRow + 2,2) = pCount
	ObjSheet.Cells(giRow + 3,2) = fCount
	ObjSheet.Cells(giRow + 4,2) = sCount
	ObjSheet.Cells(giRow + 5,2) = (fCount + pCount + sCount)
	
	ObjSheet.Columns("A:AA").AutoFit
		
	ObjWorkbook.SaveAs( strWorkBookFileName  )
	ObjWorkbook.Close
	ObjExcel.Quit

	Set WshShell = Nothing
	Set ObjSheet = Nothing
	Set ObjWorkbook = Nothing
	Set ObjExcel = Nothing
	
	If flagFolderMove = "true" Then
		For i = 0 to Ubound(arrSourceFolder)
			If arrStatus(i) = "PASS" Then
				strPassDir = strPassDir & "\"
				ObjFSO.MoveFolder arrSourceFolder(i), strPassDir
			ElseIf arrStatus(i) = "FAIL" Then
				strFailDir = strFailDir & "\"
				ObjFSO.MoveFolder arrSourceFolder(i), strFailDir
			ElseIf arrStatus(i) = "STOP" Then
				strStopDir = strStopDir & "\"
				ObjFSO.MoveFolder arrSourceFolder(i), strStopDir			
			End If
		Next
	End If

	Set ObjFSO = Nothing
	MsgBox "Generate Summary Report Completed!" & vbNewLine & _
	"Saved To : " & vbNewLine & strWorkBookFileName & vbNewLine & _
	"Total Executed : " & (pCount + fCount + sCount) & vbNewLine & _
	"Total Passed : " & pCount & vbNewLine & _
	"Total Failed : " & fCount & vbNewLine & _
	"Total Stopped : " & sCount
	'strTimeStamp1 = getDateTimeFormat
	'Msgbox "StartTime: " & strTimeStamp &vbNewLine & "EndTime: " & strTimeStamp1
	
End Function

Function ShowSubFolders(strCurDir , ObjFSO, ObjSheet, strPassDir, strFailDir)
	Set Folder = ObjFSO.GetFolder(strCurDir)
	
    For Each Subfolder in Folder.SubFolders
        Set objFolder = ObjFSO.GetFolder(Subfolder.Path)
        Set colFiles = objFolder.Files
        For Each objFile in colFiles 
			'If LCase(ObjFSO.GetExtensionName(objFile.Name)) = "htm" > -1 and InStr(objFile.Name,"report.log") = false and InStr(objFile.Name,"sub.uft.rpt") = false Then
			If LCase(ObjFSO.GetExtensionName(objFile.Name)) = "html" Then
				'msgbox objFile.Name
				currFilePathName = ObjFSO.GetAbsolutePathName(objFile)
				Set f = ObjFSO.OpenTextFile(currFilePathName)
				StrAllData = f.ReadAll
				'msgbox Instr(StrAllData, "Total Passed Cases")
				If Instr(StrAllData, "Total Passed Cases") = 0 Then
					flag = true
				End If
				f.Close
				'msgbox flag
				If flag Then
					currFilePathName = ObjFSO.GetAbsolutePathName(objFile)
					Set f = ObjFSO.OpenTextFile(currFilePathName)
					do while not f.AtEndOfStream
						StrRawData = f.ReadLine
						If Instr(StrRawData,"Execution Details") Then
							StrData1 = Split(StrRawData, "Execution Details")(1)
							StrData2 = Split(StrData1, "Test Step Description")(0)
							StrLoop = Split(StrData2,"<")
						End If
						If Instr(StrRawData,"label label-danger col-sm-12 col-xs-12") Then
							StrData3 = Split(StrRawData, "label label-danger col-sm-12 col-xs-12")(1)
							StrLoop2 = Split(StrData3,"<")
						End If
					loop
					For i = 0 to (Ubound(StrLoop)-1)
						If InStr(StrLoop(i), "Total Time") Then
							strTotalTime = Split(StrLoop(i+2),">")(1)
							'msgbox strTotalTime
						ElseIf InStr(StrLoop(i), "Project Name") Then
							strProjectName = Split(StrLoop(i+1),">")(1)
							'msgbox strProjectName
						ElseIf InStr(StrLoop(i), "Test Script ID") Then
							strTCID = Split(StrLoop(i+1),">")(1)
							'msgbox strTCID
						ElseIf InStr(StrLoop(i), "App Version") Then
							strPlatform = Split(StrLoop(i+1),">")(1)
							'msgbox strPlatform
						ElseIf InStr(StrLoop(i), "Run Date") Then
							strRunDate = Split(StrLoop(i+1),">")(1)
							'msgbox strRunDate
						ElseIf InStr(StrLoop(i), "Run Start") Then
							strRunStart = Split(StrLoop(i+1),">")(1)
							'msgbox strRunStart
						ElseIf InStr(StrLoop(i), "Run Ended") Then
							strRunEnd = Split(StrLoop(i+1),">")(1)
							'msgbox strRunEnd
						ElseIf InStr(StrLoop(i), "Execution Status") Then
							strStatus = Split(StrLoop(i+1),">")(1)
							'msgbox strStatus
						End If
					Next

					For i = 0 to (Ubound(StrLoop2)-1)
						If InStr(StrLoop2(i), "Total Failed Steps") Then
							strTotalFailed = Split(StrLoop2(i+1),">")(1)
							'msgbox strTotalFailed
						ElseIf InStr(StrLoop2(i), "Total Passed Steps") Then
							strTotalPassed = Split(StrLoop2(i+1),">")(1)
							'msgbox strTotalPassed
						ElseIf InStr(StrLoop2(i), "Total Steps") Then
							strTotalSteps = Split(StrLoop2(i+1),">")(1)
							'msgbox strTotalSteps
						End If
					Next
					
					If InStr(strTCID,"SUB_") > 0 Then
						ObjSheet.Cells(giRow,1) = "Pre-Requisite"
					Else
						ObjSheet.Cells(giRow,1) = Split(strTCID,"_")(1)
					End If
					ObjSheet.Cells(giRow,2) = strTCID
					ObjSheet.Cells(giRow,2).Interior.Color = RGB(204,153,255)
					ObjSheet.Cells(giRow,4) = strProjectName
					ObjSheet.Cells(giRow,5) = strPlatform
					ObjSheet.Cells(giRow,6).Font.Bold = true
					ObjSheet.Cells(giRow,6) = strStatus
					ObjSheet.Cells(giRow,8) = strRunDate
					ObjSheet.Cells(giRow,9) = strRunStart
					ObjSheet.Cells(giRow,10) = strRunEnd
					ObjSheet.Cells(giRow,11) = strTotalTime
					ObjSheet.Cells(giRow,16) = strTotalFailed
					ObjSheet.Cells(giRow,17) = strTotalPassed
					ObjSheet.Cells(giRow,18) = strTotalSteps
					
					strTotalTime2 = Split(strTotalTime, " ")
					
					ObjSheet.Cells(giRow, 12) = CInt(Replace(strTotalTime2(0),"hrs",""))
					ObjSheet.Cells(giRow, 13) = CInt(Replace(strTotalTime2(1),"mins",""))
					ObjSheet.Cells(giRow, 14) = CInt(Replace(strTotalTime2(2),"secs",""))
					ObjSheet.Cells(giRow, 15) = (ObjSheet.Cells(giRow, 12).Value*3600) + (ObjSheet.Cells(giRow, 13).Value*60) + ObjSheet.Cells(giRow, 14)
					ObjSheet.Cells(giRow, 19) = Round(ObjSheet.Cells(giRow, 15).Value / CInt(strTotalSteps), 1)
					

						If StrComp(strStatus, "Passed") = 0 Then
							ObjSheet.Cells(giRow, 6).Interior.Color = RGB(0,255,0)
							ReDim Preserve arrSourceFolder(UBound(arrSourceFolder) + 1)
							arrSourceFolder(UBound(arrSourceFolder)) = strCurDir
							ReDim Preserve arrStatus(UBound(arrStatus) + 1)
							arrStatus(UBound(arrStatus)) = "PASS"
						ElseIf StrComp(strStatus, "Failed") = 0 Then
							ObjSheet.Cells(giRow, 6).Interior.Color = RGB(255,0,0)
							ReDim Preserve arrSourceFolder(UBound(arrSourceFolder) + 1)
							arrSourceFolder(UBound(arrSourceFolder)) = strCurDir
							ReDim Preserve arrStatus(UBound(arrStatus) + 1)
							arrStatus(UBound(arrStatus)) = "FAIL"
						ElseIf StrComp(strStatus, "Stopped") = 0 Then
							ObjSheet.Cells(giRow, 6).Interior.Color = RGB(0,0,255)
							ReDim Preserve arrSourceFolder(UBound(arrSourceFolder) + 1)
							arrSourceFolder(UBound(arrSourceFolder)) = strCurDir
							ReDim Preserve arrStatus(UBound(arrStatus) + 1)
							arrStatus(UBound(arrStatus)) = "STOP"
						End If
					
					f.Close
					Set f = Nothing
					giRow = giRow + 1
					Exit For
				End If
			End If
        Next
		Call ShowSubFolders(Subfolder, ObjFSO, ObjSheet, strPassDir, strFailDir)        
    Next
End Function

Function getDateTimeFormat()
	strNow = Now
	strDateTime = ""
	strYear = Year(strNow)
	strMonth= Month(strNow)
	If strMonth < 10 Then
		strMonth = "0" & strMonth
	End If
	strDate = Day(strNow)
	If strDate < 10 Then
		strDate = "0" & strDate
	End If
	strHour = Hour(strNow)
	If strHour < 10 Then
		strHour = "0" & strHour
	End If
	strMinute = Minute(strNow)
	If strMinute < 10 Then
		strMinute = "0" & strMinute
	End If
	strSeconds = Second(strNow)
	If strSeconds < 10 Then
		strSeconds = "0" & strSeconds
	End If
	strDateTime = strYear & strMonth & strDate & "_" & strHour & strMinute & strSeconds
	getDateTimeFormat = strDateTime
End Function

Function createFolder(strDirPath, strFolderName)
	Set WshShell = CreateObject("WScript.Shell")
	Set ObjFSO = CreateObject("Scripting.FileSystemObject")
	createFolder = objFSO.CreateFolder(strDirPath & "\" & strFolderName)
End Function