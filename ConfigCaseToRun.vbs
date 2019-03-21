
UpdateCaseToRun


Function UpdateCaseToRun()
	Dim objExcelApp, TestCaseSheet, TestCaseRowCount,  i, j
	' Create Excel Object
	Set objExcelApp = CreateObject("Excel.Application")
	
	'set file attribute to normal
	Set fso = CreateObject( "Scripting.FileSystemObject" )
	fso.GetFile (GetCurrentPath + "\TestCase.xlsx").Attributes =0
	Set fso = Nothing 
	
'	objExcelApp.Visible = true
	objExcelApp.Workbooks.Open GetCurrentPath + "\TestCase.xlsx"

    
'Locate TestCase sheet and its total count
	Set TestCaseSheet = objExcelApp.Sheets.Item("TestCase")
	TestCaseRowCount = TestCaseSheet.UsedRange.Rows.Count

    'Loop for row
	For i = 1 to TestCaseRowCount 
	
	'	judge whether the Module run or not
		If  TestCaseSheet.Cells(i, 2) = "Y" Then
    		TestCaseSheet.Cells(i, 2) = "N"
    	ElseIf TestCaseSheet.Cells(i, 2) = "N" Then
    		TestCaseSheet.Cells(i, 2) = "Y"
		End If
	
	Next
	
	 objExcelApp.DisplayAlerts = false
	 objExcelApp.Save
	 objExcelApp.Quit
	 
	Set objExcelApp = Nothing
	MsgBox "Config test cases to run is done"
	
End Function
	
Function GetCurrentPath()
	GetCurrentPath = left(wscript.scriptfullname, instrrev(wscript.scriptfullname, "\") - 1)
End Function