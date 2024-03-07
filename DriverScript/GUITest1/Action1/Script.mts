''Create Contact
'Call fn_LaunchandLogin("Chrome.exe;https://login.salesforce.com/;maruthit39-t2fv@force.com;sfdcsfdc@3")
'Call fn_CreateContact("SFDC_SYT;SFDC_Cont")
'Call fn_Logout()
'
''Create Opportunity
'Call fn_LaunchandLogin("Chrome.exe;https://login.salesforce.com/;maruthit39-t2fv@force.com;sfdcsfdc@3")
'Call fn_CreateOpportunity("SFDC_SYT;SFDC_Opp;Qualification")
'Call fn_Logout()
'
''Update Opportunity
'Call fn_LaunchandLogin("Chrome.exe;https://login.salesforce.com/;maruthit39-t2fv@force.com;sfdcsfdc@3")
'Call fn_SearchandOpenanyRecord("Sfdc_opp;Opportunities")
'Call fn_ClickButtonsonDetailsPage("Edit")
'Call fn_EditOpportunity("")
'Call fn_Logout()
'
'*********************************************************************************************
'Action Name	: Driver Script
'Created by		: Maruthi
'Description		: This Driverscript loads the test cases, teststeps, testdata into Datatable and connect with associated actions to invoke keywords
'Inputs			: If any new action created that action need to be added here...
'*********************************************************************************************
Dim sTestCaseID	'current test case ID 
Dim sStepResult	'result of test step
Datatable.AddSheet "TCases"
Datatable.AddSheet "TSteps"
'DataTable.Addsheet "DataInput"
'Associating Function Library
LoadFunctionLibrary "D:\Automation\FunctionLibrary\Salesforce.vbs"
'Import 2 sheets in UFT Datatable
Datatable.ImportSheet "D:\Automation\TestCaseOrganiser\Testcases.xlsx","TCases","TCases"
Datatable.ImportSheet "D:\Automation\TestCaseOrganiser\Testcases.xlsx","TSteps","TSteps"
'Count total no.of rows in Datatable
iTestCases = Datatable.GetSheet("TCases").GetRowCount
For i = 1 To iTestCases
	Datatable.GetSheet("TCases").SetCurrentRow(i)
	If UCase(Datatable.Value("Execute","TCases")) = "Y" Then
		sTestCaseID = Datatable.GetSheet("TCases").GetParameter("TestCaseID")		'Get Testcase ID
		iTestSteps = Datatable.GetSheet("TSteps").GetRowCount	'Total row count of Tsteps sheet
		Reporter.ReportEvent micPass, "Test case Started", "Test Case: "&sTestCaseID&" is Started"
		'Launch Application
		For j = 1 To iTestSteps
			Datatable.GetSheet("TSteps").SetCurrentRow(j)
			If Datatable.GetSheet("TSteps").GetParameter("TestCaseID") = sTestCaseID Then
				'***Actions needs to be added here***
				LoadAndRunAction "D:\Automation\ExternalActions\AutoSales","AutoSales"
				If Datatable("Result","TSteps") = "Failed" Then
					Datatable("Result","TSteps") = "Failed"
					'Added to save screenshot of Failed step			
					strScrnshtFoldPath = "D:\Automation\TestResult\"&sTestCaseID
					Set oScrnsht = CreateObject("Scripting.FileSystemObject")
					If oScrnsht.FolderExists(strScrnshtFoldPath) = False Then
						oScrnsht.CreateFolder strScrnshtFoldPath
						Desktop.CaptureBitmap strScrnshtFoldPath&"\"&sTestCaseID&" .png", True
					Else
						Desktop.CaptureBitmap strScrnshtFoldPath&"\"&sTestCaseID&" .png", True
					End If
					Set oScrnsht = Nothing
					Exit For
				Else
					Datatable("Result","TSteps") = "Passed"
				End If
			End If
		Next
		'Close Application
		Reporter.ReportEvent micPass, "Test case Ended", "Test Case: "&sTestCaseID&" is Ended"
		strScrnshtFoldPath = "D:\Automation\TestResult\"&sTestCaseID
		Set oResFold = Createobject("Scripting.FileSystemObject")
		If oResFold.FolderExists(strScrnshtFoldPath) = False Then
			oResFold.CreateFolder strScrnshtFoldPath
			Datatable.ExportSheet strScrnshtFoldPath&"\Result.xlsx", "TSteps"
		Else
			Datatable.ExportSheet strScrnshtFoldPath&"\Result.xlsx", "TSteps"
		End If
		Set oResFold = Nothing
	End If
Next

''Publish Result
'sTimestamp = Replace(Replace(Replace(Now," ","_"),"/",""),":","")
'Datatable.ExportSheet "D:\Automation\TestResult\Result_Tcases"&"_"&sTimestamp&".xlsx","TCases"
'Datatable.ExportSheet "D:\Automation\TestResult\Result_TSteps"&"_"&sTimestamp&".xlsx","TSteps"

Set sTestCaseID = Nothing
Set sStepResult	= Nothing
Set iTestCases = Nothing
Set iTestSteps = Nothing
'Set sTimestamp = Nothing
