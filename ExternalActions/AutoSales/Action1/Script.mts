'*********************************************************************************************
'Action Name	: AutoSales
'Created by		: Maruthi
'Description		: This Action belongs to Implementation of Salesforce Functions
'Inputs			: Case "FunctionName"
					'Datatable("Result","TSteps") = FunctionName("DataInput","TSteps")
'*********************************************************************************************
'Associating Repositories
RepositoriesCollection.Add "D:\Automation\ObjectRepositories\Salesforce.tsr"
Select Case Datatable.Value("Keywords","TSteps")
	Case "fn_LaunchandLogin"
		Datatable("Result","TSteps") = fn_LaunchandLogin(Datatable("DataInput","TSteps"))
	Case "fn_CreateContact"
		Datatable("Result","TSteps") = fn_CreateContact(Datatable("DataInput","TSteps"))
	Case "fn_CreateOpportunity"
		Datatable("Result","TSteps") = fn_CreateOpportunity(Datatable("DataInput","TSteps"))
	Case "fn_SearchandOpenanyRecord"
		Datatable("Result","TSteps") = fn_SearchandOpenanyRecord(Datatable("DataInput","TSteps"))
	Case "fn_ClickButtonsonDetailsPage"
		Datatable("Result","TSteps") = fn_ClickButtonsonDetailsPage(Datatable("DataInput","TSteps"))
	Case "fn_EditOpportunity"
		Datatable("Result","TSteps") = fn_EditOpportunity(Datatable("DataInput","TSteps"))
	Case "fn_Logout"
		Datatable("Result","TSteps") = fn_Logout()
End Select
