'*********************************************************************************************
'Function Name	: Launch and Login
'Created by		: Maruthi
'Description		: This Function helps to clear all existing browsers after that it will launch & login the Application
'Inputs			: Browsername; AppURL; Username; Password
'*********************************************************************************************
Function fn_LaunchandLogin(StrLoginParams)
	On Error Resume Next
	Dim Strallvals, Strbrowser, StrURL, Strusername, Strpassword
'	StrLoginParams = "Chrome.exe;https://login.salesforce.com/;maruthit39-t2fv@force.com;sfdcsfdc@3"
	Strallvals = Split(StrLoginParams,";")
	Strbrowser = Strallvals(0)	'"Chrome.exe"
	StrURL = Strallvals(1)	'"https://login.salesforce.com/"
	Strusername = Strallvals(2)	'"maruthit39-t2fv@force.com"
	Strpassword = Strallvals(3)	'"sfdcsfdc@3"
	'Close all Existing Applications
	Systemutil.CloseProcessByName Strbrowser
	wait 2
	'Launching the Application
	Systemutil.Run Strbrowser,StrURL
	If Browser("Br_Login_Salesforce").Page("Pg_Login_Salesforce").WebButton("WbBtn_LogIn").Exist(20) Then
		Browser("Br_Login_Salesforce").Page("Pg_Login_Salesforce").WebButton("WbBtn_LogIn").Highlight
		Reporter.ReportEvent micPass, "Salesforce Application should be Launched", "Salesforce Application got Launched Successfully..."
	Else
		Browser("Br_Login_Salesforce").Page("Pg_Login_Salesforce").CaptureBitmap "D:\Automation\TestResult\Applicationlaunch"&Replace(time,":","")&".png",False
		Reporter.ReportEvent micFail, "Salesforce Application should be Launched", "Failed to Launch Salesforce Application"
	End If
	'Login to the Application
	If Browser("Br_Login_Salesforce").Page("Pg_Login_Salesforce").WebEdit("WbEd_Username").GetROProperty("disabled") = 0 Then
		Browser("Br_Login_Salesforce").Page("Pg_Login_Salesforce").WebEdit("WbEd_Username").Set Strusername
	Else
		Browser("Br_Login_Salesforce").Page("Pg_Login_Salesforce").CaptureBitmap "D:\Automation\TestResult\username"&Replace(time,":","")&".png",False
		Reporter.ReportEvent micFail, "Username Field should be Enabled", "Username Field got Disabled"
	End If
	If Browser("Br_Login_Salesforce").Page("Pg_Login_Salesforce").WebEdit("WbEd_Password").GetROProperty("disabled") = 0 Then
		Browser("Br_Login_Salesforce").Page("Pg_Login_Salesforce").WebEdit("WbEd_Password").Set Strpassword
	Else
		Browser("Br_Login_Salesforce").Page("Pg_Login_Salesforce").CaptureBitmap "D:\Automation\TestResult\password"&Replace(time,":","")&".png",False
		Reporter.ReportEvent micFail, "Password Field should be Enabled", "Password Field got Disabled"
	End If
	If Browser("Br_Login_Salesforce").Page("Pg_Login_Salesforce").WebButton("WbBtn_LogIn").GetROProperty("disabled") = 0 Then
		Browser("Br_Login_Salesforce").Page("Pg_Login_Salesforce").WebButton("WbBtn_LogIn").Click
	Else
		Browser("Br_Login_Salesforce").Page("Pg_Login_Salesforce").CaptureBitmap "D:\Automation\TestResult\Loginbtn"&Replace(time,":","")&".png",False
		Reporter.ReportEvent micFail, "Login button should be Enabled", "Login button got Disabled"
	End If
	If Browser("Br_SalesforcePE").Page("Pg_SalesforcePE").WebButton("WbBtn_Refresh").Exist(20) Then
		Reporter.ReportEvent micPass, "Salesforce Application should be Loggedin", "Salesforce Application got Loggedin Successfully..."
	Else
		Browser("Br_Login_Salesforce").Page("Pg_Login_Salesforce").CaptureBitmap "D:\Automation\TestResult\LogintoApplication"&Replace(time,":","")&".png",False
		Reporter.ReportEvent micFail, "Salesforce Application should be Loggedin", "Failed to Login Salesforce Application"
	End If
	On Error Goto 0
End Function
'*********************************************************************************************
'Function Name	: Create Contact
'Created by		: Maruthi
'Description		: This Function helps to Create Contact with an Account
'Inputs			: Accountname; Opportunityname,
'*********************************************************************************************
Function fn_CreateContact(StrContparams)
	On Error Resume Next
	Dim Strallvals, StrAct, StrLstNm, Osf, OLkResTbl, StrContAct, StrContNm
'	StrContparams = "SFDC_SYT;SFDC_Cont"&Replace(time,":","")"
	Strallvals = Split(StrContparams,";")
	StrAct = Strallvals(0)	'"SFDC_SYT"
	StrLstNm = Strallvals(1)&Replace(time,":","")	'"SFDC_Cont"&Replace(time,":","")
	'Open Edit Contacts Page
	Set Osf = Browser("Br_SalesforcePE").Page("Pg_SalesforcePE")
	If Osf.Link("Lnk_Contacts").GetROProperty("disabled") = 0 Then
		Osf.Link("Lnk_Contacts").Click
		If Osf.WebElement("html tag:=H1","innertext:=Contacts.*.").Exist Then
			Reporter.ReportEvent micPass, "Contacts Home Page should be Opened", "Contacts Home Page Opened Successfully..."
		Else
			Reporter.ReportEvent micFail, "Contacts Home Page should be Opened", "Failed to Open Contacts Home Page"
		End If
	Else
		Reporter.ReportEvent micFail, "Contacts Link should be Enabled", "Contacts Link got disabled"
	End If
	If Osf.WebButton("html tag:=INPUT","name:=New").Exist Then
		If Osf.WebButton("html tag:=INPUT","name:=New").GetROProperty("disabled") = 0 Then
			Osf.WebButton("html tag:=INPUT","name:=New").Click	
		Else
			Reporter.ReportEvent micFail, "New Button should be enabled", "New Button got disabled"
		End If
	Else 
		Reporter.ReportEvent micFail, "New Button should Exist", "New Button doesn't Exist"
	End If
	If Osf.WebElement("html tag:=H1","innertext:=Contact Edit.*.").Exist(10) Then
		Reporter.ReportEvent micPass, "Edit Contacts Page should be Opened", "Edit Contacts Page got Opened Successfully..."
	Else
		Reporter.ReportEvent micFail, "Edit Contacts Page should be Opened", "Failed to Open Edit Contacts Page"
	End If
	If Osf.WebEdit("html tag:=INPUT","html id:=name_lastcon.*.").GetROProperty("disabled") = 0 Then
		Osf.WebEdit("html tag:=INPUT","html id:=name_lastcon.*.").Set StrLstNm
	Else
		Reporter.ReportEvent micFail, "Firstname Button should be Enabled", "Firstname Button got Disabled"
	End If
	'Clicking on Lookup Icon
	If Osf.Link("html id:=con4_lkwgt").Exist Then
		If Osf.Link("html id:=con4_lkwgt").GetROProperty("disabled") = 0 Then
			Osf.Link("html id:=con4_lkwgt").Click	
		Else
			Reporter.ReportEvent micFail, "Lookup Icon should be enabled", "Lookup Icon got disabled"
		End If	
	Else
		Reporter.ReportEvent micFail, "Lookup Icon should exist", "Lookup Icon doesn't exist"
	End If
	If Browser("name:=.*.","application version:=Chrome.*.","creationtime:=1").Page("title:=.*.").Frame("html id:=searchFrame").WebEdit("html id:=lksrch").Exist(25) Then
		Browser("name:=.*.","application version:=Chrome.*.","creationtime:=1").Page("title:=.*.").Frame("html id:=searchFrame").WebEdit("html id:=lksrch").Set StrAct
		Browser("name:=.*.","application version:=Chrome.*.","creationtime:=1").Page("title:=.*.").Frame("html id:=searchFrame").WebButton("html tag:=INPUT","name:= Go.*.").Click
	Else
		Reporter.ReportEvent micFail, "Lookup Search Edit field should exist", "Lookup Search Edit field doesn't exist"
	End If
	Set OLkResTbl = Browser("name:=.*.","application version:=Chrome.*.","creationtime:=1").Page("title:=.*.").Frame("html id:=resultsFrame").WebTable("html tag:=TABLE","name:=Account Name")
	intrw = OLkResTbl.GetRowWithCellText(StrAct)
	OLkResTbl.ChildItem(intrw,0,"Link",0).Click
	If Osf.WebEdit("xpath:=(//*[@class='dataCol col02'])[6]//span//input").Exist Then
		StrContAct = Osf.WebEdit("xpath:=(//*[@class='dataCol col02'])[6]//span//input").GetROProperty("value")
		If Trim(Ucase(StrAct)) = Trim(Ucase(StrContAct)) Then
			Reporter.ReportEvent micPass, StrAct&" Account should be selected", StrAct&" Account got selected Successfully..."
		Else
			Reporter.ReportEvent micFail, StrAct&" Account should be selected", "Failed to select "&StrAct&" Account"
		End If
	Else
		Reporter.ReportEvent micFail, " Account Field should Exist", " Account Field doesn't Exist"
	End If
	Osf.WebButton("html tag:=INPUT","name:= Save.*.","index:=0").Click
	If Osf.WebElement("class:=pageDescription","html tag:=H2").Exist Then
		StrContNm = Osf.WebElement("class:=pageDescription","html tag:=H2").GetROProperty("innertext")
		If Trim(ucase(StrContNm)) = Trim(ucase(StrLstNm)) Then
			Reporter.ReportEvent micPass, "Contact should be Created", StrContNm&" Contact Created Successfully..."
		Else
			Reporter.ReportEvent micFail, "Contact should be Created", "Failed to create Contact."
		End If
	Else
		Reporter.ReportEvent micFail, "Contact Details page should Exist", "Contact Details page doesn't Exist"
	End If
	On Error Goto 0
End Function
'*********************************************************************************************
'Function Name	: Create Opportunity
'Created by		: Maruthi
'Description		: This Function helps to Create Opportunity 
'Inputs			: Accountname; Opportunityname; Stagename
'*********************************************************************************************
Function fn_CreateOpportunity(StrOpptyparams)
	On Error Resume Next
	Dim Osf, Strallvals, StrAct, StrOppNm, Strstage, Strstgval, StrOppty
	Set Osf = Browser("Br_SalesforcePE").Page("Pg_SalesforcePE")
'	StrOpptyparams = "SFDC_SYT;SFDC_Opp;Qualification"
	Strallvals = Split(StrOpptyparams,";")
	StrAct = Strallvals(0)	'"SFDC_SYT"
	StrOppNm = Strallvals(1)&Replace(time,":","")	'"SFDC_Opp"&Replace(time,":","")
	Strstage = Strallvals(2)	'"Qualification"
	If Osf.Link("Lnk_Opportunities").GetROProperty("disabled") = 0 Then
		Osf.Link("Lnk_Opportunities").Click
		If Osf.WebElement("html tag:=H1","innertext:=Opportunities.*.").Exist Then
			Reporter.ReportEvent micPass, "Opportunities Home Page should be Opened", "Opportunities Home Page Opened Successfully..."
		Else
			Reporter.ReportEvent micFail, "Opportunities Home Page should be Opened", "Failed to Open Opportunities Home Page"
		End If
	Else
		Reporter.ReportEvent micFail, "Opportunities Link should be Enabled", "Opportunities Link got disabled"
	End If
	If Osf.WebButton("html tag:=INPUT","name:=New").Exist Then
		If Osf.WebButton("html tag:=INPUT","name:=New").GetROProperty("disabled") = 0 Then
			Osf.WebButton("html tag:=INPUT","name:=New").Click	
		Else
			Reporter.ReportEvent micFail, "New Button should be enabled", "New Button got disabled"
		End If
	Else 
		Reporter.ReportEvent micFail, "New Button should Exist", "New Button doesn't Exist"
	End If
	If Osf.WebElement("html tag:=H1","innertext:=Opportunity Edit.*.").Exist(10) Then
		Reporter.ReportEvent micPass, "Edit Opportunity Page should be Opened", "Edit Opportunity Page got Opened Successfully..."
	Else
		Reporter.ReportEvent micFail, "Edit Opportunity Page should be Opened", "Failed to Open Edit Opportunity Page"
	End If
	If Osf.WebEdit("xpath:=(//*[@class='requiredInput'])[1]//input").Exist(25) Then
		Osf.WebEdit("xpath:=(//*[@class='requiredInput'])[1]//input").Set StrOppNm
	Else
		Reporter.ReportEvent micFail, "Opportunity Name field should exist", "Opportunity Name doesn't exist"
	End If
	'Clicking on Lookup Icon
	If Osf.Link("html id:=opp4_lkwgt").Exist Then
		If Osf.Link("html id:=opp4_lkwgt").GetROProperty("disabled") = 0 Then
			Osf.Link("html id:=opp4_lkwgt").Click	
		Else
			Reporter.ReportEvent micFail, "Lookup Icon should be enabled", "Lookup Icon got disabled"
		End If	
	Else
		Reporter.ReportEvent micFail, "Lookup Icon should exist", "Lookup Icon doesn't exist"
	End If
	If Browser("name:=.*.","application version:=Chrome.*.","creationtime:=1").Page("title:=.*.").Frame("html id:=searchFrame").WebEdit("html id:=lksrch").Exist(25) Then
		Browser("name:=.*.","application version:=Chrome.*.","creationtime:=1").Page("title:=.*.").Frame("html id:=searchFrame").WebEdit("html id:=lksrch").Set StrAct
		Browser("name:=.*.","application version:=Chrome.*.","creationtime:=1").Page("title:=.*.").Frame("html id:=searchFrame").WebButton("html tag:=INPUT","name:= Go.*.").Click
	Else
		Reporter.ReportEvent micFail, "Lookup Search Edit field should exist", "Lookup Search Edit field doesn't exist"
	End If
	'Select Lookup Account:
	Set OLkResTbl = Browser("name:=.*.","application version:=Chrome.*.","creationtime:=1").Page("title:=.*.").Frame("html id:=resultsFrame").WebTable("html tag:=TABLE","name:=Account Name")
	intrw = OLkResTbl.GetRowWithCellText(StrAct)
	OLkResTbl.ChildItem(intrw,0,"Link",0).Click
	If Osf.WebEdit("xpath:=(//*[@class='requiredInput'])[2]//span//input").Exist Then
		StrOptyAct = Osf.WebEdit("xpath:=(//*[@class='requiredInput'])[2]//span//input").GetROProperty("value")
		If Trim(Ucase(StrAct)) = Trim(Ucase(StrOptyAct)) Then
			Reporter.ReportEvent micPass, StrAct&" Account should be selected", StrAct&" Account got selected Successfully..."
		Else
			Reporter.ReportEvent micFail, StrAct&" Account should be selected", "Failed to select "&StrAct&" Account"
		End If
	Else
		Reporter.ReportEvent micFail, " Account Field should Exist", " Account Field doesn't Exist"
	End If
	'Select Stage:
	Osf.WebList("xpath:=(//*[@class='requiredInput'])[4]//span//select").Select Strstage
	Strstgval = Osf.WebList("xpath:=(//*[@class='requiredInput'])[4]//span//select").GetROProperty("value")
	If Trim(Ucase(Strstage)) = Trim(Ucase(Strstgval)) Then
		Reporter.ReportEvent micPass, Strstage&" Stage value should selected", Strstage&" Stage value got selected successfully..."
	Else
		Reporter.ReportEvent micFail, Strstage&" Stage value should selected", "Unable to select stage value "&Strstage
	End If
	'Click on Save button:
	Osf.WebButton("html tag:=INPUT","name:= Save.*.","index:=0").Click
	If Osf.WebElement("xpath:=//*[@class='errorMsg']").Exist Then
		'Select Close Date:
		If Osf.Link("xpath:=//*[@class='dateInput dateOnlyInput']//span//a").Exist Then
			Osf.Link("xpath:=//*[@class='dateInput dateOnlyInput']//span//a").Click
		Else
			Reporter.ReportEvent micFail, "Closedate link should exist", "Closedate link doesn't exist"
		End If
	End If
	'Click on Save button:
	Osf.WebButton("html tag:=INPUT","name:= Save.*.","index:=0").Click
	If Osf.WebElement("class:=pageDescription","html tag:=H2").Exist Then
		StrOppty = Osf.WebElement("class:=pageDescription","html tag:=H2").GetROProperty("innertext")
		If Trim(ucase(StrOppty)) = Trim(ucase(StrOppNm)) Then
			Reporter.ReportEvent micPass, "Opportunity should be Created", StrOppty&" Opportunity Created Successfully..."
		Else
			Reporter.ReportEvent micFail, "Opportunity should be Created", "Failed to create Opportunity."
		End If
	Else
		Reporter.ReportEvent micFail, "Opportunity Details page should Exist", "Opportunity Details page doesn't Exist"
	End If
	On Error Goto 0
End Function
'*********************************************************************************************
'Function Name	: Search and Open any Record
'Created by		: Maruthi
'Description		: This Function helps to Search and open any Record i.e; Account,Contact,Opportunity,Order etc....
'Inputs			: Recordname; Tablename
'*********************************************************************************************
Function fn_SearchandOpenanyRecord(Strrecparams)
	On Error Resume Next
	Dim Strallvals, Strsrchval, Strtblnm, Osf, OUnvtbl, StrWbEle
'	Strrecparams = "Sfdc_opp;Opportunities"
	Strallvals = Split(Strrecparams,";")
	Strsrchval = Strallvals(0)	'"Sfdc_opp"
	Strtblnm = Strallvals(1)	'"Opportunities"
	Set Osf = Browser("Br_SalesforcePE").Page("Pg_SalesforcePE")
	Osf.WebEdit("html id:=phSearchInput").Set Strsrchval
	If Osf.WebButton("html id:=phSearchButton").Exist Then
		Osf.WebButton("html id:=phSearchButton").Click
		Reporter.ReportEvent micPass, "Search Button should exist", "Search Button exist"
	Else
		Reporter.ReportEvent micFail, "Search Button should exist", "Search Button doesn't exist"
	End If
	Select Case Strtblnm
		Case "Accounts"
			Set OUnvtbl = Osf.WebTable("html tag:=TABLE","name:=Account Name","class:=list")
		Case "Contacts"
			Set OUnvtbl = Osf.WebTable("html tag:=TABLE","name:=Name","class:=list")
		Case "Opportunities"
			Set OUnvtbl = Osf.WebTable("html tag:=TABLE","name:=Opportunity Name","class:=list")
	End Select
	intrw = OUnvtbl.GetRowWithCellText(Strsrchval)
	OUnvtbl.ChildItem(intrw,2,"Link",0).Click
	If Osf.WebElement("class:=pageDescription","html tag:=H2").Exist Then
		StrWbEle = Osf.WebElement("class:=pageDescription","html tag:=H2").GetROProperty("innertext")
		If Trim(ucase(StrWbEle)) = Trim(ucase(Strsrchval)) Then
			Reporter.ReportEvent micPass, Strtblnm&" Details page should be opened", Strtblnm&" Details page got opened successfully..."
		Else
			Reporter.ReportEvent micFail, Strtblnm&" Details page should be opened", "Failed to open "&Strtblnm&" Details page." 
		End If
	Else
		Reporter.ReportEvent micFail, Strtblnm&" Details page should Exist", Strtblnm&" Details page doesn't Exist"
	End If
	On Error Goto 0
End Function
'*********************************************************************************************
'Function Name	: Click any Button on DetailsPage
'Created by		: Maruthi
'Description		: This Function helps to Click any button for all record types i.e; Edit,Clone,Delete etc...
'Inputs			: Buttonname,
'*********************************************************************************************
Function fn_ClickButtonsonDetailsPage(Strbtnparams)
	On Error Resume Next
	Dim Osf, StrBtn, Strallvals
'	Strallvals = Split(Strbtnparams,";")
	StrBtn = Strbtnparams
	Set Osf = Browser("Br_SalesforcePE").Page("Pg_SalesforcePE")
	Select Case StrBtn
		Case "Edit"
			If Osf.WebButton("html tag:=INPUT","name:= Edit.*.","index:=0").Exist Then
				Osf.WebButton("html tag:=INPUT","name:= Edit.*.","index:=0").Click
			Else
				Osf.CaptureBitmap "D:\Automation\TestResult\Edit Button"&Replace(time,":","")&".png",False
				Reporter.ReportEvent micFail, "Edit Button should exist", "Edit Button doesn't exist"
			End If
		Case "Clone"
			If Osf.WebButton("html tag:=INPUT","name:=  Clone.*.","index:=0").Exist Then
				Osf.WebButton("html tag:=INPUT","name:=  Clone.*.","index:=0").Click
			Else
				Osf.CaptureBitmap "D:\Automation\TestResult\ Clone Button"&Replace(time,":","")&".png",False
				Reporter.ReportEvent micFail, " Clone Button should exist", "Edit Button doesn't exist"
			End If
	End Select
	On Error Goto 0
End Function
'*********************************************************************************************
'Function Name	: Edit Opportunity
'Created by		: Maruthi
'Description		: This Function helps to Update any changes in Opportunity level
'Inputs			: Opptyname,
'*********************************************************************************************
Function fn_EditOpportunity(Stroptyparams)
	On Error Resume Next
	Dim Osf, Strallvals
'	Strallvals = Split(Stroptyparams,";")
'	StrOppty = ""
	StrOppty = Stroptyparams
	Set Osf = Browser("Br_SalesforcePE").Page("Pg_SalesforcePE")
	If Osf.WebEdit("xpath:=(//*[@class='requiredInput'])[1]//input").Exist(25) Then
		Osf.WebEdit("xpath:=(//*[@class='requiredInput'])[1]//input").Set StrOppty
	Else
		Osf.CaptureBitmap "D:\Automation\TestResult\Opportunityfield"&Replace(time,":","")&".png",False
		Reporter.ReportEvent micFail, "Opportunity Name field should exist", "Opportunity Name doesn't exist"
	End If
	'Click on Save button:
	Osf.WebButton("html tag:=INPUT","name:= Save.*.","index:=0").Click
	If Osf.WebElement("class:=pageDescription","html tag:=H2").Exist Then
		StrOppty = Osf.WebElement("class:=pageDescription","html tag:=H2").GetROProperty("innertext")
		If Trim(ucase(StrOppty)) = Trim(ucase(StrOppNm)) Then
			Reporter.ReportEvent micPass, "Opportunity should be Created", StrOppty&" Opportunity Created Successfully..."
		Else
			Osf.CaptureBitmap "D:\Automation\TestResult\Opportunityfield"&Replace(time,":","")&".png",False
			Reporter.ReportEvent micFail, "Opportunity should be Created", "Failed to create Opportunity."
		End If
	Else
		Osf.CaptureBitmap "D:\Automation\TestResult\Opportunityfield"&Replace(time,":","")&".png",False
		Reporter.ReportEvent micFail, "Opportunity Details page should Exist", "Opportunity Details page doesn't Exist"
	End If
	On Error Goto 0
End Function
'*********************************************************************************************
'Function Name	: LogOut
'Created by		: Maruthi
'Description		: This Function helps to Logout from the Salesforce Application
'Inputs			: NA
'*********************************************************************************************
Function fn_Logout()
	On Error Resume Next
	If Browser("Br_SalesforcePE").Page("Pg_SalesforcePE").WebElement("WbEl_userNavarrow").GetROProperty("disabled") = 0 Then
		Browser("Br_SalesforcePE").Page("Pg_SalesforcePE").WebElement("WbEl_userNavarrow").Click
			wait 2
			If Browser("Br_SalesforcePE").Page("Pg_SalesforcePE").Link("Lnk_Logout").Exist(15) Then
				Browser("Br_SalesforcePE").Page("Pg_SalesforcePE").Link("Lnk_Logout").Click
			Else
				Reporter.ReportEvent micFail, "Logout Link should Exist", "Logout Link Doesn't Exist"
			End If
	Else
		Reporter.ReportEvent micFail, "UserNavarrow should be Enabled", "UserNavarrow got Disabled"
	End If
	If Browser("Br_Login_Salesforce").Page("Pg_Login_Salesforce").WebButton("WbBtn_LogIn").Exist(20) Then
		Browser("Br_Login_Salesforce").Page("Pg_Login_Salesforce").WebButton("WbBtn_LogIn").Highlight
		Reporter.ReportEvent micPass, "Salesforce Application should be Loggedout", "Salesforce Application got Logged out Successfully..."
	Else
		Reporter.ReportEvent micFail, "Salesforce Application should be Loggedout", "Failed to Loggedout Salesforce Application"
	End If	
	On Error Goto 0
End Function
