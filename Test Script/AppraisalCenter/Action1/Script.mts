'@******************************************************************************************
'@ TestStory: Appraisal Center QA Smoke Test
'@ Test Automation Document :  QASmokeTest_AppraisalCenter
'@ TestData: testaccountappraisal & testaccountappraisal@yahoo.com
'@ Pre-conditions: 
'@ Description:  
'@ TestSteps: documented in QASmokeTest_AppraisalCenter doc.
'@ ExpectedResult: 
                 

'********************************************************************************************
FRM_RT_SetupTest(null)



'====== Login to the Encompass as admin ======
BIZ_Login_UserLogin "QAAppraisalCenter" @@ hightlight id_;_-1_;_script infofile_;_ZIP::ssf9.xml_;_

'===========Added this line as Encompass is loading very slow on my machine=========
'BIZ_Nav_WaitForTabControlX "Pipeline",120

'BIZ_Pipeline_SelectPipelineViewAndLoanFolder"Super Administrator - Default View","AppTitleSmokeTest"
'=====================Select Pipeline View and Create a new blank loan====================
'BIZ_Loan_AddNewBlankLoanUnderSelectedLoanFolder "", "AppTitleSmokeTest"
BIZ_Pipeline_SelectPipelineView "AppraisalAndTitleView"

'=== If the Clear button is enabled click on it to clear out the search ===
ClearSearch()

'=== Search the loan by Loan number and open it ===
'BIZ_Loan_OpenByLoanNumber "1810QA1621252INTEG" '=== QA Environment

BIZ_Loan_OpenByLoanNumber "1903QA1621312INTEG" '=== Production Environment


'=== Selecting Services Tab & show in Alpha Order ===
BIZ_Services_ShowInOrder

'=== Selecting Appraisal services ===
BIZ_Services_Open "Appraisal"


' === Submit Appraisal Order from Encompass Application ====
BIZ_Services_OrderAppraisal "QAAppraisalCenter","testaccountappraisal@yahoo.com","510-252-2525","testaccountappraisal@yahoo.com","510-252-2525","QAAppraisalCenter"

'=== Login to Appraisal Center ===
BIZ_Services_LoginAppraisalCenter "QAAppraisalCenter"


' === Process and send document back to Encompass Application === 
BIZ_Services_AppraisalCenterProcessAndSendDocument "500000","50","Appraisal Document.pdf"

'=== Shutdown the Browser
CloseAllTabs()



BIZ_Services_ImportAppraisalDocument()


'=== Logout of Encompass Application ===
BIZ_Login_UserLogout()


Function ClearSearch()
     If SwfWindow("swfname:=MainForm").Swfbutton("swfname:=btnClearSearch").GetROProperty("enabled") = True Then
    	   GUI_SwfButton_Click SwfWindow("swfname:=MainForm").Swfbutton("swfname:=btnClearSearch")
    End If
End Function

'=== Use this function to close all browser application ===
Function CloseAllTabs()
	'=== Close All Browser session ===
	Browser("name:=My Orders").Page("title:=My Orders").Sync
	Browser("name:=My Orders").CloseAllTabs

End Function


