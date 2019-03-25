'@******************************************************************************************
'@ TestStory: Automate Sprint 8 - CBIZ-812 Construction-to-Perm Loans - Loan Purpose (LE1.X4) logic (EBS)
'@ TestCase: CBIZ-1255 TC 1 - Construction-to-Perm Loans - Loan Purpose (LE1.X4) logic
'@ TestCase: CBIZ-1256 TC 2 - Construction-to-Perm Loans - Loan Purpose (LE1.X4) logic
'@ TestCase: CBIZ-1257 TC 3 - Construction-to-Perm Loans - Loan Purpose (LE1.X4) logic
'@ TestCase: CBIZ-1259 TC 4 - Construction-to-Perm Loans - Loan Purpose (LE1.X4) logic
'@ TestCase: CBIZ-1285 TC 5 - Construction-to-Perm Loans - Loan Purpose (LE1.X4) logic
'@ TestCase: CBIZ-1286 TC 6 - Construction-to-Perm Loans - Loan Purpose (LE1.X4) logic
'@ TestCase: CBIZ-1287 TC 7 - Construction-to-Perm Loans - Loan Purpose (LE1.X4) logic
'@ TestCase: CBIZ-1288 TC 8 - Construction-to-Perm Loans - Loan Purpose (LE1.X4) logic
'@ TestCase: CBIZ-1289 TC 9 - Construction-to-Perm Loans - Loan Purpose (LE1.X4) logic
'@ TestCase: CBIZ-1335 TC 10 - Construction-to-Perm Loans - Loan Purpose (LE1.X4) logic
'@ Test Automation JIRA Task: TA-4727
'@ TestData: "Forms_BorrowerSummaryOrigination","SetBorrower","28161_BorrowerInfo"
'@ TestData: "Forms_BorrowerSummaryOrigination","SetProperty","28161_PropertyInfo"
'@ TestData: "Forms_BorrowerSummaryOrigination","SetTransactionDetails","28161_TransactionDetail"
'@ Pre-conditions: 
'@ Description:  
'@ TestSteps:
	'1 Create new loan.
	'2 In Summary - Origination Form input data.
	'3 Set Purpose of Loan - ConstructionToPerm,Initial Aquisition of Land - checked.
	'4 Set Purpose of Loan - ConstructionToPerm,Initial Aquisition of Land - UnChecked,Refinance - Checked
	'5 Set Purpose of Loan - ConstructionToPerm,Initial Aquisition of Land - UnChecked,Refinance - UnChecked
	'6 Set Purpose of Loan - Purchase
	'7 Set Purpose of Loan - Other
	'8 Set Purpose of Loan - ConstructionOnly,Initial Aquisition of Land - checked.
	'9 Enter value of Existing lien as 0 on 1003 Page 1
	'10 Enter value of Refinance_filed as 0 on 1003 Page 3
	'11 Set Purpose of Loan - Cash Out Refinance
	'12 Set Purpose of Loan - No Cash Out Refinance
	'13 Add liabilities data
	'14 Set Purpose of Loan - Cash Out Refinance
	
'@ ExpectedResult: 1. for step 3, Field LE1.X4-Purchase
                  '2. for step 4, Field LE1.X4-Refinance
                  '3. for step 5, Field LE1.X4-Construction
                  '4. for step 6, LE1.X4 = Purchase
                  '5. for step 7, LE1.X4 = Home Equity Loan
                  '6. for step 8, LE1.X4 = Purchase
                  '7. for step 11, LE1.X4 = Home Equity Loan
                  '8. for step 12, LE1.X4 = Home Equity Loan
                  '9. for step 14, LE1.X4 = Home Equity Loan
      
'********************************************************************************************
strLoanPurpose = Parameter("LoanPurpose")
strLE1X4ValueToVerify = Parameter("LE1X4Value")
strInitialLoanAquistionofLandCheckbox = Parameter("InitialLoanAquistionofLandCheckbox")
strTestCaseNo = Parameter("TestCaseNo")
strRefinanceCheckbox = Parameter("RefinanaceCheckBox")

'=================Go to Borrower Summary Origination Form==========
BIZ_Forms_Open "Borrower Summary - Origination"

Dim objMainForm
Set objMainForm = SwfWindow("swfname:=MainForm").Page("title:=.*","index:=0")

If strLoanPurpose = "Other" Then
    GUI_WebCheckBox_Set objMainForm.WebCheckbox("html id:=__cid_CheckBox14_Ctrl"), "ON"    
Else
    '============Set Loan Purpose=====
    GUI_WebCheckBox_Set objMainForm.WebCheckbox("html id:=__cid_CheckBox.*_Ctrl", "value:="&strLoanPurpose), "ON"
End If

'===========Go to RegZ - LE form============
BIZ_Forms_Open "RegZ - LE"

'==========Check/Uncheck Initial Acquisition of Land checkbox=============
GUI_WebCheckBox_Set objMainForm.Webcheckbox("html id:=__cid_CheckBox82_Ctrl"),strInitialLoanAquistionofLandCheckbox

'==========Check/Uncheck Refinance checkbox=============
GUI_WebCheckBox_Set objMainForm.Webcheckbox("html id:=__cid_CheckBox83_Ctrl"),strRefinanceCheckbox

'============Go to Loan Estimate Page 1===============
BIZ_Forms_Open "Loan Estimate Page 1"

FRM_Logger_ReportStepEvent strTestCaseNo,"value of LE1.X4 should be "&strLE1X4ValueToVerify&" when loan purpose="&strLoanPurpose&_
",IntAqOfLand="&strInitialLoanAquistionofLandCheckbox&" and RefinanceCheckBox="&strRefinanceCheckbox,NULL

'===========Verify the value of LE.X4==============
GUI_Object_ValidateValue objMainForm.WebList("html id:=dd_LE1X4"),strLE1X4ValueToVerify,"LE1.X4"


