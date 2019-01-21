Attribute VB_Name = "MMAIN1"
Option Explicit

'George added for County of Essex on May 20,2005
Global glbWeek
Global glbFrom, glbTo
Global glbPayP
Global glbflgFU As Boolean
' dkostka - 03/19/2002 - Hardcoded SMTP servers for all known plants, use a global for this.
Global glbSMTPServerIP As String

Const ODBC_ADD_DSN = 1        ' Add a new data source.
Const ODBC_CONFIG_DSN = 2     ' Configure (edit) existing data source.
Const ODBC_REMOVE_DSN = 3     ' Remove existing data source.
Public Const ODBC_DSN = "INFOHR" '
Public Const ODBC_CONNECT_SQL = "ODBC;DSN=" & ODBC_DSN
'Public Const BasicCountry = "CANADA&U.S.A.&BAHAMAS&MEXICO&GERMANY&SINGAPORE&ENGLAND"
Global glbMsgBoxResult As VbMsgBoxResult
Global glbTLAY
Global RptODBC_SQL
Global glbDateFmt(3), glbDateSeparator
Global glbFrench As Boolean     ' dkostka - 03/20/2001 - Added support for French Windows environment
Global glbTrsDept As String
Global glbTransDiv As String
Global glbDIVCount, glbSDIV
Global glbDIVList
Global glbtermopen
Global glbPayrollID As String
Global gSec_Upd_EmpFlags As Boolean
Global glbCourseCodeSele As Boolean
Global glbFromDate
Global glbIsUseIHRDS As Boolean 'Ticket #20310 Franks 05/10/2011

Global glbSQL As Boolean
'Sam add for Oracle
Global glbOracle As Boolean
' danielk - 01/02/03 - PayWeb interface
Global glbPayWeb As Boolean
Global glbAdv As Boolean
' danielk - 01/14/2003 - Payweb interface work
Global glbGP As Boolean 'geo added
Global glbMediPay As Boolean 'geo added
Global glbCwis As Boolean ' Simona added for ticket #14890 -LEEDS Grenville CAS
Global glbPayWebEXE As String
Global glbVadimEXE As String
Global glbIntegrationEXE As String
Global glbChgTermDate
Global glbSpouseSIN
Global glbChgPT
Global glbChgUseProfile
Global glbChgTermReason As String
Global glbChgNewEmpnbr
Global glbChgBenTermDate
Global glbDisabled As Boolean
Global glbDBPassFlag As Boolean
Global glbTempFlag As Boolean
Global glbNewRept
Global glbWorkVisaNo
Global glbWorkExpDate

' Frank - Dec 22 2003 for Surrey Place
Global glbSPCTermDate
Global glbSPCTermReason As String
Global glbSPCNewEmpNo
Global glbSPCPPay As String

Global glbFormCaption

Global frmFind
Global glbRest

Global glbNo                            '
Global glbUEnt                          '
Global glbSkip                          '
Global gSec_Upd_Requisition As Boolean  '
Global glbUS As Boolean                 'temp
Global glbReq$, glbReqPos$              '
Global glbCReq$, glbCReqPos$            '
Global gdbAdoIhrappt As New ADODB.Connection

Global glbLinamar As Boolean
'Global glbLinamar1 As Boolean
Global glbLinNewPosSal As Boolean
Global glbAxxent As Boolean
Global glbNiagaraFulls As Boolean
Global glbCElgin As Boolean
Global glbGuelph As Boolean
Global glbSyndesis As Boolean
Global glbCBrant As Boolean
Global glbWHSCC As Boolean
Global glbOttawaCCAC As Boolean
Global glbSoroc As Boolean
Global glbDundasACL As Boolean
Global glbBrantCount As Boolean
Global glbLambton As Boolean
Global glbBurlTech As Boolean

Global glbSetPos, glbSetSal, glbSetPer
Global gflHelp$             ' help file name
Global glbUserID As String  ' employee signed on to system
Global glbSSOPwd As String    '7.9 SSO Password
Global glbSSO As String          '7.9 SSO
Global glbUserNAME As String
Global glbLUserID As String
Global glbLUserNAME As String
Global glbEmpNbr As Long
Global glbWorkDir As String    'working directory
Global glbsDateFormat As String
Global lCurrentKey As Long
Global SQLDatabaseName As String
Global SQLServerName As String
Global SQLUserName As String
Global SQLUserPassword As String
Global SQLDriver As String
Global glbHostFile As String
Global glbHosted As Boolean
Global glbLicenseKey As Date
Global glbUserEmpNo As Long
Global glbHRSSSecure As Boolean

Global glbNoNONE As Boolean   ' if "-NON" union belong to the Login's user
Global glbNoEXEC As Boolean   ' if "-EXE" union belong to the Login's user
Global glbUNION As String
Global glbUnionForm As Boolean
Global glbUNIONTe As String
Global glbUnionFormTe As Boolean
Global glbFTPT As String
Global glbBand As String 'WFC Only
Global glbEmalType As String
Global glbCRWPrintSetup As Boolean

Global glbLabels(2) As New Collection


Global gdbAdoIhr001 As New ADODB.Connection
Global gdbAdoIhr001K As New ADODB.Connection
Global gdbAdoIhr001X As New ADODB.Connection
Global gdbAdoIhr001W As New ADODB.Connection
Global gdbAdoIhr001O As New ADODB.Connection
Global gdbAdoIhr001B As New ADODB.Connection
Global gdbAdoIhrWFC As New ADODB.Connection
Global gdbAdoIhrWFCA As New ADODB.Connection
Global gdbAdoSN2322 As New ADODB.Connection
Global gdbAdoIHREDU As New ADODB.Connection
Global gdbPayroll As New ADODB.Connection
Global gdbAdoIhr001_DOC As New ADODB.Connection

Global NewHireForms As New Collection

Global glbEEFIND_Refresh As Integer     'refresh for find?
Global glbINIFileName$      ' file name of ihr ini file (includes directory)
Global glbOHSEdit%          ' in edit mode of OHS records
Global glbStopPerform%      ' are salary, position records for person
Global glbStopSalary%       ' are there position records pre salary
Global glbDATA1Recs%        ' global are there records in data1?
                            ' used for module to set controls
Global glbOClass$           ' occupational classs
Global glbOClassDesc$       ' o class description (short)
Global glbOClassMode%       ' reviewing/selecting oclasses?
Global glbBasicChg%         ' refresh ee list?
Global glbSecUSERID     As String       'hold for security lookup deptartments
Global glbSecEEName$
Global glbDeptInhSel%       ' inhibit select
Global glbOHRSDeptInhSel
Global glbDivInhSel%
Global glbJobMasterInhSel%, glbJobMaster, glbJobMasterDesc, glbJobSection As String, glbJobSectionDesc As String
Global glbSalDistInhSel%
Global glbPayCategoryInhSel%
Global glbHOMEInhSel%
Global glbLgrInhSel%        ' inhibit select
Global glbDiv, glbDivDesc, glbLgr, glbLgrDesc
Global glbJobFam, glbJobFamDesc 'Ticket #26233 Franks 11/21/2014
Global glbSubJobFam, glbSubJobFamDesc 'Ticket #26233 Franks 11/21/2014
Global glbGroupJob, glbGroupJobDesc 'Ticket #26233 Franks 11/21/2014
Global DemoSystem, DemoMaxEmp%   ' Demo System True/False

Global glbCrsCode As String, glbCrsCodeDesc As String

Global glbBenAdded As String
Global glbBenChanged As String
Global glbBenDeleted As String
Global glbBenEffDate

Global glbiOneWhere As Integer  ' set in reports
Global glbstrSelCri As String  'selection criteria

Global glbCodeRef ' if lookup (? or double click on code and then hit new
                    ' checked on lost focus and if true bypasses code desc as passed

Global glbAddHisWarning%, glbFOLLOWUPS%, glbFOLLOWUPDAYS%
Global glbFollowUpsRemain%, glbFOLLOWUPSCOMP%, glbFollwUpsFound%
Global glbExclusiveDB%
Global gstrAccPWord$, gstrAccUID$
Global glbMultiUserNum%
Global Database_Type As String

Global gintRollBack%
Global glbPos$, glbPosDesc$, glbPosSkill

Global glbIHRAUDIT As String
Global glbDBDir As String
Global glbIHRDB As String
Global glbIHRWFC As String              'Jaddy 8/9/99
Global glbIHRDBA As String
Global glbIHRDBW  As String
Global glbIHRDBO As String
Global glbIHRDBB As String
Global glbIHRREPORTS As String
Global glbIHRWFCA As String
Global glbSN2322 As String
Global glbIHREDU As String
Global glbPayrollDB As String


Global glbAdoIHRAUDIT As String
Global glbAdoIHRDB As String
Global glbAdoIHRWFC As String
Global glbAdoIHRDBA As String
Global glbAdoIHRDBW  As String
Global glbAdoIHRDBO As String
Global glbAdoIHRDBB As String
Global glbAdoIHRWFCA As String
Global glbAdoSN2322 As String

Global glbadoIHREDU As String
Global glbAdoIHRDB_DOC As String

Global glbCompName As String ' company name
Global glbCompLvl As Double
Global glbCompSerial As String
Global glbCompEdFrom As Variant, glbCompEdTo As Variant
Global glbCompEdFromS As Variant, glbCompEdToS As Variant
Global glbCompEntSick$, glbCompEntVac$
Global glbCompWDate$
Global glbEntOutStanding$, glbEntOutStandingS$
Global glbCompDecHR, glbMulti, glbMultiGrid
Global glbCompEntVacDaily

Global glbBYPASS380 As Integer

Global glbCompNo As String
Global glbCode As String, glbCodeDesc As String
Global glbHome As String, glbHomeDesc As String
Global glbDept As String, glbDeptDesc As String
Global glbGLNum As String, glbGLNumD As String
Global glbProv As String, glbProvDesc  As String
Global glbJob$, glbJobDesc$, glbTermDate$, glbRehireDt$
Global glbPlan$, glbPlanDesc$, glbSurvDate$, glbDueDate$   'laura 03/12/98
Global glbLEE_ID As Long, glbLEE_SName As String, glbLEE_ProdLine As String
Global glbTERM_ID As Long, glbTerm_SName As String
Global glbTran_ID As Long, glbTran_SName, glbTran_Fname
Global glbTran_Seq As Long
Global glbSIN As String 'Ticket #18566
Global glbTermCancel As Boolean
Global glbCandidate As Long
Global glbCand_SF_ID As Long
Global glbCommentType, glbCounselType
Global glbCommentDate, glbCounselDate

Global glbEEFIND_New As Long ' is new allowed on find ee?
Global glbLEE_FName As String    'last ee looked up
Global glbTerm_FName As String
Global glbTERM_Seq As Double  'Termination Sequence Number
Global glbSort, glbOnTop, glbEEOK, glbTermOK, PrevOnTop
Global glbCountry       'Jaddy 6/7/99
Global glbEmpCountry
Global glbWSIB      'Jaddy 6/7/99
Global glbTrsEE_ID As Variant
Global glbTrsDIV
Global glbTrsStatus 'Ticket #23247 Franks 04/19/2013
Global glbTrsUnion 'Ticket #23247 Franks 04/19/2013
Global glbTrsHourWeek 'Ticket #23247 Franks 04/19/2013
Global glbOMERS_Date As Boolean
Global glbTrsVadim1
Global glbUnionCode As Boolean
Global glbUnionDemog As Boolean
Global glbLabLang As String

Global glbDolType As String
Global glbDolFDate
Global glbDolTDate

Global glbTabNam As String  ' table name for codes descriptions
Global glbFrmCaption$   'FormCaption
Global glbErrNum&       'error number

Global glbGridReason
Global glbGridEDate
Global glbGridNDate

Global glbF7FirmAcct As String
Global glbF7FirmAcctNo As String
Global glbF7CaseNo As Long


Global glbFollowUpList As String

'=======================================================
Global glbAbout As Integer
Global glbNDepts As Integer ' number of departments user has
                            ' -1 implies all
Global glbDepts(50) As String 'changed from 10 to 20 by RAUBREY 4/8/97
                              'holds an array of departments the user can see

Global glbCrsCodeStrArr(21) As String 'Hold all field values on frmMCourseCode

Global glbENTRecalc   'Entitlements Need ReCalculate...
Global glbENTScreen   'Entitlement Overview Needs Refresh
'========================================================
' values held for frmSECURITY
'=========================
Global glbFNo
Global glbChkPass
Global glbConfPass As String
Global glbEmployeeNo&
Global glbPassword$
Global glbTxtPassword As String

Global glbAccessPswd As Boolean
Global glbAnnMonth As Integer
Global glbSenMonth As Integer

Global glbIsGWL As Boolean 'Ticket #21518 Franks 04/26/2012

'Security for individual entering the application
'=================================================
' boolean frmSECURITY flags
Global gSec_Emp_Based As Boolean
Global gSec_Inq_Basic As Boolean
Global gSec_Upd_Basic As Boolean
Global gSec_Upd_Banking As Boolean
Global gSec_Inq_Banking As Boolean
Global gSec_Upd_EmploymentEQT As Boolean
Global gSec_Inq_EmploymentEQT As Boolean
Global gSec_Upd_PayEQT As Boolean
Global gSec_Inq_PayEQT As Boolean
Global gSec_Upd_Dependents As Boolean
Global gSec_Del_Dependents As Boolean 'Ticket #22009 Franks 05/10/2012
Global gSec_Inq_Dependents As Boolean
Global gSec_Upd_Skills As Boolean
Global gSec_Inq_Skills As Boolean
Global gSec_Upd_Formal_Education As Boolean
Global gSec_Inq_Formal_Education As Boolean
Global gSec_Upd_Education_Seminars As Boolean
Global gSec_Inq_Education_Seminars As Boolean
Global gSec_Upd_Salary As Boolean
Global gSec_Inq_Salary As Boolean
Global gSec_Inq_Performance As Boolean
Global gSec_Upd_Performance As Boolean
Global gSec_Inq_Position As Boolean
Global gSec_Upd_Position As Boolean
Global gSec_Upd_Benefits As Boolean
Global gSec_Inq_Benefits As Boolean
Global gSec_Upd_Beneficiary As Boolean
Global gSec_Inq_Beneficiary As Boolean
Global gSec_Upd_Entitlements As Boolean
Global gSec_Inq_Entitlements As Boolean
Global gSec_Upd_Associations As Boolean
Global gSec_Inq_Associations As Boolean
Global gSec_Upd_UserDefineTbl As Boolean
Global gSec_Inq_UserDefineTbl As Boolean
Global gSec_Upd_PayrollTrans As Boolean
Global gSec_Inq_PayrollTrans As Boolean
Global gSec_Upd_Follow_Ups As Boolean
Global gSec_Inq_Follow_Ups As Boolean
Global gSec_Upd_Health_Safety As Boolean
Global gSec_Inq_Health_Safety As Boolean
Global gSec_Upd_Attendance As Boolean
Global gSec_Inq_Attendance As Boolean
Global gSec_Upd_Attendance_History As Boolean
Global gSec_Inq_Attendance_History As Boolean
Global gSec_Upd_Hrly_Entitlements As Boolean
Global gSec_Inq_Hrly_Entitlements As Boolean

'Mostafa Attendance Group Code Matrix
Global gSec_Upd_Attendance_Group_Code_Matrix As Boolean
Global gSec_Inq_Attendance_Group_Code_Matrix As Boolean

Global gSec_Upd_Job_Files_Attachment As Boolean
Global gSec_Inq_Job_Files_Attachment As Boolean

Global gSec_Upd_Temp_Cross_Training As Boolean
Global gSec_Inq_Temp_Cross_Training As Boolean

Global gSec_Upd_Training_List As Boolean
Global gSec_Inq_Training_List As Boolean

Global gSec_Upd_Work_Schedule As Boolean
Global gSec_Inq_Work_Schedule As Boolean

Global glbCaseFiles 'Case file number Leeds & grenville - Mostafa
Global glbCaseNum
Global glbCaseAssociate

Global gSec_Show_SIN_SSN As Boolean
Global gSec_Show_DOB As Boolean
Global gSec_Show_ADDRESS As Boolean
Global gSec_Show_Marital As Boolean
Global gSec_Add_Attendance As Boolean
Global gSec_Add_NewHire As Boolean
Global gSec_Add_Comments As Boolean
Global gSec_SP_ViewOwn As Boolean
Global gSec_Comments_ViewOwn As Boolean
Global gSec_Counsel_ViewOwn As Boolean
Global gSec_FollUp_ViewOwn As Boolean
Global gSec_OthInfo_ViewOwn As Boolean
Global gSec_EmpFlags_ViewOwn As Boolean
Global gSec_EmpHis_ViewOwn As Boolean
Global gSec_GLDist_ViewOwn As Boolean
Global gSec_Performance_ViewOwn As Boolean

'Ticket #18406 - Farmers' Mutual Insurance
Global gSec_Lock_Password As Boolean

Global gSec_Upd_Counselling As Boolean
Global gSec_Inq_Counselling As Boolean
Global gSec_Upd_Comments As Boolean
Global gSec_Inq_Comments As Boolean
Global gSec_Upd_OtherInformation As Boolean
Global gSec_Inq_OtherInformation As Boolean

Global gSec_Upd_Other_Entitlements As Boolean
Global gSec_Inq_Other_Entitlements As Boolean
Global gSec_Upd_Earnings As Boolean
Global gSec_Inq_Earnings As Boolean
Global gSec_Upd_Job_Classes As Boolean
Global gSec_Inq_Job_Classes As Boolean
Global gSec_Upd_Job_Master As Boolean
Global gSec_Inq_Job_Master As Boolean

Global gSec_Upd_Profit_Sharing As Boolean
Global gSec_Inq_Profit_Sharing As Boolean
Global gSec_Rpt_Profit_Sharing As Boolean
Global gSec_Rpt_Red_Circled As Boolean

Global gSec_Upd_SAMTableMasterLinks As Boolean
Global gSec_Inq_SAMTableMasterLinks As Boolean

Global gSec_Upd_PayPeriod_Master As Boolean
Global gSec_Inq_PayPeriod_Master As Boolean

Global glbPlantCode
Global glbPlantDesc
Global glbWFCUserSecList As String
Global glbDeptAllRight As Boolean
Global glbWFCFullRights As Boolean
Global glbEESection
Global gSec_Upd_Job_Eval As Boolean
Global gSec_Inq_Job_Eval As Boolean
Global gSec_Upd_Job_Skills As Boolean
Global gSec_Inq_Job_Skills As Boolean
Global gSec_Inq_Termination_Report As Boolean
Global gSec_Upd_Terminations As Boolean
Global gSec_Inq_Terminations As Boolean
Global gSec_Upd_RetirementProc As Boolean 'Ticket #18566
Global gSec_Inq_RetirementProc As Boolean 'Ticket #18566
Global gSec_Upd_DeathProc As Boolean 'Ticket #18566
Global gSec_Inq_DeathProc As Boolean 'Ticket #18566
Global gSec_Upd_Company As Boolean
Global gSec_Inq_Company As Boolean
Global gSec_Upd_Master_Table As New Collection
Global gSec_Inq_Master_Table As New Collection
Global gSec_Upd_Departments As Boolean
Global gSec_Inq_Departments As Boolean
Global gSec_Upd_Divisions As Boolean
Global gSec_Inq_Divisions As Boolean
Global gSec_Upd_SalDist As Boolean
Global gSec_Inq_SalDist As Boolean
Global gSec_Upd_AffirmAction_Data As Boolean 'Ticket #18790
Global gSec_Inq_AffirmAction_Data As Boolean 'Ticket #18790
Global gSec_Upd_AffirmAction_Purge As Boolean 'Ticket #18790
Global gSec_Inq_AffirmAction_Purge As Boolean 'Ticket #18790

Global gSec_Upd_Payroll_Category As Boolean
Global gSec_Inq_Payroll_Category As Boolean
Global gSec_Upd_OHRSDepartments As Boolean
Global gSec_Inq_OHRSDepartments As Boolean

'7.6
Global gSec_Upd_EMP_FLAGS As Boolean
Global gSec_Inq_EMP_FLAGS As Boolean
Global gSec_Upd_EMP_HISTORY As Boolean
Global gSec_Inq_EMP_HISTORY As Boolean
Global gSec_Upd_GLDist As Boolean
Global gSec_Inq_GLDist As Boolean
Global gSec_Upd_EMP_LANG As Boolean
Global gSec_Inq_EMP_LANG As Boolean
Global gSec_Upd_SUCCESSION As Boolean
Global gSec_Inq_SUCCESSION As Boolean
Global gSec_Upd_EmergContacts As Boolean
Global gSec_Inq_EmergContacts As Boolean

Global gSec_Upd_Charge_Code As Boolean
Global gSec_Inq_Charge_Code As Boolean

Global gSec_Upd_Project_Code As Boolean
Global gSec_Inq_Project_Code As Boolean
Global gSec_Upd_Machine As Boolean
Global gSec_Inq_Machine As Boolean
Global gSec_Upd_AttendCode_Matrix As Boolean
Global gSec_Inq_AttendCode_Matrix As Boolean
Global gSec_Upd_FollowUpEmail_Matrix As Boolean
Global gSec_Inq_FollowUpEmail_Matrix As Boolean
Global gSec_Upd_DeptGL_Matrix As Boolean
Global gSec_Inq_DeptGL_Matrix As Boolean
Global gSec_Upd_BenRates As Boolean
Global gSec_Inq_BenRates As Boolean

Global gSec_Upd_Ledgers As Boolean
Global gSec_Inq_Ledgers As Boolean

Global gSec_Upd_Security As Boolean
Global gSec_Inq_Security As Boolean
Global gSec_Upd_Quick_ESS As Boolean
Global gSec_Inq_Quick_ESS As Boolean
Global gSec_Upd_Email_Setup As Boolean
Global gSec_Inq_Email_Setup As Boolean

Global gSec_Compress_Fix As Boolean
Global gSec_CompanyPreference As Boolean
Global gSec_EmpFlagsSetup As Boolean
Global gSec_MultiDataSourceSetup As Boolean
Global gSec_HelpDescSetup As Boolean
Global gSec_BenefitGroupSetup As Boolean
Global gSec_ChangeYourPassword As Boolean
Global gSec_ITAdmin As Boolean

Global gSec_Inq_Audit As Boolean
Global gSec_Upd_Audit As Boolean
'Ticket #23409 - Samuel, Son & Co., Limited - Discipline Audit Table Report
Global gSec_Inq_CounselAudit As Boolean
Global gSec_Upd_CounselAudit As Boolean
'Ticket #24655 - Wellington-Dufferin-Guelph Public Health - On Call Hours
Global gSec_Inq_OnCallHours As Boolean
Global gSec_Upd_OnCallHours As Boolean

Global gSec_Inq_DoorAccess As Boolean
Global gSec_Upd_DoorAccess As Boolean
Global gSec_Inq_CustomReport As Boolean
Global gSec_Upd_CustomReport As Boolean

Global gSec_Inq_SalaryGrids As Boolean
Global gSec_Upd_SalaryGrids As Boolean
'Global gSec_WFC_Bonus_Intergration_Interface As Boolean
Global gSec_WFC_Band_Security As Boolean
Global gSec_WFC_UnlockSmokerStatus As Boolean
'Ticket #29846 Franks 03/07/2017 ----------------- begin
Global gSec_WFC_IPExchangeRate As Boolean
Global gSec_WFC_IPIncentiveFactors As Boolean
Global gSec_WFC_IPCreateSpreadsheet As Boolean
Global gSec_WFC_IPImportSpreadsheet As Boolean
Global gSec_WFC_IPUpdateEarnings As Boolean
Global gSec_WFC_IPPreparePayrollFile As Boolean
Global gSec_WFC_IPPrintSpreadsheet As Boolean
Global gSec_WFC_IPPrintLetter As Boolean
'Ticket #29846 Franks 03/07/2017 ----------------- end
Global gSec_Upd_Holiday As Boolean
Global gSec_Inq_Holiday As Boolean
Global gSec_Upd_New_Hire As Boolean
Global gSec_Inq_New_Hire As Boolean
Global gSec_Upd_Label As Boolean
Global gSec_Inq_Label As Boolean
    
Global gSec_Mass_Codes As Boolean
        
Global gSec_Export_Attendance As Boolean
Global gSec_Export_Salaries As Boolean
Global gSec_Export_Benefits As Boolean
Global gSec_Export_Employee As Boolean
Global gSec_Export_Table As Boolean
Global gSec_Export_YTD As Boolean
Global gSec_Export_PayrollTrans As Boolean
Global gSec_Export_ContEdu As Boolean

Global gSec_Import_Attendance As Boolean
Global gSec_Import_Salaries As Boolean
Global gSec_Import_Benefits As Boolean
Global gSec_Import_Employee As Boolean
Global gSec_Import_Table As Boolean
Global gSec_Import_YTD As Boolean
Global gSec_Import_PayrollTrans As Boolean
Global gSec_Import_ContEdu As Boolean

'Ticket #29122 - New Database Setup and Integration Setup securities
Global gSec_Inq_IntegrtDBSetup As Boolean
Global gSec_Upd_IntegrtDBSetup As Boolean
Global gSec_Inq_IntegrtSetup As Boolean
Global gSec_Upd_IntegrtSetup As Boolean

Global gSec_Province As Boolean
Global gSec_Entitle As Boolean
Global gSec_Matrix As Boolean
Global gSec_DoorName As Boolean
Global gSec_Summarize_Attendance As Boolean

'Ticket #30508 - Applicant Tracking Enhancement
Global gSec_Inq_LettersPosType As Boolean
Global gSec_Upd_LettersPosType As Boolean
Global gSec_Inq_AppFormWorkFlow As Boolean
Global gSec_Upd_AppFormWorkFlow As Boolean
Global gSec_Inq_AppFormDefaults As Boolean
Global gSec_Upd_AppFormDefaults As Boolean

Global gLast_Date
Global gvarLast_Time$
Global gLast_User&

Global glbWFC As Boolean
Global glbSamuel As Boolean
Global glbMitchellPlastics As Boolean
Global glbVadim As Boolean
Global glbInsync As Boolean
Global glbWSIBModule  As Boolean

Global glbDocName As String
Global glbDocKey As String
Global glbDocTmp As String
Global glbEmpFlagNo As Integer
Global glbEmpFlagDate As Date
Global glbEmpFlag As String
Global glbAttReason As String
Global glbAttDOA As Date
Global glbAssocCode As String
Global glbBeginDt As Date
Global glbLOAComments As String

'REPORTS
'===========
Global gSec_Rpt_Age As Boolean
Global gSec_Rpt_Benefits As Boolean
Global gSec_Rpt_Compensatory_Time As Boolean
Global gSec_Rpt_Cost_Of_Employment As Boolean
Global gSec_Rpt_Emergecy_Contacts As Boolean
Global gSec_Rpt_Employee_Labels As Boolean
Global gSec_Rpt_Job_List As Boolean
Global gSec_Rpt_Profiles As Boolean
Global gSec_Rpt_Entitlements As Boolean
Global gSec_Rpt_Follow_Ups As Boolean
Global gSec_Rpt_Home_Address As Boolean
Global gSec_Rpt_Dependents As Boolean     'Laura oct 27, 1997
Global gSec_Rpt_Salary_Performance As Boolean
Global gSec_Rpt_Staff_Profile As Boolean   'Ticket #27795 - Friesens Corporation
Global gSec_Rpt_Seniority As Boolean
Global gSec_Rpt_Skills As Boolean   'laura oct 23, 1997
Global gSec_Rpt_Languages As Boolean ' laura nov 3, 1997
Global gSec_Rpt_Telephone_Extensions As Boolean
Global gSec_Rpt_Associations As Boolean
Global gSec_Rpt_Master_Attendance As Boolean
Global gSec_Rpt_Bonus_Attendance As Boolean
Global gSec_Rpt_Calendar_Attendance As Boolean
Global gSec_Rpt_Costed_Attendance As Boolean
Global gSec_Rpt_Master_Benefits As Boolean
Global gSec_Rpt_Master_Division As Boolean
Global gSec_Rpt_Master_DolEnt As Boolean
Global gSec_Rpt_Master_Education_Seminars As Boolean
Global gSec_Rpt_Master_Formal_Education As Boolean
Global gSec_Rpt_Training_Plan As Boolean
Global gSec_Rpt_Master_Job As Boolean
Global gSec_Rpt_Master_OtherEarn As Boolean
Global gSec_Rpt_Master_HourEnt As Boolean
Global gSec_Rpt_Master_Passwords As Boolean
Global gSec_Rpt_Master_Salaries As Boolean
Global gSec_Rpt_Master_Table_Codes As Boolean
Global gSec_Rpt_Master_Termination As Boolean
Global gSec_Rpt_Heatlh_Safety As Boolean
Global gSec_Rpt_Turnover As Boolean   'laura nov 17, 1997
Global gSec_Rpt_Counselling As Boolean
Global gSec_Rpt_DocumentType As Boolean
Global gSec_Rpt_DoorAccess As Boolean
Global gSec_Rpt_Emergency_Leave As Boolean
Global gSec_Rpt_External_Hire As Boolean
Global gSec_Rpt_Internal_Hire As Boolean
Global gSec_Rpt_Key_Workforce As Boolean
Global gSec_Rpt_Manpower_Plan As Boolean
Global gSec_Rpt_Staff_Management As Boolean
Global gSec_Rpt_WC_Time As Boolean
Global gSec_Rpt_WC_Work As Boolean
Global gSec_Rpt_Paid_Sick As Boolean
Global gSec_Rpt_User_Defined_Table As Boolean
Global gsec_rpt_Future_Entitlement As Boolean
Global gSec_Rpt_Employee_Flags As Boolean
Global gSec_Rpt_Temp_CrossTraining As Boolean
Global gSec_Rpt_Req_Course_Hist As Boolean
Global gSec_Rpt_Friesens_IWantToKnowYou As Boolean
Global gSec_Rpt_Friesens_ITHireForm As Boolean
Global gSec_Rpt_Friesens_ITNoticeOfChange As Boolean
Global gSec_Rpt_Friesens_NoticeOfChange As Boolean
Global gSec_Rpt_Friesens_PerfImproveActionPlan As Boolean
Global gSec_Rpt_Friesens_PerformanceReviewRpt As Boolean
Global gSec_Rpt_Friesens_SeparationRpt As Boolean
Global gSec_Rpt_Friesens_TerminationRpt As Boolean
Global gSec_Rpt_Friesens_UpdateMeetingRpt As Boolean
Global gSec_Rpt_Friesens_WarningRpt As Boolean
Global gSec_Rpt_AffirmAction As Boolean
Global gSec_Rpt_Work_Schedule As Boolean

Global gSec_Rpt_EmailAddress As Boolean
Global gSec_Rpt_LOA As Boolean
Global gSec_Rpt_POE As Boolean
Global gSec_Rpt_SINSSN As Boolean
Global gSec_Rpt_Succession As Boolean
Global gSec_Rpt_GapAnalysis As Boolean

Global gSec_Rpt_GLDistribution As Boolean

Global gSec_Rpt_Attendance_Hist As Boolean
Global gSec_Rpt_Comments As Boolean
Global gSec_Rpt_Employee_Hist As Boolean
Global gSec_Rpt_Payroll_Trans As Boolean
Global gSec_Rpt_AttWrkSch_Descrepancy As Boolean
Global gSec_Rpt_EnviroServices As Boolean
Global gSec_Rpt_ESSReq_TransAudit As Boolean
Global gSec_Rpt_Employee_Dates As Boolean
Global gSec_Rpt_Length_Of_Service As Boolean
Global gSec_Rpt_FlexTime As Boolean     'Ticket #26576 - WDGPHU - Flex Time report

'Report Forms
Global gSec_RptF_Attendance_SignIn As Boolean
Global gSec_RptF_ATT_Discipline As Boolean
Global gSec_RptF_COC_Discipline As Boolean

'Security for WHSCC
Global gSec_Upd_WHSCC_ASL As Boolean
Global gSec_Inq_WHSCC_ASL As Boolean
Global gSec_Upd_WHSCC_BUDPOS As Boolean
Global gSec_Inq_WHSCC_BUDPOS As Boolean
Global gSec_Upd_WHSCC_USB As Boolean
Global gSec_Inq_WHSCC_USB As Boolean
Global gSec_Rpt_WHSCC_PLAN_ESTABLISMNET As Boolean
'Security for WHSCC

'Security for Samuel
Global gSec_SAM_Show_CustomFeatures As Boolean

Global gSec_Upd_Ovt_Overview As Boolean
Global gSec_Inq_Ovt_Overview As Boolean
Global gSec_Upd_Ovt_Master As Boolean
Global gSec_Inq_Ovt_Master As Boolean
Global gSec_Rpt_Ovt_Bank As Boolean
Global gSec_Rpt_Ovt_Lost_Hours As Boolean

Global gSec_Upd_ADP_Data As Boolean
Global gSec_Inq_ADP_Data As Boolean

Global gSec_Upd_CourseCodeMaster As Boolean
Global gSec_Inq_CourseCodeMaster As Boolean
Global gSec_Upd_BudgetedMP As Boolean
Global gSec_Inq_BudgetedMP As Boolean
Global gSec_Upd_ReqCourses As Boolean
Global gSec_Inq_ReqCourses As Boolean
Global gSec_Upd_BudgetedPos As Boolean
Global gSec_Inq_BudgetedPos As Boolean
Global gSec_Upd_AppProcess As Boolean
Global gSec_Inq_AppProcess As Boolean
Global gSec_Upd_WorkSchRule As Boolean
Global gSec_Inq_WorkSchRule As Boolean
Global gSec_Upd_DashboardRule As Boolean
Global gSec_Inq_DashboardRule As Boolean
Global gSec_Upd_AddPayrollIDData As Boolean
Global gSec_Inq_AddPayrollIDData As Boolean

Global gSec_Upd_LOADateChange As Boolean
Global gSec_Inq_LOADateChange As Boolean
Global gSec_Upd_Rehire As Boolean
Global gSec_Inq_Rehire As Boolean
Global gSec_Upd_EnterLeave As Boolean
Global gSec_Inq_EnterLeave As Boolean
Global gSec_Upd_HSClaimMed As Boolean
Global gSec_Inq_HSClaimMed As Boolean
Global gSec_Upd_HSContacts As Boolean
Global gSec_Inq_HSContacts As Boolean
Global gSec_Upd_HSCost As Boolean
Global gSec_Inq_HSCost As Boolean
Global gSec_Upd_HSCorrectiveAct As Boolean
Global gSec_Inq_HSCorrectiveAct As Boolean
Global gSec_Upd_HSRootCause As Boolean
Global gSec_Inq_HSRootCause As Boolean
Global gSec_Upd_HSW7CmpMst As Boolean
Global gSec_Inq_HSW7CmpMst As Boolean
Global gSec_Upd_HSW7Injury As Boolean
Global gSec_Inq_HSW7Injury As Boolean
Global gSec_Upd_HSWF9 As Boolean
Global gSec_Inq_HSWF9 As Boolean

' variables for overwrite forms
Global FormDoll%       'Laura Oct 20, 1997
Global FormAssoc%       'Laura
Global FormEduc%        'laura
Global FormOther%       'Laura

Global FormHomeAddress%   'laura  Oct 30, 1997
Global FormDepend%        'Laura  Oct 30, 1997

Global FormLanguages%     'laura nov 3, 1997
Global FormEmplPosition%  'laura nov 3, 1997

'Variables for global procedures
Global glbsnapDepts As New ADODB.Recordset
Global glbsnapDiv  As New ADODB.Recordset 'Laura  Oct 30, 1997
Global glbsnapEENames As New ADODB.Recordset 'Laura  Oct 30, 1997
Global glbHome_Snap() As New ADODB.Recordset
'variable for FETERM
Global glbchkSum  'laura nov 5, 1997
'variable for FSComp    'laura nov 28, 1997
Global glbNextEmpl
Global glbSysGen
Global glbTermTran As Boolean
'variable for Attendance and Attendance_History
Global xAttendance
Global glbSeleDeptUn, glbSeleDept, glbSeleUnion, glbSeleDiv, glbSeleSection, glbSeleAdminBy 'jdy 4/28/00
Global glbSeleLoc, glbSeleRegion, glbSeleSupCode, glbSeleVadim2
Global glbSelePEDiv, glbSelePESection
Global glbTblName As New Collection
Global glbTblDesc() As New Collection
'Globals added and declarations added July 24 1998 SBH Alpha Systems...
'for linamar only
Global gSec_Inq_Productline_Operation
Global gSec_Upd_Productline_Operation
Global gSec_Inq_LinamarSkills
Global gSec_Upd_LinamarSkills
Global glbLinHS As Boolean
Global glbLinHSDivNo, glbLinEmpNo

Global glbBatchNumber 'London CCAC only #9014

Global glbSDate 'For import attatched file, passing Start_Date George Jan 19,2006 #10266

Global glbNoAccessGrp As String

Type SYSTEM_INFO
        dwOemID As Long
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOrfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        dwReserved As Long
End Type

Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cX As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByVal lpType As Long, ByVal lpData As Any, ByVal lpcbData As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long


Public Const ERROR_SUCCESS = 0&
Public Const ERROR_FILE_NOT_FOUND = 2&
Public Const ERROR_PATH_NOT_FOUND = 3&
Public Const ERROR_ACCESS_DENIED = 5&

Public Const KEY_QUERY_VALUE = &H1

Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const READ_CONTROL = &H20000
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY)) ' And (Not SYNCHRONIZE))

Public Const glbMaxDollar = 100000000000000# 'Ticket #19697 Frank 02/09/2011

Public gsSystemDb As String     'hold the path to the system db file...
Public gsMultiLang As String   'hold the Language

'Preference setup
Public gsAttachment_DB As Boolean   'Attachment database flag
Public gsCompaRatio As Boolean   'Show Compa-Ratio
Public gsEMAIL_SENDING As Boolean
Public gsEMAIL_ONNEWHIRE As Boolean
Public gsEMAIL_ONPOSITION As Boolean 'Ticket #21444 Franks 02/10/2012
Public gsEMAIL_ONSALARY As Boolean
Public gsEMAIL_ONBENEFIT As Boolean
Public gsEMAIL_ONTERM As Boolean
Public gsEMAIL_ONREHIRE As Boolean
Public gsEMAIL_ONLEAVECHANGES As Boolean
Public gsEMAIL_ONPERFORMANCE As Boolean
Public gsEMAIL_ONDEPENDENT As Boolean
Public gsEMAIL_ONDEPEND30DAYS4_WFC As Boolean 'Ticket #22061 Franks 05/24/2012
Public gsEMAIL_ONEMPLOYEEFLAGS As Boolean   'Ticket #26934 - Oshawa Community Health Centre - Employee Flags
Public gsEMAIL_ONHSINCIDENT As Boolean 'Ticket #28664 Franks 05/30/2016
Public gsTRAININGMATRIX As Boolean
Public gsFRIESENSWORDPATH As Boolean
Public gsGPHold As Boolean
Public gsSECURED_PSW As Boolean
Public gsEMPLOYEEPHOTO As Boolean
Public gsWS_ROTATIONWEEKS As Integer
Public gsWS_ROTATIONWEEKSEFFDATE As Date
Public gsDB_CONNECT_ENCRYPT As Boolean
Public gsSMTPINFO As Boolean
Public gsFLEX_LOGIC As Boolean
Public gsDISABLE_COMPTIME As Boolean


Public giGar As Integer  'for all return values we just do not care about...
Public glGar As Long     'dito...
Public glbEntPeriodFrom, glbEntPeriodTo, glbEntExcept
Public Enum BenefitUpdateSource
    EmployeeBenefitMaster
    MassUpdateBenefit
    MassUpdateBenefitGroup
    GroupMasterAdd
    GroupMasterEdit
    GroupMasterDelete
    GroupMasterRecal
End Enum


'Access the GetUserNameA function in advapi32.dll and call the function GetUserName.
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'To write to .INI files
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
                        (ByVal lpApplicationName As String, _
                        ByVal lpKeyName As Any, _
                        ByVal lpString As Any, _
                        ByVal lpFileName As String) As Long
                        
'To read from .INI files
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
                        (ByVal lpApplicationName As String, _
                        ByVal lpKeyName As Any, _
                        ByVal lpDefault As String, _
                        ByVal lpReturnedString As String, _
                        ByVal nSize As Long, _
                        ByVal lpFileName As String) As Long

'Main routine to Dimension variables, retrieve user name and display answer.
Public Function GetCurrentWinUser()

'Dimension variables
Dim lpBuff As String * 25
Dim ret As Long
Dim UserName As String

'Get the user name minus any trailing spaces found in the name.
ret = GetUserName(lpBuff, 25)
UserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)

'Return User Name
GetCurrentWinUser = UserName
End Function

Public Function VacSickHourlyFollowUp_OLD(xReason, xDate) 'created by Zahoor(Sam) Butt  on 02/07/2006

Dim SQLQ As String
Dim SQLHRTABL As String
Dim rsEmp As New ADODB.Recordset
Dim rsFU As New ADODB.Recordset
Dim dynHRAT As New ADODB.Recordset
Dim rsTABL As New ADODB.Recordset
Dim VacOut, SickOut, HourOut



'xDate = Date_SQL(xDate)
On Error GoTo CrFollow_Err


'*****02/08/2006 Zahoor(Sam) Butt

'''''''''
SQLHRTABL = "SELECT * FROM HRTABL WHERE TB_NAME = 'FURE' AND TB_KEY ='" & xReason & "'"

rsTABL.Open SQLHRTABL, gdbAdoIhr001, adOpenStatic, adLockPessimistic
'''''''''



If Left(xReason, 3) = "VAC" Then

    SQLQ = "SELECT ED_EMPNBR,ED_PVAC,ED_VAC,ED_VACT,ED_PSICK,ED_SICK,ED_SICKT,ED_EFDATE,ED_ETDATE FROM HREMP"
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
            If rsTABL.EOF Then
    
                rsTABL.AddNew
                rsTABL("TB_COMPNO") = "001"
                rsTABL("TB_NAME") = "FURE"
                rsTABL("TB_KEY") = "VAC"
                rsTABL("TB_DESC") = "Vacation Exceeded"
                rsTABL("TB_LDATE") = Date
                rsTABL("TB_LTIME") = Time$
                rsTABL("TB_LUSER") = glbUserID
                rsTABL.Update
        
            End If
    
    If Not rsEmp.EOF Then
            
    '*** CHECK TO SEE IF TB_NAME= FURE AND TB_KEY= VAC DOES NOT ALREADY EXIST THEN ADD THEM OTHERWISE DONT ***

            If Not glbSQL And Not glbOracle Then Call Pause(0.5)
        
            rsEmp.MoveFirst
            Do Until rsEmp.EOF
                If DateValue(xDate) >= rsEmp("ED_EFDATE") And DateValue(xDate) <= rsEmp("ED_ETDATE") Then
                    VacOut = (rsEmp("ED_VAC") + rsEmp("ED_PVAC")) - (rsEmp("ED_VACT"))
                End If
          
    
        If VacOut < 0 Then
            VacOut = 0 - VacOut
            
                If Not glbSQL And Not glbOracle Then Call Pause(0.5)
            
        
      
        
        '************* Vacation FollowUp Starts here
            
                SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & rsEmp("ED_EMPNBR")
                SQLQ = SQLQ & " AND EF_FREAS = 'VAC'"
                SQLQ = SQLQ & " AND EF_FDATE =" & Date_SQL(xDate)

                dynHRAT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
                rsFU.Open "HR_FOLLOW_UP", gdbAdoIhr001, adOpenKeyset, adLockPessimistic, adCmdTableDirect
                If dynHRAT.EOF Then
                    rsFU.AddNew
                    rsFU("EF_COMPNO") = "001"
                    rsFU("EF_EMPNBR") = rsEmp("ED_EMPNBR")
                    rsFU("EF_FDATE") = xDate
                    rsFU("EF_FREAS_TABL") = "FURE"
                    'Ticket #24257 - Do not update Admin By for them only
                    If glbCompSerial <> "S/N - 2262W" Then
                        rsFU("EF_ADMINBY_TABL") = "EDAB"
                        rsFU("EF_ADMINBY") = GetEmpData(rsEmp("ED_EMPNBR"), "ED_ADMINBY", Null)
                    End If
                    rsFU("EF_FREAS") = "VAC"
                    rsFU("EF_COMMENTS") = "Employee #:" & rsEmp("ED_EMPNBR") & " Vacation Entitlement has been exceeded by (" & VacOut & ") Hours"
                    rsFU("EF_LDATE") = Date
                    rsFU("EF_LTIME") = Time$
                    rsFU("EF_LUSER") = glbUserID
                    rsFU.Update
                    'rsFU.Close
            
            '        Msg = "A Follow Up Record was created!"
            '        MsgBox Msg
                ElseIf glbflgFU = False And dynHRAT.EOF Then
                    rsFU.AddNew
                    rsFU("EF_COMPNO") = "001"
                    rsFU("EF_EMPNBR") = rsEmp("ED_EMPNBR")
                    rsFU("EF_FDATE") = xDate
                    rsFU("EF_FREAS_TABL") = "FURE"
                    rsFU("EF_FREAS") = "VAC"
                    'Ticket #24257 - Do not update Admin By for them only
                    If glbCompSerial <> "S/N - 2262W" Then
                        rsFU("EF_ADMINBY_TABL") = "EDAB"
                        rsFU("EF_ADMINBY") = GetEmpData(rsEmp("ED_EMPNBR"), "ED_ADMINBY", Null)
                    End If
                    rsFU("EF_COMMENTS") = "Employee #:" & rsEmp("ED_EMPNBR") & " Vacation Entitlement has been exceeded by (" & VacOut & ") Hours"
                    rsFU("EF_LDATE") = Date
                    rsFU("EF_LTIME") = Time$
                    rsFU("EF_LUSER") = glbUserID
                    rsFU.Update
                    'rsFU.Close
            
            '        Msg = "A Follow Up Record was created!"
            '        MsgBox Msg
                ElseIf glbflgFU = False And Not dynHRAT.EOF Then
                    dynHRAT.MoveFirst
                    Do Until dynHRAT.EOF
                        rsFU("EF_COMPNO") = "001"
                        rsFU("EF_EMPNBR") = rsEmp("ED_EMPNBR")
                        rsFU("EF_FDATE") = xDate
                        rsFU("EF_FREAS_TABL") = "FURE"
                        rsFU("EF_FREAS") = "VAC"
                        'Ticket #24257 - Do not update Admin By for them only
                        If glbCompSerial <> "S/N - 2262W" Then
                            rsFU("EF_ADMINBY_TABL") = "EDAB"
                            rsFU("EF_ADMINBY") = GetEmpData(rsEmp("ED_EMPNBR"), "ED_ADMINBY", Null)
                        End If
                        rsFU("EF_COMPLETED") = dynHRAT("EF_COMPLETED")
                        rsFU("EF_COMMENTS") = "Employee #:" & rsEmp("ED_EMPNBR") & " Vacation Entitlement has been exceeded by (" & VacOut & ") Hours"
                        rsFU("EF_LDATE") = Date
                        rsFU("EF_LTIME") = Time$
                        rsFU("EF_LUSER") = glbUserID
                        rsFU.Update
                        dynHRAT.MoveNext
                
                    Loop
            
                End If
        End If
    
         rsEmp.MoveNext
        Loop
    End If
            rsEmp.Close
End If
        'dynHRAT.Close
                        'Msg = "A Follow Up Record was updated!"
                       'MsgBox Msg
     '                   VacSickHourlyFollowUp = True
                      
                 '********* Vacation FollowUP ends here


    '************* Sick FollowUp Starts here *******************
            
If Left(xReason, 3) = "SIC" Then
        SQLQ = "SELECT ED_EMPNBR,ED_PVAC,ED_VAC,ED_VACT,ED_PSICK,ED_SICK,ED_SICKT,ED_EFDATES,ED_ETDATES FROM HREMP"
        rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
            
      '*** CHECK TO SEE IF TB_NAME= FURE AND TB_KEY= SICK AND TB_DESC='Sick Exceeded' DOES NOT ALREADY EXIST THEN ADD THEM OTHERWISE DONT ***

            If rsTABL.EOF Then
    
                rsTABL.AddNew
                rsTABL("TB_COMPNO") = "001"
                rsTABL("TB_NAME") = "FURE"
                rsTABL("TB_KEY") = "SICK"
                rsTABL("TB_DESC") = "Sick Exceeded"
                rsTABL("TB_LDATE") = Date
                rsTABL("TB_LTIME") = Time$
                rsTABL("TB_LUSER") = glbUserID
                rsTABL.Update
        
            End If
        If Not rsEmp.EOF Then
            rsEmp.MoveFirst
            Do Until rsEmp.EOF
            
                If DateValue(xDate) >= rsEmp("ED_EFDATES") And DateValue(xDate) <= rsEmp("ED_ETDATES") Then
                    SickOut = (rsEmp("ED_SICK") + rsEmp("ED_PSICK")) - (rsEmp("ED_SICKT"))
                End If
                
                
      
            If SickOut < 0 Then
                SickOut = 0 - SickOut
                If Not glbSQL And Not glbOracle Then Call Pause(0.5)
                SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & rsEmp("ED_EMPNBR")
                SQLQ = SQLQ & " AND EF_FREAS = 'SICK'"
                SQLQ = SQLQ & " AND EF_FDATE =" & Date_SQL(xDate)

                dynHRAT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
                rsFU.Open "HR_FOLLOW_UP", gdbAdoIhr001, adOpenKeyset, adLockPessimistic, adCmdTableDirect
            
                If dynHRAT.BOF And dynHRAT.EOF Then
                    rsFU.AddNew
                    rsFU("EF_COMPNO") = "001"
                    rsFU("EF_EMPNBR") = rsEmp("ED_EMPNBR")
                    rsFU("EF_FDATE") = xDate
                    rsFU("EF_FREAS_TABL") = "FURE"
                    rsFU("EF_FREAS") = "SICK"
                    'Ticket #24257 - Do not update Admin By for them only
                    If glbCompSerial <> "S/N - 2262W" Then
                        rsFU("EF_ADMINBY_TABL") = "EDAB"
                        rsFU("EF_ADMINBY") = GetEmpData(rsEmp("ED_EMPNBR"), "ED_ADMINBY", Null)
                    End If
                    rsFU("EF_COMMENTS") = "Employee #:" & rsEmp("ED_EMPNBR") & " Sick Entitlement has been exceeded by (" & SickOut & ") Hours"
                    rsFU("EF_LDATE") = Date
                    rsFU("EF_LTIME") = Time$
                    rsFU("EF_LUSER") = glbUserID
                    rsFU.Update
                    'rsFU.Close
                    'VacSickHourlyFollowUp = True
        '        Msg = "A Follow Up Record was created!"
        '        MsgBox Msg
                
                ElseIf glbflgFU = False And dynHRAT.EOF Then
                    rsFU.AddNew
                    rsFU("EF_COMPNO") = "001"
                    rsFU("EF_EMPNBR") = rsEmp("ED_EMPNBR")
                    rsFU("EF_FDATE") = xDate
                    rsFU("EF_FREAS_TABL") = "FURE"
                    rsFU("EF_FREAS") = "SICK"
                    'Ticket #24257 - Do not update Admin By for them only
                    If glbCompSerial <> "S/N - 2262W" Then
                        rsFU("EF_ADMINBY_TABL") = "EDAB"
                        rsFU("EF_ADMINBY") = GetEmpData(rsEmp("ED_EMPNBR"), "ED_ADMINBY", Null)
                    End If
                    rsFU("EF_COMMENTS") = "Employee #:" & rsEmp("ED_EMPNBR") & " Sick Entitlement has been exceeded by (" & SickOut & ") Hours"
                    rsFU("EF_LDATE") = Date
                    rsFU("EF_LTIME") = Time$
                    rsFU("EF_LUSER") = glbUserID
                    rsFU.Update
                    'rsFU.Close
                    'VacSickHourlyFollowUp = True
        '        Msg = "A Follow Up Record was created!"
        '        MsgBox Msg
                
                
                
                ElseIf glbflgFU = False And Not dynHRAT.EOF Then
                    dynHRAT.MoveFirst
                    Do Until dynHRAT.EOF
                        rsFU("EF_COMPNO") = "001"
                        rsFU("EF_EMPNBR") = rsEmp("ED_EMPNBR")
                        rsFU("EF_FDATE") = xDate
                        rsFU("EF_FREAS_TABL") = "FURE"
                        rsFU("EF_FREAS") = "SICK"
                        'Ticket #24257 - Do not update Admin By for them only
                        If glbCompSerial <> "S/N - 2262W" Then
                            rsFU("EF_ADMINBY_TABL") = "EDAB"
                            rsFU("EF_ADMINBY") = GetEmpData(rsEmp("ED_EMPNBR"), "ED_ADMINBY", Null)
                        End If
                        rsFU("EF_COMPLETED") = dynHRAT("EF_COMPLETED")
                        rsFU("EF_COMMENTS") = "Employee #:" & rsEmp("ED_EMPNBR") & " Sick Entitlement has been exceeded by (" & SickOut & ") Hours"
                        rsFU("EF_LDATE") = Date
                        rsFU("EF_LTIME") = Time$
                        rsFU("EF_LUSER") = glbUserID
                        rsFU.Update
                        dynHRAT.MoveNext

                    Loop
                End If
            
            End If
            rsEmp.MoveNext
            Loop
            
        End If
End If
     'dynHRAT.Close
                    'Msg = "A Follow Up Record was updated!"
                    'MsgBox Msg
                    'VacSickHourlyFollowUp = True
            
               
            
         '********* Sick FollowUP ends here
        
    
If Left(xReason, 3) <> "SIC" And Left(xReason, 3) <> "VAC" Then '*** Hourly FollowUp Starts here
    
    SQLQ = " SELECT HE_EMPNBR,HE_TYPE,HE_ENTITLE,HE_TAKEN FROM HRENTHRS" ' WHERE HE_EMPNBR=" & rsEmp!ED_EMPNBR
    SQLQ = SQLQ & " WHERE HE_TYPE ='" & xReason & "'"
    SQLQ = SQLQ & " AND HE_FDATE<= " & Date_SQL(xDate)
    SQLQ = SQLQ & " AND HE_TDATE>= " & Date_SQL(xDate)

    
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
      '*** CHECK TO SEE IF TB_NAME= FURE AND TB_KEY= for hourly  DOES NOT ALREADY EXIST THEN ADD THEM OTHERWISE DONT ***

            If rsTABL.EOF Then
    
                rsTABL.AddNew
                rsTABL("TB_COMPNO") = "001"
                rsTABL("TB_NAME") = "FURE"
                rsTABL("TB_KEY") = xReason
                rsTABL("TB_DESC") = "Hourly Exceeded"
                rsTABL("TB_LDATE") = Date
                rsTABL("TB_LTIME") = Time$
                rsTABL("TB_LUSER") = glbUserID
                rsTABL.Update
        
            End If
    
    If Not rsEmp.EOF Then
    
      Do Until rsEmp.EOF
        HourOut = (rsEmp("HE_ENTITLE")) - (rsEmp("HE_TAKEN"))
        
        If HourOut < 0 Then
            HourOut = 0 - HourOut
            If Not glbSQL And Not glbOracle Then Call Pause(0.5)
      
                
            
                SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & rsEmp("HE_EMPNBR")
                SQLQ = SQLQ & " AND EF_FREAS ='" & xReason & "'"
                SQLQ = SQLQ & " AND EF_FDATE =" & Date_SQL(xDate)

                dynHRAT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
                rsFU.Open "HR_FOLLOW_UP", gdbAdoIhr001, adOpenKeyset, adLockPessimistic, adCmdTableDirect
            
            If dynHRAT.EOF Then
                rsFU.AddNew
                rsFU("EF_COMPNO") = "001"
                rsFU("EF_EMPNBR") = rsEmp("HE_EMPNBR")
                rsFU("EF_FDATE") = Date_SQL(xDate)
                rsFU("EF_FREAS_TABL") = "FURE"
                rsFU("EF_FREAS") = xReason
                'Ticket #24257 - Do not update Admin By for them only
                If glbCompSerial <> "S/N - 2262W" Then
                    rsFU("EF_ADMINBY_TABL") = "EDAB"
                    rsFU("EF_ADMINBY") = GetEmpData(rsEmp("HE_EMPNBR"), "ED_ADMINBY", Null)
                End If
                rsFU("EF_COMMENTS") = "Employee #:" & rsEmp("HE_EMPNBR") & " Sick Entitlement has been exceeded by (" & HourOut & ") Hours"
                rsFU("EF_LDATE") = Date
                rsFU("EF_LTIME") = Time$
                rsFU("EF_LUSER") = glbUserID
                rsFU.Update
                'rsFU.Close
            '    VacSickHourlyFollowUp = True
        '        Msg = "A Follow Up Record was created!"
        '        MsgBox Msg
            ElseIf glbflgFU = False And dynHRAT.EOF Then
                rsFU.AddNew
                rsFU("EF_COMPNO") = "001"
                rsFU("EF_EMPNBR") = rsEmp("HE_EMPNBR")
                rsFU("EF_FDATE") = Date_SQL(xDate)
                rsFU("EF_FREAS_TABL") = "FURE"
                rsFU("EF_FREAS") = xReason
                'Ticket #24257 - Do not update Admin By for them only
                If glbCompSerial <> "S/N - 2262W" Then
                    rsFU("EF_ADMINBY_TABL") = "EDAB"
                    rsFU("EF_ADMINBY") = GetEmpData(rsEmp("HE_EMPNBR"), "ED_ADMINBY", Null)
                End If
                rsFU("EF_COMMENTS") = "Employee #:" & rsEmp("HE_EMPNBR") & " Sick Entitlement has been exceeded by (" & HourOut & ") Hours"
                rsFU("EF_LDATE") = Date
                rsFU("EF_LTIME") = Time$
                rsFU("EF_LUSER") = glbUserID
                rsFU.Update
                'rsFU.Close
                'VacSickHourlyFollowUp = True
        '        Msg = "A Follow Up Record was created!"
        '        MsgBox Msg
            
            ElseIf glbflgFU = False And Not dynHRAT.EOF Then
                dynHRAT.MoveFirst
                Do Until dynHRAT.EOF
                    rsFU("EF_COMPNO") = "001"
                    rsFU("EF_EMPNBR") = rsEmp("HE_EMPNBR")
                    rsFU("EF_FDATE") = Date_SQL(xDate)
                    rsFU("EF_FREAS_TABL") = "FURE"
                    rsFU("EF_FREAS") = xReason
                    'Ticket #24257 - Do not update Admin By for them only
                    If glbCompSerial <> "S/N - 2262W" Then
                        rsFU("EF_ADMINBY_TABL") = "EDAB"
                        rsFU("EF_ADMINBY") = GetEmpData(rsEmp("HE_EMPNBR"), "ED_ADMINBY", Null)
                    End If
                    rsFU("EF_COMPLETED") = dynHRAT("EF_COMPLETED")
                    rsFU("EF_COMMENTS") = "Employee #:" & rsEmp("HE_EMPNBR") & " Hourly Entitlement has been exceeded by (" & HourOut & ") Hours"
                    rsFU("EF_LDATE") = Date
                    rsFU("EF_LTIME") = Time$
                    rsFU("EF_LUSER") = glbUserID
                    rsFU.Update
                    dynHRAT.MoveNext
                        
                Loop
            End If
           
            
        End If
            rsEmp.MoveNext
            Loop
    
    End If
                    'dynHRAT.Close
                    'Msg = "A Follow Up Record was updated!"
                    'MsgBox Msg
      '              VacSickHourlyFollowUp = True
            
            
            
         
    
    
End If '********* Hourly FollowUP ends here


Exit Function


'*************
CrFollow_Err:
If Err = 3022 Then
    MsgBox "The record is not entered"
    Err = 0
    Resume Next
    Exit Function
End If



glbFrmCaption$ = "Vaction/Sick/Hourly FollowUP"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Vacation Follow UP", "HR_FOLLOW_UP", "UPDATE TABLE")
Resume Next
      
     
End Function

Public Sub GenTabl(zName, zCode, Optional zDesc)
Dim rsTABL As New ADODB.Recordset
Dim SQLQ
    If zCode = "" Then Exit Sub
    SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = '" & zName & "' AND TB_KEY = '" & zCode & "' "
    rsTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsTABL.EOF Then
        rsTABL.AddNew
        rsTABL!TB_COMPNO = "001"
        rsTABL!TB_NAME = zName
        rsTABL!TB_KEY = zCode
        If Not (IsMissing(zDesc) Or IsEmpty(zDesc)) Then
           rsTABL!TB_DESC = zDesc
        End If
        rsTABL!TB_LDATE = Format(Now, "Short Date")
        rsTABL!TB_LTIME = Time$
        rsTABL!TB_LUSER = glbUserID
        rsTABL.Update
    End If
    rsTABL.Close
End Sub

Public Function VacSickHourlyFollowUp(ByVal WSQLQ As String, Optional xAction, Optional AttReason, Optional AttDOA) 'created by Zahoor(Sam) Butt  on 02/07/2006
Dim xEmpNo, xReason, xDate, xFDate, xTDate
Dim SQLQ As String
Dim SQLHRTABL As String
Dim rsEmp As New ADODB.Recordset
Dim rsATT As New ADODB.Recordset
Dim rsHRSENT As New ADODB.Recordset
Dim rsFU As New ADODB.Recordset
Dim dynHRAT As New ADODB.Recordset
Dim rsTABL As New ADODB.Recordset
Dim VacOut, SickOut, HourOut

On Error GoTo CrFollow_Err
If IsMissing(xAction) Then
    WSQLQ = glbSeleDeptUn & IIf(WSQLQ = "", " ", " AND ") & WSQLQ
    WSQLQ = Replace(WSQLQ, "ED_", "HREMP.ED_")
    SQLQ = "SELECT ED_EMPNBR,ED_PVAC,ED_VAC,ED_VACT,ED_PSICK,ED_SICK,ED_SICKT,ED_EFDATE,ED_ETDATE,ED_EFDATES,ED_ETDATES FROM HREMP "
    SQLQ = SQLQ & "WHERE " & WSQLQ
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
    If Not rsEmp.EOF Then
        Call GenTabl("FURE", "VAC", "Vacation Exceeded")
        Call GenTabl("FURE", "SICK", "Sick Exceeded")
        'Release 8.0 - Grant permission to this Follow Up for this user as well so the user can see the
        'follow up record
        Call Grant_FollowUpCode_Security(glbUserID, "VAC", "Vacation Exceeded")
        Call Grant_FollowUpCode_Security(glbUserID, "SICK", "Sick Exceeded")
    End If
    Do While Not rsEmp.EOF
        xEmpNo = rsEmp("ED_EMPNBR")
        'VAC -- Begin
        If IsDate(rsEmp("ED_EFDATE")) Then
            xFDate = rsEmp("ED_EFDATE")
        Else
            xFDate = ""
        End If
        If IsDate(rsEmp("ED_ETDATE")) Then
            xTDate = rsEmp("ED_ETDATE")
        Else
            xTDate = ""
        End If
        
        If Not IsMissing(AttReason) And Not IsMissing(AttDOA) Then
            If IsDate(xFDate) And IsDate(xTDate) And Left(AttReason, 3) = "VAC" Then
                If CVDate(AttDOA) >= CVDate(xFDate) And CVDate(AttDOA) <= CVDate(xTDate) Then
                    'Continue with the rest of the procedure
                Else
                    GoTo Exit_FollowUp
                End If
            Else
                GoTo Sick_FollowUp
            End If
        End If
        
        If IsDate(xFDate) And IsDate(xTDate) Then
            'No follow up record for 0 or Null or Blank Current Vacation entitlement. - Ticket #13221
            If Not IsNull(rsEmp("ED_VAC")) And rsEmp("ED_VAC") <> "" And rsEmp("ED_VAC") > 0 Then
                VacOut = (rsEmp("ED_VAC") + rsEmp("ED_PVAC")) - (rsEmp("ED_VACT"))
                If VacOut < 0 Then
                    VacOut = 0 - VacOut
                    '************* Vacation FollowUp Starts here
                    'Ticket #11651 Dont use Attendance table, it will solw down the system
                    'SQLQ = "SELECT AD_EMPNBR,AD_DOA FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & xEmpNo & " "
                    'SQLQ = SQLQ & "AND LEFT(AD_REASON,3)= 'VAC' "
                    'SQLQ = SQLQ & "AND AD_DOA >= " & Date_SQL(xFDate) & " "
                    'SQLQ = SQLQ & "AND AD_DOA <= " & Date_SQL(xTDate) & " "
                    'SQLQ = SQLQ & "ORDER BY AD_DOA DESC"
                    'rsATT.Open SQLQ, gdbAdoIhr001, adOpenStatic
                    'If Not rsATT.EOF Then
                        'xDATE = rsATT("AD_DOA")
                        xDate = xFDate
                        SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & xEmpNo
                        SQLQ = SQLQ & " AND EF_FREAS = 'VAC'"
                        
                        'Hemu - Jerry asked me to make this change - save System date for Followup Effective Date
                        'Ticket #12425
                        'SQLQ = SQLQ & "AND EF_FDATE >= " & Date_SQL(xFDate) & " "
                        'SQLQ = SQLQ & "AND EF_FDATE <= " & Date_SQL(xTDate) & " "
                        SQLQ = SQLQ & "AND EF_FDATE <= " & Date_SQL(Date) & " "
                        
                        rsFU.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
                        If rsFU.EOF Then
                            rsFU.AddNew
                            rsFU("EF_COMPNO") = "001"
                            rsFU("EF_EMPNBR") = xEmpNo
                            rsFU("EF_FREAS_TABL") = "FURE"
                            rsFU("EF_FREAS") = "VAC"
                            'Ticket #24257 - Do not update Admin By for them only
                            If glbCompSerial <> "S/N - 2262W" Then
                                rsFU("EF_ADMINBY_TABL") = "EDAB"
                                rsFU("EF_ADMINBY") = GetEmpData(xEmpNo, "ED_ADMINBY", Null)
                            End If
                        End If
                        'Hemu - Jerry asked me to make this change - save System date for Followup Effective Date
                        'Ticket #12425
                        'rsFU("EF_FDATE") = xDATE
                        rsFU("EF_FDATE") = CVDate(Date)
                        
                        rsFU("EF_COMMENTS") = "Employee #:" & xEmpNo & " Vacation Entitlement has been exceeded by (" & VacOut & ") Hours"
                        rsFU("EF_LDATE") = Date
                        rsFU("EF_LTIME") = Time$
                        rsFU("EF_LUSER") = glbUserID
                        rsFU.Update
                        rsFU.Close
                        
                    'End If
                    'rsATT.Close
                End If
            End If
        End If
        'VAC -- End
        
Sick_FollowUp:
        'SICK -- Begin
        If IsDate(rsEmp("ED_EFDATES")) Then
            xFDate = rsEmp("ED_EFDATES")
        Else
            xFDate = ""
        End If
        If IsDate(rsEmp("ED_ETDATES")) Then
            xTDate = rsEmp("ED_ETDATES")
        Else
            xTDate = ""
        End If
        
        If Not IsMissing(AttReason) And Not IsMissing(AttDOA) Then
            If IsDate(xFDate) And IsDate(xTDate) And Left(AttReason, 3) = "SIC" Then
                If CVDate(AttDOA) >= CVDate(xFDate) And CVDate(AttDOA) <= CVDate(xTDate) Then
                    'Continue with the rest of the procedure
                Else
                    GoTo Exit_FollowUp
                End If
            Else
                GoTo Hourly_FollowUp
            End If
        End If
        
        If IsDate(xFDate) And IsDate(xTDate) Then
            'No follow up record for 0 or Null or Blank Current Sick entitlement. - Ticket #13221
            If Not IsNull(rsEmp("ED_SICK")) And rsEmp("ED_SICK") <> "" And rsEmp("ED_SICK") > 0 Then
                SickOut = (rsEmp("ED_SICK") + rsEmp("ED_PSICK")) - (rsEmp("ED_SICKT"))
                If SickOut < 0 Then
                    SickOut = 0 - SickOut 'VacOut
                    '************* Vacation FollowUp Starts here
                    'Ticket #11651 Dont use Attendance table, it will solw down the system
                    'SQLQ = "SELECT AD_EMPNBR,AD_DOA FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & xEmpNo & " "
                    'SQLQ = SQLQ & "AND LEFT(AD_REASON,3)='SIC' "
                    'SQLQ = SQLQ & "AND AD_DOA >= " & Date_SQL(xFDate) & " "
                    'SQLQ = SQLQ & "AND AD_DOA <= " & Date_SQL(xTDate) & " "
                    'SQLQ = SQLQ & "ORDER BY AD_DOA DESC"
                    'rsATT.Open SQLQ, gdbAdoIhr001, adOpenStatic
                    'If Not rsATT.EOF Then
                        'xDATE = rsATT("AD_DOA")
                        xDate = xFDate
                        SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & xEmpNo
                        SQLQ = SQLQ & " AND EF_FREAS = 'SICK'"
                        
                        'Hemu - Jerry asked me to make this change - save System date for Followup Effective Date
                        'Ticket #12425
                        'SQLQ = SQLQ & "AND EF_FDATE >= " & Date_SQL(xFDate) & " "
                        'SQLQ = SQLQ & "AND EF_FDATE <= " & Date_SQL(xTDate) & " "
                        SQLQ = SQLQ & "AND EF_FDATE = " & Date_SQL(Date) & " "
                        
                        rsFU.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
                        If rsFU.EOF Then
                            rsFU.AddNew
                            rsFU("EF_COMPNO") = "001"
                            rsFU("EF_EMPNBR") = xEmpNo
                            rsFU("EF_FREAS_TABL") = "FURE"
                            rsFU("EF_FREAS") = "SICK"
                            'Ticket #24257 - Do not update Admin By for them only
                            If glbCompSerial <> "S/N - 2262W" Then
                                rsFU("EF_ADMINBY_TABL") = "EDAB"
                                rsFU("EF_ADMINBY") = GetEmpData(xEmpNo, "ED_ADMINBY", Null)
                            End If
                        End If
                        'Hemu - Jerry asked me to make this change - save System date for Followup Effective Date
                        'Ticket #12425
                        'rsFU("EF_FDATE") = xDATE
                        rsFU("EF_FDATE") = CVDate(Date)
                        
                        rsFU("EF_COMMENTS") = "Employee #:" & xEmpNo & " Sick Entitlement has been exceeded by (" & SickOut & ") Hours"
                        rsFU("EF_LDATE") = Date
                        rsFU("EF_LTIME") = Time$
                        rsFU("EF_LUSER") = glbUserID
                        rsFU.Update
                        rsFU.Close
                        
                    'End If
                    'rsATT.Close
                End If
            End If
        End If
        'SICK -- End
        
Hourly_FollowUp:
        'Hourly Entitlement - Begin
        SQLQ = " SELECT HE_EMPNBR,HE_TYPE,HE_ENTITLE,HE_TAKEN,HE_FDATE,HE_TDATE FROM HRENTHRS WHERE HE_EMPNBR=" & xEmpNo & " "
        If Not IsMissing(AttReason) Then
            SQLQ = SQLQ & " AND HE_TYPE = '" & AttReason & "'"
        End If
        SQLQ = SQLQ & "ORDER BY HE_TYPE, HE_FDATE DESC "
        rsHRSENT.Open SQLQ, gdbAdoIhr001, adOpenStatic
        Do While Not rsHRSENT.EOF
            If IsDate(rsHRSENT("HE_FDATE")) Then
                xFDate = rsHRSENT("HE_FDATE")
            Else
                xFDate = ""
            End If
            If IsDate(rsHRSENT("HE_TDATE")) Then
                xTDate = rsHRSENT("HE_TDATE")
            Else
                xTDate = ""
            End If
            
            If Not IsMissing(AttReason) And Not IsMissing(AttDOA) Then
                If IsDate(xFDate) And IsDate(xTDate) Then
                    If CVDate(AttDOA) >= CVDate(xFDate) And CVDate(AttDOA) <= CVDate(xTDate) Then
                        'Continue with the rest of the procedure
                    Else
                        GoTo Exit_FollowUp
                    End If
                Else
                    GoTo Exit_FollowUp
                End If
            End If
            
            If IsDate(xFDate) And IsDate(xTDate) Then
                'Only do this for Current year Hourly Entitlements - Ticket #13221
                'If CVDate(xFDate) >= CVDate("1/1/" & Year(Now)) And xTDate <= CVDate("12/31/" & Year(Now)) Then
                    HourOut = (rsHRSENT("HE_ENTITLE")) - (rsHRSENT("HE_TAKEN"))
                    If HourOut < 0 Then
                        HourOut = 0 - HourOut
                        'Ticket #11651 Dont use Attendance table, it will solw down the system
                        'SQLQ = "SELECT AD_EMPNBR,AD_DOA FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & xEmpNo & " "
                        'SQLQ = SQLQ & "AND AD_REASON='" & rsHRSENT("HE_TYPE") & "' "
                        'SQLQ = SQLQ & "AND AD_DOA >= " & Date_SQL(xFDate) & " " '
                        'SQLQ = SQLQ & "AND AD_DOA <= " & Date_SQL(xTDate) & " "
                        'SQLQ = SQLQ & "ORDER BY AD_DOA DESC"
                        'rsATT.Open SQLQ, gdbAdoIhr001, adOpenStatic
                        'If Not rsATT.EOF Then
                            'xDATE = rsATT("AD_DOA")
                            xDate = xTDate
                            SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & xEmpNo
                            SQLQ = SQLQ & " AND EF_FREAS = '" & rsHRSENT("HE_TYPE") & "'"
                            
                            'Hemu - Jerry asked me to make this change - save System date for Followup Effective Date
                            'Ticket #12425
                            'SQLQ = SQLQ & "AND EF_FDATE >= " & Date_SQL(xFDate) & " "
                            'SQLQ = SQLQ & "AND EF_FDATE <= " & Date_SQL(xTDate) & " "
                            SQLQ = SQLQ & "AND EF_FDATE = " & Date_SQL(Date) & " "
                
                            rsFU.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
                            If rsFU.EOF Then
                                rsFU.AddNew
                                rsFU("EF_COMPNO") = "001"
                                rsFU("EF_EMPNBR") = xEmpNo
                                rsFU("EF_FREAS_TABL") = "FURE"
                                rsFU("EF_FREAS") = rsHRSENT("HE_TYPE")
                                'Ticket #24257 - Do not update Admin By for them only
                                If glbCompSerial <> "S/N - 2262W" Then
                                    rsFU("EF_ADMINBY_TABL") = "EDAB"
                                    rsFU("EF_ADMINBY") = GetEmpData(xEmpNo, "ED_ADMINBY", Null)
                                End If
                            End If
                            'Hemu - Jerry asked me to make this change - save System date for Followup Effective Date
                            'Ticket #12425
                            'rsFU("EF_FDATE") = xDATE
                            rsFU("EF_FDATE") = CVDate(Date)
                            
                            rsFU("EF_COMMENTS") = "Employee #:" & xEmpNo & " Hourly Entitlement has been exceeded by (" & HourOut & ") Hours"
                            rsFU("EF_LDATE") = Date
                            rsFU("EF_LTIME") = Time$
                            rsFU("EF_LUSER") = glbUserID
                            rsFU.Update
                            rsFU.Close
                            
                        'End If
                        'rsATT.Close
                        Call GenTabl("FURE", rsHRSENT("HE_TYPE"), rsHRSENT("HE_TYPE") & " Exceeded")
                        'Release 8.0 - Grant permission to this Follow Up for this user as well so the user can see the
                        'follow up record
                        Call Grant_FollowUpCode_Security(glbUserID, rsHRSENT("HE_TYPE"), rsHRSENT("HE_TYPE") & " Exceeded")
                    End If
                'End If
            End If
            rsHRSENT.MoveNext
        Loop
        rsHRSENT.Close
        
        'Hourly Entitlement - End
        rsEmp.MoveNext
    Loop
    rsEmp.Close
Else 'Attendance Deleted and he related followup record to be deleted too
    'Optional xAction, Optional AttReason, Optional AttDOA)
    If xAction = "Delete" Then
        WSQLQ = Replace(WSQLQ, "ED_", "EF_")
        SQLQ = "DELETE FROM HR_FOLLOW_UP WHERE " & WSQLQ & " "
        SQLQ = SQLQ & " AND EF_FREAS = '" & AttReason & "'"
        SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(AttDOA) & " "
        gdbAdoIhr001.Execute SQLQ
    End If
End If

Exit_FollowUp:

Exit Function

'*************
CrFollow_Err:
If Err = 3022 Then
    MsgBox "The record is not entered"
    Err = 0
    Resume Next
    Exit Function
End If

glbFrmCaption$ = "Vaction/Sick/Hourly FollowUP"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Entitlement Follow UP", "HR_FOLLOW_UP", "UPDATE TABLE")
Resume Next
    
End Function

'Release 8.0 - Ticket #22682: Delete the Exceeding Follow Up records of Previous Years
Sub Delete_Exceeding_FollowUp(xEmpNo, xFollowUpCode, xYear)
Dim SQLQ As String
    
    SQLQ = "DELETE FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & xEmpNo
    SQLQ = SQLQ & " AND EF_FREAS = '" & xFollowUpCode & "'"
    SQLQ = SQLQ & " AND YEAR(EF_FDATE) < " & xYear & " "
    SQLQ = SQLQ & " AND EF_COMMENTS like '% has been exceeded %'"
    gdbAdoIhr001.Execute SQLQ

End Sub

Sub ApplicationEnd()
Dim Response%, Msg$, Title$, DgDef As Double
Msg$ = "Do you wish to end info:HR? " & Chr(10)
'Msg$ = Msg$ & "Any changes pending will be disregarded."
Title$ = "info:HR - END APPLICATION"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
If Not Response% = IDYES Then    ' Evaluate response
   Exit Sub
End If
End
End Sub

Sub ERR_Hndlr(ByVal iErrorNumber As Long, FName As String, WDesc As String, TabName As String, TCall As String)
Dim Msg  As String, intLock As Integer
Dim DgDef As Long, Response As Integer, Title  As String
Dim strMsg As String, intMsgType  As Integer
Dim xStr, x
intLock% = False
'Jaddy change for multi employee selection
If Len(glbstrSelCri) >= 1000 Then
    xStr = glbstrSelCri
    x = 1
    Do Until InStr(xStr, ",") = 0
        xStr = Mid(xStr, InStr(xStr, ",") + 1)
        x = x + 1
    Loop
    If x > 1000 Then
        Screen.MousePointer = DEFAULT
        Msg = x & " employees are typed in the Employee Number Selection Criteria." & Chr(10)
        Msg = Msg & "Please make that less than 1000."
        Title = "info:HR - ERROR IN APPLICATION"
        DgDef = MB_OK + MB_ICONSTOP + MB_DEFBUTTON2
        Response = MsgBox(Msg, DgDef, Title)    ' Get user response.
        Err = 0
        gintRollBack% = False
        Exit Sub
    End If
End If

' dkostka - 03/21/01 - Commented out 'record locking error' message.  Not true 90% of the time, just confuses people.
'Select Case iErrorNumber
'    Case 627: intLock% = True  ' dataset not updatable
'    Case 3006: intLock% = True ' DB is exclusively locked
'    Case 3008: intLock% = True ' Table exclusively locked
'    Case 3009: intLock% = True ' couldn't lock currently in use
'    Case 3027: intLock% = True  ' db is read only
'    Case 3028: intLock% = True ' system mda can't be opened
'    Case 3033: intLock% = True ' no permission for item
'    Case 3045: intLock% = True ' couldn't use item file already in use
'    Case 3046: intLock% = True  ' locked by another user
'    Case 3050: intLock% = True ' no share or insufficient locks
'    Case 3052: intLock% = True ' no share or insufficient locks
'    Case 3113: intLock% = True ' not updatable
'    Case 3158: intLock% = True  ' can't save record currently locked
'    Case 3164: intLock% = True ' can't update field
'    Case 3186: intLock% = True ' can't save locked by item 2
'    Case 3187: intLock% = True '  can't read locked bye item 2
'    Case 3188: intLock% = True  '  can't update on thi
'    Case 3189: intLock% = True '   exlucive lock
'    Case 3197: intLock% = True '   data changed operation stopped
'    Case 3202: intLock% = True '   couldn't save locked
'    Case 3211: intLock% = True  '  couldn't lock table - currently in use
'    Case 3212: intLock% = True '   couldn't lock table
'    Case 3218: intLock% = True '   couldn't update currently locked
'    Case 3260: intLock% = True '   couldn't update
'End Select
'If intLock Then
'    strMsg = "A record locking error occured." & Chr(10)
'    strMsg = strMsg & "This warning typically occurs in multi-user environments "
'    strMsg = strMsg & "when two people try to update/edit the same information. " & Chr(10)
'    strMsg = strMsg & "After you press 'OK' the screen you will likely be closed. " & Chr(10)
'    strMsg = strMsg & "This is not a system error - but a warning that your actions "
'    strMsg = strMsg & "were not entirely carried out. " & Chr(10)
'    strMsg = strMsg & "The specific Warning Number and message was. " & Chr(10)
'    strMsg = strMsg & "Warning # " & iErrorNumber & Chr(10)
'    strMsg = strMsg & "Message : " & Error$                    ' sub with Error$ strerrd$
'    strMsg = strMsg & Chr(10) & "Press Ok to continue. "
'    intMsgType = MB_OK + MB_ICONEXCLAMATION
'    MsgBox strMsg, intMsgType, "WARNING - Records Locked"
'    Err = 0             ' reset the error number - resets message
'    gintRollBack% = True
'    Exit Sub
'End If

'~~~~~~~ERROR 3000 RESERVED ERROR CODES FROM JET 2.0 JET ENGINE~~~~~~~~~~~~~~~~~~
If iErrorNumber = 3000 Then
    Screen.MousePointer = DEFAULT
    Msg = "ERROR DESCRIPTION  " & Chr(10)
    Msg = Msg & " - " & Error$(iErrorNumber) & Chr(10) & Chr(10)
    Msg = Msg & "ERROR # - " & CStr(iErrorNumber) & Chr(10)
    
    Msg = Msg & "JET 2.0 RESERVED ERROR CODE MESSAGE:" & Chr(10)
    Msg = Msg & Error$ & Chr(10)
    
    Msg = Msg & "FORM - " & FName$ & Chr(10)
    Msg = Msg & "SUB CALL - " & WDesc$ & Chr(10)
    Msg = Msg & "TABLE NAME - " & TabName$ & Chr(10)
    Msg = Msg & "DATA CALL - " & TCall$ & Chr(10)
    Msg = Msg & Chr(10) & Chr(10)
    Msg = Msg & "Function Cancelled" & Chr(10)
    Msg = Msg & "Please report this error to the info:HR Support desk."
    Title = "info:HR - ERROR IN APPLICATION"
    DgDef = MB_OK + MB_ICONSTOP + MB_DEFBUTTON2
    Response = MsgBox(Msg, DgDef, Title)    ' Get user response.
    Err = 0
    gintRollBack% = True
    Exit Sub
End If
If iErrorNumber = 3001 Then
    Screen.MousePointer = DEFAULT
    Msg = "Data type mismatch, or invalid character."
    Title = "info:HR - ERROR IN APPLICATION"
    DgDef = MB_OK + MB_ICONSTOP + MB_DEFBUTTON2
    Response = MsgBox(Msg, DgDef, Title)    ' Get user response.
    Err = 0
    gintRollBack% = False
    Exit Sub
End If
'Jaddy changed to fix the "Subscript out of range Attendance Report
If iErrorNumber = 9 Then
    Screen.MousePointer = DEFAULT
'    Msg = "Close previous screen before opening this one."
'    Title = "info:HR - ERROR IN APPLICATION"
'    DgDef = MB_OK + MB_ICONSTOP + MB_DEFBUTTON2
'    Response = MsgBox(Msg, DgDef, Title)    ' Get user response.
    Err = 0
    gintRollBack% = False
    Exit Sub
End If

'If iErrorNumber = -2147217900 Then
'    Screen.MousePointer = DEFAULT
'    Msg = "Can not insert duplicate key in the table."
'    Title = "info:HR - ERROR IN APPLICATION"
'    DgDef = MB_OK + MB_ICONSTOP + MB_DEFBUTTON2
'    Response = MsgBox(Msg, DgDef, Title)    ' Get user response.
'    Err = 0
'    gintRollBack% = False
'    Exit Sub
'End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
gintRollBack% = False
  
Screen.MousePointer = DEFAULT
Msg = "ERROR DESCRIPTION  " & Chr(10)
Msg = Msg & " - " & Error$(iErrorNumber) & Chr(10) & Chr(10)
Msg = Msg & "ERROR # - " & CStr(iErrorNumber) & Chr(10)
Msg = Msg & "FORM - " & FName$ & Chr(10)
Msg = Msg & "SUB CALL - " & WDesc$ & Chr(10)
Msg = Msg & "TABLE NAME - " & TabName$ & Chr(10)
Msg = Msg & "DATA CALL - " & TCall$ & Chr(10)
Msg = Msg & Chr(10) & Chr(10)
' dkostka - 03/21/01 - Added better error message for 'unrecognised database format'
If iErrorNumber = 3343 Then
    Msg = Msg & "Please run the info:HR Repair & Compress utility (found in the INFO HR program group) to repair your databases." & Chr$(10)
Else
    Msg = Msg & "Function Cancelled" & Chr(10)
    Msg = Msg & "Please report this error to the info:HR Support desk."
End If
Title = "info:HR - ERROR IN APPLICATION"
DgDef = MB_OK + MB_ICONSTOP + MB_DEFBUTTON2
Response = MsgBox(Msg, DgDef, Title)    ' Get user response.
' danielk - 06/21/2002 - Try to roll back any transactions that were in progress when the error happened
On Error Resume Next
gdbAdoIhr001.RollbackTrans
gdbAdoIhr001X.RollbackTrans
gdbAdoIhr001W.RollbackTrans
' danielk - 06/21/2002 - end
Err = 0
gintRollBack% = True
' dkostka - 03/21/01 - Added better error message for 'unrecognised database format'
If iErrorNumber = 3343 Then End
Exit Sub
End Sub

Function DaysBetween(txtfld1, txtfld2)
    Dim datfld1 As Variant, datfld2 As Variant
    datfld1 = CVDate(txtfld1)
    datfld2 = CVDate(txtfld2)
    DaysBetween = DateDiff("d", datfld1, datfld2)
End Function

Function Dept_Secure()
Dim NoDepts As Integer
Dim xSnap As New ADODB.Recordset, countr   As Integer
Dim DeptUn_Snap As New ADODB.Recordset
Dim SQLQ As String
Dim xDept, xUnion, xDiv, xSECTION, xAdminBy, xLoc, xRegion, xSupCode, xVadim2
Dim xInclEmp, xExclEmp

On Error GoTo Dept_Err

Dept_Secure = False

SQLQ = "Select HRPASDEP.* from HRPASDEP"
SQLQ = SQLQ & " where HRPASDEP.PD_USERID = '" & Replace(glbUserID, "'", "''") & "'"
xSnap.Open SQLQ, gdbAdoIhr001, adOpenStatic

glbSeleDept = " ("
Do Until xSnap.EOF
    xDept = xSnap("PD_DEPT")
    If xDept = "ALL" Then
        glbSeleDept = "( 1=1 OR "
        Exit Do
    Else
        glbSeleDept = glbSeleDept & " DF_NBR='" & xDept & "' or "
    End If
    xSnap.MoveNext
Loop
glbSeleDept = glbSeleDept & " 1=2 ) "
xSnap.Close


SQLQ = "Select HRPASDEP.PD_ORG from HRPASDEP"
SQLQ = SQLQ & " where HRPASDEP.PD_USERID = '" & Replace(glbUserID, "'", "''") & "'"
SQLQ = SQLQ & " Group by PD_ORG"

xSnap.Open SQLQ, gdbAdoIhr001, adOpenStatic

glbUnionForm = False
glbSeleUnion = " ("
Do Until xSnap.EOF
    xUnion = xSnap("PD_ORG")
    If xUnion = "-NON" Or xUnion = "-EXE" Then  'Hemu -EXE
        If xUnion = "-NON" Then glbNoNONE = True
        If xUnion = "-EXE" Then glbNoEXEC = True     'Hemu -EXE
        If xSnap.RecordCount = 1 Then
            glbSeleUnion = "( 1=1 or "
        ElseIf xSnap.RecordCount = 2 Then 'Hemu -EXE
            glbSeleUnion = "( 1=1 or "
        End If
    Else
        If IsNull(xUnion) Then
            glbSeleUnion = "( 1=1 or "
        Else
            If Len(xUnion) = 0 Then
                glbSeleUnion = "( 1=1 or "
            Else
'                If Left(xUnion, 1) = "-" Then ' for listewol
'                    glbSeleUnion = glbSeleUnion & " TB_KEY <>'" & Mid(xUnion, 2) & "'  or "
'                Else
                    If glbCompSerial = "S/N - 2288W" And Left(xUnion, 1) = "-" Then 'Musashi - Ticket #12690
                        glbSeleUnion = glbSeleUnion & " TB_KEY='" & Mid(xUnion, 2) & "' or "
                    Else
                        glbSeleUnion = glbSeleUnion & " TB_KEY='" & xUnion & "' or "
                    End If
'                End If
            End If
        End If
    
    End If
    xSnap.MoveNext
Loop
glbSeleUnion = glbSeleUnion & " 1=2 )"
xSnap.Close


SQLQ = "Select HRPASDEP.PD_DIV from HRPASDEP"
SQLQ = SQLQ & " where HRPASDEP.PD_USERID = '" & Replace(glbUserID, "'", "''") & "'"
SQLQ = SQLQ & " Group by PD_DIV"

xSnap.Open SQLQ, gdbAdoIhr001, adOpenStatic

glbSeleDiv = " ("
Do Until xSnap.EOF
    xDiv = xSnap("PD_DIV")
    If IsNull(xDiv) Then
        glbSeleDiv = "( 1=1 or "
        Exit Do
    Else
        If Len(xDiv) = 0 Then
            glbSeleDiv = "( 1=1 or "
            Exit Do
        Else
            glbSeleDiv = glbSeleDiv & " DIV='" & xDiv & "' or "
        End If
    End If
    xSnap.MoveNext
Loop
glbSeleDiv = glbSeleDiv & " 1=2 )"
xSnap.Close


'Ticket #18235
SQLQ = "Select HRPASDEP.PD_ADMINBY from HRPASDEP"
SQLQ = SQLQ & " where HRPASDEP.PD_USERID = '" & Replace(glbUserID, "'", "''") & "'"
SQLQ = SQLQ & " Group by PD_ADMINBY"

xSnap.Open SQLQ, gdbAdoIhr001, adOpenStatic

glbSeleAdminBy = " ("
Do Until xSnap.EOF
    xAdminBy = xSnap("PD_ADMINBY")
    If IsNull(xAdminBy) Then
        glbSeleAdminBy = "( 1=1 or "
        Exit Do
    Else
        If Len(xAdminBy) = 0 Then
            glbSeleAdminBy = "( 1=1 or "
            Exit Do
        Else
            glbSeleAdminBy = glbSeleAdminBy & " TB_KEY='" & xAdminBy & "' or "
        End If
    End If
    xSnap.MoveNext
Loop
glbSeleAdminBy = glbSeleAdminBy & " 1=2 )"
xSnap.Close

'Ticket #22682 - Release 8.0
SQLQ = "Select HRPASDEP.PD_LOC from HRPASDEP"
SQLQ = SQLQ & " where HRPASDEP.PD_USERID = '" & Replace(glbUserID, "'", "''") & "'"
SQLQ = SQLQ & " Group by PD_LOC"

xSnap.Open SQLQ, gdbAdoIhr001, adOpenStatic

glbSeleLoc = " ("
Do Until xSnap.EOF
    xLoc = xSnap("PD_LOC")
    If IsNull(xLoc) Then
        glbSeleLoc = "( 1=1 or "
        Exit Do
    Else
        If Len(xLoc) = 0 Then
            glbSeleLoc = "( 1=1 or "
            Exit Do
        Else
            glbSeleLoc = glbSeleLoc & " TB_KEY='" & xLoc & "' or "
        End If
    End If
    xSnap.MoveNext
Loop
glbSeleLoc = glbSeleLoc & " 1=2 )"
xSnap.Close

'Ticket #22682 - Release 8.0
SQLQ = "Select HRPASDEP.PD_REGION from HRPASDEP"
SQLQ = SQLQ & " where HRPASDEP.PD_USERID = '" & Replace(glbUserID, "'", "''") & "'"
SQLQ = SQLQ & " Group by PD_REGION"

xSnap.Open SQLQ, gdbAdoIhr001, adOpenStatic

glbSeleRegion = " ("
Do Until xSnap.EOF
    xRegion = xSnap("PD_REGION")
    If IsNull(xRegion) Then
        glbSeleRegion = "( 1=1 or "
        Exit Do
    Else
        If Len(xRegion) = 0 Then
            glbSeleRegion = "( 1=1 or "
            Exit Do
        Else
            glbSeleRegion = glbSeleRegion & " TB_KEY='" & xRegion & "' or "
        End If
    End If
    xSnap.MoveNext
Loop
glbSeleRegion = glbSeleRegion & " 1=2 )"
xSnap.Close

'Ticket #24161 - Samuel only - Release 8.0
If glbSamuel Then
    'Supervisor Code
    SQLQ = "Select HRPASDEP.PD_SUPCODE from HRPASDEP"
    SQLQ = SQLQ & " where HRPASDEP.PD_USERID = '" & Replace(glbUserID, "'", "''") & "'"
    SQLQ = SQLQ & " Group by PD_SUPCODE"
    
    xSnap.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    glbSeleSupCode = " ("
    Do Until xSnap.EOF
        xSupCode = xSnap("PD_SUPCODE")
        If IsNull(xSupCode) Then
            glbSeleSupCode = "( 1=1 or "
            Exit Do
        Else
            If Len(xSupCode) = 0 Then
                glbSeleSupCode = "( 1=1 or "
                Exit Do
            Else
                glbSeleSupCode = glbSeleSupCode & " TB_KEY='" & xSupCode & "' or "
            End If
        End If
        xSnap.MoveNext
    Loop
    glbSeleSupCode = glbSeleSupCode & " 1=2 )"
    xSnap.Close
    
    'Vadim Field 2
    SQLQ = "Select HRPASDEP.PD_VADIM2 from HRPASDEP"
    SQLQ = SQLQ & " where HRPASDEP.PD_USERID = '" & Replace(glbUserID, "'", "''") & "'"
    SQLQ = SQLQ & " Group by PD_VADIM2"
    
    xSnap.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    glbSeleVadim2 = " ("
    Do Until xSnap.EOF
        xVadim2 = xSnap("PD_VADIM2")
        If IsNull(xVadim2) Then
            glbSeleVadim2 = "( 1=1 or "
            Exit Do
        Else
            If Len(xVadim2) = 0 Then
                glbSeleVadim2 = "( 1=1 or "
                Exit Do
            Else
                glbSeleVadim2 = glbSeleVadim2 & " TB_KEY='" & xVadim2 & "' or "
            End If
        End If
        xSnap.MoveNext
    Loop
    glbSeleVadim2 = glbSeleVadim2 & " 1=2 )"
    xSnap.Close
End If

SQLQ = "Select HRPASDEP.PD_SECTION from HRPASDEP"
SQLQ = SQLQ & " where HRPASDEP.PD_USERID = '" & Replace(glbUserID, "'", "''") & "'"
SQLQ = SQLQ & " Group by PD_SECTION"

xSnap.Open SQLQ, gdbAdoIhr001, adOpenStatic

glbSeleSection = " ("
Do Until xSnap.EOF
    xSECTION = xSnap("PD_Section")
    If IsNull(xSECTION) Then
        glbSeleSection = "( 1=1 or "
        Exit Do
    Else
        If Len(xSECTION) = 0 Then
            glbSeleSection = "( 1=1 or "
            Exit Do
        Else
            glbSeleSection = glbSeleSection & " TB_KEY='" & xSECTION & "' or "
        End If
    End If
    xSnap.MoveNext
Loop
glbSeleSection = glbSeleSection & " 1=2 )"
xSnap.Close

glbSelePESection = Replace(glbSeleSection, "TB_KEY", "PE_SECTION")

If gSec_Emp_Based Then
    glbSeleDeptUn = " (ED_EMPNBR=" & glbEmpNbr & ") "
Else
    SQLQ = "Select HRPASDEP.* from HRPASDEP"
    SQLQ = SQLQ & " where HRPASDEP.PD_USERID = '" & Replace(glbUserID, "'", "''") & "'"
    SQLQ = SQLQ & " ORDER by PD_DEPT,PD_DIV,PD_SECTION,PD_ORG,PD_ADMINBY,PD_LOC,PD_REGION "
    
    'Ticket #24161 - Samuel only - Release 8.0
    If glbSamuel Then
        SQLQ = SQLQ & ",PD_SUPCODE,PD_VADIM2 "
    End If
    SQLQ = SQLQ & " DESC"
    
    DeptUn_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If DeptUn_Snap.EOF Then Exit Function
    
    glbSeleDeptUn = " ("
    
    Do Until DeptUn_Snap.EOF
        xUnion = DeptUn_Snap("PD_ORG")
        xDept = DeptUn_Snap("PD_DEPT")
        xDiv = DeptUn_Snap("PD_DIV")
        xSECTION = DeptUn_Snap("PD_SECTION")
        xAdminBy = DeptUn_Snap("PD_ADMINBY")
        xLoc = DeptUn_Snap("PD_LOC")
        xRegion = DeptUn_Snap("PD_REGION")
        
        'Ticket #24161 - Samuel only - Release 8.0
        If glbSamuel Then
            xSupCode = DeptUn_Snap("PD_SUPCODE")
            xVadim2 = DeptUn_Snap("PD_VADIM2")
        End If
        
        xInclEmp = DeptUn_Snap("PD_INCLEMPNBR")
        xExclEmp = DeptUn_Snap("PD_EXCLEMPNBR")
        
'Hemu (Ticket #21484)       '7.9 Enhancement
'        If Not IsNull(xInclEmp) And Len(xInclEmp) <> 0 Then
'            glbSeleDeptUn = glbSeleDeptUn & " ("
'        End If
        
        If xDept = "ALL" Then
            glbSeleDeptUn = glbSeleDeptUn & " (1=1 and "
        Else
            glbSeleDeptUn = glbSeleDeptUn & " (ED_DEPTNO='" & xDept & "' and "
        End If
        
        If xUnion = "-NON" Or xUnion = "-EXE" Then      'Hemu -EXE
            If InStr(glbSeleUnion, "1=1") = 0 Then
                glbSeleDeptUn = glbSeleDeptUn & " ED_ORG IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDOR' AND " & glbSeleUnion & ") and "
            End If
        Else
            If IsNull(xUnion) Then
                glbSeleDeptUn = glbSeleDeptUn & " 1=1 and "
            Else
                If Len(xUnion) = 0 Then
                    glbSeleDeptUn = glbSeleDeptUn & " 1=1 and "
                Else
                    If glbCompSerial = "S/N - 2288W" And Left(xUnion, 1) = "-" Then 'Musashi - Ticket #12690
                        glbSeleDeptUn = glbSeleDeptUn & " ED_ORG='" & Mid(xUnion, 2) & "' and "
                    ElseIf Left(xUnion, 1) = "-" Then   'Listowel
                        glbSeleDeptUn = Replace(glbSeleDeptUn, ") or  (", ") and  (")
                        glbSeleDeptUn = glbSeleDeptUn & " (ED_ORG<>'" & Mid(xUnion, 2, Len(xUnion)) & "' or ED_ORG IS NULL)  and "
                    Else
                        glbSeleDeptUn = glbSeleDeptUn & " ED_ORG='" & xUnion & "' and "
                    End If
                End If
            End If
        End If
        
        If IsNull(xDiv) Then
            glbSeleDeptUn = glbSeleDeptUn & " 1=1 and "
        Else
            If Len(xDiv) = 0 Then
                glbSeleDeptUn = glbSeleDeptUn & " 1=1 and "
            Else
                glbSeleDeptUn = glbSeleDeptUn & " ED_DIV='" & xDiv & "' and "
            End If
        End If
        
        'Ticket #18235
        If IsNull(xAdminBy) Then
'Hemu (Ticket #21484) If IsNull(xInclEmp) Or Len(xInclEmp) = 0 Then
                glbSeleDeptUn = glbSeleDeptUn & " 1=1 and "
'            Else
'                glbSeleDeptUn = glbSeleDeptUn & " 1=1 or "
'            End If
        Else
            If Len(xAdminBy) = 0 Then
'                If Len(xInclEmp) = 0 Or IsNull(xInclEmp) Then
                    glbSeleDeptUn = glbSeleDeptUn & " 1=1 and "
'                Else
'                    glbSeleDeptUn = glbSeleDeptUn & " 1=1 or "
'                End If
            Else
'                If IsNull(xInclEmp) Or Len(xInclEmp) = 0 Then
                    glbSeleDeptUn = glbSeleDeptUn & " ED_ADMINBY='" & xAdminBy & "' and "
'                Else
'                    glbSeleDeptUn = glbSeleDeptUn & " ED_ADMINBY='" & xAdminBy & "' or "
'                End If
            End If
        End If
        
        
        'Ticket #22682 - Release 8.0
        If IsNull(xLoc) Then
            glbSeleDeptUn = glbSeleDeptUn & " 1=1 and "
        Else
            If Len(xLoc) = 0 Then
                glbSeleDeptUn = glbSeleDeptUn & " 1=1 and "
            Else
                glbSeleDeptUn = glbSeleDeptUn & " ED_LOC='" & xLoc & "' and "
            End If
        End If
        
        'Ticket #22682 - Release 8.0
        If IsNull(xRegion) Then
            glbSeleDeptUn = glbSeleDeptUn & " 1=1 and "
        Else
            If Len(xRegion) = 0 Then
                glbSeleDeptUn = glbSeleDeptUn & " 1=1 and "
            Else
                glbSeleDeptUn = glbSeleDeptUn & " ED_REGION='" & xRegion & "' and "
            End If
        End If
        
        'Ticket #24161 - Samuel only - Release 8.0
        If glbSamuel Then
            'Supervisor Code
            If IsNull(xSupCode) Then
                glbSeleDeptUn = glbSeleDeptUn & " 1=1 and "
            Else
                If Len(xSupCode) = 0 Then
                    glbSeleDeptUn = glbSeleDeptUn & " 1=1 and "
                Else
                    glbSeleDeptUn = glbSeleDeptUn & " ED_SUPCODE='" & xSupCode & "' and "
                End If
            End If
            
            'Vadim Field 2
            If IsNull(xVadim2) Then
                glbSeleDeptUn = glbSeleDeptUn & " 1=1 and "
            Else
                If Len(xVadim2) = 0 Then
                    glbSeleDeptUn = glbSeleDeptUn & " 1=1 and "
                Else
                    glbSeleDeptUn = glbSeleDeptUn & " ED_VADIM2='" & xVadim2 & "' and "
                End If
            End If
        End If
        
'Hemu (Ticket #21484) '7.9 Enhancement
'        'Include Employee #s
'        If IsNull(xInclEmp) Then
'            'If IsNull(xExclEmp) Then
'                glbSeleDeptUn = glbSeleDeptUn & " 1=1 and "
'            'Else
'            '    glbSeleDeptUn = glbSeleDeptUn & " 1=1 or "
'            'End If
'        Else
'            If Len(xInclEmp) = 0 Then
'                'If Len(xExclEmp) = 0 Then
'                    glbSeleDeptUn = glbSeleDeptUn & " 1=1 and "
'                'Else
'                '    glbSeleDeptUn = glbSeleDeptUn & " 1=1 or "
'                'End If
'            Else
'                'If IsNull(xExclEmp) Or Len(xExclEmp) = 0 Then
'                    glbSeleDeptUn = glbSeleDeptUn & " ED_EMPNBR IN (" & getEmpnbr(xInclEmp) & ")) and "
'                'Else
'                '    glbSeleDeptUn = glbSeleDeptUn & " ED_EMPNBR IN (" & getEmpnbr(xInclEmp) & ") or "
'                'End If
'            End If
'        End If
        
        '7.9 Enhancement
        'Exclude Employee #s
        If IsNull(xExclEmp) Then
            glbSeleDeptUn = glbSeleDeptUn & " 1=1 and "
        Else
            If Len(xExclEmp) = 0 Then
                glbSeleDeptUn = glbSeleDeptUn & " 1=1 and "
            Else
                glbSeleDeptUn = glbSeleDeptUn & " (ED_EMPNBR NOT IN (" & getEmpnbr(xExclEmp) & ")) and "
            End If
        End If
        
        
        'Section
        If IsNull(xSECTION) Then
            'Ticket #21484
            If Len(xInclEmp) > 0 Then
                glbSeleDeptUn = glbSeleDeptUn & " 1=1 OR (ED_EMPNBR IN (" & getEmpnbr(xInclEmp) & "))) or "
            Else
                glbSeleDeptUn = glbSeleDeptUn & " 1=1) or "
            End If
        Else
            If Len(xSECTION) = 0 Then
                'Ticket #21484
                If Len(xInclEmp) > 0 Then
                    glbSeleDeptUn = glbSeleDeptUn & " 1=1 OR (ED_EMPNBR IN (" & getEmpnbr(xInclEmp) & "))) or "
                Else
                    glbSeleDeptUn = glbSeleDeptUn & " 1=1) or "
                End If
            Else
                'Ticket #21484
                If Len(xInclEmp) > 0 Then
                    glbSeleDeptUn = glbSeleDeptUn & " ED_SECTION='" & xSECTION & "' OR (ED_EMPNBR IN (" & getEmpnbr(xInclEmp) & "))) or "
                Else
                    glbSeleDeptUn = glbSeleDeptUn & " ED_SECTION='" & xSECTION & "') or "
                End If
            End If
        End If
        
        DeptUn_Snap.MoveNext
    Loop
    glbSeleDeptUn = glbSeleDeptUn & " 1=2 ) "
    
'    'Hemu - Ticket #21484
'    If Len(xInclEmp) > 0 Then
'        glbSeleDeptUn = glbSeleDeptUn & " OR (ED_EMPNBR IN (" & getEmpnbr(xInclEmp) & ")) "
'    End If
    
    DeptUn_Snap.Close
End If

Call Set_Div_List

Dept_Secure = True

Exit Function
Dept_Err:
If Err = 3704 Or Err = -2147217904 Then
    MsgBox "      The database has not been converted to the most recent release." & Chr(10) & Chr(10) & _
    "Please contact HR Systems Strategies Inc. Support Department for assistance."
    End
Else
    MsgBox "Error occured # " & Err & " on HRDEPT open - call support"
    If gintRollBack% = False Then
        Resume Next
    End If
End If
End Function

'Function EmpHisCalc(xType As Integer, XEMPNBR, NDept, NDiv, NStat, NPT, NOrg, NFte, NFteHr, nChgDate, Optional NBgroup)
'Dim SQLQ
'Dim rsTB As New ADODB.Recordset
'Dim rsTD As New ADODB.Recordset
'Dim rsTC As New ADODB.Recordset
'Dim rsTE As New ADODB.Recordset
'Dim xSalary, xSalcd, xFte, xFteHr
'
'On Error GoTo EMPHIS_ERR
'EmpHisCalc = False
'If Not IsNumeric(XEMPNBR) Then Exit Function
'If XEMPNBR = 0 Then Exit Function
'
'
'rsTB.Open "SELECT ED_EMPNBR,ED_DEPTNO,ED_DIV,ED_EMP,ED_PT,ED_ORG,ED_BENEFIT_GROUP FROM HREMP WHERE ED_EMPNBR = " & XEMPNBR, gdbAdoIhr001, adOpenKeyset
'If xType <> 2 And rsTB.EOF Then
'    rsTB.Close
'    Exit Function
'End If
'SQLQ = "SELECT JH_FTENUM,JH_FTEHRS FROM HR_JOB_HISTORY WHERE JH_EMPNBR = " & XEMPNBR & " and JH_CURRENT<>0"
'rsTC.Open SQLQ, gdbAdoIhr001, adOpenKeyset
'If Not rsTC.EOF Then
'    If IsNumeric(rsTC("JH_FTENUM")) Then xFte = rsTC("JH_FTENUM") Else xFte = 0
'    If IsNumeric(rsTC("JH_FTEHRS")) Then xFteHr = rsTC("JH_FTEHRS") Else xFteHr = 0
'Else
'    xFte = 0
'    xFteHr = 0
'End If
'rsTC.Close
'
'SQLQ = "SELECT SH_SALARY,SH_SALCD FROM HR_SALARY_HISTORY WHERE SH_EMPNBR = " & XEMPNBR & " and SH_CURRENT <>0"
'rsTD.Open SQLQ, gdbAdoIhr001, adOpenKeyset
'
'If Not rsTD.EOF Then
'    If IsNumeric(rsTD("SH_SALARY")) Then xSalary = rsTD("SH_SALARY") Else xSalary = 0
'    If Len(rsTD("SH_SALCD")) > 0 Then xSalcd = rsTD("SH_SALCD") Else xSalcd = " "
'Else
'    xSalary = 0
'    xSalcd = " "
'End If
'
'rsTD.Close
'
'
'rsTE.Open "HREMPHIS", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
'
'rsTE.AddNew
'rsTE("EE_COMPNO") = glbCompNo
'rsTE("EE_EMPNBR") = XEMPNBR
'rsTE("EE_SALARY") = xSalary
'rsTE("EE_SALCD") = xSalcd
'If Len(NDept) > 0 Then rsTE("EE_NEWDEPT") = NDept
'If Len(NDiv) > 0 Then rsTE("EE_NEWDIV") = NDiv
'If Len(NStat) > 0 Then rsTE("EE_NEWSTAT") = NStat
'If Len(NPT) > 0 Then rsTE("EE_NEWPT") = NPT
'If Len(NOrg) > 0 Then rsTE("EE_NEWORG") = NOrg
'
'If xType = 1 Then
'    If Len(NDept) > 0 Then rsTE("EE_OLDDEPT") = rsTB("ED_DEPTNO")
'    If Len(NDiv) > 0 Then rsTE("EE_OLDDIV") = rsTB("ED_DIV")
'    If Len(NStat) > 0 Then rsTE("EE_OLDSTAT") = rsTB("ED_EMP")
'    If Len(NPT) > 0 Then rsTE("EE_OLDPT") = rsTB("ED_PT")
'    If Len(NOrg) > 0 Then rsTE("EE_OLDORG") = rsTB("ED_ORG")
'End If
'If xType = 3 Then
'    rsTE("EE_OLDFTE") = IIf(NFte = "", 0, NFte)
'    rsTE("EE_OLDFTEHR") = IIf(NFteHr = "", 0, NFteHr)
'    rsTE("EE_NEWFTE") = xFte
'    rsTE("EE_NEWFTEHR") = xFteHr
'End If
'If Not IsMissing(NBgroup) Then
'    rsTE("EE_OLDBENEGROUP") = rsTB("ED_BENEFIT_GROUP")
'    rsTE("EE_NEWBENEGROUP") = NBgroup
'End If
'rsTE("EE_CHGDATE") = Format(nChgDate, "SHORT DATE")
'rsTE("EE_LDATE") = Date
'rsTE("EE_LUSER") = glbUserID
'rsTE("EE_LTIME") = Time$
'rsTE.Update
'rsTE.Close
'
'EmpHisCalc = True
'Exit Function
'
'EMPHIS_ERR:
'If Err = 13 Then
'    Err = 0
'    Resume Next
'End If
'
'glbFrmCaption$ = "EMPLOYEE HIS"
'glbErrNum& = Err
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EMPHISCALC", "EMPHIS", "Insert")
'Screen.MousePointer = DEFAULT
'If gintRollBack% = False Then
'    Resume Next
'End If
'
'End Function

Function EmpHisCalc(xType As Integer, xEmpnbr, NDept, NDiv, NStat, NPT, NOrg, NFte, NFteHr, nChgDate, Optional FieldName, Optional NCode, Optional nFromDate, Optional nTodate, Optional OldSmoker, Optional NeedOldVal = "Y", Optional xSaveOldVal)
Dim SQLQ
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim rsTD As New ADODB.Recordset
Dim rsTC As New ADODB.Recordset
Dim rsTE As New ADODB.Recordset
Dim xSalary, xSalCD, xFte, xFteHr
Dim xStrTmp As String
Dim xNewPos, xNewRep1

On Error GoTo EMPHIS_ERR

EmpHisCalc = False

If Not IsNumeric(xEmpnbr) Then Exit Function
If xEmpnbr = 0 Then Exit Function

'Ticket #21118 Franks 10/25/2011 add ED_SMOKER,ED_MSTAT
rsTB.Open "SELECT ED_EMPNBR,ED_DEPTNO,ED_DIV,ED_EMP,ED_PT,ED_ORG,ED_LOC,ED_REGION,ED_SECTION,ED_ADMINBY, ED_BENEFIT_GROUP,ED_SMOKER,ED_MSTAT FROM HREMP WHERE ED_EMPNBR = " & xEmpnbr, gdbAdoIhr001, adOpenKeyset
If xType <> 2 And rsTB.EOF Then
    rsTB.Close
    EmpHisCalc = True
    Exit Function
End If

SQLQ = "SELECT JH_FTENUM,JH_FTEHRS,JH_JOB,JH_REPTAU FROM HR_JOB_HISTORY WHERE JH_EMPNBR = " & xEmpnbr & " and JH_CURRENT<>0"
rsTC.Open SQLQ, gdbAdoIhr001, adOpenKeyset
If Not rsTC.EOF Then
    If IsNumeric(rsTC("JH_FTENUM")) Then xFte = rsTC("JH_FTENUM") Else xFte = 0
    If IsNumeric(rsTC("JH_FTEHRS")) Then xFteHr = rsTC("JH_FTEHRS") Else xFteHr = 0
    xNewPos = rsTC("JH_JOB") 'Ticket #27553 Franks 09/21/2015
    If IsNull(rsTC("JH_REPTAU")) Then xNewRep1 = "" Else xNewRep1 = rsTC("JH_REPTAU") 'Ticket #27553 Franks 09/21/2015
Else
    xFte = 0
    xFteHr = 0
    xNewPos = ""
    xNewRep1 = ""
End If
rsTC.Close

'Ticket #28048 - Jerry said the system should update when the new value is blank - changed from a value
'Ticket #27553 Franks - Rept. Authority 1 must be enterred
If xType = 8 Then
    'If Len(xNewRep1) = 0 Then
    '    EmpHisCalc = True   'Ticket #28048 - was giving EMPHIS Error because this function was not returning True.
    '    Exit Function
    'End If
End If


SQLQ = "SELECT SH_SALARY,SH_SALCD FROM HR_SALARY_HISTORY WHERE SH_EMPNBR = " & xEmpnbr & " and SH_CURRENT <>0"
rsTD.Open SQLQ, gdbAdoIhr001, adOpenKeyset
If Not rsTD.EOF Then
    If IsNumeric(rsTD("SH_SALARY")) Then xSalary = rsTD("SH_SALARY") Else xSalary = 0
    If Len(rsTD("SH_SALCD")) > 0 Then xSalCD = rsTD("SH_SALCD") Else xSalCD = " "
Else
    xSalary = 0
    xSalCD = " "
End If
rsTD.Close


rsTE.Open "HREMPHIS", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
rsTE.AddNew
rsTE("EE_COMPNO") = glbCompNo
rsTE("EE_EMPNBR") = xEmpnbr
rsTE("EE_SALARY") = xSalary
rsTE("EE_SALCD") = xSalCD
If Len(NDept) > 0 Then rsTE("EE_NEWDEPT") = NDept
If Len(NDiv) > 0 Then rsTE("EE_NEWDIV") = NDiv
If Len(NStat) > 0 Then rsTE("EE_NEWSTAT") = NStat
If Len(NPT) > 0 Then rsTE("EE_NEWPT") = NPT
If Len(NOrg) > 0 Then rsTE("EE_NEWORG") = NOrg

If xType = 1 Then
    If Len(NDept) > 0 Then rsTE("EE_OLDDEPT") = rsTB("ED_DEPTNO")
    If Len(NDiv) > 0 Then rsTE("EE_OLDDIV") = rsTB("ED_DIV")
    If Len(NStat) > 0 Then rsTE("EE_OLDSTAT") = rsTB("ED_EMP")
    If Len(NPT) > 0 Then rsTE("EE_OLDPT") = rsTB("ED_PT")
    If Len(NOrg) > 0 Then rsTE("EE_OLDORG") = rsTB("ED_ORG")
End If

If xType = 2 Then
    'If glbWFC Then 'Ticket #24317 Franks 09/16/2013
        If Not IsMissing(xSaveOldVal) Then
            If Len(NDept) > 0 Then rsTE("EE_OLDDEPT") = xSaveOldVal 'rsTB("ED_DEPTNO")
            If Len(NDiv) > 0 Then rsTE("EE_OLDDIV") = xSaveOldVal 'rsTB("ED_DIV")
            If Len(NStat) > 0 Then rsTE("EE_OLDSTAT") = xSaveOldVal
            If Len(NPT) > 0 Then rsTE("EE_OLDPT") = xSaveOldVal
            If Len(NOrg) > 0 Then rsTE("EE_OLDORG") = xSaveOldVal
        End If
    'End If
End If

If xType = 3 Then
    rsTE("EE_OLDFTE") = IIf(NFte = "", 0, NFte)
    rsTE("EE_OLDFTEHR") = IIf(NFteHr = "", 0, NFteHr)
    rsTE("EE_NEWFTE") = xFte
    rsTE("EE_NEWFTEHR") = xFteHr
End If

'Ticket #21118 Franks 10/25/2011 - begin
If xType = 5 Then
    If IsMissing(OldSmoker) Then
        If Not IsNull(rsTB("ED_SMOKER")) Then
            rsTE("EE_OLDSMOKER") = IIf(rsTB("ED_SMOKER"), "Yes", "No")
        End If
    Else
        rsTE("EE_OLDSMOKER") = OldSmoker 'Ticket #23491 Franks 04/02/2013
    End If
    rsTE("EE_NEWSMOKER") = NCode
End If

If xType = 6 Then
    rsTE("EE_OLDMSTAT") = getMSDesc(rsTB("ED_MSTAT"))
    xStrTmp = NCode
    rsTE("EE_NEWMSTAT") = getMSDesc(xStrTmp)
End If
'Ticket #21118 Franks 10/25/2011 - end

'Ticket #27553 Franks 09/21/2015 - begin
If xType = 7 Then
    'Ticket #29722 - For Multi Position, it is unable to retrieve the Position from the rsTD recordset hence updating with blank values
    'rsTE("EE_OLDPOSITION") = NCode
    'rsTE("EE_NEWPOSITION") = xNewPos
    rsTE("EE_OLDPOSITION") = xSaveOldVal
    rsTE("EE_NEWPOSITION") = NCode
End If
If xType = 8 Then
    'Ticket #29722 - For Multi Position, it is unable to retrieve the RA 1 from the rsTD recordset hence updating with blank values
    'rsTE("EE_OLDREPORT1") = NCode
    'rsTE("EE_NEWREPORT1") = xNewRep1
    rsTE("EE_OLDREPORT1") = xSaveOldVal
    rsTE("EE_NEWREPORT1") = NCode
End If
'Ticket #27553 Franks 09/21/2015 - end

If Not IsMissing(FieldName) And Not IsMissing(NCode) Then
    rsTE("EE_NEW" & FieldName) = NCode
    If Not rsTB.EOF Then
        If NeedOldVal = "Y" Then 'Ticket #23875 Franks 06/14/2013
            If FieldName = "BENEGROUP" Then
                rsTE("EE_OLD" & FieldName) = rsTB("ED_BENEFIT_GROUP")
            Else
                If FieldName = "DEPT" Then
                    rsTE("EE_OLD" & FieldName) = rsTB("ED_DEPTNO")
                Else
                    rsTE("EE_OLD" & FieldName) = rsTB("ED_" & FieldName)
                End If
            End If
        End If
        If FieldName = "REGION" Then 'Ticket #12708
            If Len(NCode) > 0 Then
                rsTE("EE_NEWREGIONDESC") = Left(GetTABLDesc("EDRG", NCode), 30)
            End If
            If NeedOldVal = "Y" Then 'Ticket #23875 Franks 06/14/2013
                If Not IsNull(rsTB("ED_" & FieldName)) Then
                    rsTE("EE_OLDREGIONDESC") = GetTABLDesc("EDRG", rsTB("ED_" & FieldName))
                End If
            End If
        End If
    End If
    
    If Not IsMissing(xSaveOldVal) Then 'Ticket #24317 Franks 09/17/2013
        If FieldName = "LOC" Then rsTE("EE_OLDLOC") = xSaveOldVal
        If FieldName = "ADMINBY" Then rsTE("EE_OLDADMINBY") = xSaveOldVal
        If FieldName = "REGION" Then rsTE("EE_OLDREGION") = xSaveOldVal
        If FieldName = "SECTION" Then rsTE("EE_OLDSECTION") = xSaveOldVal
        If FieldName = "LOC" Then rsTE("EE_OLDLOC") = xSaveOldVal
    End If
End If

If Len(NStat) > 0 Then
    'Ticket #24889 - The Change Date was being replaced from From Date to Change Date
    If Not IsMissing(nFromDate) Then
        rsTE("EE_CHGDATE") = Format(nFromDate, "SHORT DATE")
    Else
        rsTE("EE_CHGDATE") = Format(nChgDate, "SHORT DATE")
    End If
    If Not IsMissing(nTodate) Then
        'If nTodate <> "" And Not IsMissing(nTodate) Then
        If nTodate <> "" Then
            rsTE("EE_TODATE") = Format(nTodate, "SHORT DATE")
        End If
    End If
    'rsTE("EE_CHGDATE") = Format(nChgDate, "SHORT DATE")
Else
    rsTE("EE_CHGDATE") = Format(nChgDate, "SHORT DATE")
End If

rsTE("EE_LDATE") = Format(Now, "SHORT DATE")
rsTE("EE_LUSER") = glbUserID
rsTE("EE_LTIME") = Time$
rsTE.Update
rsTE.Close

EmpHisCalc = True

'Call DelIncorrectEmpHisRecs

Exit Function

EMPHIS_ERR:
If Err = 13 Then
    Err = 0
    Resume Next
End If

glbFrmCaption$ = "EMPLOYEE HIS"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EMPHISCALC", "EMPHIS", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function
 
Sub DelIncorrectEmpHisRecs()
Dim SQLQ As String

SQLQ = "SELECT * FROM HREMPHIS WHERE EE_OLDDEPT IS NULL "
SQLQ = SQLQ & " AND EE_NEWDEPT IS NULL"
SQLQ = SQLQ & " AND EE_OLDDIV IS NULL"
SQLQ = SQLQ & " AND EE_NEWDIV IS NULL"
SQLQ = SQLQ & " AND EE_OLDSTAT IS NULL"
SQLQ = SQLQ & " AND EE_NEWSTAT IS NULL"
SQLQ = SQLQ & " AND EE_OLDPT   IS NULL"
SQLQ = SQLQ & " AND EE_NEWPT   IS NULL"
SQLQ = SQLQ & " AND EE_OLDORG   IS NULL"
SQLQ = SQLQ & " AND EE_NEWORG   IS NULL"
SQLQ = SQLQ & " AND EE_OLDFTE   IS NULL"
SQLQ = SQLQ & " AND EE_NEWFTE   IS NULL"
SQLQ = SQLQ & " AND EE_OLDFTEHR   IS NULL"
SQLQ = SQLQ & " AND EE_NEWFTEHR   IS NULL"
SQLQ = SQLQ & " AND EE_OLDREGION   IS NULL"
SQLQ = SQLQ & " AND EE_NEWREGION   IS NULL"
SQLQ = SQLQ & " AND EE_OLDSECTION   IS NULL"
SQLQ = SQLQ & " AND EE_NEWSECTION   IS NULL"
SQLQ = SQLQ & " AND EE_OLDADMINBY   IS NULL"
SQLQ = SQLQ & " AND EE_NEWADMINBY   IS NULL"
SQLQ = SQLQ & " AND EE_OLDLOC   IS NULL"
SQLQ = SQLQ & " AND EE_NEWLOC   IS NULL"
SQLQ = SQLQ & " AND EE_OLDBENEGROUP   IS NULL"
SQLQ = SQLQ & " AND EE_NEWBENEGROUP   IS NULL"
SQLQ = SQLQ & " AND EE_TODATE   IS NULL"
SQLQ = SQLQ & " AND EE_HRSWOLD   IS NULL"
SQLQ = SQLQ & " AND EE_OLDGLNO   IS NULL"
SQLQ = SQLQ & " AND EE_NEWGLNO   IS NULL"
SQLQ = SQLQ & " AND EE_OLDREGIONDESC   IS NULL"
SQLQ = SQLQ & " AND EE_NEWREGIONDESC   IS NULL"

gdbAdoIhr001.Execute SQLQ

End Sub

Function IsValidDate(xDate, xDay, xMonth, xYear)
    Dim xNewDate
    
    If Not IsDate(xDate) Then
        'Invalid dates because of last date of the month
        If xDay = 31 And (xMonth = 4 Or xMonth = 6 Or xMonth = 9 Or xMonth = 11) Then
            xNewDate = Format(xMonth & "/" & "30" & "/" & xYear, "mm/dd/yyyy")
            If IsDate(xNewDate) Then IsValidDate = xNewDate Else IsValidDate = ""
        ElseIf xDay = 31 And xMonth = 2 Then
            xNewDate = Format(xMonth & "/" & "29" & "/" & xYear, "mm/dd/yyyy")
            If IsDate(xNewDate) Then
                IsValidDate = xNewDate
            Else
                xNewDate = Format(xMonth & "/" & "28" & "/" & xYear, "mm/dd/yyyy")
                If IsDate(xNewDate) Then
                    IsValidDate = xNewDate
                Else
                    IsValidDate = ""
                End If
            End If
        ElseIf xDay = 29 And xMonth = 2 Then
            xNewDate = Format(xMonth & "/" & "28" & "/" & xYear, "mm/dd/yyyy")
            If IsDate(xNewDate) Then IsValidDate = xNewDate Else IsValidDate = ""
        End If
    'Valid dates but not last date of the month
    ElseIf xDay = 30 And (xMonth = 1 Or xMonth = 3 Or xMonth = 5 Or xMonth = 7 Or xMonth = 8 Or xMonth = 10 Or xMonth = 12) Then
        xNewDate = Format(xMonth & "/" & "31" & "/" & xYear, "mm/dd/yyyy")
        If IsDate(xNewDate) Then IsValidDate = xNewDate Else IsValidDate = xDate
    ElseIf xDay = 28 And xMonth = 2 Then
        xNewDate = Format(xMonth & "/" & "29" & "/" & xYear, "mm/dd/yyyy")
        If IsDate(xNewDate) Then IsValidDate = xNewDate Else IsValidDate = xDate
    Else
        IsValidDate = xDate
    End If
End Function

Sub EntReCalcPeriod_Daily(WSQLQ, xType, Optional xFromDate, Optional xToDate, Optional xFromDateS, Optional xToDateS, Optional flgVacDates As Boolean)
Dim xlen, xxx, xx1, x
Dim rsTA As New ADODB.Recordset
Dim fglbWDate$, fglbWDateS$
Dim SQLQ
Dim nFrom, nTo
Dim RecCNT
Dim xORG, xLoc, xEMP, xPT, xEmpExcl
Dim rsVT As New ADODB.Recordset
Dim rsVTDate As New ADODB.Recordset

On Error GoTo ErrorHandler

MDIMain.panHelp(1).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(1).FloodPercent = 1
MDIMain.panHelp(1).FloodPercent = 3


Select Case glbEntOutStanding$
    Case "2": fglbWDate$ = "ED_DOH"
    Case "3": fglbWDate$ = "ED_SENDTE"
    Case "4": fglbWDate$ = "ED_LTHIRE"
    Case "5": fglbWDate$ = "ED_USRDAT1"
    Case "6": fglbWDate$ = "ED_UNION"
End Select
'Select Case glbEntOutStandingS$ ' sets field reference for basic 'which date'
'    Case "2": fglbWDateS$ = "ED_DOH"
'    Case "3": fglbWDateS$ = "ED_SENDTE"
'    Case "4": fglbWDateS$ = "ED_LTHIRE"
'    Case "5": fglbWDateS$ = "ED_USRDAT1"
'    Case "6": fglbWDateS$ = "ED_UNION"
'End Select

'Part 2 - Vacation - Begin
If xType = "VAC" Then
    '2.1 -- Set ED_EFDATE,ED_ETDATE in HREMP table - Begin
    MDIMain.panHelp(1).FloodPercent = 40
    If glbEntOutStanding$ = "1" Then 'Based on 1 (Entitlement Date)
        SQLQ = "SELECT VD_ORG,VD_EMP,VD_PT,VD_EMPEXCL "
        SQLQ = SQLQ & " FROM HRVACENTDAILY "
        If IsDate(xFromDate) And IsDate(xToDate) Then
            SQLQ = SQLQ & " WHERE VD_FRDATE=" & Date_SQL(xFromDate)
            SQLQ = SQLQ & " AND VD_TODATE=" & Date_SQL(xToDate)
        End If
        SQLQ = SQLQ & " GROUP BY VD_ORG,VD_EMP,VD_PT,VD_EMPEXCL "
        rsVT.Open SQLQ, gdbAdoIhr001, adOpenStatic
        RecCNT = rsVT.RecordCount
        
        Do Until rsVT.EOF
            'MDIMain.panHelp(0).FloodPercent = (I / RecCNT) * 50 + 10: I = I + 1
            xORG = rsVT("VD_ORG") & ""
            xEMP = rsVT("VD_EMP") & ""
            xPT = rsVT("VD_PT") & ""
            xEmpExcl = rsVT("VD_EMPEXCL") & ""
        
            If IsMissing(xFromDate) Or IsMissing(xToDate) Then
                SQLQ = "SELECT VD_FRDATE,VD_TODATE "
                SQLQ = SQLQ & " FROM HRVACENTDAILY "
                SQLQ = SQLQ & " WHERE (VD_ORG = '" & xORG & "' " & IIf(Len(xORG) = 0, " OR VD_ORG IS NULL ", "") & ")"
                SQLQ = SQLQ & " AND (VD_EMP = '" & xEMP & "' " & IIf(Len(xEMP) = 0, " OR VD_EMP IS NULL ", "") & ")"
                SQLQ = SQLQ & " AND (VD_PT = '" & xPT & "' " & IIf(Len(xPT) = 0, " OR VD_PT IS NULL ", "") & ")"
                SQLQ = SQLQ & " AND (VD_EMPEXCL = '" & xEmpExcl & "' " & IIf(Len(xEmpExcl) = 0, " OR VD_EMPEXCL IS NULL ", "") & ")"
                SQLQ = SQLQ & " ORDER BY VD_FRDATE,VD_TODATE"
                
                If rsVTDate.State <> 0 Then rsVTDate.Close
                rsVTDate.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                nFrom = rsVTDate("VD_FRDATE")
                nTo = rsVTDate("VD_TODATE")
            Else
                nFrom = xFromDate
                nTo = xToDate
            End If
            
            If IsDate(nFrom) And IsDate(nTo) Then
                SQLQ = "UPDATE HREMP "
                SQLQ = SQLQ & " SET ED_EFDATE=" & Date_SQL(nFrom)
                SQLQ = SQLQ & " , ED_ETDATE=" & Date_SQL(nTo)
                SQLQ = SQLQ & " WHERE " & WSQLQ
                If Len(xORG) > 0 Then SQLQ = SQLQ & " AND ED_ORG = '" & xORG & "'"
                If Len(xEMP) > 0 Then SQLQ = SQLQ & " AND ED_EMP = '" & xEMP & "'"
                If Len(xPT) > 0 Then SQLQ = SQLQ & " AND ED_PT = '" & xPT & "'"
                If Len(xEmpExcl) > 0 Then SQLQ = SQLQ & " AND ED_EMP NOT IN ('" & Replace(xEmpExcl, ",", "','") & "')"
                'Hemu - Ticket #12649
                'If glbCompSerial = "S/N - 2296W" Then 'Essex County Library
                '    SQLQ = SQLQ & " AND NOT (HREMP.ED_PT = 'PT' AND HREMP.ED_ORG = 'CUPE') "
                'End If
                gdbAdoIhr001.BeginTrans
                gdbAdoIhr001.Execute SQLQ
                gdbAdoIhr001.CommitTrans
                                
            End If
            rsVT.MoveNext
        Loop
    End If
        
    'VAC taken - Begin
    If glbOracle Then
        MDIMain.panHelp(1).FloodPercent = 85
        
        'Hemu - Ticket #11332 - This is because ED_VACT does not get replaced with 0 when there are no records
        'in the HR_Attendance matching the criteria
        SQLQ = " Update HREMP SET "
        SQLQ = SQLQ & " ED_VACT = 0"
        SQLQ = SQLQ & " WHERE " & WSQLQ
        'Hemu - Ticket #12649
        'If glbCompSerial = "S/N - 2296W" Then 'Essex County Library
        '    SQLQ = SQLQ & " AND NOT (HREMP.ED_PT = 'PT' AND HREMP.ED_ORG = 'CUPE') "
        'End If
        gdbAdoIhr001.Execute SQLQ
        
        SQLQ = " Update HREMP SET "
        SQLQ = SQLQ & " ED_VACT =(SELECT SUM(AD_HRS) FROM HR_ATTENDANCE"
        SQLQ = SQLQ & " Where ED_EMPNBR = AD_EMPNBR"
        SQLQ = SQLQ & " AND (AD_DOA>= ED_EFDATE) AND (AD_DOA<=ED_ETDATE) "
        SQLQ = SQLQ & " AND (AD_REASON Like 'VAC%') )"
        SQLQ = SQLQ & " WHERE ED_EMPNBR IN"
        SQLQ = SQLQ & " (SELECT AD_EMPNBR FROM HR_ATTENDANCE WHERE HR_ATTENDANCE.AD_EMPNBR=HREMP.ED_EMPNBR"
        SQLQ = SQLQ & " AND (AD_DOA >= ED_EFDATE) AND (AD_DOA<=ED_ETDATE)"
        SQLQ = SQLQ & " AND (AD_REASON Like 'VAC%') )"
        SQLQ = SQLQ & " AND " & WSQLQ
        'Hemu - Ticket #12649
        'If glbCompSerial = "S/N - 2296W" Then 'Essex County Library
        '    SQLQ = SQLQ & " AND NOT (HREMP.ED_PT = 'PT' AND HREMP.ED_ORG = 'CUPE') "
        'End If
        gdbAdoIhr001.Execute SQLQ
    ElseIf glbSQL Then
    
        'Hemu - Ticket #11332 - This is because ED_VACT does not get replaced with 0 when there are no records
        'in the HR_Attendance matching the criteria
        SQLQ = " Update HREMP SET "
        SQLQ = SQLQ & " ED_VACT = 0"
        SQLQ = SQLQ & " WHERE " & WSQLQ
        'Hemu - Ticket #12649
        'If glbCompSerial = "S/N - 2296W" Then 'Essex County Library
        '    SQLQ = SQLQ & " AND NOT (HREMP.ED_PT = 'PT' AND HREMP.ED_ORG = 'CUPE') "
        'End If
        gdbAdoIhr001.Execute SQLQ
    
        SQLQ = " Update HREMP SET "
        SQLQ = SQLQ & " ED_VACT =(SELECT SUM(AD_HRS) FROM HR_ATTENDANCE"
        SQLQ = SQLQ & " Where ED_EMPNBR = AD_EMPNBR"
        SQLQ = SQLQ & " AND AD_DOA BETWEEN ED_EFDATE AND ED_ETDATE"
        SQLQ = SQLQ & " AND AD_REASON Like 'VAC%')"
        SQLQ = SQLQ & " WHERE ED_EMPNBR IN"
        SQLQ = SQLQ & " (SELECT AD_EMPNBR FROM HR_ATTENDANCE INNER JOIN HREMP ON HR_ATTENDANCE.AD_EMPNBR=HREMP.ED_EMPNBR"
        SQLQ = SQLQ & " WHERE (AD_DOA BETWEEN ED_EFDATE AND ED_ETDATE)"
        SQLQ = SQLQ & " AND AD_REASON Like 'VAC%')"
        SQLQ = SQLQ & " AND " & WSQLQ
        'Hemu - Ticket #12649
        'If glbCompSerial = "S/N - 2296W" Then 'Essex County Library
        '    SQLQ = SQLQ & " AND NOT (HREMP.ED_PT = 'PT' AND HREMP.ED_ORG = 'CUPE') "
        'End If
        gdbAdoIhr001.Execute SQLQ

    Else
        'Added by Bryan Ticket #11236, need to reset vacation taken, if there are no attendance records it will skip the person, leaving the existing taken value
        SQLQ = "UPDATE HREMP SET ED_VACT=0 WHERE " & WSQLQ
        gdbAdoIhr001.Execute SQLQ
        'end bryan
        SQLQ = "SELECT ED_EMPNBR, Sum(AD_HRS) AS SumHRS"
        SQLQ = SQLQ & " FROM HREMP INNER JOIN HR_ATTENDANCE ON HREMP.ED_EMPNBR = HR_ATTENDANCE.AD_EMPNBR"
        SQLQ = SQLQ & " WHERE AD_DOA>=ED_EFDATE And AD_DOA<=ED_ETDATE AND LEFT(AD_REASON,3)='VAC' AND " & WSQLQ
        'Hemu - Ticket #12649
        'If glbCompSerial = "S/N - 2296W" Then 'Essex County Library
        '    SQLQ = SQLQ & " AND NOT (HREMP.ED_PT = 'PT' AND HREMP.ED_ORG = 'CUPE') "
        'End If
        SQLQ = SQLQ & " GROUP BY ED_EMPNBR"
        rsTA.Open SQLQ, gdbAdoIhr001, adOpenKeyset
        Do Until rsTA.EOF
            gdbAdoIhr001.Execute "UPDATE HREMP SET ED_VACT=" & rsTA("SUMHRS") & " WHERE ED_EMPNBR=" & rsTA("ED_EMPNBR")
            rsTA.MoveNext
        Loop
        rsTA.Close
    End If
End If
'Part 2 - Vacation - End

''Part 3 - Sick - Begin
'If xType = "SICK" Then
'    '3.1 -- Set ED_EFDATES,ED_ETDATES in HREMP table - Begin
'    MDIMain.panHelp(1).FloodPercent = 60
'    If glbEntOutStandingS$ = "1" Then 'Based on 1 (Entitlement Date)
'        SQLQ = "SELECT VE_DIV,VE_DEPT,VE_ORG,VE_LOC,VE_EMP,VE_PT,VE_GRPCD,VE_SECTION "
'        SQLQ = SQLQ & " FROM HRSICKENT "
'        If IsDate(xFromDateS) And IsDate(xToDateS) Then
'            SQLQ = SQLQ & " WHERE VE_FRDATE=" & Date_SQL(xFromDateS)
'            SQLQ = SQLQ & " AND VE_TODATE=" & Date_SQL(xToDateS)
'        End If
'        SQLQ = SQLQ & " GROUP BY VE_DIV,VE_DEPT,VE_ORG,VE_LOC,VE_EMP,VE_PT,VE_GRPCD,VE_SECTION "
'
'        If rsVT.State <> 0 Then rsVT.Close
'        rsVT.Open SQLQ, gdbAdoIhr001, adOpenStatic
'        RecCNT = rsVT.RecordCount
'
'        Do Until rsVT.EOF
'            'MDIMain.panHelp(0).FloodPercent = (I / RecCNT) * 50 + 10: I = I + 1
'            xDiv = rsVT("VE_DIV") & ""
'            xDept = rsVT("VE_DEPT") & ""
'            xORG = rsVT("VE_ORG") & ""
'            xLoc = rsVT("VE_LOC") & ""
'            xEMP = rsVT("VE_EMP") & ""
'            xPT = rsVT("VE_PT") & ""
'            xGRPCD = rsVT("VE_GRPCD") & ""
'            xSec = rsVT("VE_SECTION") & ""
'
'            If IsMissing(xFromDateS) Or IsMissing(xToDateS) Then
'                SQLQ = "SELECT VE_FRDATE,VE_TODATE "
'                SQLQ = SQLQ & " FROM HRSICKENT "
'                SQLQ = SQLQ & " WHERE (VE_DIV = '" & xDiv & "' " & IIf(Len(xDiv) = 0, " OR VE_DIV IS NULL ", "") & ")"
'                SQLQ = SQLQ & " AND (VE_DEPT = '" & xDept & "' " & IIf(Len(xDept) = 0, " OR VE_DEPT IS NULL ", "") & ")"
'                SQLQ = SQLQ & " AND (VE_ORG = '" & xORG & "' " & IIf(Len(xORG) = 0, " OR VE_ORG IS NULL ", "") & ")"
'                SQLQ = SQLQ & " AND (VE_LOC = '" & xLoc & "' " & IIf(Len(xLoc) = 0, " OR VE_LOC IS NULL ", "") & ")"
'                SQLQ = SQLQ & " AND (VE_EMP = '" & xEMP & "' " & IIf(Len(xEMP) = 0, " OR VE_EMP IS NULL ", "") & ")"
'                SQLQ = SQLQ & " AND (VE_PT = '" & xPT & "' " & IIf(Len(xPT) = 0, " OR VE_PT IS NULL ", "") & ")"
'                SQLQ = SQLQ & " AND (VE_GRPCD = '" & xGRPCD & "' " & IIf(Len(xGRPCD) = 0, " OR VE_GRPCD IS NULL ", "") & ")"
'                SQLQ = SQLQ & " AND (VE_SECTION = '" & xSec & "' " & IIf(Len(xSec) = 0, " OR VE_SECTION IS NULL ", "") & ")"
'                SQLQ = SQLQ & " ORDER BY VE_FRDATE,VE_TODATE"
'
'                If rsVTDate.State <> 0 Then rsVTDate.Close
'                rsVTDate.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
'                nFrom = rsVTDate("VE_FRDATE")
'                nTo = rsVTDate("VE_TODATE")
'            Else
'                nFrom = xFromDateS
'                nTo = xToDateS
'            End If
'            If IsDate(nFrom) And IsDate(nTo) Then
'                SQLQ = "UPDATE HREMP "
'                SQLQ = SQLQ & " SET ED_EFDATES=" & Date_SQL(nFrom)
'                SQLQ = SQLQ & " , ED_ETDATES=" & Date_SQL(nTo)
'                SQLQ = SQLQ & " WHERE " & WSQLQ
'                If Len(xDiv) > 0 Then SQLQ = SQLQ & " AND ED_DIV = '" & xDiv & "'"
'                If Len(xDept) > 0 Then SQLQ = SQLQ & " AND ED_DEPTNO = '" & xDept & "'"
'                If Len(xORG) > 0 Then SQLQ = SQLQ & " AND ED_ORG = '" & xORG & "'"
'                If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #18235
'                    If Len(xLoc) > 0 Then SQLQ = SQLQ & " AND ED_VADIM2 = '" & xLoc & "'"
'                Else
'                    If Len(xLoc) > 0 Then SQLQ = SQLQ & " AND ED_LOC = '" & xLoc & "'"
'                End If
'                If Len(xEMP) > 0 Then SQLQ = SQLQ & " AND ED_EMP = '" & xEMP & "'"
'                If Len(xPT) > 0 Then SQLQ = SQLQ & " AND ED_PT = '" & xPT & "'"
'                If Len(xGRPCD) > 0 Then SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_JOB IN (SELECT JB_CODE FROM HRJOB WHERE JB_GRPCD = '" & xGRPCD & "'))"
'                If glbLinamar Then  'Hemu - Ticket #12494 - Section is actually Salary Distribution field for them on Entitl. Mstr
'                    If Len(xSec) > 0 Then SQLQ = SQLQ & " AND ED_SALDIST = '" & xSec & "' "
'                Else
'                    If Len(xSec) > 0 Then SQLQ = SQLQ & " AND ED_SECTION = '" & xSec & "'"
'                End If
'
'                gdbAdoIhr001.BeginTrans
'                gdbAdoIhr001.Execute SQLQ
'                gdbAdoIhr001.CommitTrans
'            End If
'            rsVT.MoveNext
'        Loop
'    End If
'
'    'SICK taken - Begin
'    If glbOracle Then
'        SQLQ = " Update HREMP SET "
'        SQLQ = SQLQ & " HREMP.ED_SICKT =(SELECT SUM(AD_HRS) FROM HR_ATTENDANCE"
'        SQLQ = SQLQ & " Where ED_EMPNBR = AD_EMPNBR"
'        SQLQ = SQLQ & " AND (AD_DOA>= ED_EFDATES) AND (AD_DOA<=ED_ETDATES )"
'        SQLQ = SQLQ & " AND (AD_REASON Like 'SIC%') )"
'        SQLQ = SQLQ & " WHERE ED_EMPNBR IN"
'        SQLQ = SQLQ & " (SELECT AD_EMPNBR FROM HR_ATTENDANCE WHERE HR_ATTENDANCE.AD_EMPNBR=HREMP.ED_EMPNBR"
'        SQLQ = SQLQ & " AND (AD_DOA>= ED_EFDATES) AND (AD_DOA<=ED_ETDATES)"
'        SQLQ = SQLQ & " AND (AD_REASON Like 'SIC%') )"
'        SQLQ = SQLQ & " AND " & WSQLQ
'        gdbAdoIhr001.Execute SQLQ
'
'    ElseIf glbSQL Then
'        SQLQ = " Update HREMP SET "
'        SQLQ = SQLQ & " HREMP.ED_SICKT =(SELECT SUM(AD_HRS) FROM HR_ATTENDANCE"
'        SQLQ = SQLQ & " Where ED_EMPNBR = AD_EMPNBR"
'        SQLQ = SQLQ & " AND AD_DOA BETWEEN ED_EFDATES AND ED_ETDATES"
'        SQLQ = SQLQ & " AND AD_REASON Like 'SIC%')"
'        SQLQ = SQLQ & " WHERE ED_EMPNBR IN"
'        SQLQ = SQLQ & " (SELECT AD_EMPNBR FROM HR_ATTENDANCE INNER JOIN HREMP ON HR_ATTENDANCE.AD_EMPNBR=HREMP.ED_EMPNBR"
'        SQLQ = SQLQ & " WHERE (AD_DOA BETWEEN ED_EFDATES AND ED_ETDATES)"
'        SQLQ = SQLQ & " AND AD_REASON Like 'SIC%')"
'        SQLQ = SQLQ & " AND " & WSQLQ
'        gdbAdoIhr001.Execute SQLQ
'    Else
'        'Added by Bryan Ticket #11236, need to reset vacation taken, if there are no attendance records it will skip the person, leaving the existing taken value
'        SQLQ = "UPDATE HREMP SET ED_SICKT=0 WHERE " & WSQLQ
'        gdbAdoIhr001.Execute SQLQ
'        SQLQ = "SELECT ED_EMPNBR, Sum(AD_HRS) AS SumHRS"
'        SQLQ = SQLQ & " FROM HREMP INNER JOIN HR_ATTENDANCE ON HREMP.ED_EMPNBR = HR_ATTENDANCE.AD_EMPNBR"
'        SQLQ = SQLQ & " WHERE AD_DOA>=ED_EFDATES And AD_DOA<=ED_ETDATES AND LEFT(AD_REASON,3)='SIC' AND " & WSQLQ
'        SQLQ = SQLQ & " GROUP BY ED_EMPNBR "
'        rsTA.Open SQLQ, gdbAdoIhr001, adOpenDynamic
'        Do Until rsTA.EOF
'            gdbAdoIhr001.Execute "UPDATE HREMP SET ED_SICKT=" & rsTA("SUMHRS") & " WHERE ED_EMPNBR=" & rsTA("ED_EMPNBR")
'            rsTA.MoveNext
'        Loop
'        rsTA.Close
'    End If
'End If
''Part 3 - Sick - End

glbENTScreen = True

MDIMain.panHelp(1).FloodPercent = 100
MDIMain.panHelp(1).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""

Exit Sub


ErrorHandler:
glbFrmCaption$ = "Daily Entitlement Recalculation"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EntReCalcPeriod_Daily", "", CStr(SQLQ))
If gintRollBack% = False Then
    Resume Next
End If
End Sub


Sub EntReCalcPeriod(WSQLQ, xType, Optional xFromDate, Optional xToDate, Optional xFromDateS, Optional xToDateS, Optional flgVacDates As Boolean, Optional strEntType As String)
Dim xlen, xxx, xx1, x
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim rsT_PARCO As New ADODB.Recordset
Dim rsEmpVacB As New ADODB.Recordset 'For County of Brant
Dim rsEmpVacE As New ADODB.Recordset 'For County of Brant
Dim rsEmpBack As New ADODB.Recordset   'For County of Brant
Dim rsTC As New ADODB.Recordset
Dim rsTD As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset 'For Multi Vacation Entitlement Periods, glbEntOutStanding$ = "1"
Dim XCNTIND
Dim EmpCNT, xEmpnbr, I
Dim xWDate, xWDateS
Dim fglbWDate$, fglbWDateS$
Dim SQLQ, SQLQ1, SQLQ2
Dim nFrom, nTo

Dim RecCNT
Dim xDiv, xDept, xORG, xLoc, xEMP, xPT, xGRPCD, xSec
Dim rsVT As New ADODB.Recordset
Dim rsVTDate As New ADODB.Recordset

On Error GoTo ErrorHandler

MDIMain.panHelp(1).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(1).FloodPercent = 1
MDIMain.panHelp(1).FloodPercent = 3


Select Case glbEntOutStanding$
    Case "2": fglbWDate$ = "ED_DOH"
    Case "3": fglbWDate$ = "ED_SENDTE"
    Case "4": fglbWDate$ = "ED_LTHIRE"
    Case "5": fglbWDate$ = "ED_USRDAT1"
    Case "6": fglbWDate$ = "ED_UNION"
End Select
Select Case glbEntOutStandingS$ ' sets field reference for basic 'which date'
    Case "2": fglbWDateS$ = "ED_DOH"
    Case "3": fglbWDateS$ = "ED_SENDTE"
    Case "4": fglbWDateS$ = "ED_LTHIRE"
    Case "5": fglbWDateS$ = "ED_USRDAT1"
    Case "6": fglbWDateS$ = "ED_UNION"
End Select

'Part 2 - Vacation - Begin
If xType = "VAC" Then
    '2.1 -- Set ED_EFDATE,ED_ETDATE in HREMP table - Begin
    MDIMain.panHelp(1).FloodPercent = 40
    If glbEntOutStanding$ = "1" Then 'Based on 1 (Entitlement Date)
        If Not IsMissing(strEntType) Then
            If strEntType = "HOURSBASED" Then
                SQLQ = "SELECT VH_DIV as VE_DIV,VH_DEPT as VE_DEPT,VH_ORG as VE_ORG,VH_LOC as VE_LOC,VH_EMP as VE_EMP,VH_PT as VE_PT,VH_GRPCD as VE_GRPCD,VH_SECTION as VE_SECTION "
                SQLQ = SQLQ & " FROM HRHRSVACENT "
                If IsDate(xFromDate) And IsDate(xToDate) Then
                    SQLQ = SQLQ & " WHERE VH_FRDATE=" & Date_SQL(xFromDate)
                    SQLQ = SQLQ & " AND VH_TODATE=" & Date_SQL(xToDate)
                End If
                SQLQ = SQLQ & " GROUP BY VH_DIV,VH_DEPT,VH_ORG,VH_LOC,VH_EMP,VH_PT,VH_GRPCD,VH_SECTION "
            Else
                SQLQ = "SELECT VE_DIV,VE_DEPT,VE_ORG,VE_LOC,VE_EMP,VE_PT,VE_GRPCD,VE_SECTION "
                If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #13979
                    SQLQ = SQLQ & " FROM HRANNVACENT "
                Else
                    SQLQ = SQLQ & " FROM HRVACENT "
                End If
                If IsDate(xFromDate) And IsDate(xToDate) Then
                    SQLQ = SQLQ & " WHERE VE_FRDATE=" & Date_SQL(xFromDate)
                    SQLQ = SQLQ & " AND VE_TODATE=" & Date_SQL(xToDate)
                End If
                SQLQ = SQLQ & " GROUP BY VE_DIV,VE_DEPT,VE_ORG,VE_LOC,VE_EMP,VE_PT,VE_GRPCD,VE_SECTION "
            End If
        End If
        rsVT.Open SQLQ, gdbAdoIhr001, adOpenStatic
        RecCNT = rsVT.RecordCount
        
        Do Until rsVT.EOF
            'MDIMain.panHelp(0).FloodPercent = (I / RecCNT) * 50 + 10: I = I + 1
            xDiv = rsVT("VE_DIV") & ""
            xDept = rsVT("VE_DEPT") & ""
            xORG = rsVT("VE_ORG") & ""
            xLoc = rsVT("VE_LOC") & ""
            xEMP = rsVT("VE_EMP") & ""
            xPT = rsVT("VE_PT") & ""
            xGRPCD = rsVT("VE_GRPCD") & ""
            xSec = rsVT("VE_SECTION") & ""
        
            If IsMissing(xFromDate) Or IsMissing(xToDate) Then
                If Not IsMissing(strEntType) Then
                    If strEntType = "HOURSBASED" Then
                        SQLQ = "SELECT VH_FRDATE as VE_FRDATE,VH_TODATE as VE_TODATE "
                        SQLQ = SQLQ & " FROM HRHRSVACENT "
                        SQLQ = SQLQ & " WHERE (VH_DIV = '" & xDiv & "' " & IIf(Len(xDiv) = 0, " OR VH_DIV IS NULL ", "") & ")"
                        SQLQ = SQLQ & " AND (VH_DEPT = '" & xDept & "' " & IIf(Len(xDept) = 0, " OR VH_DEPT IS NULL ", "") & ")"
                        SQLQ = SQLQ & " AND (VH_ORG = '" & xORG & "' " & IIf(Len(xORG) = 0, " OR VH_ORG IS NULL ", "") & ")"
                        SQLQ = SQLQ & " AND (VH_LOC = '" & xLoc & "' " & IIf(Len(xLoc) = 0, " OR VH_LOC IS NULL ", "") & ")"
                        SQLQ = SQLQ & " AND (VH_EMP = '" & xEMP & "' " & IIf(Len(xEMP) = 0, " OR VH_EMP IS NULL ", "") & ")"
                        SQLQ = SQLQ & " AND (VH_PT = '" & xPT & "' " & IIf(Len(xPT) = 0, " OR VH_PT IS NULL ", "") & ")"
                        SQLQ = SQLQ & " AND (VH_GRPCD = '" & xGRPCD & "' " & IIf(Len(xGRPCD) = 0, " OR VH_GRPCD IS NULL ", "") & ")"
                        SQLQ = SQLQ & " AND (VH_SECTION = '" & xSec & "' " & IIf(Len(xSec) = 0, " OR VH_SECTION IS NULL ", "") & ")"
                        SQLQ = SQLQ & " ORDER BY VH_FRDATE,VH_TODATE"
                    Else
                        SQLQ = "SELECT VE_FRDATE,VE_TODATE "
                        If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #13979
                            SQLQ = SQLQ & " FROM HRANNVACENT "
                        Else
                            SQLQ = SQLQ & " FROM HRVACENT "
                        End If
                        SQLQ = SQLQ & " WHERE (VE_DIV = '" & xDiv & "' " & IIf(Len(xDiv) = 0, " OR VE_DIV IS NULL ", "") & ")"
                        SQLQ = SQLQ & " AND (VE_DEPT = '" & xDept & "' " & IIf(Len(xDept) = 0, " OR VE_DEPT IS NULL ", "") & ")"
                        SQLQ = SQLQ & " AND (VE_ORG = '" & xORG & "' " & IIf(Len(xORG) = 0, " OR VE_ORG IS NULL ", "") & ")"
                        SQLQ = SQLQ & " AND (VE_LOC = '" & xLoc & "' " & IIf(Len(xLoc) = 0, " OR VE_LOC IS NULL ", "") & ")"
                        SQLQ = SQLQ & " AND (VE_EMP = '" & xEMP & "' " & IIf(Len(xEMP) = 0, " OR VE_EMP IS NULL ", "") & ")"
                        SQLQ = SQLQ & " AND (VE_PT = '" & xPT & "' " & IIf(Len(xPT) = 0, " OR VE_PT IS NULL ", "") & ")"
                        SQLQ = SQLQ & " AND (VE_GRPCD = '" & xGRPCD & "' " & IIf(Len(xGRPCD) = 0, " OR VE_GRPCD IS NULL ", "") & ")"
                        SQLQ = SQLQ & " AND (VE_SECTION = '" & xSec & "' " & IIf(Len(xSec) = 0, " OR VE_SECTION IS NULL ", "") & ")"
                        SQLQ = SQLQ & " ORDER BY VE_FRDATE,VE_TODATE"
                    End If
                Else
                    SQLQ = "SELECT VE_FRDATE,VE_TODATE "
                    If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #13979
                        SQLQ = SQLQ & " FROM HRANNVACENT "
                    Else
                        SQLQ = SQLQ & " FROM HRVACENT "
                    End If
                    SQLQ = SQLQ & " WHERE (VE_DIV = '" & xDiv & "' " & IIf(Len(xDiv) = 0, " OR VE_DIV IS NULL ", "") & ")"
                    SQLQ = SQLQ & " AND (VE_DEPT = '" & xDept & "' " & IIf(Len(xDept) = 0, " OR VE_DEPT IS NULL ", "") & ")"
                    SQLQ = SQLQ & " AND (VE_ORG = '" & xORG & "' " & IIf(Len(xORG) = 0, " OR VE_ORG IS NULL ", "") & ")"
                    SQLQ = SQLQ & " AND (VE_LOC = '" & xLoc & "' " & IIf(Len(xLoc) = 0, " OR VE_LOC IS NULL ", "") & ")"
                    SQLQ = SQLQ & " AND (VE_EMP = '" & xEMP & "' " & IIf(Len(xEMP) = 0, " OR VE_EMP IS NULL ", "") & ")"
                    SQLQ = SQLQ & " AND (VE_PT = '" & xPT & "' " & IIf(Len(xPT) = 0, " OR VE_PT IS NULL ", "") & ")"
                    SQLQ = SQLQ & " AND (VE_GRPCD = '" & xGRPCD & "' " & IIf(Len(xGRPCD) = 0, " OR VE_GRPCD IS NULL ", "") & ")"
                    SQLQ = SQLQ & " AND (VE_SECTION = '" & xSec & "' " & IIf(Len(xSec) = 0, " OR VE_SECTION IS NULL ", "") & ")"
                    SQLQ = SQLQ & " ORDER BY VE_FRDATE,VE_TODATE"
                End If
                If rsVTDate.State <> 0 Then rsVTDate.Close
                rsVTDate.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                nFrom = rsVTDate("VE_FRDATE")
                nTo = rsVTDate("VE_TODATE")
            Else
                nFrom = xFromDate
                nTo = xToDate
            End If
            If IsDate(nFrom) And IsDate(nTo) Then
                SQLQ = "UPDATE HREMP "
                SQLQ = SQLQ & " SET ED_EFDATE=" & Date_SQL(nFrom)
                SQLQ = SQLQ & " , ED_ETDATE=" & Date_SQL(nTo)
                SQLQ = SQLQ & " WHERE " & WSQLQ
                If Len(xDiv) > 0 Then SQLQ = SQLQ & " AND ED_DIV = '" & xDiv & "'"
                If Len(xDept) > 0 Then SQLQ = SQLQ & " AND ED_DEPTNO = '" & xDept & "'"
                If Len(xORG) > 0 Then SQLQ = SQLQ & " AND ED_ORG = '" & xORG & "'"
                If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #18235
                    If Len(xLoc) > 0 Then SQLQ = SQLQ & " AND ED_VADIM1 = '" & xLoc & "'"
                Else
                    If Len(xLoc) > 0 Then SQLQ = SQLQ & " AND ED_LOC = '" & xLoc & "'"
                End If
                If Len(xEMP) > 0 Then SQLQ = SQLQ & " AND ED_EMP = '" & xEMP & "'"
                If Len(xPT) > 0 Then SQLQ = SQLQ & " AND ED_PT = '" & xPT & "'"
                If Len(xGRPCD) > 0 Then SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_JOB IN (SELECT JB_CODE FROM HRJOB WHERE JB_GRPCD = '" & xGRPCD & "'))"
                If glbLinamar Then  'Hemu - Ticket #12494 - Section is actually Salary Distribution field for them on Entitl. Mstr
                    If Len(xSec) > 0 Then SQLQ = SQLQ & " AND ED_SALDIST = '" & xSec & "' "
                Else
                    If Len(xSec) > 0 Then SQLQ = SQLQ & " AND ED_SECTION = '" & xSec & "'"
                End If
                'Hemu - Ticket #12649
                'If glbCompSerial = "S/N - 2296W" Then 'Essex County Library
                '    SQLQ = SQLQ & " AND NOT (HREMP.ED_PT = 'PT' AND HREMP.ED_ORG = 'CUPE') "
                'End If
                gdbAdoIhr001.BeginTrans
                gdbAdoIhr001.Execute SQLQ
                gdbAdoIhr001.CommitTrans
            End If
            rsVT.MoveNext
        Loop
    End If
    
    If glbEntOutStanding$ <> "1" Then
        xWDate = IIf(fglbWDate$ = "", "ED_DOH", fglbWDate$)
        
        'Compute Vacation FROM Date
        If glbOracle Then
            SQLQ = "UPDATE HREMP SET "
            SQLQ = SQLQ & "ED_EFDATE="
            SQLQ = SQLQ & "(CASE WHEN " & xWDate & " IS NULL THEN NULL ELSE "
            SQLQ = SQLQ & " ADD_MONTHS(" & xWDate & ","
            SQLQ = SQLQ & "(CASE WHEN ADD_MONTHS(" & xWDate & ",12)>SYSDATE THEN 0 ELSE "
            SQLQ = SQLQ & " (TO_NUMBER(SYSDATE,'YYYY')-TO_NUMBER(" & xWDate & ",'YYYY'))*12 END)"
            SQLQ = SQLQ & " ) END) "
    
        ElseIf glbSQL Then
            'County of Essex - Ticket #12676
            If glbCompSerial = "S/N - 2192W" Then
                If Not IsMissing(flgVacDates) Then
                    If flgVacDates = True Then
                        SQLQ = "UPDATE HREMP SET "
                        SQLQ = SQLQ & "ED_EFDATE="
                        SQLQ = SQLQ & "(CASE WHEN " & xWDate & " IS NULL THEN NULL ELSE "
                        SQLQ = SQLQ & "DATEADD(YEAR,(CASE WHEN DATEADD(YEAR,YEAR(GETDATE())-YEAR(" & xWDate & ")," & xWDate & ")>GETDATE() "
                        SQLQ = SQLQ & "THEN YEAR(GETDATE())-YEAR(" & xWDate & ")-1 ELSE YEAR(GETDATE())-YEAR(" & xWDate & ") END),"
                        SQLQ = SQLQ & xWDate & ") END) "
                    End If
                End If
            Else
                'Ticket #25300: United Way of Lower Mainland - not going to compute new date range for them because
                'it will be compute when doing the Rollover with Anniversary Month.
                'If glbCompSerial <> "S/N - 2424W" Then
                'Ticket #25432 Franks 05/01/2014 - for new hire to setup ED_EFDATE, then no change for it, so commented this code above
                    SQLQ = "UPDATE HREMP SET "
                    SQLQ = SQLQ & "ED_EFDATE="
                    SQLQ = SQLQ & "(CASE WHEN " & xWDate & " IS NULL THEN NULL ELSE "
                    SQLQ = SQLQ & "DATEADD(YEAR,(CASE WHEN DATEADD(YEAR,YEAR(GETDATE())-YEAR(" & xWDate & ")," & xWDate & ")>GETDATE() "
                    SQLQ = SQLQ & "THEN YEAR(GETDATE())-YEAR(" & xWDate & ")-1 ELSE YEAR(GETDATE())-YEAR(" & xWDate & ") END),"
                    SQLQ = SQLQ & xWDate & ") END) "
                'End If
            End If
            
        Else
            SQLQ = "UPDATE HREMP SET "
            SQLQ = SQLQ & "ED_EFDATE="
            SQLQ = SQLQ & "IIF(" & xWDate & " IS NULL, NULL , "
            SQLQ = SQLQ & "DATEADD('yyyy',IIF(DATEADD('yyyy',1," & xWDate & ")>DATE() "
            SQLQ = SQLQ & ",0,YEAR(DATE())-YEAR(" & xWDate & ") - "
            SQLQ = SQLQ & " IIF(DATEADD('yyyy',YEAR(DATE())-YEAR(" & xWDate & ")," & xWDate & ") <=DATE() ,0, 1)"
            SQLQ = SQLQ & "),"
            SQLQ = SQLQ & xWDate & ")) "
        End If
        
        'County of Essex - Ticket #12676
        If glbCompSerial = "S/N - 2192W" Then
            If Not IsMissing(flgVacDates) Then
                If flgVacDates = True Then
                    SQLQ = SQLQ & "WHERE " & WSQLQ
                    gdbAdoIhr001.Execute SQLQ
                End If
            End If
        Else
            ''Ticket #25300: United Way of Lower Mainland - not going to compute new date range for them because
            ''it will be compute when doing the Rollover with Anniversary Month.
            'If glbCompSerial <> "S/N - 2424W" Then
                SQLQ = SQLQ & "WHERE " & WSQLQ
                'Only compute the Vacation From Date if From Date is Null, e.g. New Hires. For existing employees
                'do not change the date range (rollover to new period on Anniversary) because they Rollover OS
                'entitlement. The entire rollover will happen from Vacation Entitlement Master screen -> Year End.
                'Ticket #25432 Franks 05/01/2014 - for new hire to setup ED_EFDATE, then no change for it
                If glbCompSerial = "S/N - 2424W" Or glbCompSerial = "S/N - 2451W" Then
                    SQLQ = SQLQ & " AND ED_EFDATE IS NULL "
                End If
                gdbAdoIhr001.Execute SQLQ
            'End If
        End If
        'Hemu - Ticket #12649
        'If glbCompSerial = "S/N - 2296W" Then 'Essex County Library
        '    SQLQ = SQLQ & " AND NOT (HREMP.ED_PT = 'PT' AND HREMP.ED_ORG = 'CUPE') "
        'End If
        'gdbAdoIhr001.Execute SQLQ
        'MDIMain.panHelp(0).FloodPercent = 70
        
        'Compute Vacation TO Date
        If glbOracle Then
            SQLQ = "UPDATE HREMP SET "
            SQLQ = SQLQ & "ED_ETDATE= "
            SQLQ = SQLQ & "TO_DATE(add_months(ED_EFDATE,12) - 1) "
        ElseIf glbSQL Then
            'County of Essex - Ticket #12676
            If glbCompSerial = "S/N - 2192W" Then
                If Not IsMissing(flgVacDates) Then
                    If flgVacDates = True Then
                        SQLQ = "UPDATE HREMP SET "
                        SQLQ = SQLQ & "ED_ETDATE= "
                        SQLQ = SQLQ & "DATEADD(DAY,-1,DATEADD(YEAR,1,ED_EFDATE))"
                    End If
                End If
            Else
                SQLQ = "UPDATE HREMP SET "
                SQLQ = SQLQ & "ED_ETDATE= "
                SQLQ = SQLQ & "DATEADD(DAY,-1,DATEADD(YEAR,1,ED_EFDATE))"
            End If
        Else
            SQLQ = "UPDATE HREMP SET "
            SQLQ = SQLQ & "ED_ETDATE= "
            SQLQ = SQLQ & "DATEADD('d',-1,DATEADD('yyyy',1,ED_EFDATE)) "
        End If
        
        'County of Essex - Ticket #12676
        If glbCompSerial = "S/N - 2192W" Then
            If Not IsMissing(flgVacDates) Then
                If flgVacDates = True Then
                    SQLQ = SQLQ & "WHERE " & WSQLQ
                    gdbAdoIhr001.Execute SQLQ
                End If
            End If
        Else
            SQLQ = SQLQ & "WHERE " & WSQLQ
            gdbAdoIhr001.Execute SQLQ
        End If
        'Hemu - Ticket #12649
        'If glbCompSerial = "S/N - 2296W" Then 'Essex County Library
        '    SQLQ = SQLQ & " AND NOT (HREMP.ED_PT = 'PT' AND HREMP.ED_ORG = 'CUPE') "
        'End If
        'gdbAdoIhr001.Execute SQLQ
    End If
    '2.1 -- Set ED_EFDATE,ED_ETDATE in HREMP table - End
    
    'VAC taken - Begin
    If glbOracle Then
        MDIMain.panHelp(1).FloodPercent = 85
        
        'Hemu - Ticket #11332 - This is because ED_VACT does not get replaced with 0 when there are no records
        'in the HR_Attendance matching the criteria
        SQLQ = " Update HREMP SET "
        SQLQ = SQLQ & " ED_VACT = 0"
        SQLQ = SQLQ & " WHERE " & WSQLQ
        'Hemu - Ticket #12649
        'If glbCompSerial = "S/N - 2296W" Then 'Essex County Library
        '    SQLQ = SQLQ & " AND NOT (HREMP.ED_PT = 'PT' AND HREMP.ED_ORG = 'CUPE') "
        'End If
        gdbAdoIhr001.Execute SQLQ
        
        SQLQ = " Update HREMP SET "
        SQLQ = SQLQ & " ED_VACT =(SELECT SUM(AD_HRS) FROM HR_ATTENDANCE"
        SQLQ = SQLQ & " Where ED_EMPNBR = AD_EMPNBR"
        SQLQ = SQLQ & " AND (AD_DOA>= ED_EFDATE) AND (AD_DOA<=ED_ETDATE) "
        SQLQ = SQLQ & " AND (AD_REASON Like 'VAC%') )"
        SQLQ = SQLQ & " WHERE ED_EMPNBR IN"
        SQLQ = SQLQ & " (SELECT AD_EMPNBR FROM HR_ATTENDANCE WHERE HR_ATTENDANCE.AD_EMPNBR=HREMP.ED_EMPNBR"
        SQLQ = SQLQ & " AND (AD_DOA >= ED_EFDATE) AND (AD_DOA<=ED_ETDATE)"
        SQLQ = SQLQ & " AND (AD_REASON Like 'VAC%') )"
        SQLQ = SQLQ & " AND " & WSQLQ
        'Hemu - Ticket #12649
        'If glbCompSerial = "S/N - 2296W" Then 'Essex County Library
        '    SQLQ = SQLQ & " AND NOT (HREMP.ED_PT = 'PT' AND HREMP.ED_ORG = 'CUPE') "
        'End If
        gdbAdoIhr001.Execute SQLQ
    ElseIf glbSQL Then
    
        'Hemu - Ticket #11332 - This is because ED_VACT does not get replaced with 0 when there are no records
        'in the HR_Attendance matching the criteria
        SQLQ = " Update HREMP SET "
        SQLQ = SQLQ & " ED_VACT = 0"
        SQLQ = SQLQ & " WHERE " & WSQLQ
        'Hemu - Ticket #12649
        'If glbCompSerial = "S/N - 2296W" Then 'Essex County Library
        '    SQLQ = SQLQ & " AND NOT (HREMP.ED_PT = 'PT' AND HREMP.ED_ORG = 'CUPE') "
        'End If
        gdbAdoIhr001.Execute SQLQ
    
        SQLQ = " Update HREMP SET "
        SQLQ = SQLQ & " ED_VACT =(SELECT SUM(AD_HRS) FROM HR_ATTENDANCE"
        SQLQ = SQLQ & " Where ED_EMPNBR = AD_EMPNBR"
        SQLQ = SQLQ & " AND AD_DOA BETWEEN ED_EFDATE AND ED_ETDATE"
        SQLQ = SQLQ & " AND AD_REASON Like 'VAC%')"
        SQLQ = SQLQ & " WHERE ED_EMPNBR IN"
        SQLQ = SQLQ & " (SELECT AD_EMPNBR FROM HR_ATTENDANCE INNER JOIN HREMP ON HR_ATTENDANCE.AD_EMPNBR=HREMP.ED_EMPNBR"
        SQLQ = SQLQ & " WHERE (AD_DOA BETWEEN ED_EFDATE AND ED_ETDATE)"
        SQLQ = SQLQ & " AND AD_REASON Like 'VAC%')"
        SQLQ = SQLQ & " AND " & WSQLQ
        'Hemu - Ticket #12649
        'If glbCompSerial = "S/N - 2296W" Then 'Essex County Library
        '    SQLQ = SQLQ & " AND NOT (HREMP.ED_PT = 'PT' AND HREMP.ED_ORG = 'CUPE') "
        'End If
        gdbAdoIhr001.Execute SQLQ

    Else
        'Added by Bryan Ticket #11236, need to reset vacation taken, if there are no attendance records it will skip the person, leaving the existing taken value
        SQLQ = "UPDATE HREMP SET ED_VACT=0 WHERE " & WSQLQ
        gdbAdoIhr001.Execute SQLQ
        'end bryan
        SQLQ = "SELECT ED_EMPNBR, Sum(AD_HRS) AS SumHRS"
        SQLQ = SQLQ & " FROM HREMP INNER JOIN HR_ATTENDANCE ON HREMP.ED_EMPNBR = HR_ATTENDANCE.AD_EMPNBR"
        SQLQ = SQLQ & " WHERE AD_DOA>=ED_EFDATE And AD_DOA<=ED_ETDATE AND LEFT(AD_REASON,3)='VAC' AND " & WSQLQ
        'Hemu - Ticket #12649
        'If glbCompSerial = "S/N - 2296W" Then 'Essex County Library
        '    SQLQ = SQLQ & " AND NOT (HREMP.ED_PT = 'PT' AND HREMP.ED_ORG = 'CUPE') "
        'End If
        SQLQ = SQLQ & " GROUP BY ED_EMPNBR"
        rsTA.Open SQLQ, gdbAdoIhr001, adOpenKeyset
        Do Until rsTA.EOF
            gdbAdoIhr001.Execute "UPDATE HREMP SET ED_VACT=" & rsTA("SUMHRS") & " WHERE ED_EMPNBR=" & rsTA("ED_EMPNBR")
            rsTA.MoveNext
        Loop
        rsTA.Close
    End If
End If
'Part 2 - Vacation - End

'Part 3 - Sick - Begin
If xType = "SICK" Then
    '3.1 -- Set ED_EFDATES,ED_ETDATES in HREMP table - Begin
    MDIMain.panHelp(1).FloodPercent = 60
    If glbEntOutStandingS$ = "1" Then 'Based on 1 (Entitlement Date)
        SQLQ = "SELECT VE_DIV,VE_DEPT,VE_ORG,VE_LOC,VE_EMP,VE_PT,VE_GRPCD,VE_SECTION "
        SQLQ = SQLQ & " FROM HRSICKENT "
        If IsDate(xFromDateS) And IsDate(xToDateS) Then
            SQLQ = SQLQ & " WHERE VE_FRDATE=" & Date_SQL(xFromDateS)
            SQLQ = SQLQ & " AND VE_TODATE=" & Date_SQL(xToDateS)
        End If
        SQLQ = SQLQ & " GROUP BY VE_DIV,VE_DEPT,VE_ORG,VE_LOC,VE_EMP,VE_PT,VE_GRPCD,VE_SECTION "
        
        If rsVT.State <> 0 Then rsVT.Close
        rsVT.Open SQLQ, gdbAdoIhr001, adOpenStatic
        RecCNT = rsVT.RecordCount
        
        Do Until rsVT.EOF
            'MDIMain.panHelp(0).FloodPercent = (I / RecCNT) * 50 + 10: I = I + 1
            xDiv = rsVT("VE_DIV") & ""
            xDept = rsVT("VE_DEPT") & ""
            xORG = rsVT("VE_ORG") & ""
            xLoc = rsVT("VE_LOC") & ""
            xEMP = rsVT("VE_EMP") & ""
            xPT = rsVT("VE_PT") & ""
            xGRPCD = rsVT("VE_GRPCD") & ""
            xSec = rsVT("VE_SECTION") & ""
        
            If IsMissing(xFromDateS) Or IsMissing(xToDateS) Then
                SQLQ = "SELECT VE_FRDATE,VE_TODATE "
                SQLQ = SQLQ & " FROM HRSICKENT "
                SQLQ = SQLQ & " WHERE (VE_DIV = '" & xDiv & "' " & IIf(Len(xDiv) = 0, " OR VE_DIV IS NULL ", "") & ")"
                SQLQ = SQLQ & " AND (VE_DEPT = '" & xDept & "' " & IIf(Len(xDept) = 0, " OR VE_DEPT IS NULL ", "") & ")"
                SQLQ = SQLQ & " AND (VE_ORG = '" & xORG & "' " & IIf(Len(xORG) = 0, " OR VE_ORG IS NULL ", "") & ")"
                SQLQ = SQLQ & " AND (VE_LOC = '" & xLoc & "' " & IIf(Len(xLoc) = 0, " OR VE_LOC IS NULL ", "") & ")"
                SQLQ = SQLQ & " AND (VE_EMP = '" & xEMP & "' " & IIf(Len(xEMP) = 0, " OR VE_EMP IS NULL ", "") & ")"
                SQLQ = SQLQ & " AND (VE_PT = '" & xPT & "' " & IIf(Len(xPT) = 0, " OR VE_PT IS NULL ", "") & ")"
                SQLQ = SQLQ & " AND (VE_GRPCD = '" & xGRPCD & "' " & IIf(Len(xGRPCD) = 0, " OR VE_GRPCD IS NULL ", "") & ")"
                SQLQ = SQLQ & " AND (VE_SECTION = '" & xSec & "' " & IIf(Len(xSec) = 0, " OR VE_SECTION IS NULL ", "") & ")"
                SQLQ = SQLQ & " ORDER BY VE_FRDATE,VE_TODATE"
                If rsVTDate.State <> 0 Then rsVTDate.Close
                rsVTDate.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                nFrom = rsVTDate("VE_FRDATE")
                nTo = rsVTDate("VE_TODATE")
            Else
                nFrom = xFromDateS
                nTo = xToDateS
            End If
            If IsDate(nFrom) And IsDate(nTo) Then
                SQLQ = "UPDATE HREMP "
                SQLQ = SQLQ & " SET ED_EFDATES=" & Date_SQL(nFrom)
                SQLQ = SQLQ & " , ED_ETDATES=" & Date_SQL(nTo)
                SQLQ = SQLQ & " WHERE " & WSQLQ
                If Len(xDiv) > 0 Then SQLQ = SQLQ & " AND ED_DIV = '" & xDiv & "'"
                If Len(xDept) > 0 Then SQLQ = SQLQ & " AND ED_DEPTNO = '" & xDept & "'"
                If Len(xORG) > 0 Then SQLQ = SQLQ & " AND ED_ORG = '" & xORG & "'"
                If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #18235
                    If Len(xLoc) > 0 Then SQLQ = SQLQ & " AND ED_VADIM2 = '" & xLoc & "'"
                Else
                    If Len(xLoc) > 0 Then SQLQ = SQLQ & " AND ED_LOC = '" & xLoc & "'"
                End If
                If Len(xEMP) > 0 Then SQLQ = SQLQ & " AND ED_EMP = '" & xEMP & "'"
                If Len(xPT) > 0 Then SQLQ = SQLQ & " AND ED_PT = '" & xPT & "'"
                If Len(xGRPCD) > 0 Then SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_JOB IN (SELECT JB_CODE FROM HRJOB WHERE JB_GRPCD = '" & xGRPCD & "'))"
                If glbLinamar Then  'Hemu - Ticket #12494 - Section is actually Salary Distribution field for them on Entitl. Mstr
                    If Len(xSec) > 0 Then SQLQ = SQLQ & " AND ED_SALDIST = '" & xSec & "' "
                Else
                    If Len(xSec) > 0 Then SQLQ = SQLQ & " AND ED_SECTION = '" & xSec & "'"
                End If
                
                gdbAdoIhr001.BeginTrans
                gdbAdoIhr001.Execute SQLQ
                gdbAdoIhr001.CommitTrans
            End If
            rsVT.MoveNext
        Loop
    End If
    
    If glbEntOutStandingS$ <> "1" Then
        xWDateS = IIf(fglbWDateS$ = "", "ED_DOH", fglbWDateS$)
        If glbOracle Then
            SQLQ = "UPDATE HREMP SET "
            SQLQ = SQLQ & "ED_EFDATES="
            SQLQ = SQLQ & "(CASE WHEN " & xWDateS & " IS NULL THEN NULL ELSE "
            SQLQ = SQLQ & " ADD_MONTHS(" & xWDateS & ","
            SQLQ = SQLQ & "(CASE WHEN ADD_MONTHS(" & xWDateS & ",12)>SYSDATE THEN 0 ELSE "
            SQLQ = SQLQ & " (TO_NUMBER(SYSDATE,'YYYY')-TO_NUMBER(" & xWDateS & ",'YYYY'))*12 END)"
            SQLQ = SQLQ & " ) END) "
        ElseIf glbSQL Then
            SQLQ = "UPDATE HREMP SET "
            SQLQ = SQLQ & "ED_EFDATES="
            SQLQ = SQLQ & "(CASE WHEN " & xWDateS & " IS NULL THEN NULL ELSE "
            SQLQ = SQLQ & "DATEADD(YEAR,(CASE WHEN DATEADD(YEAR,YEAR(GETDATE())-YEAR(" & xWDateS & ")," & xWDateS & ")>GETDATE() "
            SQLQ = SQLQ & "THEN YEAR(GETDATE())-YEAR(" & xWDateS & ")-1 ELSE YEAR(GETDATE())-YEAR(" & xWDateS & ") END),"
            SQLQ = SQLQ & xWDateS & ") END) "
        Else
            SQLQ = "UPDATE HREMP SET "
            SQLQ = SQLQ & "ED_EFDATES="
            SQLQ = SQLQ & "IIF(" & xWDateS & " IS NULL , NULL,"
            SQLQ = SQLQ & "DATEADD('yyyy',IIF(DATEADD('yyyy',1," & xWDateS & ")>DATE() "
            SQLQ = SQLQ & ",0,YEAR(DATE())-YEAR(" & xWDateS & ") - "
            SQLQ = SQLQ & " IIF(DATEADD('yyyy',YEAR(DATE())-YEAR(" & xWDateS & ")," & xWDateS & ") <=DATE() ,0, 1)"
            SQLQ = SQLQ & "),"
            SQLQ = SQLQ & xWDateS & ")) "
        End If
        SQLQ = SQLQ & "WHERE " & WSQLQ
        gdbAdoIhr001.Execute SQLQ
        If glbOracle Then
            SQLQ = "UPDATE HREMP SET "
            SQLQ = SQLQ & "ED_ETDATES= "
            SQLQ = SQLQ & "TO_DATE(add_months(ED_EFDATES,12) - 1) "
        ElseIf glbSQL Then
            SQLQ = "UPDATE HREMP SET "
            SQLQ = SQLQ & "ED_ETDATES= "
            SQLQ = SQLQ & "DATEADD(DAY,-1,DATEADD(YEAR,1,ED_EFDATES)) "
        Else
            SQLQ = "UPDATE HREMP SET "
            SQLQ = SQLQ & "ED_ETDATES= "
            SQLQ = SQLQ & "DATEADD('d',-1,DATEADD('yyyy',1,ED_EFDATES)) "
        End If
        SQLQ = SQLQ & "WHERE " & WSQLQ
        gdbAdoIhr001.Execute SQLQ
    End If
    
    'SICK taken - Begin
    If glbOracle Then
        SQLQ = " Update HREMP SET "
        SQLQ = SQLQ & " HREMP.ED_SICKT =(SELECT SUM(AD_HRS) FROM HR_ATTENDANCE"
        SQLQ = SQLQ & " Where ED_EMPNBR = AD_EMPNBR"
        SQLQ = SQLQ & " AND (AD_DOA>= ED_EFDATES) AND (AD_DOA<=ED_ETDATES )"
        SQLQ = SQLQ & " AND (AD_REASON Like 'SIC%') )"
        SQLQ = SQLQ & " WHERE ED_EMPNBR IN"
        SQLQ = SQLQ & " (SELECT AD_EMPNBR FROM HR_ATTENDANCE WHERE HR_ATTENDANCE.AD_EMPNBR=HREMP.ED_EMPNBR"
        SQLQ = SQLQ & " AND (AD_DOA>= ED_EFDATES) AND (AD_DOA<=ED_ETDATES)"
        SQLQ = SQLQ & " AND (AD_REASON Like 'SIC%') )"
        SQLQ = SQLQ & " AND " & WSQLQ
        gdbAdoIhr001.Execute SQLQ

    ElseIf glbSQL Then
        SQLQ = " Update HREMP SET "
        SQLQ = SQLQ & " HREMP.ED_SICKT =(SELECT SUM(AD_HRS) FROM HR_ATTENDANCE"
        SQLQ = SQLQ & " Where ED_EMPNBR = AD_EMPNBR"
        SQLQ = SQLQ & " AND AD_DOA BETWEEN ED_EFDATES AND ED_ETDATES"
        SQLQ = SQLQ & " AND AD_REASON Like 'SIC%')"
        SQLQ = SQLQ & " WHERE ED_EMPNBR IN"
        SQLQ = SQLQ & " (SELECT AD_EMPNBR FROM HR_ATTENDANCE INNER JOIN HREMP ON HR_ATTENDANCE.AD_EMPNBR=HREMP.ED_EMPNBR"
        SQLQ = SQLQ & " WHERE (AD_DOA BETWEEN ED_EFDATES AND ED_ETDATES)"
        SQLQ = SQLQ & " AND AD_REASON Like 'SIC%')"
        SQLQ = SQLQ & " AND " & WSQLQ
        gdbAdoIhr001.Execute SQLQ
    Else
        'Added by Bryan Ticket #11236, need to reset vacation taken, if there are no attendance records it will skip the person, leaving the existing taken value
        SQLQ = "UPDATE HREMP SET ED_SICKT=0 WHERE " & WSQLQ
        gdbAdoIhr001.Execute SQLQ
        SQLQ = "SELECT ED_EMPNBR, Sum(AD_HRS) AS SumHRS"
        SQLQ = SQLQ & " FROM HREMP INNER JOIN HR_ATTENDANCE ON HREMP.ED_EMPNBR = HR_ATTENDANCE.AD_EMPNBR"
        SQLQ = SQLQ & " WHERE AD_DOA>=ED_EFDATES And AD_DOA<=ED_ETDATES AND LEFT(AD_REASON,3)='SIC' AND " & WSQLQ
        SQLQ = SQLQ & " GROUP BY ED_EMPNBR "
        rsTA.Open SQLQ, gdbAdoIhr001, adOpenDynamic
        Do Until rsTA.EOF
            gdbAdoIhr001.Execute "UPDATE HREMP SET ED_SICKT=" & rsTA("SUMHRS") & " WHERE ED_EMPNBR=" & rsTA("ED_EMPNBR")
            rsTA.MoveNext
        Loop
        rsTA.Close
    End If
End If
'Part 3 - Sick - End





'Part 4 - Calculate the Entitlement from the Accrual Table - Begin
If glbVadim Then    'And glbLambton Then
    If glbOracle Then
        'Sick Entitlement
        SQLQ = " Update HREMP SET "
        SQLQ = SQLQ & " HREMP.ED_SICK =(SELECT SUM(AC_HRS) FROM HR_ACCRUAL"
        SQLQ = SQLQ & " WHERE ED_EMPNBR = AC_EMPNBR"
        SQLQ = SQLQ & " AND (AC_EDATE>= ED_EFDATES) AND (AC_EDATE<=ED_ETDATES )"
        SQLQ = SQLQ & " AND (AC_TYPE Like 'SIC%') AND (AC_HRS>0) AND (NOT (AC_COMMENTS LIKE '%Prev%')) )"
        SQLQ = SQLQ & " WHERE ED_EMPNBR IN"
        SQLQ = SQLQ & " (SELECT AC_EMPNBR FROM HR_ACCRUAL WHERE HR_ACCRUAL.AC_EMPNBR=HREMP.ED_EMPNBR"
        SQLQ = SQLQ & " AND (AC_EDATE>= ED_EFDATES) AND (AC_EDATE<=ED_ETDATES)"
        SQLQ = SQLQ & " AND (AC_TYPE Like 'SIC%') AND (NOT (AC_COMMENTS LIKE '%Prev%')) )"
        SQLQ = SQLQ & " AND " & WSQLQ
        gdbAdoIhr001.Execute SQLQ
        
        'Vacation Entitlement
        'If glbEntOutStanding$ <> "1" Then
            SQLQ = "UPDATE HREMP SET "
            SQLQ = SQLQ & " ED_VAC = (SELECT SUM(AC_HRS) FROM HR_ACCRUAL WHERE "
            SQLQ = SQLQ & " ED_EMPNBR = AC_EMPNBR "
            SQLQ = SQLQ & " AND (AC_EDATE>=ED_EFDATE) AND (AC_EDATE<=ED_ETDATE) "
            SQLQ = SQLQ & " AND (AC_TYPE LIKE 'VAC%') AND (AC_HRS>0) AND (NOT (AC_COMMENTS LIKE '%Prev%')))"
            SQLQ = SQLQ & " WHERE ED_EMPNBR IN "
            SQLQ = SQLQ & " (SELECT AC_EMPNBR FROM HR_ACCRUAL WHERE HR_ACCRUAL.AC_EMPNBR = HREMP.ED_EMPNBR "
            SQLQ = SQLQ & " AND (AC_EDATE >= ED_EFDATE) AND (AC_EDATE<=ED_ETDATE)"
            SQLQ = SQLQ & " AND (AC_TYPE Like 'VAC%') AND (NOT (AC_COMMENTS LIKE '%Prev%')) )"
            SQLQ = SQLQ & " AND " & WSQLQ
            gdbAdoIhr001.Execute SQLQ
        'End If
    ElseIf glbSQL Then
        SQLQ = " Update HREMP SET "
        SQLQ = SQLQ & " HREMP.ED_SICK =(SELECT SUM(AC_HRS) FROM HR_ACCRUAL "
        SQLQ = SQLQ & " WHERE ED_EMPNBR = AC_EMPNBR"
        SQLQ = SQLQ & " AND AC_EDATE BETWEEN ED_EFDATES AND ED_ETDATES"
        SQLQ = SQLQ & " AND AC_TYPE LIKE 'SIC%' AND AC_HRS>0 AND (NOT (AC_COMMENTS LIKE '%Prev%')))"
        SQLQ = SQLQ & " WHERE ED_EMPNBR IN"
        SQLQ = SQLQ & " (SELECT AC_EMPNBR FROM HR_ACCRUAL INNER JOIN HREMP ON HR_ACCRUAL.AC_EMPNBR=HREMP.ED_EMPNBR"
        SQLQ = SQLQ & " WHERE (AC_EDATE BETWEEN ED_EFDATES AND ED_ETDATES)"
        SQLQ = SQLQ & " AND AC_TYPE Like 'SIC%' AND (NOT (AC_COMMENTS LIKE '%Prev%')))"
        SQLQ = SQLQ & " AND " & WSQLQ
        gdbAdoIhr001.Execute SQLQ
        
        'If glbEntOutStanding$ <> "1" Then
            SQLQ = " Update HREMP SET "
            SQLQ = SQLQ & " ED_VAC =(SELECT SUM(AC_HRS) FROM HR_ACCRUAL"
            SQLQ = SQLQ & " WHERE ED_EMPNBR = AC_EMPNBR"
            SQLQ = SQLQ & " AND AC_EDATE BETWEEN ED_EFDATE AND ED_ETDATE"
            SQLQ = SQLQ & " AND AC_TYPE LIKE 'VAC%' AND (AC_HRS>0) AND (NOT (AC_COMMENTS LIKE '%Prev%')))"
            SQLQ = SQLQ & " WHERE ED_EMPNBR IN"
            SQLQ = SQLQ & " (SELECT AC_EMPNBR FROM HR_ACCRUAL INNER JOIN HREMP ON HR_ACCRUAL.AC_EMPNBR=HREMP.ED_EMPNBR"
            SQLQ = SQLQ & " WHERE (AC_EDATE BETWEEN ED_EFDATE AND ED_ETDATE)"
            SQLQ = SQLQ & " AND AC_TYPE Like 'VAC%' AND (NOT (AC_COMMENTS LIKE '%Prev%')))"
            SQLQ = SQLQ & " AND " & WSQLQ
            gdbAdoIhr001.Execute SQLQ
        'End If
    Else
        SQLQ = "SELECT ED_EMPNBR, Sum(AC_HRS) AS SumHRS"
        SQLQ = SQLQ & " FROM HREMP INNER JOIN HR_ACCRUAL ON HREMP.ED_EMPNBR = HR_ACCRUAL.AC_EMPNBR"
        SQLQ = SQLQ & " WHERE AC_EDATE>=ED_EFDATES And AC_EDATE<=ED_ETDATES AND LEFT(AC_TYPE,3)='SIC' AND (AC_HRS>0) AND (NOT (AC_COMMENTS LIKE '%Prev%')) AND " & WSQLQ
        SQLQ = SQLQ & " GROUP BY ED_EMPNBR "
        rsTA.Open SQLQ, gdbAdoIhr001, adOpenDynamic
        Do Until rsTA.EOF
            gdbAdoIhr001.Execute "UPDATE HREMP SET ED_SICK=" & rsTA("SUMHRS") & " WHERE ED_EMPNBR=" & rsTA("ED_EMPNBR")
            rsTA.MoveNext
        Loop
        rsTA.Close
                
        'If glbEntOutStanding$ <> "1" Then
            SQLQ = "SELECT ED_EMPNBR, Sum(AC_HRS) AS SumHRS"
            SQLQ = SQLQ & " FROM HREMP INNER JOIN HR_ACCRUAL ON HREMP.ED_EMPNBR = HR_ACCRUAL.AD_EMPNBR"
            SQLQ = SQLQ & " WHERE AC_EDATE>=ED_EFDATE And AC_EDATE<=ED_ETDATE AND LEFT(AC_TYPE,3)='VAC' AND (AC_HRS>0) AND (NOT (AC_COMMENTS LIKE '%Prev%')) AND " & WSQLQ
            SQLQ = SQLQ & " GROUP BY ED_EMPNBR"
            rsTA.Open SQLQ, gdbAdoIhr001, adOpenKeyset
            Do Until rsTA.EOF
                gdbAdoIhr001.Execute "UPDATE HREMP SET ED_VAC=" & rsTA("SUMHRS") & " WHERE ED_EMPNBR=" & rsTA("ED_EMPNBR")
                rsTA.MoveNext
            Loop
            rsTA.Close
        'End If
    End If
    
    gdbAdoIhr001.Execute "Update HREMP SET ED_VAC=0 WHERE ED_VAC IS NULL"
    gdbAdoIhr001.Execute "Update HREMP SET ED_SICK=0 WHERE ED_SICK IS NULL"
        
End If
'Part 4 - Vacation & Sick Entitlement - End


glbENTScreen = True

MDIMain.panHelp(1).FloodPercent = 100
MDIMain.panHelp(1).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""
Exit Sub


ErrorHandler:
glbFrmCaption$ = "Entitlement Recalculation"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EntReCalcPeriod", "", CStr(SQLQ))
If gintRollBack% = False Then
    Resume Next
End If
End Sub

Public Sub Recalculate_KerrysPlaceLieu(ByVal WSQLQ As String, Optional xOutVal)
'for Webview:
'Outstanding Lieu from info:HR ed_user_num2
'logic:
'After the attendance is imported from WebTime, the view will contain the outstanding hours
'against an employee number. Outstanding is calculated by summing all Attendance Master records
'beginning with "OT" and subtracting all Attendance Master records beginning with "CT".
Dim rsEmp As New ADODB.Recordset
Dim xOutstanding As Double
Dim SQLQ
Dim I, xNTot

SQLQ = "SELECT * FROM HREMP WHERE (1=1) AND "
SQLQ = SQLQ & WSQLQ

rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If Not rsEmp.EOF Then
    rsEmp.MoveFirst
    I = 0
    xNTot = rsEmp.RecordCount
    Do While Not rsEmp.EOF
        DoEvents
        MDIMain.panHelp(0).FloodPercent = (I / xNTot) * 100: I = I + 1
        If IsMissing(xOutVal) Then
            xOutstanding = (Get_OvertimeBank(rsEmp("ED_EMPNBR"), "", "") - Get_OvertimeTaken(rsEmp("ED_EMPNBR"), "", ""))
        Else
            xOutstanding = xOutVal
        End If
        If IsNumeric(xOutstanding) Then
            rsEmp("ED_USER_NUM2") = xOutstanding
            rsEmp.Update
        End If
        rsEmp.MoveNext
    Loop
End If
rsEmp.Close

End Sub

Private Sub Recalculate_VitalaireOTBANK(ByVal WSQLQ As String)
Dim rsEmp As New ADODB.Recordset
Dim rsAttend As New ADODB.Recordset
Dim rsAttendCT As New ADODB.Recordset
Dim SQLQ

'Set ED_OTBANK to zero for the first time otherwise Null will be updated if some Value - Null
SQLQ = "UPDATE HREMP SET ED_OTBANK = 0 WHERE ED_OTBANK IS NULL"
gdbAdoIhr001.Execute SQLQ

SQLQ = "SELECT ED_EMPNBR, ED_EFDATE, ED_ETDATE, ED_OTBANK FROM HREMP WHERE (1=1) AND "
SQLQ = SQLQ & WSQLQ

rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If Not rsEmp.EOF Then
    rsEmp.MoveFirst
    
    Do While Not rsEmp.EOF
        If Not IsNull(rsEmp("ED_EFDATE")) And Not IsNull(rsEmp("ED_ETDATE")) Then
            SQLQ = "SELECT SUM(AD_HRS) AS OT_SUM FROM HR_ATTENDANCE WHERE AD_REASON = 'VCO' AND AD_EMPNBR = " & rsEmp("ED_EMPNBR") & " "
            SQLQ = SQLQ & " AND AD_DOA >= " & Date_SQL(rsEmp("ED_EFDATE")) & " "
            SQLQ = SQLQ & " AND AD_DOA <= " & Date_SQL(rsEmp("ED_ETDATE")) & " "
            SQLQ = SQLQ & " GROUP BY AD_EMPNBR"
            If rsAttend.State <> 0 Then rsAttend.Close
            rsAttend.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic

            If Not rsAttend.EOF Then
                If rsAttend("OT_SUM") > 0 Then
                    rsEmp("ED_OTBANK") = rsAttend("OT_SUM")
                    rsEmp.Update
                End If
            End If
            rsAttend.Close
            'rsAttendCT.Close
        End If
        rsEmp.MoveNext
    Loop
End If
rsEmp.Close

End Sub


Sub EntReCalc(ByVal WSQLQ As String, Optional flgVacDates As Boolean, Optional strType As String, Optional strEntType As String)
'Apr 3,2003 Ticket#3943
'Provide the ability to have multiple vacation years overriding the company master file
Dim xlen, xxx, xx1, x
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim rsT_PARCO As New ADODB.Recordset
Dim rsEmpVacB As New ADODB.Recordset 'For County of Brant
Dim rsEmpVacE As New ADODB.Recordset 'For County of Brant
Dim rsEmpBack As New ADODB.Recordset   'For County of Brant
Dim rsTC As New ADODB.Recordset
Dim rsTD As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset 'For Multi Vacation Entitlement Periods, glbEntOutStanding$ = "1"
Dim XCNTIND
Dim EmpCNT, xEmpnbr, I
Dim xWDate, xWDateS
Dim fglbWDate$, fglbWDateS$
Dim SQLQ, SQLQ1, SQLQ2

Dim RecCNT
Dim xDiv, xDept, xORG, xLoc, xEMP, xPT, xGRPCD, xSec
Dim rsVT As New ADODB.Recordset
Dim rsVTDate As New ADODB.Recordset
Dim xFromDate, xToDate, xFromDateS, xToDateS

On Error GoTo ErrorHandler

'Merged in EntReCalc since v7.6, old codes in v7.4
'If glbVadim And glbLambton Then
'    Call Vadim_EntReCalc(WSQLQ)
'    Exit Sub
'End If

'Merged in EntReCalc since v7.6, old codes in v7.4
'For Essex County Library
'The Recalculation Function doesn't include Vacation Entitlement if ED_PT = 'PT' and ED_ORG = 'CUPE'
'Vacation Entitlement if ED_PT = 'PT' and ED_ORG = 'CUPE' only can be calcultaed
'in Annual Mass Update of SN2296.exe
'If glbCompSerial = "S/N - 2296W" Then
'    Call EntReCalc2296(WSQLQ)
'    Exit Sub
'End If

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 1
MDIMain.panHelp(0).FloodPercent = 3
WSQLQ = glbSeleDeptUn & IIf(WSQLQ = "", " ", " AND ") & WSQLQ
WSQLQ = Replace(WSQLQ, "ED_", "HREMP.ED_")

If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #14635
    Call Recalculate_VitalaireOTBANK(WSQLQ)
End If

If glbCompSerial = "S/N - 2433W" Then 'Kerry's Place Ticket #22332 Franks 07/26/2012
    Call Recalculate_KerrysPlaceLieu(WSQLQ)
End If

'removed Brant - Bryan 18/Apr/2006 Ticket# 10495
If glbCompSerial = "S/N - 2288W" Or glbCompSerial = "S/N - 2371W" Or (glbVadim And glbLambton) Then     'Also Vacation Rule for Musashi
    Dim ArrVac(1000, 6), NumBrant
    
    If glbCompSerial <> "S/N - 2288W" And glbCompSerial <> "S/N - 2371W" And Not glbLambton Then
        SQLQ = "DELETE FROM HRVacBrant"
        gdbAdoIhr001B.BeginTrans
        gdbAdoIhr001B.Execute SQLQ
        gdbAdoIhr001B.CommitTrans
    End If
    
    SQLQ = "SELECT HREMP.ED_EMPNBR,HREMP.ED_PVAC,HREMP.ED_VAC,HREMP.ED_VACT,HREMP.ED_EFDATE,HREMP.ED_ETDATE FROM HREMP"
    SQLQ = SQLQ & " WHERE " & WSQLQ
    rsEmpVacB.Open SQLQ, gdbAdoIhr001, adOpenStatic
    NumBrant = 0
    If Not rsEmpVacB.EOF Then
        rsEmpVacB.MoveLast
        rsEmpVacB.MoveFirst
    End If
    x = 1
    Do While Not rsEmpVacB.EOF And x < 1000
        ArrVac(x, 1) = rsEmpVacB("ED_EMPNBR")
        ArrVac(x, 2) = rsEmpVacB("ED_PVAC")
        ArrVac(x, 3) = rsEmpVacB("ED_VAC")
        ArrVac(x, 4) = rsEmpVacB("ED_VACT")
        ArrVac(x, 5) = rsEmpVacB("ED_EFDATE")
        ArrVac(x, 6) = rsEmpVacB("ED_ETDATE")
        x = x + 1
        rsEmpVacB.MoveNext
    Loop
    NumBrant = x
End If

'Part 1 - Set Default values - Begin
MDIMain.panHelp(0).FloodPercent = 8
DoEvents
If glbOracle Then
    SQLQ = "Update HREMP SET "
    SQLQ = SQLQ & "( HREMP.ED_ENTOPT, HREMP.ED_ENTOPTS, HREMP.ED_EMLT,HREMP.ED_SICKT )="
    SQLQ = SQLQ & " ( select "
    SQLQ = SQLQ & " HRPARCO.PC_ENTOUT, HRPARCO.PC_ENTOUTS, 0, 0 from HRPARCO where HREMP.ED_COMPNO = HRPARCO.PC_CO ) "
ElseIf glbSQL Then
    SQLQ = "Update HREMP SET "
    SQLQ = SQLQ & " HREMP.ED_ENTOPT = HRPARCO.PC_ENTOUT ,"
    SQLQ = SQLQ & " HREMP.ED_ENTOPTS = HRPARCO.PC_ENTOUTS ,"
    SQLQ = SQLQ & " HREMP.ED_DHRS = qry_Assigned_Jobs.JH_DHRS ,"
    SQLQ = SQLQ & " HREMP.ED_EMLT = 0,"
    SQLQ = SQLQ & " HREMP.ED_SICKT = 0 "
    SQLQ = SQLQ & " FROM (HREMP INNER JOIN HRPARCO ON HREMP.ED_COMPNO = HRPARCO.PC_CO) LEFT JOIN qry_Assigned_Jobs"
    SQLQ = SQLQ & " ON HREMP.ED_EMPNBR = qry_Assigned_Jobs.ED_EMPNBR "
Else
    SQLQ = "UPDATE (HREMP INNER JOIN HRPARCO ON HREMP.ED_COMPNO = HRPARCO.PC_CO) "
    SQLQ = SQLQ & " LEFT JOIN qry_Assigned_Jobs ON HREMP.ED_EMPNBR = qry_Assigned_Jobs.ED_EMPNBR "
    SQLQ = SQLQ & " SET "
    SQLQ = SQLQ & " HREMP.ED_ENTOPT = [HRPARCO].[PC_ENTOUT],"
    SQLQ = SQLQ & " HREMP.ED_ENTOPTS = [HRPARCO].[PC_ENTOUTS], "
    SQLQ = SQLQ & " HREMP.ED_DHRS = [qry_Assigned_Jobs].[JH_DHRS], "
    SQLQ = SQLQ & " HREMP.ED_EMLT = 0, "
    SQLQ = SQLQ & " HREMP.ED_SICKT = 0"
End If
SQLQ = SQLQ & " WHERE " & WSQLQ
gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans

DoEvents
If glbOracle Then 'ED_DHRS
    SQLQ = "Update HREMP SET "
    SQLQ = SQLQ & "HREMP.ED_DHRS="
    SQLQ = SQLQ & " ( select DISTINCT "
    SQLQ = SQLQ & " HR_JOB_HISTORY.JH_DHRS FROM HR_JOB_HISTORY where HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_CURRENT<>0 ) "
    SQLQ = SQLQ & " WHERE " & WSQLQ
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
End If
Select Case glbEntOutStanding$
    Case "2": fglbWDate$ = "ED_DOH"
    Case "3": fglbWDate$ = "ED_SENDTE"
    Case "4": fglbWDate$ = "ED_LTHIRE"
    Case "5": fglbWDate$ = "ED_USRDAT1"
    Case "6": fglbWDate$ = "ED_UNION"
End Select
Select Case glbEntOutStandingS$ ' sets field reference for basic 'which date'
    Case "2": fglbWDateS$ = "ED_DOH"
    Case "3": fglbWDateS$ = "ED_SENDTE"
    Case "4": fglbWDateS$ = "ED_LTHIRE"
    Case "5": fglbWDateS$ = "ED_USRDAT1"
    Case "6": fglbWDateS$ = "ED_UNION"
End Select

gdbAdoIhr001.BeginTrans
MDIMain.panHelp(0).FloodPercent = 10
If glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #24729 01/21/2014 Franks
    SQLQ = "SELECT DISTINCT ED_SIN FROM HREMP "
    rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset
    If Not rsTB.EOF Then
        rsT_PARCO.Open "HRPARCO", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
        rsT_PARCO("PC_NUMBER_EMPLOYEES") = rsTB.RecordCount
        rsT_PARCO.Update
        rsT_PARCO.Close
        rsTB.Close
    End If
Else
    SQLQ = "SELECT COUNT(ED_EMPNBR) AS EMPCOUNT FROM HREMP "
    If glbCompSerial = "S/N - 2394W" Then  'St. John - Ticket #17446
        SQLQ = SQLQ & "WHERE ED_EMP<>'TERM'"
    End If
    rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset
    rsT_PARCO.Open "HRPARCO", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    rsT_PARCO("PC_NUMBER_EMPLOYEES") = rsTB("EMPCOUNT")
    rsT_PARCO.Update
    rsT_PARCO.Close
    rsTB.Close
End If


SQLQ = "UPDATE HREMP SET ED_DHRS=NULL WHERE ED_DHRS=0 "
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans

SQLQ = "UPDATE HREMP SET ED_VAC=0 WHERE ED_VAC IS NULL "
gdbAdoIhr001.Execute SQLQ
SQLQ = "UPDATE HREMP SET ED_PVAC=0 WHERE ED_PVAC IS NULL "
gdbAdoIhr001.Execute SQLQ

SQLQ = "UPDATE HREMP SET ED_SICK=0 WHERE ED_SICK IS NULL "
gdbAdoIhr001.Execute SQLQ
SQLQ = "UPDATE HREMP SET ED_PSICK=0 WHERE ED_PSICK IS NULL "
gdbAdoIhr001.Execute SQLQ

SQLQ = "UPDATE HREMP SET ED_INCIDCNT=0,"
SQLQ = SQLQ & " ED_ENTOPT='" & glbEntOutStanding$ & "',"
SQLQ = SQLQ & " ED_ENTOPTS='" & glbEntOutStandingS$ & "'"
SQLQ = SQLQ & " WHERE " & WSQLQ
gdbAdoIhr001.Execute SQLQ
MDIMain.panHelp(0).FloodPercent = 20

DoEvents
' ED_INCIDCNT - Begin
If glbOracle Then
    SQLQ = "UPDATE HREMP SET HREMP.ED_INCIDCNT=( select DISTINCT qry_INCID.INCIDNBR "
    SQLQ = SQLQ & " FROM qry_INCID where  HREMP.ED_EMPNBR = qry_INCID.EMPNBR )"
    SQLQ = SQLQ & " WHERE " & WSQLQ
    gdbAdoIhr001.Execute SQLQ
ElseIf glbSQL Then
'Incident Number for Musashi May 27,2002
    If glbCompSerial = "S/N - 2288W" Then
        SQLQ = "SELECT * FROM HREMP "
        SQLQ = SQLQ & " WHERE " & WSQLQ
        If rsTC.State <> 0 Then rsTC.Close
        rsTC.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
        Do While Not rsTC.EOF
            XCNTIND = 0
            SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsTC("ED_EMPNBR") & ""
            SQLQ = SQLQ & " AND AD_INCID<>0"
            If rsTD.State <> 0 Then rsTD.Close
            rsTD.Open SQLQ, gdbAdoIhr001, adOpenStatic
            Do While Not rsTD.EOF
                If IsDate(rsTD("AD_DOA")) Then
                    If DateDiff("d", CVDate(rsTD("AD_DOA")), DateAdd("m", -6, Now)) <= 0 Then
                        If rsTD("AD_INCID") Then
                            XCNTIND = XCNTIND + 1
                        End If
                    End If
                End If
                rsTD.MoveNext
            Loop
            rsTD.Close
            rsTC("ED_INCIDCNT") = XCNTIND
            rsTC.Update
            rsTC.MoveNext
        Loop
        rsTC.Close
    Else
        SQLQ = "UPDATE HREMP SET HREMP.ED_INCIDCNT=qry_INCID.INCIDNBR "
        SQLQ = SQLQ & " FROM HREMP INNER JOIN qry_INCID ON HREMP.ED_EMPNBR = qry_INCID.EMPNBR"
        SQLQ = SQLQ & " WHERE " & WSQLQ
        gdbAdoIhr001.Execute SQLQ
    End If
Else
    SQLQ = "UPDATE HREMP RIGHT JOIN HR_ATTENDANCE ON HREMP.ED_EMPNBR = HR_ATTENDANCE.AD_EMPNBR "
    SQLQ = SQLQ & " SET HREMP.ED_INCIDCNT = HREMP.ED_INCIDCNT-HR_ATTENDANCE.AD_INCID "
    SQLQ = SQLQ & " WHERE AD_INCID<>0"
    SQLQ = SQLQ & " AND " & WSQLQ
    gdbAdoIhr001.Execute SQLQ

End If
' ED_INCIDCNT - End

' Shift - Begin

If Not glbMulti Then
    If glbOracle Then
        SQLQ = "UPDATE HREMP SET HREMP.ED_SHIFT="
        SQLQ = SQLQ & "(SELECT DISTINCT HR_JOB_HISTORY.JH_SHIFT FROM HR_JOB_HISTORY "
        SQLQ = SQLQ & "WHERE HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND JH_CURRENT<>0)"
        SQLQ = SQLQ & " WHERE " & WSQLQ
    ElseIf glbSQL Then
        SQLQ = "UPDATE HREMP SET HREMP.ED_SHIFT=HR_JOB_HISTORY.JH_SHIFT "
        SQLQ = SQLQ & " FROM HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR"
        SQLQ = SQLQ & " WHERE JH_CURRENT<>0"
        SQLQ = SQLQ & " AND " & WSQLQ
    Else
        SQLQ = "UPDATE HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
        SQLQ = SQLQ & " SET HREMP.ED_SHIFT = HR_JOB_HISTORY.JH_SHIFT"
        SQLQ = SQLQ & " WHERE JH_CURRENT<>0"
        SQLQ = SQLQ & " AND " & WSQLQ
    End If
    gdbAdoIhr001.Execute SQLQ

End If
' Shift - End

'Part 1 - Set Default values - End

DoEvents
'Part 2 - Vacation - Begin
'Set ED_EFDATE,ED_ETDATE in HREMP table , VAC Taken
MDIMain.panHelp(0).FloodPercent = 40
'County of Essex - Ticket #12676
If glbCompSerial = "S/N - 2192W" Then
    If Not IsMissing(flgVacDates) Then
        Call EntReCalcPeriod(WSQLQ, "VAC", , , , , flgVacDates)
    End If
Else
    'Essex County Library
    If glbCompSerial = "S/N - 2296W" Then
        SQLQ = "UPDATE HREMP SET ED_VACT=0 WHERE " & WSQLQ
        gdbAdoIhr001.Execute SQLQ
        SQLQ = "SELECT ED_EMPNBR, Sum(AD_HRS) AS SumHRS"
        SQLQ = SQLQ & " FROM HREMP INNER JOIN HR_ATTENDANCE ON HREMP.ED_EMPNBR = HR_ATTENDANCE.AD_EMPNBR"
        SQLQ = SQLQ & " WHERE AD_DOA>=ED_EFDATE And AD_DOA<=ED_ETDATE AND LEFT(AD_REASON,3)='VAC' AND " & WSQLQ
        SQLQ = SQLQ & " GROUP BY ED_EMPNBR"
        rsTA.Open SQLQ, gdbAdoIhr001, adOpenKeyset
        Do Until rsTA.EOF
            gdbAdoIhr001.Execute "UPDATE HREMP SET ED_VACT=" & rsTA("SUMHRS") & " WHERE ED_EMPNBR=" & rsTA("ED_EMPNBR")
            rsTA.MoveNext
        Loop
        rsTA.Close
    ElseIf Not IsMissing(strType) Then
        If strType = "TAKEN ONLY" Then
            SQLQ = " Update HREMP SET ED_VACT = 0 WHERE " & WSQLQ
            gdbAdoIhr001.Execute SQLQ
        
            SQLQ = " Update HREMP SET "
            SQLQ = SQLQ & " ED_VACT =(SELECT SUM(AD_HRS) FROM HR_ATTENDANCE"
            SQLQ = SQLQ & " Where ED_EMPNBR = AD_EMPNBR"
            SQLQ = SQLQ & " AND AD_DOA BETWEEN ED_EFDATE AND ED_ETDATE"
            SQLQ = SQLQ & " AND AD_REASON Like 'VAC%')"
            SQLQ = SQLQ & " WHERE ED_EMPNBR IN"
            SQLQ = SQLQ & " (SELECT AD_EMPNBR FROM HR_ATTENDANCE INNER JOIN HREMP ON HR_ATTENDANCE.AD_EMPNBR=HREMP.ED_EMPNBR"
            SQLQ = SQLQ & " WHERE (AD_DOA BETWEEN ED_EFDATE AND ED_ETDATE)"
            SQLQ = SQLQ & " AND AD_REASON Like 'VAC%')"
            SQLQ = SQLQ & " AND " & WSQLQ
            gdbAdoIhr001.Execute SQLQ
        Else
            If Not IsMissing(strEntType) Then
                If strEntType = "HOURSBASED" Then
                    Call EntReCalcPeriod(WSQLQ, "VAC", , , , , , strEntType)
                Else
                    'Ticket #29230 - Daily Entitlement
                    If glbCompEntVacDaily Then
                        Call EntReCalcPeriod_Daily(WSQLQ, "VAC")
                    Else
                        Call EntReCalcPeriod(WSQLQ, "VAC")
                    End If
                End If
            Else
                'Ticket #29230 - Daily Entitlement
                If glbCompEntVacDaily Then
                    Call EntReCalcPeriod_Daily(WSQLQ, "VAC")
                Else
                    Call EntReCalcPeriod(WSQLQ, "VAC")
                End If
            End If
        End If
    Else
        If Not IsMissing(strEntType) Then
            If strEntType = "HOURSBASED" Then
                Call EntReCalcPeriod(WSQLQ, "VAC", , , , , , strEntType)
            Else
                'Ticket #29230 - Daily Entitlement
                If glbCompEntVacDaily Then
                    Call EntReCalcPeriod_Daily(WSQLQ, "VAC")
                Else
                    Call EntReCalcPeriod(WSQLQ, "VAC")
                End If
            End If
        Else
            'Ticket #29230 - Daily Entitlement
            If glbCompEntVacDaily Then
                Call EntReCalcPeriod_Daily(WSQLQ, "VAC")
            Else
                Call EntReCalcPeriod(WSQLQ, "VAC")
            End If
        End If
    End If
End If

'Part 3 - Sick - Begin
'Set ED_EFDATES,ED_ETDATES in HREMP table , SICK Taken
MDIMain.panHelp(0).FloodPercent = 60
If Not IsMissing(strType) Then
    If strType = "TAKEN ONLY" Then
        SQLQ = " Update HREMP SET "
        SQLQ = SQLQ & " HREMP.ED_SICKT =(SELECT SUM(AD_HRS) FROM HR_ATTENDANCE"
        SQLQ = SQLQ & " Where ED_EMPNBR = AD_EMPNBR"
        SQLQ = SQLQ & " AND AD_DOA BETWEEN ED_EFDATES AND ED_ETDATES"
        SQLQ = SQLQ & " AND AD_REASON Like 'SIC%')"
        SQLQ = SQLQ & " WHERE ED_EMPNBR IN"
        SQLQ = SQLQ & " (SELECT AD_EMPNBR FROM HR_ATTENDANCE INNER JOIN HREMP ON HR_ATTENDANCE.AD_EMPNBR=HREMP.ED_EMPNBR"
        SQLQ = SQLQ & " WHERE (AD_DOA BETWEEN ED_EFDATES AND ED_ETDATES)"
        SQLQ = SQLQ & " AND AD_REASON Like 'SIC%')"
        SQLQ = SQLQ & " AND " & WSQLQ
        gdbAdoIhr001.Execute SQLQ
    Else
        Call EntReCalcPeriod(WSQLQ, "SICK")
    End If
Else
    Call EntReCalcPeriod(WSQLQ, "SICK")
End If

'Part 4 - EMLT - Begin
MDIMain.panHelp(0).FloodPercent = 80

DoEvents
If glbOracle Then
    MDIMain.panHelp(0).FloodPercent = 90
    SQLQ = " Update HREMP SET "
    'linamar stuff added by Bryan 19/Oct/05 Ticket#9552
    If Not glbLinamar And (glbCompSerial <> "S/N - 2288W") Then  'Musashi - Ticket #16786
        SQLQ = SQLQ & " ED_EMLT =(SELECT SUM(AD_HRS) FROM HR_ATTENDANCE"
    Else
        'Release 8.0 - Ticket #24545: Jerry wants this to be same as the EML Report, i.e.
        'to show actual # of days taken and outstanding instead of rounding up to 1.
        'SQLQ = SQLQ & " ED_EMLT =(SELECT count(AD_HRS) FROM HR_ATTENDANCE"
        SQLQ = SQLQ & " ED_EMLT =(SELECT SUM(AD_HRS) FROM HR_ATTENDANCE"
    End If
    SQLQ = SQLQ & " Where ED_EMPNBR = AD_EMPNBR"
    SQLQ = SQLQ & " AND TO_CHAR(AD_DOA,'YYYY')='" & Format(Date, "yyyy") & "'" '& "TO_DATE(" & Date_SQL(Date) & ",'YYYY') "
    SQLQ = SQLQ & " AND AD_EMELEA <>0)"
    SQLQ = SQLQ & " WHERE ED_EMPNBR IN"
    SQLQ = SQLQ & " (SELECT AD_EMPNBR FROM HR_ATTENDANCE WHERE HR_ATTENDANCE.AD_EMPNBR=HREMP.ED_EMPNBR"
    SQLQ = SQLQ & " AND TO_CHAR(AD_DOA,'YYYY')='" & Format(Date, "yyyy") & "'" '" TO_DATE( " & Date_SQL(Date) & " ,'YYYY')"
    SQLQ = SQLQ & " AND AD_EMELEA <>0)"
    SQLQ = SQLQ & " AND " & WSQLQ
    gdbAdoIhr001.Execute SQLQ
    MDIMain.panHelp(0).FloodPercent = 95
ElseIf glbSQL Then
    MDIMain.panHelp(0).FloodPercent = 90
    SQLQ = " Update HREMP SET "
    'linamar stuff added by Bryan 19/Oct/05 Ticket#9552
    If Not glbLinamar And (glbCompSerial <> "S/N - 2288W") Then  'Musashi - Ticket #16786
        SQLQ = SQLQ & " ED_EMLT =(SELECT SUM(AD_HRS) FROM HR_ATTENDANCE"
    Else
        'Release 8.0 - Ticket #24545: Jerry wants this to be same as the EML Report, i.e.
        'to show actual # of days taken and outstanding instead of rounding up to 1.
        If glbCompSerial = "S/N - 2282W" Or glbCompSerial = "S/N - 2393W" Then   'Ticket #28102 - KTH Shelburne - Every occurence of EML - 1 day
            SQLQ = SQLQ & " ED_EMLT =(SELECT Count(AD_HRS) FROM HR_ATTENDANCE"
        Else
            SQLQ = SQLQ & " ED_EMLT =(SELECT SUM(AD_HRS) FROM HR_ATTENDANCE"
        End If
    End If
    SQLQ = SQLQ & " Where ED_EMPNBR = AD_EMPNBR"
    SQLQ = SQLQ & " AND YEAR(AD_DOA)=" & Year(Date)
    SQLQ = SQLQ & " AND AD_EMELEA <>0)"
    SQLQ = SQLQ & " WHERE ED_EMPNBR IN"
    SQLQ = SQLQ & " (SELECT AD_EMPNBR FROM HR_ATTENDANCE INNER JOIN HREMP ON HR_ATTENDANCE.AD_EMPNBR=HREMP.ED_EMPNBR"
    SQLQ = SQLQ & " WHERE YEAR(AD_DOA)=" & Year(Date)
    SQLQ = SQLQ & " AND AD_EMELEA <>0)"
    SQLQ = SQLQ & " AND " & WSQLQ
    gdbAdoIhr001.Execute SQLQ
    MDIMain.panHelp(0).FloodPercent = 95
Else
    'Emergency Leave update - Hemu
    'linamar stuff added by Bryan 19/Oct/05 Ticket#9552
    If Not glbLinamar And (glbCompSerial <> "S/N - 2288W") Then  'Musashi - Ticket #16786
        SQLQ = "SELECT ED_EMPNBR, Sum(AD_HRS) AS SumHRS"
    Else
        'Release 8.0 - Ticket #24545: Jerry wants this to be same as the EML Report, i.e.
        'to show actual # of days taken and outstanding instead of rounding up to 1.
        'SQLQ = "SELECT ED_EMPNBR, count(AD_HRS) AS SumHRS"
        SQLQ = "SELECT ED_EMPNBR, sum(AD_HRS) AS SumHRS"
    End If
    SQLQ = SQLQ & " FROM HREMP INNER JOIN HR_ATTENDANCE ON HREMP.ED_EMPNBR = HR_ATTENDANCE.AD_EMPNBR"
    SQLQ = SQLQ & " WHERE YEAR(AD_DOA)= " & Year(Date) & " AND AD_EMELEA <>0 AND " & WSQLQ
    SQLQ = SQLQ & " GROUP BY ED_EMPNBR"
    rsTA.Open SQLQ, gdbAdoIhr001, adOpenKeyset
    Do Until rsTA.EOF
        gdbAdoIhr001.Execute "UPDATE HREMP SET ED_EMLT=" & rsTA("SUMHRS") & " WHERE ED_EMPNBR=" & rsTA("ED_EMPNBR")
        rsTA.MoveNext
    Loop
    rsTA.Close

End If
'gdbAdoIhr001.CommitTrans
'Part 4 - EMLT  - End
'removed Brant - Bryan 18/Apr/2006 Ticket# 10495
'Part 5 - Brant - Begin
'If glbCBrant Then 'This function is only for VACATION Setup "Annual"
'                  'If VACATION Setup "Monthly", rsEmpVacE("ED_VAC") = rsEmpVacE("ED_VAC") + Previous Vaction
'                  'There is another Recalculation function in Status/Dates Screen
'    Call Pause(2)
'    rsEmpBack.Open "HRVacBrant", gdbAdoIhr001B, adOpenKeyset, adLockOptimistic, adCmdTableDirect
'    For x = 1 To NumBrant - 1
'        SQLQ = "SELECT HREMP.ED_EMPNBR,HREMP.ED_PVAC,HREMP.ED_VAC,HREMP.ED_VACT,HREMP.ED_EFDATE,HREMP.ED_ETDATE FROM HREMP"
'        SQLQ = SQLQ & " WHERE ED_EMPNBR = " & ArrVac(x, 1)
'        rsEmpVacE.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'
'        If Not rsEmpVacE.EOF Then
'            If ArrVac(x, 5) <> rsEmpVacE("ED_EFDATE") Then
'                rsEmpVacE("ED_VAC") = 0
'                rsEmpVacE("ED_PVAC") = ArrVac(x, 2) + ArrVac(x, 3) - ArrVac(x, 4)
'                rsEmpVacE.Update
'                rsEmpBack.AddNew
'                rsEmpBack("ED_COMPNO") = "001"
'                rsEmpBack("ED_EMPNBR") = ArrVac(x, 1)
'                rsEmpBack("ED_PVAC") = ArrVac(x, 2)
'                rsEmpBack("ED_VAC") = ArrVac(x, 3)
'                rsEmpBack("ED_VACT") = ArrVac(x, 4)
'                rsEmpBack("ED_EFDATE") = ArrVac(x, 5)
'                rsEmpBack("ED_ETDATE") = ArrVac(x, 6)
'                rsEmpBack("ED_REDATE") = Format(Now, "Short Date")
'                rsEmpBack.Update
'            End If
'        End If
'        rsEmpVacE.Close
'    Next
'    rsEmpBack.Close
'End If
'Part 5 - Brant - End

'Part 6 - Calculate the Entitlement from the Accrual Table - Begin
If glbVadim Then    'And glbLambton Then
    If glbOracle Then
        'Sick Entitlement
        SQLQ = " Update HREMP SET "
        SQLQ = SQLQ & " HREMP.ED_SICK =(SELECT (CASE WHEN SUM(AC_HRS) IS NULL THEN 0 ELSE SUM(AC_HRS) END) FROM HR_ACCRUAL"
        SQLQ = SQLQ & " WHERE ED_EMPNBR = AC_EMPNBR"
        SQLQ = SQLQ & " AND (AC_EDATE>= ED_EFDATES) AND (AC_EDATE<=ED_ETDATES )"
        '===========================
        'SQLQ = SQLQ & " AND (AC_TYPE Like 'SIC%') AND (AC_HRS>0) AND (NOT (AC_COMMENTS LIKE '%Prev%')) )"
        SQLQ = SQLQ & " AND (AC_TYPE Like 'SIC%') AND (NOT (AC_COMMENTS LIKE '%Prev%')) )"
        'SQLQ = SQLQ & " WHERE ED_EMPNBR IN"
        SQLQ = SQLQ & " WHERE " & WSQLQ
        '===========================
        SQLQ = SQLQ & " (SELECT AC_EMPNBR FROM HR_ACCRUAL WHERE HR_ACCRUAL.AC_EMPNBR=HREMP.ED_EMPNBR"
        SQLQ = SQLQ & " AND (AC_EDATE>= ED_EFDATES) AND (AC_EDATE<=ED_ETDATES)"
        SQLQ = SQLQ & " AND (AC_TYPE Like 'SIC%') AND (NOT (AC_COMMENTS LIKE '%Prev%')) )"
        SQLQ = SQLQ & " AND " & WSQLQ
        gdbAdoIhr001.Execute SQLQ
        
        'Vacation Entitlement
        'If glbEntOutStanding$ <> "1" Then
            SQLQ = "UPDATE HREMP SET "
            SQLQ = SQLQ & " ED_VAC = (SELECT (CASE WHEN SUM(AC_HRS) IS NULL THEN 0 ELSE SUM(AC_HRS) END) FROM HR_ACCRUAL WHERE "
            SQLQ = SQLQ & " ED_EMPNBR = AC_EMPNBR "
            SQLQ = SQLQ & " AND (AC_EDATE>=ED_EFDATE) AND (AC_EDATE<=ED_ETDATE) "
            '===========================
            'SQLQ = SQLQ & " AND (AC_TYPE LIKE 'VAC%') AND (AC_HRS>0) AND (NOT (AC_COMMENTS LIKE '%Prev%')))"
            SQLQ = SQLQ & " AND (AC_TYPE LIKE 'VAC%') AND (NOT (AC_COMMENTS LIKE '%Prev%')))"
            'SQLQ = SQLQ & " WHERE ED_EMPNBR IN "
            SQLQ = SQLQ & " WHERE " & WSQLQ
            '===========================
            'SQLQ = SQLQ & " (SELECT AC_EMPNBR FROM HR_ACCRUAL WHERE HR_ACCRUAL.AC_EMPNBR = HREMP.ED_EMPNBR "
            'SQLQ = SQLQ & " AND (AC_EDATE >= ED_EFDATE) AND (AC_EDATE<=ED_ETDATE)"
            'SQLQ = SQLQ & " AND (AC_TYPE Like 'VAC%') AND (NOT (AC_COMMENTS LIKE '%Prev%')) )"
            'SQLQ = SQLQ & " AND " & WSQLQ
            gdbAdoIhr001.Execute SQLQ
        'End If
    ElseIf glbSQL Then
        SQLQ = " Update HREMP SET "
        SQLQ = SQLQ & " HREMP.ED_SICK =(SELECT (CASE WHEN SUM(AC_HRS) IS NULL THEN 0 ELSE SUM(AC_HRS) END) FROM HR_ACCRUAL "
        SQLQ = SQLQ & " WHERE ED_EMPNBR = AC_EMPNBR"
        SQLQ = SQLQ & " AND AC_EDATE BETWEEN ED_EFDATES AND ED_ETDATES"
        '===========================
        'SQLQ = SQLQ & " AND AC_TYPE LIKE 'SIC%' AND AC_HRS>0 AND (NOT (AC_COMMENTS LIKE '%Prev%')))"
        SQLQ = SQLQ & " AND AC_TYPE LIKE 'SIC%' AND (NOT (AC_COMMENTS LIKE '%Prev%')))"
        SQLQ = SQLQ & " WHERE " & WSQLQ
        'SQLQ = SQLQ & " WHERE ED_EMPNBR IN"
        '===========================
        'SQLQ = SQLQ & " (SELECT AC_EMPNBR FROM HR_ACCRUAL INNER JOIN HREMP ON HR_ACCRUAL.AC_EMPNBR=HREMP.ED_EMPNBR"
        'SQLQ = SQLQ & " WHERE (AC_EDATE BETWEEN ED_EFDATES AND ED_ETDATES)"
        'SQLQ = SQLQ & " AND AC_TYPE Like 'SIC%' AND (NOT (AC_COMMENTS LIKE '%Prev%')))"
        'SQLQ = SQLQ & " AND " & WSQLQ
        gdbAdoIhr001.Execute SQLQ
        
        'If glbEntOutStanding$ <> "1" Then
            SQLQ = " Update HREMP SET "
            SQLQ = SQLQ & " ED_VAC =(SELECT (CASE WHEN SUM(AC_HRS) IS NULL THEN 0 ELSE SUM(AC_HRS) END) FROM HR_ACCRUAL"
            SQLQ = SQLQ & " WHERE ED_EMPNBR = AC_EMPNBR"
            SQLQ = SQLQ & " AND AC_EDATE BETWEEN ED_EFDATE AND ED_ETDATE"
            '===========================
            'SQLQ = SQLQ & " AND AC_TYPE LIKE 'VAC%' AND (AC_HRS>0) AND (NOT (AC_COMMENTS LIKE '%Prev%')))"
            SQLQ = SQLQ & " AND AC_TYPE LIKE 'VAC%' AND (NOT (AC_COMMENTS LIKE '%Prev%')))"
            SQLQ = SQLQ & " WHERE " & WSQLQ
            'SQLQ = SQLQ & " WHERE ED_EMPNBR IN "
            '===========================
            'SQLQ = SQLQ & " (SELECT AC_EMPNBR FROM HR_ACCRUAL INNER JOIN HREMP ON HR_ACCRUAL.AC_EMPNBR=HREMP.ED_EMPNBR"
            'SQLQ = SQLQ & " WHERE (AC_EDATE BETWEEN ED_EFDATE AND ED_ETDATE)"
            'SQLQ = SQLQ & " AND AC_TYPE Like 'VAC%' AND (NOT (AC_COMMENTS LIKE '%Prev%')))"
            'SQLQ = SQLQ & " AND " & WSQLQ
            gdbAdoIhr001.Execute SQLQ
        'End If
    Else
        SQLQ = "SELECT ED_EMPNBR, Sum(AC_HRS) AS SumHRS"
        SQLQ = SQLQ & " FROM HREMP INNER JOIN HR_ACCRUAL ON HREMP.ED_EMPNBR = HR_ACCRUAL.AC_EMPNBR"
        '===========================
        'SQLQ = SQLQ & " WHERE AC_EDATE>=ED_EFDATES And AC_EDATE<=ED_ETDATES AND LEFT(AC_TYPE,3)='SIC' AND (AC_HRS>0) AND (NOT (AC_COMMENTS LIKE '%Prev%')) AND " & WSQLQ
        SQLQ = SQLQ & " WHERE AC_EDATE>=ED_EFDATES And AC_EDATE<=ED_ETDATES AND LEFT(AC_TYPE,3)='SIC' AND (NOT (AC_COMMENTS LIKE '%Prev%')) AND " & WSQLQ
        '===========================
        SQLQ = SQLQ & " GROUP BY ED_EMPNBR "
        rsTA.Open SQLQ, gdbAdoIhr001, adOpenDynamic
        Do Until rsTA.EOF
            gdbAdoIhr001.Execute "UPDATE HREMP SET ED_SICK=" & rsTA("SUMHRS") & " WHERE ED_EMPNBR=" & rsTA("ED_EMPNBR")
            rsTA.MoveNext
        Loop
        rsTA.Close
                
        'If glbEntOutStanding$ <> "1" Then
            SQLQ = "SELECT ED_EMPNBR, Sum(AC_HRS) AS SumHRS"
            SQLQ = SQLQ & " FROM HREMP INNER JOIN HR_ACCRUAL ON HREMP.ED_EMPNBR = HR_ACCRUAL.AD_EMPNBR"
            '===========================
            'SQLQ = SQLQ & " WHERE AC_EDATE>=ED_EFDATE And AC_EDATE<=ED_ETDATE AND LEFT(AC_TYPE,3)='VAC' AND (AC_HRS>0) AND (NOT (AC_COMMENTS LIKE '%Prev%')) AND " & WSQLQ
            SQLQ = SQLQ & " WHERE AC_EDATE>=ED_EFDATE And AC_EDATE<=ED_ETDATE AND LEFT(AC_TYPE,3)='VAC' AND (NOT (AC_COMMENTS LIKE '%Prev%')) AND " & WSQLQ
            '===========================
            SQLQ = SQLQ & " GROUP BY ED_EMPNBR"
            rsTA.Open SQLQ, gdbAdoIhr001, adOpenKeyset
            Do Until rsTA.EOF
                gdbAdoIhr001.Execute "UPDATE HREMP SET ED_VAC=" & rsTA("SUMHRS") & " WHERE ED_EMPNBR=" & rsTA("ED_EMPNBR")
                rsTA.MoveNext
            Loop
            rsTA.Close
        'End If
    End If
    
    gdbAdoIhr001.Execute "Update HREMP SET ED_VAC=0 WHERE ED_VAC IS NULL"
    gdbAdoIhr001.Execute "Update HREMP SET ED_SICK=0 WHERE ED_SICK IS NULL"
    
    'gdbAdoIhr001.Execute "Update HREMP SET ED_VAC=ED_VAC-ED_PVAC WHERE ED_PVAC < ED_VAC"
    'gdbAdoIhr001.Execute "Update HREMP SET ED_SICK=ED_SICK-ED_PSICK WHERE ED_PSICK < ED_SICK"
    
    'gdbAdoIhr001.Execute "Update HREMP SET ED_VAC=ED_PVAC-ED_VAC WHERE ED_VAC < ED_PVAC"
    'gdbAdoIhr001.Execute "Update HREMP SET ED_SICK=ED_PSICK-ED_SICK WHERE ED_SICK < ED_PSICK"
    
    'gdbAdoIhr001.Execute "Update HREMP SET ED_VAC=0 WHERE ED_VAC < 0"
    'gdbAdoIhr001.Execute "Update HREMP SET ED_SICK=0 WHERE ED_SICK < 0"
    
End If
'Part 6 - Calculate the Entitlement from the Accrual Table - End

MDIMain.panHelp(0).FloodPercent = 97

'Part 7 - ED_PVAC - Begin
If (glbCompSerial = "S/N - 2288W" Or glbCompSerial = "S/N - 2371W" Or (glbVadim And glbLambton)) And glbEntOutStanding$ <> "1" Then     'Vacation Rule for Musashi
    For x = 1 To NumBrant - 1
        SQLQ = "SELECT HREMP.ED_EMPNBR,HREMP.ED_PVAC,HREMP.ED_EFDATE FROM HREMP"
        SQLQ = SQLQ & " WHERE ED_EMPNBR = " & ArrVac(x, 1)
        rsEmpVacE.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
      
        If Not rsEmpVacE.EOF Then
            If ArrVac(x, 5) <> rsEmpVacE("ED_EFDATE") Then
                rsEmpVacE("ED_PVAC") = (ArrVac(x, 2) + ArrVac(x, 3)) - ArrVac(x, 4)
                rsEmpVacE.Update
            End If
        End If
        rsEmpVacE.Close
    Next
End If
'Part 7 - ED_PVAC - End

glbENTScreen = True

MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""
Exit Sub


ErrorHandler:
glbFrmCaption$ = "Entitlement Recalculation"
glbErrNum& = Err

'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EntReCalc", "", CStr(SQLQ))
If gintRollBack% = False Then
    Resume Next
End If
End Sub

Sub Vadim_EntReCalcHr()
Dim SQLQ
On Error GoTo ErrorHandler

MDIMain.panHelp(0).FloodType = 1            '28July99 js
MDIMain.panHelp(1).Caption = " Please Wait" '
MDIMain.panHelp(2).Caption = ""             '
MDIMain.panHelp(0).FloodPercent = 10

gdbAdoIhr001.BeginTrans
If glbOracle Then
    SQLQ = "Update HRENTHRS"
    SQLQ = SQLQ & " SET HRENTHRS.HE_TAKEN = 0,"
    SQLQ = SQLQ & " HRENTHRS.HE_DHRS = (SELECT SUM(JH_DHRS) FROM qry_Assigned_jobs"
    SQLQ = SQLQ & " WHERE ED_EMPNBR=HE_EMPNBR)"
    gdbAdoIhr001.Execute SQLQ
ElseIf glbSQL Then
    SQLQ = "Update HRENTHRS"
    SQLQ = SQLQ & " SET HRENTHRS.HE_TAKEN = 0,"
    SQLQ = SQLQ & " HRENTHRS.HE_DHRS = (SELECT SUM(JH_DHRS) FROM qry_Assigned_Jobs"
    SQLQ = SQLQ & " WHERE ED_EMPNBR=HE_EMPNBR)"
    gdbAdoIhr001.Execute SQLQ
Else
    gdbAdoIhr001.Execute "qry_HrEntitle"
End If

MDIMain.panHelp(0).FloodPercent = 30

If glbOracle Then
     SQLQ = " Update HRENTHRS "
     SQLQ = SQLQ & "SET HE_TAKEN = HE_TAKEN + (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE (AD_DOA BETWEEN HE_FDATE And HE_TDATE) AND AD_REASON= HE_TYPE AND AD_EMPNBR= HE_EMPNBR)"
     gdbAdoIhr001.Execute SQLQ
ElseIf glbSQL Then
    SQLQ = "UPDATE  HRENTHRS "
    SQLQ = SQLQ & " SET HE_TAKEN = HE_TAKEN + (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE (AD_DOA BETWEEN HE_FDATE And HE_TDATE) AND AD_REASON= HE_TYPE AND AD_EMPNBR= HE_EMPNBR)"
    SQLQ = SQLQ & " WHERE HE_EMPNBR IN (SELECT AD_EMPNBR FROM HR_ATTENDANCE,HRENTHRS WHERE (AD_DOA BETWEEN HE_FDATE And HE_TDATE) AND AD_REASON= HE_TYPE)"
    gdbAdoIhr001.Execute SQLQ
Else
    SQLQ = "UPDATE  HRENTHRS LEFT JOIN HR_ATTENDANCE ON (HRENTHRS.HE_TYPE = HR_ATTENDANCE.AD_REASON) AND "
    SQLQ = SQLQ & "(HRENTHRS.HE_EMPNBR = HR_ATTENDANCE.AD_EMPNBR) SET HRENTHRS.HE_TAKEN = [HRENTHRS].[HE_TAKEN]+[HR_ATTENDANCE].[AD_HRS] "
    SQLQ = SQLQ & "WHERE (((HR_ATTENDANCE.AD_DOA)>=[HRENTHRS].[HE_FDATE] And (HR_ATTENDANCE.AD_DOA)<=[HRENTHRS].[HE_TDATE]) AND "
    SQLQ = SQLQ & "((HR_ATTENDANCE.AD_REASON)=[HRENTHRS].[HE_TYPE]) AND ((HR_ATTENDANCE.AD_EMPNBR)=[HRENTHRS].[HE_EMPNBR]))"
    gdbAdoIhr001.Execute SQLQ
End If
gdbAdoIhr001.Execute "Update HRENTHRS SET HE_TAKEN=0 WHERE HE_TAKEN IS NULL"
gdbAdoIhr001.CommitTrans
    
MDIMain.panHelp(0).FloodPercent = 50


gdbAdoIhr001.BeginTrans
'Calculate the Entitlement from the HR_ACCRUAL
If glbOracle Then
    SQLQ = "UPDATE HRENTHRS "
    'SQLQ = SQLQ & " SET HE_ENTITLE = (SELECT SUM(AC_HRS) FROM HR_ACCRUAL WHERE (AC_EDATE BETWEEN HE_FDATE AND HE_TDATE) AND AC_TYPE = HE_TYPE AND AC_EMPNBR = HE_EMPNBR AND AC_HRS>0)"
    'Ticket #24653
    'SQLQ = SQLQ & " SET HE_ENTITLE = (SELECT SUM(AC_HRS) FROM HR_ACCRUAL WHERE (AC_EDATE BETWEEN HE_FDATE AND HE_TDATE) AND AC_TYPE = HE_TYPE AND AC_EMPNBR = HE_EMPNBR AND (AC_HRS>0 OR AC_ACTION='M'))"
    SQLQ = SQLQ & " SET HE_ENTITLE = (SELECT SUM(AC_HRS) FROM HR_ACCRUAL WHERE (AC_EDATE BETWEEN HE_FDATE AND HE_TDATE) AND AC_TYPE = HE_TYPE AND AC_EMPNBR = HE_EMPNBR AND AC_COMMENTS NOT LIKE 'Prev.%' AND (AC_HRS>0 OR AC_ACTION='M' OR AC_ACTION='C'))"
    gdbAdoIhr001.Execute SQLQ
ElseIf glbSQL Then
    SQLQ = "UPDATE  HRENTHRS "
    'SQLQ = SQLQ & " SET HE_ENTITLE = (SELECT SUM(AC_HRS) FROM HR_ACCRUAL WHERE (AC_EDATE BETWEEN HE_FDATE AND HE_TDATE) AND AC_TYPE = HE_TYPE AND AC_EMPNBR = HE_EMPNBR AND AC_HRS>0)"
    'Ticket #24653
    'SQLQ = SQLQ & " SET HE_ENTITLE = (SELECT SUM(AC_HRS) FROM HR_ACCRUAL WHERE (AC_EDATE BETWEEN HE_FDATE AND HE_TDATE) AND AC_TYPE = HE_TYPE AND AC_EMPNBR = HE_EMPNBR AND (AC_HRS>0 OR AC_ACTION='M'))"
    SQLQ = SQLQ & " SET HE_ENTITLE = (SELECT SUM(AC_HRS) FROM HR_ACCRUAL WHERE (AC_EDATE BETWEEN HE_FDATE AND HE_TDATE) AND AC_TYPE = HE_TYPE AND AC_EMPNBR = HE_EMPNBR AND AC_COMMENTS NOT LIKE 'Prev.%' AND (AC_HRS>0 OR AC_ACTION='M' OR AC_ACTION='C'))"
    SQLQ = SQLQ & " WHERE HE_EMPNBR IN (SELECT AC_EMPNBR FROM HR_ACCRUAL,HRENTHRS WHERE (AC_EDATE BETWEEN HE_FDATE And HE_TDATE) AND AC_TYPE = HE_TYPE)"
    gdbAdoIhr001.Execute SQLQ
Else
    SQLQ = "UPDATE  HRENTHRS LEFT JOIN HR_ACCRUAL ON (HRENTHRS.HE_TYPE = HR_ACCRUAL.AC_TYPE) AND "
    SQLQ = SQLQ & "(HRENTHRS.HE_EMPNBR = HR_ACCRUAL.AC_EMPNBR) SET HRENTHRS.HE_ENTITLE = [HRENTHRS].[HE_ENTITLE] + [HR_ACCRUAL].[AC_HRS] "
    SQLQ = SQLQ & "WHERE (((HR_ACCRUAL.AC_EDATE)>=[HRENTHRS].[HE_FDATE] And (HR_ACCRUAL.AC_EDATE)<=[HRENTHRS].[HE_TDATE]) AND "
    'SQLQ = SQLQ & "((HR_ACCRUAL.AC_TYPE)=[HRENTHRS].[HE_TYPE]) AND ((HR_ACCRUAL.AC_EMPNBR)=[HRENTHRS].[HE_EMPNBR]) AND ([HR_ACCRUAL].[AC_HRS]>0))"
    'Ticket #24653
    'SQLQ = SQLQ & "((HR_ACCRUAL.AC_TYPE)=[HRENTHRS].[HE_TYPE]) AND ((HR_ACCRUAL.AC_EMPNBR)=[HRENTHRS].[HE_EMPNBR]) AND ([HR_ACCRUAL].[AC_HRS]>0 OR [HR_ACCRUAL].[AC_ACTION]='M'))"
    SQLQ = SQLQ & "((HR_ACCRUAL.AC_TYPE)=[HRENTHRS].[HE_TYPE]) AND ((HR_ACCRUAL.AC_EMPNBR)=[HRENTHRS].[HE_EMPNBR]) AND [HR_ACCRUAL].[AC_COMMENTS] NOT LIKE 'Prev.%' AND ([HR_ACCRUAL].[AC_HRS]>0 OR [HR_ACCRUAL].[AC_ACTION]='M' OR [HR_ACCRUAL].[AC_ACTION]='C'))"
    gdbAdoIhr001.Execute SQLQ
End If
gdbAdoIhr001.Execute "Update HRENTHRS SET HE_ENTITLE=0 WHERE HE_ENTITLE IS NULL"
gdbAdoIhr001.CommitTrans

MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""
Exit Sub

ErrorHandler:
glbFrmCaption$ = "Hourly Entitlement Recalculation"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EntReCalcHr", "", "qry_HrEntitle1/qry_HrEntitle")
If gintRollBack% = False Then
    Resume Next
End If

End Sub

Sub EntReCalcHrTerm(Optional xTermSEQ) 'Ticket #23417 Franks 05/31/2013
Dim SQLQ

On Error GoTo ErrorHandler


MDIMain.panHelp(0).FloodType = 1            '28July99 js
MDIMain.panHelp(1).Caption = " Please Wait" '
MDIMain.panHelp(2).Caption = ""             '
MDIMain.panHelp(0).FloodPercent = 10

SQLQ = "UPDATE Term_ENTHRS SET HE_TAKEN = 0 "
If Not IsMissing(xTermSEQ) Then SQLQ = SQLQ & "WHERE Term_ENTHRS.TERM_SEQ = " & xTermSEQ & " "
gdbAdoIhr001.Execute SQLQ

SQLQ = "UPDATE Term_ENTHRS SET Term_ENTHRS.HE_DHRS = Term_JOB_HISTORY.JH_DHRS "
SQLQ = SQLQ & "FROM Term_ENTHRS LEFT JOIN Term_JOB_HISTORY ON Term_ENTHRS.TERM_SEQ = Term_JOB_HISTORY.TERM_SEQ "
SQLQ = SQLQ & "WHERE NOT Term_JOB_HISTORY.JH_CURRENT = 0 "
If Not IsMissing(xTermSEQ) Then SQLQ = SQLQ & "AND Term_ENTHRS.TERM_SEQ = " & xTermSEQ & " "
gdbAdoIhr001.Execute SQLQ


MDIMain.panHelp(0).FloodPercent = 30

SQLQ = "UPDATE Term_ENTHRS SET Term_ENTHRS.HE_TAKEN = "
SQLQ = SQLQ & "(SELECT SUM(AD_HRS) FROM Term_ATTENDANCE WHERE (AD_DOA BETWEEN HE_FDATE And HE_TDATE) AND AD_REASON= HE_TYPE AND Term_ENTHRS.TERM_SEQ= Term_ATTENDANCE.TERM_SEQ) "
SQLQ = SQLQ & "WHERE Term_ENTHRS.TERM_SEQ IN (SELECT Term_ATTENDANCE.TERM_SEQ FROM Term_ATTENDANCE,Term_ENTHRS WHERE (AD_DOA BETWEEN HE_FDATE And HE_TDATE) AND AD_REASON= HE_TYPE) "
If Not IsMissing(xTermSEQ) Then SQLQ = SQLQ & "AND Term_ENTHRS.TERM_SEQ = " & xTermSEQ & " "
gdbAdoIhr001.Execute SQLQ

gdbAdoIhr001.Execute "Update Term_ENTHRS SET HE_TAKEN=0 WHERE HE_TAKEN IS NULL"

    
'Ticket #17924 - Begin
'New Logic using XXX+ for Entitlement and XXX- for Taken
'ENTITLEMENT (+)
'If there are XXX+ coded records then update HE_ENTITLE with those values
'Note: Entitlement update from Hourly Entitlement Master screen on a RIGHT(HE_TYPE,1) = '+' will create a new record
'HR_ATTENDANCE with AD_REASON = HE_TYPE, AD_HRS = HE_ENTITLE, AD_DOA = HE_FDATE
'gdbAdoIhr001.BeginTrans
    
'Alternatively
SQLQ = "UPDATE Term_ENTHRS "
SQLQ = SQLQ & " SET HE_ENTITLE = (SELECT SUM(AD_HRS) FROM Term_ATTENDANCE WHERE (AD_DOA BETWEEN HE_FDATE And HE_TDATE) AND AD_REASON = HE_TYPE AND Term_ENTHRS.TERM_SEQ= Term_ATTENDANCE.TERM_SEQ)"
SQLQ = SQLQ & " WHERE RIGHT(HE_TYPE,1) = '+' "
If Not IsMissing(xTermSEQ) Then SQLQ = SQLQ & " AND Term_ENTHRS.TERM_SEQ = " & xTermSEQ & " "

gdbAdoIhr001.Execute SQLQ

gdbAdoIhr001.Execute "Update Term_ENTHRS SET HE_ENTITLE=0 WHERE HE_ENTITLE IS NULL"
'gdbAdoIhr001.CommitTrans

'TAKEN & (-)
'If there are XXX- coded records then update HE_TAKEN with those values
'gdbAdoIhr001.BeginTrans
    
'Alternatively
SQLQ = "UPDATE  Term_ENTHRS "
'Ticket #18559 - Additional logic to FLEX logic - Multiple Codes taking out from one Bank
'SQLQ = SQLQ & " SET HE_TAKEN = (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE (AD_DOA BETWEEN HE_FDATE And HE_TDATE) AND AD_REASON = LEFT(HE_TYPE,LEN(HE_TYPE)-1) + '-' AND AD_EMPNBR= HE_EMPNBR)"
SQLQ = SQLQ & " SET HE_TAKEN = (SELECT SUM(AD_HRS) FROM Term_ATTENDANCE WHERE (AD_DOA BETWEEN HE_FDATE And HE_TDATE) AND (AD_REASON = LEFT(HE_TYPE,LEN(HE_TYPE)-1) + '-' OR AD_REASON like LEFT(HE_TYPE,2) + '-%') AND Term_ENTHRS.TERM_SEQ= Term_ATTENDANCE.TERM_SEQ)"
SQLQ = SQLQ & " WHERE RIGHT(HE_TYPE,1) = '+'"  '+ because HE_TYPE will always have '+', only Attendance will have AD_REASON with '-'
If Not IsMissing(xTermSEQ) Then SQLQ = SQLQ & " AND Term_ENTHRS.TERM_SEQ = " & xTermSEQ & " "

gdbAdoIhr001.Execute SQLQ

'Also the original Logic for TAKEN calculation without '-' suffix is needed
SQLQ = "UPDATE  Term_ENTHRS "
SQLQ = SQLQ & " SET HE_TAKEN = (SELECT SUM(AD_HRS) FROM Term_ATTENDANCE WHERE (AD_DOA BETWEEN HE_FDATE And HE_TDATE) AND AD_REASON= HE_TYPE AND Term_ENTHRS.TERM_SEQ= Term_ATTENDANCE.TERM_SEQ)"
SQLQ = SQLQ & " WHERE RIGHT(HE_TYPE,1) <> '+'" '+ because HE_TYPE will always have '+', only Attendance will have AD_REASON with '-'
If Not IsMissing(xTermSEQ) Then SQLQ = SQLQ & " AND Term_ENTHRS.TERM_SEQ = " & xTermSEQ & " "

gdbAdoIhr001.Execute SQLQ


gdbAdoIhr001.Execute "Update Term_ENTHRS SET HE_TAKEN=0 WHERE HE_TAKEN IS NULL"
gdbAdoIhr001.Execute "Update Term_ENTHRS SET HE_PREV=0 WHERE HE_PREV IS NULL"
'gdbAdoIhr001.CommitTrans
'Ticket #17924 - End

MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""

Exit Sub

ErrorHandler:
glbFrmCaption$ = "Hourly Entitlement Recalculation"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EntReCalcHrTerm", "", "EntReCalcHrTerm")
If gintRollBack% = False Then
    Resume Next
End If

End Sub

Sub EntReCalcHr(Optional xEmpnbr)
Dim SQLQ

On Error GoTo ErrorHandler

If glbVadim Then    'And glbLambton Then
    Call Vadim_EntReCalcHr
    Exit Sub
End If

MDIMain.panHelp(0).FloodType = 1            '28July99 js
MDIMain.panHelp(1).Caption = " Please Wait" '
MDIMain.panHelp(2).Caption = ""             '
MDIMain.panHelp(0).FloodPercent = 10

gdbAdoIhr001.BeginTrans
If glbOracle Then
    SQLQ = "Update HRENTHRS"
    SQLQ = SQLQ & " SET HRENTHRS.HE_TAKEN = 0,"
    SQLQ = SQLQ & " HRENTHRS.HE_DHRS = (SELECT SUM(JH_DHRS) FROM qry_Assigned_jobs"
    SQLQ = SQLQ & " WHERE ED_EMPNBR=HE_EMPNBR)"
    
    If Not IsMissing(xEmpnbr) Then
        SQLQ = SQLQ & " WHERE HE_EMPNBR=" & xEmpnbr
    End If
    
    gdbAdoIhr001.Execute SQLQ
ElseIf glbSQL Then
    SQLQ = "Update HRENTHRS"
    SQLQ = SQLQ & " SET HRENTHRS.HE_TAKEN = 0,"
    SQLQ = SQLQ & " HRENTHRS.HE_DHRS = (SELECT TOP 1 JH_DHRS FROM qry_Assigned_Jobs"
    SQLQ = SQLQ & " WHERE ED_EMPNBR=HE_EMPNBR)"
    
    If Not IsMissing(xEmpnbr) Then
        SQLQ = SQLQ & " WHERE HE_EMPNBR=" & xEmpnbr
    End If
    
    gdbAdoIhr001.Execute SQLQ
Else
    gdbAdoIhr001.Execute "qry_HrEntitle"
End If

MDIMain.panHelp(0).FloodPercent = 30

If glbOracle Then
    SQLQ = " Update HRENTHRS "
    SQLQ = SQLQ & "SET HE_TAKEN = HE_TAKEN + (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE (AD_DOA BETWEEN HE_FDATE And HE_TDATE) AND AD_REASON= HE_TYPE AND AD_EMPNBR= HE_EMPNBR)"
    If Not IsMissing(xEmpnbr) Then
        SQLQ = SQLQ & " WHERE HE_EMPNBR=" & xEmpnbr
    End If
     
     gdbAdoIhr001.Execute SQLQ
ElseIf glbSQL Then
    SQLQ = "UPDATE  HRENTHRS "
    SQLQ = SQLQ & " SET HE_TAKEN = HE_TAKEN + (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE (AD_DOA BETWEEN HE_FDATE And HE_TDATE) AND AD_REASON= HE_TYPE AND AD_EMPNBR= HE_EMPNBR)"
    SQLQ = SQLQ & " WHERE HE_EMPNBR IN (SELECT AD_EMPNBR FROM HR_ATTENDANCE,HRENTHRS WHERE (AD_DOA BETWEEN HE_FDATE And HE_TDATE) AND AD_REASON= HE_TYPE)"
    If Not IsMissing(xEmpnbr) Then
        SQLQ = SQLQ & " AND HE_EMPNBR=" & xEmpnbr
    End If
    
    gdbAdoIhr001.Execute SQLQ
Else
    SQLQ = "UPDATE  HRENTHRS LEFT JOIN HR_ATTENDANCE ON (HRENTHRS.HE_TYPE = HR_ATTENDANCE.AD_REASON) AND "
    SQLQ = SQLQ & "(HRENTHRS.HE_EMPNBR = HR_ATTENDANCE.AD_EMPNBR) SET HRENTHRS.HE_TAKEN = [HRENTHRS].[HE_TAKEN]+[HR_ATTENDANCE].[AD_HRS] "
    SQLQ = SQLQ & "WHERE (((HR_ATTENDANCE.AD_DOA)>=[HRENTHRS].[HE_FDATE] And (HR_ATTENDANCE.AD_DOA)<=[HRENTHRS].[HE_TDATE]) AND "
    SQLQ = SQLQ & "((HR_ATTENDANCE.AD_REASON)=[HRENTHRS].[HE_TYPE]) AND ((HR_ATTENDANCE.AD_EMPNBR)=[HRENTHRS].[HE_EMPNBR]))"
    gdbAdoIhr001.Execute SQLQ
End If
gdbAdoIhr001.Execute "Update HRENTHRS SET HE_TAKEN=0 WHERE HE_TAKEN IS NULL"
gdbAdoIhr001.CommitTrans
    
DoEvents

'Ticket #17924 - Begin
'New Logic using XXX+ for Entitlement and XXX- for Taken
'ENTITLEMENT (+)
'If there are XXX+ coded records then update HE_ENTITLE with those values
'Note: Entitlement update from Hourly Entitlement Master screen on a RIGHT(HE_TYPE,1) = '+' will create a new record
'HR_ATTENDANCE with AD_REASON = HE_TYPE, AD_HRS = HE_ENTITLE, AD_DOA = HE_FDATE
gdbAdoIhr001.BeginTrans
If glbSQL Then
    'SQLQ = "UPDATE  HRENTHRS "
    'SQLQ = SQLQ & " SET HE_ENTITLE = HE_ENTITLE + (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE (AD_DOA BETWEEN HE_FDATE And HE_TDATE) AND AD_REASON = LEFT(HE_TYPE,3) + '+' AND AD_EMPNBR= HE_EMPNBR)"
    'SQLQ = SQLQ & " WHERE HE_EMPNBR IN (SELECT AD_EMPNBR FROM HR_ATTENDANCE,HRENTHRS WHERE (AD_DOA BETWEEN HE_FDATE And HE_TDATE) AND LEFT(AD_REASON,3) + '+' = LEFT(HE_TYPE,3) + '+')"

    'Alternatively
    SQLQ = "UPDATE  HRENTHRS "
    SQLQ = SQLQ & " SET HE_ENTITLE = (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE (AD_DOA BETWEEN HE_FDATE And HE_TDATE) AND AD_REASON = HE_TYPE AND AD_EMPNBR= HE_EMPNBR)"
    SQLQ = SQLQ & " WHERE RIGHT(HE_TYPE,1) = '+'"
    If Not IsMissing(xEmpnbr) Then
        SQLQ = SQLQ & " AND HE_EMPNBR=" & xEmpnbr
    End If

    gdbAdoIhr001.Execute SQLQ
Else
'    'SQLQ = "UPDATE  HRENTHRS LEFT JOIN HR_ATTENDANCE ON (LEFT(HRENTHRS.HE_TYPE,3) + '+' = HR_ATTENDANCE.AD_REASON) AND "
'    'SQLQ = SQLQ & "(HRENTHRS.HE_EMPNBR = HR_ATTENDANCE.AD_EMPNBR) SET HRENTHRS.HE_ENTITLE = [HRENTHRS].[HE_ENTITLE]+[HR_ATTENDANCE].[AD_HRS] "
'    'SQLQ = SQLQ & "WHERE (((HR_ATTENDANCE.AD_DOA)>=[HRENTHRS].[HE_FDATE] And (HR_ATTENDANCE.AD_DOA)<=[HRENTHRS].[HE_TDATE]) AND "
'    'SQLQ = SQLQ & "((HR_ATTENDANCE.AD_REASON)=LEFT([HRENTHRS].[HE_TYPE],3) + '+') AND ((HR_ATTENDANCE.AD_EMPNBR)=[HRENTHRS].[HE_EMPNBR]))"
'
'    'Alternatively - gives an error - but in any case we are not doing this for MS Access version.
'    SQLQ = "UPDATE  HRENTHRS "
'    SQLQ = SQLQ & " SET HE_ENTITLE = (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE (AD_DOA BETWEEN HE_FDATE And HE_TDATE) AND AD_REASON = HE_TYPE AND AD_EMPNBR= HE_EMPNBR)"
'    SQLQ = SQLQ & " WHERE RIGHT(HE_TYPE,1) = '+'"
'
'    gdbAdoIhr001.Execute SQLQ
End If
gdbAdoIhr001.Execute "Update HRENTHRS SET HE_ENTITLE=0 WHERE HE_ENTITLE IS NULL"
gdbAdoIhr001.CommitTrans

DoEvents

'TAKEN & (-)
'If there are XXX- coded records then update HE_TAKEN with those values
gdbAdoIhr001.BeginTrans
If glbSQL Then
    'SQLQ = "UPDATE  HRENTHRS "
    'SQLQ = SQLQ & " SET HE_TAKEN = HE_TAKEN + (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE (AD_DOA BETWEEN HE_FDATE And HE_TDATE) AND AD_REASON = LEFT(HE_TYPE,3) + '-' AND AD_EMPNBR= HE_EMPNBR)"
    'SQLQ = SQLQ & " WHERE HE_EMPNBR IN (SELECT AD_EMPNBR FROM HR_ATTENDANCE,HRENTHRS WHERE (AD_DOA BETWEEN HE_FDATE And HE_TDATE) AND AD_REASON = LEFT(HE_TYPE,3) + '-')"

    'Alternatively
    SQLQ = "UPDATE  HRENTHRS "
    'Ticket #18559 - Additional logic to FLEX logic - Multiple Codes taking out from one Bank
    'SQLQ = SQLQ & " SET HE_TAKEN = (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE (AD_DOA BETWEEN HE_FDATE And HE_TDATE) AND AD_REASON = LEFT(HE_TYPE,LEN(HE_TYPE)-1) + '-' AND AD_EMPNBR= HE_EMPNBR)"
    SQLQ = SQLQ & " SET HE_TAKEN = (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE (AD_DOA BETWEEN HE_FDATE And HE_TDATE) AND (AD_REASON = LEFT(HE_TYPE,LEN(HE_TYPE)-1) + '-' OR AD_REASON like LEFT(HE_TYPE,2) + '-%') AND AD_EMPNBR= HE_EMPNBR)"
    SQLQ = SQLQ & " WHERE RIGHT(HE_TYPE,1) = '+'"  '+ because HE_TYPE will always have '+', only Attendance will have AD_REASON with '-'

    If Not IsMissing(xEmpnbr) Then
        SQLQ = SQLQ & " AND HE_EMPNBR=" & xEmpnbr
    End If

    gdbAdoIhr001.Execute SQLQ

    'Also the original Logic for TAKEN calculation without '-' suffix is needed
    SQLQ = "UPDATE  HRENTHRS "
    SQLQ = SQLQ & " SET HE_TAKEN = (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE (AD_DOA BETWEEN HE_FDATE And HE_TDATE) AND AD_REASON= HE_TYPE AND AD_EMPNBR= HE_EMPNBR)"
    SQLQ = SQLQ & " WHERE RIGHT(HE_TYPE,1) <> '+'" '+ because HE_TYPE will always have '+', only Attendance will have AD_REASON with '-'
    
    If Not IsMissing(xEmpnbr) Then
        SQLQ = SQLQ & " AND HE_EMPNBR=" & xEmpnbr
    End If
    
    gdbAdoIhr001.Execute SQLQ

Else
'    'SQLQ = "UPDATE  HRENTHRS LEFT JOIN HR_ATTENDANCE ON (LEFT(HRENTHRS.HE_TYPE,3) + '-' = HR_ATTENDANCE.AD_REASON) AND "
'    'SQLQ = SQLQ & "(HRENTHRS.HE_EMPNBR = HR_ATTENDANCE.AD_EMPNBR) SET HRENTHRS.HE_TAKEN = [HRENTHRS].[HE_TAKEN]+[HR_ATTENDANCE].[AD_HRS] "
'    'SQLQ = SQLQ & "WHERE (((HR_ATTENDANCE.AD_DOA)>=[HRENTHRS].[HE_FDATE] And (HR_ATTENDANCE.AD_DOA)<=[HRENTHRS].[HE_TDATE]) AND "
'    'SQLQ = SQLQ & "((HR_ATTENDANCE.AD_REASON)=LEFT([HRENTHRS].[HE_TYPE],3) + '-') AND ((HR_ATTENDANCE.AD_EMPNBR)=[HRENTHRS].[HE_EMPNBR]))"
'
'   'Alternatively  - gives an error - but in any case we are not doing this for MS Access version.
'    SQLQ = "UPDATE  HRENTHRS "
'    SQLQ = SQLQ & " SET HE_TAKEN = (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE (AD_DOA BETWEEN HE_FDATE And HE_TDATE) AND AD_REASON = LEFT(HE_TYPE,LEN(HE_TYPE)-1) + '-' AND AD_EMPNBR= HE_EMPNBR)"
'    SQLQ = SQLQ & " WHERE RIGHT(HE_TYPE,1) = '+'"  '+ because HE_TYPE will always have '+', only Attendance will have AD_REASON with '-'
'
'    gdbAdoIhr001.Execute SQLQ
'
'    'Also the original Logic for TAKEN calculation without '-' suffix is needed
'    SQLQ = "UPDATE  HRENTHRS "
'    SQLQ = SQLQ & " SET HE_TAKEN = (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE (AD_DOA BETWEEN HE_FDATE And HE_TDATE) AND AD_REASON= HE_TYPE AND AD_EMPNBR= HE_EMPNBR)"
'    SQLQ = SQLQ & " WHERE RIGHT(HE_TYPE,1) <> '+'" '+ because HE_TYPE will always have '+', only Attendance will have AD_REASON with '-'
'    gdbAdoIhr001.Execute SQLQ

End If
gdbAdoIhr001.Execute "Update HRENTHRS SET HE_TAKEN=0 WHERE HE_TAKEN IS NULL"
gdbAdoIhr001.Execute "Update HRENTHRS SET HE_PREV=0 WHERE HE_PREV IS NULL"
gdbAdoIhr001.CommitTrans
'Ticket #17924 - End

DoEvents

MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""

Exit Sub

ErrorHandler:
glbFrmCaption$ = "Hourly Entitlement Recalculation"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EntReCalcHr", "", "qry_HrEntitle1/qry_HrEntitle")
If gintRollBack% = False Then
    Resume Next
End If

End Sub

'Sub Get_Code(TabNam, Captin, Optional Multiline As Boolean)
'Dim SQLQ As String
'On Error GoTo GCode_Err
'If Not gSec_Inq_Master_Table(TabNam) Then
'    MsgBox "You Do Not Have Authority For This Transaction"
'    Exit Sub
'End If
'glbTabNam = TabNam
'Load frmMTABL
'frmMTABL.vbxTrueGrid.MultiSelect = IIf(Multiline, 2, 1)
'frmMTABL.Caption = Captin
'
'If TabNam = "ALL" Then ' display for all
'    SQLQ = "SELECT * FROM HRTABL WHERE NOT (TB_NAME = 'EDOR' and TB_KEY = '-NON') "
'    SQLQ = SQLQ & " ORDER BY TB_NAME, TB_DESC, TB_KEY;"
'Else
'    SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = '" & TabNam & "'"
'    If TabNam = "EDOR" Then
'        If glbUnionForm Then
'            SQLQ = SQLQ & " AND " & glbSeleUnion
'        Else
'            SQLQ = SQLQ & " AND " & glbSeleUnion & " AND TB_KEY <> '-NON' "
'        End If
'
'    End If
'    If TabNam = "EDSE" Then
'        SQLQ = SQLQ & " AND " & glbSeleSection
'    End If
'    SQLQ = SQLQ & " ORDER BY TB_DESC, TB_KEY"
'End If
'
'
'
'frmMTABL.Data1.ConnectionString = glbAdoIHRDB
'frmMTABL.Data1.RecordSource = SQLQ
'frmMTABL.Data1.Refresh
'If (frmMTABL.Data1.Recordset.EOF Or frmMTABL.Data1.Recordset.BOF) Then
'    frmMTABL.cmdModify.Enabled = False
'    frmMTABL.cmdDelete.Enabled = False
'End If
'frmMTABL.Show 1
'
'Exit Sub
'
'
'GCode_Err:
'If Err.Number = 5 Then
'    MsgBox "You Do Not Have Authority For This Transaction"
'    Exit Sub
'End If
'
'glbFrmCaption$ = "GET CODE PROC"
'glbErrNum& = Err
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Get_Code", "TABL", "SELECT")
'If gintRollBack% = False Then
'    Resume Next
'End If
'
'End Sub

Sub Get_Code_Normal(TabNam, Captin, zDiv As String)
Dim SQLQ As String
On Error GoTo GCode_Err
If Not gSec_Inq_Master_Table(TabNam) Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If
glbTabNam = TabNam
glbTransDiv = zDiv
frmTABLMASTER.Caption = Captin
frmTABLMASTER.Show 1

Exit Sub


GCode_Err:
If Err.Number = 5 Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

glbFrmCaption$ = "GET CODE PROC"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Get_Code", "TABL", "SELECT")
If gintRollBack% = False Then
    Resume Next
End If

End Sub
Sub Get_Code_Linamar(TabNam, Captin, zDiv As String)
Dim SQLQ As String
On Error GoTo GCode_Err
If Not gSec_Inq_Master_Table(TabNam) Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If
glbTabNam = TabNam
glbTransDiv = zDiv
frmMTABLin.Caption = Captin
frmMTABLin.Show 1

Exit Sub


GCode_Err:
If Err.Number = 5 Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

glbFrmCaption$ = "GET CODE PROC"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Get_Code", "TABL", "SELECT")
If gintRollBack% = False Then
    Resume Next
End If

End Sub
Sub Get_HOME(TabNam, Captin, Master%, zDiv As String)
Dim SQLQ As String
On Error GoTo GCode_Err
If Not gSec_Inq_Master_Table(TabNam) Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If
glbTabNam = TabNam
glbTransDiv = zDiv
Load frmHomeMaster
frmHomeMaster.Caption = Captin
If Master Then
    frmHomeMaster.cmdSelect.Enabled = False
    glbHOMEInhSel% = True
Else
    frmHomeMaster.cmdSelect.Enabled = True
    glbHOMEInhSel% = False
End If
frmHomeMaster.Show 1

Exit Sub


GCode_Err:
If Err.Number = 5 Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If
glbFrmCaption$ = "GET CODE PROC"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Get_Code", "TABL", "SELECT")
If gintRollBack% = False Then
    Resume Next
End If

End Sub

Sub Get_Dept(Master%)
Dim SQLQ As String, countr As Integer

On Error GoTo GDept_Err

Load frmDEPTS

SQLQ = "SELECT * FROM HRDEPT "
If Not Master Then
    SQLQ = SQLQ & " Where " & glbSeleDept
End If

SQLQ = SQLQ & " ORDER BY DF_NAME "

'frmDEPTS.Data1.DatabaseName = glbIHRDB
frmDEPTS.Data1.ConnectionString = glbAdoIHRDB
frmDEPTS.Data1.RecordSource = SQLQ
frmDEPTS.Data1.Refresh
If Master Then
    frmDEPTS.Caption = frmDEPTS.Caption & " Master"
    frmDEPTS.cmdSelect.Enabled = False
    glbDeptInhSel% = True
Else
    frmDEPTS.cmdSelect.Enabled = True
    glbDeptInhSel% = False
End If
frmDEPTS.Show 1
Exit Sub

GDept_Err:
glbFrmCaption$ = "Get Dept Proc"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Get_Dept", "DEPT", "SELECT")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Sub

Sub Get_OHRSDept(Master%)
Dim SQLQ As String, countr As Integer

On Error GoTo GOHRSDept_Err

Load frmOHRSDEPTS

SQLQ = "SELECT * FROM HR_OHRSDEPT "
'If Not Master Then
'    SQLQ = SQLQ & " Where " & glbSeleDept
'End If
SQLQ = SQLQ & " ORDER BY OH_NAME "

'frmOHRSDEPTS.Data1.DatabaseName = glbIHRDB
frmOHRSDEPTS.Data1.ConnectionString = glbAdoIHRDB
frmOHRSDEPTS.Data1.RecordSource = SQLQ
frmOHRSDEPTS.Data1.Refresh
If Master Then
    frmOHRSDEPTS.Caption = frmOHRSDEPTS.Caption & " Master"
    frmOHRSDEPTS.cmdSelect.Enabled = False
    glbOHRSDeptInhSel = True
Else
    frmOHRSDEPTS.cmdSelect.Enabled = True
    glbOHRSDeptInhSel = False
End If

frmOHRSDEPTS.Show 1

Exit Sub

GOHRSDept_Err:
glbFrmCaption$ = "Get OHRS Dept Proc"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Get_OHRSDept", "HR_OHRSDEPT", "SELECT")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Sub

Sub Get_DeptBonus(Master%)
Dim SQLQ As String, countr As Integer

On Error GoTo GDept_Err

Load frmDEPTSBonus

SQLQ = "SELECT * FROM WFC_Bonus_Loc_Department "
SQLQ = SQLQ & " ORDER BY Dept_Name "

'frmDEPTS.Data1.DatabaseName = glbIHRDB
frmDEPTSBonus.Data1.ConnectionString = glbAdoIHRDB
frmDEPTSBonus.Data1.RecordSource = SQLQ
frmDEPTSBonus.Data1.Refresh
If Master Then
    frmDEPTSBonus.Caption = frmDEPTSBonus.Caption & " Master"
    frmDEPTSBonus.cmdSelect.Enabled = False
    glbDeptInhSel% = True
Else
    frmDEPTSBonus.cmdSelect.Enabled = True
    glbDeptInhSel% = False
End If
frmDEPTSBonus.Show 1
Exit Sub

GDept_Err:
glbFrmCaption$ = "Get Dept Proc"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Get_DeptBonus", "Bonus DEPT", "SELECT")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Sub

Sub Get_CourseCode(Master%, Optional xType As String)
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim xSQL As String

On Error GoTo GDiv_Err

glbCourseCodeSele = Not Master%

'SQLQ = "SELECT *  FROM HR_COURSECODE_MASTER WHERE (1=1) "
'Ticket #18210
SQLQ = "SELECT HR_COURSECODE_MASTER.*,HRTABL.TB_DESC AS COURSEDESC FROM HR_COURSECODE_MASTER, HRTABL WHERE HR_COURSECODE_MASTER.ES_CRSCODE_TABL = HRTABL.TB_NAME "
SQLQ = SQLQ & "AND HRTABL.TB_NAME = 'ESCD' AND HR_COURSECODE_MASTER.ES_CRSCODE = HRTABL.TB_KEY "

If Len(xType) > 0 Then 'Ticket #13520
    xSQL = "SELECT *  FROM HR_COURSECODE_MASTER WHERE ES_CTYPE = '" & xType & "' "
    rsTemp.Open xSQL, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        SQLQ = SQLQ & "AND ES_CTYPE = '" & xType & "' "
    End If
    rsTemp.Close
End If
If glbCourseCodeSele Then
    SQLQ = SQLQ & "AND NOT (ES_STATUS = 0) "
    If Not glbDeptAllRight Then
        SQLQ = SQLQ & "AND (ES_CORPONLY = 0) "
    End If
End If
SQLQ = SQLQ & " ORDER BY ES_CRSCODE "

frmMCourseCode.Data1.ConnectionString = glbAdoIHRDB
frmMCourseCode.Data1.RecordSource = SQLQ
frmMCourseCode.Data1.Refresh
'Display_Value

If Master Then
    frmMCourseCode.Caption = frmMCourseCode.Caption & " Master"
    frmMCourseCode.cmdSelect.Enabled = False
    'glbDivInhSel% = True
Else
    frmMCourseCode.cmdSelect.Enabled = True
    'glbDivInhSel% = False
End If

frmMCourseCode.Show 1
Exit Sub

GDiv_Err:
glbFrmCaption$ = lStr("Get Course Code Proc")
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Get_CourseCode", "HR_COURSECODE_MASTER", "SELECT")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Sub

Sub Get_JobMaster(Master%) 'Ticket #25911 Franks 09/24/2014.
Dim SQLQ As String
Dim OLDid

On Error GoTo GDiv_Err

'Load frmMJobMaster

SQLQ = "SELECT * FROM HRJOBMASTER ORDER BY JB_JOBDESCR "

'frmMJobMaster.Data1.DatabaseName = glbIHRDB
frmMJobMaster.Data1.ConnectionString = glbAdoIHRDB
frmMJobMaster.Data1.RecordSource = SQLQ
frmMJobMaster.Data1.Refresh

If Master Then
    'frmMJobMaster.Caption = frmMJobMaster.Caption & " Master"   'Serbo
    frmMJobMaster.cmdSelect.Enabled = False
    glbJobMasterInhSel% = True
Else
    frmMJobMaster.cmdSelect.Enabled = True
    glbJobMasterInhSel% = False
End If

OLDid = glbJobMaster
frmMJobMaster.Show 1
If Not Master Then
    If glbJobMaster <> OLDid Then
        ''Call ReDisplayForms(RelatePOS)
        'Unload frmMJobMasterMain
        Call ReDisplayForms(RelateJobMaster)
    End If
End If

Exit Sub

GDiv_Err:
glbFrmCaption$ = lStr("Get Job Master Proc")
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Get_JobMaster", "HRJOBMASTER", "SELECT")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Sub

Sub Get_Div(Master%)
Dim SQLQ As String

On Error GoTo GDiv_Err

'Load frmDIVISIONS

SQLQ = "SELECT * FROM HR_DIVISION "
SQLQ = SQLQ & " WHERE " & glbSeleDiv
SQLQ = SQLQ & " ORDER BY Division_Name "

'frmDIVISIONS.Data1.DatabaseName = glbIHRDB
frmDIVISIONS.Data1.ConnectionString = glbAdoIHRDB
frmDIVISIONS.Data1.RecordSource = SQLQ
frmDIVISIONS.Data1.Refresh

If Master Then
    frmDIVISIONS.Caption = frmDIVISIONS.Caption & " Master"   'Serbo
    frmDIVISIONS.cmdSelect.Enabled = False
    glbDivInhSel% = True
Else
    frmDIVISIONS.cmdSelect.Enabled = True
    glbDivInhSel% = False
End If

frmDIVISIONS.Show 1
Exit Sub

GDiv_Err:
glbFrmCaption$ = lStr("Get Division Proc")
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Get_Div", "HR_DIVISION", "SELECT")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Sub
Sub Get_SalDist(Master%)
Dim SQLQ As String

On Error GoTo GSALSIST_Err

Load frmDIVISIONS

SQLQ = "SELECT * FROM HRSALDIST"
'SQLQ = SQLQ & " WHERE " & glbSeleDiv
SQLQ = SQLQ & " ORDER BY SD_DESC"

'frmDIVISIONS.Data1.DatabaseName = glbIHRDB
frmSalDist.Data1.ConnectionString = glbAdoIHRDB
frmSalDist.Data1.RecordSource = SQLQ
frmSalDist.Data1.Refresh

If Master Then
    frmSalDist.Caption = frmSalDist.Caption & " Master"   'Serbo
    frmSalDist.cmdSelect.Enabled = False
    glbSalDistInhSel% = True
Else
    frmSalDist.cmdSelect.Enabled = True
    glbSalDistInhSel% = False
End If

frmSalDist.Show 1
Exit Sub

GSALSIST_Err:
glbFrmCaption$ = lStr("Get Salary Distribution")
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Get_SalDist", "HRSALDIST", "SELECT")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Sub

Sub Get_PayCategory(Master%)
Dim SQLQ As String

On Error GoTo GSALSIST_Err

Load frmPayCategory

'SQLQ = "SELECT * FROM HR_PAYROLL_CATEGORY"
''SQLQ = SQLQ & " WHERE " & glbSeleDiv
'SQLQ = SQLQ & " ORDER BY PC_DESC"

'frmDIVISIONS.Data1.DatabaseName = glbIHRDB
'frmPayCategory.Data1.ConnectionString = glbAdoIHRDB
'frmPayCategory.Data1.RecordSource = SQLQ
'frmPayCategory.Data1.Refresh

frmPayCategory.Caption = frmPayCategory.Caption & " Master"   'Serbo
frmPayCategory.cmdSelect.Enabled = False
glbPayCategoryInhSel% = True

frmPayCategory.Show 1
Exit Sub

GSALSIST_Err:
glbFrmCaption$ = lStr("Get Payroll Category")
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Get_PayCategory", "HRPayCategory", "SELECT")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Sub

Sub Get_ChargeCode(Master%)
Dim SQLQ As String

On Error GoTo ChargeCode_Err

Load frmCHARGECODE

frmCHARGECODE.cmdSelect.Enabled = False

frmCHARGECODE.Show 1
Exit Sub

ChargeCode_Err:
glbFrmCaption$ = lStr("Get Charge Code")
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Get_ChargeCode", "HR_CHARGE_CODE", "SELECT")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Sub

Sub Get_ProjectCode(Master%)
Dim SQLQ As String

On Error GoTo ProjectCode_Err

Load frmProjectCode

frmProjectCode.cmdSelect.Enabled = False

frmProjectCode.Show 1
Exit Sub

ProjectCode_Err:
glbFrmCaption$ = "Get Project Code"

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Get_ProjectCode", "HR_ProjectCode", "SELECT")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Sub

Sub Get_Machine(Master%)
Dim SQLQ As String

On Error GoTo Machine_Err

Load frmMachine

frmMachine.cmdSelect.Enabled = False

frmMachine.Show 1
Exit Sub

Machine_Err:
glbFrmCaption$ = "Get Machine"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Get_Machine", "HR_Machine", "SELECT")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Sub

Sub GET_Division()
Dim OLDid, XScn
Dim tMode As RelateModeEnum
    OLDid = glbDiv
    Call Get_Div(False)
    tMode = RelateEMP
    XScn = (glbDiv <> OLDid)
    If XScn Then
        If glbLinHS Then
            glbLEE_ID = Val("999999" & glbDiv)
            Call ReDisplayForms(tMode)
        End If
    End If

End Sub

Sub GET_EMP()
Dim OLDid, XScn
Dim tMode As RelateModeEnum
On Error GoTo Err_Deal
XScn = False
If glbOnTop = "FRMEREHIRE" Or (glbtermopen And glbTermTran) Then
    OLDid = glbTERM_Seq
    frmTERMEMPL.Show 1
    tMode = RelateTermEmp
    XScn = (glbTERM_Seq <> OLDid)
ElseIf glbOnTop = "FRMETRANIN" Or (glbtermopen And Not glbTermTran) Then
    Unload frmETRANIN
    Load frmETRANIN
    frmETRANIN.ZOrder 0
Else
    OLDid = glbLEE_ID
    frmEEFIND.Show 1
    XScn = (glbLEE_ID <> OLDid)
    tMode = RelateEMP
End If
If XScn Then
    Call ReDisplayForms(tMode)
'    If MDIMain.ActiveForm Is Nothing Then
'        glbOnTop = ""
'    Else
'        glbOnTop = UCase(MDIMain.ActiveForm.name)
'    End If
'
'    Unload frmEEBASIC
'    If glbOnTop = "FRMEEBASIC" Then Load frmEEBASIC
'    Unload frmEESTATS
'    If glbOnTop = "FRMEESTATS" Then Load frmEESTATS
'    Unload frmEMERG
'    If glbOnTop = "FRMEMERG" Then Load frmEMERG
'    Unload frmDEPNDTS
'    If glbOnTop = "FRMDEPNDTS" Then Load frmDEPNDTS
'    Unload frmEBANK
'    If glbOnTop = "FRMEBANK" Then Load frmEBANK
'    Unload frmESKILLS
'    If glbOnTop = "FRMESKILLS" Then Load frmESKILLS
'    Unload frmFORMALED
'    If glbOnTop = "FRMFORMALED" Then Load frmFORMALED
'    Unload frmESEMINARS
'    If glbOnTop = "FRMESEMINARS" Then Load frmESEMINARS
'    Unload frmEASSOC
'    If glbOnTop = "FRMEASSOC" Then Load frmEASSOC
'    Unload frmEBENEFITS
'    If glbOnTop = "FRMEBENEFITS" Then Load frmEBENEFITS
'    Unload frmEODOLLAR
'    If glbOnTop = "FRMEODOLLAR" Then Load frmEODOLLAR
'    Unload frmOTHERERN
'    If glbOnTop = "FRMOTHERERN" Then Load frmOTHERERN
'    Unload frmEPOSITION
'    If glbOnTop = "FRMEPOSITION" Then Load frmEPOSITION
'    Unload frmEPERFORM
'    If glbOnTop = "FRMEPERFORM" Then Load frmEPERFORM
'    Unload frmESALARY
'    If glbOnTop = "FRMESALARY" Then Load frmESALARY
'    Unload frmVATTEND
'    If glbOnTop = "FRMVATTEND" Then Load frmVATTEND
'    Unload frmVACSICK
'    If glbOnTop = "FRMVACSICK" Then Load frmVACSICK
'    Unload frmVACSICKO
'    If glbOnTop = "FRMVACSICKO" Then Load frmVACSICKO
'    Unload frmEHSINCIDENT
'    If glbOnTop = "FRMEHSINCIDENT" Then Load frmEHSINCIDENT
'    Unload frmECOMMENTS
'    If glbOnTop = "FRMECOMMENTS" Then Load frmECOMMENTS
'    Unload frmEFOLLOWUP
'    If Not glbtermopen And glbOnTop = "FRMEFOLLOWUP" Then Load frmEFOLLOWUP
'    Unload frmEHSWCB
'    If glbOnTop = "FRMEHSWCB" Then Load frmEHSWCB
'    Unload frmEHSWCBC
'    If glbOnTop = "FRMEHSWCBC" Then Load frmEHSWCBC
'    Unload frmEHSINJURY
'    If glbOnTop = "FRMEHSINJURY" Then Load frmEHSINJURY
'    Unload frmHrEnt
'    If glbOnTop = "FRMHRENT" Then Load frmHrEnt
'    Unload frmETERM
'    If glbOnTop = "FRMETERM" Then Load frmETERM
'    Unload frmEHSCause
'    If glbOnTop = "FRMEHSCAUSE" Then Load frmEHSCause
'    Unload frmEHSCorrective
'    If glbOnTop = "FRMEHSCORRECTIVE" Then Load frmEHSCorrective
'    Unload frmEHSContact
'    If glbOnTop = "FRMEHSCONTACT" Then Load frmEHSContact
'    Unload frmCobra
'    If glbOnTop = "FRMCOBRA" Then Load frmCobra
'
'    Unload frmEREHIRE
'    If glbOnTop = "FRMEREHIRE" Then Load frmEREHIRE
'
    Unload frmETRANIN
    If glbOnTop = "FRMETRANIN" Then Load frmETRANIN
    Unload frmEREHIRE
    'If glbOnTop = "FRMEREHIRE" Then Load frmEREHIRE

    '
'    If glbtermopen Then Unload frmvFOLOWUP
'    Unload frmEComPlan
'    If glbOnTop = "FRMECOMPLAN" And Not glbtermopen Then Load frmEComPlan
''    Unload frmAXXRSP
''    If glbOnTop = "frmAXXRSP" Then Load frmAXXRSP
'    Unload frmECounsel
'    If glbOnTop = "FRMECOUNSEL" Then Load frmECounsel
'    Unload frmETLAY
'    If glbOnTop = "FRMETLAY" Then Load frmETLAY
End If
Exit Sub

Err_Deal:
If Err = 364 Then Resume Next

End Sub



Sub GET_JOB()
Dim OLDid
OLDid = glbJob
frmJOBS.Show 1
'If glbJOB <> OLDid Then
'    Unload frmMPOSITIONS
'    If glbOnTop = "FRMMPOSITIONS" Then Load frmMPOSITIONS
'End If
End Sub

Sub Get_JobFamily(Master%, xType, Optional xParentCode) 'Ticket #26233 Franks 11/21/2014 VitalAire Canada Inc.
Dim SQLQ As String
Dim xCaption As String
On Error GoTo GLgr_Err

'Load frmJobFamily

frmJobFamily.LinkItem = xType
SQLQ = "SELECT * FROM HRJOBFAMILY WHERE JB_TYPE = '" & xType & "' "
If Not IsMissing(xParentCode) Then
    If Len(xParentCode) > 0 Then
        SQLQ = SQLQ & "AND JB_PARENTCODE = '" & xParentCode & "' "
        frmJobFamily.locParentCode = xParentCode
    End If
End If
SQLQ = SQLQ & " ORDER BY JB_DESCR "

frmJobFamily.Data1.ConnectionString = glbAdoIHRDB
frmJobFamily.Data1.RecordSource = SQLQ
frmJobFamily.Data1.Refresh

If frmJobFamily.Data1.Recordset.BOF And frmJobFamily.Data1.Recordset.EOF Then
    frmJobFamily.cmdModify.Enabled = False
    frmJobFamily.cmdDelete.Enabled = False
End If

If xType = "JOBFAMILY" Then xCaption = "Job Family"
If xType = "SUBFAMILY" Then xCaption = "Sub-Job Family"
If xType = "GROUPJOBS" Then xCaption = "Group Jobs"

If Master Then
    xCaption = xCaption & " Master"
    frmJobFamily.cmdSelect.Enabled = False
    glbLgrInhSel% = True
Else
    glbLgrInhSel% = False
    frmJobFamily.cmdSelect.Enabled = True
End If
frmJobFamily.Caption = xCaption
'frmJobFamily.LinkItem = xType
frmJobFamily.Show 1

Exit Sub

GLgr_Err:
glbFrmCaption$ = "Get JobFamily Proc"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Get_JobFamily", "HRJOBFAMILY", "SELECT")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Sub


Sub Get_Ledgers(Master%)
Dim SQLQ As String

On Error GoTo GLgr_Err

Load frmLEDGER

SQLQ = "SELECT * FROM HRGL "
SQLQ = SQLQ & " ORDER BY GL_DESCR "

'frmLEDGER.Data1.DatabaseName = glbIHRDB
frmLEDGER.Data1.ConnectionString = glbAdoIHRDB
frmLEDGER.Data1.RecordSource = SQLQ
frmLEDGER.Data1.Refresh

If frmLEDGER.Data1.Recordset.BOF And frmLEDGER.Data1.Recordset.EOF Then
    frmLEDGER.cmdModify.Enabled = False
    frmLEDGER.cmdDelete.Enabled = False

End If
If Master Then
frmLEDGER.Caption = frmLEDGER.Caption & " Master"
    frmLEDGER.cmdSelect.Enabled = False
    glbLgrInhSel% = True
Else
    glbLgrInhSel% = False
    frmLEDGER.cmdSelect.Enabled = True
End If
frmLEDGER.Show 1

Exit Sub

GLgr_Err:
glbFrmCaption$ = "Get Ledger Proc"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Get_Lgr", "HRGL", "SELECT")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If


End Sub

Sub Get_Attendance_Group_Master_Code()


Dim SQLQ As String
Dim rsSR As New ADODB.Recordset
On Error GoTo GMCode_Err

'Load frmTABLMASTER

'rsSR.Open "SELECT * FROM HRTABL WHERE TB_NAME NOT IN (SELECT CODENAME FROM HR_SECURE_ACCESS WHERE USERID='" & glbUserID & "' AND CODENAME IS NOT NULL AND ACCESSABLE<>0) ", gdbAdoIhr001, adOpenKeyset
'If Not rsSR.EOF Then
'    MsgBox "You Do Not Have Authority For This Transaction"
'    Exit Sub
'End If
'rsSR.Close
SQLQ = "SELECT * FROM HRTABL WHERE NOT(TB_NAME = 'EDOR' and TB_KEY = '-NON' and TB_KEY = '-EXE') "     'Hemu -EXE
If glbLinamar Then
    SQLQ = SQLQ & " AND TB_NAME<>'EDSE' AND TB_NAME<>'EDRG' AND TB_NAME <> 'BNCD'"
End If
SQLQ = SQLQ & " AND TB_NAME IN(SELECT CODENAME FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbUserID, "'", "''") & "' AND CODENAME IS NOT NULL AND ACCESSABLE<>0) "
SQLQ = SQLQ & " ORDER BY TB_NAME, TB_KEY"

'Mostafa
SQLQ = "SELECT * FROM HRATTGRP"

frmTABLATTGroupMASTER.Data1.RecordSource = SQLQ
frmTABLATTGroupMASTER.Data1.Refresh
frmTABLATTGroupMASTER.Show 1

Exit Sub
GMCode_Err:
glbFrmCaption$ = "GET CODE PROC"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Get_Group_Master_Code", "TABL", "SELECT")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Sub

Sub Get_Master_Code()
Dim SQLQ As String
Dim rsSR As New ADODB.Recordset
On Error GoTo GMCode_Err

'Load frmTABLMASTER

'rsSR.Open "SELECT * FROM HRTABL WHERE TB_NAME NOT IN (SELECT CODENAME FROM HR_SECURE_ACCESS WHERE USERID='" & glbUserID & "' AND CODENAME IS NOT NULL AND ACCESSABLE<>0) ", gdbAdoIhr001, adOpenKeyset
'If Not rsSR.EOF Then
'    MsgBox "You Do Not Have Authority For This Transaction"
'    Exit Sub
'End If
'rsSR.Close
SQLQ = "SELECT * FROM HRTABL WHERE NOT(TB_NAME = 'EDOR' and TB_KEY = '-NON' and TB_KEY = '-EXE') "     'Hemu -EXE
If glbLinamar Then
    SQLQ = SQLQ & " AND TB_NAME<>'EDSE' AND TB_NAME<>'EDRG' AND TB_NAME <> 'BNCD'"
End If
SQLQ = SQLQ & " AND TB_NAME IN(SELECT CODENAME FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbUserID, "'", "''") & "' AND CODENAME IS NOT NULL AND ACCESSABLE<>0) "
SQLQ = SQLQ & " ORDER BY TB_NAME, TB_KEY"

glbTabNam = ""
glbTransDiv = ""
frmTABLMASTER.Data1.RecordSource = SQLQ
frmTABLMASTER.Data1.Refresh
frmTABLMASTER.cmdSelect.Enabled = False
frmTABLMASTER.Show 1

Exit Sub

GMCode_Err:
glbFrmCaption$ = "GET CODE PROC"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Get_Master_Code", "TABL", "SELECT")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Sub

'this function takes a sql statement and returns the value of a column of the first row
Function GetValue(sql As String, colname As String) As String
    Dim rsDATA As New ADODB.Recordset
    rsDATA.Open sql, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
     
End Function

Sub Get_Positions()

Dim SQLQ As String
On Error GoTo GJOB_Err

Load frmJOBS

SQLQ = "SELECT * FROM HRJOB"
frmJOBS.Data1.RecordSource = SQLQ
frmJOBS.Data1.Refresh
frmJOBS.Show 1
Exit Sub


GJOB_Err:
glbFrmCaption$ = "Get Positions Proc"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Get_Positions", "Job", "SELECT")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Sub

Sub Get_Prov(Master%)

Dim SQLQ As String
On Error GoTo GProv_Err

Load frmPROV

SQLQ = "SELECT * FROM HRPROV"
frmPROV.Data1.RecordSource = SQLQ
frmPROV.Data1.Refresh
frmPROV.Show 1

frmPROV.cmdModify.Enabled = True        '09June99 js
frmPROV.cmdNew.Enabled = True           '
frmPROV.cmdDelete.Enabled = True        '
frmPROV.cmdFind.Enabled = True          '
frmPROV.cmdOk.Enabled = False           '

If Master Then
    frmPROV.cmdSelect.Enabled = False
End If

Exit Sub

GProv_Err:
glbFrmCaption$ = "Get Province Proc"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Get_Prov", "Prov", "SELECT")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Sub


Sub glbCri_DeptUN(txtDeptCD As String)
Dim DeptUnCri As String, SQLQ
Dim xDept, xUnion, xDiv, xSECTION, xAdminBy, xLoc, xRegion, xSupCode, xVadim2
Dim DeptUn_Snap As New ADODB.Recordset
Dim RecUnion As New ADODB.Recordset
Dim xInclEmp, xExclEmp

If gSec_Emp_Based Then
    DeptUnCri = DeptUnCri & "({HREMP.ED_EMPNBR}=" & glbEmpNbr & ") "
Else
    SQLQ = "Select HRPASDEP.* from HRPASDEP"
    SQLQ = SQLQ & " where HRPASDEP.PD_USERID = '" & Replace(glbUserID, "'", "''") & "'"
    SQLQ = SQLQ & " ORDER by PD_DEPT, PD_ORG "
    RecUnion.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    DeptUn_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic
    DeptUnCri = "("
    Do While Not DeptUn_Snap.EOF
        xDept = DeptUn_Snap("PD_DEPT")
        xUnion = DeptUn_Snap("PD_ORG")
        xDiv = DeptUn_Snap("PD_DIV")
        xSECTION = DeptUn_Snap("PD_SECTION")
        xAdminBy = DeptUn_Snap("PD_ADMINBY")
        xLoc = DeptUn_Snap("PD_LOC")
        xRegion = DeptUn_Snap("PD_REGION")
        
        'Ticket #24161 - Samuel only - Release 8.0
        If glbSamuel Then
            xSupCode = DeptUn_Snap("PD_SUPCODE")
            xVadim2 = DeptUn_Snap("PD_VADIM2")
        End If
        
        xInclEmp = DeptUn_Snap("PD_INCLEMPNBR")
        xExclEmp = DeptUn_Snap("PD_EXCLEMPNBR")
        
'Ticket #21484 - Hemu        '7.9 Enhancement
'        If Not IsNull(xInclEmp) And Len(xInclEmp) <> 0 Then
'            DeptUnCri = DeptUnCri & " ("
'        End If
        
        If xDept = "ALL" Then
            DeptUnCri = DeptUnCri & "(1=1 and "
        Else
            DeptUnCri = DeptUnCri & "(({HREMP.ED_DEPTNO}='" & xDept & "') and "
        End If
        
        If IsNull(xUnion) Then
            DeptUnCri = DeptUnCri & "1=1 and "
        Else
            If Len(xUnion) = 0 Then
                DeptUnCri = DeptUnCri & "1=1 and "
            Else

                If xUnion = "-NON" Or xUnion = "-EXE" Then     'Hemu -EXE
                
                    If InStr(glbSeleUnion, "1=1") = 0 Then

                        If RecUnion.RecordCount > 0 Then
                            DeptUnCri = DeptUnCri & "("
                            RecUnion.MoveFirst
                            Do While Not RecUnion.EOF
                                If RecUnion("PD_ORG") <> "-NON" And RecUnion("PD_ORG") <> "-EXE" Then   'Hemu -EXE
                                    DeptUnCri = DeptUnCri & " ({HREMP.ED_ORG} = '" & RecUnion("PD_ORG") & "') or "
                                End If
                                RecUnion.MoveNext
                            Loop
                            DeptUnCri = DeptUnCri & " (1=2)) and "
                        End If
                    End If
                    'DeptUnCri = DeptUnCri & "({HREMP.ED_ORG} <> 'NONE') and "
                Else
                    'If Left(xUnion, 1) = "-" Then   'Listowel
                    '    DeptUnCri = DeptUnCri & " ({HREMP.ED_ORG}<>'" & Mid(xUnion, 2, Len(xUnion)) & "') or (" & DeptUnCri & " isnull({HREMP.ED_ORG})))  and "
                    'Else
                    If glbCompSerial = "S/N - 2288W" And Left(xUnion, 1) = "-" Then 'Musashi - Ticket #12690
                        DeptUnCri = DeptUnCri & "({HREMP.ED_ORG}='" & Mid(xUnion, 2) & "') and "
                    Else
                        DeptUnCri = DeptUnCri & "({HREMP.ED_ORG}='" & xUnion & "') and "
                    End If
                    'End If
                End If
            End If
        End If
        
        If IsNull(xDiv) Then
            DeptUnCri = DeptUnCri & "1=1 and "
        Else
            If Len(xDiv) = 0 Then
                DeptUnCri = DeptUnCri & "1=1 and "
            Else
                DeptUnCri = DeptUnCri & "({HREMP.ED_DIV}='" & xDiv & "') and "
            End If
        End If
        
        'Ticket #18235
        If IsNull(xAdminBy) Then
'Ticket #21484 - Hemu            If IsNull(xInclEmp) Or Len(xInclEmp) = 0 Then
                DeptUnCri = DeptUnCri & "1=1 and "
'            Else
'                DeptUnCri = DeptUnCri & "1=1 or "
'            End If
        Else
            If Len(xAdminBy) = 0 Then
'                If Len(xInclEmp) = 0 Then
                    DeptUnCri = DeptUnCri & "1=1 and "
'                Else
'                    DeptUnCri = DeptUnCri & "1=1 or "
'                End If
            Else
'                If IsNull(xInclEmp) Or Len(xInclEmp) = 0 Then
                    DeptUnCri = DeptUnCri & "({HREMP.ED_ADMINBY}='" & xAdminBy & "') and "
'                Else
'                    DeptUnCri = DeptUnCri & "({HREMP.ED_ADMINBY}='" & xAdminBy & "') or "
'                End If
            End If
        End If
                
        'Ticket #22682 - Release 8.0
        If IsNull(xLoc) Then
            DeptUnCri = DeptUnCri & "1=1 and "
        Else
            If Len(xLoc) = 0 Then
                DeptUnCri = DeptUnCri & "1=1 and "
            Else
                DeptUnCri = DeptUnCri & "({HREMP.ED_LOC}='" & xLoc & "') and "
            End If
        End If
                
        'Ticket #22682 - Release 8.0
        If IsNull(xRegion) Then
            DeptUnCri = DeptUnCri & "1=1 and "
        Else
            If Len(xRegion) = 0 Then
                DeptUnCri = DeptUnCri & "1=1 and "
            Else
                DeptUnCri = DeptUnCri & "({HREMP.ED_REGION}='" & xRegion & "') and "
            End If
        End If
                
                
        'Ticket #24161 - Samuel only - Release 8.0
        If glbSamuel Then
            'Supervisor Code
            If IsNull(xSupCode) Then
                DeptUnCri = DeptUnCri & "1=1 and "
            Else
                If Len(xSupCode) = 0 Then
                    DeptUnCri = DeptUnCri & "1=1 and "
                Else
                    DeptUnCri = DeptUnCri & "({HREMP.ED_SUPCODE}='" & xSupCode & "') and "
                End If
            End If
            
            'Vadim Field 2
            If IsNull(xVadim2) Then
                DeptUnCri = DeptUnCri & "1=1 and "
            Else
                If Len(xVadim2) = 0 Then
                    DeptUnCri = DeptUnCri & "1=1 and "
                Else
                    DeptUnCri = DeptUnCri & "({HREMP.ED_VADIM2}='" & xVadim2 & "') and "
                End If
            End If
        End If
                
'Ticket #21484 - Hemu        '7.9 Enhancement
'        'Include Employee #s
'        If IsNull(xInclEmp) Then
'            'If IsNull(xExclEmp) Then
'                DeptUnCri = DeptUnCri & "1=1 and "
'            'Else
'            '    DeptUnCri = DeptUnCri & "1=1 or "
'            'End If
'        Else
'            If Len(xInclEmp) = 0 Then
'                'If Len(xExclEmp) = 0 Then
'                    DeptUnCri = DeptUnCri & "1=1 and "
'                'Else
'                '    DeptUnCri = DeptUnCri & "1=1 or "
'                'End If
'            Else
'                'If Len(xExclEmp) = 0 Then
'                    DeptUnCri = DeptUnCri & "({HREMP.ED_EMPNBR} IN [" & getEmpnbr(xInclEmp) & "])) and "
'                'Else
'                '    DeptUnCri = DeptUnCri & "({HREMP.ED_EMPNBR} IN [" & getEmpnbr(xInclEmp) & "]) or "
'                'End If
'            End If
'        End If
                
        '7.9 Enhancement
        'Exclude Employee #s
        If IsNull(xExclEmp) Then
            DeptUnCri = DeptUnCri & "1=1 and "
        Else
            If Len(xExclEmp) = 0 Then
                DeptUnCri = DeptUnCri & "1=1 and "
            Else
                DeptUnCri = DeptUnCri & " NOT ({HREMP.ED_EMPNBR} IN [" & getEmpnbr(xExclEmp) & "]) and "
            End If
        End If
                
                
        If IsNull(xSECTION) Then
            'Ticket #21484
            If Len(xInclEmp) > 0 Then
                DeptUnCri = DeptUnCri & "1=1 or ({HREMP.ED_EMPNBR} IN [" & getEmpnbr(xInclEmp) & "]) ) or "
            Else
                DeptUnCri = DeptUnCri & "1=1) or "
            End If
        Else
            If Len(xSECTION) = 0 Then
                'Ticket #21484
                If Len(xInclEmp) > 0 Then
                    DeptUnCri = DeptUnCri & "1=1 or ({HREMP.ED_EMPNBR} IN [" & getEmpnbr(xInclEmp) & "]) ) or "
                Else
                    DeptUnCri = DeptUnCri & "1=1) or "
                End If
            Else
                'Ticket #21484
                If Len(xInclEmp) > 0 Then
                    DeptUnCri = DeptUnCri & "({HREMP.ED_SECTION}='" & xSECTION & "') OR ({HREMP.ED_EMPNBR} IN [" & getEmpnbr(xInclEmp) & "])) or "
                Else
                    DeptUnCri = DeptUnCri & "({HREMP.ED_SECTION}='" & xSECTION & "')) or "
                End If
            End If
        End If
    
        DeptUn_Snap.MoveNext
    Loop
    DeptUnCri = DeptUnCri & " 1<>1) "
    
    'Ticket #21484 - Hemu
'    If Len(xInclEmp) > 0 Then
'        DeptUnCri = DeptUnCri & "OR ({HREMP.ED_EMPNBR} IN [" & getEmpnbr(xInclEmp) & "]) "
'    End If
    
    'Hemu - 06/02/2004 Begin
    'If Len(txtDeptCD) <> 0 Then DeptUnCri = DeptUnCri & " And {HREMP.ED_DEPTNO} = '" & txtDeptCD & "' "
    If glbSQL Or glbOracle Then
        'If Len(txtDeptCD) <> 0 Then DeptUnCri = DeptUnCri & " And {HREMP.ED_DEPTNO} IN ('" & getCodes(txtDeptCD) & "') " '[] to () changed by Bryan 05-08-05 Ticket #9063
        'Frank 06/19/2006, ticket #11014, () to []
        If Len(txtDeptCD) <> 0 Then DeptUnCri = DeptUnCri & " And {HREMP.ED_DEPTNO} IN ['" & getCodes(txtDeptCD) & "']"
    Else
        Dim strTemp As String
        strTemp = getCodes(txtDeptCD)
        
        If Len(txtDeptCD) <> 0 And InStr(1, strTemp, ",", vbTextCompare) > 0 Then 'Edit by Bryan 20/07/05 Ticket #8963
           DeptUnCri = DeptUnCri & " And {HREMP.ED_DEPTNO} IN ['" & strTemp & "'] " '() to [] changed by Sam 06-02-06
        ElseIf Len(txtDeptCD) <> 0 And InStr(1, strTemp, ",", vbTextCompare) = 0 Then
            DeptUnCri = DeptUnCri & " And {HREMP.ED_DEPTNO} = '" & strTemp & "' "
        End If 'Bryan End
    End If
    'Hemu - 06/02/2004 End
    
End If

If Len(DeptUnCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = DeptUnCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & DeptUnCri
    End If
    glbiOneWhere = True
End If

End Sub


Sub Main()
Dim strAP$
Dim xUID, xPWD, xCmd
Dim strHostFile As String
Dim isEmpt

'===============================================================
ChDir App.Path
strAP$ = App.Path
gflHelp$ = strAP$ & "\INFOHR.HLP"
App.HelpFile = gflHelp$

'Check if ihrhost.ini file exists to see if this system is hosted
'if IHRHost.ini exists then get the connection string info from the host.ini file instead of Registry
glbHosted = False
glbHostFile = App.Path & "\IHRHost.ini"
If File(glbHostFile) Then
    If FileLen(glbHostFile) <> 0 Then
        'Hosted Environment
        'Get connection values from ihrhost.ini and Set it to the global connection variables
        Call modLoadHostINI
        
        glbHosted = True
    End If
End If

If Not glbHosted Then
    'Non-Hosted Environment
    Call modLoadINI ' load ini file values - both win ini and prog
    
    'Frank 04/22/2003 Ticket# 4033
    If File(glbIHRWFC) And (Not glbSQL) And (Not glbOracle) Then
        If GetPasswordADO Then
            MsgBox "Database out of Sync-Encryption." & Chr(10) & "Please contact your system administrator to repair the database. "
            End
        End If
    End If
End If

glbINIFileName$ = glbWorkDir & "\IHR.INI"

frmSPLASH.Show
CenterForm frmSPLASH

glbEEOK = True

End Sub

Function GetPasswordADO() As Boolean
Dim db01 As New ADODB.Connection
Dim xAdoIHRDB As String
On Error GoTo Err_Deal
    GetPasswordADO = False
    xAdoIHRDB = glbAdoIHRDB
    xAdoIHRDB = Replace(xAdoIHRDB, "Jet OLEDB:Database Password=petman;", "")
    If db01.State = adStateOpen Then db01.Close
    db01.CommandTimeout = 600
    db01.Open xAdoIHRDB
    db01.Close
    GetPasswordADO = True
    Exit Function
Err_Deal:

End Function

Function GetPassword() As Boolean
Dim db01 As DAO.Database
Dim db01X As DAO.Database

On Error GoTo ErrorHandle:
    GetPassword = False
    Set db01 = OpenDatabase(glbIHRDB, True, False, ";PWD=petman")
    If File(glbIHRWFC) Then
        On Error Resume Next
        db01.NewPassword "", "petman"
        On Error GoTo 0
        
        On Error Resume Next
        Set db01X = OpenDatabase(glbIHRAUDIT, True, False, ";PWD=HRSS001")
        On Error GoTo 0
        On Error Resume Next
        db01X.NewPassword "", "petman"
        On Error GoTo 0
        On Error Resume Next
        db01X.Close
        On Error GoTo 0
    End If
    db01.Close
    GetPassword = True
    Exit Function
ErrorHandle:

End Function

Private Sub ChangePassword()
    Dim db01 As DAO.Database
    Dim db01X As DAO.Database
    On Error Resume Next
    Set db01 = OpenDatabase(glbIHRDB, True, False, ";PWD=HRSS001")
    On Error GoTo 0
    On Error Resume Next
    db01.NewPassword "HRSS001", "petman"
    On Error GoTo 0
    
    On Error Resume Next
    Set db01X = OpenDatabase(glbIHRAUDIT, True, False, ";PWD=HRSS001")
    On Error GoTo 0
    On Error Resume Next
    db01X.NewPassword "HRSS001", "petman"
    On Error GoTo 0
    
End Sub

Function modECount_FamilyDay()
Dim snapECount As New ADODB.Recordset
Dim SQLQ As String

    modECount_FamilyDay = 0
    
    SQLQ = "SELECT DISTINCT ED_SIN FROM HREMP "
    snapECount.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not snapECount.EOF Then
        modECount_FamilyDay = snapECount.RecordCount
    End If
    snapECount.Close
    Set snapECount = Nothing
    
End Function

Function modECount(Updat As Integer)
Dim ECount%
Dim snapECount As New ADODB.Recordset
Dim SQLQ As String

On Error GoTo CEE_Err

modECount = 10
SQLQ = "SELECT Count(HREMP.ED_EMPNBR) AS ECount FROM HREMP"

snapECount.Open SQLQ, gdbAdoIhr001, adOpenStatic

If snapECount.RecordCount < 1 Then Exit Function
snapECount.MoveFirst

If glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #27829
    ECount% = modECount_FamilyDay
Else
    ECount% = snapECount("ECount")
End If

snapECount.Close

If Updat Then
    SQLQ = "UPDATE HRPARCO "
    SQLQ = SQLQ & "SET HRPARCO.PC_NUMBER_EMPLOYEES =" & ECount%
    SQLQ = SQLQ & ", HRPARCO.PC_WHEN_COUNTED = " & Date_SQL(Date)
End If

gdbAdoIhr001.Execute SQLQ

modECount = ECount%


Exit Function

CEE_Err:
glbFrmCaption$ = "Count Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EE Count", "HREMP", "Count")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function modECountChk()
Dim ECount As Long
Dim EMax As Long

Dim snapECount As New ADODB.Recordset
Dim SQLQ

On Error GoTo ChkEE_Err

' returns true if less than or equal to their max

modECountChk = False
SQLQ = "SELECT * FROM HRPARCO"

snapECount.Open SQLQ, gdbAdoIhr001, adOpenStatic

If snapECount.RecordCount < 1 Then Exit Function
snapECount.MoveFirst

ECount = snapECount("PC_NUMBER_EMPLOYEES")

'---------------------------------------------------------------------------------
'Release 8.1 - The Max Employee License is encrypted now.
'Check if it's encrypted. If so, then decrypt it first and then assign to EMax.
If Not snapECount("PC_OPTUPD") Or IsNull(snapECount("PC_OPTUPD")) Then
    'Not Encrypted yet
    EMax = snapECount("PC_MAX_EMPLOYEES")
Else
    'Encrypted. Decrypt it first and then assign
    'Encrypted value is stored in the different field.
    If gsMultiLang = "Y" Then 'For Listowel only
        EMax = DecryptPasswordMultiLang_First(snapECount("PC_OPT"))
    ElseIf UCase(gsMultiLang) = "YES" Then 'whscc
        EMax = DecryptPasswordMultiLang(snapECount("PC_OPT"))
    Else
        EMax = DecryptPassword(snapECount("PC_OPT"))
    End If
End If
'EMax = snapECount("PC_MAX_EMPLOYEES")
'---------------------------------------------------------------------------------

glbNextEmpl = snapECount("PC_NEXT_AVAILABLE_NBR") 'laura nov 28, 1997
glbSysGen = snapECount("PC_SYSTEM_EMPLOYEE")      'laura nov 28, 1997

If DemoSystem = True Then EMax = DemoMaxEmp%

If ECount <= EMax - 1 Then
    modECountChk = True
Else
    modECountChk = False
End If

snapECount.Close


Exit Function

ChkEE_Err:
glbFrmCaption$ = "Count Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EE Count", "HREMP", "Count")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Sub modLoadINI()

Dim sPath As String, sSetting As String, strT  As String
Dim x%, Value$, I, DateStr
    
    '===============================================================
    'Reading windows defaults - ie date
    ' get the user's windows default date format
    lCurrentKey = HKEY_CURRENT_USER
    sPath = "CONTROL PANEL\INTERNATIONAL"
    ' dkostka - 07/12/2001 - Added check to make sure key is there before getting value.
    If DoesKeyExist(lCurrentKey, sPath) Then
        sSetting = "sShortDate"
        glbsDateFormat = "None Found"
        giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, glbsDateFormat)
        glbsDateFormat = UCase$(glbsDateFormat)
    Else
        glbsDateFormat = String(255, "0")
    End If
    ' dkostka - 03/20/01 - Added support for French Windows environment
    If IsDate("2 Mars 2003") Then glbFrench = True
    
    ' dkostka - 09/15/2000 - Changed for Linamar, Elgin, etc - If no date format in registry, try to find it manually.
    If glbsDateFormat = String(255, "0") Then
        If glbFrench Then DateStr = "Janvier 2, 2003" Else DateStr = "January 2, 2003"
        Select Case Format(DateStr, "Short Date")
            Case "01/02/2003", "1/02/2003", "01/2/2003", "1/2/2003"
                glbsDateFormat = "MM/DD/YYYY"
            Case "01/02/03", "1/02/03", "01/2/03", "1/2/03"
                glbsDateFormat = "MM/DD/YY"
            Case "02/01/2003", "2/01/2003", "02/1/2003", "2/1/2003"
                glbsDateFormat = "DD/MM/YYYY"
            Case "02/01/03", "2/01/03", "02/1/03", "2/1/03"
                glbsDateFormat = "DD/MM/YY"
            Case "2003/01/02", "2003/1/02", "2003/01/2", "2003/1/2"
                glbsDateFormat = "YYYY/MM/DD"
            Case "03/01/02", "03/1/02", "03/01/2", "03/1/2"
                glbsDateFormat = "YY/MM/DD"
            ' dkostka - 03/13/01 - Changed for Assumption Life - Added date formats with dashes
            Case "01-02-2003", "1-02-2003", "01-2-2003", "1-2-2003"
                glbsDateFormat = "DD-MM-YYYY"
            Case "01-02-03", "1-02-03", "01-2-03", "1-2-03"
                glbsDateFormat = "MM-DD-YY"
            Case "02-01-2003", "2-01-2003", "02-1-2003", "2-1-2003"
                glbsDateFormat = "DD-MM-YYYY"
            Case "02-01-03", "2-01-03", "02-1-03", "2-1-03"
                glbsDateFormat = "DD-MM-YY"
            Case "2003-01-02", "2003-1-02", "2003-01-2", "2003-1-2"
                glbsDateFormat = "YYYY-MM-DD"
            Case "03-01-02", "03-1-02", "03-01-2", "03-1-2"
                glbsDateFormat = "YY-MM-DD"
            Case Else
                glbsDateFormat = "Date Format Not Set"
        End Select
    End If
    ' dkostka - 09/15/2000 - end

    Call iniDateFormat  'Jaddy - May 1,2001 for date entry
    
    If Not DoesKeyExist(HKEY_LOCAL_MACHINE, REG_NAME) Then
        lCurrentKey = HKEY_CURRENT_USER
    Else
        lCurrentKey = HKEY_LOCAL_MACHINE
    End If
    
    sPath = REG_NAME & "INFOHR Files"
    '========================================
    'Jaddy Changed to remove all others registy keys except IHRDB
    'All database's Location will be set under IHRDB
    'IHRDB could have IHR001.MDB words or not
    'For example, F:\IHR; F:\IHR\ or F:\IHR\IHR001.MDB
    
    sSetting = "IHRDB"
    glbIHRDB = glbWorkDir & "\ihr001.mdb"
    giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, glbIHRDB)
    glbDBDir = Replace(UCase(glbIHRDB), "IHR001.MDB", "")
    glbDBDir = glbDBDir & IIf(Right(glbDBDir, 1) = "\", "", "\")
    If glbDBDir <> "" Then
        glbIHRDB = glbDBDir & "IHR001.MDB"
        glbIHRAUDIT = glbDBDir & "IHR001X.MDB"
        glbIHRDBW = glbDBDir & "IHR001W.MDB"
        glbIHRWFC = glbDBDir & "IHRWFC.MDB"
        glbIHRWFCA = glbDBDir & "IHRWFC-A.MDB"
        glbSN2322 = glbDBDir & "SN2322.MDB"
        glbIHRDBO = glbDBDir & "IHROPUS.MDB"
        glbIHRDBB = glbDBDir & "IHR001B.MDB"
        glbIHREDU = glbDBDir & "IHREDU.MDB"
    End If
    sSetting = "IHRREPORTS"  'Compressed database location
    glbIHRREPORTS = glbWorkDir
    giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, glbIHRREPORTS)
    glbIHRREPORTS = glbIHRREPORTS & IIf(Right$(glbIHRREPORTS, 1) <> "\", "\", "")
    
    '==================================================
    sPath = REG_NAME & "Network"
    sSetting = "EXCLUSIVE"  'Compressed database location
    strT = "N"
    giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    If strT = "Y" Then
        glbExclusiveDB% = True
    Else
        glbExclusiveDB% = False
    End If
    
    
    sSetting = "MULTIUSERNUM"  'Compressed database location
    strT = "0"
    giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    If IsNumeric(strT) Then glbMultiUserNum% = CInt(strT)
    

    '=========================================================
    sPath = REG_NAME & "Options"
    sSetting = "DatabaseType"
    strT = ""
    giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    gsSystemDb = UCase(strT)
    
    'Get Language
    sPath = REG_NAME & "Options"
    sSetting = "MultiLang"
    strT = ""
    giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    gsMultiLang = UCase(strT)
    If Len(gsMultiLang) > 250 Then
        gsMultiLang = "N"
    Else
        If gsMultiLang <> "Y" And gsMultiLang <> "YES" Then
            gsMultiLang = "N"
        End If
    End If
    'this changes were for Listowel. it can be used for every body
    
    '=================================================
'    sPath = REG_NAME & "CUSTOM SETUP"
'    sSetting = "SOUND"  ' do they want sound effects?
'    strT = "N"
'    giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
'    If strT = "N" Then
'        glbSound% = False
'    Else
'        glbSound% = True
'    End If
    
    sSetting = "ADDHISWARNING"  'Warning on when adding history
    strT = "N"
    giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    If strT = "Y" Then
        glbAddHisWarning% = True
    Else
        glbAddHisWarning% = False
    End If
    
    sPath = REG_NAME & "FOLLOWUPS"
    sSetting = "FOLLOWUPS"  'check for follow-ups when entering
    strT = "N"
    giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    If strT = "Y" Then
        glbFOLLOWUPS% = True
    Else
        glbFOLLOWUPS% = False
    End If
    
    sSetting = "FOLLOWUPDAYS"  'check for follow-ups when entering
    strT = "5"
    giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    If Not IsNumeric(strT) Then strT = "5"
    glbFOLLOWUPDAYS% = CInt(strT)
    
    
    sSetting = "SHOWCOMPLETED"  'check for follow-ups when entering
    strT = "N"
    giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    If strT = "Y" Then
        glbFOLLOWUPSCOMP% = True
    Else
        glbFOLLOWUPSCOMP% = False
    End If
        
    '=========================================================
    sPath = REG_NAME & "ODBC Setup"
    sSetting = "ODBCIHR"  'check for follow-ups when entering
    strT = "N"
    giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    If strT = "N" Then
        'Call mod_ODBC_Register("IHR")
        x% = WriteRegistrySetting(lCurrentKey, sPath, sSetting, "Y")
    End If
    
    sSetting = "ODBCIHRX"  'check for follow-ups when entering
    strT = "N"
    giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    If strT = "N" Then
        'Call mod_ODBC_Register("IHRX")
        x% = WriteRegistrySetting(lCurrentKey, sPath, sSetting, "Y")
    End If
            
    sSetting = "DATABASENAME"
    strT = ""
    giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    ' dkostka - 06/17/2002 - Don't force the SQL Server params to uppercase, if the server is set like
    '   WHSCC's server, you have to have the case right or it won't let you log in.
    'SQLDatabaseName = UCase(strT)
    SQLDatabaseName = strT
    If Len(SQLDatabaseName) = 255 Then
        SQLDatabaseName = "INFOHR"
    End If
    
    'SQL SERVER LOGIN SERVER NAME, USER NAME AND PASSWORD
    sSetting = "SERVERNAME"
    strT = ""
    giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    ' dkostka - 06/17/2002 - Don't force the SQL Server params to uppercase, if the server is set like
    '   WHSCC's server, you have to have the case right or it won't let you log in.
    'SQLServerName = UCase(strT)
    SQLServerName = strT
    If Len(SQLServerName) = 255 Then
        SQLServerName = ""
    End If
           
    sSetting = "USERNAME"
    strT = ""
    giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    ' dkostka - 06/17/2002 - Don't force the SQL Server params to uppercase, if the server is set like
    '   WHSCC's server, you have to have the case right or it won't let you log in.
    'SQLUserName = UCase(strT)
    SQLUserName = strT
    If Len(SQLUserName) = 255 Then
        SQLUserName = ""
    End If
    
    sSetting = "USERPSW"
    strT = ""
    giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    If gsMultiLang = "Y" Then
        SQLUserPassword = DecryptPasswordMultiLang_First(strT)
    ElseIf gsMultiLang = "YES" Then 'WHSCC
        SQLUserPassword = DecryptPasswordMultiLang(strT)
    Else
        glbDBPassFlag = True
        SQLUserPassword = DecryptPassword(strT)
    End If
    
    'Oracle driver NAME, USER NAME AND PASSWORD
    sSetting = "DRIVERNAME"
    strT = ""
    giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    ' dkostka - 06/17/2002 - Don't force the SQL Server params to uppercase, if the server is set like
    '   WHSCC's server, you have to have the case right or it won't let you log in.
    'SQLServerName = UCase(strT)
    SQLDriver = strT
    If Len(SQLDriver) = 255 Then
        SQLDriver = ""
    End If
        
    '=========================================================
    
    '-------------------------------------------------------------------------------------------------------------
    'gsDB_CONNECT_ENCRYPT = RetrieveCompanyPreference_Value("DB_CONNNECT_ENCRYPT")
    
    'The Encryption of Database Connection is turned-ON, retrieve encrypted database connection information from License Key
    'If gsDB_CONNECT_ENCRYPT Then
        'Ticket #24352 - PIPEDA
        'Retrieve the License key and Decrypt it
        Dim SQLHRSSLicense As String
        glbHRSSSecure = True
        sPath = REG_NAME & "Options"
        sSetting = "License"
        strT = ""
        giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
        SQLHRSSLicense = strT
        If Len(SQLHRSSLicense) = 344 Then
            SQLHRSSLicense = ""
        Else
            'Decrypt the License key
            SQLHRSSLicense = DecryptDatabaseSettings(SQLHRSSLicense)
    
            'Break down the license key to appropriate database connection variables
            If Len(SQLHRSSLicense) > 0 Then
                Call DatabaseConnection_License(SQLHRSSLicense)
            End If
        End If
        glbHRSSSecure = False
    'End If
    '-------------------------------------------------------------------------------------------------------------
    
    sPath = REG_NAME & "Options"
    sSetting = "DatabaseType"
    strT = ""
    giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    gsSystemDb = UCase(strT)
    
    ' dkostka - 02/16/01 - Added option to pass directories via command line options
    Call SetCmdLinePath
    '--- SET THE VALUES TO glbAdoIHRDB,glbAdoIHRAUDIT, ...
    ' Jaddy Jan 19, 2005 set default linamar login
    'Call SetLinamarLogin
    
    Call glbAdo_Value
        
End Sub

Private Sub SetLinamarLogin()
Dim SECTION$
Dim Key$
Dim x%, Value$, I, valtmp
    
    'This function is not called from anywhere now. I am not going to add the IHRHOST.INI option
    'for this function.

    SQLUserName = "ihradm"
    SQLUserPassword = "3mp10y"
  
    'Write Server Name to Register File
    SECTION$ = REG_NAME & "ODBC Setup"
    Key$ = "USERNAME"
    x% = WriteRegistrySetting(lCurrentKey, SECTION$, Key$, SQLUserName)
    
  
    'Write Server Name to Register File
    SECTION$ = REG_NAME & "ODBC Setup"
    Key$ = "USERPSW"
    Value$ = SQLUserPassword
    If Len(Value$) > 0 Then
      For I = 1 To Len(Value$)
          valtmp = Value$
          valtmp = Asc(Mid(Value$, I, 1))
          valtmp = Chr(valtmp + 80)
          valtmp = Mid(Value$, 1, I - 1) & valtmp & Mid(Value$, I + 1, Len(Value$) - I)
          Value$ = valtmp
      Next I
    End If
    x% = WriteRegistrySetting(lCurrentKey, SECTION$, Key$, Value$)
            
End Sub

Private Sub SetCmdLinePath()
    Dim Path As String
    
    ' Only set the path if they passed one
    If InStr(UCase(Command), "/PATH") <> 0 And gsSystemDb <> "MS SQL SERVER" Then
        Path = Mid(Command, InStr(UCase(Command), "/PATH") + 6)
        Path = Path & IIf(Right(Path, 1) = "\", "", "\")
        glbIHRDB = Path & "IHR001.MDB"
        glbIHRAUDIT = Path & "IHR001X.MDB"
        glbIHRDBW = Path & "IHR001W.MDB"
        glbIHRWFC = Path & "IHRWFC.MDB"
        glbIHRWFCA = Path & "IHRWFC-A.MDB"
        glbIHRDBO = Path & "IHROPUS.MDB"
        glbIHRDBB = Path & "IHR001B.MDB"
        glbIHREDU = Path & "IHREDU.mdb"
        glbIHRREPORTS = Path
    End If
End Sub

Sub glbAdo_Value()

    'If you want to switch to SQL Native Client Provider (newer) then read the link below to get the right connection string construction:
    'https://docs.microsoft.com/en-us/sql/relational-databases/native-client/applications/using-ado-with-sql-server-native-client

    Select Case gsSystemDb
    Case "MS SQL SERVER"
        glbSQL = True
        'glbAdoIHRDB = "Provider=SQLOLEDB.1;Persist Security Info=False;Network Library=DBMSSOCN;User ID=" & SQLUserName & ";Password=" & SQLUserPassword & ";Initial Catalog=" & SQLDatabaseName & ";Data Source=" & SQLServerName
        glbAdoIHRDB = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & SQLUserName & ";Password=" & SQLUserPassword & ";Initial Catalog=" & SQLDatabaseName & ";Data Source=" & SQLServerName
        'glbAdoIHRDB = "Provider=SQLNCLI;Server=(local);Persist Security Info=False;User ID=" & SQLUserName & ";Password=" & SQLUserPassword & ";Initial Catalog=" & SQLDatabaseName & ";Data Source=" & SQLServerName
        'glbAdoIHRDB = "Provider=SQLNCLI11;Data Source=" & SQLServerName & ";Database=" & SQLDatabaseName & ";User ID=" & SQLUserName & ";Password=" & SQLUserPassword & ";Persist Security Info=False;DataTypeCompatibility=80;MARS Connection=True;"
        
        glbAdoIHRDB_DOC = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & SQLUserName & ";Password=" & SQLUserPassword & ";Initial Catalog=" & SQLDatabaseName & "_DOC;Data Source=" & SQLServerName
        'glbAdoIHRDB_DOC = "Provider=SQLOLEDB.1;Persist Security Info=False;Network Library=DBMSSOCN;User ID=" & SQLUserName & ";Password=" & SQLUserPassword & ";Initial Catalog=" & SQLDatabaseName & "_DOC;Data Source=" & SQLServerName
        'glbAdoIHRDB_DOC = "Provider=SQLNCLI11;Data Source=" & SQLServerName & ";Database=" & SQLDatabaseName & "_DOC;User ID=" & SQLUserName & ";Password=" & SQLUserPassword & ";Persist Security Info=False;DataTypeCompatibility=80;MARS Connection=True;"
    
        glbAdoIHRAUDIT = glbAdoIHRDB
        glbAdoIHRDBW = glbAdoIHRDB
        glbadoIHREDU = glbAdoIHRDB
        RptODBC_SQL = ODBC_CONNECT_SQL & ";UID=" & SQLUserName & ";PWD=" & SQLUserPassword
        If SQLServerName <> "" Then
            DBEngine.RegisterDatabase "INFOHR", "SQL Server", True, "Database=" & SQLDatabaseName & vbCr & "Server=" & SQLServerName & vbCr
            'DBEngine.RegisterDatabase "INFOHR", "SQL Native Client", True, "Database=" & SQLDatabaseName & vbCr & "Server=" & SQLServerName & vbCr
        End If
   Case "ORACLE"
        glbOracle = True
        glbAdoIHRDB = "Provider=OraOLEDB.Oracle.1;Password=" & SQLUserPassword & ";Persist Security Info=True;User ID=" & SQLUserName & ";Data Source=" & SQLServerName & ""
        glbAdoIHRDB_DOC = "Provider=OraOLEDB.Oracle.1;Password=" & SQLUserPassword & ";Persist Security Info=True;User ID=" & SQLUserName & ";Data Source=" & SQLServerName & "_DOC"
        glbAdoIHRAUDIT = glbAdoIHRDB
        glbAdoIHRDBW = glbAdoIHRDB
        glbadoIHREDU = glbAdoIHRDB
        RptODBC_SQL = ODBC_CONNECT_SQL & ";UID=" & SQLUserName & ";PWD=" & SQLUserPassword
'        If SQLServerName <> "" Then
'            DBEngine.RegisterDatabase "INFOHR", "SQL Server", True, "Database=" & SQLDatabaseName & vbCr & "Server=" & SQLServerName & vbCr
'        End If
    Case Else
        glbSQL = False
        glbAdoIHRDB = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password=petman;Data Source=" & glbIHRDB
        glbAdoIHRAUDIT = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password=petman;Data Source=" & glbIHRAUDIT
        glbAdoIHRDBW = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & glbIHRDBW
        glbAdoIHRWFC = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & glbIHRWFC
        glbAdoIHRWFCA = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & glbIHRWFCA
        glbAdoIHRDBO = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & glbIHRDBO
        glbAdoIHRDBB = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & glbIHRDBB
        glbAdoSN2322 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & glbSN2322
        glbadoIHREDU = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & glbIHREDU
        RptODBC_SQL = "DSN=IHR001;PWD=petman;DSQ=" & glbIHRDB
        DBEngine.RegisterDatabase "IHR001", "Microsoft Access Driver (*.mdb)", True, "DBQ=" & glbIHRDB & vbCr
    End Select
    
End Sub
Function modNukeEETerm(EEID As Long)
Dim SQLQ As String
Dim TabName$, EEIDAlias$
Dim iRow As Integer, Msg As String
Dim snapEETables As New ADODB.Recordset
modNukeEETerm = False

On Error GoTo modNukeEETerm_Err

SQLQ = "SELECT * FROM INFO_HR_TABLES "
SQLQ = SQLQ & " WHERE Employee_Keyed <>0"
SQLQ = SQLQ & " AND TERMINATION_TABLE<>0"
'Ticket #20415 - Add Serial # to the select statement so custom tables also gets employee # changed.
'Serial 9999 is by default for all standard info:HR table.
'SQLQ = SQLQ & " AND (SERIAL = 'S/N - 9999W' OR SERIAL = '" & glbCompSerial & "')"
'Ticket #20893 Franks 09/02/2011 - only remove data for the standard INFO:HR tables
SQLQ = SQLQ & " AND (SERIAL = 'S/N - 9999W' OR SERIAL IS NULL) "

snapEETables.Open SQLQ, gdbAdoIhr001, adOpenStatic

iRow = 0
Do While Not snapEETables.EOF
    iRow = iRow + 1
    TabName$ = snapEETables("Table_Name")
    EEIDAlias$ = "TERM_SEQ"
    Call NukeEERows(TabName$, EEIDAlias$, EEID&)
    snapEETables.MoveNext
Loop

snapEETables.Close
'if termination tables are missing need a new updefaults76.exe
'bryan Mar 26, 2007 Ticket#12852
If iRow = 0 Then
    MsgBox "Termination Tables are missing, please contact HR Systems Strategies Support for a Defaults Update", vbInformation + vbOKOnly, "Missing Defaults"
End If

If glbAxxent Then
    SQLQ = "DELETE FROM Term_HRRSP WHERE TERM_SEQ = " & EEID
    gdbAdoIhr001X.Execute SQLQ
End If

If glbCompSerial = "S/N - 2279W" Then  'Friesens Corporation
    SQLQ = "DELETE FROM Term_PERFORM_FRIESEN WHERE TERM_SEQ =" & EEID & " "
    gdbAdoIhr001X.Execute SQLQ
End If

modNukeEETerm = True

Exit Function

modNukeEETerm_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modNukeEETerm", "modNukeEETerm", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If


End Function

Function modSecurity_Check(UserID As String, PASSWD As String) As Boolean
Dim Secure_Snap As New ADODB.Recordset
Dim Secure_Access_Snap As New ADODB.Recordset
Dim rsTD As New ADODB.Recordset
Dim rsSecAcc As New ADODB.Recordset
Dim rsLabel As New ADODB.Recordset
Dim SQLQ As String
Dim xStr As String, x
Dim xSecTemplate As String

On Error GoTo Secure_Err

modSecurity_Check = False

'Check first it database has been converted to 7.9 with this 7.9 exe
rsLabel.Open "SELECT LB_LANG FROM HRLABEL WHERE 1 = 2", gdbAdoIhr001, adOpenStatic, adLockOptimistic
If rsLabel.EOF Then
    'DO NOTHING - DATABASE HAS BEEN CONVERTED
End If
rsLabel.Close
Set rsLabel = Nothing

If glbSSO <> "YES" Then
    frmSECURITY.panHelpEntry.Caption = "Verifying Password..." 'Added by Bryan 11/07/05 Ticket #8855
End If

glbPassword$ = PASSWD
glbTxtPassword = PASSWD

If gsMultiLang = "Y" Then
    'Temporary
    'WriteFile ("In Multi Lang = Y; Password: " & glbPassword)
    
    gdbAdoIhr001.BeginTrans
    SQLQ = "QRY_SETPASSWORD ('" & Replace(UserID, "'", "''") & "',-80)"
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
    
    SQLQ = "Select USERID,PASSWORD,WRKEMP from HRSECWRK WHERE WRKEMP=USERID AND USERID='" & Replace(UserID, "'", "''") & "' AND PASSWORD='" & glbPassword & "'"
    Secure_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If Not Secure_Snap.EOF Then
        'Temporary
        'WriteFile ("In Multi Lang = Y; Query before Delete: " & SQLQ)
        
        Secure_Snap.Delete
        
        SQLQ = "Select * from HR_SECURE_BASIC"
        SQLQ = SQLQ & " where USERID = '" & Replace(UserID, "'", "''") & "'"
        Secure_Snap.Close
        Secure_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic
        
        'Temporary
        'WriteFile ("In Multi Lang = YES; Query after Delete: " & SQLQ)
    End If
ElseIf gsMultiLang = "YES" Then 'whscc
    'Temporary
    'WriteFile ("In Multi Lang = YES; Password: " & glbPassword)
    
    SQLQ = "Select * from HR_SECURE_BASIC"
    SQLQ = SQLQ & " where USERID = '" & Replace(UserID, "'", "''") & "'"
    SQLQ = SQLQ & " AND PassWord = " & "'" & EncryptPasswordMultiLang(glbPassword) & "'"
    Secure_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    'Temporary
    'WriteFile ("In Multi Lang = YES; Query: " & SQLQ)
Else
    'Temporary
    'WriteFile ("In Regular Encryption; Password: " & glbPassword)
    
    SQLQ = "Select * from HR_SECURE_BASIC"
    SQLQ = SQLQ & " where USERID = '" & Replace(UserID, "'", "''") & "'"
    'SQLQ = SQLQ & " AND PassWord = " & "'" & glbPassword & "'"
    SQLQ = SQLQ & " AND PassWord = " & "'" & EncryptPassword(glbPassword) & "'"
    Secure_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    'Temporary
    'WriteFile ("In Regular Encryption; Query: " & SQLQ)
    
    If glbCompSerial = "S/N - 2188W" Then  'Ticket #14707
        If Secure_Snap.EOF And Secure_Snap.BOF Then
            Secure_Snap.Close
            SQLQ = "Select * from HR_SECURE_BASIC"
            SQLQ = SQLQ & " where USERID = '" & Replace(UserID, "'", "''") & "'"
            SQLQ = SQLQ & " AND PassWord = " & "'" & PASSWD & "'"
            Secure_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic
            glbPassword = PASSWD
        End If
    End If
    
End If

If Secure_Snap.EOF And Secure_Snap.BOF Then
    'Temporary
    'WriteFile ("Security record not found; Query: " & SQLQ)
    
    Exit Function
Else
    'Ticket #24808 - Moving Lock Password into HR_SECURE_BASIC table because if the User's are Template based then this
    'Lock option will not work as the User's profile is retrieved from Template.
    'Password should not be locked
    If glbCompSerial = "S/N - 2407W" Then 'Ticket #18406 - Farmers' Mutual Insurance
        'Ticket #24808 - Check if User's Password is locked.
        If Not IsNull(Secure_Snap("LOCK_PASSWORD")) Then
            If Secure_Snap("LOCK_PASSWORD") Then
                'Password is locked
                Secure_Snap.Close
                Set Secure_Snap = Nothing
                Exit Function
            End If
        End If
        SQLQ = "SELECT " & Field_SQL("FUNCTION") & ", ACCESSABLE FROM HR_SECURE_ACCESS "
        SQLQ = SQLQ & " WHERE USERID='" & Replace(UserID, "'", "''") & "'"
        SQLQ = SQLQ & " AND " & Field_SQL("FUNCTION") & " = 'Lock_Password'"
        rsSecAcc.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsSecAcc.EOF Then
            If rsSecAcc("ACCESSABLE") Then
                'Password is locked
                rsSecAcc.Close
                Set rsSecAcc = Nothing
                Exit Function
            End If
        End If
        rsSecAcc.Close
        Set rsSecAcc = Nothing
    End If
End If

If glbSSO <> "YES" Then
    frmSECURITY.panHelpEntry.Caption = "Updating Security..." 'Added by Bryan 11/07/05 Ticket #8855
End If

'????Ticket #24808 -  Get User's Template if there
xSecTemplate = ""
If Not IsNull(Secure_Snap("SECURE_TEMPLATE")) Then
    xSecTemplate = Secure_Snap("SECURE_TEMPLATE")
End If

glbUserNAME = Secure_Snap("USERNAME")
gSec_Emp_Based = Secure_Snap("Empnbr_Based")
glbCompNo$ = Secure_Snap("COMPNO")
glbEmpNbr = 0

If Not IsNull(Secure_Snap("EMPNBR")) Then glbEmpNbr = Secure_Snap("EMPNBR")

'????Ticket #24808 -  Retrieve Template's security profile if User is Template based
If xSecTemplate = "" Or xSecTemplate = "TEMPLATE" Then
    Secure_Access_Snap.Open "select * from HR_SECURE_ACCESS WHERE USERID='" & Replace(UserID, "'", "''") & "' AND CODENAME IS NULL", gdbAdoIhr001, adOpenStatic
Else
    Secure_Access_Snap.Open "select * from HR_SECURE_ACCESS WHERE USERID='" & Replace(xSecTemplate, "'", "''") & "' AND CODENAME IS NULL", gdbAdoIhr001, adOpenStatic
End If
Do Until Secure_Access_Snap.EOF
    If Secure_Access_Snap("FUNCTION") = "Basic_Inquiry" Then gSec_Inq_Basic = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Basic_Update" Then gSec_Upd_Basic = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Banking_Inquiry" Then gSec_Inq_Banking = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Banking_Update" Then gSec_Upd_Banking = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "EmploymentEQT_Update" Then gSec_Upd_EmploymentEQT = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "EmploymentEQT_Inquiry" Then gSec_Inq_EmploymentEQT = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "PayEQT_Update" Then gSec_Upd_PayEQT = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "PayEQT_Inquiry" Then gSec_Inq_PayEQT = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Dependents_Update" Then gSec_Upd_Dependents = Secure_Access_Snap("ACCESSABLE")
    'Ticket #22009 Franks 05/10/2012
    If Secure_Access_Snap("FUNCTION") = "Del_Dependents" Then gSec_Del_Dependents = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Dependents_Inquiry" Then gSec_Inq_Dependents = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Skills_Update" Then gSec_Upd_Skills = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Skills_Inquiry" Then gSec_Inq_Skills = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Formal_Education_Update" Then gSec_Upd_Formal_Education = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Formal_Education_Inquiry" Then gSec_Inq_Formal_Education = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Education_Seminars_Update" Then gSec_Upd_Education_Seminars = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Education_Seminars_Inquiry" Then gSec_Inq_Education_Seminars = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Salary_Update" Then gSec_Upd_Salary = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Salary_Inquiry" Then gSec_Inq_Salary = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Performance_Update" Then gSec_Upd_Performance = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Performance_Inquiry" Then gSec_Inq_Performance = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Position_Update" Then gSec_Upd_Position = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Position_Inquiry" Then gSec_Inq_Position = Secure_Access_Snap("ACCESSABLE")
                
                
    If Secure_Access_Snap("FUNCTION") = "Benefits_Update" Then gSec_Upd_Benefits = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Benefits_Inquiry" Then gSec_Inq_Benefits = Secure_Access_Snap("ACCESSABLE")
    
    '7.9 Enhancement
    If Secure_Access_Snap("FUNCTION") = "Beneficiary_Update" Then gSec_Upd_Beneficiary = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Beneficiary_Inquiry" Then gSec_Inq_Beneficiary = Secure_Access_Snap("ACCESSABLE")
    
    If Secure_Access_Snap("FUNCTION") = "Entitlements_Update" Then gSec_Upd_Entitlements = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Entitlements_Inquiry" Then gSec_Inq_Entitlements = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Associations_Update" Then gSec_Upd_Associations = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Associations_Inquiry" Then gSec_Inq_Associations = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "UserDefineTbl_Update" Then gSec_Upd_UserDefineTbl = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "UserDefineTbl_Inquiry" Then gSec_Inq_UserDefineTbl = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "PayrollTrans_Update" Then gSec_Upd_PayrollTrans = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "PayrollTrans_Inquiry" Then gSec_Inq_PayrollTrans = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Follow_Ups_Update" Then gSec_Upd_Follow_Ups = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Follow_Ups_Inquiry" Then gSec_Inq_Follow_Ups = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Health_Safety_Update" Then gSec_Upd_Health_Safety = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Health_Safety_Inquiry" Then gSec_Inq_Health_Safety = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Attendance_Update" Then gSec_Upd_Attendance = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Attendance_Inquiry" Then gSec_Inq_Attendance = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Attendance_History_Update" Then gSec_Upd_Attendance_History = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Attendance_History_Inquiry" Then gSec_Inq_Attendance_History = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Other_Entitlements_Update" Then gSec_Upd_Other_Entitlements = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Other_Entitlements_Inquiry" Then gSec_Inq_Other_Entitlements = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Other_Earnings_Update" Then gSec_Upd_Earnings = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Other_Earnings_Inquiry" Then gSec_Inq_Earnings = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Hrly_Entitlements_Update" Then gSec_Upd_Hrly_Entitlements = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Hrly_Entitlements_Inquiry" Then gSec_Inq_Hrly_Entitlements = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Show_SIN_SSN" Then gSec_Show_SIN_SSN = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Show_DOB" Then gSec_Show_DOB = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Show_ADDRESS" Then gSec_Show_ADDRESS = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Show_MARITAL" Then gSec_Show_Marital = Secure_Access_Snap("ACCESSABLE")
    If glbCompSerial = "S/N - 2407W" Then 'Ticket #18406 - Farmers' Mutual Insurance
        If Secure_Access_Snap("FUNCTION") = "Lock_Password" Then gSec_Lock_Password = Secure_Access_Snap("ACCESSABLE")
    End If
  'tkt#10423 Jerry said remove serial#control for add_Attendance
  '  If glbCompSerial = "S/N - 2173W" Then
        If Secure_Access_Snap("FUNCTION") = "Add_Attendance" Then gSec_Add_Attendance = Secure_Access_Snap("ACCESSABLE")
   ' End If
    'Ticket #22682 - Release 8.0
    If Secure_Access_Snap("FUNCTION") = "Add_NewHire" Then gSec_Add_NewHire = Secure_Access_Snap("ACCESSABLE")
    
    'Release 8.1
    If Secure_Access_Snap("FUNCTION") = "Add_Comments" Then gSec_Add_Comments = Secure_Access_Snap("ACCESSABLE")
    
    'Ticket #23923 - Release 8.0 - View Own
    If Secure_Access_Snap("FUNCTION") = "ScsPlan_ViewOwn" Then gSec_SP_ViewOwn = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Comments_ViewOwn" Then gSec_Comments_ViewOwn = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Counsel_ViewOwn" Then gSec_Counsel_ViewOwn = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "FollUp_ViewOwn" Then gSec_FollUp_ViewOwn = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "OthInfo_ViewOwn" Then gSec_OthInfo_ViewOwn = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "EmpFlags_ViewOwn" Then gSec_EmpFlags_ViewOwn = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "EmpHis_ViewOwn" Then gSec_EmpHis_ViewOwn = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "GLDist_ViewOwn" Then gSec_GLDist_ViewOwn = Secure_Access_Snap("ACCESSABLE")
               
    'Ticket #28635 - Add View Own security
    If Secure_Access_Snap("FUNCTION") = "Perform_ViewOwn" Then gSec_Performance_ViewOwn = Secure_Access_Snap("ACCESSABLE")
    
    If Secure_Access_Snap("FUNCTION") = "Counselling_Update" Then gSec_Upd_Counselling = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Counselling_Inquiry" Then gSec_Inq_Counselling = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Comments_Update" Then gSec_Upd_Comments = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Comments_Inquiry" Then gSec_Inq_Comments = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "OtherInformation_Update" Then gSec_Upd_OtherInformation = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "OtherInformation_Inquiry" Then gSec_Inq_OtherInformation = Secure_Access_Snap("ACCESSABLE")
    
    'Ticket #20052 Franks 07/22/2011
    'Samuel Ticket #21000 Franks 09/26/2011 - move to Custom Security
    'If Secure_Access_Snap("FUNCTION") = "Profit_Sharing_Update" Then gSec_Upd_Profit_Sharing = Secure_Access_Snap("ACCESSABLE")
    'If Secure_Access_Snap("FUNCTION") = "Profit_Sharing_Inquiry" Then gSec_Inq_Profit_Sharing = Secure_Access_Snap("ACCESSABLE")
    'If Secure_Access_Snap("FUNCTION") = "Report_Profit_Sharing" Then gSec_Rpt_Profit_Sharing = Secure_Access_Snap("ACCESSABLE")
    If glbCompSerial = "S/N - 2382W" Then 'Samuel Ticket #21000 Franks 09/26/2011
        If Secure_Access_Snap("FUNCTION") = "SAM_Profit_Sharing_Upt" Then gSec_Upd_Profit_Sharing = Secure_Access_Snap("ACCESSABLE")
        If Secure_Access_Snap("FUNCTION") = "SAM_Profit_Sharing_Inq" Then gSec_Inq_Profit_Sharing = Secure_Access_Snap("ACCESSABLE")
        If Secure_Access_Snap("FUNCTION") = "SAM_Profit_Sharing_Rpt" Then gSec_Rpt_Profit_Sharing = Secure_Access_Snap("ACCESSABLE")
        If Secure_Access_Snap("FUNCTION") = "SAM_Red_Circled_Rpt" Then gSec_Rpt_Red_Circled = Secure_Access_Snap("ACCESSABLE")
        'Ticket #21581 Franks 02/14/2012
        If Secure_Access_Snap("FUNCTION") = "SAM_Table_Master_Links_Upt" Then gSec_Upd_SAMTableMasterLinks = Secure_Access_Snap("ACCESSABLE")
        If Secure_Access_Snap("FUNCTION") = "SAM_Table_Master_Links_Inq" Then gSec_Inq_SAMTableMasterLinks = Secure_Access_Snap("ACCESSABLE")
    End If
    
    If Secure_Access_Snap("FUNCTION") = "Job_Classes_Update" Then gSec_Upd_Job_Classes = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Job_Classes_Inquiry" Then gSec_Inq_Job_Classes = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Job_Master_Update" Then gSec_Upd_Job_Master = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Job_Master_Inquiry" Then gSec_Inq_Job_Master = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Job_Eval_Update" Then gSec_Upd_Job_Eval = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Job_Eval_Inquiry" Then gSec_Inq_Job_Eval = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Job_Skills_Update" Then gSec_Upd_Job_Skills = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Job_Skills_Inquiry" Then gSec_Inq_Job_Skills = Secure_Access_Snap("ACCESSABLE")
                
    'If Secure_Access_Snap("FUNCTION") = "Termination_Report" Then  gSec_Inq_Termination_Report  = SECURE_ACCESS_SNAP("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Terminations_Update" Then gSec_Upd_Terminations = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Termination_Inquiry" Then gSec_Inq_Terminations = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Company_Update" Then gSec_Upd_Company = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Company_Inquiry" Then gSec_Inq_Company = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Department_Inquiry" Then gSec_Inq_Departments = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Departments_Update" Then gSec_Upd_Departments = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Divisions_Update" Then gSec_Upd_Divisions = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Divisions_Inquiry" Then gSec_Inq_Divisions = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Sal_Distribute_Update" Then gSec_Upd_SalDist = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Sal_Distribute_Inquiry" Then gSec_Inq_SalDist = Secure_Access_Snap("ACCESSABLE")
    
    If Secure_Access_Snap("FUNCTION") = "EMP_FLAGS_Update" Then gSec_Upd_EMP_FLAGS = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "EMP_FLAGS_Inquiry" Then gSec_Inq_EMP_FLAGS = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "EMP_HISTORY_Update" Then gSec_Upd_EMP_HISTORY = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "EMP_HISTORY_Inquiry" Then gSec_Inq_EMP_HISTORY = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "GL_DIST_Update" Then gSec_Upd_GLDist = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "GL_DIST_Inquiry" Then gSec_Inq_GLDist = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "EMP_LANG_Update" Then gSec_Upd_EMP_LANG = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "EMP_LANG_Inquiry" Then gSec_Inq_EMP_LANG = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "EMP_SUCCESSION_Update" Then gSec_Upd_SUCCESSION = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "EMP_SUCCESSION_Inquiry" Then gSec_Inq_SUCCESSION = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Emergency_Contacts_Update" Then gSec_Upd_EmergContacts = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Emergency_Contacts_Inquiry" Then gSec_Inq_EmergContacts = Secure_Access_Snap("ACCESSABLE")
    
    If Secure_Access_Snap("FUNCTION") = "Work_Schedule_Update" Then gSec_Upd_Work_Schedule = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Work_Schedule_Inquiry" Then gSec_Inq_Work_Schedule = Secure_Access_Snap("ACCESSABLE")
    
    If Secure_Access_Snap("FUNCTION") = "Pay_Period_Update" Then gSec_Upd_PayPeriod_Master = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Pay_Period_Inquiry" Then gSec_Inq_PayPeriod_Master = Secure_Access_Snap("ACCESSABLE")
    
    'If Secure_Access_Snap("FUNCTION") = "Quick_ESS_Update" Then gSec_Upd_Quick_ESS = Secure_Access_Snap("ACCESSABLE")
    'If Secure_Access_Snap("FUNCTION") = "Quick_ESS_Inquiry" Then gSec_Inq_Quick_ESS = Secure_Access_Snap("ACCESSABLE")
    
    If Secure_Access_Snap("FUNCTION") = "Email_Setup_Update" Then gSec_Upd_Email_Setup = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Email_Setup_Inquiry" Then gSec_Inq_Email_Setup = Secure_Access_Snap("ACCESSABLE")
    
    If Secure_Access_Snap("FUNCTION") = "Payroll_Category_Update" Then gSec_Upd_Payroll_Category = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Payroll_Category_Inquiry" Then gSec_Inq_Payroll_Category = Secure_Access_Snap("ACCESSABLE")
    
    If Secure_Access_Snap("FUNCTION") = "Charge_Code_Update" Then gSec_Upd_Charge_Code = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Charge_Code_Inquiry" Then gSec_Inq_Charge_Code = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Project_Code_Update" Then gSec_Upd_Project_Code = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Project_Code_Inquiry" Then gSec_Inq_Project_Code = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Machine_Update" Then gSec_Upd_Machine = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Machine_Inquiry" Then gSec_Inq_Machine = Secure_Access_Snap("ACCESSABLE")
    
    If Secure_Access_Snap("FUNCTION") = "AttendCode_Matrix_Update" Then gSec_Upd_AttendCode_Matrix = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "AttendCode_Matrix_Inquiry" Then gSec_Inq_AttendCode_Matrix = Secure_Access_Snap("ACCESSABLE")
    
    'Ticket #22682 - Release 8.0 - Follow Up Code Email Matrix
    If Secure_Access_Snap("FUNCTION") = "FollowUpCodeEmail_Matrix_Update" Then gSec_Upd_FollowUpEmail_Matrix = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "FollowUpCodeEmail_Matrix_Inquiry" Then gSec_Inq_FollowUpEmail_Matrix = Secure_Access_Snap("ACCESSABLE")
    
    'Ticket #25922 - OHRS Reporting for CHC
    If Secure_Access_Snap("FUNCTION") = "OHRS_Department_Update" Then gSec_Upd_OHRSDepartments = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "OHRS_Department_Inquiry" Then gSec_Inq_OHRSDepartments = Secure_Access_Snap("ACCESSABLE")
    
    'Ticket #25746 - Department/GL Number Matrix
    If Secure_Access_Snap("FUNCTION") = "DeptGL_Matrix_Update" Then gSec_Upd_DeptGL_Matrix = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "DeptGL_Matrix_Inquiry" Then gSec_Inq_DeptGL_Matrix = Secure_Access_Snap("ACCESSABLE")

    If Secure_Access_Snap("FUNCTION") = "Ledgers_Update" Then gSec_Upd_Ledgers = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Ledgers_Inquiry" Then gSec_Inq_Ledgers = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Security_Update" Then gSec_Upd_Security = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Security_Inquiry" Then gSec_Inq_Security = Secure_Access_Snap("ACCESSABLE")
    If glbLinamar Then
        If Secure_Access_Snap("FUNCTION") = "DoorAccess_Update" Then gSec_Upd_DoorAccess = Secure_Access_Snap("ACCESSABLE")
        If Secure_Access_Snap("FUNCTION") = "DoorAccess_Inquiry" Then gSec_Inq_DoorAccess = Secure_Access_Snap("ACCESSABLE")

        If Secure_Access_Snap("FUNCTION") = "DoorName" Then gSec_DoorName = Secure_Access_Snap("ACCESSABLE")
        If Secure_Access_Snap("FUNCTION") = "Summarize_Attendance" Then gSec_Summarize_Attendance = Secure_Access_Snap("ACCESSABLE")
        If Secure_Access_Snap("FUNCTION") = "Report_DoorAccess" Then gSec_Rpt_DoorAccess = Secure_Access_Snap("ACCESSABLE")
    End If
    If glbCompSerial = "S/N - 2380W" Then ' For VitalAire Canada Inc. Ticket #26233 Franks 11/20/2014
        If Secure_Access_Snap("FUNCTION") = "DoorAccess_Update" Then gSec_Upd_DoorAccess = Secure_Access_Snap("ACCESSABLE")
        If Secure_Access_Snap("FUNCTION") = "DoorAccess_Inquiry" Then gSec_Inq_DoorAccess = Secure_Access_Snap("ACCESSABLE")
    End If
    If Secure_Access_Snap("FUNCTION") = "CustomReport_Update" Then gSec_Upd_CustomReport = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "CustomReport_Inquiry" Then gSec_Inq_CustomReport = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Holiday_Update" Then gSec_Upd_Holiday = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Holiday_Inquiry" Then gSec_Inq_Holiday = Secure_Access_Snap("ACCESSABLE")
    
    If Secure_Access_Snap("FUNCTION") = "New_Hire_Update" Then gSec_Upd_New_Hire = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "New_Hire_Inquiry" Then gSec_Inq_New_Hire = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Label_Update" Then gSec_Upd_Label = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Label_Inquiry" Then gSec_Inq_Label = Secure_Access_Snap("ACCESSABLE")
    
    If Secure_Access_Snap("FUNCTION") = "Codes" Then gSec_Mass_Codes = Secure_Access_Snap("ACCESSABLE")
    
    If glbWFC Then
        If Secure_Access_Snap("FUNCTION") = "SalaryGrids_Update" Then gSec_Upd_SalaryGrids = Secure_Access_Snap("ACCESSABLE")
        If Secure_Access_Snap("FUNCTION") = "SalaryGrids_Inquiry" Then gSec_Inq_SalaryGrids = Secure_Access_Snap("ACCESSABLE")
        'If Secure_Access_Snap("FUNCTION") = "WFC_Bonus_Intergration_Interface" Then gSec_WFC_Bonus_Intergration_Interface = Secure_Access_Snap("ACCESSABLE")
        If Secure_Access_Snap("FUNCTION") = "WFC_Band_Security" Then gSec_WFC_Band_Security = Secure_Access_Snap("ACCESSABLE")
        'Ticket #18566 - begin
        If Secure_Access_Snap("FUNCTION") = "RetirementProc_Update" Then gSec_Upd_RetirementProc = Secure_Access_Snap("ACCESSABLE")
        If Secure_Access_Snap("FUNCTION") = "RetirementProc_Inquiry" Then gSec_Inq_RetirementProc = Secure_Access_Snap("ACCESSABLE")
        If Secure_Access_Snap("FUNCTION") = "DeathProc_Update" Then gSec_Upd_DeathProc = Secure_Access_Snap("ACCESSABLE")
        If Secure_Access_Snap("FUNCTION") = "DeathProc_Inquiry" Then gSec_Inq_DeathProc = Secure_Access_Snap("ACCESSABLE")
        'Ticket #18566 - end
        'Ticket #22533 Franks 09/10/2012
        If Secure_Access_Snap("FUNCTION") = "WFC_UnlockSmokerStatus" Then gSec_WFC_UnlockSmokerStatus = Secure_Access_Snap("ACCESSABLE")
        'Ticket #29846 Franks 03/07/2017 ----------------- begin
        If Secure_Access_Snap("FUNCTION") = "WFC_IPExchangeRate" Then gSec_WFC_IPExchangeRate = Secure_Access_Snap("ACCESSABLE")
        If Secure_Access_Snap("FUNCTION") = "WFC_IPIncentiveFactors" Then gSec_WFC_IPIncentiveFactors = Secure_Access_Snap("ACCESSABLE")
        If Secure_Access_Snap("FUNCTION") = "WFC_IPCreateSpreadsheet" Then gSec_WFC_IPCreateSpreadsheet = Secure_Access_Snap("ACCESSABLE")
        If Secure_Access_Snap("FUNCTION") = "WFC_IPImportSpreadsheet" Then gSec_WFC_IPImportSpreadsheet = Secure_Access_Snap("ACCESSABLE")
        If Secure_Access_Snap("FUNCTION") = "WFC_IPUpdateEarnings" Then gSec_WFC_IPUpdateEarnings = Secure_Access_Snap("ACCESSABLE")
        If Secure_Access_Snap("FUNCTION") = "WFC_IPPreparePayrollFile" Then gSec_WFC_IPPreparePayrollFile = Secure_Access_Snap("ACCESSABLE")
        If Secure_Access_Snap("FUNCTION") = "WFC_IPPrintSpreadsheet" Then gSec_WFC_IPPrintSpreadsheet = Secure_Access_Snap("ACCESSABLE")
        If Secure_Access_Snap("FUNCTION") = "WFC_IPPrintLetter" Then gSec_WFC_IPPrintLetter = Secure_Access_Snap("ACCESSABLE")
        'Ticket #29846 Franks 03/07/2017 ----------------- end
    End If
    If Secure_Access_Snap("FUNCTION") = "AffirmAction_Data_Update" Then gSec_Upd_AffirmAction_Data = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "AffirmAction_Data_Inquiry" Then gSec_Inq_AffirmAction_Data = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "AffirmAction_Purge_Update" Then gSec_Upd_AffirmAction_Purge = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "AffirmAction_Purge_Inquiry" Then gSec_Inq_AffirmAction_Purge = Secure_Access_Snap("ACCESSABLE")
    
    If Secure_Access_Snap("FUNCTION") = "Compress_Fix" Then gSec_Compress_Fix = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "CompanyPreference" Then gSec_CompanyPreference = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "EmpFlagsSetup" Then gSec_EmpFlagsSetup = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "MultiDataSourceSetup" Then gSec_MultiDataSourceSetup = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "HelpDescSetup" Then gSec_HelpDescSetup = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "BenefitGroupSetup" Then gSec_BenefitGroupSetup = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "ChangeYourPassword" Then gSec_ChangeYourPassword = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "ITAdmin" Then gSec_ITAdmin = Secure_Access_Snap("ACCESSABLE")
    
    If Secure_Access_Snap("FUNCTION") = "Audit_Inquiry" Then gSec_Inq_Audit = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Audit_Update" Then gSec_Upd_Audit = Secure_Access_Snap("ACCESSABLE")
    
    'Ticket #23409 - Samuel, Son & Co., Limited - Discipline Audit Table Report
    If glbCompSerial = "S/N - 2382W" Then
        If Secure_Access_Snap("FUNCTION") = "CounselAudit_Inquiry" Then gSec_Inq_CounselAudit = Secure_Access_Snap("ACCESSABLE")
        If Secure_Access_Snap("FUNCTION") = "CounselAudit_Update" Then gSec_Upd_CounselAudit = Secure_Access_Snap("ACCESSABLE")
    End If
    
    'Ticket #24655 - Wellington-Dufferin-Guelph Public Health - On Call Hours
    If glbCompSerial = "S/N - 2411W" Then
        If Secure_Access_Snap("FUNCTION") = "On_Call_Hours_Update" Then gSec_Upd_OnCallHours = Secure_Access_Snap("ACCESSABLE")
        If Secure_Access_Snap("FUNCTION") = "On_Call_Hours_Inquiry" Then gSec_Inq_OnCallHours = Secure_Access_Snap("ACCESSABLE")
    End If
    
    If Secure_Access_Snap("FUNCTION") = "Export_Attendance" Then gSec_Export_Attendance = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Export_Salaries" Then gSec_Export_Salaries = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Export_Benefits" Then gSec_Export_Benefits = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Export_Employee" Then gSec_Export_Employee = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Export_Table" Then gSec_Export_Table = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Export_YTD" Then gSec_Export_YTD = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Export_PayrollTrans" Then gSec_Export_PayrollTrans = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Export_ContEdu" Then gSec_Export_ContEdu = Secure_Access_Snap("ACCESSABLE")
                
    If Secure_Access_Snap("FUNCTION") = "Import_Attendance" Then gSec_Import_Attendance = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Import_Salaries" Then gSec_Import_Salaries = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Import_Benefits" Then gSec_Import_Benefits = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Import_Employee" Then gSec_Import_Employee = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Import_Table" Then gSec_Import_Table = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Import_YTD" Then gSec_Import_YTD = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Import_PayrollTrans" Then gSec_Import_PayrollTrans = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Import_ContEdu" Then gSec_Import_ContEdu = Secure_Access_Snap("ACCESSABLE")
                
    If Secure_Access_Snap("FUNCTION") = "Province" Then gSec_Province = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Entitle" Then gSec_Entitle = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Matrix" Then gSec_Matrix = Secure_Access_Snap("ACCESSABLE")
    
    'Ticket #29122 - New Database Setup and Integration Setup securities
    If Secure_Access_Snap("FUNCTION") = "IntegrtDBSetup_Inquiry" Then gSec_Inq_IntegrtDBSetup = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "IntegrtDBSetup_Update" Then gSec_Upd_IntegrtDBSetup = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "IntegrtSetup_Inquiry" Then gSec_Inq_IntegrtSetup = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "IntegrtSetup_Update" Then gSec_Upd_IntegrtSetup = Secure_Access_Snap("ACCESSABLE")
    
    
    If Secure_Access_Snap("FUNCTION") = "LDate" Then gLast_Date = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "LTime" Then gvarLast_Time$ = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "LUser" Then gLast_User& = Secure_Access_Snap("ACCESSABLE")
                
                
    If Secure_Access_Snap("FUNCTION") = "Report_Age" Then gSec_Rpt_Age = Secure_Access_Snap("ACCESSABLE")
    'If Secure_Access_Snap("FUNCTION") = "Report_Benefits" Then gSec_Rpt_Benefits= SECURE_ACCESS_SNAP("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Dependents" Then gSec_Rpt_Dependents = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Compensatory_Time" Then gSec_Rpt_Compensatory_Time = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Cost_Of_Employment" Then gSec_Rpt_Cost_Of_Employment = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Emergecy_Contacts" Then gSec_Rpt_Emergecy_Contacts = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Employee_Labels" Then gSec_Rpt_Employee_Labels = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Job_List" Then gSec_Rpt_Job_List = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Profiles" Then gSec_Rpt_Profiles = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Entitlements" Then gSec_Rpt_Entitlements = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Follow_Ups" Then gSec_Rpt_Follow_Ups = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Home_Address" Then gSec_Rpt_Home_Address = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Salary_Performance" Then gSec_Rpt_Salary_Performance = Secure_Access_Snap("ACCESSABLE")
    'Ticket #27795 - Friesens Corporation
    If Secure_Access_Snap("FUNCTION") = "Report_Staff_Profile" Then gSec_Rpt_Staff_Profile = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Seniority" Then gSec_Rpt_Seniority = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Skills" Then gSec_Rpt_Skills = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Languages" Then gSec_Rpt_Languages = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Telephone_Extensions" Then gSec_Rpt_Telephone_Extensions = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Associations" Then gSec_Rpt_Associations = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Master_Attendance" Then gSec_Rpt_Master_Attendance = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Bonus_Attendance" Then gSec_Rpt_Bonus_Attendance = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Calendar_Attendance" Then gSec_Rpt_Calendar_Attendance = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Costed_Attendance" Then gSec_Rpt_Costed_Attendance = Secure_Access_Snap("ACCESSABLE")
    
    If Secure_Access_Snap("FUNCTION") = "Report_Master_Benefits" Then gSec_Rpt_Master_Benefits = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Master_Division" Then gSec_Rpt_Master_Division = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Master_DolEnt" Then gSec_Rpt_Master_DolEnt = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Master_Termination" Then gSec_Rpt_Master_Termination = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Master_Formal_Education" Then gSec_Rpt_Master_Formal_Education = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Training_Plan" Then gSec_Rpt_Training_Plan = Secure_Access_Snap("ACCESSABLE") 'Ticket #21709
    If Secure_Access_Snap("FUNCTION") = "Report_Master_Job" Then gSec_Rpt_Master_Job = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Master_OtherEarn" Then gSec_Rpt_Master_OtherEarn = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Master_Passwords" Then gSec_Rpt_Master_Passwords = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Master_Salaries" Then gSec_Rpt_Master_Salaries = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Master_Edu_Seminars" Then gSec_Rpt_Master_Education_Seminars = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Master_Table_Codes" Then gSec_Rpt_Master_Table_Codes = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Heatlh_Safety" Then gSec_Rpt_Heatlh_Safety = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Hourly_Entitlements" Then gSec_Rpt_Master_HourEnt = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Employee_Turnover" Then gSec_Rpt_Turnover = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Counselling" Then gSec_Rpt_Counselling = Secure_Access_Snap("ACCESSABLE")
    'Release 8.1
    If Secure_Access_Snap("FUNCTION") = "Report_DocumentType" Then gSec_Rpt_DocumentType = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Emergency_Leave" Then gSec_Rpt_Emergency_Leave = Secure_Access_Snap("ACCESSABLE")
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Report_External_Hire") Then gSec_Rpt_External_Hire = Secure_Access_Snap("ACCESSABLE")
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Report_Internal_Hire") Then gSec_Rpt_Internal_Hire = Secure_Access_Snap("ACCESSABLE")
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Report_Key_Workforce") Then gSec_Rpt_Key_Workforce = Secure_Access_Snap("ACCESSABLE")
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Report_Manpower_Plan") Then gSec_Rpt_Manpower_Plan = Secure_Access_Snap("ACCESSABLE")
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Report_Staff_Management") Then gSec_Rpt_Staff_Management = Secure_Access_Snap("ACCESSABLE")
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Report_WC_Time") Then gSec_Rpt_WC_Time = Secure_Access_Snap("ACCESSABLE")
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Report_WC_Work") Then gSec_Rpt_WC_Work = Secure_Access_Snap("ACCESSABLE")
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Report_Paid_Sick") Then gSec_Rpt_Paid_Sick = Secure_Access_Snap("ACCESSABLE")
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Report_User_Defined_Table") Then gSec_Rpt_User_Defined_Table = Secure_Access_Snap("ACCESSABLE")
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Report_Future_Entitlement") Then gsec_rpt_Future_Entitlement = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Employee_Flags" Then gSec_Rpt_Employee_Flags = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Temp_CrossTraining" Then gSec_Rpt_Temp_CrossTraining = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Required_Course_Hist" Then gSec_Rpt_Req_Course_Hist = Secure_Access_Snap("ACCESSABLE")
    
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Report_Email_Address") Then gSec_Rpt_EmailAddress = Secure_Access_Snap("ACCESSABLE")
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Report_LOA") Then gSec_Rpt_LOA = Secure_Access_Snap("ACCESSABLE")
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Report_POE") Then gSec_Rpt_POE = Secure_Access_Snap("ACCESSABLE")
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Report_SINSSN") Then gSec_Rpt_SINSSN = Secure_Access_Snap("ACCESSABLE")
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Report_Succession") Then gSec_Rpt_Succession = Secure_Access_Snap("ACCESSABLE")
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Report_Gap_Analysis") Then gSec_Rpt_GapAnalysis = Secure_Access_Snap("ACCESSABLE")
    
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Report_GL_Distribution") Then gSec_Rpt_GLDistribution = Secure_Access_Snap("ACCESSABLE")
    
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Report_Attendance_Hist") Then gSec_Rpt_Attendance_Hist = Secure_Access_Snap("ACCESSABLE")
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Report_AttWrkSch_Descrepancy") Then gSec_Rpt_AttWrkSch_Descrepancy = Secure_Access_Snap("ACCESSABLE")
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Report_Comments") Then gSec_Rpt_Comments = Secure_Access_Snap("ACCESSABLE")
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Report_Employee_Hist") Then gSec_Rpt_Employee_Hist = Secure_Access_Snap("ACCESSABLE")
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Report_Payroll_Transactions") Then gSec_Rpt_Payroll_Trans = Secure_Access_Snap("ACCESSABLE")
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Report_Environmental_Serv") Then gSec_Rpt_EnviroServices = Secure_Access_Snap("ACCESSABLE")
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Report_ESSReq_TransAudit") Then gSec_Rpt_ESSReq_TransAudit = Secure_Access_Snap("ACCESSABLE")
    
    'Release 8.0 - Ticket #22682
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Report_Employee_Dates") Then gSec_Rpt_Employee_Dates = Secure_Access_Snap("ACCESSABLE")
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Report_Length_of_Service") Then gSec_Rpt_Length_Of_Service = Secure_Access_Snap("ACCESSABLE")
    
    'Ticket #24663
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Form_Attendance_SignIn") Then gSec_RptF_Attendance_SignIn = Secure_Access_Snap("ACCESSABLE")
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Form_ATT_Discipline") Then gSec_RptF_ATT_Discipline = Secure_Access_Snap("ACCESSABLE")
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Form_COC_Discipline") Then gSec_RptF_COC_Discipline = Secure_Access_Snap("ACCESSABLE")
    
    'Ticket #26576 - WDGPHU - Flex Time report
    If UCase(Secure_Access_Snap("FUNCTION")) = UCase("Report_FlexTime") Then gSec_Rpt_FlexTime = Secure_Access_Snap("ACCESSABLE")
    
    If Secure_Access_Snap("FUNCTION") = "Report_Friesens_IWantToKnowYou" Then gSec_Rpt_Friesens_IWantToKnowYou = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Friesens_ITHireForm" Then gSec_Rpt_Friesens_ITHireForm = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Friesens_ITNoticeOfChange" Then gSec_Rpt_Friesens_ITNoticeOfChange = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Friesens_NoticeOfChange" Then gSec_Rpt_Friesens_NoticeOfChange = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Friesens_PerfImproveActionPlan" Then gSec_Rpt_Friesens_PerfImproveActionPlan = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Friesens_PerformanceReviewRpt" Then gSec_Rpt_Friesens_PerformanceReviewRpt = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Friesens_SeparationRpt" Then gSec_Rpt_Friesens_SeparationRpt = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Friesens_TerminationRpt" Then gSec_Rpt_Friesens_TerminationRpt = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Friesens_UpdateMeetingRpt" Then gSec_Rpt_Friesens_UpdateMeetingRpt = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Friesens_WarningRpt" Then gSec_Rpt_Friesens_WarningRpt = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_AffirmAction" Then gSec_Rpt_AffirmAction = Secure_Access_Snap("ACCESSABLE")

    If Secure_Access_Snap("FUNCTION") = "Report_WorkSchedule" Then gSec_Rpt_Work_Schedule = Secure_Access_Snap("ACCESSABLE")
    
    If Secure_Access_Snap("FUNCTION") = "ProductLine_Operation_Update" Then gSec_Upd_Productline_Operation = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "ProductLine_Operation_Inquiry" Then gSec_Inq_Productline_Operation = Secure_Access_Snap("ACCESSABLE")
    
    If Secure_Access_Snap("FUNCTION") = "LinamarSkills_Update" Then gSec_Upd_LinamarSkills = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "LinamarSkills_Inquiry" Then gSec_Inq_LinamarSkills = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "WHSCC_ASL" Then
        gSec_Upd_WHSCC_ASL = Secure_Access_Snap("Maintainable")
        gSec_Inq_WHSCC_ASL = Secure_Access_Snap("ACCESSABLE")
    End If
    If Secure_Access_Snap("FUNCTION") = "WHSCC_BUDPOS" Then
        gSec_Upd_WHSCC_BUDPOS = Secure_Access_Snap("Maintainable")
        gSec_Inq_WHSCC_BUDPOS = Secure_Access_Snap("ACCESSABLE")
    End If
    If Secure_Access_Snap("FUNCTION") = "WHSCC_USB" Then
        gSec_Upd_WHSCC_USB = Secure_Access_Snap("Maintainable")
        gSec_Inq_WHSCC_USB = Secure_Access_Snap("ACCESSABLE")
    End If
    If Secure_Access_Snap("FUNCTION") = "WHSCC_PLAN_ESTABLISMNET_REPORT" Then
        gSec_Rpt_WHSCC_PLAN_ESTABLISMNET = Secure_Access_Snap("ACCESSABLE")
    End If
    'For WHSCC
    
    'For Samuel
    If Secure_Access_Snap("FUNCTION") = "SAM_Show_CustomFeatures" Then
        gSec_SAM_Show_CustomFeatures = Secure_Access_Snap("ACCESSABLE")
    End If
    
    'Overtime
    If Secure_Access_Snap("FUNCTION") = "Report_Overtime_Lost_Hours" Or Secure_Access_Snap("FUNCTION") = "Report_Overtime_Bank" Or Secure_Access_Snap("FUNCTION") = "Overtime_Master" Then
        gSec_Upd_Ovt_Overview = Secure_Access_Snap("Maintainable")
        gSec_Inq_Ovt_Overview = Secure_Access_Snap("ACCESSABLE")
    End If
    
    If Secure_Access_Snap("FUNCTION") = "Overtime_Master_Update" Then gSec_Upd_Ovt_Master = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Overtime_Master_Inquiry" Then gSec_Inq_Ovt_Master = Secure_Access_Snap("ACCESSABLE")
    
    If Secure_Access_Snap("FUNCTION") = "Report_Overtime_Bank" Then gSec_Rpt_Ovt_Bank = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Report_Overtime_Lost_Hours" Then gSec_Rpt_Ovt_Lost_Hours = Secure_Access_Snap("ACCESSABLE")
    
    If Secure_Access_Snap("FUNCTION") = "ADP_Data_Update" Then gSec_Upd_ADP_Data = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "ADP_Data_Inquiry" Then gSec_Inq_ADP_Data = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "CourseCodeMaster_Update" Then gSec_Upd_CourseCodeMaster = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "CourseCodeMaster_Inquiry" Then gSec_Inq_CourseCodeMaster = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "BudgetedManpower_Update" Then gSec_Upd_BudgetedMP = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "BudgetedManpower_Inquiry" Then gSec_Inq_BudgetedMP = Secure_Access_Snap("ACCESSABLE")
    
    'Ticket #22220
    If Secure_Access_Snap("FUNCTION") = "WorkScheduleRule_Update" Then gSec_Upd_WorkSchRule = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "WorkScheduleRule_Inquiry" Then gSec_Inq_WorkSchRule = Secure_Access_Snap("ACCESSABLE")
    
    'Ticket #22541
    If Secure_Access_Snap("FUNCTION") = "DashboardSetup_Update" Then gSec_Upd_DashboardRule = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "DashboardSetup_Inquiry" Then gSec_Inq_DashboardRule = Secure_Access_Snap("ACCESSABLE")
    
    If Secure_Access_Snap("FUNCTION") = "RequiredCourses_Update" Then gSec_Upd_ReqCourses = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "RequiredCourses_Inquiry" Then gSec_Inq_ReqCourses = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "BudgetedPosition_Update" Then gSec_Upd_BudgetedPos = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "BudgetedPosition_Inquiry" Then gSec_Inq_BudgetedPos = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "ApplicationProcess_Update" Then gSec_Upd_AppProcess = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "ApplicationProcess_Inquiry" Then gSec_Inq_AppProcess = Secure_Access_Snap("ACCESSABLE")
    
    'Ticket #25015 - Macaulay
    If Secure_Access_Snap("FUNCTION") = "AddPayrollIDData_Update" Then gSec_Upd_AddPayrollIDData = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "AddPayrollIDData_Inquiry" Then gSec_Inq_AddPayrollIDData = Secure_Access_Snap("ACCESSABLE")
    
    If Secure_Access_Snap("FUNCTION") = "Rehire_Update" Then gSec_Upd_Rehire = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Rehire_Inquiry" Then gSec_Inq_Rehire = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "EnterLeave_Update" Then gSec_Upd_EnterLeave = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "EnterLeave_Inquiry" Then gSec_Inq_EnterLeave = Secure_Access_Snap("ACCESSABLE")
    
    
    If Secure_Access_Snap("FUNCTION") = "HS_ClaimMed_Update" Then gSec_Upd_HSClaimMed = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "HS_ClaimMed_Inquiry" Then gSec_Inq_HSClaimMed = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "HS_Contacts_Update" Then gSec_Upd_HSContacts = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "HS_Contacts_Inquiry" Then gSec_Inq_HSContacts = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "HS_Cost_Update" Then gSec_Upd_HSCost = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "HS_Cost_Inquiry" Then gSec_Inq_HSCost = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "HS_CorrectAction_Update" Then gSec_Upd_HSCorrectiveAct = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "HS_CorrectAction_Inquiry" Then gSec_Inq_HSCorrectiveAct = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "HS_RootCause_Update" Then gSec_Upd_HSRootCause = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "HS_RootCause_Inquiry" Then gSec_Inq_HSRootCause = Secure_Access_Snap("ACCESSABLE")
    
    If glbWSIBModule Then   'WSIB Form 7 - Billable Module
        If Secure_Access_Snap("FUNCTION") = "HS_W7CompanyMaster_Update" Then gSec_Upd_HSW7CmpMst = Secure_Access_Snap("ACCESSABLE")
        If Secure_Access_Snap("FUNCTION") = "HS_W7CompanyMaster_Inquiry" Then gSec_Inq_HSW7CmpMst = Secure_Access_Snap("ACCESSABLE")
        If Secure_Access_Snap("FUNCTION") = "HS_W7Injury_Update" Then gSec_Upd_HSW7Injury = Secure_Access_Snap("ACCESSABLE")
        If Secure_Access_Snap("FUNCTION") = "HS_W7Injury_Inquiry" Then gSec_Inq_HSW7Injury = Secure_Access_Snap("ACCESSABLE")
        
        'Form 9
        If Secure_Access_Snap("FUNCTION") = "HS_WF9_Update" Then gSec_Upd_HSWF9 = Secure_Access_Snap("ACCESSABLE")
        If Secure_Access_Snap("FUNCTION") = "HS_WF9_Inquiry" Then gSec_Inq_HSWF9 = Secure_Access_Snap("ACCESSABLE")
    End If
    
    'Mostafa - Attedance Group Code Matrix
    If Secure_Access_Snap("FUNCTION") = "Attendance_Group_Code_Matrix_Update" Then gSec_Upd_Attendance_Group_Code_Matrix = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Attendance_Group_Code_Matrix_Inquiry" Then gSec_Inq_Attendance_Group_Code_Matrix = Secure_Access_Snap("ACCESSABLE")
    
    If Secure_Access_Snap("FUNCTION") = "Job_Files_Attachment_Update" Then gSec_Upd_Job_Files_Attachment = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Job_Files_Attachment_Inquiry" Then gSec_Inq_Job_Files_Attachment = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Temp_Cross_Training_Update" Then gSec_Upd_Temp_Cross_Training = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Temp_Cross_Training_Inquiry" Then gSec_Inq_Temp_Cross_Training = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Training_List_Update" Then gSec_Upd_Training_List = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "Training_List_Inquiry" Then gSec_Inq_Training_List = Secure_Access_Snap("ACCESSABLE")
    
    'Ticket #30508 - Applicant Tracking Enhancement
    If Secure_Access_Snap("FUNCTION") = "App_LetterPosType_Update" Then gSec_Upd_LettersPosType = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "App_LetterPosType_Inquiry" Then gSec_Inq_LettersPosType = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "App_FormWorkflow_Update" Then gSec_Upd_AppFormWorkFlow = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "App_FormWorkflow_Inquiry" Then gSec_Inq_AppFormWorkFlow = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "App_FormDefaults_Update" Then gSec_Upd_AppFormDefaults = Secure_Access_Snap("ACCESSABLE")
    If Secure_Access_Snap("FUNCTION") = "App_FormDefaults_Inquiry" Then gSec_Inq_AppFormDefaults = Secure_Access_Snap("ACCESSABLE")
    
    
    Secure_Access_Snap.MoveNext
    
Loop

Do While gSec_Upd_Master_Table.count > 0
    gSec_Upd_Master_Table.Remove 1
Loop
Do While gSec_Inq_Master_Table.count > 0
    gSec_Inq_Master_Table.Remove 1
Loop

Secure_Access_Snap.Close

'????Ticket #24808 -  Retrieve Template's security profile if User is Template based
If xSecTemplate = "" Or xSecTemplate = "TEMPLATE" Then
    Secure_Access_Snap.Open "select * from HR_SECURE_ACCESS WHERE USERID='" & Replace(UserID, "'", "''") & "' AND CODENAME IS NOT NULL", gdbAdoIhr001, adOpenStatic
Else
    Secure_Access_Snap.Open "select * from HR_SECURE_ACCESS WHERE USERID='" & Replace(xSecTemplate, "'", "''") & "' AND CODENAME IS NOT NULL", gdbAdoIhr001, adOpenStatic
End If
Do Until Secure_Access_Snap.EOF
    xStr = Trim(Secure_Access_Snap("CODENAME"))
    gSec_Upd_Master_Table.Add IIf(Secure_Access_Snap("Maintainable") <> 0, True, False), xStr
    gSec_Inq_Master_Table.Add IIf(Secure_Access_Snap("ACCESSABLE") <> 0, True, False), xStr
    Secure_Access_Snap.MoveNext
Loop
Secure_Access_Snap.Close

modSecurity_Check = True

Exit Function


Secure_Err:
If Err = 3265 Or Err = 3704 Or Err = -2147217904 Or Err = -2147217900 Then
    MsgBox "      The database has not been converted to the most recent release." & Chr(10) & Chr(10) & _
    "Please contact HR Systems Strategies Inc. Support Department for assistance."
    End
Else
    glbFrmCaption$ = "Module - Security"
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Read Security Rcd", "Security", "Select")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
        Resume Next
    End If
End If
End Function

Sub modSetModControls(frmName As Form, intData%)
' referenced in cancel, deletes, finds and form loads
' not in new/modify. this is passed if data is in
'control 1, and form name - if data found it will
're enable/disable delete,modify and print controls
' which are only valid if there is data on the screen!
If Not intData Then
    frmName.cmdDelete.Enabled = False
    frmName.cmdModify.Enabled = False
    'frmName.cmdPrint.Enabled = False
Else
    frmName.cmdDelete.Enabled = True
    frmName.cmdModify.Enabled = True
    'frmName.cmdPrint.Enabled = True
End If

End Sub

Sub NukeEE_SerialNo(EEID&) 'Ticket #23116 Franks 01/23/2013 for WFC
Dim snapEETables As New ADODB.Recordset
Dim SQLQ As String, TabName$
Dim EEIDAlias$
Dim rsSE As New ADODB.Recordset
Dim xUserID As String
'On Error GoTo NukeEE_Err

SQLQ = "SELECT * FROM INFO_HR_TABLES "
SQLQ = SQLQ & " WHERE Employee_Keyed <>0"
SQLQ = SQLQ & " AND TERMINATION_TABLE=0"
SQLQ = SQLQ & " AND (SERIAL = '" & glbCompSerial & "') "

snapEETables.Open SQLQ, gdbAdoIhr001, adOpenStatic

If snapEETables.RecordCount < 1 Then Exit Sub
snapEETables.MoveFirst

While Not snapEETables.EOF
    TabName$ = snapEETables("Table_Name")
    If UCase(Right(TabName$, 3)) <> "WRK" Then
        If Not IsNull(snapEETables("EMPNBR_Alias")) Then
            EEIDAlias$ = snapEETables("EMPNBR_Alias")
        Else
            EEIDAlias$ = ""
        End If
      Call NukeEERows(TabName$, EEIDAlias$, EEID&)
    End If
    snapEETables.MoveNext
Wend

snapEETables.Close

Exit Sub
End Sub


Sub NukeEE(EEID&)

Dim snapEETables As New ADODB.Recordset
Dim SQLQ As String, TabName$
Dim EEIDAlias$


On Error GoTo NukeEE_Err
Dim rsSE As New ADODB.Recordset
Dim xUserID As String
rsSE.Open "SELECT USERID FROM HR_SECURE_BASIC WHERE EMPNBR=" & EEID&, gdbAdoIhr001, adOpenStatic
If Not rsSE.EOF Then
    xUserID = rsSE("USERID")
    Call NukeUSERID(xUserID)
End If
rsSE.Close


SQLQ = "SELECT * FROM INFO_HR_TABLES "
SQLQ = SQLQ & " WHERE Employee_Keyed <>0"
SQLQ = SQLQ & " AND TERMINATION_TABLE=0"
'Ticket #20415 - Add Serial # to the select statement so custom tables also gets employee # changed.
'Serial 9999 is by default for all standard info:HR table.
'SQLQ = SQLQ & " AND (SERIAL = 'S/N - 9999W' OR SERIAL = '" & glbCompSerial & "')"
'Ticket #20893 Franks 09/02/2011 - only remove data for the standard INFO:HR tables
SQLQ = SQLQ & " AND (SERIAL = 'S/N - 9999W' OR SERIAL IS NULL)"

snapEETables.Open SQLQ, gdbAdoIhr001, adOpenStatic

If snapEETables.RecordCount < 1 Then Exit Sub
snapEETables.MoveFirst

While Not snapEETables.EOF
    TabName$ = snapEETables("Table_Name")
    If UCase(Right(TabName$, 3)) <> "WRK" Then
        If Not IsNull(snapEETables("EMPNBR_Alias")) Then
            EEIDAlias$ = snapEETables("EMPNBR_Alias")
        Else
            EEIDAlias$ = ""
        End If
      Call NukeEERows(TabName$, EEIDAlias$, EEID&)
    End If
    snapEETables.MoveNext
Wend

snapEETables.Close
Call UpdVacTimeRequest(EEID&, "D")
Call UpdTimesheetEMPNum(EEID&, "D") 'Ticket #16654

If glbCompSerial = "S/N - 2279W" Then  'Friesens Corporation
    SQLQ = "DELETE FROM HR_PERFORM_FRIESEN WHERE PH_EMPNBR =" & EEID & " "
    gdbAdoIhr001.Execute SQLQ
End If


Exit Sub

NukeEE_Err:
glbFrmCaption$ = "Delete Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "HR_TABLES Error", "TabName$", "Search")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Sub

Sub NukeEERows(TabName$, EEIDAlias$, EEID&)
' returns number of records found for ee in table

Dim Rows%, SQLQ As String
Dim gdbESS As New ADODB.Connection


On Error GoTo NukeEERowsErr

If EEIDAlias$ <> "" Then
    SQLQ = "DELETE FROM " & TabName
    
    If InStr(EEIDAlias$, "USER") > 0 Then
        SQLQ = SQLQ & " WHERE " & EEIDAlias & " = '" & EEID & "'"
    Else
        SQLQ = SQLQ & " WHERE " & EEIDAlias & " = " & EEID
    End If
End If

If SQLQ <> "" Then
    If glbtermopen Or glbRest Then
        gdbAdoIhr001X.Execute SQLQ
    Else
        'Users were getting error when the ESS mdb was not there for MS Access users.
        If Not glbSQL And Not glbOracle And (TabName = "HR_TIMESHEET" Or TabName = "HR_TIMESHEET_MODS" Or TabName = "HR_VACTIMEOFF_REQ" Or TabName = "HR_VACTIMEOFF_REQ_ARCHIVE") Then
            If gdbESS = "" Then
                gdbESS.Open Replace(glbAdoIHRDB, "IHR001", "IHRESS")
            End If
        End If
    
        If Not glbSQL And Not glbOracle And (TabName = "HR_TIMESHEET" Or TabName = "HR_TIMESHEET_MODS" Or TabName = "HR_VACTIMEOFF_REQ" Or TabName = "HR_VACTIMEOFF_REQ_ARCHIVE") Then
            If gdbESS <> "" Then
                gdbESS.Execute SQLQ
            End If
        Else
            gdbAdoIhr001.Execute SQLQ
        End If
        
    End If
End If

Exit Sub

NukeEERowsErr:

If Err.Number = -2147467259 Then
    gdbESS = ""
    Resume Next
End If

glbFrmCaption$ = "Nuke Rows"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Delete EE Rows", TabName$, "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Sub

Sub NukeUSERID(zUSERID As String)
Dim Rows%, SQLQ As String


On Error GoTo NukeUSERIDErr


SQLQ = "DELETE FROM HR_SECURE_ACCESS WHERE USERID = '" & Replace(zUSERID, "'", "''") & "'"
gdbAdoIhr001.Execute SQLQ
SQLQ = "DELETE FROM HRPASDEP WHERE PD_USERID = '" & Replace(zUSERID, "'", "''") & "'"
gdbAdoIhr001.Execute SQLQ
SQLQ = "DELETE FROM HR_EMAIL WHERE EM_USERID = '" & Replace(zUSERID, "'", "''") & "'"
gdbAdoIhr001.Execute SQLQ
If glbLinamar Then
    SQLQ = "DELETE FROM LN_DOORS WHERE USERID = '" & Replace(zUSERID, "'", "''") & "'"
    gdbAdoIhr001.Execute SQLQ
End If

Exit Sub

NukeUSERIDErr:
glbFrmCaption$ = "Nuke Rows"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Delete EE Rows", "USERID", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Sub

Sub NukeHSCosts(InNo As Long)
Dim SQLQ As String

On Error GoTo NukeHSCosts_Err

SQLQ = "DELETE FROM HROHSCOS"
SQLQ = SQLQ & " WHERE CC_CASE = " & InNo & ""

gdbAdoIhr001.Execute SQLQ

Exit Sub

NukeHSCosts_Err:
glbFrmCaption$ = "Inc/Injury"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Nuke HS Costs", "HROHSCOS", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Sub

'Ticket #22682 - Delete the rest of the incident related records
Sub NukeHSRootCauses(InNo As Long, xEmpNo)
    Dim SQLQ As String
    
    On Error GoTo NukeHSRootCauses_Err
    
    SQLQ = "DELETE FROM HR_OHS_ROOT_CAUSES"
    SQLQ = SQLQ & " WHERE RC_Case = " & InNo & ""
    SQLQ = SQLQ & " AND RC_Empnbr = " & xEmpNo
    gdbAdoIhr001.Execute SQLQ
    
    Exit Sub
    
NukeHSRootCauses_Err:
    glbFrmCaption$ = "Inc/Injury"
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Nuke HS Root Causes", "HR_OHS_ROOT_CAUSES", "Delete")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
        Resume Next
    End If

End Sub

'Ticket #22682 - Delete the rest of the incident related records
Sub NukeHSContacts(InNo As Long, xEmpNo)
    Dim SQLQ As String
    
    On Error GoTo NukeHSContacts_Err
    
    SQLQ = "DELETE FROM HR_OHS_CONTACT"
    SQLQ = SQLQ & " WHERE CT_Case = " & InNo & ""
    SQLQ = SQLQ & " AND CT_Empnbr = " & xEmpNo
    gdbAdoIhr001.Execute SQLQ
    
    Exit Sub
    
NukeHSContacts_Err:
    glbFrmCaption$ = "Inc/Injury"
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Nuke HS Contacts", "HR_OHS_CONTACT", "Delete")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
        Resume Next
    End If

End Sub

'Ticket #22682 - Delete the rest of the incident related records
Sub NukeHSCorrective(InNo As Long, xEmpNo)
    Dim SQLQ As String
    
    On Error GoTo NukeHSCorrective_Err
    
    SQLQ = "DELETE FROM HR_OHS_CORRECTIVE"
    SQLQ = SQLQ & " WHERE CR_Case = " & InNo & ""
    SQLQ = SQLQ & " AND CR_Empnbr = " & xEmpNo
    gdbAdoIhr001.Execute SQLQ
    
    Exit Sub
    
NukeHSCorrective_Err:
    glbFrmCaption$ = "Inc/Injury"
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Nuke HS Corrective", "HR_OHS_CORRECTIVE", "Delete")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
        Resume Next
    End If

End Sub

'Ticket #22682 - Delete the rest of the incident related records
Sub NukeHSForm9(InNo As Long, xEmpNo)
    Dim SQLQ As String
    
    On Error GoTo NukeHSForm9_Err
    
    SQLQ = "DELETE FROM HR_OHS_FORM9"
    SQLQ = SQLQ & " WHERE F9_CASE = " & InNo & ""
    SQLQ = SQLQ & " AND F9_EMPNBR = " & xEmpNo
    gdbAdoIhr001.Execute SQLQ
    
    Exit Sub
    
NukeHSForm9_Err:
    glbFrmCaption$ = "Inc/Injury"
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Nuke HS Form 9", "HR_OHS_FORM9", "Delete")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
        Resume Next
    End If

End Sub

'Ticket #22682 - Delete the rest of the incident related records
Sub NukeHSAttachment(InNo As Long, xEmpNo)
    Dim SQLQ As String
    
    On Error GoTo NukeHSAttachment_Err
        
    If Not gsAttachment_DB Then
        Exit Sub
    End If
    
    SQLQ = "DELETE FROM HRDOC_HEALTH_SAFETY"
    SQLQ = SQLQ & " WHERE DE_CASE = " & InNo & ""
    SQLQ = SQLQ & " AND DE_EMPNBR = " & xEmpNo
    gdbAdoIhr001_DOC.Execute SQLQ
    
    Exit Sub
    
NukeHSAttachment_Err:
    glbFrmCaption$ = "Inc/Injury"
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Nuke HS Attachment DOC", "HRDOC_HEALTH_SAFETY", "Delete")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
        Resume Next
    End If

End Sub

Function PrtForm(Captn As String, FrmNam As Form)
Dim DgDef, Msg As String, Response%, Title$ '  variables.

Title = "Print Screen?"

Msg = "Do You Wish to Print " & Chr(10)
Msg = Msg & Captn & " Screen?"

DgDef = MB_YESNOCANCEL + MB_ICONQUESTION + MB_DEFBUTTON2    ' Describe dialog.

Response = MsgBox(Msg, DgDef, Title)    ' Get user response.
Select Case Response
    Case IDNO
        PrtForm = True
        Exit Function
    Case IDYES
        PrtForm = True
        FrmNam.PrintForm
    Case Else
        PrtForm = False
        
End Select

End Function

Function setCompInfo(CompNo)
Dim SQLQ As String, PARCO_Snap As New ADODB.Recordset
Dim rsUser As New ADODB.Recordset
setCompInfo = False    ' returns true if found records
On Error GoTo Comp_Err

SQLQ = "SELECT * FROM HRPARCO WHERE PC_CO = '" & CompNo & "'"


PARCO_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic
'adOpenForwardOnly
If PARCO_Snap.BOF And PARCO_Snap.EOF Then
    setCompInfo = False
Else
    setCompInfo = True    ' returns true if found records
    glbCompName = PARCO_Snap("PC_NAME")
    glbCompLvl = PARCO_Snap("PC_LVLNBR")
    glbCompSerial = PARCO_Snap("PC_SERIAL")
    glbCompEdFrom = PARCO_Snap("PC_FDATE")
    glbCompEdTo = PARCO_Snap("PC_TDATE")
    glbCompEdFromS = PARCO_Snap("PC_FDATES")
    glbCompEdToS = PARCO_Snap("PC_TDATES")
    glbCompEntSick$ = PARCO_Snap("PC_SICKENT")
    glbCompEntVac$ = PARCO_Snap("PC_VACENT")
    glbCompWDate$ = PARCO_Snap("PC_WDATE")
    glbEntOutStanding$ = PARCO_Snap("PC_ENTOUT")
    glbEntOutStandingS$ = PARCO_Snap("PC_ENTOUTS")
    glbMulti = IIf(PARCO_Snap("PC_MULTI") <> 0, True, False)
    glbMultiGrid = PARCO_Snap("PC_MULTIGRID")
    If IsNull(PARCO_Snap("PC_COUNTRY")) Then
        glbCountry = "CANADA"
    Else
        glbCountry = PARCO_Snap("PC_COUNTRY")
    End If
    glbCompDecHR = PARCO_Snap("PC_DECHR")
    PARCO_Snap.Close
    
    'Ticket #29230 - Daily Entitlement Setup & Update Flag
    If glbCompEntVac$ = "D" Then
        glbCompEntVacDaily = True
    Else
        glbCompEntVacDaily = False
    End If
End If

'Frank 03/11/04 Ticket# 5733
'Check User's Country, if it exists, then replace glbCountry with User's Country
SQLQ = "SELECT USERID, COUNTRY FROM HR_SECURE_BASIC WHERE USERID = '" & Replace(glbUserID, "'", "''") & "' "
rsUser.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsUser.EOF Then
    If Not IsNull(rsUser("COUNTRY")) Then
        glbCountry = rsUser("COUNTRY")
    End If
End If
rsUser.Close

If glbCompSerial = "S/N - 2309W" Then ' For Linamar
    glbLinamar = True
Else
    glbLinamar = False
End If
If glbCompSerial = "S/N - 2336W" Then ' For WHSCC
    glbWHSCC = True
Else
    glbWHSCC = False
End If
If glbCompSerial = "S/N - 2308W" Then ' For Axxent
    glbAxxent = True
Else
    glbAxxent = False
End If
If glbCompSerial = "S/N - 2276W" Then ' For City of Niagara Fulls
    glbNiagaraFulls = True
Else
    glbNiagaraFulls = False
End If
If glbCompSerial = "S/N - 2292W" Then ' For County of Elgin
    glbCElgin = True
Else
    glbCElgin = False
End If
If glbCompSerial = "S/N - 2322W" Then ' For Guelph-Wellington
    glbGuelph = True
    'Ticket #24677 - SQL Conversion
    'If Dir$(glbSN2322) <> "" Then
        If gdbAdoSN2322.State = adStateOpen Then gdbAdoSN2322.Close
        gdbAdoSN2322.Mode = adModeReadWrite
        'Ticket #24677 - SQL Conversion
        'gdbAdoSN2322.Open glbAdoSN2322
        gdbAdoSN2322.Open glbAdoIHRDB
    'End If
Else
    glbGuelph = False
End If
If glbCompSerial = "S/N - 2291W" Then ' For Syndesis
    glbSyndesis = True
Else
    glbSyndesis = False
End If
If glbCompSerial = "S/N - 2323W" Then ' For County of Brant
    glbCBrant = True
    'If Left(glbIHRDBB, 8) <> "00000000" Then
    '    If Dir$(glbIHRDBB) <> "" Then
    '        If gdbAdoIhr001B.State = adStateOpen Then gdbAdoIhr001B.Close
    '        gdbAdoIhr001B.Mode = adModeReadWrite
    '        gdbAdoIhr001B.Open glbAdoIHRDBB
    '    End If
    'End If
    'Ticket #23810 Franks 06/17/2013
    If gdbAdoIhr001B.State = adStateOpen Then gdbAdoIhr001B.Close
    gdbAdoIhr001B.Mode = adModeReadWrite
    gdbAdoIhr001B.Open glbAdoIHRDB
Else
    glbCBrant = False
End If
If glbCompSerial = "S/N - 2343W" Then
    glbOttawaCCAC = True
Else
    glbOttawaCCAC = False
End If
If glbCompSerial = "S/N - 2326W" Then
    glbSoroc = True
Else
    glbSoroc = False
End If
If glbCompSerial = "S/N - 2341W" Then
    glbDundasACL = True
Else
    glbDundasACL = False
End If
If glbCompSerial = "S/N - 2226W" Then
    glbBrantCount = True
Else
    glbBrantCount = False
End If
If glbCompSerial = "S/N - 2355W" Then
    glbLambton = True
Else
    glbLambton = False
End If
If glbCompSerial = "S/N - 2351W" Then 'Burlington Tech
    glbBurlTech = True
Else
    glbBurlTech = False
End If
If glbCompSerial = "S/N - 2233W" Then 'Simona - Leeds Grenville CAS ticket #14890
    glbCwis = True
Else
    glbCwis = False
End If

'Get setup values from HRPREFERENCE table

Exit Function

Comp_Err:
glbFrmCaption$ = "Set Company Info"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Company Snap", "PARCO", "SELECT")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Public Sub setPreference()
Dim rsPrefer As New ADODB.Recordset
    rsPrefer.Open "SELECT * FROM HRPREFERENCE", gdbAdoIhr001, adOpenStatic
    Do While Not rsPrefer.EOF
        If glbSQL Or glbOracle Then
            If rsPrefer("HP_FUN_NAME") = "ATTACHMENT" Then
                gsAttachment_DB = rsPrefer("HP_ENABLED")
                If gsAttachment_DB Then
                    If gdbAdoIhr001_DOC.State = adStateOpen Then gdbAdoIhr001_DOC.Close
                    On Error GoTo DOC_HNDLR
                    gdbAdoIhr001_DOC.CommandTimeout = 600
                    gdbAdoIhr001_DOC.Open glbAdoIHRDB_DOC
                    On Error Resume Next
                End If
            End If
        End If
        If rsPrefer("HP_FUN_NAME") = "EMAIL_SENDING" Then
            gsEMAIL_SENDING = rsPrefer("HP_ENABLED")
        End If
        If rsPrefer("HP_FUN_NAME") = "COMPA-RATIO" Then
            gsCompaRatio = rsPrefer("HP_ENABLED")
        End If
        If rsPrefer("HP_FUN_NAME") = "EMAIL_ONNEWHIRE" Then
            gsEMAIL_ONNEWHIRE = rsPrefer("HP_ENABLED")
        End If
        If rsPrefer("HP_FUN_NAME") = "EMAIL_ONPOSITION" Then 'Ticket #21444 Franks 02/10/2012
            gsEMAIL_ONPOSITION = rsPrefer("HP_ENABLED")
        End If
        If rsPrefer("HP_FUN_NAME") = "EMAIL_ONSALARY" Then
            gsEMAIL_ONSALARY = rsPrefer("HP_ENABLED")
        End If
        If rsPrefer("HP_FUN_NAME") = "EMAIL_ONBENEFIT" Then
            gsEMAIL_ONBENEFIT = rsPrefer("HP_ENABLED")
        End If
        If rsPrefer("HP_FUN_NAME") = "EMAIL_ONTERM" Then
            gsEMAIL_ONTERM = rsPrefer("HP_ENABLED")
        End If
        If rsPrefer("HP_FUN_NAME") = "EMAIL_ONREHIRE" Then
            gsEMAIL_ONREHIRE = rsPrefer("HP_ENABLED")
        End If
        If rsPrefer("HP_FUN_NAME") = "EMAIL_ONLEAVECHANGES" Then
            gsEMAIL_ONLEAVECHANGES = rsPrefer("HP_ENABLED")
        End If
        If rsPrefer("HP_FUN_NAME") = "EMAIL_ONPERFORMANCE" Then
            gsEMAIL_ONPERFORMANCE = rsPrefer("HP_ENABLED")
        End If
        If rsPrefer("HP_FUN_NAME") = "EMAIL_ONDEPENDENT" Then
            gsEMAIL_ONDEPENDENT = rsPrefer("HP_ENABLED")
        End If
        If rsPrefer("HP_FUN_NAME") = "DEPENDENT30DAYSEMAIL" Then 'Ticket #22061 Franks 05/24/2012
            gsEMAIL_ONDEPEND30DAYS4_WFC = rsPrefer("HP_ENABLED")
        End If
        If rsPrefer("HP_FUN_NAME") = "TRAININGMATRIX" Then
            gsTRAININGMATRIX = rsPrefer("HP_ENABLED")
        End If
        If rsPrefer("HP_FUN_NAME") = "GP_HOLDING" Then
            gsGPHold = rsPrefer("HP_ENABLED")
        End If
        If rsPrefer("HP_FUN_NAME") = "SECURED_PSW" Then
            gsSECURED_PSW = rsPrefer("HP_ENABLED")
        End If
        
        'Friesens - Ticket #17029
        If rsPrefer("HP_FUN_NAME") = "FRIESENSWORDPATH" Then
            gsFRIESENSWORDPATH = rsPrefer("HP_ENABLED")
        End If
        
        '8.0 = Ticket #22682 - Move Photo out of database into a folder
        If rsPrefer("HP_FUN_NAME") = "EMPLOYEEPHOTOPATH" Then
            gsEMPLOYEEPHOTO = rsPrefer("HP_ENABLED")
        End If
                
        'Ticket #24485 - Work Schedule Rotation Weeks
        If rsPrefer("HP_FUN_NAME") = "WS_ROTATIONWEEKS" Then
            gsWS_ROTATIONWEEKS = rsPrefer("HP_NUM")
            gsWS_ROTATIONWEEKSEFFDATE = CVDate(rsPrefer("HP_DATE"))
        End If
        
        '8.0 - Ticket #24352 - Database Connection Encryption
        If rsPrefer("HP_FUN_NAME") = "DB_CONNNECT_ENCRYPT" Then
            gsDB_CONNECT_ENCRYPT = rsPrefer("HP_ENABLED")
        End If
        
        '8.1 - Ticket #26529 -  Email Enhancement
        If rsPrefer("HP_FUN_NAME") = "SMTP_INFORMATION" Then
            gsSMTPINFO = rsPrefer("HP_ENABLED")
        End If
        
        'Ticket #26576 - WDGPHU - Flex Logic
        If rsPrefer("HP_FUN_NAME") = "FLEX_LOGIC" Then
            gsFLEX_LOGIC = rsPrefer("HP_ENABLED")
        End If
        
        'Ticket #30305 - Disable Compensatory Time Entries
        If rsPrefer("HP_FUN_NAME") = "DISABLE_COMPTIME" Then
            gsDISABLE_COMPTIME = rsPrefer("HP_ENABLED")
        End If

        'Ticket #26934 - Oshawa Community Health Centre - Employee Flags
        If rsPrefer("HP_FUN_NAME") = "EMAIL_ONEMPLOYEEFLAGS" Then
            gsEMAIL_ONEMPLOYEEFLAGS = rsPrefer("HP_ENABLED")
        End If
        
        'Ticket #28664 Franks 05/31/2016
        If rsPrefer("HP_FUN_NAME") = "EMAIL_ONHSINCIDENT" Then
            gsEMAIL_ONHSINCIDENT = rsPrefer("HP_ENABLED")
        End If
        
        rsPrefer.MoveNext
    Loop
    rsPrefer.Close
    
exH:
    Exit Sub
DOC_HNDLR:
    MsgBox "Attachments Database not found, please check with your IT department, Thank you"
    gsAttachment_DB = False
    Resume Next
End Sub

Function setLabels()
Dim SQLQ As String, rsLB As New ADODB.Recordset
Dim xStr As String
Dim x, Y
setLabels = False    ' returns true if found records
On Error GoTo LB_Err

'Delete the records for Flag 1, Flag 2, Flag 3, Flag 4 and Flag 5 as they have
'been replaced with UFlag 1, UFlag 2, UFlag 3, UFlag 4 and UFlag 5
SQLQ = "DELETE FROM HRLABEL WHERE LB_ORG IN ('Flag 1','Flag 2','Flag 3','Flag 4','Flag 5')"
gdbAdoIhr001.Execute SQLQ

For x = 1 To glbLabels(1).count: glbLabels(1).Remove 1: Next

glbLabels(1).Add "Department"
glbLabels(1).Add "G/L"
glbLabels(1).Add "Division"
glbLabels(1).Add "Location"
glbLabels(1).Add "Administered By"
glbLabels(1).Add "Region"
glbLabels(1).Add "Section"
glbLabels(1).Add "Union"
glbLabels(1).Add "Category"
glbLabels(1).Add "Original Hire"
glbLabels(1).Add "Seniority"
glbLabels(1).Add "Last Hire"
glbLabels(1).Add "Union Date"
glbLabels(1).Add "First Day"
glbLabels(1).Add "Last Day"
glbLabels(1).Add "OMERS Date"
glbLabels(1).Add "User Defined"
glbLabels(1).Add "Eligibility"
glbLabels(1).Add "Earliest Retirement"
glbLabels(1).Add "Normal Retirement"
glbLabels(1).Add "Latest Retirement"

'glbLabels(1).Add "Category"

glbLabels(1).Add "Hire Code"
glbLabels(1).Add "Pay Period"
glbLabels(1).Add "Driver License #"
glbLabels(1).Add "Type of Vehicle"
glbLabels(1).Add "Parking Permit #1"
glbLabels(1).Add "Parking Permit #2"
glbLabels(1).Add "License Plate #1"
glbLabels(1).Add "License Plate #2"
glbLabels(1).Add "Locker #"
glbLabels(1).Add "Combination"
glbLabels(1).Add "Position Group"
glbLabels(1).Add "Position Status"
glbLabels(1).Add "Grid Category"
glbLabels(1).Add "Salary Distribution"
glbLabels(1).Add "Supervisor Code"
glbLabels(1).Add "Vadim Field 1"
glbLabels(1).Add "Vadim Field 2"
'commented by Bryan 27/Mar/06, moved these labels into index 1 of glbLabels where they will actually be used.
'For Y = 1 To glbLabels(0).count: glbLabels(0).Remove 1: Next
glbLabels(1).Add "Employee Flag 1"
glbLabels(1).Add "Employee Flag 2"
glbLabels(1).Add "Employee Flag 3"
glbLabels(1).Add "Employee Flag 4"
glbLabels(1).Add "Employee Flag 5"
glbLabels(1).Add "Employee Flag 6"
glbLabels(1).Add "Employee Flag 7"
glbLabels(1).Add "Employee Flag 8"
glbLabels(1).Add "Employee Flag 9"
glbLabels(1).Add "Employee Flag 10"
glbLabels(1).Add "Employee Flag 11"
glbLabels(1).Add "Employee Flag 12"
glbLabels(1).Add "Employee Flag 13"
glbLabels(1).Add "Employee Flag 14"
glbLabels(1).Add "Employee Flag 15"
glbLabels(1).Add "Employee Flag 16"
glbLabels(1).Add "Employee Flag 17"
glbLabels(1).Add "Employee Flag 18"
glbLabels(1).Add "Employee Flag 19"
glbLabels(1).Add "Employee Flag 20"

'Employee Other Information
glbLabels(1).Add "Passport Expiration Date"
glbLabels(1).Add "Visa/Work Permit Expiration Date"
'INFOHR Menus
glbLabels(1).Add "Counseling"
glbLabels(1).Add "Comments"

'Employee Other Information
glbLabels(1).Add "Citizenship"
glbLabels(1).Add "Passport Country"
glbLabels(1).Add "Passport Number"
glbLabels(1).Add "Visa/Work Permit #"

glbLabels(1).Add "Other Expenses $"
glbLabels(1).Add "Employer $" '-68
glbLabels(1).Add "Accommodation $" ' 69

glbLabels(1).Add "Employee $" '70

glbLabels(1).Add "Performance" '71
glbLabels(1).Add "Associations" '72
glbLabels(1).Add "Charge Code" '73
glbLabels(1).Add "Machine #" '74
glbLabels(1).Add "Account Code" '75
glbLabels(1).Add "Claim #" '76

'Continuing Education
glbLabels(1).Add "Course Code" '77
glbLabels(1).Add "Scheduled Date" '78
glbLabels(1).Add "Start Date" '79
glbLabels(1).Add "Course Name" '80
glbLabels(1).Add "Course Description" '81
glbLabels(1).Add "Date Completed" '82
glbLabels(1).Add "Renewal Date" '83
glbLabels(1).Add "Conducted By" '84
glbLabels(1).Add "Company Name" '85
glbLabels(1).Add "Co-Ordinated By" '86
glbLabels(1).Add "Method Used" '87
glbLabels(1).Add "Trainer Name" '88
glbLabels(1).Add "Results" '89
glbLabels(1).Add "Account #" '90
glbLabels(1).Add "Keyword" '91
glbLabels(1).Add "Course Type" '92
glbLabels(1).Add "Course Hours" '93
glbLabels(1).Add "Presenter" '94
glbLabels(1).Add "Learning Material $" '95
glbLabels(1).Add "User Text 1"   '96
glbLabels(1).Add "User Text 2"   '97
glbLabels(1).Add "User Number 1" '98
glbLabels(1).Add "User Number 2" '99
glbLabels(1).Add "User Date"     '100
glbLabels(1).Add "Dependent Status" '101
glbLabels(1).Add "COB Dental" '102
glbLabels(1).Add "Dependent Smoker" '103
glbLabels(1).Add "COB Medical" '104
glbLabels(1).Add "Benefit Eligible Date" '105
glbLabels(1).Add "COB Other" '106
glbLabels(1).Add "Benefit End Date" '107
glbLabels(1).Add "Dependent Comment" '108
glbLabels(1).Add "Dependent Number" '109

'User Defined
glbLabels(1).Add "User Defined Table" '110
glbLabels(1).Add "Code 1"
glbLabels(1).Add "Code 2"
glbLabels(1).Add "Code 3"
glbLabels(1).Add "Code 4"
glbLabels(1).Add "Code 5"
glbLabels(1).Add "Date 1"
glbLabels(1).Add "Date 2"
glbLabels(1).Add "Date 3"
glbLabels(1).Add "Date 4"
glbLabels(1).Add "Date 5"
glbLabels(1).Add "UFlag 1"
glbLabels(1).Add "UFlag 2"
glbLabels(1).Add "UFlag 3"
glbLabels(1).Add "UFlag 4"
glbLabels(1).Add "UFlag 5"
glbLabels(1).Add "UText 1"
glbLabels(1).Add "UText 2"
glbLabels(1).Add "UComments" '128

glbLabels(1).Add "Follow-ups" '129
glbLabels(1).Add "Follow-Up" '130
 
'Pension Date 1 - 6
glbLabels(1).Add "Pension Date 1" '131
glbLabels(1).Add "Pension Date 2" '132
glbLabels(1).Add "Pension Date 3" '133
glbLabels(1).Add "Pension Date 4" '134
glbLabels(1).Add "Pension Date 5" '135
glbLabels(1).Add "Pension Date 6" '136

'Other Date 1 - 10
glbLabels(1).Add "Other Date 1" '137
glbLabels(1).Add "Other Date 2" '138
glbLabels(1).Add "Other Date 3" '139
glbLabels(1).Add "Other Date 4" '140
glbLabels(1).Add "Other Date 5" '141
glbLabels(1).Add "Other Date 6" '142
glbLabels(1).Add "Other Date 7" '143
glbLabels(1).Add "Other Date 8" '144
glbLabels(1).Add "Other Date 9" '145
glbLabels(1).Add "Other Date 10" '146

'Missed fields on Status/Date screen 'Ticket #15576 but DO NOT ADD Employment Status as Jerry asked
glbLabels(1).Add "Employment Type" '147
glbLabels(1).Add "Benefit Group" '148
glbLabels(1).Add "Internal Phone Extension" '149
glbLabels(1).Add "Email Address" '150

glbLabels(1).Add "From Date" '151
glbLabels(1).Add "To Date" '152
glbLabels(1).Add "Reason" '153
glbLabels(1).Add "AttSupervisor" '154
glbLabels(1).Add "Hours" '155
'glbLabels(1).Add "Charge Code" '73 'already appearing at the top
glbLabels(1).Add "Shift" '156
'glbLabels(1).Add "Claim #" '76 'already appearing at the top
glbLabels(1).Add "Point" '157
'glbLabels(1).Add "Account Code" '75    'already appearing at the top
'glbLabels(1).Add "Machine #" '74   'already appearing at the top
glbLabels(1).Add "Acting Position" '158
glbLabels(1).Add "PShift" '159
glbLabels(1).Add "Notes 1" '160
glbLabels(1).Add "Notes 2" '161
glbLabels(1).Add "SComments" '162
glbLabels(1).Add "Bonus $" '163

'Ticket #20609 Franks 09/06/2011 - Other Text fields on the Employee Other Information
glbLabels(1).Add "Other Text 1" '164
glbLabels(1).Add "Other Text 2" '165
glbLabels(1).Add "Other Text 3" '166
glbLabels(1).Add "Other Text 4" '167
'Ticket #20609 Franks 09/06/2011 - Other Text fields on the Employee Dependent screen
glbLabels(1).Add "Dependent Text 1" '168
glbLabels(1).Add "Dependent Text 2" '169
glbLabels(1).Add "Dependent Text 3" '170
glbLabels(1).Add "Dependent Text 4" '171
'Ticket #21462 Franks 02/09/2012 - Employee Position
glbLabels(1).Add "Rept. Authority 1" '172
glbLabels(1).Add "Rept. Authority 2" '173
glbLabels(1).Add "Rept. Authority 3" '174
glbLabels(1).Add "Rept. Authority 4" '175

'Ticket #22682 Hemu - Release 8.0 - Continuing Education
glbLabels(1).Add "CEU Type" '176
glbLabels(1).Add "CEU Credit" '177

'Ticket #23537 and Release 8.0 - Employee Position
glbLabels(1).Add "Hours/Day" '178
glbLabels(1).Add "Hours/Week" '179
glbLabels(1).Add "Hours/Pay Period" '180
glbLabels(1).Add "FTE Hours/Year" '181

'Ticket #23537 and Release 8.0 - Demographics
glbLabels(1).Add "Telephone #2" '182
glbLabels(1).Add "Cellular Telephone" '183
glbLabels(1).Add "Pager Number" '184

'Ticket #24164 - Re-ordering and new Organization fields
glbLabels(1).Add "Organization 1" '185
glbLabels(1).Add "Organization 2" '186

'Release 8.0 - Ticket #22682: Add Province to Label Master
glbLabels(1).Add "Prov. Name"   '187
glbLabels(1).Add "Prov. #"      '188
glbLabels(1).Add "Prov. Num1"   '189
glbLabels(1).Add "Prov. Num2"   '190
glbLabels(1).Add "Prov. Text1"  '191
glbLabels(1).Add "Prov. Text2"  '192
glbLabels(1).Add "Prov. Text3"  '193

'Ticket #25015 - Macaulay: New Additional Payroll ID Data
glbLabels(1).Add "Additional Payroll ID Data"   '194
glbLabels(1).Add "ADP Branch #"  '195
glbLabels(1).Add "ADP GL #"  '196
glbLabels(1).Add "ADP Department"  '197

'Release 8.0 - Ticket #2268: Add Payroll ID to Label Master
glbLabels(1).Add "Payroll ID"  '198

'Ticket #25911 Franks 10/20/2014 WFC - begin
glbLabels(1).Add "Position Description"  '199
glbLabels(1).Add "Position Alternate"  '200
glbLabels(1).Add "Position User Defined 1"  '201
glbLabels(1).Add "Position User Defined 2"  '202
glbLabels(1).Add "Job Group"  '203
glbLabels(1).Add "Position Level"  '204
glbLabels(1).Add "Job Status"  '205
glbLabels(1).Add "Job User Defined 1"  '206
glbLabels(1).Add "Job User Defined 2"  '207
'Ticket #25911 Franks 10/20/2014 WFC - end
'Note: don't change the order of glbLabels(1),
'it should keep the same order of txtNew(i)
glbLabels(1).Add "Other Email Address" '208

For x = 1 To glbLabels(2).count: glbLabels(2).Remove 1: Next

If glbLabLang = "" Then glbLabLang = "EN"
SQLQ = "SELECT * FROM HRLABEL WHERE LB_LANG = '" & glbLabLang & "'"

rsLB.Open SQLQ, gdbAdoIhr001, adOpenStatic
Do Until rsLB.EOF
    glbLabels(2).Add Trim(rsLB("LB_NEW")), Trim(rsLB("LB_ORG"))
    rsLB.MoveNext
Loop


Exit Function

LB_Err:
If Err.Number = 5 Then Resume Next
Exit Function
glbFrmCaption$ = "Set Labels"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Labels Snap", "HRLABEL", "SELECT")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Sub SetPanHelp(Contl As Control)
'*******************************************************
'*                                                     *
'*      Procedure Name: SetPanHelp                     *
'*                                                     *
'*             Created:              By:               *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: to set mdi pannel help/assist  *
'*                                                     *
'*******************************************************


Dim MsgPart1 As String, MsgPart2 As String, MsgPart3 As String
Dim lcontl As Integer, MsgPArt4 As String, strContl As String
Dim FChar As String, sChar As String
If Contl Is Nothing Then Exit Sub
lcontl = Len(Contl.Tag)
If lcontl > 1 Then
    If lcontl >= 3 Then
        strContl = Contl.Tag
        If Mid$(strContl, 5, 1) = "-" Then
            lcontl = lcontl - 5
            MsgPart2 = "AlphaNumeric"
            MsgPart1 = Right$(Contl.Tag, lcontl)
        ElseIf Mid$(strContl, 3, 1) = "-" Then
            FChar = Left$(strContl, 1)
            sChar = Mid$(strContl, 2, 1)
            Select Case FChar   'first char indicates type input/for assist
                Case "0"
                    MsgPart2 = "AlphaNumeric"
                Case "1"
                    MsgPart2 = "Numeric"
                Case "2"
                    MsgPart2 = "Currency"
                Case "3"
                    MsgPart2 = "Alpha"
                Case "4"
                    MsgPart2 = glbsDateFormat
                Case "5"
                    MsgPart2 = "X = True"
                Case "6"             'laura dec 16, 1997
                    MsgPart2 = "Year"
                Case Else
                    MsgPart2 = " "
            End Select
            Select Case sChar   '2nd char indicates required
                Case "0"  ' no assist
                    MsgPArt4 = " "
                Case "1"
                    MsgPArt4 = "Req."
                Case "2"
                    MsgPArt4 = "Con."
                Case Else
                    MsgPArt4 = " "
            End Select
            lcontl = lcontl - 3
            MsgPart1 = Right$(Contl.Tag, lcontl)
            If Mid$(MsgPart1, 5, 1) = "-" Then
                lcontl = lcontl - 5
                MsgPart2 = "AlphaNumeric"
                MsgPart1 = Right$(Contl.Tag, lcontl)
            End If
        Else
            MsgPart1 = Contl.Tag
        End If
    Else
        MsgPart1 = Contl.Tag
    End If
Else
    MsgPart1 = " "
End If


If TypeOf Contl Is CommandButton Then
    MsgPart2 = "Button"
End If
If TypeOf Contl Is ListBox Then
    MsgPart2 = "List box"
End If
If TypeOf Contl Is TextBox Then
 '   MsgPart2 = "Input Field"
     MsgPart3 = CStr(Contl.MaxLength)
     If MsgPart3 = "0" Then MsgPart3 = ""  'Jaddy Sep 21,1999
End If
If TypeOf Contl Is ComboBox Then
    MsgPart2 = "Combo Box"
End If
If TypeOf Contl Is MaskEdBox Then
    'MsgPart2 = CStr(Contl.Mask)
    MsgPart3 = CStr(Contl.MaxLength)
    If Contl.MaxLength = 64 Then MsgPart3 = ""
End If
If TypeOf Contl Is OptionButton Then
    MsgPart2 = "Option Button"
End If
If TypeOf Contl Is SSCheck Then
    MsgPart2 = "Check Box"
    MsgPart3 = 1
End If
If TypeOf Contl Is CheckBox Then 'Frank 5/24/2000
    MsgPart2 = "Check Box"
    MsgPart3 = 1
End If
If TypeOf Contl Is TDBGrid Then
    MsgPart2 = "Look-up"
End If
If TypeOf Contl Is CodeLookup Then ' Sam add 06/25/2002
    MsgPart3 = CStr(Contl.MaxLength)
End If
If Contl.Parent.name = "frmSLabel" Then
    MDIMain.panHelp(0).Caption = MsgPart1
Else
    MDIMain.panHelp(0).Caption = lStr(MsgPart1)
End If
If Contl.Parent.name = "frmSEmpFlags" Then
    MDIMain.panHelp(0).Caption = MsgPart1
Else
    MDIMain.panHelp(0).Caption = lStr(MsgPart1)
End If
MDIMain.panHelp(1).Caption = MsgPart2
MDIMain.panHelp(2).Caption = MsgPart3
MDIMain.panHelp(3).Caption = MsgPArt4
End Sub

Function SIN_chk(ssin$)
SIN_chk = False
Dim csumset1 As String

Dim c1$, c2$, c3$, c4$, c5$
Dim n1%, n2%, n3%, n4%, n5%

Dim sin1$, sin2$, sin3$, sin4$, sin5$, sin6$, sin7$, sin8$, sin9$
Dim ssin1%, ssin3%, ssin5%, ssin7%, ssin9%
Dim sset1$, sset2$

Dim charset1$, Sssin$

Dim numset1%, numset2%, sumset1%, nrest%
Dim ntotal%, snext%, chkdigit%, n1to5%

'Sssin = Left$(ssin, 3) & Mid$(ssin, 4, 3) & Right$(ssin, 3)
'ssin = CStr(Sssin)

sin1 = Mid$(ssin, 1, 1)
sin2 = Mid$(ssin, 2, 1)
sin3 = Mid$(ssin, 3, 1)
sin4 = Mid$(ssin, 4, 1)
sin5 = Mid$(ssin, 5, 1)
sin6 = Mid$(ssin, 6, 1)
sin7 = Mid$(ssin, 7, 1)
sin8 = Mid$(ssin, 8, 1)
sin9 = Mid$(ssin, 9, 1)
ssin1 = Val(sin1)
ssin3 = Val(sin3)
ssin5 = Val(sin5)
ssin7 = Val(sin7)
ssin9 = Val(sin9)
sset1 = sin2 + sin4 + sin6 + sin8
sset2 = sin2 + sin4 + sin6 + sin8
numset1 = Val(sset1)
numset2 = Val(sset2)
sumset1 = numset1 + numset2
csumset1 = CStr(sumset1)
'charset1 = Str(sumset1, 5, 0)
charset1 = Left$(csumset1, 5)

c1 = Mid$(charset1, 1, 1)
c2 = Mid$(charset1, 2, 1)
c3 = Mid$(charset1, 3, 1)
c4 = Mid$(charset1, 4, 1)
c5 = Mid$(charset1, 5, 1)
n1 = Val(c1)
n2 = Val(c2)
n3 = Val(c3)
n4 = Val(c4)
n5 = Val(c5)
n1to5 = n1 + n2 + n3 + n4 + n5
nrest = ssin1 + ssin3 + ssin5 + ssin7
ntotal = nrest + n1to5
If ntotal > 0 And ntotal < 10 Then
 snext = 10
 chkdigit = snext - ntotal
End If
If ntotal = 10 Then
 chkdigit = 0
End If
If ntotal > 10 And ntotal < 20 Then
 snext = 20
 chkdigit = snext - ntotal
End If
If ntotal = 20 Then
 chkdigit = 0
End If
If ntotal > 20 And ntotal < 30 Then
 snext = 30
 chkdigit = snext - ntotal
End If
If ntotal = 30 Then
 chkdigit = 0
End If
If ntotal > 30 And ntotal < 40 Then
 snext = 40
 chkdigit = snext - ntotal
End If
If ntotal = 40 Then
 chkdigit = 0
End If
If ntotal > 40 And ntotal < 50 Then
 snext = 50
 chkdigit = snext - ntotal
End If
If ntotal = 50 Then
 chkdigit = 0
End If
If ntotal > 50 And ntotal < 60 Then
 snext = 60
 chkdigit = snext - ntotal
End If
If ntotal = 60 Then
 chkdigit = 0
End If
If ntotal > 60 And ntotal < 70 Then
 snext = 70
 chkdigit = snext - ntotal
End If

If chkdigit = ssin9 Then
    SIN_chk = True
End If



End Function

Function SIN_chk_USA(ssin$)

If Not Len(ssin$) = 9 Then
  SIN_chk_USA = False
Else
  SIN_chk_USA = True
End If

End Function

Sub UnloadFrms(Optional xNewRecord)
     'Edited by Bryan on 29/Mar/2006
     'this way is more effcient, compares open forms to glbOnTop. if it's not on top it's unloaded
     Dim frm As Form
     For Each frm In Forms
        'If frm.name <> "MDIMain" And glbOnTop <> "FRMJOBS" Then 'Ticket #11941
        If frm.name <> "MDIMain" And glbOnTop <> "FRMJOBS" And glbOnTop <> "FRMJOBSWFC" Then 'Ticket #11941
            If frm.Visible = True And frm.MDIChild = True Then
                    If UCase(frm.name) <> UCase(glbOnTop) Then
                        Unload frm
                    End If
            End If
        End If
    Next frm
     
     
'    If IsMissing(xNewRecord) Then xNewRecord = ""
'    If xNewRecord = "newform" And glbOnTop = "FRMEEBASIC" Then
'        'don't unload the form
'    Else
'        Unload frmEEBASIC
'        'If glbOnTop = "FRMEEBASIC" Then Load frmEEBasic
'    End If
'    If xNewRecord = "newform" And glbOnTop = "FRMEMPLOYEEFLAGS" Then
'        'don't unload the form
'    Else
'        Unload frmEmployeeFlags
'        'If glbOnTop = "FRMEEBASIC" Then Load frmEEBasic
'    End If
'    If xNewRecord = "newform" And glbOnTop = "FRMEESTATS" Then
'        'don't unload the form
'    Else
'        Unload frmEESTATS
'       ' If glbOnTop = "FRMEESTATS" Then Load frmEEStats
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMEMERG" Then
'        'don't unload the form
'    Else
'        Unload frmEMERG
'       ' If glbOnTop = "FRMEMERG" Then Load frmEMERG
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMDEPNDTS" Then
'        'don't unload the form
'    Else
'         Unload frmDEPNDTS
'        ' If glbOnTop = "FRMDEPNDTS" Then Load frmDEPNDTS
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMEBANK" Then
'        'don't unload the form
'    Else
'         Unload frmEBANK
'        'If glbOnTop = "FRMEBANK" Then Load frmEBANK
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMEMPOTHER" Then
'        'don't unload the form
'    Else
'         Unload frmEmpOther
'        'If glbOnTop = "FRMEMPOTHER" Then Load frmEBANK
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMESKILLS" Then
'        'don't unload the form
'    Else
'         Unload frmESkills
'        'If glbOnTop = "FRMESKILLS" Then Load frmESKILLS
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMFORMALED" Then
'        'don't unload the form
'    Else
'         Unload frmFORMALED
'        ' If glbOnTop = "FRMFORMALED" Then Load frmFORMALED
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMESEMINARS" Then
'        'don't unload the form
'    Else
'         Unload frmESEMINARS
'        ' If glbOnTop = "FRMESEMINARS" Then Load frmESEMINARS
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMEASSOC" Then
'        'don't unload the form
'    Else
'         Unload frmEASSOC
'        ' If glbOnTop = "FRMEASSOC" Then Load frmEASSOC
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMEBENEFITS" Then
'        'don't unload the form
'    Else
'         Unload frmEBENEFITS
'        ' If glbOnTop = "FRMEBENEFITS" Then Load frmEBENEFITS
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMEODOLLAR" Then
'        'don't unload the form
'    Else
'         Unload frmEODOLLAR
'        ' If glbOnTop = "FRMEODOLLAR" Then Load frmEODOLLAR
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMOTHERERN" Then
'        'don't unload the form
'    Else
'         Unload frmOTHERERN
'        ' If glbOnTop = "FRMOTHERERN" Then Load frmOTHERERN
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMEPOSITION" Then
'        'don't unload the form
'    Else
'         Unload frmEPOSITION
'        ' If glbOnTop = "FRMEPOSITION" Then Load frmEPOSITION
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMEPERFORM" Then
'        'don't unload the form
'    Else
'         Unload frmEPERFORM
'        ' If glbOnTop = "FRMEPERFORM" Then Load frmEPERFORM
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMESALARY" Then
'        'don't unload the form
'    Else
'        Unload frmESALARY
'       ' If glbOnTop = "FRMESALARY" Then Load frmESALARY
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMVATTEND" Then
'        'don't unload the form
'    Else
'         Unload frmVATTEND
'        ' If glbOnTop = "FRMVATTEND" Then Load frmVATTEND
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMVACSICK" Then
'        'don't unload the form
'    Else
'         Unload frmVACSICK
'        ' If glbOnTop = "FRMVACSICK" Then Load frmVACSICK
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMVACSICKO" Then
'        'don't unload the form
'    Else
'         Unload frmVACSICKO
'        ' If glbOnTop = "FRMVACSICKO" Then Load frmVACSICKO
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMEHSINCIDENT" Then
'        'don't unload the form
'    Else
'         Unload frmEHSINCIDENT
'        ' If glbOnTop = "FRMEHSINCIDENT" Then Load frmEHSINCIDENT
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMECOMMENTS" Then
'        'don't unload the form
'    Else
'         Unload frmECOMMENTS
'        ' If glbOnTop = "FRMECOMMENTS" Then Load frmECOMMENTS
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMEFOLLOWUP" Then
'        'don't unload the form
'    Else
'         Unload frmEFOLLOWUP
'        ' If glbOnTop = "FRMEFOLLOWUP" Then Load frmEFOLLOWUP
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMEHSWCB" Then
'        'don't unload the form
'    Else
'         Unload frmEHSWCB
'        ' If glbOnTop = "FRMEHSWCB" Then Load frmEHSWCB
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMEHSWCBC" Then
'        'don't unload the form
'    Else
'         Unload frmEHSWCBC
'        ' If glbOnTop = "FRMEHSWCBC" Then Load frmEHSWCBC
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMEHSINJURY" Then
'        'don't unload the form
'    Else
'         Unload frmEHSINJURY
'        ' If glbOnTop = "FRMEHSINJURY" Then Load frmEHSINJURY
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMHRENT" Then
'        'don't unload the form
'    Else
'         Unload frmHrEnt
'        ' If glbOnTop = "FRMHRENT" Then Load frmHRENT
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMETERM" Then
'        'don't unload the form
'    Else
'         Unload frmETERM
'        ' If glbOnTop = "FRMETERM" Then Load frmETERM
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMEHSCAUSE" Then
'        'don't unload the form
'    Else
'        Unload frmEHSCause
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMEHSCORRECTIVE" Then
'        'don't unload the form
'    Else
'        Unload frmEHSCorrective
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMEHSCONTACT" Then
'        'don't unload the form
'    Else
'        Unload frmEHSContact
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMCOBRA" Then
'        'don't unload the form
'    Else
'        Unload frmCobra
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMECOMPLAN" Then
'        'don't unload the form
'    Else
'        Unload frmEComPlan
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMECOUNSEL" Then
'        'don't unload the form
'    Else
'        Unload frmECounsel
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMTLAY" Then
'        'don't unload the form
'    Else
'        Unload frmTLAY
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMETLAY" Then
'        'don't unload the form
'    Else
'        Unload frmETLAY
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMMPOSITIONS" Then
'        'don't unload the form
'    Else
'        Unload frmMPOSITIONS
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMPOSSKILLS" Then
'        'don't unload the form
'    Else
'        Unload frmPosSkills
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMPOSEVAL" Then
'        'don't unload the form
'    Else
'        Unload frmPosEval
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMPOSRESP" Then
'        'don't unload the form
'    Else
'        Unload frmPosResp
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMPOSDUTIES" Then
'        'don't unload the form
'    Else
'        Unload frmPosDuties
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMPOSCOURSE" Then
'        'don't unload the form
'    Else
'        Unload frmPosCourse
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMPOSBUDGET" Then
'        'don't unload the form
'    Else
'        Unload frmPosBudget
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMPOSAPPPROC" Then
'        'don't unload the form
'    Else
'        Unload frmPosAppProc
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMUATTEND" Then
'        'don't unload the form
'    Else
'        Unload frmUATTEND
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMUATTHIS" Then
'        'don't unload the form
'    Else
'        Unload frmUATTHIS
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMUBENEFITS" Then
'        'don't unload the form
'    Else
'        Unload frmUBENEFITS
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMUCODE" Then
'        'don't unload the form
'    Else
'        Unload frmUCode
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMUDOORS" Then
'        'don't unload the form
'    Else
'        Unload frmUDOORS
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMUEMPNUM" Then
'        'don't unload the form
'    Else
'        Unload frmUEmpNum
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMUENTITLE" Then
'        'don't unload the form
'    Else
'        Unload frmUEntitle
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMUFOLLOW" Then
'        'don't unload the form
'    Else
'        Unload frmUFollow
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMUJOBS" Then
'        'don't unload the form
'    Else
'        Unload frmUJobs
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMUOTHEREARN" Then
'        'don't unload the form
'    Else
'        Unload frmUOtherEarn
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMUREPAUTH" Then
'        'don't unload the form
'    Else
'        Unload frmURepAuth
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMUSALARY" Then
'        'don't unload the form
'    Else
'        Unload frmUSalary
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMUSEMINARS" Then
'        'don't unload the form
'    Else
'        Unload frmUSEMINARS
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMUTD1DOLLAR" Then
'        'don't unload the form
'    Else
'        Unload frmUTd1Dollar
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMUTERM" Then
'        'don't unload the form
'    Else
'        Unload frmUTERM
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMSVACENT" Then
'        'don't unload the form
'    Else
'        Unload frmSVacEnt
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMSICKENT" Then
'        'don't unload the form
'    Else
'        Unload frmSickEnt
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMSHRSENT" Then
'        'don't unload the form
'    Else
'        Unload frmSHrsEnt
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMSECURE" Then
'        'don't unload the form
'    Else
'        Unload frmSECURE
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMSPENENT" Then
'        'don't unload the form
'    Else
'        Unload frmSPenEnt
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMDOLENTIT" Then
'        'don't unload the form
'    Else
'        Unload frmDolEntit
'    End If
'    If xNewRecord = "newform" And glbOnTop = "FRMSHOLIDAY" Then
'        'don't unload the form
'    Else
'         Unload frmSHoliday
'        ' If glbOnTop = "FRMECOMMENTS" Then Load frmECOMMENTS
'    End If
'    If xNewRecord = "newform" And glbOnTop = "FRMEMPLOYEEFLAGS" Then
'        'don't unload the form
'    Else
'         Unload frmEmployeeFlags
'        ' If glbOnTop = "FRMECOMMENTS" Then Load frmECOMMENTS
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMUACCRCLR" Then
'        'don't unload the form
'    Else
'        Unload frmUAccrClr
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "frmEGLDist" Then
'
'    Else
'        Unload frmEGLDist
'    End If
'
'    If xNewRecord = "newform" And glbOnTop = "FRMVFOLLOWUP" Then
'
'    Else
'        Unload frmvFOLOWUP
'    End If
'
'    If glbLinamar Or glbWFC Then
'        Unload frmETRANIN
'    End If
'    If glbAxxent Then
''    Unload frmAXXRSP
'    End If
End Sub

Sub UpdUStats(frmName As Form)

frmName.Updstats(0).Text = Format(Now, "SHORT DATE")
frmName.Updstats(1).Text = Time$
frmName.Updstats(2).Text = glbUserID

End Sub


Sub LoadForm(FormName As String)
FormName = UCase(FormName)
'Sam changed it for Friesens as they have only Performance Review form
If glbCompSerial = "S/N - 2279W" Then  'Friesens Corporation - Ticket #11784
    If FormName = "FRMEPERFORM" Then
        FormName = "FRMEPERFORMREVIEW"
    End If
End If

Select Case FormName
    Case "FRMEEBASIC": Load frmEEBASIC
    Case "FRMEESTATS": Load frmEESTATS
    Case "FRMEMERG": Load frmEMERG
    Case "FRMDEPNDTS": Load frmDEPNDTS
    Case "FRMEBANK": Load frmEBANK
    Case "FRMEMPOTHER": Load frmEmpOther
    Case "FRMEMPLOYEEFLAGS": Load frmEmployeeFlags
    Case "FRMESKILLS": Load frmESkills
    Case "FRMFORMALED": Load frmFORMALED
    Case "FRMESEMINARS": Load frmESEMINARS
    Case "FRMEASSOC": Load frmEASSOC
    Case "FRMEBENEFITS": Load frmEBENEFITS
    Case "FRMEODOLLAR": Load frmEODOLLAR
    Case "FRMOTHERERN": Load frmOTHERERN
    Case "FRMEPOSITION": Load frmEPOSITION
    
    Case "FRMEPERFORM": Load frmEPERFORM
    
    Case "FRMEPERFORMREVIEW": Load frmEPERFORMReview
    
    Case "FRMESALARY": Load frmESALARY
    Case "FRMESALARYMusashi": Load frmESALARYMusashi
    Case "FRMVATTEND": Load frmVATTEND
    Case "FRMVACSICK": Load frmVACSICK
    Case "FRMVACSICKO": Load frmVACSICKO
    Case "FRMHRENT": Load frmHrEnt
    Case "FRMEHSINCIDENT": Load frmEHSINCIDENT
    Case "FRMEHSINJURY": Load frmEHSINJURY
    Case "FRMEHSCAUSE": Load frmEHSCause
    Case "FRMEHSCORRECTIVE": Load frmEHSCorrective
    Case "FRMEHSWCB": Load frmEHSWCB
    Case "FRMEHSCONTACT": Load frmEHSContact
    Case "FRMEHSWCBC": Load frmEHSWCBC
    Case "FRMECOMMENTS": Load frmECOMMENTS
    Case "FRMEFOLLOWUP": Load frmEFOLLOWUP
    Case "FRMCOBRA": Load frmCobra
    Case "FRMECOMPLAN": If Not glbtermopen Then Load frmEComPlan
    Case "FRMECOUNSEL": Load frmECounsel
    Case UCase("frmEEOTHER"): Load frmEmpOther
    'v7.6
    Case UCase("frmEmployeeFlags"): Load frmEmployeeFlags
    Case UCase("frmEGLDist"): Load frmEGLDist
    Case UCase("frmELang"): Load frmELang
    Case UCase("frmESuccession"): Load frmESuccession
    Case UCase("frmEmpADP"): Load frmEmpADP
    
    ' danielk - 12/31/2002 - added EEO form
    Case "FRMEEO": Load frmEEO
    Case "FRMEUSERDEF": Load frmEUserDef 'Ticket #30482 Franks 08/16/2017
    
End Select
End Sub

Sub NextForm()
If NewHireForms.count > 0 Then
    If UCase(NewHireForms(1)) = UCase(MDIMain.ActiveForm.name) Then
        If glbWFC Then 'Ticket #24184 Franks 09/11/2013
            Call WFCHRSoftProcUpt(MDIMain.ActiveForm.name)
        End If
        NewHireForms.Remove 1
        If NewHireForms.count > 0 Then
'            fGLBNew = True
            Call LoadForm(NewHireForms(1))
        End If
    End If
End If
End Sub

Function NextFormIF(Msg1 As String)
Dim Msg$, VReturn%
NextFormIF = False
If NewHireForms.count > 0 Then
    Msg$ = "Do you wish to add another " & Msg1 & " record?"
    VReturn% = MsgBox(Msg$, MB_YESNO)
    If VReturn% = IDYES Then
        NextFormIF = True
    Else
        Call NextForm
    End If
End If
End Function

Public Function IsWFCHRSFUptSuccess(xFormName, xCandID)
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As Boolean
    retVal = False
    If xFormName = "frmEPOSITION" Then
        SQLQ = "SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE NOT JH_CURRENT = 0 "
        SQLQ = SQLQ & "AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_CANDIDATE =  " & xCandID & " ) "
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTemp.EOF Then
            retVal = True
        End If
        rsTemp.Close
    End If
    If xFormName = "frmESALARY" Then
        SQLQ = "SELECT SH_EMPNBR FROM HR_SALARY_HISTORY WHERE NOT SH_CURRENT = 0 "
        SQLQ = SQLQ & "AND SH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_CANDIDATE =  " & xCandID & " ) "
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTemp.EOF Then
            retVal = True
        End If
        rsTemp.Close
    End If
    IsWFCHRSFUptSuccess = retVal
End Function

Sub Pause(ByVal nSecond As Single)
   Dim t0 As Single
   t0 = Timer
   If nSecond = 0.5 Then nSecond = 0.8
   Do While Timer - t0 < nSecond
      Dim dummy As Integer
       dummy = DoEvents()
      ' if we cross midnight, back up one day
      If Timer < t0 Then
         t0 = t0 - CLng(24) * CLng(60) * CLng(60)
      End If
   Loop
End Sub
Sub setCaption(xCtrl As Object)
On Error GoTo setCaption_err
xCtrl.Caption = lStr(xCtrl.Caption)

setCaption_err:
End Sub

Function getCodes(codes As Variant)

    'Remove extra commas
    Do Until InStr(codes, ",,") = 0
        codes = Replace(codes, ",,", ",")
    Loop
    
    'If 1st char is comma, remove it
    If Left(codes, 1) = "," Then
        codes = Mid(codes, 2, Len(codes) - 1)
    End If
    
    'If lasts char is comma, remove it
    If Right(codes, 1) = "," Then
        codes = Mid(codes, 1, Len(codes) - 1)
    End If
    
    'Add the single quotes enclosing each code
    'In the selection criteria - don't forget to add quotes at the beginning
    'of the text and at the end of the text - like any string.
    codes = Replace(codes, ",", "','")
    getCodes = codes
    
End Function
Function getPayrollID(PayID As Variant)
Dim xPayID
    xPayID = "'" & PayID & "'"
    xPayID = Replace(xPayID, ",", "','")
    getPayrollID = xPayID
End Function
Function getEmpnbr(EmpID As Variant)
Dim xEMP

'Add by Frank on Jun 27,02 for causing problem if extra "," at the end of line
Do Until InStr(EmpID, ",,") = 0
    EmpID = Replace(EmpID, ",,", ",")
Loop
If Left(EmpID, 1) = "," Then
    EmpID = Mid(EmpID, 2, Len(EmpID) - 1)
End If
If Right(EmpID, 1) = "," Then
    EmpID = Mid(EmpID, 1, Len(EmpID) - 1)
End If
'Add by Frank on Jun 27,02

getEmpnbr = EmpID
If Len(EmpID) < 3 Then Exit Function
If glbLinamar Then
    
    If InStr(EmpID, ",") = 0 Then
        getEmpnbr = Val(Mid(EmpID, 5)) & Format(Val(Left(EmpID, 3)), "000")
    Else
        getEmpnbr = ""
        Do While InStr(EmpID, ",") <> 0
            xEMP = Left(EmpID, InStr(EmpID, ",") - 1)
            EmpID = Mid(EmpID, InStr(EmpID, ",") + 1)
            If Len(xEMP) > 3 Then
                getEmpnbr = getEmpnbr & Val(Mid(xEMP, 5)) & Format(Val(Left(xEMP, 3)), "000") & ","
            End If
        Loop
        If Len(xEMP) > 3 Then
            getEmpnbr = getEmpnbr & Val(Mid(EmpID, 5)) & Format(Val(Left(EmpID, 3)), "000")
        Else
            getEmpnbr = Left(getEmpnbr, Len(getEmpnbr) - 1)
        End If
    End If
End If
End Function
Function ShowEmpnbr(EmpID As Variant)

ShowEmpnbr = EmpID
If Len(EmpID) < 3 Then Exit Function

If glbLinamar Then
    ShowEmpnbr = Right(EmpID, 3) + "-" + Left(EmpID, Len(EmpID) - 3)
End If
End Function

Function GetNewEmpnbr()
    GetNewEmpnbr = 0
    glbLEE_ID = 0
    frmNewEmployee.Show 1
    GetNewEmpnbr = glbTrsEE_ID
End Function

Function lStr(xStr As String)
Dim xOld, xNew
Dim xOBJ, x
On Error GoTo LStr_Err
lStr = xStr

If glbLinamar Then
    xOld = "First Name"
    xNew = "First/Second Name"
    If InStr(xStr, xOld) <> 0 Then lStr = Replace(xStr, xOld, xNew)
End If
For x = 1 To glbLabels(1).count
    xStr = Replace(xStr, "General Ledger", "G/L")
    xStr = Replace(xStr, "G/L #", "G/L")
    xOld = glbLabels(1).Item(x)
    xNew = glbLabels(2)(xOld)
    If xOld = "Seniority" Then
        If InStr(xStr, xOld) <> 0 Then
            If InStr(xStr, "Seniority Report") = 0 And InStr(xStr, "Seniority Hour") = 0 Then
                lStr = Replace(xStr, xOld, xNew)
            Else
                lStr = Replace(xStr, xOld & " Date", xNew & " Date")
            End If
        End If
    Else
        If InStr(xStr, xOld) <> 0 Then
            lStr = Replace(xStr, xOld, xNew)
        Else
            If InStr(xStr, UCase(xOld)) <> 0 Then
                lStr = Replace(xStr, UCase(xOld), UCase(xNew))
            End If
        End If
    End If
Next x



Exit Function
LStr_Err:
    If Err.Number = 5 Then
        xNew = xOld
        Resume Next
    End If
End Function

Sub setRptCaption(frmName As Form)
On Error Resume Next
Call setCaption(frmName.lblDiv)
Call setCaption(frmName.lblDept)
Call setCaption(frmName.lblLocation)
Call setCaption(frmName.lblRegion)
Call setCaption(frmName.lblAdmin)
Call setCaption(frmName.lblSection)
Call setCaption(frmName.lblUnion)
Call setCaption(frmName.lblSen)
Call setCaption(frmName.lblPT)
Call setCaption(frmName.lblGrid)
Call setCaption(frmName.lblPosGroup)
Call setCaption(frmName.lblPosStatus)

End Sub

Sub setRptLabel(frmName As Form, xType)

If xType = 0 Then
    frmName.vbxCrystal.Formulas(10) = "lblDept='" & Replace(lStr("Department"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(11) = "lblDivision='" & Replace(lStr("Division"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(12) = "lblUnion='" & Replace(lStr("Union"), "'", "''") & "'"
    If frmName.name = "frmSVacEnt" Or frmName.name = "frmSalPerctg" Or frmName.name = "frmVacPerctg" Then
        frmName.vbxCrystal.Formulas(15) = "lblLocation='" & Replace(lStr("Location") & IIf(xType = 1, ":", ""), "'", "''") & "'"
        If glbLinamar Then
            frmName.vbxCrystal.Formulas(18) = "lblSection='Vacation Group'"
        Else
            frmName.vbxCrystal.Formulas(18) = "lblSection='" & Replace(lStr("Section") & IIf(xType = 1, ":", ""), "'", "''") & "'"
        End If
    End If
End If
If xType = 1 Or xType = 2 Then
    frmName.vbxCrystal.Formulas(10) = "lblPT='" & Replace(lStr("Category") & IIf(xType = 1, ":", ""), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(11) = "lblDept='" & Replace(lStr("Department") & IIf(xType = 1, ":", ""), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(12) = "lblDivision='" & Replace(lStr("Division") & IIf(xType = 1, ":", ""), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(13) = "lblUnion='" & Replace(lStr("Union") & IIf(xType = 1, ":", ""), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(14) = "lblGL='" & Replace(lStr("G/L#") & IIf(xType = 1, ":", ""), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(15) = "lblLocation='" & Replace(lStr("Location") & IIf(xType = 1, ":", ""), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(16) = "lblAdmin='" & Replace(lStr("Administered By") & IIf(xType = 1, ":", ""), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(17) = "lblRegion='" & Replace(lStr("Region") & IIf(xType = 1, ":", ""), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(18) = "lblSection='" & Replace(lStr("Section") & IIf(xType = 1, ":", ""), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(19) = "lblOHireDate='" & Replace(lStr("Original Hire Date") & IIf(xType = 1, ":", ""), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(20) = "lblSeniority='" & Replace(lStr("Seniority Date") & IIf(xType = 1, ":", ""), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(21) = "lblLHireDate='" & Replace(lStr("Last Hire Date") & IIf(xType = 1, ":", ""), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(22) = "lblUnionDate='" & Replace(lStr("Union Date") & IIf(xType = 1, ":", ""), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(23) = "lblFDay='" & Replace(lStr("First Day") & IIf(xType = 1, ":", ""), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(24) = "lblLDay='" & Replace(lStr("Last Day") & IIf(xType = 1, ":", ""), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(25) = "lblOMERS='" & Replace(lStr("OMERS Date") & IIf(xType = 1, ":", ""), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(26) = Replace("lblUserDate='" & Replace(lStr("User Defined Date") & IIf(xType = 1, ":", ""), "'", "''") & "'", "Date Date", "Date")
    frmName.vbxCrystal.Formulas(27) = "lblEligibility='" & Replace(lStr("Eligibility") & IIf(xType = 1, ":", ""), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(28) = "Age55Title='" & Replace(lStr("Earliest Retirement") & IIf(xType = 1, ":", ""), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(29) = "Age60Title='" & Replace(lStr("Normal Retirement") & IIf(xType = 1, ":", ""), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(30) = "Age65Title='" & Replace(lStr("Latest Retirement") & IIf(xType = 1, ":", ""), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(31) = "lblDeptStart='" & Replace(lStr("Department Start Date") & IIf(xType = 1, ":", ""), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(32) = "lblDivStart='" & Replace(lStr("Division Start Date") & IIf(xType = 1, ":", ""), "'", "''") & "'"
    If frmName.name = "frmRPosition" And frmName.Caption = "Employee Profile Report" Then
        frmName.vbxCrystal.Formulas(33) = "lblDriverLicense='" & Replace(lStr("Driver License #") & IIf(xType = 1, ":", ""), "'", "''") & "'"
        frmName.vbxCrystal.Formulas(34) = "lblTypeVehicle='" & Replace(lStr("Type of Vehicle") & IIf(xType = 1, ":", ""), "'", "''") & "'"
        frmName.vbxCrystal.Formulas(35) = "lblParkingPermit1='" & Replace(lStr("Parking Permit #1") & IIf(xType = 1, ":", ""), "'", "''") & "'"
        frmName.vbxCrystal.Formulas(36) = "lblParkingPermit2='" & Replace(lStr("Parking Permit #2") & IIf(xType = 1, ":", ""), "'", "''") & "'"
        frmName.vbxCrystal.Formulas(37) = "lblLicensePlate1='" & Replace(lStr("License Plate #1") & IIf(xType = 1, ":", ""), "'", "''") & "'"
        frmName.vbxCrystal.Formulas(38) = "lblLicensePlate2='" & Replace(lStr("License Plate #2") & IIf(xType = 1, ":", ""), "'", "''") & "'"
        frmName.vbxCrystal.Formulas(39) = "lblLocker='" & Replace(lStr("Locker #") & IIf(xType = 1, ":", ""), "'", "''") & "'"
        frmName.vbxCrystal.Formulas(40) = "lblCombination='" & Replace(lStr("Combination") & IIf(xType = 1, ":", ""), "'", "''") & "'"

        If glbVadim Then
            frmName.vbxCrystal.Formulas(41) = "lblWCBCode='EI rate" & IIf(xType = 1, ":", "") & "'"
        ElseIf glbPayWeb Then
            frmName.vbxCrystal.Formulas(41) = "lblWCBCode='E.I. Reduced Rate" & IIf(xType = 1, ":", "") & "'"
        ElseIf glbInsync Then
            frmName.vbxCrystal.Formulas(41) = "lblWCBCode='Status Federal Tax" & IIf(xType = 1, ":", "") & "'"
        End If
    ElseIf frmName.name = "frmRPosition" Then    'Position only
        frmName.vbxCrystal.Formulas(42) = "lblGridCategory='" & Replace(lStr("Grid Category") & IIf(xType = 1, ":", ""), "'", "''") & "'"
    End If
    If frmName.name = "frmEEBASIC" Then
        frmName.vbxCrystal.Formulas(33) = "lblDriverLicense='" & Replace(lStr("Driver License #") & IIf(xType = 1, ":", ""), "'", "''") & "'"
        frmName.vbxCrystal.Formulas(34) = "lblTypeVehicle='" & Replace(lStr("Type of Vehicle") & IIf(xType = 1, ":", ""), "'", "''") & "'"
        frmName.vbxCrystal.Formulas(35) = "lblParkingPermit1='" & Replace(lStr("Parking Permit #1") & IIf(xType = 1, ":", ""), "'", "''") & "'"
        frmName.vbxCrystal.Formulas(36) = "lblParkingPermit2='" & Replace(lStr("Parking Permit #2") & IIf(xType = 1, ":", ""), "'", "''") & "'"
        frmName.vbxCrystal.Formulas(37) = "lblLicensePlate1='" & Replace(lStr("License Plate #1") & IIf(xType = 1, ":", ""), "'", "''") & "'"
        frmName.vbxCrystal.Formulas(38) = "lblLicensePlate2='" & Replace(lStr("License Plate #2") & IIf(xType = 1, ":", ""), "'", "''") & "'"
        frmName.vbxCrystal.Formulas(39) = "lblLocker='" & Replace(lStr("Locker #") & IIf(xType = 1, ":", ""), "'", "''") & "'"
        frmName.vbxCrystal.Formulas(40) = "lblCombination='" & Replace(lStr("Combination") & IIf(xType = 1, ":", ""), "'", "''") & "'"
    End If
End If
If xType = 3 Then   'Ticket #15276 - Employee Flags Reports
    frmName.vbxCrystal.Formulas(51) = "lblEmpFlag1='" & Replace(lStr("Employee Flag 1"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(52) = "lblEmpFlag2='" & Replace(lStr("Employee Flag 2"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(53) = "lblEmpFlag3='" & Replace(lStr("Employee Flag 3"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(54) = "lblEmpFlag4='" & Replace(lStr("Employee Flag 4"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(55) = "lblEmpFlag5='" & Replace(lStr("Employee Flag 5"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(56) = "lblEmpFlag6='" & Replace(lStr("Employee Flag 6"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(57) = "lblEmpFlag7='" & Replace(lStr("Employee Flag 7"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(58) = "lblEmpFlag8='" & Replace(lStr("Employee Flag 8"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(59) = "lblEmpFlag9='" & Replace(lStr("Employee Flag 9"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(60) = "lblEmpFlag10='" & Replace(lStr("Employee Flag 10"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(61) = "lblEmpFlag11='" & Replace(lStr("Employee Flag 11"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(62) = "lblEmpFlag12='" & Replace(lStr("Employee Flag 12"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(63) = "lblEmpFlag13='" & Replace(lStr("Employee Flag 13"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(64) = "lblEmpFlag14='" & Replace(lStr("Employee Flag 14"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(65) = "lblEmpFlag15='" & Replace(lStr("Employee Flag 15"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(66) = "lblEmpFlag16='" & Replace(lStr("Employee Flag 16"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(67) = "lblEmpFlag17='" & Replace(lStr("Employee Flag 17"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(68) = "lblEmpFlag18='" & Replace(lStr("Employee Flag 18"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(69) = "lblEmpFlag19='" & Replace(lStr("Employee Flag 19"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(70) = "lblEmpFlag20='" & Replace(lStr("Employee Flag 20"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(71) = "lblFollowUp='" & Replace(lStr("Follow-Up"), "'", "''") & "'"
End If

If xType = 4 Then   'Ticket #18668 - info:HR 7.9 release
    frmName.vbxCrystal.Formulas(31) = "lblAttFromDate='" & Replace(lStr("From Date"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(32) = "lblAttToDate='" & Replace(lStr("To Date"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(33) = "lblAttReason='" & Replace(lStr("Reason"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(34) = "lblSupervisor='" & Replace(lStr("Supervisor"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(35) = "lblAttHours='" & Replace(lStr("Hours"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(36) = "lblChargeCode='" & Replace(lStr("Charge Code"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(37) = "lblShift='" & Replace(lStr("Shift"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(38) = "lblClaimNo='" & Replace(lStr("Claim #"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(39) = "lblPoint='" & Replace(lStr("Point"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(40) = "lblAccountCode='" & Replace(lStr("Account Code"), "'", "''") & "'"
    frmName.vbxCrystal.Formulas(41) = "lblMachineNo='" & Replace(lStr("Machine #"), "'", "''") & "'"
End If
End Sub

Sub Set_Div_List()
Dim SQLQ
Dim rsDiv As New ADODB.Recordset
SQLQ = "select * from HR_DIVISION WHERE " & glbSeleDiv
rsDiv.Open SQLQ, gdbAdoIhr001, adOpenStatic
glbDIVList = "("
glbDIVCount = 0
Do Until rsDiv.EOF
    glbDIVList = glbDIVList & "'" & rsDiv("DIV") & "',"
    glbDIVCount = glbDIVCount + 1
    glbSDIV = rsDiv("DIV")
    rsDiv.MoveNext
Loop
glbDIVList = Left(glbDIVList, Len(glbDIVList) - 1)
glbDIVList = glbDIVList & ", 'ALL')"
rsDiv.Close
End Sub

Sub iniDateFormat()
Dim xPos, xStr, x

glbDateSeparator = IIf(InStr(glbsDateFormat, "/") <> 0, "/", "-")

xStr = UCase(glbsDateFormat)
For x = 1 To 3
    xPos = InStr(xStr, glbDateSeparator)
    If xPos <> 0 Then
        glbDateFmt(x) = Left(xStr, xPos - 1)
        xStr = Mid(xStr, xPos + 1)
    Else
        glbDateFmt(x) = xStr
    End If
Next
End Sub

Sub Date_Change(txtAllDate As Object) '(txtAllDate As TextBox)
Dim xStr, xPos, xLin, KeyAscii, x
On Error GoTo Can_Err

If (txtAllDate Is Nothing) Then Exit Sub
If Len(txtAllDate) = 0 Then Exit Sub
If txtAllDate.CausesValidation Then Exit Sub
txtAllDate.CausesValidation = True

If Right(txtAllDate, 1) = glbDateSeparator Then
    If Right(txtAllDate, 2) = glbDateSeparator & glbDateSeparator Then
        SendKeys "{Backspace}"
    End If
    Exit Sub
End If

KeyAscii = Asc(Right(txtAllDate, 1))
xStr = txtAllDate
x = 0
Do While True
    xPos = InStr(xStr, glbDateSeparator)
    x = x + 1
    If xPos <> 0 Then
        xStr = Mid(xStr, xPos + 1)
    Else
        Exit Do
    End If
Loop
If x > 3 Then Exit Sub

Select Case glbDateFmt(x)
Case "MMM", "MM", "M"
    Select Case Len(xStr)
    Case 1
        If KeyAscii > 49 And KeyAscii < 57 Then
            txtAllDate = txtAllDate + glbDateSeparator
            SendKeys "{End}"
        End If
    Case 2
        If Val(xStr) = 0 Then
            SendKeys "{Backspace}"
            Exit Sub
        End If
        If Val(xStr) > 12 Then
            If x <> 3 Then
                txtAllDate = Left(txtAllDate, Len(txtAllDate) - 1) + glbDateSeparator + Right(txtAllDate, 1)
                SendKeys "{End}"
            Else
                SendKeys "{Backspace}"
                Exit Sub
            End If
        Else
            If x <> 3 Then
                txtAllDate = txtAllDate + glbDateSeparator
                SendKeys "{End}"
            End If
        End If
    Case 3
        SendKeys "{Backspace}"
    End Select
Case "DD", "D"
    Select Case Len(xStr)
    Case 1
        If (KeyAscii > 51 And KeyAscii < 57) Or (KeyAscii > 99 And KeyAscii < 105) Then
            If x <> 3 Then
                txtAllDate = txtAllDate + glbDateSeparator
                SendKeys "{End}"
            End If
        End If
    Case 2
        If Val(xStr) = 0 Then
            SendKeys "{Backspace}"
            Exit Sub
        End If
        If Val(xStr) > 31 Then
            If x <> 3 Then
                txtAllDate = Left(txtAllDate, Len(txtAllDate) - 1) + glbDateSeparator + Right(txtAllDate, 1)
                SendKeys "{End}"
            Else
                SendKeys "{Backspace}"
                Exit Sub
            End If
        Else
            If x <> 3 Then
                txtAllDate = txtAllDate + glbDateSeparator
                SendKeys "{End}"
            End If
        End If
    Case 3
        SendKeys "{Backspace}"
    End Select
Case "YYYY"
    Select Case Len(xStr)
    Case 2
'        If Val(xStr) = 0 Then
'            SendKeys "{Backspace}"
'            Exit Sub
'        End If
    Case 3
        If Left(xStr, 1) = "0" Then SendKeys "{Backspace}"
    Case 4
        If x <> 3 Then
            txtAllDate = txtAllDate + glbDateSeparator
            SendKeys "{End}"
        End If
    Case 5
        SendKeys "{Backspace}"
    End Select
Case "YY"
    Select Case Len(xStr)
    Case 2
        If x <> 3 Then
            txtAllDate = txtAllDate + glbDateSeparator
            SendKeys "{End}"
        End If
    Case 3
        If Left(xStr, 1) = "0" Then SendKeys "{Backspace}"
    Case 4
        If x <> 3 Then
            txtAllDate = txtAllDate + glbDateSeparator
            SendKeys "{End}"
        End If
    Case 5
        SendKeys "{Backspace}"
    End Select
End Select
Exit Sub
Can_Err:
If Err.Number = 91 Then
    Exit Sub
End If
End Sub

Private Function Get_Fields1(db As ADODB.Connection, TableName As String, UnField As String)
Dim rsTB As New ADODB.Recordset
Dim FdList As String
Dim x As Integer

rsTB.Open TableName, db
FdList = ""
For x = 0 To rsTB.Fields.count - 1
    If UnField = "" Or InStr(UCase(UnField) & ",", UCase(rsTB.Fields(x).name) & ",") = 0 Then
        FdList = FdList & ", " & rsTB.Fields(x).name
    End If
Next
Get_Fields1 = Mid(FdList, 2)
rsTB.Close
End Function


Public Sub Set_Control(Act As String, CurrentForm As Form, Optional rsTA As ADODB.Recordset, Optional SecondRecordset As Boolean)
'Act:if 'B' blank object and do not need pass the rsta control
'Act:if 'U' Update
'Act:if 'R' Refresh
On Error GoTo err_Control
Dim Ctrl  As Control
Dim cName As String

For Each Ctrl In CurrentForm
    If TypeOf Ctrl Is Label _
    Or TypeOf Ctrl Is ComboBox _
    Or TypeOf Ctrl Is CodeLookup _
    Or TypeOf Ctrl Is TextBox _
    Or TypeOf Ctrl Is DateLookup _
    Or TypeOf Ctrl Is EmployeeLookup _
    Or TypeOf Ctrl Is MaskEdBox _
    Or TypeOf Ctrl Is CheckBox _
    Or TypeOf Ctrl Is SSCheck Then
        
        If Ctrl.DataField <> "" Then
            If CurrentForm.name = "frmEESTATS" Then 'Ticket #15576
                If Not SecondRecordset Then ' rsDATA
                    If Left(Ctrl.DataField, 3) = "ER_" Then
                        GoTo ToNext
                    End If
                Else 'rsDAT_Other
                    If Left(Ctrl.DataField, 3) = "ED_" Then
                        GoTo ToNext
                    End If
                End If
            End If
        End If
    
        If TypeOf Ctrl Is Label Then
            If Ctrl.DataField <> "" Then
                If Act = "U" Then
                    If Len(Ctrl.Caption) = 0 Then
                        rsTA(Ctrl.DataField) = Null
                    Else
                        rsTA(Ctrl.DataField) = Ctrl
                    End If
                ElseIf Act = "B" Then
                    Ctrl = ""
                ElseIf Act = "R" Then
                    Ctrl = ""
                    If rsTA.EOF Or rsTA.BOF Then Exit Sub
                    If IsNull(rsTA(Ctrl.DataField)) Then
                        Ctrl = ""
                    Else
                        Ctrl = rsTA(Ctrl.DataField)
                    End If
                End If
            End If
        End If
        
        If TypeOf Ctrl Is CodeLookup Then
            If Ctrl.DataField <> "" Then
                If Act = "U" Then
                    If Len(Ctrl.Text) = 0 Then
                        rsTA(Ctrl.DataField) = Null
                    Else
                        rsTA(Ctrl.DataField) = Ctrl.Text
                    End If
                ElseIf Act = "B" Then
                    Ctrl.Text = ""
                    Ctrl.Caption = " "
                ElseIf Act = "R" Then
                    If rsTA.EOF Or rsTA.BOF Then Exit Sub
                    If IsNull(rsTA(Ctrl.DataField)) Then
                        Ctrl.Text = ""
                    Else
                        Ctrl.Text = rsTA(Ctrl.DataField)
                    End If
                End If
            End If
        End If
        
        If TypeOf Ctrl Is TextBox Then
            If Ctrl.DataField <> "" Then
                If Act = "U" Then
                    If Len(Ctrl.Text) = 0 Then
                        rsTA(Ctrl.DataField) = Null
                    Else
                        rsTA(Ctrl.DataField) = Ctrl.Text
                    End If
                ElseIf Act = "B" Then
                    Ctrl.Text = ""
                ElseIf Act = "R" Then
                    Ctrl.Text = ""
                    If rsTA.EOF Or rsTA.BOF Then Exit Sub
                    If IsNull(rsTA(Ctrl.DataField)) Then
                        
                        Ctrl.Text = ""
                    Else
                        Ctrl.Text = rsTA(Ctrl.DataField)
                    End If
                 End If
            End If
        End If
    
        If TypeOf Ctrl Is DateLookup Then
              If Ctrl.DataField <> "" Then
    
                If Act = "U" Then
                    If Len(Ctrl.Text) = 0 Then
                       rsTA(Ctrl.DataField) = Null
                    Else
                        rsTA(Ctrl.DataField) = Ctrl.Text
                    End If
                ElseIf Act = "B" Then
                    Ctrl.Text = ""
                ElseIf Act = "R" Then
                    Ctrl.Text = ""
                    If rsTA.EOF Or rsTA.BOF Then Exit Sub
                    If IsNull(rsTA(Ctrl.DataField)) Then
                        Ctrl.Text = ""
                    Else
                        Ctrl.Text = rsTA(Ctrl.DataField)
                    End If
                End If
            End If
        End If
        
        If TypeOf Ctrl Is EmployeeLookup Then
              If Ctrl.DataField <> "" Then
                If Act = "U" Then
                    If Len(Ctrl.Text) = 0 Then
                       rsTA(Ctrl.DataField) = Null
                    Else
                        rsTA(Ctrl.DataField) = Ctrl.Text
                    End If
                ElseIf Act = "B" Then
                    Ctrl.Text = ""
                ElseIf Act = "R" Then
                    Ctrl.Text = ""
                    If rsTA.EOF Or rsTA.BOF Then Exit Sub
                    If IsNull(rsTA(Ctrl.DataField)) Then
                        Ctrl.Text = ""
                    Else
                        Ctrl.Text = rsTA(Ctrl.DataField)
                    End If
                End If
            End If
        End If
        If TypeOf Ctrl Is MaskEdBox Then
            If Ctrl.DataField <> "" Then
                If Act = "U" Then
                    If Len(Ctrl) = 0 Then
                       rsTA(Ctrl.DataField) = Null
                    Else
                        If glbFrench Then
                            If IsNumeric(Ctrl) Then Ctrl = Replace(Ctrl, ",", ".")
                            rsTA(Ctrl.DataField) = Ctrl
                        Else
                            rsTA(Ctrl.DataField) = Ctrl
                        End If
                    End If
                ElseIf Act = "B" Then
                    Ctrl = ""
                ElseIf Act = "R" Then
                    Ctrl = ""
                    If rsTA.EOF Or rsTA.BOF Then Exit Sub
                    If IsNull(rsTA(Ctrl.DataField)) Then
                        Ctrl = ""
                    Else
                        Ctrl = rsTA(Ctrl.DataField)
                    End If
                End If
            End If
        End If
    
        If TypeOf Ctrl Is CheckBox Then
            If Ctrl.DataField <> "" Then
                If Act = "U" Then
                    rsTA(Ctrl.DataField) = Ctrl.Value
                ElseIf Act = "B" Then
                    Ctrl.Value = 0
                ElseIf Act = "R" Then
                    Ctrl.Value = 0
                    If rsTA.EOF Or rsTA.BOF Then Exit Sub
                    Ctrl.Value = IIf(rsTA(Ctrl.DataField).Value <> 0, 1, 0)
                End If
            End If
        End If
    
        If TypeOf Ctrl Is SSCheck Then
            If Ctrl.DataField <> "" Then
                If Act = "U" Then
                    rsTA(Ctrl.DataField) = Ctrl.Value
                ElseIf Act = "B" Then
                    Ctrl.Value = 0
                ElseIf Act = "R" Then
                    Ctrl.Value = 0
                    If rsTA.EOF Or rsTA.BOF Then Exit Sub
                    Ctrl.Value = IIf(rsTA(Ctrl.DataField).Value <> 0, True, False)
                End If
            End If
        End If
        
    End If

ToNext:

Next

Exit Sub
err_Control:
'    MsgBox Err.Description & " " & Ctrl.name
    Resume Next
End Sub

Public Sub INI_Controls(frmName As Form)
Dim Ctrl As Control
On Error GoTo INI_Error
For Each Ctrl In frmName
        If Left(Ctrl.name, 3) = "clp" Then
            Ctrl.TagTransfer = Ctrl.Tag
            Ctrl.AttachConnection gdbAdoIhr001
            Set Ctrl.MDIMain = MDIMain
            Ctrl.CompSerial = glbCompSerial
            Ctrl.UserID = glbUserID
            Ctrl.RptODBC = RptODBC_SQL
            Ctrl.IHRREPORTS = glbIHRREPORTS
            Ctrl.ADOIHRDB = glbAdoIHRDB
     '       Ctrl.SecurityAccessable = gSec_Inq_Requisition      'temp
            Ctrl.SecurityMaintainable = gSec_Upd_Requisition    '
            
'            If glbCompSerial = "S/N - 2363W" Then ' CITY OF K LAKES
'               If Ctrl.LookupType = HRTABL And Ctrl.TABLName = "EDRG" Then
'                    Ctrl.LookupType = PayrollCategory
'               End If
'            End If
            'Jaddy notes: this changes is for syncho
            
            Select Case Ctrl.LookupType
            Case HRTABL
                'Ticket #15312 - Begin
                Ctrl.SecurityAccessable = False
                Ctrl.SecurityMaintainable = False
                'Ticket #15312 - End
                Ctrl.SecurityAccessable = gSec_Inq_Master_Table(Ctrl.TablName)
                Ctrl.SecurityMaintainable = gSec_Upd_Master_Table(Ctrl.TablName)
                Ctrl.seleEMPCode = ""
                If Ctrl.TablName = "EDOR" Then
                    If Not Ctrl.MultiSelect Then
                        If glbVadim Then Ctrl.MaxLength = 1
                        'Except City of Niagara Falls, Dist. of Muskoka, Town of Marathon (Ticket #23001),Town of Lasalle, Town of Greater Napanee
                        If glbCompSerial = "S/N - 2276W" Or glbCompSerial = "S/N - 2373W" Or glbCompSerial = "S/N - 2330W" Or glbCompSerial = "S/N - 2379W" Or glbCompSerial = "S/N - 2447W" Then Ctrl.MaxLength = 4
                    End If
                    Ctrl.seleUnion = glbSeleUnion
                End If
                If Ctrl.TablName = "EDSE" Then Ctrl.seleSection = glbSeleSection
                If glbLinamar Then
                    If Ctrl.TablName = "EDRG" Or Ctrl.TablName = "EDSE" Or Ctrl.TablName = "EDSK" Or Ctrl.TablName = "BNCD" Then Ctrl.seleDiv = glbSeleDiv
                    If Ctrl.TablName = "EDRG" Then
                        Ctrl.MaxLength = 8
                        Ctrl.TextBoxWidth = 1000
                    End If
                    If Ctrl.TablName = "EDSK" Then Ctrl.MaxLength = 20
                    If Left(Ctrl.TablName, 2) = "HM" Then Ctrl.TextBoxWidth = 1200
                End If
                If glbVadim Then
                    If Ctrl.TablName = Vadim_PayType_TABLName Then Ctrl.MaxLength = 1
                    If Ctrl.TablName = Vadim_EmpType_TABLName Then Ctrl.MaxLength = 2
                End If
                
            Case Department
                Ctrl.SecurityAccessable = gSec_Inq_Departments
                Ctrl.SecurityMaintainable = gSec_Upd_Departments
                If glbVadim Then Ctrl.MaxLength = 4
            Case Division
                Ctrl.SecurityAccessable = gSec_Inq_Divisions
                Ctrl.SecurityMaintainable = gSec_Upd_Divisions
            Case GL
                Ctrl.SecurityAccessable = gSec_Inq_Ledgers
                Ctrl.SecurityMaintainable = gSec_Upd_Ledgers
            Case Province
                Ctrl.SecurityAccessable = True
                Ctrl.SecurityMaintainable = True
            Case NOC
                Ctrl.SecurityAccessable = gSec_Inq_Job_Classes
                Ctrl.SecurityMaintainable = gSec_Upd_Job_Classes
            Case Plan
                Ctrl.SecurityAccessable = gSec_Inq_EmploymentEQT
                Ctrl.SecurityMaintainable = True
            Case Job
                Ctrl.SecurityAccessable = gSec_Inq_Job_Master
                Ctrl.SecurityMaintainable = True
                Ctrl.seleEMPCode = ""
            Case JobMaster
                Ctrl.SecurityAccessable = gSec_Inq_Job_Master
                Ctrl.SecurityMaintainable = True
                Ctrl.seleEMPCode = ""
            Case SalaryDistribution
                Ctrl.SecurityAccessable = gSec_Inq_SalDist
                Ctrl.SecurityMaintainable = gSec_Upd_SalDist
            Case PayrollCategory
                Ctrl.SecurityAccessable = gSec_Inq_Payroll_Category
                Ctrl.SecurityMaintainable = gSec_Upd_Payroll_Category
            Case ChargeCode
                Ctrl.SecurityAccessable = gSec_Inq_Charge_Code
                Ctrl.SecurityMaintainable = gSec_Upd_Charge_Code
            Case ProjectCode
                Ctrl.SecurityAccessable = gSec_Inq_Project_Code
                Ctrl.SecurityMaintainable = gSec_Upd_Project_Code
            Case Machine
                Ctrl.SecurityAccessable = gSec_Inq_Machine
                Ctrl.SecurityMaintainable = gSec_Upd_Machine
            End Select
            Ctrl.SecurityForm = False
            If Ctrl.LookupType = Division Then Ctrl.seleDiv = glbSeleDiv
            If Ctrl.LookupType = Department Then Ctrl.seleDept = glbSeleDept
        ElseIf Left(Ctrl.name, 3) = "elp" Then
            Ctrl.TagTransfer = Ctrl.Tag
            Ctrl.AttachConnection gdbAdoIhr001, gdbAdoIhr001X
            Ctrl.ADOIHRDB = glbAdoIHRDB
            Set Ctrl.MDIMain = MDIMain
            Ctrl.CompSerial = glbCompSerial
            Ctrl.UserID = glbUserID
            Ctrl.RptODBC = RptODBC_SQL
            Ctrl.IHRREPORTS = glbIHRREPORTS
            Ctrl.SortType = glbSort
            Ctrl.seleDeptUn = glbSeleDeptUn
        ElseIf Left(Ctrl.name, 3) = "dlp" Then
             Ctrl.TagTransfer = Ctrl.Tag
        End If

 Next Ctrl
INI_Error:
    If Err = 5 Then Resume Next
End Sub
Function Date_SQL(xDate) As String
Date_SQL = " NULL "
If IsDate(xDate) Then
    If glbOracle Then
        Date_SQL = " TO_DATE('" & Format(xDate, "DD-MM-YYYY") & "','DD-MM-YYYY') "
    ElseIf glbSQL Then
        Date_SQL = " ('" & Format(xDate, "MMM DD,YYYY") & "') "
        If glbFrench Then
            Dim RtnDateStr As String
            
            RtnDateStr = " ('" & Format(xDate, "MMMM DD,YYYY") & "') "
            
            Date_SQL = TranslateDateString(RtnDateStr)
        End If
        
    Else
        Date_SQL = " CVDATE('" & xDate & "') "
    End If
End If
End Function
Function in_SQL(xDatabase) As String
in_SQL = ""
If Not glbOracle And Not glbSQL Then
    If InStr(UCase(xDatabase), "IHR001X.MDB") <> 0 Then
        in_SQL = " IN '" & glbIHRAUDIT & "' [;PWD=petman;DATABASE=" & glbIHRAUDIT & "] "
    ElseIf InStr(UCase(xDatabase), "IHR001.MDB") <> 0 Then
        in_SQL = " IN '" & glbIHRDB & "' [;PWD=petman;DATABASE=" & glbIHRDB & "] "
    Else
        in_SQL = " IN '" & xDatabase & "' "
    End If
End If
End Function

Function Field_SQL(xField) As String
Field_SQL = xField
If glbSQL Then Field_SQL = "[" & xField & "]"
End Function

Function Upper_SQL(xField) As String
Upper_SQL = xField
If glbOracle Then Upper_SQL = "UPPER(" & xField & ")"
End Function

Function getEGroup(ShowStr As String)
Dim vPosGroup

If Not glbSyndesis Then
     vPosGroup = "Position Group"
Else
     vPosGroup = "Position Grade"
End If

Select Case ShowStr
    Case lStr("Division"):              getEGroup = "{HR_DIVISION.Division_Name}"
    Case lStr("Department"):            getEGroup = "{HRDEPT.DF_NAME}"
    Case lStr("Location"):              getEGroup = "{HRTABL.TB_DESC}"
    Case lStr("Section"):               getEGroup = "{tblSec.TB_DESC}"
    Case lStr("Region"):                getEGroup = "{tblRegion.TB_DESC}"
    Case lStr("Administered By"):       getEGroup = "{tblAdminBy.TB_DESC}"
    Case lStr("Union"):                 getEGroup = "{tblUnion.TB_DESC}"
    Case lStr("Category"):              getEGroup = "{HREMP.ED_PT}"
    Case vPosGroup:                     getEGroup = "{tblPosGroup.TB_DESC}"     '"{HRJOB.JB_GRPCD}"
    Case "Position Code":               getEGroup = "{HRJOB.JB_DESCR}"
    Case lStr("Position Description"):  getEGroup = "{HRJOB.JB_DESCR}"
    Case "Employment Type":             getEGroup = "{@EMPTYPE}"
    Case "Employment Status":           getEGroup = "{tblEMP.TB_DESC}"
    Case "Home Line":                   getEGroup = "{LN_HOMES.TB_DESC}"
    Case "Employee Name":               getEGroup = "{@EFullName}"
    Case "Shift":                       getEGroup = "{HREMP.ED_SHIFT}"
    Case "Year of Birth":               getEGroup = "{@BirthYear}"
    
    Case "Performance Category":        getEGroup = "{tblCategory.TB_DESC}"
    Case "Performance Event":           getEGroup = "{tblEvent.TB_DESC}"
        
    'Hemu - 07/10/2003 Begin
    Case "Month of Birth":              getEGroup = "{@BirthMonth}"
    'Hemu - 07/10/2003 End
    
    'Ticket #14189 Frank 01/17/08
    Case lStr("G/L #"):                 getEGroup = "{HRGL.GL_DESCR}"
        
    'Ticket #18663 - Hemu
    Case lStr("Rept. Authority 1"):     getEGroup = "{@RepFullName1}"
    
    Case lStr("Rept. Authority 2"):     getEGroup = "{@RepFullName2}"
    Case lStr("Rept. Authority 3"):     getEGroup = "{@RepFullName3}"
    Case lStr("Rept. Authority 4"):     getEGroup = "{@RepFullName4}"
    Case "Date Sent":                   getEGroup = "{HR_FOLLOWUP_EMAIL_LOG.FL_SENTDT}"
    
    Case "Effective Date":              getEGroup = "{HR_SCHEDULER.SD_EDATE}"
    
    Case lStr("Skills"):                getEGroup = "{HREMPSKL.SE_SKILL}"
    
    Case lStr("Course Code"):           getEGroup = "{tblCourseCode.TB_DESC}"
    
    Case "Length of Service":            getEGroup = "{@Seniority}"     'Ticket #22682: Release 8.0
    
    Case lStr("Account Code"):          getEGroup = "{HR_PROJECT_CODE.DESCRIPTION}"
    
    'Release 8.1
    Case "Document Type":               getEGroup = "{tblDocType.TB_DESC}"
    
    Case "(none)":                      getEGroup = "(none)"
End Select
  
End Function

Public Function MonthDiff(startDate As Date, EndDate As Date)
Dim xSDay, xEday, xLDay

xSDay = Day(startDate)
xLDay = Day(MonthLastDate(EndDate))
If xSDay > xLDay Then xSDay = xLDay
xEday = Day(EndDate)

MonthDiff = DateDiff("m", startDate, EndDate)
If xSDay <> xEday Then
    MonthDiff = MonthDiff - (xSDay - xEday) / xLDay
End If

End Function

Public Function MonthLastDate(xDate As Date)
Dim x, DateStr
For x = 28 To 31
    DateStr = Format(xDate, "mmm," & x & " yyyy")
    If IsDate(DateStr) Then
        MonthLastDate = DateStr
    Else
        Exit For
    End If
Next
End Function

Public Function DaysInMonth(xDate As Date)
    DaysInMonth = DateSerial(Year(xDate), month(xDate) + 1, 1) - DateSerial(Year(xDate), month(xDate), 1)
End Function

Public Function DecryptPasswordMultiLang_First(Value As String)
Dim valtmp, I
If Len(Value) > 0 And Len(Value) < 255 Then
  For I = 1 To Len(Value)
    valtmp = Asc(Mid(Value, I, 1))
    valtmp = Chr(valtmp - 1)
    valtmp = Mid(Value, 1, I - 1) & valtmp & Mid(Value, I + 1, Len(Value) - I)
    Value = valtmp
  Next I
End If
DecryptPasswordMultiLang_First = Value
If Len(DecryptPasswordMultiLang_First) = 255 Then
    DecryptPasswordMultiLang_First = ""
End If
End Function
Public Function DecryptPasswordMultiLang(Value As String)
Dim valtmp, I
If Len(Value) > 1 Then
    Value = Mid(Value, 2, Len(Value) - 1)
    Value = Mid(Value, 1, Len(Value) - 1)
End If
If Len(Value) > 0 Then
  valtmp = ""
  For I = 1 To Len(Value)
    valtmp = Mid(Value, I, 1) & valtmp
  Next I
  Value = valtmp
End If
If Len(Value) > 0 And Len(Value) < 255 Then
  For I = 1 To Len(Value)
    valtmp = Asc(Mid(Value, I, 1))
    valtmp = Chr(valtmp - 2)
    valtmp = Mid(Value, 1, I - 1) & valtmp & Mid(Value, I + 1, Len(Value) - I)
    Value = valtmp
  Next I
End If
DecryptPasswordMultiLang = Value
If Len(DecryptPasswordMultiLang) = 255 Then
    DecryptPasswordMultiLang = ""
End If

End Function
Public Function DecryptPassword(ByVal Value As String)
Dim valtmp, I
If Not (glbCompSerial = "S/N - 2336W") Or glbDBPassFlag Then 'No Encrypted passwords for WHSCC
    If Len(Value) > 0 And Len(Value) < 255 Then
      For I = 1 To Len(Value)
    '    valtmp = Value
        valtmp = Asc(Mid(Value, I, 1))
        If valtmp - 80 < 0 Then
            valtmp = Chr(0)
        Else
            valtmp = Chr(valtmp - 80)
        End If
        valtmp = Mid(Value, 1, I - 1) & valtmp & Mid(Value, I + 1, Len(Value) - I)
        Value = valtmp
      Next I
    End If
End If
DecryptPassword = Value
glbDBPassFlag = False
If Len(DecryptPassword) = 255 Then
    DecryptPassword = ""
End If
End Function
Public Function EncryptPasswordMultiLang_First(Value As String)
Dim I, valtmp
If Len(Value) > 0 Then
  For I = 1 To Len(Value)
      valtmp = Value
      valtmp = Asc(Mid(Value, I, 1))
      
         valtmp = Chr(valtmp + 1)
     
      valtmp = Mid(Value, 1, I - 1) & valtmp & Mid(Value, I + 1, Len(Value) - I)
      Value = valtmp
  Next I
End If
EncryptPasswordMultiLang_First = Value
End Function
Public Function EncryptPasswordMultiLang(Value As String)
Dim I, valtmp
valtmp = ""
If Len(Value) > 0 Then
  For I = 1 To Len(Value)
    valtmp = Value
    valtmp = Asc(Mid(Value, I, 1))
    valtmp = Chr(valtmp + 2)
    valtmp = Mid(Value, 1, I - 1) & valtmp & Mid(Value, I + 1, Len(Value) - I)
    Value = valtmp
  Next I
End If
If Len(Value) > 0 Then
  valtmp = ""
  For I = 1 To Len(Value)
    valtmp = Mid(Value, I, 1) & valtmp
  Next I
End If
Value = "/" & valtmp & "f"
EncryptPasswordMultiLang = Value

End Function
Public Function EncryptPassword(Value As String)
Dim I, valtmp
If Not (glbCompSerial = "S/N - 2336W") Or glbDBPassFlag Then
    If Len(Value) > 0 Then
      For I = 1 To Len(Value)
          valtmp = Value
          valtmp = Asc(Mid(Value, I, 1))
          
             valtmp = Chr(valtmp + 80)
         
          valtmp = Mid(Value, 1, I - 1) & valtmp & Mid(Value, I + 1, Len(Value) - I)
          Value = valtmp
      Next I
    End If
End If

EncryptPassword = Value
glbDBPassFlag = False
End Function

Public Function GetMassUpdateSecurities(xSecurity As String, xUserID) As Boolean
Dim rsSECM As New ADODB.Recordset
Dim rsSEC As New ADODB.Recordset
Dim SQLQ
Dim zSecurity
Dim xTemplate As String

'????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
xTemplate = ""
xTemplate = Get_Template(xUserID)

If xTemplate = "" Or xTemplate = "TEMPLATE" Then
    SQLQ = "SELECT * FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
Else
    '????Ticket #24808 -  Retrieve template's security profile
    SQLQ = "SELECT * FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
End If
SQLQ = SQLQ & " AND " & Field_SQL("FUNCTION") & "='" & xSecurity & "'"

rsSECM.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic

If rsSECM.EOF Then
    zSecurity = Replace(xSecurity, "MassUpdate", "Update")
    If zSecurity = "Attendance_His_Update" Then zSecurity = "Attendance_Update"
    SQLQ = "SELECT ACCESSABLE FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
    If zSecurity <> "Job_Master_Update" Then
        SQLQ = SQLQ & " AND " & Field_SQL("FUNCTION") & "='" & zSecurity & "'"
    Else
        SQLQ = SQLQ & " AND (" & Field_SQL("FUNCTION") & "='" & zSecurity & "'"
        SQLQ = SQLQ & " OR " & Field_SQL("FUNCTION") & "='Salary_Update'"
        SQLQ = SQLQ & " OR " & Field_SQL("FUNCTION") & "='Position_Update')"
    End If

    rsSEC.Open SQLQ, gdbAdoIhr001, adOpenStatic
    rsSECM.AddNew
    rsSECM("COMPNO") = "001"
    rsSECM("USERID") = xUserID
    rsSECM("FUNCTION") = xSecurity
    If rsSEC.EOF Then
        rsSECM("ACCESSABLE") = 0
    Else
        If zSecurity = "Job_Master_Update" Then
            If rsSEC.RecordCount = 3 Then
                rsSECM("ACCESSABLE") = 1
                Do Until rsSEC.EOF
                    If rsSEC("ACCESSABLE") = 0 Then rsSECM("ACCESSABLE") = 0: Exit Do
                    rsSEC.MoveNext
                Loop
            Else
                rsSECM("ACCESSABLE") = 0
            End If
        Else
            rsSECM("ACCESSABLE") = rsSEC("ACCESSABLE")
        End If
    End If
    rsSECM.Update
    rsSEC.Close
End If
GetMassUpdateSecurities = rsSECM("ACCESSABLE") <> 0
rsSECM.Close
End Function

Public Sub BTIPoint(SelStr, Optional RecalFlag As Boolean)  'Ticket #10210 BTI Points Calculate Function
Dim rsTAtt As New ADODB.Recordset
Dim rsBEmp As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim rsMain As New ADODB.Recordset
Dim SQLQ, xNum, xCode, xSec, xEmpNo
Dim xUnexFlag, xEmlFlag, xUnexVal, xExcuVal, xEmlVal, glbDiv, glbYear, xYear, glbPointType
Dim I, xNTot
Dim xDOA
Dim xABSCarryover, xLLECarryover
Dim xPointVal
Dim glbBTIDate, xQRED_Date
    glbBTIDate = CVDate(GetMonth("Jan") & " 1, 2006")
    If IsDate(glbCompEdFrom) Then
        If CVDate(glbCompEdFrom) > CVDate(glbBTIDate) Then
            glbBTIDate = glbCompEdFrom
        End If
    End If

    'Points Recalculate on Company master - SelStr = "ALL"
    'Reset the Point, Unexcused, Excused, EML flags from Attendance Reason codes - Begin
    'For attendance reason codes with point or EML only
    If RecalFlag Then
        'SQLQ = "SELECT ED_EMPNBR,ED_DIV,ED_SECTION FROM HREMP WHERE ED_SECTION='HRLY'"
        SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE (1=1) "
        SQLQ = SQLQ & "AND AD_DOA >=" & Date_SQL(CVDate(glbBTIDate)) & " "  'New Policy begins on Jan 1, 2006
        SQLQ = SQLQ & "AND AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='ADRE' AND (TB_USR2>0 OR TB_USR3<>0)) " 'TB_USR3 -EML flag
        SQLQ = SQLQ & "AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SECTION='HRLY' ) "
        If IsNumeric(SelStr) Then
            SQLQ = SQLQ & "AND AD_EMPNBR = " & SelStr & " "
        End If
        SQLQ = SQLQ & "ORDER BY AD_DOA "
        If rsTAtt.State <> 0 Then rsTAtt.Close
        rsTAtt.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsTAtt.EOF Then
            I = 0
            xNTot = rsTAtt.RecordCount
        End If
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait"
        Do While Not rsTAtt.EOF
            DoEvents
            MDIMain.panHelp(0).FloodPercent = (I / xNTot) * 100: I = I + 1
            SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME='ADRE' AND TB_KEY ='" & rsTAtt("AD_REASON") & "' "
            rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsTemp.EOF Then
                If rsTemp("TB_USR2") > 0 Then 'Points
                    rsTAtt("AD_POINT") = rsTemp("TB_USR2")
                Else
                    rsTAtt("AD_POINT") = Null
                End If
                If rsTemp("TB_USR3") <> 0 Then 'EML
                    rsTAtt("AD_EMELEA") = rsTemp("TB_USR3")
                Else
                    rsTAtt("AD_EMELEA") = 0
                End If
                rsTAtt("AD_INDICATOR") = rsTemp("TB_INDICATOR")
                rsTAtt("AD_SEN") = rsTemp("TB_SEN")
                rsTAtt.Update
            End If
            rsTemp.Close
            rsTAtt.MoveNext
        Loop
        rsTAtt.Close
    End If
    'Reset the Point, Unexcused, Excused, EML flags from Attendance Reason codes - End

    If IsNumeric(SelStr) Then
        xEmpNo = SelStr
        SQLQ = "SELECT ED_EMPNBR,ED_DIV,ED_SECTION FROM HREMP WHERE ED_SECTION='HRLY' AND ED_EMPNBR=" & SelStr & ""
    Else
        If SelStr = "ALL" Then
            SQLQ = "SELECT ED_EMPNBR,ED_DIV,ED_SECTION FROM HREMP WHERE ED_SECTION='HRLY'"
        Else
            SQLQ = "SELECT ED_EMPNBR,ED_DIV,ED_SECTION FROM HREMP WHERE ED_SECTION='HRLY' AND " & SelStr & ""
        End If
    End If
    rsMain.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsMain.EOF Then
        I = 0
        xNTot = rsMain.RecordCount
    End If

    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    Do While Not rsMain.EOF
        DoEvents
        MDIMain.panHelp(0).FloodPercent = (I / xNTot) * 100: I = I + 1
        xEmpNo = rsMain("ED_EMPNBR")
        'Get glbDiv
        glbDiv = rsMain("ED_DIV")
        xSec = rsMain("ED_SECTION")

        
        'Get an Attendance recordset for this employee
        SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR=" & xEmpNo & " "
        SQLQ = SQLQ & "AND AD_DOA >=" & Date_SQL(CVDate(glbBTIDate)) & " "  'New Policy begins on Jan 1, 2006
        'SQLQ = SQLQ & "AND AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='ADRE' AND TB_USR2>0) "
        SQLQ = SQLQ & "AND (AD_POINT<>0) " 'show all attendance records with point <> 0
        SQLQ = SQLQ & "ORDER BY AD_DOA "
        If rsBEmp.State <> 0 Then rsBEmp.Close
        rsBEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
        Screen.MousePointer = HOURGLASS
        Do While Not rsBEmp.EOF
            xDOA = rsBEmp("AD_DOA")
            glbYear = Trim(Str(Year(rsBEmp("AD_DOA"))))
    
            xCode = rsBEmp("AD_REASON")
            xPointVal = ""
            If IsNumeric(rsBEmp("AD_POINT")) Then
                If rsBEmp("AD_POINT") > 0 Then
                    xPointVal = rsBEmp("AD_POINT")
                End If
            End If
            'Total Points -Begin
            xNum = 0
            SQLQ = "SELECT SUM(AD_POINT) AS TOTNUM FROM HR_ATTENDANCE WHERE AD_EMPNBR=" & xEmpNo & " "
            SQLQ = SQLQ & "AND (AD_POINT<>0) "
            SQLQ = SQLQ & "AND to_char(AD_DOA,'yyyy')='" & glbYear & "' "
            SQLQ = SQLQ & "AND AD_DOA <=" & Date_SQL(xDOA) & " "
            SQLQ = SQLQ & "AND AD_DOA >=" & Date_SQL(CVDate(glbBTIDate)) & " "
            rsTAtt.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsTAtt.EOF Then
                If Not IsNull(rsTAtt("TOTNUM")) Then
                    xNum = rsTAtt("TOTNUM")
                End If
            End If
            rsTAtt.Close
            'Total Points -End
            
            'Ticket #10712
            'If a period of 3 months is achieved without an unexcused absence - Begin
            'the employee will receive a 1 point reduction.
            If Val(glbYear) = Year(glbCompEdFrom) Then '= Year(Date) Then
                xQRED_Date = DateAdd("d", 90, CVDate(xDOA))
                If CVDate(xQRED_Date) > Date Then
                    xQRED_Date = Date
                End If
                xExcuVal = BIT3Months(xEmpNo, glbYear, xNum, xQRED_Date, glbBTIDate) 'BIT3Months(xEmpNo, glbYear, xNum, Date)
                xNum = xNum - xExcuVal
            End If
            'If a period of 3 months is achieved without an unexcused absence - End
            
            If xPointVal > 0 Then
                If xNum > 0 Then
                    'Check if the AttCode with point, No -> Skip; Yes -> go ahead
                    SQLQ = "SELECT TB_KEY FROM HRTABL WHERE TB_NAME='ADRE' AND TB_USR2>0 AND TB_KEY ='" & xCode & "' "
                    If rsTAtt.State <> 0 Then rsTAtt.Close
                    rsTAtt.Open SQLQ, gdbAdoIhr001, adOpenStatic
                    If Not rsTAtt.EOF Then
                        Call ModEmpCounsel(xEmpNo, glbYear, glbPointType, xNum, glbDiv, xDOA, xCode)
                    End If
                    rsTAtt.Close
                End If
            End If 'End for Points Check
            rsBEmp.MoveNext
        Loop
        rsMain.MoveNext
    Loop
    MDIMain.panHelp(0).FloodPercent = 100
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    Screen.MousePointer = DEFAULT
End Sub


Public Sub ModEmpPoint(xEmpNo, xYear, xType, xNum)
Dim rsPAtt As New ADODB.Recordset
Dim SQLQ
    SQLQ = "SELECT * FROM HR_EMPPOINTS WHERE EP_EMPNBR=" & xEmpNo & " "
    SQLQ = SQLQ & "AND EP_YEAR='" & xYear & "' "
    rsPAtt.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsPAtt.EOF Then
        rsPAtt.AddNew
        rsPAtt("EP_COMPNO") = "001"
        rsPAtt("EP_EMPNBR") = xEmpNo
        rsPAtt("EP_YEAR") = Trim(xYear)
        rsPAtt("EP_POINT") = 0
        rsPAtt("EP_LEPOINT") = 0
        rsPAtt("EP_EML") = 0
        rsPAtt("EP_LUSER") = glbUserID
        rsPAtt("EP_LDATE") = Format(Now, "Short Date")
        rsPAtt("EP_LTIME") = Time$
    End If
    If xType = "ABS" Then
        rsPAtt("EP_POINT") = xNum '+ 1
    End If
    If xType = "LLE" Then
        rsPAtt("EP_LEPOINT") = xNum '+ 1
    End If
    If xType = "EML" Then
        rsPAtt("EP_EML") = xNum '+ 1
    End If
    rsPAtt.Update
    rsPAtt.Close
End Sub

Public Function BIT3Months(xEmpNo, xYear, xNum, xDate, xBTIDate)
Dim rsPAtt As New ADODB.Recordset
Dim rsPCounsel As New ADODB.Recordset
Dim rsPTemp As New ADODB.Recordset
Dim SQLQ
Dim xDays, xRedPoints, xRedDate
Dim xNo90
Dim BIT3MonthStart

    'If a period of 3 months is achieved without an unexcused absence, the employee will receive a 1 point reduction.
    'This would therefore mean that it could occur 4 times per year.
    'Point Reduction - Begin
        'Get the last date of Absence
        SQLQ = "SELECT AD_EMPNBR,AD_DOA, AD_POINT,AD_DISCIPLINE FROM HR_ATTENDANCE WHERE AD_EMPNBR=" & xEmpNo & " "
        'SQLQ = SQLQ & "AND to_char(AD_DOA,'yyyy')='" & xYear & "' AND NOT (AD_DISCIPLINE IS NULL) AND AD_POINT>0 "
        'If there is no points within 3 months, and then do points reduction.
        'Ticket #12938 - Begin - should check the last 3 months records even in last year
        ''SQLQ = SQLQ & "AND to_char(AD_DOA,'yyyy')='" & xYear & "' AND AD_POINT<>0 "
        BIT3MonthStart = DateAdd("d", -90, CVDate(xBTIDate))
        SQLQ = SQLQ & "AND AD_POINT<>0 "
        SQLQ = SQLQ & "AND AD_REASON<>'2200' " 'Year End Carry Over
        SQLQ = SQLQ & "AND AD_DOA >" & Date_SQL(BIT3MonthStart) & " "
        'Ticket #12938 - End
        SQLQ = SQLQ & "AND AD_DOA <" & Date_SQL(xDate) & " "
        SQLQ = SQLQ & "ORDER BY AD_DOA DESC "
        rsPCounsel.Open SQLQ, gdbAdoIhr001, adOpenStatic
        xRedPoints = 0 '""
        If Not rsPCounsel.EOF Then
            xDays = DateDiff("d", CVDate(rsPCounsel("AD_DOA")), CVDate(xDate))
            xRedPoints = Int(xDays / 90)
            xNo90 = xRedPoints
            If xRedPoints > 0 Then
                If xRedPoints > xNum Then
                    xRedPoints = xNum
                End If
                If Not xRedPoints = 0 Then 'Ticket #12046, don't create a new record when xRedPoints = 0
                    'Insert the Reduce points - Begin '
                    'xRedDate = DateAdd("d", 90 * xRedPoints, CVDate(rsPCounsel("AD_DOA"))) '
                    xRedDate = DateAdd("d", 90 * xNo90, CVDate(rsPCounsel("AD_DOA")))
                    SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR=" & xEmpNo & " "
                    SQLQ = SQLQ & "AND AD_DOA =" & Date_SQL(xRedDate) & " "
                    SQLQ = SQLQ & "AND AD_REASON = 'QRED' "
                    rsPTemp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If rsPTemp.EOF Then
                        rsPTemp.AddNew
                        rsPTemp("AD_COMPNO") = "001"
                        rsPTemp("AD_EMPNBR") = xEmpNo
                        rsPTemp("AD_DOA") = xRedDate
                        rsPTemp("AD_REASON") = "QRED"
                        rsPTemp("AD_POINT") = -xRedPoints
                        rsPTemp("AD_HRS") = 0
                        rsPTemp("AD_LDATE") = Date
                        rsPTemp("AD_LTIME") = Time$
                        rsPTemp("AD_LUSER") = glbUserID
                        rsPTemp.Update
                    End If
                    rsPTemp.Close
                    'Insert the Reduce points - End
                    'xNum = xNum - xRedPoints
                End If
            End If
        Else 'Check the Attendance History too 'Ticket #13009
            rsPCounsel.Close
            SQLQ = "SELECT AH_EMPNBR,AH_DOA, AH_POINT,AH_DISCIPLINE FROM HR_ATTENDANCE_HISTORY WHERE AH_EMPNBR=" & xEmpNo & " "
            BIT3MonthStart = DateAdd("d", -90, CVDate(xBTIDate))
            SQLQ = SQLQ & "AND AH_POINT<>0 "
            SQLQ = SQLQ & "AND AH_REASON<>'2200' " 'Year End Carry Over
            SQLQ = SQLQ & "AND AH_DOA >" & Date_SQL(BIT3MonthStart) & " "
            SQLQ = SQLQ & "AND AH_DOA <" & Date_SQL(xDate) & " "
            SQLQ = SQLQ & "ORDER BY AH_DOA DESC "
            rsPCounsel.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsPCounsel.EOF Then
                xDays = DateDiff("d", CVDate(rsPCounsel("AH_DOA")), CVDate(xDate))
                xRedPoints = Int(xDays / 90)
                xNo90 = xRedPoints
                If xRedPoints > 0 Then
                    If xRedPoints > xNum Then
                        xRedPoints = xNum
                    End If
                    If Not xRedPoints = 0 Then
                        'Insert the Reduce points - Begin '
                        'xRedDate = DateAdd("d", 90 * xRedPoints, CVDate(rsPCounsel("AD_DOA"))) '
                        xRedDate = DateAdd("d", 90 * xNo90, CVDate(rsPCounsel("AH_DOA")))
                        SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR=" & xEmpNo & " "
                        SQLQ = SQLQ & "AND AD_DOA =" & Date_SQL(xRedDate) & " "
                        SQLQ = SQLQ & "AND AD_REASON = 'QRED' "
                        rsPTemp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If rsPTemp.EOF Then
                            rsPTemp.AddNew
                            rsPTemp("AD_COMPNO") = "001"
                            rsPTemp("AD_EMPNBR") = xEmpNo
                            rsPTemp("AD_DOA") = xRedDate
                            rsPTemp("AD_REASON") = "QRED"
                            rsPTemp("AD_POINT") = -xRedPoints
                            rsPTemp("AD_HRS") = 0
                            rsPTemp("AD_LDATE") = Date
                            rsPTemp("AD_LTIME") = Time$
                            rsPTemp("AD_LUSER") = glbUserID
                            rsPTemp.Update
                        End If
                        rsPTemp.Close
                        'Insert the Reduce points - End
                        'xNum = xNum - xRedPoints
                    End If
                End If
            
            End If
        End If
        rsPCounsel.Close
        BIT3Months = xRedPoints
    'Point Reduction - End

End Function

Public Sub ModEmpCounsel(xEmpNo, xYear, xType, xNum, xDiv, xDOA, xAttCode)
Dim rsPAtt As New ADODB.Recordset
Dim rsPCounsel As New ADODB.Recordset
Dim rsPTemp As New ADODB.Recordset
Dim SQLQ, xCounlStep, xCounlReason, xCounlAction
Dim xFLAG As Boolean
Dim xTarget
Dim xDays, xRedPoints, xRedDate
'
    xCounlStep = "": xCounlReason = "": xCounlAction = "": xTarget = ""
    SQLQ = "SELECT * FROM HR_COUNSEL_ABSENCE WHERE CL_DIVISION ='" & xDiv & "' ORDER BY CL_TARGET"
    If rsPAtt.State <> 0 Then rsPAtt.Close
    rsPAtt.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsPAtt.EOF
        If xNum >= rsPAtt("CL_TARGET") Then 'Ticket# 9356
        'If xNum = rsPAtt("CL_TARGET") Then
            xCounlStep = rsPAtt("CL_TYPE")
            xCounlReason = rsPAtt("CL_REASON")
            xCounlAction = rsPAtt("CL_ACTION")
            xTarget = rsPAtt("CL_TARGET")
        End If
        rsPAtt.MoveNext
    Loop
    rsPAtt.Close
    If Len(xCounlStep) = 0 Then Exit Sub 'No Counsel Type found for this Point
    
    'Search in Emp Counsel table For this Attendance record, if it exists and then exit sub
    SQLQ = "SELECT * FROM HR_COUNSEL WHERE CL_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND (CL_INCDATE)=" & Date_SQL(xDOA) & " "
    If rsPCounsel.State <> 0 Then rsPCounsel.Close
    rsPCounsel.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsPCounsel.EOF Then
        Exit Sub
    End If
    rsPCounsel.Close
    
    'Search in Emp Counsel table
    'SQLQ = "SELECT * FROM HR_COUNSEL WHERE CL_EMPNBR = " & xEmpNo & " "
    'SQLQ = SQLQ & "AND (CL_INCDATE)=" & Date_SQL(xDOA) & " "
    'If rsPCounsel.State <> 0 Then rsPCounsel.Close
    'rsPCounsel.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    'If rsPCounsel.EOF Then
        'To check if there is Counsel record which no Counsel Action by HR -Begin
        SQLQ = "SELECT * FROM HR_COUNSEL WHERE CL_EMPNBR = " & xEmpNo & " "
        'Comment this to fix a bug - If there is a Counselling date is null for this year,
        'don't create a new counselling record untill this case is over
        'SQLQ = SQLQ & "AND CL_TYPE <>'" & xCounlStep & "' "
        SQLQ = SQLQ & "AND to_char(CL_INCDATE,'yyyy')='" & xYear & "' "
        SQLQ = SQLQ & "AND CL_COUDATE IS NULL "
        xFLAG = False
        rsPTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If rsPTemp.EOF Then
            xFLAG = True
            ''Ticket# 9356, check the last record if it's duplicate record for this type
            rsPTemp.Close
            'SQLQ = "SELECT * FROM HR_COUNSEL WHERE CL_EMPNBR = " & xEmpNo & " "
            'SQLQ = SQLQ & "AND to_char(CL_INCDATE,'yyyy')=" & xYear & " "
            'SQLQ = SQLQ & "ORDER BY CL_INCDATE DESC "
            SQLQ = "SELECT AD_EMPNBR, AD_POINT,AD_DISCIPLINE FROM HR_ATTENDANCE WHERE AD_EMPNBR=" & xEmpNo & " "
            SQLQ = SQLQ & "AND to_char(AD_DOA,'yyyy')='" & xYear & "' AND AD_POINT<>0 AND AD_DOA <" & Date_SQL(xDOA) & " "
            'SQLQ = SQLQ & "AND to_char(AD_DOA,'yyyy')=" & xYear & "AND AD_POINT<>0 AND AD_DOA <" & Date_SQL(xDOA) & " "
            SQLQ = SQLQ & "ORDER BY AD_DOA DESC "
            rsPTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsPTemp.EOF Then
                If rsPTemp("AD_DISCIPLINE") = xCounlStep Then
                    xFLAG = False
                End If
            End If
            ''Ticket# 9356 End
        End If
        rsPTemp.Close
        'To check if there is Counsel record which no Counsel Action by HR- End
        If xFLAG Then
            SQLQ = "SELECT * FROM HR_COUNSEL WHERE CL_EMPNBR = " & xEmpNo & " "
            SQLQ = SQLQ & "AND (1=2) " 'Open a blank recordset
            If rsPCounsel.State <> 0 Then rsPCounsel.Close
            rsPCounsel.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If rsPCounsel.EOF Then
                rsPCounsel.AddNew
                rsPCounsel("CL_COMPNO") = "001"
                rsPCounsel("CL_EMPNBR") = xEmpNo
                rsPCounsel("CL_TYPE") = xCounlStep
                rsPCounsel("CL_REASON") = xCounlReason
                'rsPCounsel("CL_COUDATE") = xDOA
                rsPCounsel("CL_INCDATE") = xDOA
                rsPCounsel("CL_ATTDATE") = xDOA
                rsPCounsel("CL_ATTREASON") = xAttCode
                rsPCounsel("CL_REASON") = xCounlReason
                rsPCounsel("CL_COMMENTS") = xCounlAction
                rsPCounsel("CL_LDATE") = Format(Now, "SHORT DATE")
                rsPCounsel("CL_LTIME") = Time$
                rsPCounsel("CL_LUSER") = glbUserID
                rsPCounsel.Update
            End If
            rsPCounsel.Close
            SQLQ = "UPDATE HR_ATTENDANCE SET AD_DISCIPLINE = '" & xCounlStep & "' "
            SQLQ = SQLQ & "WHERE AD_EMPNBR=" & xEmpNo & " AND AD_DOA = " & Date_SQL(xDOA) & " "
            SQLQ = SQLQ & "AND AD_REASON = '" & xAttCode & "' "
            gdbAdoIhr001.Execute SQLQ
            'Sending email ...
            Call EmailSendingForBTI(xDiv, xType, xCounlStep, xNum, xEmpNo, xDOA, xAttCode)
            '(xEmpNo, xYear, xType, xNum, xDiv, xDOA, xAttCode)
        End If
    'End If
    'rsPCounsel.Close
    
End Sub


Private Sub EmailSendingForBTI(xDiv, xPointType, xCType, xNum, xEmpNo, xDOA, xATTReason)
Dim rsTemp As New ADODB.Recordset
Dim MailBody As String
Dim AttDesc As String, TypeDesc As String
Dim xEmpName, SQLQ, xTStr, AbortTerm
Dim xToEmai, xCCEmail
    Exit Sub
    If xPointType = "ABS" Then
        SQLQ = "SELECT * FROM HR_COUNSEL_ABSENCE WHERE CL_DIVISION ='" & xDiv & "' AND CL_TYPE = '" & xCType & "' ORDER BY CL_TARGET"
    Else
        SQLQ = "SELECT * FROM HR_COUNSEL_LE WHERE CL_DIVISION ='" & xDiv & "' AND CL_TYPE = '" & xCType & "' ORDER BY CL_TARGET"
    End If
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        xToEmai = rsTemp("CL_EMAIL1")
        xCCEmail = ""
        If Not IsNull(rsTemp("CL_EMAIL2")) Then
            xCCEmail = xCCEmail & rsTemp("CL_EMAIL2")
        End If
        If Not IsNull(rsTemp("CL_EMAIL3")) Then
            If Len(xCCEmail) = 0 Then
                xCCEmail = xCCEmail & rsTemp("CL_EMAIL3")
            Else
                xCCEmail = xCCEmail & ";" & rsTemp("CL_EMAIL3")
            End If
        End If
    End If
    rsTemp.Close
    If Len(xToEmai) = 0 Then Exit Sub
    
    glbWFCEmailTest = False
    SQLQ = "SELECT ED_SURNAME,ED_FNAME FROM HREMP WHERE ED_EMPNBR =" & xEmpNo & " "
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        xEmpName = rsTemp("ED_SURNAME") & ", " & rsTemp("ED_FNAME")
    End If
    rsTemp.Close
    If xPointType = "ABS" Then xTStr = "Absence " Else xTStr = "L/LE "
    MailBody = "The employee below has been counselled." & vbCrLf & vbCrLf
    MailBody = MailBody & "Employee #: " & xEmpNo & vbCrLf
    MailBody = MailBody & "Name: " & xEmpName & vbCrLf
    MailBody = MailBody & "Counselling Type: " & xCType & vbCrLf
    MailBody = MailBody & "Incident/Attendance Date: " & xDOA & vbCrLf
    MailBody = MailBody & "Attendance Reason: " & xATTReason & vbCrLf
    MailBody = MailBody & "Total " & xTStr & "Points: " & xNum & vbCrLf
    
    frmSendEmail.txtBody.Text = MailBody

    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = "Sending email..."
    'frmSendEmail.txtTo.Text = "hotline@woodbridgegroup.com"
    frmSendEmail.txtTo.Text = xToEmai
    If Len(xCCEmail) > 0 Then
        frmSendEmail.txtCC.Text = xCCEmail
    End If
    frmSendEmail.Tag = ""
    frmSendEmail.cmdSend_Click
    Do
        DoEvents
    Loop Until frmSendEmail.Tag <> ""   ' MC - dkostka - 05/03/01 - Changed from = "DONE" to <> ""
    ' AC - dkostka - 05/03/01 - Added checking to make sure the email went through,
    '   otherwise refuse to terminate the employee.
    If frmSendEmail.Tag = "DONE" Then
        Unload frmSendEmail
        AbortTerm = False
    Else
        Unload frmSendEmail
        AbortTerm = True
    End If
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(0).FloodType = 1


End Sub

Public Function GetLeapYear(yr As Long) As Boolean
    GetLeapYear = False
    If yr / 4 = Int(yr / 4) Then
        GetLeapYear = True
    End If
End Function

Public Function buildSec() As String
    'created by Bryan 14/Oct/05 Ticket#9424
    'build working table for Terminated comments
    'changed by Bryan 21/Nov/05 Ticket#9806
    'working table not good for Access, changed from a working table to just a list of acceptable codes
    'Dim rsIN As New ADODB.Recordset
    Dim rsOUT As New ADODB.Recordset
    Dim strSQL As String
    Dim retVal As String
    Dim c As Long
    Dim xTemplate As String
    
    '????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
    xTemplate = ""
    xTemplate = Get_Template(glbUserID)
    
    
'
'    strSQL = "DELETE FROM HRCOMWRK WHERE WRKEMP='" & glbUserID & "'"
'    gdbAdoIhr001W.Execute strSQL
'
'    rsOUT.Open "HRCOMWRK", gdbAdoIhr001W, adOpenDynamic, adLockOptimistic, adCmdTable
'    strSQL = "SELECT CO_COMPNO, CO_EMPNBR, CO_EDATE, CO_TYPE_TABL, CO_TYPE, CO_COMMENT_ID, CO_COMMENTS, CO_LDATE, CO_LTIME, CO_LUSER, TERM_SEQ from Term_COMMENTS WHERE TERM_SEQ = " & glbTERM_Seq
'    rsIN.Open strSQL, gdbAdoIhr001X, adOpenStatic, adLockOptimistic, adCmdText
'    If rsIN.EOF = False And rsIN.BOF = False Then
'       Do
'            rsOUT.AddNew
'                rsOUT("CO_COMPNO") = rsIN("CO_COMPNO")
'                rsOUT("CO_EMPNBR") = rsIN("CO_EMPNBR")
'                rsOUT("CO_EDATE") = rsIN("CO_EDATE")
'                rsOUT("CO_TYPE_TABL") = rsIN("CO_TYPE_TABL")
'                rsOUT("CO_TYPE") = rsIN("CO_TYPE")
'               ' rsOUT("CO_COMMENT_ID") = rsIN("CO_COMMENT_ID")
'                rsOUT("CO_COMMENTS") = rsIN("CO_COMMENTS")
'                rsOUT("CO_LDATE") = rsIN("CO_LDATE")
'                rsOUT("CO_LTIME") = rsIN("CO_LTIME")
'                rsOUT("CO_LUSER") = rsIN("CO_LUSER")
'                rsOUT("TERM_SEQ") = rsIN("TERM_SEQ")
'                rsOUT("WRKEMP") = glbUserID
'            rsOUT.Update
'            rsIN.MoveNext
'        Loop Until rsIN.EOF
'        rsIN.Close
'        rsOUT.Close
'
        c = 0
        retVal = ""
        If xTemplate = "" Or xTemplate = "TEMPLATE" Then
            strSQL = "SELECT CODENAME, USERID, ACCESSABLE, MAINTAINABLE FROM HR_SECURE_COMMENTS WHERE USERID='" & Replace(glbUserID, "'", "''") & "'"
        Else
            '????Ticket #24808 -  Retrieve template's security profile
            strSQL = "SELECT CODENAME, USERID, ACCESSABLE, MAINTAINABLE FROM HR_SECURE_COMMENTS WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
        End If
        strSQL = strSQL & " AND ACCESSABLE <> 0 "
        rsOUT.Open strSQL, gdbAdoIhr001, adOpenDynamic, adLockOptimistic, adCmdText
        If rsOUT.EOF = False And rsOUT.BOF = False Then
            Do
                retVal = retVal & "'" & rsOUT("CODENAME") & "', "
'                strSQL = "select ACCESSABLE, MAINTAINABLE from HR_SECURE_COMMENTS where "
'                strSQL = strSQL & "USERID='" & glbUserID & "' AND CODENAME='" & rsOUT("CO_TYPE") & "' AND TB_NAME='ECOM'"
'                rsIN.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
'                If rsIN.EOF = False And rsIN.BOF = False Then
'                    rsOUT("SC_MAINTAINABLE") = rsIN("MAINTAINABLE")
'                    rsOUT("SC_ACCESSABLE") = rsIN("ACCESSABLE")
'                    rsOUT.Update
'                End If
'                rsIN.Close
                c = c + 1
                rsOUT.MoveNext
            Loop Until rsOUT.EOF
        End If
        If c = 0 Then
            retVal = "=''"
        ElseIf c = 1 Then
            retVal = "=" & Left(retVal, Len(retVal) - 2)
        ElseIf c > 1 Then
            retVal = "IN (" & Left(retVal, Len(retVal) - 2) & ")"
        End If
        rsOUT.Close
        
         buildSec = retVal
 '   End If
    
    
End Function

Public Function buildSec_FollowUp() As String
    'working table not good for Access, changed from a working table to just a list of acceptable codes
    Dim rsOUT As New ADODB.Recordset
    Dim strSQL As String
    Dim retVal As String
    Dim c As Long
    Dim xTemplate As String
    
    '????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
    xTemplate = ""
    xTemplate = Get_Template(glbUserID)
        
        
    c = 0
    retVal = ""
    If xTemplate = "" Or xTemplate = "TEMPLATE" Then
        strSQL = "SELECT CODENAME, USERID, ACCESSABLE, MAINTAINABLE FROM HR_SECURE_FOLLOW_UP WHERE USERID='" & Replace(glbUserID, "'", "''") & "'"
    Else
        '????Ticket #24808 -  Retrieve template's security profile
        strSQL = "SELECT CODENAME, USERID, ACCESSABLE, MAINTAINABLE FROM HR_SECURE_FOLLOW_UP WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
    End If
    strSQL = strSQL & " AND ACCESSABLE <> 0 "
    rsOUT.Open strSQL, gdbAdoIhr001, adOpenDynamic, adLockOptimistic, adCmdText
    If rsOUT.EOF = False And rsOUT.BOF = False Then
        Do
           retVal = retVal & "'" & rsOUT("CODENAME") & "', "
           c = c + 1
           rsOUT.MoveNext
       Loop Until rsOUT.EOF
    End If
    If c = 0 Then
       retVal = "=''"
    ElseIf c = 1 Then
       retVal = "=" & Left(retVal, Len(retVal) - 2)
    ElseIf c > 1 Then
       retVal = "IN (" & Left(retVal, Len(retVal) - 2) & ")"
    End If
    rsOUT.Close
    
    buildSec_FollowUp = retVal
    
End Function

Public Function lockctl(ByRef frm As Form, dolock As Boolean)
    On Error GoTo Eh
    Dim ctl As Control
    
    For Each ctl In frm.Controls
         If TypeOf ctl Is TextBox Then
            ctl.Locked = Not dolock
        ElseIf TypeOf ctl Is ListBox Then
            ctl.Enabled = dolock
        ElseIf TypeOf ctl Is ComboBox Then
            ctl.Enabled = dolock
'        ElseIf TypeOf ctl Is CommandButton And ctl.Tag = "L" Then
'            ctl.Enabled = notdolock
        ElseIf TypeOf ctl Is CheckBox Then
            ctl.Locked = Not dolock
        ElseIf TypeOf ctl Is CodeLookup Then
            ctl.Enabled = dolock
        ElseIf TypeOf ctl Is DateLookup Then
            ctl.Enabled = dolock
        ElseIf TypeOf ctl Is EmployeeLookup Then
            ctl.Enabled = dolock
        End If
    Next ctl
    
exH:
    Exit Function
Eh:
    glbFrmCaption$ = frm.Caption
    glbErrNum& = Err

    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "lockctl", "Security", "Security")
    Call RollBack '28July99 js
End Function
Public Function GetUserEmpNo(xUser)
Dim rsUser As New ADODB.Recordset
Dim retVal
    retVal = ""
    If Len(xUser) > 0 Then
        rsUser.Open "SELECT USERID,USERNAME,EMPNBR FROM HR_SECURE_BASIC WHERE USERID='" & Replace(xUser, "'", "''") & "' ", gdbAdoIhr001, adOpenStatic
        If Not rsUser.EOF Then
            If Not IsNull(rsUser("EMPNBR")) Then
                retVal = rsUser("EMPNBR")
            End If
        End If
        rsUser.Close
    End If
    GetUserEmpNo = retVal
End Function
Public Function GetUserDesc(xUser)
Dim rsUser As New ADODB.Recordset
Dim xDesc
    If Len(xUser) = 0 Then
        xDesc = ""
    Else
        rsUser.Open "SELECT USERID,USERNAME FROM HR_SECURE_BASIC WHERE USERID='" & Replace(xUser, "'", "''") & "' ", gdbAdoIhr001, adOpenStatic
        If rsUser.EOF Then
            xDesc = xUser
        Else
            xDesc = rsUser("USERNAME")
        End If
        rsUser.Close
    End If
    GetUserDesc = xDesc
End Function


Public Function Get_No_Access_Group_List()
    Dim rsDeptSec As New ADODB.Recordset
    Dim SQLQ As String
            
    'Get list of Union Codes the user does not have access to
    SQLQ = "Select HRPASDEP.* from HRPASDEP"
    SQLQ = SQLQ & " WHERE HRPASDEP.PD_USERID = '" & Replace(glbUserID, "'", "''") & "'"
    SQLQ = SQLQ & " AND LEFT(PD_ORG,1) = '-'"
    SQLQ = SQLQ & " ORDER by PD_DEPT, PD_ORG, PD_DIV, PD_SECTION, PD_ADMINBY, PD_LOC, PD_REGION "
    rsDeptSec.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    glbNoAccessGrp = ""
    Do While Not rsDeptSec.EOF
        
        glbNoAccessGrp = glbNoAccessGrp & IIf(Len(glbNoAccessGrp) > 0, " AND NOT (", "NOT (")
        
        If Not IsNull(rsDeptSec("PD_DEPT")) And rsDeptSec("PD_DEPT") <> "ALL" And rsDeptSec("PD_DEPT") <> "" Then
            glbNoAccessGrp = glbNoAccessGrp & "ED_DEPTNO = '" & rsDeptSec("PD_DEPT") & "'"
        End If
        If Not IsNull(rsDeptSec("PD_ORG")) And rsDeptSec("PD_ORG") <> "" Then
            glbNoAccessGrp = glbNoAccessGrp & IIf(Len(glbNoAccessGrp) > 0 And Right(glbNoAccessGrp, 1) = "(", "", " AND ") & "ED_ORG = '" & Mid(rsDeptSec("PD_ORG"), 2) & "'"
        End If
        If Not IsNull(rsDeptSec("PD_DIV")) And rsDeptSec("PD_DIV") <> "" Then
            glbNoAccessGrp = glbNoAccessGrp & IIf(Len(glbNoAccessGrp) > 0 And Right(glbNoAccessGrp, 1) = "(", "", " AND ") & "ED_DIV = '" & rsDeptSec("PD_DIV") & "'"
        End If
        'Ticket #18235
        If Not IsNull(rsDeptSec("PD_ADMINBY")) And rsDeptSec("PD_ADMINBY") <> "" Then
            glbNoAccessGrp = glbNoAccessGrp & IIf(Len(glbNoAccessGrp) > 0 And Right(glbNoAccessGrp, 1) = "(", "", " AND ") & "ED_ADMINBY = '" & rsDeptSec("PD_ADMINBY") & "'"
        End If
        
        'Ticket #22682 Release 8.0
        If Not IsNull(rsDeptSec("PD_LOC")) And rsDeptSec("PD_LOC") <> "" Then
            glbNoAccessGrp = glbNoAccessGrp & IIf(Len(glbNoAccessGrp) > 0 And Right(glbNoAccessGrp, 1) = "(", "", " AND ") & "ED_LOC = '" & rsDeptSec("PD_LOC") & "'"
        End If
        If Not IsNull(rsDeptSec("PD_REGION")) And rsDeptSec("PD_REGION") <> "" Then
            glbNoAccessGrp = glbNoAccessGrp & IIf(Len(glbNoAccessGrp) > 0 And Right(glbNoAccessGrp, 1) = "(", "", " AND ") & "ED_REGION = '" & rsDeptSec("PD_REGION") & "'"
        End If
        
        If Not IsNull(rsDeptSec("PD_SECTION")) And rsDeptSec("PD_SECTION") <> "" Then
            glbNoAccessGrp = glbNoAccessGrp & IIf(Len(glbNoAccessGrp) > 0 And Right(glbNoAccessGrp, 1) = "(", "", " AND ") & "ED_SECTION = '" & rsDeptSec("PD_SECTION") & "'"
        End If
    
        If glbNoAccessGrp = "NOT (" Then
            glbNoAccessGrp = ""
        ElseIf Right(glbNoAccessGrp, 10) = " AND NOT (" Then
            glbNoAccessGrp = Left(glbNoAccessGrp, Len(glbNoAccessGrp) - 10)
        Else
            glbNoAccessGrp = glbNoAccessGrp & IIf(Len(glbNoAccessGrp) > 0, ")", "")
        End If
        
        'If Not IsNull(rsDeptSec("PD_ORG")) Then
        '    If Left(rsDeptSec("PD_ORG"), 1) = "-" Then
        '        glbNoAccessGrp = glbNoAccessGrp & IIf(Len(glbNoAccessGrp) > 0, ",", "") & "'" & Mid(rsDeptSec("PD_ORG"), 2) & "'"
        '    End If
        'End If
        
        rsDeptSec.MoveNext
    Loop
    rsDeptSec.Close
End Function

Public Function Allow_User_To_View(strActTerm As String) As Boolean
    Dim rsHREmp As New ADODB.Recordset
    Dim SQLQ As String
    
    If Len(glbNoAccessGrp) <> 0 Then
        Allow_User_To_View = False
        
        If strActTerm = "ACTIVE" Then
            SQLQ = "SELECT ED_EMPNBR FROM HREMP "
        Else
            SQLQ = "SELECT ED_EMPNBR FROM Term_HREMP "
        End If
        
        SQLQ = SQLQ & " WHERE " & glbNoAccessGrp
        SQLQ = SQLQ & " AND ED_EMPNBR = " & glbLEE_ID
        rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        
        If Not rsHREmp.EOF Then
            Allow_User_To_View = True
        Else
            Allow_User_To_View = False
        End If
        rsHREmp.Close
    Else
        Allow_User_To_View = True
    End If
    
End Function

Public Function Get_Division_Name(xDivCode As String, Optional xField) As String
    Dim rsDiv As New ADODB.Recordset
    Dim SQLQ As String
    
    'SQLQ = "SELECT Division_Name FROM HR_DIVISION WHERE DIV = '" & xDivCode & "' "
    SQLQ = "SELECT * FROM HR_DIVISION WHERE DIV = '" & xDivCode & "' "
    rsDiv.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsDiv.EOF Then
        If Not IsMissing(xField) Then
            Get_Division_Name = IIf(IsNull(rsDiv(xField)), "", rsDiv(xField))
        Else
            Get_Division_Name = rsDiv("Division_Name")
        End If
    Else
        Get_Division_Name = ""
    End If
    rsDiv.Close
    Set rsDiv = Nothing
End Function

Public Function Get_Province_Name(xProvinceCode As String) As String
    Dim rsProv As New ADODB.Recordset
    Dim SQLQ As String
    
    SQLQ = "SELECT DESCR FROM HRPROV WHERE CODE = '" & xProvinceCode & "' "
    rsProv.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsProv.EOF Then
        Get_Province_Name = rsProv("DESCR")
    Else
        Get_Province_Name = ""
    End If
    rsProv.Close
    Set rsProv = Nothing
End Function

Public Function Get_ProvinceCodeData(xProvinceCode As String, xField As String) As String
    Dim rsProv As New ADODB.Recordset
    Dim SQLQ As String
    
    SQLQ = "SELECT " & xField & " FROM HRPROV WHERE CODE = '" & xProvinceCode & "' "
    rsProv.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsProv.EOF Then
        If Not IsNull(rsProv(xField)) Then
            Get_ProvinceCodeData = rsProv(xField)
        Else
            Get_ProvinceCodeData = ""
        End If
    Else
        Get_ProvinceCodeData = ""
    End If
    rsProv.Close
    Set rsProv = Nothing
End Function

Public Function Get_ProvinceNoData(xProvinceNo As String, xField As String) As String
    Dim rsProv As New ADODB.Recordset
    Dim SQLQ As String
    
    SQLQ = "SELECT " & xField & "  FROM HRPROV WHERE NBR = '" & xProvinceNo & "' "
    rsProv.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsProv.EOF Then
        If Not IsNull(rsProv(xField)) Then
            Get_ProvinceNoData = rsProv(xField)
        Else
            Get_ProvinceNoData = ""
        End If
    Else
        Get_ProvinceNoData = ""
    End If
    rsProv.Close
    Set rsProv = Nothing
End Function

'Provided a French Month returns an english month
Public Function GetEnglishMonth(month As String) As String
    Dim frenchMonths, englishMonths, I
    frenchMonths = Array("janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre")
    englishMonths = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    Dim rtnval As String
    rtnval = ""
    
        
        For I = 0 To UBound(englishMonths)
            If UCase(Left(month, 3)) = UCase(Left(frenchMonths(I), 3)) Then
                If Len(month) = 3 Then
                    rtnval = Left(englishMonths(I), 3)
                ElseIf Len(month) > 3 Then
                    rtnval = englishMonths(I)
                End If
                Exit For
            End If
        Next
        
    GetEnglishMonth = rtnval
End Function

Public Function TranslateDateString(dStr As String) As String
    Dim frenchMonths, englishMonths, I
    frenchMonths = Array("janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre")
    englishMonths = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    Dim rtnval As String
    rtnval = ""
    
        
        For I = 0 To UBound(englishMonths)
            If InStr(1, UCase(dStr), UCase(frenchMonths(I))) > 0 Then
                 rtnval = Replace(dStr, frenchMonths(I), englishMonths(I))
                Exit For
            End If
        Next
        
    TranslateDateString = rtnval
End Function

'Returns French only if the system is french
Public Function GetMonth(month As String) As String
    Dim frenchMonths, englishMonths, I
    frenchMonths = Array("janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre")
    englishMonths = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    Dim rtnval As String
    rtnval = ""
    If glbFrench = True Then
        
        For I = 0 To UBound(englishMonths)
            If UCase(Left(month, 3)) = UCase(Left(englishMonths(I), 3)) Then
                If Len(month) = 3 Then
                    rtnval = Left(frenchMonths(I), 3)
                ElseIf Len(month) > 3 Then
                    rtnval = frenchMonths(I)
                End If
                Exit For
            End If
        Next
        
    Else
        rtnval = month
    End If
    
    GetMonth = rtnval
End Function

Public Sub WriteFile(Log)
 Dim sFileText As String
 Dim iFileNo As Integer
 iFileNo = FreeFile
 'open the file for writing
 Open "c:\log.txt" For Append As #iFileNo
 'please note, if this file already exists it will be overwritten!
 'write some example text to the file
 Print #iFileNo, Log
 
 'close the file (if you dont do this, you wont be able to open it again!)
 Close #iFileNo
 'To Read    If fs.FileExists("C:\mytestfile.txt") Then        Set ts = fs.OpenTextFile("C:\mytestfile.txt")                 Do While Not ts.AtEndOfStream            MsgBox ts.ReadLine        Loop        ts.Close    End If        'clear memory used by FSO objects    Set ts = Nothing    Set fs = Nothing
End Sub

Public Function getCodeFromDesc(TblName As String, TblDesc As String) As String
    Dim SQLQ As String
    Dim rsTABL As New ADODB.Recordset
    On Error GoTo CodeDesc_Err
        
    SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = '" & TblName & "' AND TB_DESC = '" & TblDesc & "' "
    rsTABL.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTABL.EOF Then
        getCodeFromDesc = rsTABL("TB_KEY")
    Else
        getCodeFromDesc = ""
    End If
    rsTABL.Close
    Exit Function

CodeDesc_Err:
    Resume Next
End Function

Public Function GetEmployeeInfo(info As String) As String
    Dim SQLQ As String
    Dim rsE As New ADODB.Recordset
    On Error GoTo Emp_Err
    
    SQLQ = "SELECT " & info & " FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID
    rsE.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsE.EOF Then
        GetEmployeeInfo = rsE(info)
    Else
        GetEmployeeInfo = ""
    End If
    rsE.Close
    Exit Function

Emp_Err:
    Resume Next
End Function

'Ticket #18188
Public Function GetEmployeeInfoByColumn(col As String, info As String) As String
    Dim SQLQ As String
    Dim rsE As New ADODB.Recordset
    On Error GoTo Emp_Err
    
    
    
    SQLQ = "SELECT " & col & "  FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID
    rsE.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsE.EOF Then
       
            GetEmployeeInfoByColumn = rsE(info)
            
    Else
        GetEmployeeInfoByColumn = ""
    End If
    rsE.Close
    Exit Function

Emp_Err:
    Resume Next
End Function

'Ticket #18188
Public Function IsSTATHoliday(d1 As Date, d2 As Date, Optional xEmpnbr) As String
    Dim SQLQ As String
    Dim rsHoliday As New ADODB.Recordset
    Dim wProv As String
    Dim wSection As String
    
    On Error GoTo Holiday_Err
    
    SQLQ = "SELECT * FROM HR_HOLIDAY WHERE HL_DATE >= " & Date_SQL(d1) & " AND HL_DATE <= " & Date_SQL(d2) & " "
    'If glbCompSerial = "S/N - 2418W" Then
        'Province/State
        If IsMissing(xEmpnbr) Then
            wProv = GetEmployeeInfo("ED_PROVEMP")
        Else
            wProv = GetEmpData(xEmpnbr, "ED_PROVEMP")
        End If
        If wProv <> "" Then
            'Ticket #20014 Franks 03/18/2011
            'SQLQ = SQLQ & " AND (HL_STATE = '" & wProv & "' OR HL_STATE is null OR  len(HL_STATE) = 0) "
            SQLQ = SQLQ & " AND (HL_STATE = '" & wProv & "' OR HL_STATE is null OR " & lenFunction & "(HL_STATE) = 0) "
        Else
            If IsMissing(xEmpnbr) Then
                wProv = GetEmployeeInfo("ED_PROV")
            Else
                wProv = GetEmpData(xEmpnbr, "ED_PROV")
            End If
            If wProv <> "" Then
                SQLQ = SQLQ & " AND (HL_STATE = '" & wProv & "' OR HL_STATE is null OR " & lenFunction & "(HL_STATE) = 0) "
            End If
        End If
        
        'Section
        If IsMissing(xEmpnbr) Then
            wSection = GetEmployeeInfo("ED_SECTION")
        Else
            wSection = GetEmpData(xEmpnbr, "ED_SECTION")
        End If
        If wSection <> "" Then
            SQLQ = SQLQ & " AND (HL_SECTION = '" & wSection & "' OR HL_SECTION is null OR  " & lenFunction & "(HL_SECTION) = 0) "
        End If
    'End If
    rsHoliday.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsHoliday.EOF Then
        Do While Not rsHoliday.EOF
            IsSTATHoliday = IsSTATHoliday & Date_SQL(rsHoliday("HL_DATE")) & "|"
            rsHoliday.MoveNext
        Loop
    Else
        IsSTATHoliday = ""
    End If
    rsHoliday.Close
    Exit Function

Holiday_Err:
    Resume Next
End Function

Public Function getHREEO_NOC(xEmpNo, Optional xJob)
Dim rsEmpNOC As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As String
    retVal = ""
    'Ticket #20852
    If IsMissing(xJob) Then
        SQLQ = "SELECT JB_FEDGRP FROM HRJOB WHERE JB_CODE IN (SELECT JH_JOB FROM HR_JOB_HISTORY WHERE JH_EMPNBR = " & xEmpNo & " AND JH_CURRENT <> 0)"
    Else
        SQLQ = "SELECT JB_FEDGRP FROM HRJOB WHERE JB_CODE = '" & xJob & "'"
    End If
    rsEmpNOC.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmpNOC.EOF Then
        If Not IsNull(rsEmpNOC("JB_FEDGRP")) Then
            retVal = rsEmpNOC("JB_FEDGRP")
        End If
    End If
    getHREEO_NOC = retVal
End Function

Public Sub uptEEO_Fields(xEmpNo, xType, Optional xFieldName, Optional xValue, Optional xJob, Optional xETHNICITY, Optional xRACE) ', xDOT)
Dim rsEmp As New ADODB.Recordset
Dim rsEEO As New ADODB.Recordset
Dim SQLQ As String
Dim xStr As String
Dim xlocEmpNo
    
    If IsMissing(xFieldName) Then
        If xType = "New" Then
            If Len(xEmpNo) > 0 Then
                SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
                rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsEmp.EOF Then
                    SQLQ = "SELECT * FROM HREEO WHERE EO_TYPE = 'E' AND NOT (EO_EMPNBR IS NULL) "
                    SQLQ = SQLQ & "AND EO_EMPNBR = " & xEmpNo & " "
                    rsEEO.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If rsEEO.EOF Then
                        rsEEO.AddNew
                        rsEEO("EO_TYPE") = "E"
                        rsEEO("EO_EMPNBR") = xEmpNo
                        rsEEO("EO_EEONNBR") = xEmpNo
                        rsEEO("EO_VETERAN") = 0
                        rsEEO("EO_VIETNAM") = 0
                        rsEEO("EO_DISABLE_YN") = 0
                    End If
                    rsEEO("EO_Surname") = rsEmp("ED_Surname")
                    rsEEO("EO_FName") = rsEmp("ED_FName")
                    rsEEO("EO_SSN") = rsEmp("ED_SIN")
                    rsEEO("EO_DOB") = rsEmp("ED_DOB")
                    rsEEO("EO_SEX") = rsEmp("ED_SEX")
                    rsEEO("EO_DOH") = rsEmp("ED_DOH")
                    rsEEO("EO_WORKCOUNTRY") = rsEmp("ED_WORKCOUNTRY")
                    rsEEO("EO_PT") = rsEmp("ED_PT")
                    rsEEO("EO_REGION") = rsEmp("ED_REGION")
                    rsEEO("EO_LOC") = rsEmp("ED_LOC")
                    
                    ''Ticket #20852
                    'If glbMulti Then
                        If Not IsMissing(xJob) Then
                            xStr = getHREEO_NOC(xEmpNo, xJob)
                        Else
                            xStr = getHREEO_NOC(xEmpNo)
                        End If
                    'Else
                    '    xStr = getHREEO_NOC(xEmpNo)
                    'End If
                    If Len(xStr) = 0 Then
                        rsEEO("EO_OCC_CAT") = Null
                    Else
                        rsEEO("EO_OCC_CAT") = xStr
                    End If
                    'Ticket #24767 Franks 12/10/2013 - begin
                    If Not IsMissing(xETHNICITY) Then
                        If Len(xETHNICITY) > 0 Then
                            rsEEO("EO_ETHNICITY") = xETHNICITY
                        End If
                    End If
                    If Not IsMissing(xRACE) Then
                        If Len(xRACE) > 0 Then
                            rsEEO("EO_RACE") = xRACE
                        End If
                    End If
                    'Ticket #24767 Franks 12/10/2013 - end
                    rsEEO.Update
                End If
            End If
        End If
        
        If xType = "Update" Then
            'update all fields
            SQLQ = "SELECT * FROM HREEO WHERE EO_TYPE = 'E' AND NOT (EO_EMPNBR IS NULL) "
            If Len(xEmpNo) > 0 Then
                SQLQ = SQLQ & "AND EO_EMPNBR = " & xEmpNo & " "
            End If
            rsEEO.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            
            'Ticket #21525
            Do While Not rsEEO.EOF
            'If Not rsEEO.EOF Then
                xlocEmpNo = rsEEO!EO_EMPNBR
                
                'Employee info - begin
                SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xlocEmpNo
                rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsEmp.EOF Then
                    rsEEO("EO_Surname") = rsEmp("ED_Surname")
                    rsEEO("EO_FName") = rsEmp("ED_FName")
                    rsEEO("EO_SSN") = rsEmp("ED_SIN")
                    rsEEO("EO_DOB") = rsEmp("ED_DOB")
                    rsEEO("EO_SEX") = rsEmp("ED_SEX")
                    rsEEO("EO_DOH") = rsEmp("ED_DOH")
                    rsEEO("EO_WORKCOUNTRY") = rsEmp("ED_WORKCOUNTRY")
                    rsEEO("EO_PT") = rsEmp("ED_PT")
                    rsEEO("EO_REGION") = rsEmp("ED_REGION")
                    rsEEO("EO_LOC") = rsEmp("ED_LOC")
                End If
                rsEmp.Close
                Set rsEmp = Nothing
                'Employee info - end
                    
                'Ticket #20852
                If glbMulti Then
                    If Not IsMissing(xJob) Then
                        xStr = getHREEO_NOC(xlocEmpNo, xJob)
                    Else
                        xStr = getHREEO_NOC(xlocEmpNo)
                    End If
                Else
                    xStr = getHREEO_NOC(xlocEmpNo)
                End If
                If Len(xStr) = 0 Then
                    rsEEO("EO_OCC_CAT") = Null
                Else
                    rsEEO("EO_OCC_CAT") = xStr
                End If
                rsEEO.Update
            'End If
                rsEEO.MoveNext
            Loop
            rsEEO.Close
            Set rsEEO = Nothing
        End If
        
        If xType = "Delete" Then
            SQLQ = "DELETE FROM HREEO WHERE (1=1) "
            SQLQ = SQLQ & "AND EO_EMPNBR = " & xEmpNo & " "
            gdbAdoIhr001.Execute SQLQ
        End If
    Else
        'update this field
        If xFieldName = "EO_OCC_CAT" Then
            'Ticket #20852
            If glbMulti Then
                If Not IsMissing(xJob) Then
                    xStr = getHREEO_NOC(xEmpNo, xJob)
                Else
                    xStr = getHREEO_NOC(xEmpNo)
                End If
            Else
                xValue = getHREEO_NOC(xEmpNo)
            End If
        End If
        If xValue = "" Then xValue = Null
        SQLQ = "SELECT * FROM HREEO WHERE EO_EMPNBR = " & xEmpNo & " "
        rsEEO.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsEEO.EOF Then
            rsEEO(xFieldName) = xValue
            rsEEO.Update
        End If
        rsEEO.Close
    End If
End Sub

Public Function GPBDPayCode(xUnion)
Dim rsGPICMatrix As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As Boolean
    retVal = False
    SQLQ = "SELECT * FROM HR_GP_INCOMECODE_MATRIX WHERE IC_HRCODE = '" & xUnion & "' "
    'for all IC_FUNC_GROUP
    '2-Regular Rate;3-Benefit/Deduction
    'SQLQ = SQLQ & "AND (IC_FUNC_GROUP = '2' OR IC_FUNC_GROUP = '3') "
    rsGPICMatrix.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsGPICMatrix.EOF Then
        retVal = True
    End If
    rsGPICMatrix.Close
    GPBDPayCode = retVal
End Function

Public Sub UpdateGPBenefitDeduction(xEmpNo, xNewUnion, xOldUnion)
Dim rsICMST As New ADODB.Recordset
Dim rsICTMP As New ADODB.Recordset
Dim rsBGEE As New ADODB.Recordset
Dim rsTABL As New ADODB.Recordset
Dim SQLQ As String
Dim BelongOldGroup As Boolean

gdbAdoIhr001.Execute "DELETE FROM HR_GP_INCOMECODE_MTR_WRK WHERE IC_WRKEMP = '" & glbUserID & "' "

If Len(xNewUnion) > 0 Then
    gdbAdoIhr001.BeginTrans
    SQLQ = "SELECT * FROM HR_GP_INCOMECODE_MTR_WRK WHERE IC_WRKEMP = '" & glbUserID & "' "
    rsICTMP.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

    SQLQ = "SELECT * FROM HR_GP_INCOMECODE_MATRIX WHERE IC_HRCODE = '" & xNewUnion & "' "
    'SQLQ = SQLQ & "AND IC_FUNC_GROUP = '3' "
    'Ticket #19782 Franks 02/02/2011, it should for all type
    'SQLQ = SQLQ & "AND (IC_FUNC_GROUP = '2' OR IC_FUNC_GROUP = '3') "
    rsICMST.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    Do While Not rsICMST.EOF
        rsICTMP.AddNew
        rsICTMP("IC_COMPNO") = "001"
        rsICTMP("IC_INCOMECODE") = rsICMST("IC_INCOMECODE")
        rsICTMP("IC_INCOMECODEDESC") = rsICMST("IC_INCOMECODEDESC")
        rsICTMP("IC_CODETYPE") = rsICMST("IC_CODETYPE")
        rsICTMP("IC_HRCODE") = rsICMST("IC_HRCODE")
        rsICTMP("IC_FUNC_GROUP") = rsICMST("IC_FUNC_GROUP")
        If Not IsNull(rsICMST("IC_CODETYPE")) Then
            rsICTMP("IC_CODETYPE_DESC") = getGPTYPE_Desc(rsICMST("IC_CODETYPE"))
        End If
        rsICTMP("IC_MULTIPLIER") = rsICMST("IC_MULTIPLIER")
        rsICTMP("IC_COMMON") = rsICMST("IC_COMMON")
        rsICTMP("IC_EMPNBR") = xEmpNo
        rsICTMP("IC_OMERS_TIER") = rsICMST("IC_OMERS_TIER")
        rsICTMP("IC_RULE_DESC") = rsICMST("IC_RULE_DESC")
        'Ticket #19179
        If rsICMST("IC_DEF_ACTION") Then
            rsICTMP("IC_CHECK") = 1
        Else
            rsICTMP("IC_CHECK") = 0
        End If
        rsICTMP("IC_ACTION") = "Add"
        rsICTMP("IC_WRKEMP") = glbUserID
        rsICTMP.Update
        rsICMST.MoveNext
    Loop
    rsICTMP.Close
    rsICMST.Close
    gdbAdoIhr001.CommitTrans
    
    If Not glbSQL And Not glbOracle Then Call Pause(1)
    
    'Old Union
    'SQLQ = "SELECT * FROM HR_GP_INCOMECODE_MTR_WRK WHERE IC_WRKEMP = '" & glbUserID & "' "
    'rsICTMP.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    SQLQ = "SELECT * FROM HR_GP_INCOMECODE_MATRIX WHERE IC_HRCODE = '" & xOldUnion & "' "
    'SQLQ = SQLQ & "AND IC_FUNC_GROUP = '3' "
    SQLQ = SQLQ & "AND (IC_FUNC_GROUP = '2' OR IC_FUNC_GROUP = '3') "
    rsICMST.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    Do While Not rsICMST.EOF
        SQLQ = "SELECT * FROM HR_GP_INCOMECODE_MTR_WRK WHERE IC_WRKEMP = '" & glbUserID & "' "
        SQLQ = SQLQ & "AND IC_INCOMECODE = '" & rsICMST("IC_INCOMECODE") & "' "
        SQLQ = SQLQ & "AND IC_ACTION = 'Add'"
        rsICTMP.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsICTMP.EOF Then
            'not found, to be deleted
            rsICTMP.AddNew
            rsICTMP("IC_COMPNO") = "001"
            rsICTMP("IC_INCOMECODE") = rsICMST("IC_INCOMECODE")
            rsICTMP("IC_INCOMECODEDESC") = rsICMST("IC_INCOMECODEDESC")
            rsICTMP("IC_CODETYPE") = rsICMST("IC_CODETYPE")
            rsICTMP("IC_HRCODE") = rsICMST("IC_HRCODE")
            If Not IsNull(rsICMST("IC_CODETYPE")) Then
                rsICTMP("IC_CODETYPE_DESC") = getGPTYPE_Desc(rsICMST("IC_CODETYPE"))
            End If
            rsICTMP("IC_FUNC_GROUP") = rsICMST("IC_FUNC_GROUP")
            rsICTMP("IC_MULTIPLIER") = rsICMST("IC_MULTIPLIER")
            rsICTMP("IC_COMMON") = rsICMST("IC_COMMON")
            rsICTMP("IC_EMPNBR") = xEmpNo
            rsICTMP("IC_CHECK") = 1
            rsICTMP("IC_ACTION") = "Delete"
            rsICTMP("IC_WRKEMP") = glbUserID
            rsICTMP.Update
        Else
            'found, to be updated
            rsICTMP("IC_ACTION") = "Update"
            rsICTMP.Update
        End If
        rsICTMP.Close

        rsICMST.MoveNext
    Loop
    'rsICTMP.Close
    rsICMST.Close
Else
    'Deleting the Pay Codes for this Union Group
    If Len(xOldUnion) > 0 Then
        gdbAdoIhr001.BeginTrans
        SQLQ = "SELECT * FROM HR_GP_INCOMECODE_MTR_WRK WHERE IC_WRKEMP = '" & glbUserID & "' "
        rsICTMP.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
        SQLQ = "SELECT * FROM HR_GP_INCOMECODE_MATRIX WHERE IC_HRCODE = '" & xOldUnion & "' "
        'SQLQ = SQLQ & "AND IC_FUNC_GROUP = '3' "
        'SQLQ = SQLQ & "AND (IC_FUNC_GROUP = '2' OR IC_FUNC_GROUP = '3') "
        rsICMST.Open SQLQ, gdbAdoIhr001, adOpenStatic
        
        Do While Not rsICMST.EOF
            rsICTMP.AddNew
            rsICTMP("IC_COMPNO") = "001"
            rsICTMP("IC_INCOMECODE") = rsICMST("IC_INCOMECODE")
            rsICTMP("IC_INCOMECODEDESC") = rsICMST("IC_INCOMECODEDESC")
            rsICTMP("IC_CODETYPE") = rsICMST("IC_CODETYPE")
            rsICTMP("IC_HRCODE") = rsICMST("IC_HRCODE")
            rsICTMP("IC_FUNC_GROUP") = rsICMST("IC_FUNC_GROUP")
            If Not IsNull(rsICMST("IC_CODETYPE")) Then
                rsICTMP("IC_CODETYPE_DESC") = getGPTYPE_Desc(rsICMST("IC_CODETYPE"))
            End If
            rsICTMP("IC_MULTIPLIER") = rsICMST("IC_MULTIPLIER")
            rsICTMP("IC_COMMON") = rsICMST("IC_COMMON")
            rsICTMP("IC_EMPNBR") = xEmpNo
            rsICTMP("IC_OMERS_TIER") = rsICMST("IC_OMERS_TIER")
            rsICTMP("IC_RULE_DESC") = rsICMST("IC_RULE_DESC")
            rsICTMP("IC_CHECK") = 1
            rsICTMP("IC_ACTION") = "Delete"
            rsICTMP("IC_WRKEMP") = glbUserID
            rsICTMP.Update
            rsICMST.MoveNext
        Loop
        rsICTMP.Close
        rsICMST.Close
        gdbAdoIhr001.CommitTrans
    End If
End If

End Sub

Public Function getGPTYPE_Desc(xType)
Dim retVal As String
        retVal = ""
        'If xType = "1" Then
        '    retVal = "1 - Banked"
        'End If
        'If xType = "2" Then
        '    retVal = "2 - Benefit"
        'End If
        'If xType = "3" Then
        '    retVal = "3 - Deduction"
        'End If
        'If xType = "4" Then
        '    retVal = "4 - Income"
        'End If
        ''Ticket #19782 Franks 02/02/2011 for Frontenac
        If xType = "1" Then
            retVal = "1 - Income"
        End If
        If xType = "2" Then
            retVal = "2 - Benefit"
        End If
        If xType = "3" Then
            retVal = "3 - Deduction"
        End If
        If xType = "4" Then
            retVal = "4 - Banked"
        End If
        getGPTYPE_Desc = retVal
End Function

Public Function chkMaxDollar(xAmt) 'Ticket #19697
Dim retVal As Boolean
    retVal = False
    If IsNumeric(xAmt) Then
        If xAmt > glbMaxDollar Then
            retVal = True
        End If
    End If
    chkMaxDollar = retVal
End Function

Public Function lenFunction() 'Ticket #20014 Franks 03/18/2011
Dim retVal As String
    If glbOracle Then
        retVal = "Length"
    Else
        retVal = "Len"
    End If
    lenFunction = retVal
End Function

Sub modLoadHostINI()
Dim sPath As String, sSetting As String, strT  As String
Dim x%, Value$, I, DateStr
    
    'INIRead("Folders","ApplicationFolder","C:\ihrhost.ini")
    
    '===============================================================
    'Reading windows defaults - ie date
    'get the user's windows default date format
    lCurrentKey = HKEY_CURRENT_USER
    sPath = "CONTROL PANEL\INTERNATIONAL"
    ' dkostka - 07/12/2001 - Added check to make sure key is there before getting value.
    If DoesKeyExist(lCurrentKey, sPath) Then
        sSetting = "sShortDate"
        glbsDateFormat = "None Found"
        giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, glbsDateFormat)
        glbsDateFormat = UCase$(glbsDateFormat)
    Else
        glbsDateFormat = String(255, "0")
    End If
    
    'dkostka - 03/20/01 - Added support for French Windows environment
    If IsDate("2 Mars 2003") Then glbFrench = True
    
    'dkostka - 09/15/2000 - Changed for Linamar, Elgin, etc - If no date format in registry, try to find it manually.
    If glbsDateFormat = String(255, "0") Then
        If glbFrench Then DateStr = "Janvier 2, 2003" Else DateStr = "January 2, 2003"
        Select Case Format(DateStr, "Short Date")
            Case "01/02/2003", "1/02/2003", "01/2/2003", "1/2/2003"
                glbsDateFormat = "MM/DD/YYYY"
            Case "01/02/03", "1/02/03", "01/2/03", "1/2/03"
                glbsDateFormat = "MM/DD/YY"
            Case "02/01/2003", "2/01/2003", "02/1/2003", "2/1/2003"
                glbsDateFormat = "DD/MM/YYYY"
            Case "02/01/03", "2/01/03", "02/1/03", "2/1/03"
                glbsDateFormat = "DD/MM/YY"
            Case "2003/01/02", "2003/1/02", "2003/01/2", "2003/1/2"
                glbsDateFormat = "YYYY/MM/DD"
            Case "03/01/02", "03/1/02", "03/01/2", "03/1/2"
                glbsDateFormat = "YY/MM/DD"
            ' dkostka - 03/13/01 - Changed for Assumption Life - Added date formats with dashes
            Case "01-02-2003", "1-02-2003", "01-2-2003", "1-2-2003"
                glbsDateFormat = "DD-MM-YYYY"
            Case "01-02-03", "1-02-03", "01-2-03", "1-2-03"
                glbsDateFormat = "MM-DD-YY"
            Case "02-01-2003", "2-01-2003", "02-1-2003", "2-1-2003"
                glbsDateFormat = "DD-MM-YYYY"
            Case "02-01-03", "2-01-03", "02-1-03", "2-1-03"
                glbsDateFormat = "DD-MM-YY"
            Case "2003-01-02", "2003-1-02", "2003-01-2", "2003-1-2"
                glbsDateFormat = "YYYY-MM-DD"
            Case "03-01-02", "03-1-02", "03-01-2", "03-1-2"
                glbsDateFormat = "YY-MM-DD"
            Case Else
                glbsDateFormat = "Date Format Not Set"
        End Select
    End If
    'dkostka - 09/15/2000 - end

    Call iniDateFormat  'Jaddy - May 1,2001 for date entry
    
    'If Not DoesKeyExist(HKEY_LOCAL_MACHINE, REG_NAME) Then
    '    lCurrentKey = HKEY_CURRENT_USER
    'Else
    '    lCurrentKey = HKEY_LOCAL_MACHINE
    'End If
    
    'sPath = Section
    sPath = "INFOHR Files"
    '========================================
    'Jaddy Changed to remove all others registy keys except IHRDB
    'All database's Location will be set under IHRDB
    'IHRDB could have IHR001.MDB words or not
    'For example, F:\IHR; F:\IHR\ or F:\IHR\IHR001.MDB
    
    'sSetting = Key in the Section
    sSetting = "IHRDB"
    glbIHRDB = glbWorkDir & "\ihr001.mdb"
    glbIHRDB = INIRead(sPath, sSetting, glbHostFile)
    'giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, glbIHRDB)
    glbDBDir = Replace(UCase(glbIHRDB), "IHR001.MDB", "")
    glbDBDir = glbDBDir & IIf(Right(glbDBDir, 1) = "\", "", "\")
    If glbDBDir <> "" Then
        glbIHRDB = glbDBDir & "IHR001.MDB"
        glbIHRAUDIT = glbDBDir & "IHR001X.MDB"
        glbIHRDBW = glbDBDir & "IHR001W.MDB"
        glbIHRWFC = glbDBDir & "IHRWFC.MDB"
        glbIHRWFCA = glbDBDir & "IHRWFC-A.MDB"
        glbSN2322 = glbDBDir & "SN2322.MDB"
        glbIHRDBO = glbDBDir & "IHROPUS.MDB"
        glbIHRDBB = glbDBDir & "IHR001B.MDB"
        glbIHREDU = glbDBDir & "IHREDU.MDB"
    End If
    
    sSetting = "IHRREPORTS"  'Compressed database location
    glbIHRREPORTS = glbWorkDir
    'giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, glbIHRREPORTS)
    glbIHRREPORTS = INIRead(sPath, sSetting, glbHostFile)
    glbIHRREPORTS = glbIHRREPORTS & IIf(Right$(glbIHRREPORTS, 1) <> "\", "\", "")
    
    '==================================================
    sPath = "Network"
    sSetting = "EXCLUSIVE"  'Compressed database location
    strT = "N"
    'giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    strT = INIRead(sPath, sSetting, glbHostFile)
    If strT = "Y" Then
        glbExclusiveDB% = True
    Else
        glbExclusiveDB% = False
    End If
    
    
    sSetting = "MULTIUSERNUM"  'Compressed database location
    strT = "0"
    'giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    strT = INIRead(sPath, sSetting, glbHostFile)
    If IsNumeric(strT) Then glbMultiUserNum% = CInt(strT)
    

    '=========================================================
    sPath = "Options"
    sSetting = "DatabaseType"
    strT = ""
    'giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    strT = INIRead(sPath, sSetting, glbHostFile)
    gsSystemDb = UCase(strT)
    
    'Get Language
    sPath = "Options"
    sSetting = "MultiLang"
    strT = ""
    'giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    strT = INIRead(sPath, sSetting, glbHostFile)
    gsMultiLang = UCase(strT)
    If Len(gsMultiLang) > 250 Then
        gsMultiLang = "N"
    Else
        If gsMultiLang <> "Y" And gsMultiLang <> "YES" Then
            gsMultiLang = "N"
        End If
    End If
    'this changes were for Listowel. it can be used for every body
    
    
    sSetting = "ADDHISWARNING"  'Warning on when adding history
    strT = "N"
    'giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    strT = INIRead(sPath, sSetting, glbHostFile)
    If strT = "Y" Then
        glbAddHisWarning% = True
    Else
        glbAddHisWarning% = False
    End If
    
    sPath = "FOLLOWUPS"
    sSetting = "FOLLOWUPS"  'check for follow-ups when entering
    strT = "N"
    'giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    strT = INIRead(sPath, sSetting, glbHostFile)
    If strT = "Y" Then
        glbFOLLOWUPS% = True
    Else
        glbFOLLOWUPS% = False
    End If
    
    sSetting = "FOLLOWUPDAYS"  'check for follow-ups when entering
    strT = "5"
    'giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    strT = INIRead(sPath, sSetting, glbHostFile)
    If Not IsNumeric(strT) Then strT = "5"
    glbFOLLOWUPDAYS% = CInt(strT)
    
    
    sSetting = "SHOWCOMPLETED"  'check for follow-ups when entering
    strT = "N"
    'giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    strT = INIRead(sPath, sSetting, glbHostFile)
    If strT = "Y" Then
        glbFOLLOWUPSCOMP% = True
    Else
        glbFOLLOWUPSCOMP% = False
    End If
        
    '=========================================================
    
    sPath = "ODBC Setup"
    sSetting = "ODBCIHR"  'check for follow-ups when entering
    strT = "N"
    'giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    strT = INIRead(sPath, sSetting, glbHostFile)
    If strT = "N" Then
        'Call mod_ODBC_Register("IHR")
        'x% = WriteRegistrySetting(lCurrentKey, sPath, sSetting, "Y")
        x% = INIWrite(sPath, sSetting, "Y", glbHostFile)
    End If

    sSetting = "ODBCIHRX"  'check for follow-ups when entering
    strT = "N"
    'giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    strT = INIRead(sPath, sSetting, glbHostFile)
    If strT = "N" Then
        'Call mod_ODBC_Register("IHRX")
        'x% = WriteRegistrySetting(lCurrentKey, sPath, sSetting, "Y")
        x% = INIWrite(sPath, sSetting, "Y", glbHostFile)
    End If
            
    sSetting = "DATABASENAME"
    strT = ""
    'giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    strT = INIRead(sPath, sSetting, glbHostFile)
    ' dkostka - 06/17/2002 - Don't force the SQL Server params to uppercase, if the server is set like
    ' WHSCC's server, you have to have the case right or it won't let you log in.
    'SQLDatabaseName = UCase(strT)
    SQLDatabaseName = strT
    If Len(SQLDatabaseName) = 255 Then
        SQLDatabaseName = "INFOHR"
    End If
    
    'SQL SERVER LOGIN SERVER NAME, USER NAME AND PASSWORD
    sSetting = "SERVERNAME"
    strT = ""
    'giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    strT = INIRead(sPath, sSetting, glbHostFile)
    ' dkostka - 06/17/2002 - Don't force the SQL Server params to uppercase, if the server is set like
    ' WHSCC's server, you have to have the case right or it won't let you log in.
    'SQLServerName = UCase(strT)
    SQLServerName = strT
    If Len(SQLServerName) = 255 Then
        SQLServerName = ""
    End If
           
    sSetting = "USERNAME"
    strT = ""
    'giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    strT = INIRead(sPath, sSetting, glbHostFile)
    ' dkostka - 06/17/2002 - Don't force the SQL Server params to uppercase, if the server is set like
    ' WHSCC's server, you have to have the case right or it won't let you log in.
    'SQLUserName = UCase(strT)
    SQLUserName = strT
    If Len(SQLUserName) = 255 Then
        SQLUserName = ""
    End If
    
    sSetting = "USERPSW"
    strT = ""
    'giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    strT = INIRead(sPath, sSetting, glbHostFile)
    If gsMultiLang = "Y" Then
        SQLUserPassword = DecryptPasswordMultiLang_First(strT)
    ElseIf gsMultiLang = "YES" Then 'WHSCC
        SQLUserPassword = DecryptPasswordMultiLang(strT)
    Else
        glbDBPassFlag = True
        SQLUserPassword = DecryptPassword(strT)
    End If
    
    'Oracle driver NAME, USER NAME AND PASSWORD
    sSetting = "DRIVERNAME"
    strT = ""
    'giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    strT = INIRead(sPath, sSetting, glbHostFile)
    ' dkostka - 06/17/2002 - Don't force the SQL Server params to uppercase, if the server is set like
    ' WHSCC's server, you have to have the case right or it won't let you log in.
    'SQLServerName = UCase(strT)
    SQLDriver = strT
    If Len(SQLDriver) = 255 Then
        SQLDriver = ""
    End If
    '=========================================================
    
    sPath = "Options"
    sSetting = "DatabaseType"
    strT = ""
    'giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
    strT = INIRead(sPath, sSetting, glbHostFile)
    gsSystemDb = UCase(strT)
    
    ' dkostka - 02/16/01 - Added option to pass directories via command line options
    Call SetCmdLinePath
    '--- SET THE VALUES TO glbAdoIHRDB,glbAdoIHRAUDIT, ...
    ' Jaddy Jan 19, 2005 set default linamar login
    'Call SetLinamarLogin
    
    Call glbAdo_Value
End Sub

Public Function Get_Template(glbSecUSERID)
    Dim rsSecBasic As New ADODB.Recordset
    Dim SQLQ As String
    
    SQLQ = "SELECT USERID, SECURE_TEMPLATE FROM HR_SECURE_BASIC WHERE USERID = '" & Replace(glbSecUSERID, "'", "''") & "'"
    rsSecBasic.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsSecBasic.EOF Then
        If Not IsNull(rsSecBasic("SECURE_TEMPLATE")) Then
            Get_Template = rsSecBasic("SECURE_TEMPLATE")
        Else
            Get_Template = ""
        End If
    Else
        Get_Template = ""
    End If
    rsSecBasic.Close
    Set rsSecBasic = Nothing
    
End Function

Public Function Weekdays(startDate, EndDate)
    Dim varDays, varWeekendDays
 
    ' The number of weekend days per week.
    Const ncNumberOfWeekendDays As Integer = 2
    
    ' Calculate the number of days inclusive (+ 1 is to add back startDate).
    varDays = DateDiff("d", startDate, EndDate) + 1
    
    ' Calculate the number of weekend days.
    varWeekendDays = (DateDiff("ww", startDate, EndDate) * ncNumberOfWeekendDays) _
        + IIf(DatePart("w", startDate) = vbSunday, 1, 0) _
        + IIf(DatePart("w", EndDate) = vbSaturday, 1, 0)
    
    ' Calculate the number of weekdays.
    Weekdays = (varDays - varWeekendDays)
    
End Function

Public Function AddWorkingDays(startDate, NoDays, exclStatHoliday As Boolean) As Date
    'Add working days (NoDays) to the StartDate, exclude Weekends (Saturday and Sunday) and if Statutory Holidays
    'are to be excluded (exclStatHoliday = True) then exclude Statutory Holidays.
    
    AddWorkingDays = startDate
    
    'Add one day at a time to the Start Date and skip Weekend or Stat Holiday if date is not working day
    Do While NoDays >= 0
        'Check if Date is Weekend. If so skip the day without affecting NoDays
        If Weekday(AddWorkingDays) = vbSaturday Or Weekday(AddWorkingDays) = vbSunday Then
            AddWorkingDays = DateAdd("d", 1, AddWorkingDays)
        Else
            'Check if Statutory Holiday. If so skip the day without affecting NoDays
            If exclStatHoliday Then
                If InStr(IsSTATHoliday(CVDate(AddWorkingDays), CVDate(AddWorkingDays)), Date_SQL(AddWorkingDays)) > 0 Then
                    AddWorkingDays = DateAdd("d", 1, AddWorkingDays)
                Else
                    AddWorkingDays = DateAdd("d", 1, AddWorkingDays)
                    NoDays = NoDays - 1
                End If
            Else
                AddWorkingDays = DateAdd("d", 1, AddWorkingDays)
                NoDays = NoDays - 1
            End If
        End If
    Loop
    
    AddWorkingDays = DateAdd("d", -1, AddWorkingDays)
    
End Function

Private Function GetValidFolder(xPath)
Dim xFileName
On Error GoTo Errline
    'xFileName = xPath & "A_Test.txt"
    'If (Dir(xFileName)) <> "" Then Kill xFileName
    'Open xFileName For Output As #1
    'Print #1, "This is a test"
    'Close #1
    If Dir(xPath, vbDirectory) = "" Then
        GetValidFolder = ""
    Else
        GetValidFolder = xPath
    End If
Exit Function
Errline:
    GetValidFolder = ""

End Function

Public Function Check_FollowUp_Email_Sending_Log()
    Dim xLogPath As String
    Dim iFileNo As Integer
    Dim buf As String
    Dim xErrFlg  As Boolean
    Dim xMissingFlg As Boolean
    Dim xValidFlie As String
    
    'Ticket #20598 Franks 10/11/2012 - don't use this when test the project
    If UCase(Left(App.Path, 10)) = "C:\SSWORK\" Then
        Exit Function
    End If
    
    'Retrieve the Log file from the path mentioned on the Company Pref.
    'Open today's log file
    
    'Get the location to save the file in
    xLogPath = GetComPreferEmail("FOLLOWUPEMAILLOGPATH")
    If Len(xLogPath) = 0 Then
        'Path not found for Follow Up Email Sending log
        MsgBox "Follow Up Email Sending Log file path not found. Please specify the path to the log file on the 'Company Preference' screen under the 'Setup' menu", vbExclamation, lStr("Follow-ups Email Sending Log")
    Else
        If Right(xLogPath, 1) <> "\" Then
            xLogPath = xLogPath & "\"
        End If
    End If
    
    'Initialise
    xErrFlg = False
    xMissingFlg = False
    
    xValidFlie = GetValidFolder(xLogPath)
    If xValidFlie = "" Then
        Exit Function
    End If
    
    If Dir(xLogPath & "FollowUpEmailLog_" & Format(Now, "mmddyyyy") & ".csv") <> "" Then
        'Log file found
        'Check the contents of the file for error
        
        iFileNo = FreeFile
        
        'open the file for reading
        Open xLogPath & "FollowUpEmailLog_" & Format(Now, "mmddyyyy") & ".csv" For Input As #iFileNo
        
        'Read the contents of the file
        Do While Not EOF(iFileNo)
            Line Input #iFileNo, buf
            
            If InStr(1, buf, "Error") > 0 Then
                'Error string found
                xErrFlg = True
                
                If xMissingFlg Then
                    'No need to check the rest of the file as at least one "error" and "email missing" found
                    Exit Do
                End If
            ElseIf InStr(1, buf, "Reporting Authority email missing") > 0 Then
                'Reporting Authority Email missing
                xMissingFlg = True
                
                If xErrFlg Then
                    'No need to check the rest of the file as at least one "error" and "email missing" found
                    Exit Do
                End If
            End If
            
        Loop
    Else
        'Log file not found
        'May be there was no follow-ups that were due
    End If
    
    'If xErrFlg And xMissingFlg Then
    If xErrFlg Or xMissingFlg Then
        'Error and Reporting Authority #1 Email Missing found
        frmCustomMsgBox.lblMsg.Caption = "Errors were encountered when attempting to send some " & lStr("Follow-ups") & " emails." & vbCrLf & vbCrLf & "Please open the '" & xLogPath & "FollowUpEmailLog_" & Format(Now, "mmddyyyy") & ".CSV for more information."
        frmCustomMsgBox.lblMsg.Alignment = 0
        frmCustomMsgBox.Caption = "Error sending " & lStr("Follow-ups") & " emails"
        frmCustomMsgBox.NoButton.Visible = False
        frmCustomMsgBox.YesButton.Caption = "Close"
        frmCustomMsgBox.YesButton.Visible = True
        frmCustomMsgBox.UnButton.Caption = "Open Log"
        frmCustomMsgBox.Show 1
        If glbMsgCustomVal = 3 Then
            'Open the file
            If Not LanchXlsW98(xLogPath & "FollowUpEmailLog_" & Format(Now, "mmddyyyy") & ".CSV") Then
                Shell "cmd /c " & GetShortName(xLogPath & "FollowUpEmailLog_" & Format(Now, "mmddyyyy") & ".CSV")
            End If
        ElseIf glbMsgCustomVal = 1 Then
            'Close the message - do nothing
        End If
        
        'MsgBox "Errors were encountered when attempting to send some " & lStr("Follow-ups") & " emails." & vbCrLf & "Please open the '" & xLogPath & "FollowUpEmailLog_" & Format(Now, "mmddyyyy") & ".CSV for more information.", vbExclamation, "Error sending " & lStr("Follow-ups") & " emails"
        
    'ElseIf xErrFlg Then
    '    'Error found
    '    MsgBox "Error(s) found when sending the " & lStr("Follow-ups") & " emails." & vbCrLf & "Please check the log for more information : " & xLogPath & "FollowUpEmailLog_" & Format(Now, "mmddyyyy") & ".txt", vbExclamation, lStr("Follow-ups") & " Email Sending Log"
    'ElseIf xMissingFlg Then
    '    'Reporting Authority #1 Email Missing found
    '    MsgBox "Reporting Authority #1 was missing when sending the " & lStr("Follow-ups") & " emails." & vbCrLf & "Please check the log for more information : " & xLogPath & "FollowUpEmailLog_" & Format(Now, "mmddyyyy") & ".txt", vbExclamation, lStr("Follow-ups") & " Email Sending Log"
    End If
    
    'close the file
    Close #iFileNo

End Function

Function LanchXlsW98(xFileName)
On Error GoTo Error_Deal
    LanchXlsW98 = False
    Shell "Start " & GetShortName(xFileName)
    LanchXlsW98 = True
Exit Function
Error_Deal:

End Function

Public Function HRSSCharCount(xList, xChar) 'Franks 04/30/2012
Dim I As Integer
Dim retVal As Integer
    retVal = 0
    For I = 1 To Len(xList)
        If Mid(xList, I, 1) = xChar Then
            retVal = retVal + 1
        End If
    Next
    HRSSCharCount = retVal
End Function

Public Function GetString(o As Object) As String
    Dim rtnvalue As String
    rtnvalue = ""
    On Error GoTo rtn
    rtnvalue = CStr(o)
    GetString = rtnvalue
rtn:
GetString = ""
End Function

Public Function GetDouble(o As Object) As Double
    Dim rtnvalue As Double
    rtnvalue = ""
    On Error GoTo rtn
    rtnvalue = CDbl(o)
    GetDouble = rtnvalue
rtn:
GetDouble = 0
End Function

Public Function GetInt(o As Object) As Integer
    Dim rtnvalue As Integer
    rtnvalue = ""
    On Error GoTo rtn
    rtnvalue = CInt(o)
    GetInt = rtnvalue
rtn:
GetInt = 0
End Function

Public Function getCodeDesc(TblName As String, TblKey As String) As String
    Dim SQLQ As String
    Dim rsTABL As New ADODB.Recordset
    On Error GoTo CodeDesc_Err
        
    SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = '" & TblName & "' AND TB_KEY = '" & TblKey & "' "
    rsTABL.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTABL.EOF Then
        getCodeDesc = rsTABL("TB_DESC")
    Else
        getCodeDesc = ""
    End If
    rsTABL.Close
    Exit Function

CodeDesc_Err:
    Resume Next
End Function

Public Function Accrual_Rec_Exists(xEmpNo, xAccType, xAccDate, xUpdType) As Boolean
    Dim rsAccrual As New ADODB.Recordset
    Dim SQLQ As String
    
    'Function to check if Accrual record for this update already exists. If so, will warn the user.
    'This is to avoid user from doing the multiple entitlement updates for the same period/date.
    SQLQ = "SELECT * FROM HR_ACCRUAL WHERE AC_EMPNBR = " & xEmpNo & " AND AC_TYPE = '" & xAccType & "'"
    SQLQ = SQLQ & " AND AC_EDATE = " & Date_SQL(xAccDate)
    If xAccType <> "VAC" And xAccType <> "SICK" Then
        SQLQ = SQLQ & " AND AC_ACTION IN (" & xUpdType & ")"
    Else
        'Ticket #28024 - To fix the error caused by calling this function without '' apostrophes
        'SQLQ = SQLQ & " AND AC_ACTION = '" & xUpdType & "'"
        SQLQ = SQLQ & " AND AC_ACTION IN (" & xUpdType & ")"
    End If
    rsAccrual.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If Not rsAccrual.EOF Then
        'Accrual record exists
        Accrual_Rec_Exists = True
    Else
        'Accrual record do not exists
        Accrual_Rec_Exists = False
    End If
    rsAccrual.Close
    Set rsAccrual = Nothing
    
End Function

Public Function Latest_Termination(xEmpNo, xTermSEQ, xTermDate) As Boolean
    Dim SQLQ As String
    Dim rsHRTrmEmp As New ADODB.Recordset
    
    Latest_Termination = False
    
    SQLQ = "SELECT * FROM Term_HRTRMEMP WHERE "
    SQLQ = SQLQ & "Employee_Number = " & xEmpNo '& " AND TERM_SEQ = " & xTermSEQ
    SQLQ = SQLQ & " ORDER BY Term_DOT DESC"
    rsHRTrmEmp.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    If Not rsHRTrmEmp.EOF Then
        rsHRTrmEmp.MoveFirst
        
        'Check if the first record's Term Date is more recent compared to the record selected by the user
        If CVDate(rsHRTrmEmp("Term_DOT")) > CVDate(xTermDate) Then
            'User selected Term record is not the most recent term record of the employee
            Latest_Termination = False
        Else
            'User selected Term record is the most recent one
            Latest_Termination = True
        End If
    Else
        Latest_Termination = True
    End If
End Function

Public Function Get_UserID_Info(xUserID, xInfo, xDefault)
Dim rsUser As New ADODB.Recordset
Dim SQLQ As String

SQLQ = "SELECT " & xInfo & " FROM HR_SECURE_BASIC WHERE USERID = '" & Replace(xUserID, "'", "''") & "' "
rsUser.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsUser.EOF Then
    Get_UserID_Info = IIf(IsNull(rsUser(xInfo)), xDefault, rsUser(xInfo))
Else
    Get_UserID_Info = xDefault
End If
rsUser.Close
Set rsUser = Nothing
End Function

Sub SeniorityDateCalculation()
    frmSenDateCalc.Show 1
End Sub

Public Function EncryptDatabaseSettings(strHRSSLic) As String
    'Ticket #24352 - PIPEDA
    'Encrypt the database settings using the HRSS Control so it can be written on the Registry

    'Define the control
    Dim hrssCtrl As New HRSSControl.HRSSControlDll
    Dim secKey As String
    Dim EncryptedStr As String

    'Create the Key for Encryption which will be used also Decryption
    'secKey = hrssCtrl.CreateKey

    EncryptedStr = ""

    'If Len(secKey) > 0 Then
        EncryptedStr = hrssCtrl.Encrypt(strHRSSLic, "info:HR999999999master")
    'Else
    '    MsgBox "Please provide the Key to Encrypt the string"
    'End If

    EncryptDatabaseSettings = EncryptedStr
End Function

Public Function DecryptDatabaseSettings(strHRSSLic) As String
    'Ticket #24352 - PIPEDA
    'Decrypt the database settings from Registry using the HRSS Control so it can be read and connect to the database

    'Define the control
    Dim hrssCtrl As New HRSSControl.HRSSControlDll
    Dim DecryptedStr As String

    DecryptedStr = ""

    If Len(strHRSSLic) > 0 Then
        DecryptedStr = hrssCtrl.Decrypt(strHRSSLic, "info:HR999999999master")
    Else
        MsgBox "Please provide the Key to Decrypt the string"
    End If

    DecryptDatabaseSettings = DecryptedStr
End Function

Public Function DatabaseConnection_License(SQLHRSSLicense As String)
    Dim sPath As String
    Dim sSetting As String
    Dim strT As String

    sPath = "ODBC Setup"

    sSetting = "DATABASENAME"
    strT = ""
    strT = ExtractConnectionStr(sSetting, SQLHRSSLicense)
    SQLDatabaseName = strT
    If Len(SQLDatabaseName) = 255 Then
        SQLDatabaseName = "INFOHR"
    End If

    'SQL SERVER LOGIN SERVER NAME, USER NAME AND PASSWORD
    sSetting = "SERVERNAME"
    strT = ""
    strT = ExtractConnectionStr(sSetting, SQLHRSSLicense)
    SQLServerName = strT
    If Len(SQLServerName) = 255 Then
        SQLServerName = ""
    End If

    sSetting = "USERNAME"
    strT = ""
    strT = ExtractConnectionStr(sSetting, SQLHRSSLicense)
    SQLUserName = strT
    If Len(SQLUserName) = 255 Then
        SQLUserName = ""
    End If

    sSetting = "USERPSW"
    strT = ""
    strT = ExtractConnectionStr(sSetting, SQLHRSSLicense)
    If gsMultiLang = "Y" Then
        SQLUserPassword = DecryptPasswordMultiLang_First(strT)
    ElseIf gsMultiLang = "YES" Then 'WHSCC
        SQLUserPassword = DecryptPasswordMultiLang(strT)
    Else
        glbDBPassFlag = True
        SQLUserPassword = DecryptPassword(strT)
    End If

    'Oracle driver NAME, USER NAME AND PASSWORD
    sSetting = "DRIVERNAME"
    strT = ""
    strT = ExtractConnectionStr(sSetting, SQLHRSSLicense)
    SQLDriver = strT
    If Len(SQLDriver) = 255 Then
        SQLDriver = ""
    End If

End Function

Public Function ExtractConnectionStr(sSetting, SQLHRSSLicense)
    Dim sConnFields
    Dim sVal As String

    sConnFields = Split(SQLHRSSLicense, "|")
    If sSetting = "DATABASENAME" Then
        sVal = Mid(sConnFields(0), InStr(1, sConnFields(0), "=") + 1)
    End If
    If sSetting = "SERVERNAME" Then
        sVal = Mid(sConnFields(2), InStr(1, sConnFields(2), "=") + 1)
    End If
    If sSetting = "USERNAME" Then
        sVal = Mid(sConnFields(3), InStr(1, sConnFields(3), "=") + 1)
    End If
    If sSetting = "USERPSW" Then
        sVal = Mid(sConnFields(4), InStr(1, sConnFields(4), "=") + 1)
    End If
    If sSetting = "DRIVERNAME" Then
        sVal = Mid(sConnFields(1), InStr(1, sConnFields(1), "=") + 1)
    End If
    ExtractConnectionStr = sVal
End Function

Public Function Get_CompanyPreference_Value(xFunction)
    Dim rsPrefer As New ADODB.Recordset
    
    Get_CompanyPreference_Value = False
    
    rsPrefer.Open "SELECT * FROM HRPREFERENCE WHERE HP_FUN_NAME = '" & xFunction & "'", gdbAdoIhr001, adOpenStatic
    If Not rsPrefer.EOF Then
        Get_CompanyPreference_Value = rsPrefer("HP_ENABLED")
    Else
        Get_CompanyPreference_Value = False
    End If
    rsPrefer.Close
    Set rsPrefer = Nothing
End Function

Public Sub AddRemove_ODBCSetup_RegistryKey(xLicAddRemove)
    'This function Adds or Deletes the License Key from the Registry
    'Also,
    '   if it Adds the License Key then it will Remove the ODBC Setup Key for Database Connection
    '   if it Deletes the License Key then it will Add the ODBC Setup Key for Database Connection
    
    Dim Response%, w%, x%, Y%, SECTION$, Key$, xPWD$, valtmp, I
    Dim sPath As String, sSetting As String, strT  As String
    Dim hKey As Long
    Dim SQLHRSSLicense As String
    Dim strHRSSLic As String
    Dim strHRSSLicEncrypt As String
    Dim xPswd As String
        
    If Not glbHosted Then
        If Not DoesKeyExist(HKEY_LOCAL_MACHINE, REG_NAME) Then
            lCurrentKey = HKEY_CURRENT_USER
        Else
            lCurrentKey = HKEY_LOCAL_MACHINE
        End If
        
        If xLicAddRemove = "RemoveLic" Then     'Remove License Key and Add ODBC Setup
            'First - Retrieve & Decrypt the Datbaase Connection values from the License Key so it can be used for
            'recreating the ODBC Setup for Database Connection
            glbHRSSSecure = True
            sPath = REG_NAME & "Options"
            sSetting = "License"
            strT = ""
            giGar = bGetRegistrySetting(lCurrentKey, sPath, sSetting, strT)
            SQLHRSSLicense = strT
            If Len(SQLHRSSLicense) = 344 Then
                SQLHRSSLicense = ""
            Else
                'Decrypt the License key
                SQLHRSSLicense = DecryptDatabaseSettings(SQLHRSSLicense)
        
                'Break down the license key to appropriate database connection variables
                If Len(SQLHRSSLicense) > 0 Then
                    Call DatabaseConnection_License(SQLHRSSLicense)
                End If
            End If
            glbHRSSSecure = False
                        
                        
            'Second - Add back the ODBC Setup for Database Connection (HKEY_LOCAL_MACHINE\Software\HR Systems\ODBC Setup\)
            '       - DATABASENAME
            '       - DRIVERNAME
            '       - SERVERNAME
            '       - USERNAME
            '       - USERPSW
            SECTION$ = REG_NAME & "ODBC Setup"
    
            'DATABASENAME
            x% = WriteRegistrySetting(lCurrentKey, SECTION$, "DATABASENAME", SQLDatabaseName)  'gbasINI_WritePrivateString
            'SQLDatabaseName = txtDatabase.Text
    
            'DRIVERNAME
            x% = WriteRegistrySetting(lCurrentKey, SECTION$, "DRIVERNAME", SQLDriver)  'gbasINI_WritePrivateString
            'SQLDriver = cboDrivers.Text
    
            'SERVERNAME
            x% = WriteRegistrySetting(lCurrentKey, SECTION$, "SERVERNAME", SQLServerName)  'gbasINI_WritePrivateString
            'SQLServerName = txtServer.Text
    
            'USERNAME
            x% = WriteRegistrySetting(lCurrentKey, SECTION$, "USERNAME", SQLUserName)  'gbasINI_WritePrivateString
            'SQLUserName = txtUserName.Text
    
            'USERPSW
            xPswd = SQLUserPassword
            If gsMultiLang = "Y" Then 'For Listowel only
                xPWD$ = EncryptPasswordMultiLang_First(xPswd)
            ElseIf UCase(gsMultiLang) = "YES" Then 'For general multi language clients
                xPWD$ = EncryptPasswordMultiLang(xPswd)
            'For version 7.6 only ticket# 9153
            Else
                xPWD$ = EncryptPassword(xPswd)
            End If
    
            x% = WriteRegistrySetting(lCurrentKey, SECTION$, "USERPSW", xPWD$)  'gbasINI_WritePrivateString
            'SQLUserPassword = txtUserPsw.Text
                      
            
            'Finally - Delete the Licence Key (HKEY_LOCAL_MACHINE\Software\HR Systems\Options\License)
            sPath = REG_NAME & "Options"
                        
            'Deleting the License Key
            sSetting = "License"
            If DeleteRegistryValue(lCurrentKey, sPath, sSetting) = 0 Then
                'MsgBox "Cannot Delete Key"
            Else
                'MsgBox "Deleted"
            End If
            
        ElseIf xLicAddRemove = "AddLic" Then    'Add License Key and Remove ODBC Setup
            'First - Retrieve the Database Connection values from ODBC Setup for Database Connection (HKEY_LOCAL_MACHINE\Software\HR Systems\ODBC Setup\)
            '       - DATABASENAME
            '       - DRIVERNAME
            '       - SERVERNAME
            '       - USERNAME
            '       - USERPSW
            'Values already stored in the Global variables so no need to retrive from Registry
            
            strHRSSLic = ""

            'DATABASENAME
            strHRSSLic = strHRSSLic & "DATABASENAME=" & SQLDatabaseName & "|"
            'SQLDatabaseName = txtDatabase.Text
            
            'DRIVERNAME
            strHRSSLic = strHRSSLic & "DRIVERNAME=" & SQLDriver & "|"
            'SQLDriver = cboDrivers.Text
            
            'DRIVERNAME
            strHRSSLic = strHRSSLic & "DRIVERNAME=" & SQLServerName & "|"
            'SQLServerName = txtServer.Text
            
            'USERNAME
            strHRSSLic = strHRSSLic & "USERNAME=" & SQLUserName & "|"
            'SQLUserName = txtUserName.Text
                                
            'USERPSW
            xPswd = SQLUserPassword
            If gsMultiLang = "Y" Then 'For Listowel only
                xPWD$ = EncryptPasswordMultiLang_First(xPswd)
            ElseIf UCase(gsMultiLang) = "YES" Then 'For general multi language clients
                xPWD$ = EncryptPasswordMultiLang(xPswd)
            'For version 7.6 only ticket# 9153
            Else
                xPWD$ = EncryptPassword(xPswd)
            End If
            
            strHRSSLic = strHRSSLic & "USERPSW=" & xPWD$ & "|"
            'SQLUserPassword = txtUserPsw.Text
            
            
            'Second - Encrypted the Database Connection values and Create the License Key (HKEY_LOCAL_MACHINE\Software\HR Systems\Options\License)
            'Ticket #24352 - PIPEDA
            If Len(strHRSSLic) > 0 Then
                strHRSSLicEncrypt = EncryptDatabaseSettings(strHRSSLic)
                If Len(strHRSSLicEncrypt) > 0 Then
                    SECTION$ = REG_NAME & "Options"
                    x% = WriteRegistrySetting(lCurrentKey, SECTION$, "License", strHRSSLicEncrypt)
                End If
            End If
        
        
            'Finally - Delete the Database Connection values from ODBC Setup (HKEY_LOCAL_MACHINE\Software\HR Systems\ODBC Setup\)
            '       - DATABASENAME
            '       - DRIVERNAME
            '       - SERVERNAME
            '       - USERNAME
            '       - USERPSW
            sPath = REG_NAME & "ODBC Setup"
                        
            'Deleting the DATABASENAME Key
            sSetting = "DATABASENAME"
            If DeleteRegistryValue(lCurrentKey, sPath, sSetting) = 0 Then
                'MsgBox "Cannot Delete Key"
            Else
                'MsgBox "Deleted"
            End If
            
            'Deleting the DRIVERNAME Key
            sSetting = "DRIVERNAME"
            If DeleteRegistryValue(lCurrentKey, sPath, sSetting) = 0 Then
                'MsgBox "Cannot Delete Key"
            Else
                'MsgBox "Deleted"
            End If
        
            'Deleting the SERVERNAME Key
            sSetting = "SERVERNAME"
            If DeleteRegistryValue(lCurrentKey, sPath, sSetting) = 0 Then
                'MsgBox "Cannot Delete Key"
            Else
                'MsgBox "Deleted"
            End If
        
            'Deleting the USERNAME Key
            sSetting = "USERNAME"
            If DeleteRegistryValue(lCurrentKey, sPath, sSetting) = 0 Then
                'MsgBox "Cannot Delete Key"
            Else
                'MsgBox "Deleted"
            End If
        
            'Deleting the USERPSW Key
            sSetting = "USERPSW"
            If DeleteRegistryValue(lCurrentKey, sPath, sSetting) = 0 Then
                'MsgBox "Cannot Delete Key"
            Else
                'MsgBox "Deleted"
            End If
        End If
    End If
        
End Sub

Public Function Check_Daily_Accrual_Exists(xSQLQ)
    Dim SQLQ As String
    Dim rsDailyAcc As New ADODB.Recordset

    'Check Daily Accrual table to see if the accrual details already exists for employees in the current selection
    Screen.MousePointer = HOURGLASS
    
    'Condition for the selection
    'If IsMissing(xSQLQ) Then
    '    'Get the current selected rule
    '    Call getWSQLQ_DailyAccrual
    '
    '    SQLQ = "SELECT DA_EMPNBR FROM HR_DAILYVACACCR WHERE " & fglbVSQLQ & " "
    'Else
        SQLQ = "SELECT DA_EMPNBR FROM HR_DAILYVACACCR WHERE " & xSQLQ & " "
    'End If
    SQLQ = SQLQ & " GROUP BY DA_EMPNBR"
    rsDailyAcc.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsDailyAcc.BOF And rsDailyAcc.EOF Then
        'Employee's Daily Accrual Do NOT exists
        Check_Daily_Accrual_Exists = False
    Else
        'Employee's Daily Accrual exists
        Check_Daily_Accrual_Exists = True
    End If
    rsDailyAcc.Close
    Set rsDailyAcc = Nothing
    
    Screen.MousePointer = DEFAULT
End Function

Public Function Clear_Employees_Daily_Accruals(xSQLQ, Optional xEffDate)
    Dim SQLQ As String
    Dim rsDailyAcc As New ADODB.Recordset
    Dim curVac
    Dim xComments As String
    Dim lngRecs As Long, pct As Long, prec As Long
    
    On Error GoTo Clear_Employees_Daily_Accruals_Err
    
    'Delete the Daily Accrual Details of the Employees for the selected rule.
    Screen.MousePointer = HOURGLASS
    
    Clear_Employees_Daily_Accruals = False
    
    'Get condition for the selection
    'If IsMissing(xSQLQ) Then
    '    'Get the current selected rule
    '    Call getWSQLQ_DailyAccrual
    '
    '    SQLQ = "SELECT DA_EMPNBR FROM HR_DAILYVACACCR WHERE " & fglbVSQLQ & " "
    'Else
        SQLQ = "SELECT DA_EMPNBR FROM HR_DAILYVACACCR WHERE " & xSQLQ & " "
    'End If
    SQLQ = SQLQ & " GROUP BY DA_EMPNBR"
    rsDailyAcc.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsDailyAcc.BOF And rsDailyAcc.EOF Then
        Screen.MousePointer = DEFAULT
        'flgNoErrorClrAcc = True
        'MsgBox "Employees for this selection do not exist!", vbOKOnly, "Clear Daily Accrual File"
    Else
        'Clear ED_VAC for all the employees in this selection
        rsDailyAcc.MoveFirst
        
        MDIMain.panHelp(0).FloodType = 1
        lngRecs = 0
        prec = 0
        
        Do While Not rsDailyAcc.EOF
            lngRecs = rsDailyAcc.RecordCount
            prec = prec + 1
            pct = Int(100 * (prec / lngRecs))
            MDIMain.panHelp(0).FloodPercent = pct
            
            'Get current ED_VAC value
            curVac = GetEmpData(rsDailyAcc("DA_EMPNBR"), "ED_VAC", 0)
            
            'Update ED_VAC to 0
            SQLQ = "UPDATE HREMP SET ED_VAC = 0 WHERE ED_EMPNBR = " & rsDailyAcc("DA_EMPNBR")
            gdbAdoIhr001.Execute SQLQ
        
            'Update Accrual Table with the cleared value
            xComments = "Current Vac. Ent. Chg from " & curVac & " to 0"
            
            If Not IsMissing(xEffDate) Then
                If IsDate(xEffDate) Then
                    Call Append_Accrual(rsDailyAcc("DA_EMPNBR"), "VAC", CVDate(Format(xEffDate, "mm/dd/yyyy")), 0 - Round(Val(curVac), 4), "G", xComments)
                Else
                    Call Append_Accrual(rsDailyAcc("DA_EMPNBR"), "VAC", CVDate(Format(Now, "mm/dd/yyyy")), 0 - Round(Val(curVac), 4), "G", xComments)
                End If
            Else
                Call Append_Accrual(rsDailyAcc("DA_EMPNBR"), "VAC", CVDate(Format(Now, "mm/dd/yyyy")), 0 - Round(Val(curVac), 4), "G", xComments)
            End If
            
            rsDailyAcc.MoveNext
        Loop
        
        MDIMain.panHelp(0).FloodType = 0
    End If
    rsDailyAcc.Close
    Set rsDailyAcc = Nothing
    
    'Clear Daily Accrual table for the selected rule
    'Condition for the selection
    'If IsMissing(xSQLQ) Then
    '    SQLQ = "DELETE FROM HR_DAILYVACACCR WHERE " & fglbVSQLQ
    'Else
        SQLQ = "DELETE FROM HR_DAILYVACACCR WHERE " & xSQLQ
    'End If
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
    
    Clear_Employees_Daily_Accruals = True
    
    Screen.MousePointer = DEFAULT
                
Exit Function

Clear_Employees_Daily_Accruals_Err:

Clear_Employees_Daily_Accruals = False
'flgNoErrorClrAcc = False

Screen.MousePointer = DEFAULT
    
glbFrmCaption$ = "Daily Employee Accrual"
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Clear_Employees_Daily_Accruals", "HR_DAILYVACACCR", "Clear Accrual")
Screen.MousePointer = DEFAULT
'If gintRollBack% = False Then
    'Rollback
    Resume Next
'Else
'    Unload Me
'End If
                
End Function

Public Function EntRecalVacDaily(xSQLQ, Optional xEffDate) As Boolean
    Dim SQLQ As String
    Dim rsDailyAcc As New ADODB.Recordset
    Dim curVac
    Dim xComments As String
    Dim lngRecs As Long, pct As Long, prec As Long
    
    On Error GoTo EntRecalVacDaily_Err
    
    EntRecalVacDaily = False
    
    'Recalculate the Current year Vacation based on the Daily Accruals Earned to date and the Vacation Taken
    Screen.MousePointer = HOURGLASS
    
    'Condition for the selection
    'If IsMissing(xSQLQ) Then
    '    'Get the current selected rule
    '    Call getWSQLQ_DailyAccrual
    '
    '    'Sum the Daily Accruals earned to date
    '    SQLQ = "SELECT SUM(DA_ACCRAMT) AS VAC_ACCRUED, DA_EMPNBR FROM HR_DAILYVACACCR WHERE " & fglbVSQLQ & " "
    'Else
        'Sum the Daily Accruals earned to date
        SQLQ = "SELECT SUM(DA_ACCRAMT) AS VAC_ACCRUED, DA_EMPNBR FROM HR_DAILYVACACCR WHERE " & xSQLQ & " "
    'End If
    SQLQ = SQLQ & " AND DA_ACCRDATE BETWEEN DA_FRDATE AND " & Date_SQL(Date)
    SQLQ = SQLQ & " GROUP BY DA_EMPNBR"
    rsDailyAcc.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsDailyAcc.BOF And rsDailyAcc.EOF Then
        'MsgBox "Employees for this selection do not exist!"
    Else
        'Update ED_VAC and ED_VACT in HREMP for all the employees in this selection
        rsDailyAcc.MoveFirst
        
        MDIMain.panHelp(0).FloodType = 1
        lngRecs = 0
        prec = 0
        
        Do While Not rsDailyAcc.EOF
            lngRecs = rsDailyAcc.RecordCount
            prec = prec + 1
            pct = Int(100 * (prec / lngRecs))
            MDIMain.panHelp(0).FloodPercent = pct
            
            'Get current ED_VAC value
            curVac = GetEmpData(rsDailyAcc("DA_EMPNBR"), "ED_VAC", 0)
                        
            'Update Current Vacation
            SQLQ = "UPDATE HREMP SET ED_VAC = " & rsDailyAcc("VAC_ACCRUED") & " WHERE ED_EMPNBR = " & rsDailyAcc("DA_EMPNBR")
            gdbAdoIhr001.Execute SQLQ
            
            'Update Annual Vacation
            SQLQ = "UPDATE HREMP SET ED_ANNVAC = " & Get_AnnualVac_From_DailyAccrual(rsDailyAcc("DA_EMPNBR"), GetEmpData(rsDailyAcc("DA_EMPNBR"), "ED_ETDATE")) & " WHERE ED_EMPNBR = " & rsDailyAcc("DA_EMPNBR")
            gdbAdoIhr001.Execute SQLQ
            
            'Check if current value is different from computed value of Vacation accrual todate
            'Update Accrual table
            If Round(Val(curVac), 4) <> Round(Val(rsDailyAcc("VAC_ACCRUED")), 4) Then
                'Update Accrual table with the difference
                xComments = "Current Vac. Ent. Chg from " & curVac & " to " & rsDailyAcc("VAC_ACCRUED")
                
                If Not IsMissing(xEffDate) Then
                    If IsDate(xEffDate) Then
                        Call Append_Accrual(rsDailyAcc("DA_EMPNBR"), "VAC", CVDate(Format(xEffDate, "mm/dd/yyyy")), Round(Val(rsDailyAcc("VAC_ACCRUED")), 4) - Round(Val(curVac), 4), "J", xComments)
                    Else
                        Call Append_Accrual(rsDailyAcc("DA_EMPNBR"), "VAC", CVDate(Format(Now, "mm/dd/yyyy")), Round(Val(rsDailyAcc("VAC_ACCRUED")), 4) - Round(Val(curVac), 4), "J", xComments)
                    End If
                Else
                    'Call Append_Accrual(rsDailyAcc("DA_EMPNBR"), "VAC", CVDate(Format(Now, "mm/dd/yyyy")), Round(Val(curVac), 4) - Round(Val(rsDailyAcc("VAC_ACCRUED")), 4), "J", xComments)
                    Call Append_Accrual(rsDailyAcc("DA_EMPNBR"), "VAC", CVDate(Format(Now, "mm/dd/yyyy")), Round(Val(rsDailyAcc("VAC_ACCRUED")), 4) - Round(Val(curVac), 4), "J", xComments)
                End If
            End If
            
            'Update the Daily Accrual table with Process Date since it update employee's ED_VAC
            SQLQ = "UPDATE HR_DAILYVACACCR SET DA_PROCESSDATE = " & Date_SQL(Date)
            SQLQ = SQLQ & " WHERE DA_EMPNBR = " & rsDailyAcc("DA_EMPNBR")
            SQLQ = SQLQ & " AND (DA_PROCESSDATE IS NULL OR DA_PROCESSDATE = '')"
            SQLQ = SQLQ & " AND DA_ACCRDATE BETWEEN DA_FRDATE AND " & Date_SQL(Date)
            gdbAdoIhr001.Execute SQLQ
            
            'Update Vacation Taken
            SQLQ = "UPDATE HREMP SET ED_VACT = 0 WHERE ED_EMPNBR = " & rsDailyAcc("DA_EMPNBR")
            gdbAdoIhr001.Execute SQLQ
        
            SQLQ = " Update HREMP SET "
            SQLQ = SQLQ & " ED_VACT =(SELECT SUM(AD_HRS) FROM HR_ATTENDANCE"
            SQLQ = SQLQ & " WHERE ED_EMPNBR = AD_EMPNBR"
            SQLQ = SQLQ & " AND AD_DOA BETWEEN ED_EFDATE AND ED_ETDATE"
            SQLQ = SQLQ & " AND AD_REASON Like 'VAC%')"
            SQLQ = SQLQ & " WHERE ED_EMPNBR IN"
            SQLQ = SQLQ & " (SELECT AD_EMPNBR FROM HR_ATTENDANCE INNER JOIN HREMP ON HR_ATTENDANCE.AD_EMPNBR=HREMP.ED_EMPNBR"
            SQLQ = SQLQ & " WHERE (AD_DOA BETWEEN ED_EFDATE AND ED_ETDATE)"
            SQLQ = SQLQ & " AND AD_REASON Like 'VAC%')"
            SQLQ = SQLQ & " AND ED_EMPNBR = " & rsDailyAcc("DA_EMPNBR")
            gdbAdoIhr001.Execute SQLQ
            
            rsDailyAcc.MoveNext
        Loop
        
        MDIMain.panHelp(0).FloodType = 0
    End If
    rsDailyAcc.Close
    Set rsDailyAcc = Nothing
    
    EntRecalVacDaily = True
    
    Screen.MousePointer = DEFAULT
    
    'MsgBox "Recalculate Completed Successfully for the employees belonging to this Entitlement Rule", vbInformation, "Daily Accrual Recalculate"
    
Exit Function

EntRecalVacDaily_Err:

EntRecalVacDaily = False

Screen.MousePointer = DEFAULT
    
glbFrmCaption$ = "Daily Employee Accrual"
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EntRecalVacDaily", "HREMP", "Recalculate")
Screen.MousePointer = DEFAULT
'If gintRollBack% = False Then
'    'Rollback
    Resume Next
'Else
'    Unload Me
'End If
   
End Function

'Employee #, Status, Union, Category, Excluded Status, Hours/Day, FTE, Date Skipped, Accrual Missed, Reason
Public Sub Log_Skipped_Transaction(xEmpNo, xORG, xEMP, xPT, xEmpExclude, xFromDate, xToDate, xDHRS, xFte, xSkipDate, xAccAmt, xSkipReason)
    Dim rsDailyAccLog As New ADODB.Recordset

    rsDailyAccLog.Open "SELECT * FROM HR_DAILYACC_LOG WHERE 0=1", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    rsDailyAccLog.AddNew
    rsDailyAccLog("DL_COMPNO") = "001"
    rsDailyAccLog("DL_EMPNBR") = xEmpNo
    rsDailyAccLog("DL_ORG") = xORG
    rsDailyAccLog("DL_EMP") = xEMP
    rsDailyAccLog("DL_PT") = xPT
    rsDailyAccLog("DL_EMPEXCL") = xEmpExclude
    rsDailyAccLog("DL_FRDATE") = xFromDate
    rsDailyAccLog("DL_TODATE") = xToDate
    rsDailyAccLog("DL_DHRS") = xDHRS
    rsDailyAccLog("DL_FTENUM") = xFte
    rsDailyAccLog("DL_SKIPDATE") = xSkipDate
    rsDailyAccLog("DL_ACCRAMT") = IIf(xAccAmt = "", Null, xAccAmt)
    rsDailyAccLog("DL_SKIPREASON") = Left(xSkipReason, 250)
    rsDailyAccLog("DL_LUSER") = glbUserID
    rsDailyAccLog("DL_LDATE") = Date
    rsDailyAccLog("DL_LTIME") = Time$
    rsDailyAccLog.Update
    rsDailyAccLog.Close
    Set rsDailyAccLog = Nothing
End Sub

Public Function Accrued_ToDate(xEmpNo, xORG, xEMP, xPT, xEmpExclude, xFromDate, xToDate, Optional xAsOfDate)
    Dim rsDailyAcc As New ADODB.Recordset
    Dim SQLQ
    
    Accrued_ToDate = 0
    
    'Get total accrual todate
    SQLQ = "SELECT SUM(DA_ACCRAMT) AS ACCTODATE FROM HR_DAILYVACACCR WHERE DA_EMPNBR = " & xEmpNo
    SQLQ = SQLQ & " AND DA_FRDATE = " & Date_SQL(xFromDate)
    SQLQ = SQLQ & " AND DA_TODATE = " & Date_SQL(xToDate)
    If Not IsMissing(xAsOfDate) Then
        SQLQ = SQLQ & " AND DA_ACCRDATE <= " & Date_SQL(xAsOfDate)
    End If
    
    'If Len(xORG) = 0 Then
    '    SQLQ = SQLQ & " AND (DA_ORG IS NULL OR DA_ORG='') "
    'Else
    '    SQLQ = SQLQ & " AND DA_ORG = '" & xORG & "'"
    'End If
    'If Len(xEMP) = 0 Then
    '    SQLQ = SQLQ & " AND (DA_EMP IS NULL OR DA_EMP='')"
    'Else
    '    SQLQ = SQLQ & " AND DA_EMP = '" & xEMP & "'"
    'End If
    'If Len(xPT) = 0 Then
    '    SQLQ = SQLQ & " AND (DA_PT IS NULL OR DA_PT='')"
    'Else
    '    SQLQ = SQLQ & " AND DA_PT = '" & xPT & "' "
    'End If
    'If Len(xEmpExclude) = 0 Then
    '    SQLQ = SQLQ & " AND (DA_EMPEXCL IS NULL OR DA_EMPEXCL='')"
    'Else
    '    SQLQ = SQLQ & " AND DA_EMPEXCL = '" & xEmpExclude & "'"
    'End If
    
    rsDailyAcc.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsDailyAcc.EOF Then
        Accrued_ToDate = IIf(IsNull(rsDailyAcc("ACCTODATE")), 0, rsDailyAcc("ACCTODATE"))
    Else
        Accrued_ToDate = 0
    End If
    rsDailyAcc.Close
    Set rsDailyAcc = Nothing

End Function

Public Sub Append_Daily_Accrul_File(xEmpNo, xORG, xEMP, xPT, xEmpExclude, xEffDate, xFromDate, xToDate, xAnnAcc, xAccDate, xAccAmt, xProcessDate, xAccdToDate, Optional xSkipped As Boolean)
    Dim rsDailyAcc As New ADODB.Recordset

    rsDailyAcc.Open "SELECT * FROM HR_DAILYVACACCR WHERE 0=1", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    rsDailyAcc.AddNew
    rsDailyAcc("DA_COMPNO") = "001"
    rsDailyAcc("DA_EMPNBR") = xEmpNo
    rsDailyAcc("DA_ORG") = xORG
    rsDailyAcc("DA_EMP") = xEMP
    rsDailyAcc("DA_PT") = xPT
    rsDailyAcc("DA_EMPEXCL") = xEmpExclude
    rsDailyAcc("DA_EDATE") = xEffDate
    rsDailyAcc("DA_FRDATE") = xFromDate
    rsDailyAcc("DA_TODATE") = xToDate
    rsDailyAcc("DA_ANNACCR") = xAnnAcc
    rsDailyAcc("DA_ACCRDATE") = xAccDate
    rsDailyAcc("DA_ACCRAMT") = xAccAmt
    rsDailyAcc("DA_PROCESSDATE") = IIf(xProcessDate = "", Null, xProcessDate)
    rsDailyAcc("DA_ACCRDTODATE") = xAccdToDate
    rsDailyAcc("DA_SKIPPED") = xSkipped
    rsDailyAcc("DA_LUSER") = glbUserID
    rsDailyAcc("DA_LDATE") = Date
    rsDailyAcc("DA_LTIME") = Time$
    rsDailyAcc.Update
    rsDailyAcc.Close
    Set rsDailyAcc = Nothing
End Sub

Public Function Get_DailyAccrual(xEmpNo, xORG, xEMP, xPT, xEmpExclude, xFromDate, xToDate, xAccDate, xProcessed As Boolean)
    Dim rsDailyAcc As New ADODB.Recordset
    Dim SQLQ As String

    'Return the Accrual of the day
    Get_DailyAccrual = 0
    
    'Retrieve the accrual of the day
    SQLQ = "SELECT SUM(DA_ACCRAMT) AS ACCTODATE FROM HR_DAILYVACACCR WHERE DA_EMPNBR = " & xEmpNo
    SQLQ = SQLQ & " AND DA_FRDATE = " & Date_SQL(xFromDate)
    SQLQ = SQLQ & " AND DA_TODATE = " & Date_SQL(xToDate)
    SQLQ = SQLQ & " AND DA_ACCRDATE = " & Date_SQL(xAccDate)
    If Not xProcessed Then
        SQLQ = SQLQ & " AND (DA_PROCESSDATE IS NULL OR DA_PROCESSDATE = '')"
    Else
        SQLQ = SQLQ & " AND DA_PROCESSDATE IS NOT NULL"
    End If
    
    If Len(xORG) = 0 Then
        SQLQ = SQLQ & " AND (DA_ORG IS NULL OR DA_ORG='') "
    Else
        SQLQ = SQLQ & " AND DA_ORG = '" & xORG & "'"
    End If
    If Len(xEMP) = 0 Then
        SQLQ = SQLQ & " AND (DA_EMP IS NULL OR DA_EMP='')"
    Else
        SQLQ = SQLQ & " AND DA_EMP = '" & xEMP & "'"
    End If
    If Len(xPT) = 0 Then
        SQLQ = SQLQ & " AND (DA_PT IS NULL OR DA_PT='')"
    Else
        SQLQ = SQLQ & " AND DA_PT = '" & xPT & "' "
    End If
    If Len(xEmpExclude) = 0 Then
        SQLQ = SQLQ & " AND (DA_EMPEXCL IS NULL OR DA_EMPEXCL='')"
    Else
        SQLQ = SQLQ & " AND DA_EMPEXCL = '" & xEmpExclude & "'"
    End If
    
    rsDailyAcc.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsDailyAcc.EOF Then
        Get_DailyAccrual = IIf(IsNull(rsDailyAcc("ACCTODATE")), 0, rsDailyAcc("ACCTODATE"))
    Else
        Get_DailyAccrual = 0
    End If
    
    rsDailyAcc.Close
    Set rsDailyAcc = Nothing
    
End Function

Public Function DailyVacUpdatedAlready(xDate)
    Dim SQLQ As String
    Dim rsHRPARCO As New ADODB.Recordset
    
    'Check if the Daily Vacation Entitlement has already been updated
    SQLQ = "SELECT PC_LST_DAILYVAC_UPD_DATE FROM HRPARCO"
    rsHRPARCO.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHRPARCO.EOF Then
        If IsNull(rsHRPARCO("PC_LST_DAILYVAC_UPD_DATE")) Or rsHRPARCO("PC_LST_DAILYVAC_UPD_DATE") = "" Then
            DailyVacUpdatedAlready = False
        Else
            If CVDate(rsHRPARCO("PC_LST_DAILYVAC_UPD_DATE")) >= CVDate(Date) Then
                DailyVacUpdatedAlready = True
            Else
                DailyVacUpdatedAlready = False
            End If
        End If
    Else
        DailyVacUpdatedAlready = False
    End If
    rsHRPARCO.Close
    Set rsHRPARCO = Nothing
End Function

Public Function Get_AnnualVac_From_DailyAccrual(xEmpnbr, xToDate)
    Dim rsDailyAcc As New ADODB.Recordset
    Dim SQLQ As String

    'Return the Annual Vacation Accrual
    Get_AnnualVac_From_DailyAccrual = 0
    
    'Not valid entitlement End Date, exit function
    If Not IsDate(xToDate) Then
        Get_AnnualVac_From_DailyAccrual = 0
        Exit Function
    End If
    
    'Retrieve the accrual of the day
    SQLQ = "SELECT DA_ACCRDTODATE FROM HR_DAILYVACACCR WHERE DA_EMPNBR = " & xEmpnbr
    SQLQ = SQLQ & " AND DA_TODATE = " & Date_SQL(xToDate)
    SQLQ = SQLQ & " AND DA_ACCRDATE = " & Date_SQL(xToDate)
    rsDailyAcc.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsDailyAcc.EOF Then
        Get_AnnualVac_From_DailyAccrual = IIf(IsNull(rsDailyAcc("DA_ACCRDTODATE")), 0, rsDailyAcc("DA_ACCRDTODATE"))
    Else
        Get_AnnualVac_From_DailyAccrual = 0
    End If
    
    rsDailyAcc.Close
    Set rsDailyAcc = Nothing

End Function

Public Function Get_Employees_DailyEntitlement_Rule(xEmpnbr, Optional xFromDate, Optional xToDate)
    Dim rsVT As New ADODB.Recordset
    Dim rsHREmp As New ADODB.Recordset
    Dim SQLQ As String
    Dim SQLQV As String
    Dim xORG, xLoc, xEMP, xPT, xEmpExcl
    
    Get_Employees_DailyEntitlement_Rule = ""
    SQLQV = ""
    
    'Retrieve Daily Entitlement Rules
    SQLQ = "SELECT VD_ORG,VD_EMP,VD_PT,VD_EMPEXCL,VD_FRDATE,VD_TODATE "
    SQLQ = SQLQ & " FROM HRVACENTDAILY "
    If IsDate(xFromDate) And IsDate(xToDate) Then
        SQLQ = SQLQ & " WHERE VD_FRDATE=" & Date_SQL(xFromDate)
        SQLQ = SQLQ & " AND VD_TODATE=" & Date_SQL(xToDate)
    End If
    SQLQ = SQLQ & " GROUP BY VD_ORG,VD_EMP,VD_PT,VD_EMPEXCL,VD_FRDATE,VD_TODATE "
    rsVT.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsVT.EOF Then
        'For each rule see if this employee belong to this rule
        Do While Not rsVT.EOF
            xORG = rsVT("VD_ORG") & ""
            xEMP = rsVT("VD_EMP") & ""
            xPT = rsVT("VD_PT") & ""
            xEmpExcl = rsVT("VD_EMPEXCL") & ""
            
            'Check if Employee belong to this rule
            SQLQ = "SELECT ED_EMPNBR FROM HREMP "
            SQLQ = SQLQ & " WHERE ED_EMPNBR = " & xEmpnbr
            SQLQ = SQLQ & " AND ED_EFDATE=" & Date_SQL(rsVT("VD_FRDATE"))
            SQLQ = SQLQ & " AND ED_ETDATE=" & Date_SQL(rsVT("VD_TODATE"))
            If Len(xORG) > 0 Then SQLQ = SQLQ & " AND ED_ORG = '" & xORG & "'"
            If Len(xEMP) > 0 Then SQLQ = SQLQ & " AND ED_EMP = '" & xEMP & "'"
            If Len(xPT) > 0 Then SQLQ = SQLQ & " AND ED_PT = '" & xPT & "'"
            If Len(xEmpExcl) > 0 Then SQLQ = SQLQ & " AND ED_EMP NOT IN ('" & Replace(xEmpExcl, ",", "','") & "')"
            rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsHREmp.EOF Then
                'Rule found, return this rule
                SQLQV = " (DA_ORG = '" & xORG & "' " & IIf(Len(xORG) = 0, " OR DA_ORG IS NULL ", "") & ")"
                SQLQV = SQLQV & " AND (DA_EMP = '" & xEMP & "' " & IIf(Len(xEMP) = 0, " OR DA_EMP IS NULL ", "") & ")"
                SQLQV = SQLQV & " AND (DA_PT = '" & xPT & "' " & IIf(Len(xPT) = 0, " OR DA_PT IS NULL ", "") & ")"
                SQLQV = SQLQV & " AND (DA_EMPEXCL = '" & xEmpExcl & "' " & IIf(Len(xEmpExcl) = 0, " OR DA_EMPEXCL IS NULL ", "") & ")"
                SQLQV = SQLQV & " AND DA_FRDATE = " & Date_SQL(rsVT("VD_FRDATE"))
                SQLQV = SQLQV & " AND DA_TODATE = " & Date_SQL(rsVT("VD_TODATE"))
                
                Get_Employees_DailyEntitlement_Rule = SQLQV
                Exit Do
            Else
                'Rule not found - check next rule
            End If
            rsHREmp.Close
            Set rsHREmp = Nothing
            
            rsVT.MoveNext
        Loop
    Else
        'No rules found for this employee
        Get_Employees_DailyEntitlement_Rule = ""
    End If
    rsVT.Close
    Set rsVT = Nothing
End Function

Public Function Check_Daily_Entitlement_Rule_Exists()
    Dim SQLQ As String
    Dim rsDailyEnt As New ADODB.Recordset

    'Check Daily Entitlement Rules exists
    Screen.MousePointer = HOURGLASS
    
    SQLQ = "SELECT * FROM HRVACENTDAILY "
    rsDailyEnt.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsDailyEnt.BOF And rsDailyEnt.EOF Then
        'Daily Entitlement Rule Do NOT exists
        Check_Daily_Entitlement_Rule_Exists = False
    Else
        'Daily Entitlement Rule exists
        Check_Daily_Entitlement_Rule_Exists = True
    End If
    rsDailyEnt.Close
    Set rsDailyEnt = Nothing
    
    Screen.MousePointer = DEFAULT

End Function

Public Function Total_NonAbsent_Hours(xEmpnbr, xFromDate, xToDate)
    Dim rsAttend As New ADODB.Recordset
    Dim SQLQ As String
    
    Dim xTotHrs As Double
    
    xTotHrs = 0
    
    'Attendance
    SQLQ = "SELECT SUM(AD_HRS) AS TOT_HRS FROM HR_ATTENDANCE"
    SQLQ = SQLQ & " WHERE AD_EMPNBR = " & xEmpnbr
    SQLQ = SQLQ & " AND (AD_DOA >= " & Date_SQL(xFromDate)
    SQLQ = SQLQ & " AND AD_DOA <= " & Date_SQL(xToDate) & ")"
    SQLQ = SQLQ & " AND AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE = 0)"
    SQLQ = SQLQ & " GROUP BY AD_EMPNBR"
    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsAttend.EOF Then
        rsAttend.MoveFirst
                
        'Sum Total Hours
        If rsAttend("TOT_HRS") > 0 Then
            xTotHrs = xTotHrs + rsAttend("TOT_HRS")
        End If
    End If
    rsAttend.Close
    Set rsAttend = Nothing
    
    'Attendance History
    SQLQ = "SELECT SUM(AH_HRS) AS TOT_HRS FROM HR_ATTENDANCE_HISTORY"
    SQLQ = SQLQ & " WHERE AH_EMPNBR = " & xEmpnbr
    SQLQ = SQLQ & " AND (AH_DOA >= " & Date_SQL(xFromDate)
    SQLQ = SQLQ & " AND AH_DOA <= " & Date_SQL(xToDate) & ")"
    SQLQ = SQLQ & " AND AH_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE = 0)"
    SQLQ = SQLQ & " GROUP BY AH_EMPNBR"
    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsAttend.EOF Then
        rsAttend.MoveFirst
                
        'Sum Total Hours
        If rsAttend("TOT_HRS") > 0 Then
            xTotHrs = xTotHrs + rsAttend("TOT_HRS")
        End If
            
    End If
    rsAttend.Close
    Set rsAttend = Nothing

    Total_NonAbsent_Hours = xTotHrs

End Function


