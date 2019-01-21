VERSION 5.00
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmVacEarnedCalc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vacation Earned Calculation"
   ClientHeight    =   2640
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6495
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6255
      Begin VB.TextBox txtWeek 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2550
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   1
         Left            =   3960
         TabIndex        =   4
         Tag             =   "40-Date upto and including this date forward"
         Top             =   1185
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
         Enabled         =   0   'False
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   0
         Left            =   2235
         TabIndex        =   5
         Tag             =   "40-Date from and including this date forward"
         Top             =   1185
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
         Enabled         =   0   'False
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   2190
         Picture         =   "frmVacEarnedCalc.frx":0000
         Top             =   720
         Width           =   240
      End
      Begin VB.Label lblWeek 
         Caption         =   "Pay Period #"
         Height          =   195
         Left            =   720
         TabIndex        =   8
         Top             =   765
         Width           =   1395
      End
      Begin VB.Label lblDate 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Range"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   7
         Top             =   1230
         Width           =   870
      End
      Begin VB.Label lblDateRange 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Vacation Entitlement Period: "
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   270
         Visible         =   0   'False
         Width           =   6000
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Update"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1867
      TabIndex        =   0
      Top             =   2160
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3427
      TabIndex        =   1
      Top             =   2160
      Width           =   1200
   End
End
Attribute VB_Name = "frmVacEarnedCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xVacFromDate
Dim xVacToDate


'Public Property Get ChangeAction() As UpdateStateEnum
'ChangeAction = OPENING
'End Property
'
'Public Property Get RelateMode() As RelateModeEnum
''RelateMode = Reports
''RelateMode = RelateEMP
'RelateMode = MassChanges
'End Property
'
'Public Property Get UpdateRight() As Boolean
'UpdateRight = False 'True  'False
'End Property
'
'Public Property Get Addable() As Boolean
'Addable = False 'True  'False
'End Property
'
'Public Property Get Updateble() As Boolean
'Updateble = True
'End Property
'
'Public Property Get Deleteble() As Boolean
'Deleteble = False
'End Property
'
'Public Property Get Printable() As Boolean
'Printable = False 'True
'End Property

'Private Sub cmbMonth_Change()
'    If cmbMonth.ListIndex = 1 Then
'        dlpDateRange(0).Text = CVDate(Format("07/01/" & Year(Date) - 1, "mm/dd/yyyy"))
'        dlpDateRange(1).Text = CVDate(Format("12/31/" & Year(Date) - 1, "mm/dd/yyyy"))
'    ElseIf cmbMonth.ListIndex = 2 Then
'        dlpDateRange(0).Text = CVDate(Format("01/01/" & Year(Date), "mm/dd/yyyy"))
'        dlpDateRange(1).Text = CVDate(Format("06/30/" & Year(Date), "mm/dd/yyyy"))
'    Else
'        dlpDateRange(0).Text = ""
'        dlpDateRange(1).Text = ""
'    End If
'End Sub

'Private Sub cmbMonth_Click()
'    If cmbMonth.ListIndex = 1 Then
'        dlpDateRange(0).Text = CVDate(Format("07/01/" & Year(Date) - 1, "mm/dd/yyyy"))
'        dlpDateRange(1).Text = CVDate(Format("12/31/" & Year(Date) - 1, "mm/dd/yyyy"))
'    ElseIf cmbMonth.ListIndex = 2 Then
'        dlpDateRange(0).Text = CVDate(Format("01/01/" & Year(Date), "mm/dd/yyyy"))
'        dlpDateRange(1).Text = CVDate(Format("06/30/" & Year(Date), "mm/dd/yyyy"))
'    Else
'        dlpDateRange(0).Text = ""
'        dlpDateRange(1).Text = ""
'    End If
'End Sub
'
'Private Sub cmbMonth_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Private Sub cmdCancel_Click()
    'glbSenMonth = 999
    Unload Me
End Sub

Sub cmdOK_Click()
Dim Response%

'Verify the Pay Period selected is within the Vacation Entitlement Period
If Len(Trim(txtWeek.Text)) = 0 Then
    MsgBox "Pay Period # cannot be blank. Please select valid Pay Period.", vbExclamation
    txtWeek.SetFocus
    Exit Sub
End If

If IsDate(dlpDateRange(0)) And IsDate(dlpDateRange(1)) Then
    'Check if valid Pay Period selected for the displayed Vacation Entitlement Period
    If IsDate(xVacFromDate) And IsDate(xVacToDate) Then
        If CVDate(dlpDateRange(0)) >= CVDate(xVacFromDate) And CVDate(dlpDateRange(1)) <= CVDate(xVacToDate) Then
            'Pay Period Date range within Vacation Entitlement Period
        Else
            MsgBox "Pay Period Date Range outside the Vacation Entitlement Period. Please select valid Pay Period.", vbExclamation
            txtWeek.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Invalid Vacation Entitlement Period to compute the Vacation Earned Hours. Please verify the Vacation Entitlement Master screen for the Entitlement Period.", vbExclamation
        Exit Sub
    End If
Else
    MsgBox "Invalid Pay Period Date Range to compute the Vacation Earned Hours. Please select valid Pay Period.", vbExclamation
    dlpDateRange(0).SetFocus
    Exit Sub
End If

'Call procedure to calculate the Vacation for everyone
Call OshawaPL_Vacation_Update

'glbSenMonth = 0

'If cmbMonth.ListIndex <> 0 And cmbMonth.ListIndex <> -1 Then
'    'Prompt to confirm
'    Response% = MsgBox("Seniority Date will be computed for the period: " & vbCrLf & dlpDateRange(0).Text & " to " & dlpDateRange(1).Text & "." & vbCrLf & vbCrLf & "Do you wish to proceed?", vbYesNo + vbQuestion, "info:HR - Seniority Date Calculation")
'
'    If Response% = IDNO Then
'        cmbMonth.SetFocus
'        Exit Sub
'    End If
    
'    Call Compute_SeniorityDate
'Else
'    MsgBox "No Month selected to update the Seniority Date for." & vbCrLf & "Please select the Month to compute the employee's Seniority Date.", vbExclamation, "Select 'For the Month'"
'    cmbMonth.SetFocus
'    Exit Sub
'End If

Screen.MousePointer = HOURGLASS

'glbSenMonth = cmbMonth.ListIndex

Screen.MousePointer = DEFAULT

Unload Me

End Sub

Private Sub VacationEntitlementPeriod()
    Dim rsVacEnt As New ADODB.Recordset
    Dim SQLQ As String
    
    'Initialise
    xVacFromDate = ""
    xVacToDate = ""
    
    SQLQ = "SELECT * FROM HRVACENT ORDER BY VE_FRDATE DESC, VE_TODATE DESC"
    rsVacEnt.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsVacEnt.EOF Then
        rsVacEnt.MoveFirst
        
        lblDateRange.Caption = "Vacation Entitlement Period: " & rsVacEnt("VE_FRDATE") & " - " & rsVacEnt("VE_TODATE")
        lblDateRange.Visible = True
        'Store locally Vacation Entitlement period for later use
        xVacFromDate = rsVacEnt("VE_FRDATE")
        xVacToDate = rsVacEnt("VE_TODATE")
    Else
        'Clear Vacation Entitlement Period displayed and stored locally
        lblDateRange.Caption = "Vacation Entitlement Period: "
        
        xVacFromDate = ""
        xVacToDate = ""
    End If
    rsVacEnt.Close
    Set rsVacEnt = Nothing
End Sub

Private Sub Form_Activate()
glbOnTop = "frmVacEarnedCalc"
End Sub

Private Sub Form_Load()
glbOnTop = "frmVacEarnedCalc"

'Retrieve Vacation Entitlement Period
Call VacationEntitlementPeriod

End Sub

Private Sub Compute_SeniorityDate()
    Dim rsHREmp As New ADODB.Recordset
    Dim rsAttend As New ADODB.Recordset
    Dim rsAudit As New ADODB.Recordset
    Dim SQLQ As String
    
    Dim xPT As String
    Dim xDiv As String
    Dim xTotHrsWks As Double
    Dim xTotHrsDays As Double
    Dim xNewSenDate As Date
    
    'CUPE employees only
    SQLQ = "SELECT ED_EMPNBR, ED_SENDTE, ED_PT, ED_DIV, ED_LDATE, ED_LTIME, ED_LUSER FROM HREMP WHERE ED_ORG = 'CUPE'"
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHREmp.EOF Then
        'For each employee get the total hours grouped by employee's Hours/Week
        'By Reason Codes LOA and USL with Union CUPE. And
        'By Union <> CUPE
        rsHREmp.MoveFirst
        
        Do While Not rsHREmp.EOF
            xTotHrsWks = 0
            
            'Make sure Seniority Date is valid date
            If IsDate(rsHREmp("ED_SENDTE")) Then
                'Attendance
                SQLQ = "SELECT AD_EMPNBR, AD_WHRS, SUM(AD_HRS) AS TOT_HRS FROM HR_ATTENDANCE"
                SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR")
                SQLQ = SQLQ & " AND (((AD_DOA >= " & Date_SQL(dlpDateRange(0).Text)
                SQLQ = SQLQ & " AND AD_DOA <= " & Date_SQL(dlpDateRange(1).Text) & ")"
                SQLQ = SQLQ & " AND (AD_REASON IN ('LOA','USL') AND AD_ORG = 'CUPE'))"
                SQLQ = SQLQ & " OR ((AD_DOA >= " & Date_SQL(dlpDateRange(0).Text)
                SQLQ = SQLQ & " AND AD_DOA <= " & Date_SQL(dlpDateRange(1).Text) & ")"
                SQLQ = SQLQ & " AND (AD_ORG <> 'CUPE')))"
                SQLQ = SQLQ & " GROUP BY AD_EMPNBR, AD_WHRS"
                SQLQ = SQLQ & " ORDER BY AD_WHRS"
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsAttend.EOF Then
                    rsAttend.MoveFirst
                            
                    'For each Total Hours by Hours/Week compute Total Hours in Weeks
                    Do While Not rsAttend.EOF
                        If rsAttend("TOT_HRS") > 0 Then
                            xTotHrsWks = xTotHrsWks + (rsAttend("TOT_HRS") / rsAttend("AD_WHRS"))
                        End If
                        
                        rsAttend.MoveNext
                    Loop
                End If
                rsAttend.Close
                Set rsAttend = Nothing
                
                'Attendance History
                SQLQ = "SELECT AH_EMPNBR, AH_WHRS, SUM(AH_HRS) AS TOT_HRS FROM HR_ATTENDANCE_HISTORY"
                SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR")
                SQLQ = SQLQ & " AND (((AH_DOA >= " & Date_SQL(dlpDateRange(0).Text)
                SQLQ = SQLQ & " AND AH_DOA <= " & Date_SQL(dlpDateRange(1).Text) & ")"
                SQLQ = SQLQ & " AND (AH_REASON IN ('LOA','USL') AND AH_ORG = 'CUPE'))"
                SQLQ = SQLQ & " OR ((AH_DOA >= " & Date_SQL(dlpDateRange(0).Text)
                SQLQ = SQLQ & " AND AH_DOA <= " & Date_SQL(dlpDateRange(1).Text) & ")"
                SQLQ = SQLQ & " AND (AH_ORG <> 'CUPE')))"
                SQLQ = SQLQ & " GROUP BY AH_EMPNBR, AH_WHRS"
                SQLQ = SQLQ & " ORDER BY AH_WHRS"
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsAttend.EOF Then
                    rsAttend.MoveFirst
                            
                    'For each Total Hours by Hours/Week compute Total Hours in Weeks
                    Do While Not rsAttend.EOF
                        If rsAttend("TOT_HRS") > 0 Then
                            xTotHrsWks = xTotHrsWks + (rsAttend("TOT_HRS") / rsAttend("AH_WHRS"))
                        End If
                        
                        rsAttend.MoveNext
                    Loop
                End If
                rsAttend.Close
                Set rsAttend = Nothing
            
            
                'Compute Seniority Date by Work Weeks, excluding Saturday/Sunday and Holidays
                'so, adding Total Hours in Weeks (xTotHrsHWks) to Seniority Date
                If xTotHrsWks > 0 Then
                    'Convert # of Weeks to # of Days, so we can exclude non work days when adding to Seniority Date
                    xTotHrsDays = xTotHrsWks * 5
                    
                    'Add # of Total Hours in Days to the Seniority excluding Weekends and Statutory Holidays
                    xNewSenDate = AddWorkingDays(rsHREmp("ED_SENDTE"), xTotHrsDays, True)
                    
                    'Update Seniority Date in HREMP if the date has changed
                    If CVDate(rsHREmp("ED_SENDTE")) <> CVDate(xNewSenDate) Then
                        rsHREmp("ED_SENDTE") = xNewSenDate
                        rsHREmp("ED_LDATE") = Now
                        rsHREmp("ED_LTIME") = Time$
                        rsHREmp("ED_LUSER") = glbLEE_ID
                        rsHREmp.Update
                        
                        'Retrieve PT and Div from HREMP
                        If IsNull(rsHREmp("ED_PT")) Then xPT = "" Else xPT = rsHREmp("ED_PT")
                        If IsNull(rsHREmp("ED_DIV")) Then xDiv = "" Else xDiv = rsHREmp("ED_DIV")
                                            
                        'Add Audit Log
                        rsAudit.Open "SELECT * FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
                        
                        rsAudit.AddNew
                        rsAudit("AU_LOC_TABL") = "EDLC": rsAudit("AU_SECTION_TABL") = "EDSE": rsAudit("AU_EMP_TABL") = "EDEM": rsAudit("AU_SUPCODE_TABL") = "EDSP": rsAudit("AU_ORG_TABL") = "EDOR": rsAudit("AU_PAYP_TABL") = "SDPP": rsAudit("AU_BCODE_TABL") = "BNCD": rsAudit("AU_TREAS_TABL") = "TERM": rsAudit("AU_DOLENT_TABL") = "EDOL": rsAudit("AU_EARN_TABL") = "EARN"
                        rsAudit("AU_ADMINBY_TABL") = "EDAB": rsAudit("AU_LANG1_TABL") = "EDL1": rsAudit("AU_LANG2_TABL") = "EDL1"
                        
                        rsAudit("AU_NEWEMP") = "N"
                        rsAudit("AU_PTUPL") = xPT
                        rsAudit("AU_DIVUPL") = xDiv
                        rsAudit("AU_COMPNO") = "001"
                        rsAudit("AU_EMPNBR") = rsHREmp("ED_EMPNBR")
                                
                        If IsDate(xNewSenDate) Then
                            rsAudit("AU_SENDTE") = CVDate(xNewSenDate)
                        End If
                        
                        rsAudit("AU_LDATE") = Date
                        rsAudit("AU_LUSER") = glbLEE_ID
                        rsAudit("AU_LTIME") = Time$
                        rsAudit("AU_UPLOAD") = "N"
                        rsAudit("AU_TYPE") = "M"
                        
                        rsAudit.Update
                        rsAudit.Close
                        Set rsAudit = Nothing
                    End If
                End If
            End If
            rsHREmp.MoveNext
        Loop
        MsgBox "Employees Seniority Date updated successfully.", vbOKOnly, "Seniority Date Calculation"
    Else
        MsgBox "No Employees to update.", vbOKOnly, "Seniority Date Calculation"
    End If
    rsHREmp.Close
    Set rsHREmp = Nothing
End Sub

Private Sub txtWeek_Change()
Dim DateRange
DateRange = Split(getDateRange("", txtWeek, Year(Now)), "|")
    dlpDateRange(0) = DateRange(0)
    dlpDateRange(1) = DateRange(1)
End Sub

Private Sub txtWeek_DblClick()
Call imgIcon_Click
End Sub

Private Sub imgIcon_Click()
frmPayPeriodList.SelectedYear = Year(Now) 'Val(txtYear)
'frmPayPeriodList.PayPeriodCode = clpPayP.Text
frmPayPeriodList.Show 1
txtWeek = glbWeek
dlpDateRange(0) = glbFrom
dlpDateRange(1) = glbTo
End Sub

Private Sub txtWeek_LostFocus()
    If txtWeek = "" Then
        dlpDateRange(0) = ""
        dlpDateRange(1) = ""
    Else
        'FIND THE DATA RANGE FROM THE DATABASE FOR THAT WEEK #
    End If
End Sub

'Private Sub txtYear_Change()
'Dim DateRange
'DateRange = Split(getDateRange("", txtWeek, txtYear), "|")
'    dlpDateRange(0) = DateRange(0)
'    dlpDateRange(1) = DateRange(1)
'End Sub
'
Private Function getDateRange(theClientNumber, thePayNbr, theYear)
Dim rsPayPeriod As New ADODB.Recordset
Dim SQLQ, intNum

On Error Resume Next

getDateRange = "|"

If Not IsNumeric(thePayNbr) Then Exit Function
If Not IsNumeric(theYear) Then Exit Function

If Len(theClientNumber) > 0 Then
    SQLQ = "SELECT PP_NBR,PP_YEAR,PP_Start,PP_End FROM HR_PAYPERIOD WHERE PP_PAYP='" & theClientNumber & "'"
Else
    SQLQ = "SELECT PP_NBR,PP_YEAR,PP_Start,PP_End FROM HR_PAYPERIOD WHERE 1 = 1 "
End If
SQLQ = SQLQ & " AND PP_NBR = " & thePayNbr
SQLQ = SQLQ & " AND PP_YEAR = '" & theYear & "'"
rsPayPeriod.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly

If Not rsPayPeriod.EOF Then
    getDateRange = rsPayPeriod("PP_Start") & "|" & rsPayPeriod("PP_End")
End If
rsPayPeriod.Close
Exit Function

End Function

Private Function OshawaPL_Vacation_Update()
    Dim rsHREmp As New ADODB.Recordset
    Dim rsAttend As New ADODB.Recordset
    Dim rsAttPP As New ADODB.Recordset
    Dim SQLQ As String
    Dim xVacEarned As Double
    Dim xFTHsWorked As Double
    Dim xVacEarnedPT As Double
    Dim xVacEarnedFT As Double
    
    
    'For Category = PT employees
    'Get the Total Seniority Hours from HR_ATTENDANCE and HR_ATTENDANCE_HISTORY table
    SQLQ = "SELECT EMPNBR, SUM(TOT_SEN_HRS) AS TOT_SEN_HRS FROM "
    SQLQ = SQLQ & " (SELECT AD_EMPNBR AS EMPNBR, SUM(AD_HRS) AS TOT_SEN_HRS FROM HR_ATTENDANCE WHERE"
    SQLQ = SQLQ & " AD_SEN<>0 "
    SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_PT = 'PT')"
    SQLQ = SQLQ & " GROUP BY AD_EMPNBR"
    SQLQ = SQLQ & " UNION "
    SQLQ = SQLQ & " SELECT AH_EMPNBR AS EMPNBR, SUM(AH_HRS) AS TOT_SEN_HRS FROM HR_ATTENDANCE_HISTORY WHERE"
    SQLQ = SQLQ & " AH_SEN<>0 "
    SQLQ = SQLQ & " AND AH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_PT = 'PT')"
    SQLQ = SQLQ & " GROUP BY AH_EMPNBR) AS HR_ATTENDANCE"
    SQLQ = SQLQ & " GROUP BY EMPNBR"
    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If Not rsAttend.EOF Then
        rsAttend.MoveFirst
                
        Do While Not rsAttend.EOF
            'Initialise
            xVacEarned = 0
            
            'Calculate employee's Seniority Hours for the Pay Period
            SQLQ = "SELECT EMPNBR, SUM(PP_SEN_HRS) AS PP_SEN_HRS FROM "
            SQLQ = SQLQ & " (SELECT AD_EMPNBR AS EMPNBR, SUM(AD_HRS) AS PP_SEN_HRS FROM HR_ATTENDANCE WHERE"
            SQLQ = SQLQ & " AD_SEN<>0 AND AD_EMPNBR = " & rsAttend("EMPNBR")
            SQLQ = SQLQ & " AND AD_DOA >= " & Date_SQL(dlpDateRange(0)) & " AND AD_DOA <= " & Date_SQL(dlpDateRange(1))
            SQLQ = SQLQ & " GROUP BY AD_EMPNBR"
            SQLQ = SQLQ & " UNION "
            SQLQ = SQLQ & " SELECT AH_EMPNBR AS EMPNBR, SUM(AH_HRS) AS PP_SEN_HRS FROM HR_ATTENDANCE_HISTORY WHERE"
            SQLQ = SQLQ & " AH_SEN<>0 AND AH_EMPNBR = " & rsAttend("EMPNBR")
            SQLQ = SQLQ & " AND AH_DOA >= " & Date_SQL(dlpDateRange(0)) & " AND AH_DOA <= " & Date_SQL(dlpDateRange(1))
            SQLQ = SQLQ & " GROUP BY AH_EMPNBR) AS HR_ATTENDANCE"
            SQLQ = SQLQ & " GROUP BY EMPNBR"
            rsAttPP.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsAttPP.EOF Then
            
                'Compute Vacation Earned Hours based on employee's Total Seniority Hours and Pay Period Hours
                '<=9100 then Vacatio Earned Hours = 105/1820 * Pay Period Hours
                If rsAttend("TOT_SEN_HRS") <= 9100 Then
                    xVacEarned = (105 / 1820) * rsAttPP("PP_SEN_HRS")
                Else
                    '>9100 then Vacatio Earned Hours = 140/1820 * Pay Period Hours
                    xVacEarned = (140 / 1820) * rsAttPP("PP_SEN_HRS")
                End If
                
                'Update Employee's Vacation by Vacation Earned based on the Pay Period and Seniority Hours
                SQLQ = "SELECT ED_EMPNBR, ED_VAC, ED_LDATE, ED_LTIME, ED_LUSER FROM HREMP WHERE ED_EMPNBR = " & rsAttend("EMPNBR")
                rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsHREmp.EOF Then
                    If IsNumeric(rsHREmp("ED_VAC")) Then
                        rsHREmp("ED_VAC") = rsHREmp("ED_VAC") + xVacEarned
                    Else
                        rsHREmp("ED_VAC") = xVacEarned
                    End If
                    rsHREmp("ED_LDATE") = Now
                    rsHREmp("ED_LTIME") = Time$
                    rsHREmp("ED_LUSER") = glbUserID
                    rsHREmp.Update
                End If
                rsHREmp.Close
                Set rsHREmp = Nothing
            End If
            rsAttPP.Close
            Set rsAttPP = Nothing
            
            rsAttend.MoveNext
        Loop
    End If
    rsAttend.Close
    Set rsAttend = Nothing
    
    'Initialise
    xFTHsWorked = 70
    
    'For Category = TFT employees
    SQLQ = "SELECT ED_EMPNBR, ED_VAC, ED_VACPC, ED_LDATE, ED_LTIME, ED_LUSER FROM HREMP WHERE ED_PT = 'TFT'"
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsHREmp.EOF
        'Initialise
        xVacEarnedPT = 0
        xVacEarnedFT = 0
        
        If IsNumeric(rsHREmp("ED_VACPC")) Then
            xVacEarnedPT = 50 * rsHREmp("ED_VACPC")
            xVacEarnedFT = 20 * 0.04
            
            If IsNumeric(rsHREmp("ED_VAC")) Then
                rsHREmp("ED_VAC") = rsHREmp("ED_VAC") + xVacEarnedPT + xVacEarnedFT
            Else
                rsHREmp("ED_VAC") = xVacEarnedPT + xVacEarnedFT
            End If
            rsHREmp("ED_LDATE") = Now
            rsHREmp("ED_LTIME") = Time$
            rsHREmp("ED_LUSER") = glbUserID
            rsHREmp.Update
        Else
            'xVacEarnedFT = 20 * 0.04
            
            'If IsNumeric(rsHREmp("ED_VAC")) Then
            '    rsHREmp("ED_VAC") = rsHREmp("ED_VAC") + xVacEarnedPT
            'Else
            '    rsHREmp("ED_VAC") = xVacEarnedPT
            'End If
            'rsHREmp("ED_LDATE") = Now
            'rsHREmp("ED_LTIME") = Time$
            'rsHREmp("ED_LUSER") = glbUserID
            'rsHREmp.Update
        End If
        rsHREmp.MoveNext
    Loop
    rsHREmp.Close
    Set rsHREmp = Nothing
    
    
End Function

