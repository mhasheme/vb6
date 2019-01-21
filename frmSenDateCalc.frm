VERSION 5.00
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmSenDateCalc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seniority Date Calculation"
   ClientHeight    =   2640
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6255
      Begin VB.ComboBox cmbMonth 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2925
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Tag             =   "Select Seniority Month"
         Top             =   510
         Width           =   1725
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   1
         Left            =   4320
         TabIndex        =   7
         Tag             =   "40-Date upto and including this date forward"
         Top             =   945
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
         Left            =   2600
         TabIndex        =   8
         Tag             =   "40-Date from and including this date forward"
         Top             =   945
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
         Enabled         =   0   'False
      End
      Begin VB.Label lblDateRange 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Date"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2925
         TabIndex        =   6
         Top             =   150
         Visible         =   0   'False
         Width           =   1020
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
         Left            =   1080
         TabIndex        =   5
         Top             =   990
         Width           =   870
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "For the Month"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   3
         Tag             =   "41-Date Terminated"
         Top             =   570
         Width           =   1200
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
Attribute VB_Name = "frmSenDateCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Sub cmbMonth_Change()
    If cmbMonth.ListIndex = 1 Then
        dlpDateRange(0).Text = CVDate(Format("07/01/" & Year(Date) - 1, "mm/dd/yyyy"))
        dlpDateRange(1).Text = CVDate(Format("12/31/" & Year(Date) - 1, "mm/dd/yyyy"))
    ElseIf cmbMonth.ListIndex = 2 Then
        dlpDateRange(0).Text = CVDate(Format("01/01/" & Year(Date), "mm/dd/yyyy"))
        dlpDateRange(1).Text = CVDate(Format("06/30/" & Year(Date), "mm/dd/yyyy"))
    Else
        dlpDateRange(0).Text = ""
        dlpDateRange(1).Text = ""
    End If
End Sub

Private Sub cmbMonth_Click()
    If cmbMonth.ListIndex = 1 Then
        dlpDateRange(0).Text = CVDate(Format("07/01/" & Year(Date) - 1, "mm/dd/yyyy"))
        dlpDateRange(1).Text = CVDate(Format("12/31/" & Year(Date) - 1, "mm/dd/yyyy"))
    ElseIf cmbMonth.ListIndex = 2 Then
        dlpDateRange(0).Text = CVDate(Format("01/01/" & Year(Date), "mm/dd/yyyy"))
        dlpDateRange(1).Text = CVDate(Format("06/30/" & Year(Date), "mm/dd/yyyy"))
    Else
        dlpDateRange(0).Text = ""
        dlpDateRange(1).Text = ""
    End If
End Sub

Private Sub cmbMonth_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdCancel_Click()
    glbSenMonth = 999
    Unload Me
End Sub

Sub cmdOK_Click()
Dim Response%

glbSenMonth = 0

If cmbMonth.ListIndex <> 0 And cmbMonth.ListIndex <> -1 Then
    'Prompt to confirm
    Response% = MsgBox("Seniority Date will be computed for the period: " & vbCrLf & dlpDateRange(0).Text & " to " & dlpDateRange(1).Text & "." & vbCrLf & vbCrLf & "Do you wish to proceed?", vbYesNo + vbQuestion, "info:HR - Seniority Date Calculation")
    
    If Response% = IDNO Then
        cmbMonth.SetFocus
        Exit Sub
    End If
    
    'Ticket #26514- The logic changed so created a new procedure to compute the Seniority Date.
    'Call Compute_SeniorityDate
    Call Compute_SeniorityDate_NewLogic
    
Else
    MsgBox "No Month selected to update the Seniority Date for." & vbCrLf & "Please select the Month to compute the employee's Seniority Date.", vbExclamation, "Select 'For the Month'"
    cmbMonth.SetFocus
    Exit Sub
End If

Screen.MousePointer = HOURGLASS

glbSenMonth = cmbMonth.ListIndex

Screen.MousePointer = DEFAULT

Unload Me

End Sub

Private Sub PopulateMonths()
    cmbMonth.Clear
    cmbMonth.AddItem ""
    cmbMonth.AddItem "January"
    cmbMonth.AddItem "July"
End Sub

Private Sub Form_Activate()
glbOnTop = "frmSenDateCalc"
End Sub

Private Sub Form_Load()
glbOnTop = "frmSenDateCalc"

Call PopulateMonths

End Sub

Private Sub Compute_SeniorityDate_Old()
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

'Private Sub Compute_SeniorityDate_New()
'    Dim rsHREmp As New ADODB.Recordset
'    Dim rsAttend As New ADODB.Recordset
'    Dim rsAudit As New ADODB.Recordset
'    Dim SQLQ As String
'
'    Dim xPT As String
'    Dim xDiv As String
'    Dim xTotHrs As Double
'    Dim xTotHrsDays As Double
'    Dim xNewSenDate As Date
'
'    Dim xStdHrsWeek
'    Dim xHrsWeek
'    Dim xHrsDay
'    Dim x6MnthsWeek
'    Dim xHrsNotWorked
'    Dim xTotPDWHrs
'
'    'Initialise
'    xStdHrsWeek = 35
'    xHrsDay = 7
'    x6MnthsWeek = 26
'    xHrsNotWorked = 0
'    xTotPDWHrs = 0
'
'    'CUPE employees only
'    SQLQ = "SELECT ED_EMPNBR, ED_SENDTE, ED_PT, ED_DIV, ED_LDATE, ED_LTIME, ED_LUSER FROM HREMP WHERE ED_ORG = 'CUPE'"
'    SQLQ = SQLQ & " AND ED_SENDTE IS NOT NULL"
'    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    If Not rsHREmp.EOF Then
'        'For each employee get the total hours grouped by employee's Hours/Week
'        'By Reason Codes LOA and USL with Union CUPE. And
'        'By Union <> CUPE
'        rsHREmp.MoveFirst
'
'        Do While Not rsHREmp.EOF
'            xTotHrs = 0
'            xHrsNotWorked = 0
'            xTotPDWHrs = 0
'
'            'Make sure Seniority Date is valid date
'            If IsDate(rsHREmp("ED_SENDTE")) Then
'                'Get Total Employee's Hours per Week as they are multi position
'                xHrsWeek = GetJHData_TotWHRS(rsHREmp("ED_EMPNBR"), "JH_WHRS", 0)
'
'                If xHrsWeek = 0 Then GoTo NextEmployee
'
'                'Compute Hours Not Worked
'                xHrsNotWorked = xStdHrsWeek - xHrsWeek
'
'                'Attendance
'                SQLQ = "SELECT AD_EMPNBR, SUM(AD_HRS) AS TOT_HRS FROM HR_ATTENDANCE"
'                SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR")
'                SQLQ = SQLQ & " AND (AD_DOA >= " & Date_SQL(dlpDateRange(0).Text)
'                SQLQ = SQLQ & " AND AD_DOA <= " & Date_SQL(dlpDateRange(1).Text) & ")"
'                'Ticket #29283 - Other sets of codes
'                'Ticket #28921 - Additional codes
'                'SQLQ = SQLQ & " AND (AD_REASON IN ('LOA','USL'))"    ' AND AD_ORG = 'CUPE'))"
'                'SQLQ = SQLQ & " AND (AD_REASON IN ('LOA','USL','LAID','OT','UPL','VCP'))"    ' AND AD_ORG = 'CUPE'))"
'                SQLQ = SQLQ & " AND (AD_REASON IN ('LOA','USL','LAID','UPL','PDNW'))"    ' AND AD_ORG = 'CUPE'))"
'                'SQLQ = SQLQ & " OR ((AD_DOA >= " & Date_SQL(dlpDateRange(0).Text)
'                'SQLQ = SQLQ & " AND AD_DOA <= " & Date_SQL(dlpDateRange(1).Text) & ")"
'                'SQLQ = SQLQ & " AND (AD_ORG <> 'CUPE')))"
'                SQLQ = SQLQ & " GROUP BY AD_EMPNBR"  ', AD_WHRS"
'                'SQLQ = SQLQ & " ORDER BY AD_WHRS"
'                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                If Not rsAttend.EOF Then
'                    rsAttend.MoveFirst
'
'                    'For each Total Hours by Hours/Week compute Total Hours in Weeks
'                    Do While Not rsAttend.EOF
'                        If rsAttend("TOT_HRS") > 0 Then
'                            xTotHrs = xTotHrs + rsAttend("TOT_HRS")   '(rsAttend("TOT_HRS") / rsAttend("AD_WHRS"))
'                        End If
'
'                        rsAttend.MoveNext
'                    Loop
'                End If
'                rsAttend.Close
'                Set rsAttend = Nothing
'
'                'Attendance History
'                SQLQ = "SELECT AH_EMPNBR, SUM(AH_HRS) AS TOT_HRS FROM HR_ATTENDANCE_HISTORY"
'                SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR")
'                SQLQ = SQLQ & " AND (AH_DOA >= " & Date_SQL(dlpDateRange(0).Text)
'                SQLQ = SQLQ & " AND AH_DOA <= " & Date_SQL(dlpDateRange(1).Text) & ")"
'                'Ticket #29283 - Other sets of codes
'                'Ticket #28921 - Additional codes
'                'SQLQ = SQLQ & " AND (AH_REASON IN ('LOA','USL'))"    ' AND AH_ORG = 'CUPE'))"
'                'SQLQ = SQLQ & " AND (AH_REASON IN ('LOA','USL','LAID','OT','UPL','VCP'))"    ' AND AH_ORG = 'CUPE'))"
'                SQLQ = SQLQ & " AND (AH_REASON IN ('LOA','USL','LAID','UPL','PDNW'))"    ' AND AH_ORG = 'CUPE'))"
'                'SQLQ = SQLQ & " OR ((AH_DOA >= " & Date_SQL(dlpDateRange(0).Text)
'                'SQLQ = SQLQ & " AND AH_DOA <= " & Date_SQL(dlpDateRange(1).Text) & ")"
'                'SQLQ = SQLQ & " AND (AH_ORG <> 'CUPE')))"
'                SQLQ = SQLQ & " GROUP BY AH_EMPNBR" ', AH_WHRS"
'                'SQLQ = SQLQ & " ORDER BY AH_WHRS"
'                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                If Not rsAttend.EOF Then
'                    rsAttend.MoveFirst
'
'                    'For each Total Hours by Hours/Week compute Total Hours in Weeks
'                    Do While Not rsAttend.EOF
'                        If rsAttend("TOT_HRS") > 0 Then
'                            xTotHrs = xTotHrs + rsAttend("TOT_HRS")   '(rsAttend("TOT_HRS") / rsAttend("AH_WHRS"))
'                        End If
'
'                        rsAttend.MoveNext
'                    Loop
'                End If
'                rsAttend.Close
'                Set rsAttend = Nothing
'
'
'                'Ticket #29283 - New Step - Subtract PDW Hours of the same period from Total Hours (xTotHrs)
'                'Attendance - PDW
'                SQLQ = "SELECT AD_EMPNBR, SUM(AD_HRS) AS TOT_HRS FROM HR_ATTENDANCE"
'                SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR")
'                SQLQ = SQLQ & " AND (AD_DOA >= " & Date_SQL(dlpDateRange(0).Text)
'                SQLQ = SQLQ & " AND AD_DOA <= " & Date_SQL(dlpDateRange(1).Text) & ")"
'                SQLQ = SQLQ & " AND (AD_REASON IN ('PDW'))"
'                SQLQ = SQLQ & " GROUP BY AD_EMPNBR"
'                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                If Not rsAttend.EOF Then
'                    rsAttend.MoveFirst
'
'                    If rsAttend("TOT_HRS") > 0 Then
'                        xTotPDWHrs = xTotPDWHrs + rsAttend("TOT_HRS")
'                    End If
'                End If
'                rsAttend.Close
'                Set rsAttend = Nothing
'
'                'Attendance History - PDW
'                SQLQ = "SELECT AH_EMPNBR, SUM(AH_HRS) AS TOT_HRS FROM HR_ATTENDANCE_HISTORY"
'                SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR")
'                SQLQ = SQLQ & " AND (AH_DOA >= " & Date_SQL(dlpDateRange(0).Text)
'                SQLQ = SQLQ & " AND AH_DOA <= " & Date_SQL(dlpDateRange(1).Text) & ")"
'                SQLQ = SQLQ & " AND (AH_REASON IN ('PDW'))"
'                SQLQ = SQLQ & " GROUP BY AH_EMPNBR"
'                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                If Not rsAttend.EOF Then
'                    rsAttend.MoveFirst
'
'                    If rsAttend("TOT_HRS") > 0 Then
'                        xTotPDWHrs = xTotPDWHrs + rsAttend("TOT_HRS")
'                    End If
'                End If
'                rsAttend.Close
'                Set rsAttend = Nothing
'
'                'Ticket #29283 - Subtract Total PDW Hours from Total Hours
'                xTotHrs = xTotHrs - xTotPDWHrs
'
'
'                'Compute Hours Not Worked in 6 months - doing this before including the LOA and USL hours because the
'                '26weeks/6months is as per the standard # of hours not worked.
'                xHrsNotWorked = x6MnthsWeek * xHrsNotWorked
'
'                'Compute Total Hours Not Worked with LOA and USL hours included - Now adding to actual hours not worked
'                'for 6 months.
'                xHrsNotWorked = xHrsNotWorked + xTotHrs
'
'
'                'Compute Seniority Date by Work Weeks, excluding Saturday/Sunday - Stat Holidays should be included (as per client)
'                'so, adding Total Hours not worked (xHrsNotWorked) to Seniority Date
'                If xHrsNotWorked > 0 Then
'                    'Convert the Total # of Hours Not Worked to # of Days, so we can exclude non work days when
'                    'adding to Seniority Date
'                    xTotHrsDays = xHrsNotWorked / xHrsDay
'
'                    'Ticket #26514 - Include the Seniority Date itself as Day 1 missed work. Therefore, I am
'                    'subtracting 1 from total days not worked to include teh Seniority Date as Day 1.
'                    'Add # of Total Hours Not Worked in Days to the Seniority excluding Weekends
'                    'They want to include Statutory Holidays
'                    'xNewSenDate = AddWorkingDays(rsHREmp("ED_SENDTE"), xTotHrsDays, False)
'                    xNewSenDate = AddWorkingDays(rsHREmp("ED_SENDTE"), xTotHrsDays - 1, False)
'
'                    'Update Seniority Date in HREMP if the date has changed
'                    If CVDate(rsHREmp("ED_SENDTE")) <> CVDate(xNewSenDate) Then
'                        rsHREmp("ED_SENDTE") = xNewSenDate
'                        rsHREmp("ED_LDATE") = Now
'                        rsHREmp("ED_LTIME") = Time$
'                        rsHREmp("ED_LUSER") = glbUserID
'                        rsHREmp.Update
'
'                        'Retrieve PT and Div from HREMP
'                        If IsNull(rsHREmp("ED_PT")) Then xPT = "" Else xPT = rsHREmp("ED_PT")
'                        If IsNull(rsHREmp("ED_DIV")) Then xDiv = "" Else xDiv = rsHREmp("ED_DIV")
'
'                        'Add Audit Log
'                        rsAudit.Open "SELECT * FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
'
'                        rsAudit.AddNew
'                        rsAudit("AU_LOC_TABL") = "EDLC": rsAudit("AU_SECTION_TABL") = "EDSE": rsAudit("AU_EMP_TABL") = "EDEM": rsAudit("AU_SUPCODE_TABL") = "EDSP": rsAudit("AU_ORG_TABL") = "EDOR": rsAudit("AU_PAYP_TABL") = "SDPP": rsAudit("AU_BCODE_TABL") = "BNCD": rsAudit("AU_TREAS_TABL") = "TERM": rsAudit("AU_DOLENT_TABL") = "EDOL": rsAudit("AU_EARN_TABL") = "EARN"
'                        rsAudit("AU_ADMINBY_TABL") = "EDAB": rsAudit("AU_LANG1_TABL") = "EDL1": rsAudit("AU_LANG2_TABL") = "EDL1"
'
'                        rsAudit("AU_NEWEMP") = "N"
'                        rsAudit("AU_PTUPL") = xPT
'                        rsAudit("AU_DIVUPL") = xDiv
'                        rsAudit("AU_COMPNO") = "001"
'                        rsAudit("AU_EMPNBR") = rsHREmp("ED_EMPNBR")
'
'                        If IsDate(xNewSenDate) Then
'                            rsAudit("AU_SENDTE") = CVDate(xNewSenDate)
'                        End If
'
'                        rsAudit("AU_LDATE") = Date
'                        rsAudit("AU_LUSER") = glbUserID
'                        rsAudit("AU_LTIME") = Time$
'                        rsAudit("AU_UPLOAD") = "N"
'                        rsAudit("AU_TYPE") = "M"
'
'                        rsAudit.Update
'                        rsAudit.Close
'                        Set rsAudit = Nothing
'                    End If
'                End If
'            End If
'NextEmployee:
'            rsHREmp.MoveNext
'        Loop
'        MsgBox "Employees Seniority Date updated successfully.", vbOKOnly, "Seniority Date Calculation"
'    Else
'        MsgBox "No Employees to update.", vbOKOnly, "Seniority Date Calculation"
'    End If
'    rsHREmp.Close
'    Set rsHREmp = Nothing
'End Sub

Private Sub Compute_SeniorityDate_NewLogic()
    Dim rsHREmp As New ADODB.Recordset
    Dim rsAttend As New ADODB.Recordset
    Dim rsAudit As New ADODB.Recordset
    Dim SQLQ As String
    
    Dim xPT As String
    Dim xDiv As String
    Dim xTotHrs As Double
    Dim xTotHrsDays As Double
    Dim xNewSenDate As Date
    
    Dim xStdHrsWeek
    Dim xHrsWeek
    Dim xHrsDay
    Dim x6MnthsWeek
    Dim xHrsNotWorked
    Dim xTotPDWHrs
    Dim xDaysNotWorked
    
    'Initialise
    xStdHrsWeek = 35
    xHrsDay = 7
    x6MnthsWeek = 26
    xHrsNotWorked = 0
    xTotPDWHrs = 0
    
    'CUPE employees only
    SQLQ = "SELECT ED_EMPNBR, ED_SENDTE, ED_PT, ED_DIV, ED_LDATE, ED_LTIME, ED_LUSER FROM HREMP WHERE ED_ORG = 'CUPE'"
    SQLQ = SQLQ & " AND ED_SENDTE IS NOT NULL"
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHREmp.EOF Then
        'For each employee get the total hours grouped by employee's Hours/Week
        'By Reason Codes LOA and USL with Union CUPE. And
        'By Union <> CUPE
        rsHREmp.MoveFirst
        
        Do While Not rsHREmp.EOF
            xTotHrs = 0
            xHrsNotWorked = 0
            xTotPDWHrs = 0
            
            'Make sure Seniority Date is valid date
            If IsDate(rsHREmp("ED_SENDTE")) Then
                'Get Total Employee's Hours per Week as they are multi position
                xHrsWeek = GetJHData_TotWHRS(rsHREmp("ED_EMPNBR"), "JH_WHRS", 0)
                
                If xHrsWeek = 0 Then GoTo NextEmployee
                
                ''Compute Hours Not Worked
                ''xHrsNotWorked = xStdHrsWeek - xHrsWeek
                'Compute Days Not Worked
                xDaysNotWorked = DaysNotWorked(rsHREmp("ED_EMPNBR"), xHrsWeek, 0)
                
                'Attendance
                SQLQ = "SELECT AD_EMPNBR, SUM(AD_HRS) AS TOT_HRS FROM HR_ATTENDANCE"
                SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR")
                SQLQ = SQLQ & " AND (AD_DOA >= " & Date_SQL(dlpDateRange(0).Text)
                SQLQ = SQLQ & " AND AD_DOA <= " & Date_SQL(dlpDateRange(1).Text) & ")"
                'Ticket #29283 - Other sets of codes
                'Ticket #28921 - Additional codes
                'SQLQ = SQLQ & " AND (AD_REASON IN ('LOA','USL'))"    ' AND AD_ORG = 'CUPE'))"
                'SQLQ = SQLQ & " AND (AD_REASON IN ('LOA','USL','LAID','OT','UPL','VCP'))"    ' AND AD_ORG = 'CUPE'))"
                SQLQ = SQLQ & " AND (AD_REASON IN ('LOA','USL','LAID','UPL','PDNW'))"    ' AND AD_ORG = 'CUPE'))"
                'SQLQ = SQLQ & " OR ((AD_DOA >= " & Date_SQL(dlpDateRange(0).Text)
                'SQLQ = SQLQ & " AND AD_DOA <= " & Date_SQL(dlpDateRange(1).Text) & ")"
                'SQLQ = SQLQ & " AND (AD_ORG <> 'CUPE')))"
                SQLQ = SQLQ & " GROUP BY AD_EMPNBR"  ', AD_WHRS"
                'SQLQ = SQLQ & " ORDER BY AD_WHRS"
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsAttend.EOF Then
                    rsAttend.MoveFirst
                            
                    'For each Total Hours by Hours/Week compute Total Hours in Weeks
                    Do While Not rsAttend.EOF
                        If rsAttend("TOT_HRS") > 0 Then
                            xTotHrs = xTotHrs + rsAttend("TOT_HRS")   '(rsAttend("TOT_HRS") / rsAttend("AD_WHRS"))
                        End If
                        
                        rsAttend.MoveNext
                    Loop
                End If
                rsAttend.Close
                Set rsAttend = Nothing
                
                'Attendance History
                SQLQ = "SELECT AH_EMPNBR, SUM(AH_HRS) AS TOT_HRS FROM HR_ATTENDANCE_HISTORY"
                SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR")
                SQLQ = SQLQ & " AND (AH_DOA >= " & Date_SQL(dlpDateRange(0).Text)
                SQLQ = SQLQ & " AND AH_DOA <= " & Date_SQL(dlpDateRange(1).Text) & ")"
                'Ticket #29283 - Other sets of codes
                'Ticket #28921 - Additional codes
                'SQLQ = SQLQ & " AND (AH_REASON IN ('LOA','USL'))"    ' AND AH_ORG = 'CUPE'))"
                'SQLQ = SQLQ & " AND (AH_REASON IN ('LOA','USL','LAID','OT','UPL','VCP'))"    ' AND AH_ORG = 'CUPE'))"
                SQLQ = SQLQ & " AND (AH_REASON IN ('LOA','USL','LAID','UPL','PDNW'))"    ' AND AH_ORG = 'CUPE'))"
                'SQLQ = SQLQ & " OR ((AH_DOA >= " & Date_SQL(dlpDateRange(0).Text)
                'SQLQ = SQLQ & " AND AH_DOA <= " & Date_SQL(dlpDateRange(1).Text) & ")"
                'SQLQ = SQLQ & " AND (AH_ORG <> 'CUPE')))"
                SQLQ = SQLQ & " GROUP BY AH_EMPNBR" ', AH_WHRS"
                'SQLQ = SQLQ & " ORDER BY AH_WHRS"
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsAttend.EOF Then
                    rsAttend.MoveFirst
                            
                    'For each Total Hours by Hours/Week compute Total Hours in Weeks
                    Do While Not rsAttend.EOF
                        If rsAttend("TOT_HRS") > 0 Then
                            xTotHrs = xTotHrs + rsAttend("TOT_HRS")   '(rsAttend("TOT_HRS") / rsAttend("AH_WHRS"))
                        End If
                        
                        rsAttend.MoveNext
                    Loop
                End If
                rsAttend.Close
                Set rsAttend = Nothing
            
                
                'Ticket #29283 - New Step - Subtract PDW Hours of the same period from Total Hours (xTotHrs)
                'Attendance - PDW
                SQLQ = "SELECT AD_EMPNBR, SUM(AD_HRS) AS TOT_HRS FROM HR_ATTENDANCE"
                SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR")
                SQLQ = SQLQ & " AND (AD_DOA >= " & Date_SQL(dlpDateRange(0).Text)
                SQLQ = SQLQ & " AND AD_DOA <= " & Date_SQL(dlpDateRange(1).Text) & ")"
                SQLQ = SQLQ & " AND (AD_REASON IN ('PDW'))"
                SQLQ = SQLQ & " GROUP BY AD_EMPNBR"
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsAttend.EOF Then
                    rsAttend.MoveFirst
                    
                    If rsAttend("TOT_HRS") > 0 Then
                        xTotPDWHrs = xTotPDWHrs + rsAttend("TOT_HRS")
                    End If
                End If
                rsAttend.Close
                Set rsAttend = Nothing
                
                'Attendance History - PDW
                SQLQ = "SELECT AH_EMPNBR, SUM(AH_HRS) AS TOT_HRS FROM HR_ATTENDANCE_HISTORY"
                SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR")
                SQLQ = SQLQ & " AND (AH_DOA >= " & Date_SQL(dlpDateRange(0).Text)
                SQLQ = SQLQ & " AND AH_DOA <= " & Date_SQL(dlpDateRange(1).Text) & ")"
                SQLQ = SQLQ & " AND (AH_REASON IN ('PDW'))"
                SQLQ = SQLQ & " GROUP BY AH_EMPNBR"
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsAttend.EOF Then
                    rsAttend.MoveFirst
                            
                    If rsAttend("TOT_HRS") > 0 Then
                        xTotPDWHrs = xTotPDWHrs + rsAttend("TOT_HRS")
                    End If
                End If
                rsAttend.Close
                Set rsAttend = Nothing
                
                'Ticket #29283 - Subtract Total PDW Hours from Total Hours
                xTotHrs = xTotHrs - xTotPDWHrs
                                
                
                ''Compute Hours Not Worked in 6 months - doing this before including the LOA and USL hours because the
                ''26weeks/6months is as per the standard # of hours not worked.
                'xHrsNotWorked = x6MnthsWeek * xHrsNotWorked
            
                ''Compute Total Hours Not Worked with LOA and USL hours included - Now adding to actual hours not worked
                ''for 6 months.
                'xHrsNotWorked = xHrsNotWorked + xTotHrs
            
                'Convert Total Hours not worked from Attendance to Days before adding to Total Days Not Worked from employee's Hours per Week
                If xTotHrs > 0 Then
                    xTotHrsDays = Round((xTotHrs / xHrsDay), 0) + xDaysNotWorked
                Else
                    xTotHrsDays = xDaysNotWorked
                End If
                
                
                'Compute Seniority Date by Work Weeks, excluding Saturday/Sunday - Stat Holidays should be included (as per client)
                'so, adding Total Hours not worked (xHrsNotWorked) to Seniority Date
                If xTotHrsDays > 0 Then
                    ''Convert the Total # of Hours Not Worked to # of Days, so we can exclude non work days when
                    ''adding to Seniority Date
                    'xTotHrsDays = xHrsNotWorked / xHrsDay
                    
                    'Ticket #26514 - Include the Seniority Date itself as Day 1 missed work. Therefore, I am
                    'subtracting 1 from total days not worked to include teh Seniority Date as Day 1.
                    'Add # of Total Hours Not Worked in Days to the Seniority excluding Weekends
                    'They want to include Statutory Holidays
                    'xNewSenDate = AddWorkingDays(rsHREmp("ED_SENDTE"), xTotHrsDays, False)
                    xNewSenDate = AddWorkingDays(rsHREmp("ED_SENDTE"), xTotHrsDays - 1, False)
                    
                    'Update Seniority Date in HREMP if the date has changed
                    If CVDate(rsHREmp("ED_SENDTE")) <> CVDate(xNewSenDate) Then
                        rsHREmp("ED_SENDTE") = xNewSenDate
                        rsHREmp("ED_LDATE") = Now
                        rsHREmp("ED_LTIME") = Time$
                        rsHREmp("ED_LUSER") = glbUserID
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
                        rsAudit("AU_LUSER") = glbUserID
                        rsAudit("AU_LTIME") = Time$
                        rsAudit("AU_UPLOAD") = "N"
                        rsAudit("AU_TYPE") = "M"
                        
                        rsAudit.Update
                        rsAudit.Close
                        Set rsAudit = Nothing
                    End If
                End If
            End If
NextEmployee:
            rsHREmp.MoveNext
        Loop
        MsgBox "Employees Seniority Date updated successfully.", vbOKOnly, "Seniority Date Calculation"
    Else
        MsgBox "No Employees to update.", vbOKOnly, "Seniority Date Calculation"
    End If
    rsHREmp.Close
    Set rsHREmp = Nothing
End Sub

Private Function GetJHData_TotWHRS(EmpNbr, Field As String, DEFAULT)
    Dim rsJHTEMP As New ADODB.Recordset
    rsJHTEMP.Open "SELECT SUM(" & Field & ") AS JH_TOTHRS FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & EmpNbr, gdbAdoIhr001, adOpenForwardOnly
    GetJHData_TotWHRS = DEFAULT
    
    If Not rsJHTEMP.EOF Then
        If Not IsNull(rsJHTEMP("JH_TOTHRS")) Then GetJHData_TotWHRS = rsJHTEMP("JH_TOTHRS")
    End If
    rsJHTEMP.Close
    Set rsJHTEMP = Nothing
End Function

Private Function StartDate_Same_MultiPos(EmpNbr)
    Dim rsJHTEMP As New ADODB.Recordset
    Dim xStartDate
    
    StartDate_Same_MultiPos = False
    
    rsJHTEMP.Open "SELECT JH_SDATE FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & EmpNbr, gdbAdoIhr001, adOpenForwardOnly
    If Not rsJHTEMP.EOF Then
        Do While Not rsJHTEMP.EOF
            If Not IsDate(xStartDate) Then
                xStartDate = rsJHTEMP("JH_SDATE")
            ElseIf CVDate(rsJHTEMP("JH_SDATE")) = CVDate(xStartDate) Then
                StartDate_Same_MultiPos = True
            Else
                StartDate_Same_MultiPos = False
            End If
            rsJHTEMP.MoveNext
        Loop
    End If
    rsJHTEMP.Close
    Set rsJHTEMP = Nothing
End Function

Private Function DaysNotWorked(EmpNbr, xTotCurHrsWK, DEFAULT)
    Dim rsEmpJob As New ADODB.Recordset
    Dim xHrsWeek
    Dim xDaysToNextJob      '# of days from start of the period to next Job Start Date
    Dim xWeeksToNextJob     '# of weeks from start of the period to next Job Start Date
    Dim xPerWeekHrsNotWorked
    Dim xPerPeriodHrsNotWorked
    Dim xDaysNotWorked
    Dim xStdHrsWeek, xDaysPerWeek
    Dim xStartDate, xEndDate
    
    'Initialise
    xStdHrsWeek = 35
    xDaysPerWeek = 7
    xDaysNotWorked = 0
    xStartDate = dlpDateRange(0).Text
    xEndDate = dlpDateRange(1).Text
    
    rsEmpJob.Open "SELECT * FROM HR_JOB_HISTORY WHERE JH_EMPNBR=" & EmpNbr & " AND JH_SDATE <= " & Date_SQL(dlpDateRange(1).Text) & " ORDER BY JH_CURRENT DESC, JH_SDATE DESC", gdbAdoIhr001, adOpenForwardOnly
    'rsEmpJob.Open "SELECT * FROM HR_JOB_HISTORY WHERE JH_EMPNBR=" & EmpNbr & " AND JH_SDATE <= " & Date_SQL(xEndDate) & " ORDER BY JH_CURRENT DESC, JH_SDATE DESC", gdbAdoIhr001, adOpenForwardOnly
    DaysNotWorked = DEFAULT
    
    Do While Not rsEmpJob.EOF
        If CVDate(rsEmpJob("JH_SDATE")) > CVDate(xStartDate) Then
            'Job Start Date is still in 6 months range, move to previous job
            rsEmpJob.MoveNext
        Else
            'Job Start Date is either same as the 6 month range From Date or prior to From Date: i.e. 6 Month Start = 7/1/2016, Job Start Date = 4/1/2016
            '   - Get the Hours per Week (A): 22.5 Hrs/Wk from 4/1/2016
            'Move to Previous record (more recent) to get the Job Start Date - 1: i.e. Job Start Date 8/1/2016 - 1 = 7/31/2016
            '   - so we know how long the Hours per Week (from A above) were worked at from From Date of the 6 month range to Job Start Date: i.e. 7/1/2016 upto 7/31/2016 @ 22.5
            '   - get # of days worked at 22.5 (B): 7/31/2016 - 7/1/2016 = 30
            '   - get # of weeks worked at 22.5 C = Round(B/7): Round(30 / 7,0) = 4
            '   - get # of Hours not Worked D in a week = 35 - A: 35 - 22.5 = 12.5
            '   - get # of Weeks not Worked (E) for the period C, i.e. E = D * C: 12.5 * 4 = 50
            '   - get # of Days not Worked F = Round(E / 7,0): Round(50 / 7,0) = 7
            '   - Add F working days to the Seniority Date
            
            'A - Hours/Week the employee started with just prior to the 6 month range
            'If multi position and total hours per week is Standard Hours per Week (35) then user Standard Hours per Week as Hours per Week.
            'Employee #735 has two position totalling 35 hours per week but starts mid period and both positions have same start date
            'Find of Start Date of Multi Position is same
            If xTotCurHrsWK = xStdHrsWeek And rsEmpJob("JH_CURRENT") <> 0 And StartDate_Same_MultiPos(EmpNbr) Then
                xHrsWeek = xStdHrsWeek
            Else
                xHrsWeek = rsEmpJob("JH_WHRS")
            End If
            
            'Go get to the next Job Start Date where the Hours/Week may have changed
            rsEmpJob.MovePrevious
            
            'if EOF then use End Date
            If rsEmpJob.BOF Then
                'B - # of days from start of the period to next Job Start Date
                'xDaysToNextJob = DateDiff("d", CVDate(xEndDate), CVDate(xStartDate))
                xDaysToNextJob = DateDiff("d", CVDate(xStartDate), CVDate(xEndDate))
            Else
                'B - # of days from start of the period to next Job Start Date
                'xDaysToNextJob = DateDiff("d", CVDate(rsEmpJob("JH_SDATE") - 1), CVDate(xStartDate))
                xDaysToNextJob = DateDiff("d", CVDate(xStartDate), CVDate(rsEmpJob("JH_SDATE")))
            End If
            
            'C - # of weeks from start of the period to next Job Start Date based on # of days in B: Round((B / 7),0)
            xWeeksToNextJob = Round(xDaysToNextJob / xDaysPerWeek, 0)
            
            'D - # of hours not worked per week until next job B -> where 35 is Standard Hours per Week (35 - A)
            xPerWeekHrsNotWorked = xStdHrsWeek - xHrsWeek
            
            'E - # of hours not worked in C weeks based on hours not worked in D per week (D * C) until next Job
            xPerPeriodHrsNotWorked = xPerWeekHrsNotWorked * xWeeksToNextJob
            
            'F - # of days not worked based on hours not worked in E: Round((E / 7),0)
            xDaysNotWorked = xDaysNotWorked + Round((xPerPeriodHrsNotWorked / xDaysPerWeek), 0)
            
            If Not rsEmpJob.BOF Then
                'Next set of Jobs with Start Date
                xStartDate = rsEmpJob("JH_SDATE")
            Else
                Exit Do
            End If
        End If
    Loop
    rsEmpJob.Close
    Set rsEmpJob = Nothing

    'Return F "xDaysNotWorked" to add to Seniority Date as Working Days
    DaysNotWorked = xDaysNotWorked
    
End Function
