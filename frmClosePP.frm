VERSION 5.00
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmClosePP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Close a Pay Period"
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
         Left            =   2220
         Picture         =   "frmClosePP.frx":0000
         Top             =   742
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
         Caption         =   "Entitlement Period: "
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2430
         TabIndex        =   3
         Top             =   270
         Visible         =   0   'False
         Width           =   1380
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
Attribute VB_Name = "frmClosePP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xVacFromDate
Dim xVacToDate
Dim fglbWDate$


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

'Valid Pay Period #
If Not IsNumeric(txtWeek.Text) Then
    MsgBox "Invalid Pay Period #. Please select valid Pay Period.", vbExclamation
    txtWeek.SetFocus
    Exit Sub
End If

'Check if the previous PP's have been close
If Previous_PayPeriod_Open Then
    MsgBox "Pay Period prior to the selected Pay Period # " & txtWeek.Text & " has not been closed yet. Please close the previous Pay Period(s) first.", vbExclamation
    txtWeek.SetFocus
    Exit Sub
End If

If IsDate(dlpDateRange(0)) And IsDate(dlpDateRange(1)) Then
    'Proceed with the update
    Call Close_PayPeriod_HourlyEntitlement
Else
    MsgBox "Invalid Pay Period Date Range. Please select valid Pay Period.", vbExclamation
    txtWeek.SetFocus
    Exit Sub
End If

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
glbOnTop = "FRMCLOSEPP"
End Sub

Private Sub Form_Load()
glbOnTop = "FRMCLOSEPP"

    Select Case glbCompWDate$ ' sets field reference for basic 'which date'
        Case "O": fglbWDate$ = "ED_DOH"
        Case "S": fglbWDate$ = "ED_SENDTE"
        Case "U": fglbWDate$ = "ED_UNION"
        Case "L": fglbWDate$ = "ED_LTHIRE"
        Case "D": fglbWDate$ = "ED_USRDAT1"
    End Select

'Retrieve Vacation Entitlement Period
Call VacationEntitlementPeriod

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
frmPayPeriodList.SelectedYear = Year(Now)
frmPayPeriodList.Closed = False
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

Private Function Previous_PayPeriod_Open()
    Dim rsPPMaster As New ADODB.Recordset
    Dim SQLQ As String
    
    Previous_PayPeriod_Open = False
    
    SQLQ = "SELECT * FROM HR_PAYPERIOD WHERE PP_YEAR=" & Year(Now)
    SQLQ = SQLQ & " AND PP_NBR < " & glbWeek    'Previous PP
    SQLQ = SQLQ & " AND PP_UPLOADED = 0"
    SQLQ = SQLQ & " ORDER BY PP_NBR"
    rsPPMaster.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsPPMaster.EOF Then
        'Rows found means they are previous PPs not closed
        Previous_PayPeriod_Open = True
    Else
        'No rows returned so no previous PPs - so all closed
        Previous_PayPeriod_Open = False
    End If
    rsPPMaster.Close
    Set rsPPMaster = Nothing

End Function

Private Function Close_PayPeriod_HourlyEntitlement()
    Dim SQLQ As String
    Dim ESQLQ  As String
    Dim rsHREntMst As New ADODB.Recordset
    Dim snapHREntitle As New ADODB.Recordset
    
    Screen.MousePointer = HOURGLASS
    
    'Update the Vacation Pay % with the latest values
    If Update_VacationPayPercentage Then
        
        'Compute the Entitlement and Update Employees and Accrual table
        If Update_HourlyEntitlement Then
        
            'Recalculate the Hourly Entitlement
            Call EntReCalcHr
            
            'Close the Pay Period
            If Close_Pay_Period Then
                MsgBox "Successfully updated employees with Hourly Entitlements." & vbCrLf & vbCrLf & "Pay Period #" & txtWeek.Text & " is closed now.", vbInformation, "Close Pay Period"
            Else
                MsgBox "An error occured closing the Pay Period." & vbCrLf & vbCrLf & "Pay Period #" & txtWeek.Text & " has NOT been closed yet.", vbExclamation, "Failed to Close Pay Period"
            End If
        Else
            Screen.MousePointer = DEFAULT
            
            MsgBox "An error occured computing employee's Hourly Entitlement. " & vbCrLf & vbCrLf & "Pay Period #" & txtWeek.Text & " has NOT been closed yet.", vbExclamation, "Failed to Close Pay Period"
        End If
    End If
    Screen.MousePointer = DEFAULT
End Function

Private Function Update_VacationPayPercentage()
    Dim rsVacPayPerc As New ADODB.Recordset
    Dim rsVacPayPercDtl As New ADODB.Recordset
    Dim snapEntitle As New ADODB.Recordset
    Dim SQLQ As String
    Dim ESQLQ As String
    Dim EmpNo As Long
    Dim dblServiceHours#
    Dim varStartDate As Variant
    Dim lngRecs&
    Dim dblVacPayPct#, intWhereFit&, x%, Y%, z%
    Dim Msg$, Title$, DgDef As Variant
    Dim Response%, pct%
    Dim prec%, xAsOf
    Dim rsAudit As New ADODB.Recordset
    Dim xPT As String
    Dim xDiv As String
    Dim OVACPC

    On Error GoTo Update_VacationPayPercentage_Err
    
    Screen.MousePointer = HOURGLASS
    
    Update_VacationPayPercentage = False
    
    SQLQ = "SELECT DISTINCT VP_DIV,VP_DEPT,VP_ORG,VP_LOC,VP_SECTION,VP_EMP,VP_PT,VP_GRPCD,VP_FRDATE,VP_TODATE,VP_MANUAL,VP_EDATE FROM HRVACPCTENT "
    SQLQ = SQLQ & " WHERE VP_FRDATE <= " & Date_SQL(dlpDateRange(0).Text)
    SQLQ = SQLQ & " AND VP_TODATE >= " & Date_SQL(dlpDateRange(1).Text)
    If glbDIVCount = 1 And glbLinamar Then
        SQLQ = SQLQ & " AND VP_DIV IN (select DIV from HR_DIVISION WHERE " & glbSeleDiv & ")"
    End If
    rsVacPayPerc.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsVacPayPerc.EOF Then
        rsVacPayPerc.MoveFirst
        
        Do While Not rsVacPayPerc.EOF
            'Get employee List to update
            ESQLQ = glbSeleDeptUn
            
            If Len(rsVacPayPerc("VP_DEPT")) > 0 Then ESQLQ = ESQLQ & " AND  ED_DEPTNO = '" & rsVacPayPerc("VP_DEPT") & "' "
            If Len(rsVacPayPerc("VP_DIV")) > 0 Then ESQLQ = ESQLQ & " AND ED_DIV = '" & rsVacPayPerc("VP_DIV") & "' "
            If Len(rsVacPayPerc("VP_ORG")) > 0 Then ESQLQ = ESQLQ & " AND ED_ORG = '" & rsVacPayPerc("VP_ORG") & "' "
            If Len(rsVacPayPerc("VP_EMP")) > 0 Then ESQLQ = ESQLQ & " AND ED_EMP = '" & rsVacPayPerc("VP_EMP") & "' "
            If Len(rsVacPayPerc("VP_SECTION")) > 0 Then ESQLQ = ESQLQ & " AND ED_SECTION = '" & rsVacPayPerc("VP_SECTION") & "' "
            If Len(rsVacPayPerc("VP_LOC")) > 0 Then ESQLQ = ESQLQ & " AND ED_LOC = '" & rsVacPayPerc("VP_LOC") & "' "
            If Len(rsVacPayPerc("VP_PT")) > 0 Then ESQLQ = ESQLQ & " AND ED_PT = '" & rsVacPayPerc("VP_PT") & "' "
            
            SQLQ = "SELECT ED_EMPNBR,ED_VACPC,ED_PVAC,ED_VAC,ED_PSICK,ED_SICK,ED_SICKT,ED_EFDATES,ED_ETDATES, HREMP.ED_ANNVAC, HREMP.ED_ANNSICK, "
            SQLQ = SQLQ & " ED_DIV,ED_PT, ED_SECTION, ED_LOC, ED_EMP,"
            SQLQ = SQLQ & " ED_HIRECODE," 'County of Brant Ticket #12525
            SQLQ = SQLQ & " ED_DOH, ED_SENDTE,ED_UNION,ED_LTHIRE,ED_USRDAT1,ED_LUSER,ED_LDATE,ED_LTIME "
            SQLQ = SQLQ & " FROM HREMP WHERE " & ESQLQ
            If Len(rsVacPayPerc("VP_GRPCD")) > 0 Then
                SQLQ = SQLQ & " AND ED_EMPNBR IN "
                SQLQ = SQLQ & " (SELECT JH_EMPNBR FROM qry_JobCurrent "
                SQLQ = SQLQ & " WHERE JB_GRPCD = '" & rsVacPayPerc("VP_GRPCD") & "') "
            End If
            If snapEntitle.State <> 0 Then snapEntitle.Close
            snapEntitle.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
            
            If Not snapEntitle.EOF Then
                MDIMain.panHelp(0).FloodType = 1
                MDIMain.panHelp(0).FloodPercent = 5
                
                lngRecs& = snapEntitle.RecordCount
                
                'Retrieve the Complete Vacation Pay Percentage Rule
                SQLQ = "SELECT * FROM HRVACPCTENT WHERE "
                SQLQ = SQLQ & " (VP_DEPT = '" & rsVacPayPerc("VP_DEPT") & "' " & IIf(Len(rsVacPayPerc("VP_DEPT")) = 0, " OR VP_DEPT IS NULL ", "") & ")"
                SQLQ = SQLQ & " AND (VP_DIV = '" & rsVacPayPerc("VP_DIV") & "' " & IIf(Len(rsVacPayPerc("VP_DIV")) = 0, " OR VP_DIV IS NULL ", "") & ")"
                SQLQ = SQLQ & " AND (VP_ORG = '" & rsVacPayPerc("VP_ORG") & "' " & IIf(Len(rsVacPayPerc("VP_ORG")) = 0, " OR VP_ORG IS NULL ", "") & ")"
                SQLQ = SQLQ & " AND (VP_EMP = '" & rsVacPayPerc("VP_EMP") & "' " & IIf(Len(rsVacPayPerc("VP_EMP")) = 0, " OR VP_EMP IS NULL ", "") & ")"
                SQLQ = SQLQ & " AND (VP_SECTION = '" & rsVacPayPerc("VP_SECTION") & "' " & IIf(Len(rsVacPayPerc("VP_SECTION")) = 0, " OR VP_SECTION IS NULL ", "") & ")"
                SQLQ = SQLQ & " AND (VP_LOC = '" & rsVacPayPerc("VP_LOC") & "' " & IIf(Len(rsVacPayPerc("VP_LOC")) = 0, " OR VP_LOC IS NULL ", "") & ")"
                SQLQ = SQLQ & " AND (VP_PT = '" & rsVacPayPerc("VP_PT") & "' " & IIf(Len(rsVacPayPerc("VP_PT")) = 0, " OR VP_PT IS NULL ", "") & ")"
                SQLQ = SQLQ & " AND (VP_GRPCD = '" & rsVacPayPerc("VP_GRPCD") & "' " & IIf(Len(rsVacPayPerc("VP_GRPCD")) = 0, " OR VP_GRPCD IS NULL ", "") & ")"
                SQLQ = SQLQ & " AND (VP_FRDATE = " & Date_SQL(rsVacPayPerc("VP_FRDATE")) & ")"
                SQLQ = SQLQ & " AND (VP_TODATE = " & Date_SQL(rsVacPayPerc("VP_FRDATE")) & ")"
                SQLQ = SQLQ & " ORDER BY VP_DIV,VP_DEPT,VP_ORG,VP_EMP,VP_PT,VP_LOC,VP_SECTION,VP_ORDER "
                rsVacPayPercDtl.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsVacPayPercDtl.EOF Then
                                
                    While Not snapEntitle.EOF
                        prec% = prec% + 1
                        pct% = Int(100 * (prec% / lngRecs&))
                        MDIMain.panHelp(0).FloodPercent = pct%
                                    
                        EmpNo& = snapEntitle("ED_EMPNBR")
                    
                        'Ticket #29617 - Mississaugas of Scugog Island First Nation
                        'Get the length of service by months
                        If glbCompSerial = "S/N - 2485W" Then
                            'Vacation / Sick Mass Updated Based Upon
                            If IsNull(snapEntitle(fglbWDate$)) Then GoTo lblNextRec
                    
                            varStartDate = snapEntitle(fglbWDate$)
                            
                            xAsOf = dlpDateRange(1).Text
                            
                            'Length of Service in Months based on Vacation / Sick Mass Update Based Upon
                            dblServiceHours# = MonthDiff(CVDate(varStartDate), CVDate(xAsOf))
                        Else
                            'Get Total Non Absent Hours from Attendance and Attendance History
                            dblServiceHours# = Total_NonAbsent_Hours(snapEntitle("ED_EMPNBR"), dlpDateRange(0).Text, dlpDateRange(1).Text)
                        End If
                        
                        intWhereFit& = -1
                    
                        'Which range of the service month the employee falls in and if the Vacation Pay % exists for that range
                        rsVacPayPercDtl.MoveFirst
                        Do While Not rsVacPayPercDtl.EOF
                            If rsVacPayPercDtl("VP_EHOUR") > 0 Then
                                If dblServiceHours# >= CDbl(rsVacPayPercDtl("VP_BHOUR")) And dblServiceHours# <= CDbl(rsVacPayPercDtl("VP_EHOUR")) Then
                                    intWhereFit& = 1
                                    Exit Do
                                End If
                            End If
                            rsVacPayPercDtl.MoveNext
                        Loop
                                            
                        If intWhereFit& = -1 Or dblServiceHours# < 0 Then GoTo lblNextRec ' skip record if not in any of the ranges
                    
                        dblVacPayPct# = rsVacPayPercDtl("VP_PCT")
                        
                        OVACPC = ""
                        OVACPC = snapEntitle("ED_VACPC")
                        
                        snapEntitle("ED_VACPC") = dblVacPayPct# '* 100
                        snapEntitle("ED_LDATE") = Now
                        snapEntitle("ED_LTIME") = Time$
                        snapEntitle("ED_LUSER") = glbUserID
                        snapEntitle.Update
                        
                        
                        'Update Audit
                        If OVACPC <> dblVacPayPct# Then
                            'Retrieve PT and Div from HREMP
                            If IsNull(snapEntitle("ED_PT")) Then xPT = "" Else xPT = snapEntitle("ED_PT")
                            If IsNull(snapEntitle("ED_DIV")) Then xDiv = "" Else xDiv = snapEntitle("ED_DIV")
                                                
                            'Add Audit Log
                            rsAudit.Open "SELECT * FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
                            
                            rsAudit.AddNew
                            rsAudit("AU_LOC_TABL") = "EDLC": rsAudit("AU_SECTION_TABL") = "EDSE": rsAudit("AU_EMP_TABL") = "EDEM": rsAudit("AU_SUPCODE_TABL") = "EDSP": rsAudit("AU_ORG_TABL") = "EDOR": rsAudit("AU_PAYP_TABL") = "SDPP": rsAudit("AU_BCODE_TABL") = "BNCD": rsAudit("AU_TREAS_TABL") = "TERM": rsAudit("AU_DOLENT_TABL") = "EDOL": rsAudit("AU_EARN_TABL") = "EARN"
                            rsAudit("AU_ADMINBY_TABL") = "EDAB": rsAudit("AU_LANG1_TABL") = "EDL1": rsAudit("AU_LANG2_TABL") = "EDL1"
                            
                            rsAudit("AU_NEWEMP") = "N"
                            rsAudit("AU_PTUPL") = xPT
                            rsAudit("AU_DIVUPL") = xDiv
                            rsAudit("AU_COMPNO") = "001"
                            rsAudit("AU_EMPNBR") = snapEntitle("ED_EMPNBR")
                                    
                            If OVACPC <> dblVacPayPct# Then
                                If IsNumeric(dblVacPayPct#) Then rsAudit("AU_VACPC") = dblVacPayPct# * 100
                                If IsNumeric(OVACPC) Then rsAudit("AU_OLDVAC") = OVACPC * 100
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
lblNextRec:
                        DoEvents
                        
                        snapEntitle.MoveNext
                    Wend
                
                End If
                rsVacPayPercDtl.Close
                Set rsVacPayPercDtl = Nothing
            
            End If
            snapEntitle.Close
            Set snapEntitle = Nothing
            
            Update_VacationPayPercentage = True
            
            MDIMain.panHelp(0).FloodType = 0
                    
            rsVacPayPerc.MoveNext
        Loop
                
    End If
    rsVacPayPerc.Close
    Set rsVacPayPerc = Nothing

    Update_VacationPayPercentage = True
    
    Screen.MousePointer = DEFAULT

Exit Function

Update_VacationPayPercentage_Err:
'These errors are:
'13=type mismatch
'94=invalid use of null
'3018=couln't find field 'item'
If Err = 13 Or Err = 94 Or Err = 3018 Then
   ' MsgBox "Err:" & Str(Err) & Chr(10) & Error$ & Chr(10) & " modUpdateSelection" & Chr(10) & "FORM:FUENTITL.FRM"
    'commented out by RAUBREY 5/20/97
    Err = 0
    Resume Next
End If

Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "UpdateVacationPay%", "HREMP", "Close Pay Period")

Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    'Rollback
    Resume Next
Else
    Unload Me
End If

End Function

Private Function Update_HourlyEntitlement()
    Dim SQLQ As String
    Dim ESQLQ  As String
    Dim rsHREntMst As New ADODB.Recordset
    Dim rsHREntMstDtl As New ADODB.Recordset
    Dim snapHREntitle As New ADODB.Recordset
    Dim snapHREnt As New ADODB.Recordset
    Dim rzAttend As New ADODB.Recordset
    Dim rsHREmp As New ADODB.Recordset
    Dim fTablHREMP As New ADODB.Recordset
    Dim EmpNo As Long
    Dim rsCurJobSal As New ADODB.Recordset
    Dim dblServiceHours#, dblNewEntitle#, dblEntitleUpd#, dblEntitle#
    Dim lngRecs&
    Dim intWhereFit&, x%, Y%, z%
    Dim Msg$, Title$, DgDef As Variant
    Dim Response%, pct%
    Dim prec%
    Dim xComments
    Dim PayPerc
    Dim varStartDate As Variant
    Dim oldEntitleUpd
    Dim dblServiceYears#
    Dim dblNewMax#
    Dim NumRec As Integer
    Dim xKey
    Dim xUpdMethod As String

    On Error GoTo Update_HourlyEntitlement_Err

    Screen.MousePointer = HOURGLASS
    
    Update_HourlyEntitlement = False
    
    'Compute the Entitlement
    'Update Employees and Accrual table
    
    'Retrieve the specific Hourly Entitlements rules
    SQLQ = "SELECT DISTINCT EH_DIV,EH_DEPT,EH_ORG,EH_FDATE,EH_TDATE,EH_EMP,EH_SECTION,EH_LOC,EH_PT,EH_HETYPE,EH_MANUAL,EH_EDATE,EH_UPDMETHOD FROM HR_HOURLYENT "
    SQLQ = SQLQ & " WHERE EH_FDATE <= " & Date_SQL(dlpDateRange(0).Text)
    SQLQ = SQLQ & " AND EH_TDATE >= " & Date_SQL(dlpDateRange(1).Text)
    SQLQ = SQLQ & " AND EH_HETYPE IN ('IW+','VA+','SK+','PT+','PD+')"
    If glbDIVCount = 1 And glbLinamar Then
        SQLQ = SQLQ & " AND EH_DIV IN (select DIV from HR_DIVISION WHERE " & glbSeleDiv & ")"
    End If
    If glbWFC Then 'Ticket #28553 Franks 05/03/2016
        SQLQ = SQLQ & " AND " & getWFCPlantSecurity("EH_SECTION")
    End If
    rsHREntMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHREntMst.EOF Then
        rsHREntMst.MoveFirst
        
        'For each Hourly Entitlement Rule, get the list of Employees
        Do While Not rsHREntMst.EOF
            'Update Method
            xUpdMethod = IIf(Len(rsHREntMst("EH_UPDMETHOD")) = 0, "A", rsHREntMst("EH_UPDMETHOD"))
    
            'Get Employees to update
            ESQLQ = glbSeleDeptUn
                    
            If Len(rsHREntMst("EH_DEPT")) > 0 Then ESQLQ = ESQLQ & " AND ED_DEPTNO = '" & rsHREntMst("EH_DEPT") & "'"
            If Len(rsHREntMst("EH_DIV")) > 0 Then ESQLQ = ESQLQ & " AND ED_DIV = '" & rsHREntMst("EH_DIV") & "' "
            If Len(rsHREntMst("EH_ORG")) > 0 Then ESQLQ = ESQLQ & " AND ED_ORG = '" & rsHREntMst("EH_ORG") & "' "
            If Len(rsHREntMst("EH_EMP")) > 0 Then ESQLQ = ESQLQ & " AND ED_EMP = '" & rsHREntMst("EH_EMP") & "' "
            If Len(rsHREntMst("EH_SECTION")) > 0 Then ESQLQ = ESQLQ & " AND ED_SECTION = '" & rsHREntMst("EH_SECTION") & "' "
            If Len(rsHREntMst("EH_LOC")) > 0 Then ESQLQ = ESQLQ & " AND ED_LOC = '" & rsHREntMst("EH_LOC") & "' "
            If Len(rsHREntMst("EH_PT")) > 0 Then ESQLQ = ESQLQ & " AND ED_PT = '" & rsHREntMst("EH_PT") & "' "
        
            SQLQ = "SELECT JH_DHRS,JH_FTENUM,ED_EMPNBR,ED_DHRS,ED_DOH,ED_SENDTE,ED_UNION,ED_LTHIRE,ED_USRDAT1,ED_VACPC "
            SQLQ = SQLQ & "FROM HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
            SQLQ = SQLQ & "WHERE HR_JOB_HISTORY.JH_CURRENT<>0"
            SQLQ = SQLQ & " AND " & ESQLQ
            If snapHREntitle.State <> 0 Then snapHREntitle.Close
            snapHREntitle.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not snapHREntitle.EOF Then
                MDIMain.panHelp(0).FloodType = 1
                MDIMain.panHelp(0).FloodPercent = 5
                
                lngRecs& = snapHREntitle.RecordCount
                prec% = 0
                
                'Retrieve the Complete Hourly Entitlement Rule to compute the Hourly Entitlement amounts
                SQLQ = "SELECT * FROM HR_HOURLYENT WHERE "
                SQLQ = SQLQ & " (EH_DEPT = '" & rsHREntMst("EH_DEPT") & "' " & IIf(Len(rsHREntMst("EH_DEPT")) = 0, " OR EH_DEPT IS NULL ", "") & ")"
                SQLQ = SQLQ & " AND (EH_DIV = '" & rsHREntMst("EH_DIV") & "' " & IIf(Len(rsHREntMst("EH_DIV")) = 0, " OR EH_DIV IS NULL ", "") & ")"
                SQLQ = SQLQ & " AND (EH_ORG = '" & rsHREntMst("EH_ORG") & "' " & IIf(Len(rsHREntMst("EH_ORG")) = 0, " OR EH_ORG IS NULL ", "") & ")"
                SQLQ = SQLQ & " AND (EH_EMP = '" & rsHREntMst("EH_EMP") & "' " & IIf(Len(rsHREntMst("EH_EMP")) = 0, " OR EH_EMP IS NULL ", "") & ")"
                SQLQ = SQLQ & " AND (EH_SECTION = '" & rsHREntMst("EH_SECTION") & "' " & IIf(Len(rsHREntMst("EH_SECTION")) = 0, " OR EH_SECTION IS NULL ", "") & ")"
                SQLQ = SQLQ & " AND (EH_LOC = '" & rsHREntMst("EH_LOC") & "' " & IIf(Len(rsHREntMst("EH_LOC")) = 0, " OR EH_LOC IS NULL ", "") & ")"
                SQLQ = SQLQ & " AND (EH_PT = '" & rsHREntMst("EH_PT") & "' " & IIf(Len(rsHREntMst("EH_PT")) = 0, " OR EH_PT IS NULL ", "") & ")"
                SQLQ = SQLQ & " AND (EH_FDATE = " & Date_SQL(rsHREntMst("EH_FDATE")) & ")"
                SQLQ = SQLQ & " AND (EH_TDATE = " & Date_SQL(rsHREntMst("EH_TDATE")) & ")"
                SQLQ = SQLQ & " AND EH_HETYPE = '" & rsHREntMst("EH_HETYPE") & "'"
                SQLQ = SQLQ & " ORDER BY EH_DIV,EH_DEPT,EH_ORG,EH_FDATE,EH_EMP,EH_PT,EH_LOC,EH_SECTION,EH_ORDER "
                rsHREntMstDtl.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsHREntMstDtl.EOF Then
                    
                    'For each Employee and Hourly Entitlement Code, compute the entitlement based on the rule and the Attendance Code Matrix;
                    'apply Vac Pay % on it and then compare against the Maximum;
                    'Lastly Update employee's Hourly Entitlement.
                    While Not snapHREntitle.EOF
                                            
                        pct% = 100 * (prec% / lngRecs&)
                        MDIMain.panHelp(0).FloodPercent = pct%
                        prec% = prec% + 1
                        
                        'Initialise
                        dblNewEntitle = 0
                        dblEntitleUpd = 0
                    
                        If IsNull(snapHREntitle(fglbWDate$)) Then
                            GoTo lblNextRec
                        End If
                    
                    
                        EmpNo& = snapHREntitle("ED_EMPNBR")
                    
                        'Retrieve the existing entitlement
                        SQLQ = "SELECT HE_EMPNBR,HE_TYPE,HE_ID ,"
                        SQLQ = SQLQ & " HE_ENTITLE, HE_TDATE FROM HRENTHRS "
                        SQLQ = SQLQ & " WHERE HE_EMPNBR = " & snapHREntitle("ED_EMPNBR")
                        SQLQ = SQLQ & " AND HE_TYPE = '" & rsHREntMst("EH_HETYPE") & "'"
                        SQLQ = SQLQ & " AND HE_FDATE = " & Date_SQL(rsHREntMst("EH_FDATE"))
                        SQLQ = SQLQ & " AND HE_TDATE = " & Date_SQL(rsHREntMst("EH_TDATE"))
                        snapHREnt.Open SQLQ, gdbAdoIhr001, adOpenKeyset
                        If Not snapHREnt.EOF And Not snapHREnt.BOF Then
                            snapHREnt.MoveLast
                        End If
                    
                        NumRec = snapHREnt.RecordCount
                        
                        'Get existing entitlement
                        If snapHREnt.EOF Then
                            oldEntitleUpd = 0
                        Else
                            oldEntitleUpd = snapHREnt("HE_ENTITLE")
                        End If
                        
                        'Start entitlement with, based on the Update Method
                        If xUpdMethod = "A" Then
                            If NumRec > 0 Then
                                dblEntitle# = snapHREnt("HE_ENTITLE")
                            Else
                                dblEntitle# = 0
                            End If
                        Else
                            dblEntitle# = 0
                        End If
                        snapHREnt.Close
                        Set snapHREnt = Nothing
                    
                    
                        'Get Total Worked Hours from Attendance
                        dblServiceHours# = 0
                        dblServiceHours# = Total_Worked_Hours(snapHREntitle("ED_EMPNBR"), rsHREntMst("EH_HETYPE"))
                            
                        'Compute the Worked Hours based on the Vacation Pay Percentage depending on the type of Entitlement Code
                        'Get Vacation Pay Percentage based on Entitlement Code entitled
                        PayPerc = 0
                        Select Case rsHREntMst("EH_HETYPE")
                            Case "VA+"
                                If Not IsNull(snapHREntitle("ED_VACPC")) Then
                                    PayPerc = snapHREntitle("ED_VACPC")
                                End If
                            Case "SK+"
                                PayPerc = 4 / 100
                            Case "PD+"
                                PayPerc = 4 / 100
                            Case "PT+"
                                PayPerc = 4 / 100
                            Case "IW+"
                                PayPerc = 11.11 / 100
                        End Select
                        
                        'Compute the new entitlement based on the Pay %
                        dblNewEntitle# = dblServiceHours# * PayPerc
                    
                        'Check against Maximum before updating the Hourly Entitlement record
                        'Get Start Date to compute the Service Months
                        varStartDate = snapHREntitle(fglbWDate$)
                        dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(dlpDateRange(1).Text))    'Ticket #17924
                        
                        'Check if valid Service Months
                        If dblServiceYears# < 0 Then GoTo lblNextRec
                        
                        'Get the Maximum based on the Service Months and Entitlement Code
                        'dblNewMax# = Get_HourlyEntitle_Maximum(dblServiceYears#)
                        'Which range of the service month the employee falls in and if the Vacation Pay % exists for that range
                        dblNewMax# = 0
                        rsHREntMstDtl.MoveFirst
                        Do While Not rsHREntMstDtl.EOF
                            If rsHREntMstDtl("EH_BMONTH") = "" And rsHREntMstDtl("EH_EMONTH") = "" Then Exit Do
                            
                            If IsNumeric(rsHREntMstDtl("EH_BMONTH")) And rsHREntMstDtl("EH_EMONTH") = "" Then
                                If dblServiceYears# >= CDbl(rsHREntMstDtl("EH_BMONTH")) Then
                                    dblNewMax# = IIf(IsNull(rsHREntMstDtl("EH_MAX")), 0, rsHREntMstDtl("EH_MAX"))
                                    Exit Do
                                End If
                            End If
                            If IsNumeric(rsHREntMstDtl("EH_BMONTH")) And IsNumeric(rsHREntMstDtl("EH_EMONTH")) Then
                                If dblServiceYears# >= CDbl(rsHREntMstDtl("EH_BMONTH")) And dblServiceYears# <= CDbl(rsHREntMstDtl("EH_EMONTH")) Then
                                    dblNewMax# = IIf(IsNull(rsHREntMstDtl("EH_MAX")), 0, rsHREntMstDtl("EH_MAX"))
                                    Exit Do
                                End If
                            End If
                            
                            rsHREntMstDtl.MoveNext
                        Loop
                        
                        'Calculated Entitlement cannot exceed Maximum
                        If dblNewMax# > 0 Then
                            If dblNewEntitle# > dblNewMax# Then
                                dblNewEntitle# = dblNewMax#
                            End If
                        End If
                    
                        'Accumulate to the existing entitlement or replace the existing entitlement
                        If xUpdMethod = "A" Then
                            dblEntitleUpd = dblEntitle# + dblNewEntitle
                        Else
                            dblEntitleUpd = dblNewEntitle
                        End If
                        
                        'Update respective Hourly Entitlement records
                        If xUpdMethod = "A" Then
                            'ACCUMULATE METHOD
                            
                            If NumRec > 0 Then  'if accumulate and found duplicate record
                                'For Flex Codes, do not update the entitlement to the new entitlement value because the entitlement gets updated using the
                                'Flex Attendance record
                                If Right(rsHREntMst("EH_HETYPE"), 1) <> "+" Then
                                    SQLQ = "UPDATE HRENTHRS "
                                    SQLQ = SQLQ & " SET HE_ENTITLE = " & dblEntitleUpd & " "
                                    SQLQ = SQLQ & " WHERE HE_EMPNBR = " & snapHREntitle("ED_EMPNBR")
                                    SQLQ = SQLQ & " AND HRENTHRS.HE_TYPE = '" & rsHREntMst("EH_HETYPE") & "' "
                                    SQLQ = SQLQ & " AND HRENTHRS.HE_FDATE = " & Date_SQL(rsHREntMst("EH_FDATE"))
                                    SQLQ = SQLQ & " AND HRENTHRS.HE_TDATE = " & Date_SQL(rsHREntMst("EH_TDATE"))
                                    gdbAdoIhr001.Execute (SQLQ)
                                End If
                                
                                'For Flex Codes, simply add the actual new entitlement earned to the Accrual File and not update
                                If Right(rsHREntMst("EH_HETYPE"), 1) = "+" Then
                                    Call Append_Accrual(snapHREntitle("ED_EMPNBR"), rsHREntMst("EH_HETYPE"), dlpDateRange(1).Text, dblNewEntitle, "A", "Mass added the Hourly Entitlement")
                                Else
                                    'This will be a add because it's Accumulate method
                                    'Call Append_Accrual(snapHREntitle("ED_EMPNBR"), rsHREntMst("EH_HETYPE"), dlpDateRange(1).Text, dblEntitleUpd - oldEntitleUpd, "U", "Mass changed the existing Hourly Entitlement") 'Ticket #17924
                                    Call Append_Accrual(snapHREntitle("ED_EMPNBR"), rsHREntMst("EH_HETYPE"), dlpDateRange(1).Text, dblNewEntitle, "A", "Mass added the Hourly Entitlement")
                                End If
                            Else
                                'Ticket #17924 - If Flex logic (+) then update the existing Flex code hourly entitlement record instead
                                'of adding a new record.
                                If Right(rsHREntMst("EH_HETYPE"), 1) = "+" Then
                                    If fTablHREMP.State <> adStateClosed Then fTablHREMP.Close
                                    SQLQ = "SELECT * FROM HRENTHRS "
                                    SQLQ = SQLQ & " WHERE HE_EMPNBR = " & snapHREntitle("ED_EMPNBR")
                                    SQLQ = SQLQ & " AND HE_TYPE = '" & rsHREntMst("EH_HETYPE") & "'"
                                    SQLQ = SQLQ & " ORDER BY HE_FDATE DESC"
                                    fTablHREMP.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not fTablHREMP.EOF Then
                                        fTablHREMP.MoveFirst
                                    Else
                                        fTablHREMP.AddNew
                                        fTablHREMP("HE_PREV") = 0
                                    End If
                                Else
                                    fTablHREMP.AddNew     'if accumulate and no duplicate record
                                    fTablHREMP("HE_PREV") = 0
                                End If
                                
                                fTablHREMP("HE_EMPNBR") = snapHREntitle("ED_EMPNBR")
                                fTablHREMP("HE_COMPNO") = "001"
                                fTablHREMP("HE_TYPE_TABL") = "ADRE"
                                fTablHREMP("HE_TYPE") = rsHREntMst("EH_HETYPE")
                                fTablHREMP("HE_FDATE") = rsHREntMst("EH_FDATE")
                                fTablHREMP("HE_TDATE") = rsHREntMst("EH_TDATE")
                                fTablHREMP("HE_ENTITLE") = dblEntitleUpd
                                fTablHREMP("HE_COE") = True
                                fTablHREMP("HE_DHRS") = snapHREntitle("ED_DHRS")
                                fTablHREMP("HE_LDATE") = Now
                                fTablHREMP("HE_LTIME") = Time$
                                fTablHREMP("HE_LUSER") = glbUserID
                                fTablHREMP.Update
                                            
                                Call Append_Accrual(snapHREntitle("ED_EMPNBR"), rsHREntMst("EH_HETYPE"), dlpDateRange(1).Text, dblEntitleUpd, "A", "Mass added the Hourly Entitlement")
                            End If
                        Else
                            'REPLACE METHOD
                            
                            'Ticket #17924 - If Flex logic (+) then update the existing Flex code hourly entitlement record instead
                            'of adding a new record.
                            If Right(rsHREntMst("EH_HETYPE"), 1) = "+" Then
                                If fTablHREMP.State <> adStateClosed Then fTablHREMP.Close
                                SQLQ = "SELECT * FROM HRENTHRS "
                                SQLQ = SQLQ & " WHERE HE_EMPNBR = " & snapHREntitle("ED_EMPNBR")
                                SQLQ = SQLQ & " AND HE_TYPE = '" & rsHREntMst("EH_HETYPE") & "'"
                                SQLQ = SQLQ & " ORDER BY HE_FDATE DESC"
                                fTablHREMP.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not fTablHREMP.EOF Then
                                    fTablHREMP.MoveFirst
                                Else
                                    fTablHREMP.AddNew
                                    fTablHREMP("HE_PREV") = 0
                                End If
                            Else
                                'Ticket #18559 - Jerry does not want the Previous to be replaced with 0 after the rollover which
                                'creates a new record on the Hourly Entitlement screen. In which case we cannot delete an existing
                                'Hourly Entitlement record but instead update the values.
                                'SQLQ$ = "DELETE FROM HRENTHRS "
                                'SQLQ = SQLQ & " WHERE HE_EMPNBR = " & SnapAddEntitle("ED_EMPNBR")
                                'SQLQ = SQLQ & " AND HE_TYPE = '" & clpCode(2).Text & "'"
                                'SQLQ = SQLQ & " AND HE_TDATE = " & Date_SQL(dlpTo.Text)
                                'gdbAdoIhr001.Execute SQLQ
                                
                                If fTablHREMP.State <> adStateClosed Then fTablHREMP.Close
                                SQLQ = "SELECT * FROM HRENTHRS "
                                SQLQ = SQLQ & " WHERE HE_EMPNBR = " & snapHREntitle("ED_EMPNBR")
                                SQLQ = SQLQ & " AND HE_TYPE = '" & rsHREntMst("EH_HETYPE") & "'"
                                SQLQ = SQLQ & " AND HE_FDATE = " & Date_SQL(rsHREntMst("EH_FDATE"))
                                SQLQ = SQLQ & " AND HE_TDATE = " & Date_SQL(rsHREntMst("EH_TDATE"))
                                fTablHREMP.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If fTablHREMP.EOF Then
                                    fTablHREMP.AddNew
                                    fTablHREMP("HE_PREV") = 0
                                End If
                            End If
                            
                            'fTablHREMP.AddNew
                            
                            fTablHREMP("HE_EMPNBR") = snapHREntitle("ED_EMPNBR")
                            fTablHREMP("HE_COMPNO") = "001"
                            fTablHREMP("HE_TYPE_TABL") = "ADRE"
                            fTablHREMP("HE_TYPE") = rsHREntMst("EH_HETYPE")
                            fTablHREMP("HE_FDATE") = rsHREntMst("EH_FDATE")
                            fTablHREMP("HE_TDATE") = rsHREntMst("EH_TDATE")
                            fTablHREMP("HE_ENTITLE") = dblEntitleUpd
                            fTablHREMP("HE_COE") = True
                            fTablHREMP("HE_DHRS") = snapHREntitle("ED_DHRS")
                            fTablHREMP("HE_LDATE") = Now
                            fTablHREMP("HE_LTIME") = Time$
                            fTablHREMP("HE_LUSER") = glbUserID
                            fTablHREMP.Update
                            
                            If NumRec > 0 Then  'if accumulate and found duplicate record
                                'Call Append_Accrual(SnapAddEntitle("ED_EMPNBR"), clpCode(2).Text, dlpFrom.Text, dblEntitleUpd - oldEntitleUpd, "U", "Mass modified the Hourly Entitlement")
                                Call Append_Accrual(snapHREntitle("ED_EMPNBR"), rsHREntMst("EH_HETYPE"), dlpDateRange(1).Text, dblEntitleUpd - oldEntitleUpd, "U", "Mass modified the Hourly Entitlement")
                            Else
                                'Call Append_Accrual(SnapAddEntitle("ED_EMPNBR"), clpCode(2).Text, dlpFrom.Text, dblEntitleUpd, "A", "Mass added the Hourly Entitlement")
                                Call Append_Accrual(snapHREntitle("ED_EMPNBR"), rsHREntMst("EH_HETYPE"), dlpDateRange(1).Text, dblEntitleUpd, "A", "Mass added the Hourly Entitlement")
                            End If
                        End If
                    
                        'Ticket #22682 - Release 8.0: Jerry said not to check for duplicate, simply add new Attendance record, even
                        'though it is a duplicate record.
                        'Ticket #17924 - Begin
                        'If the Entitlement Code is suffixed with + then insert an Attendance record
                        'for the Hourly Entitlement earned - helps in the Recalculate function
                        If Right(rsHREntMst("EH_HETYPE"), 1) = "+" Then
                            'Add Record in Attendance screen
                            'Ticket #22682 - Release 8.0: Do not check for duplicates
                            SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE 1 = 2"
                            'SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR =" & SnapAddEntitle("ED_EMPNBR")
                            'SQLQ = SQLQ & " AND AD_REASON = '" & clpCode(2).Text & "'"
                            'Ticket #18550 - Attendance record date cannot be prior to hire date
                            'If CVDate(SnapAddEntitle("ED_DOH")) > CVDate(dlpFrom.Text) Then
                            '    SQLQ = SQLQ & " AND AD_DOA =" & Date_SQL(SnapAddEntitle("ED_DOH"))
                            'Else
                            '    SQLQ = SQLQ & " AND AD_DOA =" & Date_SQL(dlpFrom.Text)
                            'End If
                            rzAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            'If rzAttend.EOF Then
                                rzAttend.AddNew
                            'End If
                            rzAttend("AD_COMPNO") = "001"
                            rzAttend("AD_EMPNBR") = snapHREntitle("ED_EMPNBR")
                            rzAttend("AD_DOA") = dlpDateRange(1).Text
                            rzAttend("AD_REASON") = rsHREntMst("EH_HETYPE")
                            If xUpdMethod = "A" Then
                                'Accumulate = Update Attendance with the additional new entitlement
                                rzAttend("AD_HRS") = dblNewEntitle
                            Else
                                'Replace = Update with the difference so the Hourly Entitlement balance is the New entitlement
                                rzAttend("AD_HRS") = dblNewEntitle - oldEntitleUpd
                            End If
                            
                    
                            SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID,ED_ORG,ED_GLNO,ED_DOH FROM HREMP WHERE ED_EMPNBR = " & snapHREntitle("ED_EMPNBR")
                            rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                            If Not rsHREmp.EOF Then
                                rzAttend("AD_PAYROLL_ID") = rsHREmp("ED_PAYROLL_ID")
                                rzAttend("AD_GLNO") = rsHREmp("ED_GLNO")
                                rzAttend("AD_ORG") = rsHREmp("ED_ORG")
                                
                                'Ticket #18550 - Attendance record date cannot be prior to hire date
                                If CVDate(rsHREmp("ED_DOH")) > CVDate(dlpDateRange(1).Text) Then
                                    rzAttend("AD_DOA") = rsHREmp("ED_DOH")
                                End If
                            End If
                            rsHREmp.Close
                    
                            SQLQ = "SELECT SH_EMPNBR, SH_SALARY,SH_SALCD FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & snapHREntitle("ED_EMPNBR")
                            rsCurJobSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                            If Not rsCurJobSal.BOF Then
                                If rsCurJobSal("SH_SALARY") > 0 Then
                                    rzAttend("AD_SALARY") = rsCurJobSal("SH_SALARY")
                                    rzAttend("AD_SALCD") = rsCurJobSal("SH_SALCD")
                                End If
                            End If
                            rsCurJobSal.Close
                            Set rsCurJobSal = Nothing
                    
                            SQLQ = "SELECT JH_EMPNBR,JH_CURRENT,JH_JOB,JH_DHRS,JH_WHRS FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & snapHREntitle("ED_EMPNBR")
                            rsCurJobSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                            If Not rsCurJobSal.EOF Then
                                rzAttend("AD_JOB") = rsCurJobSal("JH_JOB")
                                rzAttend("AD_DHRS") = rsCurJobSal("JH_DHRS")
                                rzAttend("AD_WHRS") = rsCurJobSal("JH_WHRS")
                            End If
                            rsCurJobSal.Close
                            Set rsCurJobSal = Nothing
                    
                            'Ticket #18550
                            'rzAttend("AD_COMM") = "Entitlement earned for the period: " & dlpFrom.Text & " to " & dlpTo.Text & "."
                            rzAttend("AD_COMM") = "Accrued Hours for the period: " & dlpDateRange(0).Text & " to " & dlpDateRange(1).Text & "."
                            rzAttend("AD_LDATE") = Date
                            rzAttend("AD_LUSER") = glbUserID
                            rzAttend("AD_LTIME") = Time$
                            rzAttend.Update
                            rzAttend.Close
                            Set rzAttend = Nothing
                        End If
                        'Ticket #17924 - End
                        
                        DoEvents
                        xKey = snapHREntitle("ED_EMPNBR")
                        xKey = xKey & "|" & Format(dlpDateRange(0).Text, "dd-mmm-yyyy")
                        xKey = xKey & "|" & Format(dlpDateRange(1).Text, "dd-mmm-yyyy")
                        xKey = xKey & "|" & rsHREntMst("EH_HETYPE")
                        xKey = xKey & "|" & dblEntitleUpd
                        xKey = xKey & "|" & Format(Date, "dd-mmm-yyyy") 'Transaction Date
                        Call Entitlements_Master_Integration(xKey, 0)
                        
                        DoEvents
                    
lblNextRec:
                        DoEvents
                    
                        snapHREntitle.MoveNext  'Next Employee
                    Wend
                End If
                rsHREntMstDtl.Close
                Set rsHREntMstDtl = Nothing
            End If
            
            Update_HourlyEntitlement = True
            MDIMain.panHelp(0).FloodType = 0
        
            snapHREntitle.Close
            Set snapHREntitle = Nothing
            
            rsHREntMst.MoveNext     'Next Hourly Entitlement Rule
        Loop
    End If
    rsHREntMst.Close
    Set rsHREntMst = Nothing
    
    Update_HourlyEntitlement = True
    
    Screen.MousePointer = DEFAULT
        
Exit Function

Update_HourlyEntitlement_Err:
'These errors are:
'13=type mismatch
'94=invalid use of null
'3018=couln't find field 'item'
If Err = 13 Or Err = 94 Or Err = 3018 Then
   ' MsgBox "Err:" & Str(Err) & Chr(10) & Error$ & Chr(10) & " modUpdateSelection" & Chr(10) & "FORM:FUENTITL.FRM"
    'commented out by RAUBREY 5/20/97
    Err = 0
    Resume Next
End If

Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "UpdateHrsEntitle", "HREMP", "Close Pay Period")

Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    'Rollback
    Resume Next
Else
    Unload Me
End If

End Function

Private Function Total_Worked_Hours(xEmpNbr, xCode)
    Dim rsAttend As New ADODB.Recordset
    Dim rsHRAttMatrix As New ADODB.Recordset
    Dim SQLQ As String
    Dim SQLQMatrix As String
    Dim xTotHrs As Double
    
    'FYI: A table field value FT,PT,CA would be returned as 'FT','PT','CA' when using the following combination of REPLACE and QUOTENAME functions in TSQL
    'SELECT REPLACE(QUOTENAME(AM_PT,''''),',',''','''),* FROM HRATT_MATRIX
    
    Total_Worked_Hours = 0
    
    xTotHrs = 0
    
    'Retrieve the HR Attendance Matrix based on Entitlement Code entitled
    Select Case xCode
        Case "VA+"
            SQLQMatrix = "SELECT * FROM HRATT_MATRIX WHERE AM_VAC_HRS <> 0"
        Case "SK+"
            SQLQMatrix = "SELECT * FROM HRATT_MATRIX WHERE AM_ABSENT_HRS <> 0"
        Case "PD+"
            SQLQMatrix = "SELECT * FROM HRATT_MATRIX WHERE AM_EXTRA_HRS <> 0"
        Case "PT+"
            SQLQMatrix = "SELECT * FROM HRATT_MATRIX WHERE AM_REG_HRS <> 0"
        Case "IW+"
            SQLQMatrix = "SELECT * FROM HRATT_MATRIX WHERE AM_INCID <> 0"
    End Select
    
    'Attendance
    rsHRAttMatrix.Open SQLQMatrix, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHRAttMatrix.EOF Then
        rsHRAttMatrix.MoveFirst
        
        'For each Attendance Matrix record's Reason Code, sum the hours from Attendance for the employee if the employee is part of the Attendance Matrix Category
        Do While Not rsHRAttMatrix.EOF
            'Get Total Hours for the Attendance Matrix Reason Code
            SQLQ = "SELECT SUM(AD_HRS) AS TOT_HRS FROM HR_ATTENDANCE"
            SQLQ = SQLQ & " WHERE AD_EMPNBR = " & xEmpNbr
            SQLQ = SQLQ & " AND (AD_DOA >= " & Date_SQL(dlpDateRange(0).Text)
            SQLQ = SQLQ & " AND AD_DOA <= " & Date_SQL(dlpDateRange(1).Text) & ")"
            SQLQ = SQLQ & " AND AD_REASON = '" & rsHRAttMatrix("AM_REASON") & "'"
            'Also if the employee's Category matches the Attendance Matrix
            SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_PT IN ('" & Replace(rsHRAttMatrix("AM_PT"), ",", "','") & "'))"
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
            
            rsHRAttMatrix.MoveNext
        Loop
    End If
    rsHRAttMatrix.Close
    Set rsHRAttMatrix = Nothing
    
'    'Attendance History
'    rsHRAttMatrix.Open SQLQMatrix, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    If Not rsHRAttMatrix.EOF Then
'        rsHRAttMatrix.MoveFirst
'
'        'For each Attendance Matrix record's Reason Code, sum the hours from Attendance for the employee if the employee is part of the Attendance Matrix Category
'        Do While Not rsHRAttMatrix.EOF
'            'Get Total Hours for the Attendance Matrix Reason Code
'            SQLQ = "SELECT SUM(AH_HRS) AS TOT_HRS FROM HR_ATTENDANCE_HISTORY"
'            SQLQ = SQLQ & " WHERE AH_EMPNBR = " & xEmpNbr
'            SQLQ = SQLQ & " AND (AH_DOA >= " & Date_SQL(dlpDateRange(0).Text)
'            SQLQ = SQLQ & " AND AH_DOA <= " & Date_SQL(dlpDateRange(1).Text) & ")"
'            SQLQ = SQLQ & " AND AH_REASON = '" & rsHRAttMatrix("AM_REASON") & "'"
'            'Also if the employee's Category matches the Attendance Matrix
'            SQLQ = SQLQ & " AND AH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_PT IN ('" & Replace(rsHRAttMatrix("AM_PT"), ",", "','") & "'))"
'            SQLQ = SQLQ & " GROUP BY AH_EMPNBR"
'            rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'            If Not rsAttend.EOF Then
'                rsAttend.MoveFirst
'
'                'Sum Total Hours
'                If rsAttend("TOT_HRS") > 0 Then
'                    xTotHrs = xTotHrs + rsAttend("TOT_HRS")
'                End If
'
'            End If
'            rsAttend.Close
'            Set rsAttend = Nothing
'            rsHRAttMatrix.MoveNext
'        Loop
'    End If
'    rsHRAttMatrix.Close
'    Set rsHRAttMatrix = Nothing

    Total_Worked_Hours = xTotHrs

End Function

Private Function Close_Pay_Period()
    Dim SQLQ As String
    
    On Error GoTo Close_Pay_Period_Err
    
    Close_Pay_Period = False
    
    SQLQ = "UPDATE HR_PAYPERIOD SET PP_UPLOADED = 1 WHERE PP_START = " & Date_SQL(dlpDateRange(0).Text) & " AND PP_END = " & Date_SQL(dlpDateRange(1).Text) & " AND PP_YEAR = " & Year(Now)
    gdbAdoIhr001.Execute (SQLQ)
    
    Close_Pay_Period = True
    
    Screen.MousePointer = DEFAULT
    
Exit Function

Close_Pay_Period_Err:
'These errors are:
'13=type mismatch
'94=invalid use of null
'3018=couln't find field 'item'
If Err = 13 Or Err = 94 Or Err = 3018 Then
   ' MsgBox "Err:" & Str(Err) & Chr(10) & Error$ & Chr(10) & " modUpdateSelection" & Chr(10) & "FORM:FUENTITL.FRM"
    'commented out by RAUBREY 5/20/97
    Err = 0
    Resume Next
End If

Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Closing Pay Period", "HR_PAYPERIOD", "Close Pay Period")

Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    'Rollback
    Resume Next
Else
    Unload Me
End If
    
End Function
