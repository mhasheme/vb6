VERSION 5.00
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmForfeitHrs 
   Caption         =   "Forfeit Hours"
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6255
      Begin VB.CheckBox chkYearEnd 
         Alignment       =   1  'Right Justify
         Caption         =   "Year End"
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Top             =   840
         Width           =   1155
      End
      Begin INFOHR_Controls.DateLookup dlpMnthEndDate 
         Height          =   285
         Left            =   3120
         TabIndex        =   0
         Tag             =   "41-Month End Date"
         Top             =   360
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Forfeit Hours for the Month Ending"
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
         Left            =   120
         TabIndex        =   5
         Tag             =   "41-Date Terminated"
         Top             =   405
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
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
      TabIndex        =   2
      Top             =   1800
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
      TabIndex        =   3
      Top             =   1800
      Width           =   1200
   End
End
Attribute VB_Name = "frmForfeitHrs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Sub cmdOK_Click()
Dim Response%

'Validate Month End Date
If Len(Trim(dlpMnthEndDate.Text)) = 0 Then
    MsgBox "Forfeit Hours for the Month Ending cannot be blank", vbExclamation
    dlpMnthEndDate.SetFocus
    Exit Sub
ElseIf Not IsDate(dlpMnthEndDate) Then
    MsgBox "Invalid Forfeit Hours for the Month Ending Date", vbExclamation
    dlpMnthEndDate.SetFocus
    Exit Sub
ElseIf CVDate(dlpMnthEndDate) <> CVDate(MonthLastDate(dlpMnthEndDate)) Then
    MsgBox "Forfeit Hours for the Month Ending Date is not the End of Month Date", vbExclamation
    dlpMnthEndDate.SetFocus
    Exit Sub
Else
    'Prompt to confirm
    If chkYearEnd Then
        Response% = MsgBox("Forfeit Hours will be calculated for the Month of " & Format(dlpMnthEndDate, "mmmm, yyyy") & " and for the Year." & vbCrLf & "Do you wish to proceed?", vbYesNo + vbQuestion, "info:HR - Forfeit Hours")
    Else
        Response% = MsgBox("Forfeit Hours will be calculated for the Month of " & Format(dlpMnthEndDate, "mmmm, yyyy") & "." & vbCrLf & "Do you wish to proceed?", vbYesNo + vbQuestion, "info:HR - Forfeit Hours")
    End If
    
    If Response% = IDNO Then
        dlpMnthEndDate.SetFocus
        Exit Sub
    End If
End If

Screen.MousePointer = HOURGLASS

Call Forfeited_Hours_Computation

Screen.MousePointer = DEFAULT

If chkYearEnd Then
    MsgBox "Forfeit Hours calculation complete for the Month of " & Format(dlpMnthEndDate, "mmmm, yyyy") & " and for the Year.", vbInformation, "info:HR - Forfeit Hours"
Else
    MsgBox "Forfeit Hours calculation complete for the Month of " & Format(dlpMnthEndDate, "mmmm, yyyy") & ".", vbInformation, "info:HR - Forfeit Hours"
End If

Unload Me

End Sub

Private Sub dlpMnthEndDate_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub Forfeited_Hours_Computation()
    'Compute the Forfeited Hour for Four Villages
    'Get the total OTs for the month
    'Get the total CTs except CTF for the month
    'Get the Max OT allowed from Overtime Master
    'Find the difference between OTs and CTs - Outstanding OT for the month
    'If Outstanding OT exceeds the Maximum then Forfeit those hours as CTF for the month
        '- create attendance record.
    
    Dim rsAttend As New ADODB.Recordset
    Dim rsEmp As New ADODB.Recordset
    Dim rsTABL As New ADODB.Recordset
    Dim rsCurPos As New ADODB.Recordset
    Dim rsCurSal As New ADODB.Recordset
    Dim rsOTMaster As New ADODB.Recordset
    Dim SQLQ As String
    Dim xMonthOT As Double
    Dim xMonthCT As Double
    Dim xMaxOT As Double
    Dim xOsOT As Double
    Dim xForfeitHrs As Double
    Dim xYearOT As Double
    Dim xYearCT As Double
    
    'For each employee compute the CTF for the month
    SQLQ = "SELECT ED_EMPNBR FROM HREMP"
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsEmp.EOF
        'Get the OT Entitlement Period so the OT outstanding can be calculated from the beginning to
        'upto current Month End selected by the user
        SQLQ = "SELECT OT_EFDATE,OT_ETDATE FROM HR_OVERTIME_BANK "
        SQLQ = SQLQ & " WHERE OT_EMPNBR = " & rsEmp("ED_EMPNBR")
        rsOTMaster.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsOTMaster.EOF Then
        
            'Get total OTs for the month
            SQLQ = "SELECT SUM(AD_HRS) AS TOT_OT FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsEmp("ED_EMPNBR")
            'SQLQ = SQLQ & " AND MONTH(AD_DOA) = " & month(dlpMnthEndDate) & " AND YEAR(AD_DOA) = " & Year(dlpMnthEndDate)
            SQLQ = SQLQ & " AND AD_DOA >= " & Date_SQL(rsOTMaster("OT_EFDATE")) & " AND AD_DOA <= " & Date_SQL(Format(dlpMnthEndDate, "mm/dd/yyyy"))
            SQLQ = SQLQ & " AND LEFT(AD_REASON, 2) = 'OT'"
            rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsAttend.EOF Then
                xMonthOT = IIf(IsNull(rsAttend("TOT_OT")), 0, rsAttend("TOT_OT"))
            Else
                xMonthOT = 0
            End If
            rsAttend.Close
            Set rsAttend = Nothing
            
            'Get total CTs (except CTF) for the month
            SQLQ = "SELECT SUM(AD_HRS) AS TOT_CT FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsEmp("ED_EMPNBR")
            'SQLQ = SQLQ & " AND MONTH(AD_DOA) = " & month(dlpMnthEndDate) & " AND YEAR(AD_DOA) = " & Year(dlpMnthEndDate)
            SQLQ = SQLQ & " AND AD_DOA >= " & Date_SQL(rsOTMaster("OT_EFDATE")) & " AND AD_DOA <= " & Date_SQL(Format(dlpMnthEndDate, "mm/dd/yyyy"))
            SQLQ = SQLQ & " AND LEFT(AD_REASON, 2) = 'CT'" 'AND AD_REASON <> 'CTF'"
            rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsAttend.EOF Then
                xMonthCT = IIf(IsNull(rsAttend("TOT_CT")), 0, rsAttend("TOT_CT"))
            Else
                xMonthCT = 0
            End If
            rsAttend.Close
            Set rsAttend = Nothing
            
            'Get the Max OT
            xMaxOT = Get_Maximum_OT_Allowed(rsEmp("ED_EMPNBR"), dlpMnthEndDate)
            
            
            If xMaxOT <> 0 Then
                'Calculate Outstanding OT
                xOsOT = xMonthOT - xMonthCT
                
                'Calculate the Forfeited Hours
                If xOsOT > xMaxOT Then
                    xForfeitHrs = xOsOT - xMaxOT
                Else
                    xForfeitHrs = 0
                End If
                
                'Add/Update Attendance with Forfeited Hours (CTF)
                If xForfeitHrs > 0 Then
                    SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsEmp("ED_EMPNBR")
                    SQLQ = SQLQ & " AND AD_DOA = " & Date_SQL(MonthLastDate(dlpMnthEndDate))
                    SQLQ = SQLQ & " AND AD_REASON = 'CTF'"
                    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If rsAttend.EOF Then
                        'No CTF record found for the month end
                    
                        'Make sure CTF code exists if not then add the code
                        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_KEY = 'CTF' "
                        rsTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If rsTABL.EOF Then
                            rsTABL.AddNew
                            rsTABL("TB_COMPNO") = "001"
                            rsTABL("TB_NAME") = "ADRE"
                            rsTABL("TB_KEY") = "CTF"
                            rsTABL("TB_DESC") = "FORFEITED HOURS"
                            rsTABL("TB_LDATE") = Date
                            rsTABL("TB_LTIME") = Time$
                            rsTABL("TB_LUSER") = glbUserID
                            rsTABL.Update
                        End If
                        rsTABL.Close
                        
                        'Add a new record with CTF hours
                        rsAttend.AddNew
                        rsAttend("AD_COMPNO") = "001"
                        rsAttend("AD_EMPNBR") = rsEmp("ED_EMPNBR")
                        
                        'Update with Salary info.
                        SQLQ = "SELECT SH_EMPNBR, SH_SALARY,SH_SALCD FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & rsEmp("ED_EMPNBR")
                        rsCurSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                        If Not rsCurSal.BOF Then
                            If rsCurSal("SH_SALARY") > 0 Then
                                rsAttend("AD_SALARY") = rsCurSal("SH_SALARY")
                                rsAttend("AD_SALCD") = rsCurSal("SH_SALCD")
                            End If
                        End If
                        rsCurSal.Close
                        Set rsCurSal = Nothing
                        
                        'Update with Position info.
                        SQLQ = "SELECT JH_EMPNBR,JH_CURRENT,JH_JOB,JH_DHRS,JH_WHRS,JH_REPTAU,JH_PAYROLL_ID,JH_SHIFT,JH_GLNO,JH_ORG FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & rsEmp("ED_EMPNBR")
                        rsCurPos.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                        If Not rsCurPos.EOF Then
                            rsAttend("AD_JOB") = rsCurPos("JH_JOB")
                            rsAttend("AD_DHRS") = rsCurPos("JH_DHRS")
                            rsAttend("AD_WHRS") = rsCurPos("JH_WHRS")
                            rsAttend("AD_SUPER") = rsCurPos("JH_REPTAU")
                            rsAttend("AD_PAYROLL_ID") = rsCurPos("JH_PAYROLL_ID")
                            rsAttend("AD_SHIFT") = rsCurPos("JH_SHIFT")
                            rsAttend("AD_GLNO") = rsCurPos("JH_GLNO")
                            rsAttend("AD_ORG") = rsCurPos("JH_ORG")
                        End If
                        rsCurPos.Close
                        Set rsCurPos = Nothing
                    End If
                    
                    'Update with the CTF hours
                    rsAttend("AD_DOA") = CVDate(MonthLastDate(dlpMnthEndDate))
                    rsAttend("AD_REASON") = "CTF"
                    rsAttend("AD_HRS") = xForfeitHrs
                    rsAttend("AD_COMM") = "Monthly Forfeited Hours."
                    rsAttend("AD_LUSER") = glbUserID
                    rsAttend("AD_LDATE") = Date
                    rsAttend("AD_LTIME") = Time$
                    rsAttend("AD_SOURCE") = "IHRFOR"
                    rsAttend.Update
                    
                    rsAttend.Close
                    Set rsAttend = Nothing
                End If
                
                'Forfeit for the Year?
                If chkYearEnd Then
                    'Ticket #23641 - No need to open the recordset again - the values needed has already been retrieved
                    'Forfeit Hours for the year as well
                    'Get the Overtime Entitlement Period
                    'SQLQ = "SELECT OT_EFDATE,OT_ETDATE FROM HR_OVERTIME_BANK "
                    'SQLQ = SQLQ & " WHERE OT_EMPNBR = " & rsEmp("ED_EMPNBR")
                    'OT_EFDATE
                    'OT_ETDATE
                    'rsOTMaster.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsOTMaster.EOF Then
                        'Get total OT for Entitlement Period = minus the last date of the year?
                        SQLQ = "SELECT SUM(AD_HRS) AS TOT_OT FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsEmp("ED_EMPNBR")
                        SQLQ = SQLQ & " AND AD_DOA >= " & Date_SQL(rsOTMaster("OT_EFDATE")) & " AND AD_DOA <= " & Date_SQL(rsOTMaster("OT_ETDATE"))
                        SQLQ = SQLQ & " AND LEFT(AD_REASON, 2) = 'OT'"
                        rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsAttend.EOF Then
                            xYearOT = IIf(IsNull(rsAttend("TOT_OT")), 0, rsAttend("TOT_OT"))
                        Else
                            xYearOT = 0
                        End If
                        rsAttend.Close
                        Set rsAttend = Nothing
                        
                        'Get total CT for Entitlement Period = minus the last date of the year?
                        SQLQ = "SELECT SUM(AD_HRS) AS TOT_CT FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsEmp("ED_EMPNBR")
                        SQLQ = SQLQ & " AND AD_DOA >= " & Date_SQL(rsOTMaster("OT_EFDATE")) & " AND AD_DOA <= " & Date_SQL(rsOTMaster("OT_ETDATE"))
                        SQLQ = SQLQ & " AND LEFT(AD_REASON, 2) = 'CT'"
                        rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsAttend.EOF Then
                            xYearCT = IIf(IsNull(rsAttend("TOT_CT")), 0, rsAttend("TOT_CT"))
                        Else
                            xYearCT = 0
                        End If
                        rsAttend.Close
                        Set rsAttend = Nothing
                        
                        'Get Outstanding OT
                        xForfeitHrs = xYearOT - xYearCT
                                        
                        'Add CTF Record for Year End for the Outstanding OT - all hours forfeited
                        If xForfeitHrs > 0 Then
                            SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsEmp("ED_EMPNBR")
                            SQLQ = SQLQ & " AND AD_DOA = " & Date_SQL(rsOTMaster("OT_ETDATE"))
                            SQLQ = SQLQ & " AND AD_REASON = 'CTF'"
                            rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If rsAttend.EOF Then
                                'No CTF record found for the year end
                            
                                'Make sure CTF code exists if not then add the code
                                SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_KEY = 'CTF' "
                                rsTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If rsTABL.EOF Then
                                    rsTABL.AddNew
                                    rsTABL("TB_COMPNO") = "001"
                                    rsTABL("TB_NAME") = "ADRE"
                                    rsTABL("TB_KEY") = "CTF"
                                    rsTABL("TB_DESC") = "FORFEITED HOURS"
                                    rsTABL("TB_LDATE") = Date
                                    rsTABL("TB_LTIME") = Time$
                                    rsTABL("TB_LUSER") = glbUserID
                                    rsTABL.Update
                                End If
                                rsTABL.Close
                                
                                'Add a new record with CTF hours
                                rsAttend.AddNew
                                rsAttend("AD_COMPNO") = "001"
                                rsAttend("AD_EMPNBR") = rsEmp("ED_EMPNBR")
                                
                                'Update with Salary info.
                                SQLQ = "SELECT SH_EMPNBR, SH_SALARY,SH_SALCD FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & rsEmp("ED_EMPNBR")
                                rsCurSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                                If Not rsCurSal.BOF Then
                                    If rsCurSal("SH_SALARY") > 0 Then
                                        rsAttend("AD_SALARY") = rsCurSal("SH_SALARY")
                                        rsAttend("AD_SALCD") = rsCurSal("SH_SALCD")
                                    End If
                                End If
                                rsCurSal.Close
                                Set rsCurSal = Nothing
                                
                                'Update with Position info.
                                SQLQ = "SELECT JH_EMPNBR,JH_CURRENT,JH_JOB,JH_DHRS,JH_WHRS,JH_REPTAU,JH_PAYROLL_ID,JH_SHIFT,JH_GLNO,JH_ORG FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & rsEmp("ED_EMPNBR")
                                rsCurPos.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                                If Not rsCurPos.EOF Then
                                    rsAttend("AD_JOB") = rsCurPos("JH_JOB")
                                    rsAttend("AD_DHRS") = rsCurPos("JH_DHRS")
                                    rsAttend("AD_WHRS") = rsCurPos("JH_WHRS")
                                    rsAttend("AD_SUPER") = rsCurPos("JH_REPTAU")
                                    rsAttend("AD_PAYROLL_ID") = rsCurPos("JH_PAYROLL_ID")
                                    rsAttend("AD_SHIFT") = rsCurPos("JH_SHIFT")
                                    rsAttend("AD_GLNO") = rsCurPos("JH_GLNO")
                                    rsAttend("AD_ORG") = rsCurPos("JH_ORG")
                                End If
                                rsCurPos.Close
                                Set rsCurPos = Nothing
                            End If
                            
                            'Update with the CTF hours
                            rsAttend("AD_DOA") = CVDate(rsOTMaster("OT_ETDATE"))
                            rsAttend("AD_REASON") = "CTF"
                            rsAttend("AD_HRS") = xForfeitHrs
                            rsAttend("AD_COMM") = "End of the year Forfeited Hours."
                            rsAttend("AD_LUSER") = glbUserID
                            rsAttend("AD_LDATE") = Date
                            rsAttend("AD_LTIME") = Time$
                            rsAttend("AD_SOURCE") = "IHRFOR"
                            rsAttend.Update
                            
                            rsAttend.Close
                            Set rsAttend = Nothing
                        
                        End If
                    End If
                    'rsOTMaster.Close
                    'Set rsOTMaster = Nothing
                End If
            End If
        End If
        rsOTMaster.Close
        Set rsOTMaster = Nothing
        
        rsEmp.MoveNext
    Loop
    rsEmp.Close
    Set rsEmp = Nothing
    
End Sub

Private Function Get_Maximum_OT_Allowed(xEmpnbr, xAttDate)
    Dim rsOTMaster As New ADODB.Recordset
    Dim SQLQ As String
    
    Get_Maximum_OT_Allowed = 0
    
    'Get the Max OT allowed as per the entitlement period
    SQLQ = "SELECT OT_MBANK FROM HR_OVERTIME_BANK "
    SQLQ = SQLQ & " WHERE OT_EMPNBR = " & xEmpnbr
    SQLQ = SQLQ & " AND OT_EFDATE <= " & Date_SQL(xAttDate)
    SQLQ = SQLQ & " AND OT_ETDATE >= " & Date_SQL(xAttDate)
    rsOTMaster.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsOTMaster.EOF Then
        If Not IsNull(rsOTMaster("OT_MBANK")) Then
            Get_Maximum_OT_Allowed = rsOTMaster("OT_MBANK")
        End If
    End If
    rsOTMaster.Close
    Set rsOTMaster = Nothing

End Function
