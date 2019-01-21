VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRQuarter 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Custom Report"
   ClientHeight    =   8490
   ClientLeft      =   180
   ClientTop       =   825
   ClientWidth     =   10110
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8490
   ScaleWidth      =   10110
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtUpcoming 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   5640
      Width           =   6015
   End
   Begin VB.TextBox txtCurrent 
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   2520
      Width           =   6015
   End
   Begin VB.Frame fraLegend 
      Caption         =   "Incident Reporting"
      Height          =   2415
      Left            =   6360
      TabIndex        =   7
      Top             =   360
      Width           =   3255
      Begin VB.Label Label8 
         Caption         =   "Lost Time"
         Height          =   255
         Left            =   1680
         TabIndex        =   14
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Medical Aid"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "First Aid"
         Height          =   255
         Left            =   1680
         TabIndex        =   12
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Near Miss"
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Type"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "A B C"
         Height          =   735
         Left            =   1680
         TabIndex        =   9
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "Hazard Rating"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox txtYear 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.Frame frmDate 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   6075
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   1
         Left            =   3750
         TabIndex        =   3
         Tag             =   "40-Date upto and including this date forward"
         Top             =   0
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   0
         Left            =   1890
         TabIndex        =   2
         Tag             =   "40-Date from and including this date forward"
         Top             =   0
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin VB.Label lblFromTo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "From / To Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   8040
      Top             =   7320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowWidth     =   480
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   2
      WindowControls  =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label10 
      Caption         =   "Upcoming Events:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label Label9 
      Caption         =   "Current Associate Concerns:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Year:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblSelCri 
      Caption         =   "Selection Criteria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmRQuarter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*****************************************************
'    Form created by Bryan on 8/Nov/2005
'    for TS Tech custom Quarterly Report
'    Ticket#9720
'*****************************************************
Dim fglbEmpTable As String
Dim rsRPT As New ADODB.Recordset
Dim fglbFileName
Dim fglbDateTable
Dim fglbDateField

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm(frmRQuarter.Caption & " Criteria", Me) Then Exit Sub
    Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    x% = Cri_SetAll()
    Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
    Call set_PrintState(True)
End If
Exit Sub

PrntErr:
MsgBox "Error Printing - check your Windows Printer setup"

Resume Next

End Sub

Public Sub cmdView_Click()
Dim x%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    Screen.MousePointer = HOURGLASS
    Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    x% = Cri_SetAll()
    Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
    Call set_PrintState(True)
End If
Exit Sub

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", fglbEmpTable, "SELECT")
Resume Next

End Sub

Private Function Cri_SetAll()
Dim x%, strRName$


Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

Call CriCheck
Call Cri_Notes
SELATTWRK


Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "sn2369Quarter.rpt"


If Len(glbstrSelCri) > 0 Then
    glbstrSelCri = glbstrSelCri & "AND {HRMANWRK.WRKEMP}='" & glbUserID & "'"
Else
    glbstrSelCri = "{HRMANWRK.WRKEMP}='" & glbUserID & "'"
End If

If Len(glbstrSelCri) >= 0 Then
   Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
End If

Me.vbxCrystal.WindowTitle = "Quarterly Report"


Cri_SetAll = True



Screen.MousePointer = DEFAULT
Exit Function


modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "FR Comp Time", "Comp Report", "Select")
Cri_SetAll = False
Resume Next

End Function



Private Function CriCheck()
Dim x As Integer

CriCheck = False

If Len(txtYear) < 4 Then
    MsgBox "Invalid Year, please enter a 4 digit year"
    txtYear.SetFocus
    Exit Function
End If

For x = 0 To 1
 If Len(dlpDateRange(x).Text) > 0 Then
    If Not IsDate(dlpDateRange(x).Text) Then
        MsgBox "Not a valid date"
        dlpDateRange(x).Text = ""
        dlpDateRange(x).SetFocus
        Exit Function
    End If
 End If
Next x

If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    If DateDiff("d", dlpDateRange(1).Text, dlpDateRange(0).Text) > 0 Then
        MsgBox "Date Range To must be greater than date range From"
        dlpDateRange(1).Text = ""
        dlpDateRange(1).SetFocus
        Exit Function
    End If
End If


CriCheck = True
End Function





Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
glbOnTop = Me.name

Call setRptCaption(Me)

Call INI_Controls(Me)
Screen.MousePointer = DEFAULT


End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select function from the menu."
End Sub


Public Property Get ChangeAction() As UpdateStateEnum
ChangeAction = OPENING
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = Reports
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = False
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property
Public Property Get Updateble() As Boolean
Updateble = False
End Property
Public Property Get Deleteble() As Boolean
Deleteble = False
End Property

Public Property Get Printable() As Boolean


Printable = True


End Property

Public Sub SET_UP_MODE()
Call set_Buttons
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Cancel = (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub SELATTWRK()
    On Error GoTo EH
    
    Dim strSQL As String
    Dim rsDATA As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim DTE As String
    Dim DTE1 As String
    Dim strJobStatus As String
    Dim startMonth As Integer 'Month Sequence
    Dim c As Integer  'Month
    Dim x As Integer  'Sequence
    Dim TMP As Long
    Dim sFrom As Integer
    Dim sTo As Integer


    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(2).Caption = ""
    Screen.MousePointer = HOURGLASS
    
    For c% = 1 To 12
        x% = 0
        strSQL = "SELECT MONTH_SEQ FROM HRBUDGET WHERE BUDGET_MONTH=" & c%
        strSQL = strSQL & " AND BUDGET_YEAR=" & txtYear.Text
        rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
        If rsDATA.EOF = False And rsDATA.BOF = False Then
            x% = rsDATA("month_seq")
        End If
        rsDATA.Close
        If x = 1 Then startMonth = c
        If x% > 0 Then
            If c% < x% Then 'If month is less than the sequence then it must be next year
                DTE = getEOM(c%) & "/" & c% & "/" & (txtYear.Text + 1)
            Else
                DTE = getEOM(c%) & "/" & c% & "/" & txtYear.Text
            End If
            'Get Actualy Employee Count
            strSQL = "SELECT Count(HREMP.ED_EMPNBR) AS EMPCNT, HREMP.ED_DEPTNO, HREMP.ED_GLNO, HREMP.ED_PT, HREMP.ED_ADMINBY, HREMP.ED_DIV " ',HREMP.ED_LOC "
            If glbOracle Then
                strSQL = strSQL & "From HREMP, HR_JOB_HISTORY, HRJOB "
                strSQL = strSQL & "Where HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE AND (HREMP.ED_DOH <= " & Date_SQL(DTE) & ") AND (HRJOB.JB_STATUS<>'NA' AND HRJOB.JB_STATUS<>'L')  AND (HR_JOB_HISTORY.JH_CURRENT<>0)"
            Else
                strSQL = strSQL & "From  (HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
                strSQL = strSQL & "Where (HREMP.ED_DOH <= " & Date_SQL(DTE) & ") AND (HRJOB.JB_STATUS<>'NA' AND HRJOB.JB_STATUS<>'L') AND (HR_JOB_HISTORY.JH_CURRENT<>0)"
            End If
            strSQL = strSQL & "GROUP BY HREMP.ED_DEPTNO, HREMP.ED_GLNO, HREMP.ED_PT, HREMP.ED_ADMINBY, HREMP.ED_DIV " ' ,HREMP.ED_LOC "
            strSQL = strSQL & "HAVING (HREMP.ED_PT='FT' Or HREMP.ED_PT='TMP' Or HREMP.ED_PT='OT') "
            
            rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
            If rsDATA.EOF = False And rsDATA.BOF = False Then
                Do
'                    Select Case rsDATA("JB_STATUS")
'                    Case "L"
'                        strJobStatus = "S"
'                    Case Else
                        strJobStatus = "A"
'                    End Select
                        
                    If rsDATA("ED_PT") = "FT" Then
                        strSQL = "UPDATE HRBUDGET SET ACTUAL_FT_" & strJobStatus & " = " & rsDATA("EMPCNT")
                        strSQL = strSQL & " WHERE BD_DEPT='" & rsDATA("ED_DEPTNO") & "' AND GL_NUMBER='" & rsDATA("ED_GLNO") & "' AND BD_ADMINBY='" & rsDATA("ED_ADMINBY") & "' AND BD_DIV='" & rsDATA("ED_DIV") & "' " 'AND BD_LOCATION='" & rsDATA("ED_LOC") & "' "
                        strSQL = strSQL & "AND BUDGET_YEAR=" & txtYear.Text & " AND BUDGET_MONTH=" & c%
                        strSQL = strSQL & " AND BD_FREEZE=0"
                    ElseIf rsDATA("ED_PT") = "TMP" Then
                        strSQL = "UPDATE HRBUDGET SET ACTUAL_TMP_" & strJobStatus & " = " & rsDATA("EMPCNT")
                        strSQL = strSQL & " WHERE BD_DEPT='" & rsDATA("ED_DEPTNO") & "' AND GL_NUMBER='" & rsDATA("ED_GLNO") & "' AND BD_ADMINBY='" & rsDATA("ED_ADMINBY") & "' AND BD_DIV='" & rsDATA("ED_DIV") & "' " 'AND BD_LOCATION='" & rsDATA("ED_LOC") & "' "
                        strSQL = strSQL & "AND BUDGET_YEAR=" & txtYear.Text & " AND BUDGET_MONTH=" & c%
                        strSQL = strSQL & " AND BD_FREEZE=0"
                    ElseIf rsDATA("ED_PT") = "OT" Then
                        strSQL = "UPDATE HRBUDGET SET ACTUAL_OTHER_" & strJobStatus & " = " & rsDATA("EMPCNT")
                        strSQL = strSQL & " WHERE BD_DEPT='" & rsDATA("ED_DEPTNO") & "' AND GL_NUMBER='" & rsDATA("ED_GLNO") & "' AND BD_ADMINBY='" & rsDATA("ED_ADMINBY") & "' AND BD_DIV='" & rsDATA("ED_DIV") & "' " 'AND BD_LOCATION='" & rsDATA("ED_LOC") & "' "
                        strSQL = strSQL & "AND BUDGET_YEAR=" & txtYear.Text & " AND BUDGET_MONTH=" & c%
                        strSQL = strSQL & " AND BD_FREEZE=0"
                    End If
                    gdbAdoIhr001.BeginTrans
                    gdbAdoIhr001.Execute strSQL
                    gdbAdoIhr001.CommitTrans
    
                    rsDATA.MoveNext
                Loop Until rsDATA.EOF
            End If
            rsDATA.Close
'*********************
 'Find Actual Supervisors for this month(c)
            strSQL = "SELECT Count(HREMP.ED_EMPNBR) AS EMPCNT, HREMP.ED_DEPTNO, HREMP.ED_GLNO, HREMP.ED_PT, HREMP.ED_ADMINBY, HRJOB.JB_STATUS, HREMP.ED_DIV " ',HREMP.ED_LOC "
            If glbOracle Then
                strSQL = strSQL & "From HREMP, HR_JOB_HISTORY, HRJOB "
                strSQL = strSQL & "Where HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE AND (HREMP.ED_DOH <= " & Date_SQL(DTE) & ")  AND (HR_JOB_HISTORY.JH_CURRENT<>0)"
            Else
                strSQL = strSQL & "From  (HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
                strSQL = strSQL & "Where (HREMP.ED_DOH <= " & Date_SQL(DTE) & ") AND (HR_JOB_HISTORY.JH_CURRENT<>0) "
            End If
            strSQL = strSQL & "GROUP BY HREMP.ED_DEPTNO, HREMP.ED_GLNO, HREMP.ED_PT, HREMP.ED_ADMINBY, HRJOB.JB_STATUS, HREMP.ED_DIV " ',HREMP.ED_LOC "
            strSQL = strSQL & "HAVING (HREMP.ED_PT='FT' Or HREMP.ED_PT='TMP' Or HREMP.ED_PT='OT') AND (HRJOB.JB_STATUS='L')"
            
            rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
            If rsDATA.EOF = False And rsDATA.BOF = False Then
                Do
'                    Select Case rsDATA("JB_STATUS")
'                    Case "L"
                        strJobStatus = "S"
'                    Case Else
'                        strJobStatus = "A"
'                    End Select
                        
                    If rsDATA("ED_PT") = "FT" Then
                        strSQL = "UPDATE HRBUDGET SET ACTUAL_FT_" & strJobStatus & " = " & rsDATA("EMPCNT")
                        strSQL = strSQL & " WHERE BD_DEPT='" & rsDATA("ED_DEPTNO") & "' AND GL_NUMBER='" & rsDATA("ED_GLNO") & "' AND BD_ADMINBY='" & rsDATA("ED_ADMINBY") & "' AND BD_DIV='" & rsDATA("ED_DIV") & "' " 'AND BD_LOCATION='" & rsDATA("ED_LOC") & "' "
                        strSQL = strSQL & "AND BUDGET_YEAR=" & txtYear.Text & " AND BUDGET_MONTH=" & c%
                        strSQL = strSQL & " AND BD_FREEZE=0"
                    ElseIf rsDATA("ED_PT") = "TMP" Then
                        strSQL = "UPDATE HRBUDGET SET ACTUAL_TMP_" & strJobStatus & " = " & rsDATA("EMPCNT")
                        strSQL = strSQL & " WHERE BD_DEPT='" & rsDATA("ED_DEPTNO") & "' AND GL_NUMBER='" & rsDATA("ED_GLNO") & "' AND BD_ADMINBY='" & rsDATA("ED_ADMINBY") & "' AND BD_DIV='" & rsDATA("ED_DIV") & "' " 'AND BD_LOCATION='" & rsDATA("ED_LOC") & "' "
                        strSQL = strSQL & "AND BUDGET_YEAR=" & txtYear.Text & " AND BUDGET_MONTH=" & c%
                        strSQL = strSQL & " AND BD_FREEZE=0"
                    ElseIf rsDATA("ED_PT") = "OT" Then
                        strSQL = "UPDATE HRBUDGET SET ACTUAL_OTHER_" & strJobStatus & " = " & rsDATA("EMPCNT")
                        strSQL = strSQL & " WHERE BD_DEPT='" & rsDATA("ED_DEPTNO") & "' AND GL_NUMBER='" & rsDATA("ED_GLNO") & "' AND BD_ADMINBY='" & rsDATA("ED_ADMINBY") & "' AND BD_DIV='" & rsDATA("ED_DIV") & "' " 'AND BD_LOCATION='" & rsDATA("ED_LOC") & "' "
                        strSQL = strSQL & "AND BUDGET_YEAR=" & txtYear.Text & " AND BUDGET_MONTH=" & c%
                        strSQL = strSQL & " AND BD_FREEZE=0"
                    End If
                    gdbAdoIhr001.BeginTrans
                    gdbAdoIhr001.Execute strSQL
                    gdbAdoIhr001.CommitTrans
    
                    rsDATA.MoveNext
                Loop Until rsDATA.EOF
            End If
            rsDATA.Close



'**********************
            If c% < x% Then
                DTE = getEOM(c%) & "/" & c% & "/" & (txtYear.Text + 1)
            Else
                DTE = getEOM(c%) & "/" & c% & "/" & txtYear
            End If
            
            strSQL = "DELETE FROM HRMANTERM WHERE WRKEMP='" & glbUserID & "'"
            gdbAdoIhr001W.BeginTrans
            gdbAdoIhr001W.Execute strSQL
            gdbAdoIhr001W.CommitTrans
            'Find Actual Employees from the terminated for this month(c)
            strSQL = "SELECT TERM_HREMP.ED_EMPNBR, TERM_HREMP.ED_DEPTNO, TERM_HREMP.ED_GLNO, TERM_HREMP.ED_PT, TERM_HREMP.ED_ADMINBY, Term_JOB_HISTORY.JH_JOB, TERM_HREMP.ED_LOC, TERM_HREMP.ED_DIV  "
            If glbOracle Then
                strSQL = strSQL & "From  TERM_HREMP, TERM_HRTRMEMP, Term_JOB_HISTORY "
                strSQL = strSQL & "Where TERM_HREMP.ED_EMPNBR = TERM_HRTRMEMP.Employee_Number AND TERM_HREMP.ED_EMPNBR = Term_JOB_HISTORY.JH_EMPNBR AND (TERM_HREMP.ED_DOH <= " & Date_SQL(DTE) & ") AND (TERM_HRTRMEMP.TERM_DOT >= " & Date_SQL(DTE) & ") "
            Else
                strSQL = strSQL & "From  ((TERM_HREMP INNER JOIN TERM_HRTRMEMP ON TERM_HREMP.ED_EMPNBR = TERM_HRTRMEMP.Employee_Number) INNER JOIN Term_JOB_HISTORY ON TERM_HREMP.ED_EMPNBR = Term_JOB_HISTORY.JH_EMPNBR)"
                strSQL = strSQL & "Where (TERM_HREMP.ED_DOH <= " & Date_SQL(DTE) & ") AND (TERM_HRTRMEMP.TERM_DOT >= " & Date_SQL(DTE) & ") "
            End If
            strSQL = strSQL & "AND (((TERM_HREMP.ED_PT)='FT' Or (TERM_HREMP.ED_PT)='TMP' Or (TERM_HREMP.ED_PT)='OT')) AND Term_JOB_HISTORY.JH_CURRENT<>0"
            rsDATA.Open strSQL, gdbAdoIhr001X, adOpenStatic, adLockOptimistic, adCmdText
            If rsDATA.EOF = False And rsDATA.BOF = False Then
                Do
                    strSQL = "INSERT INTO HRMANTERM (ED_EMPNBR, ED_DEPTNO, ED_GLNO, ED_PT, ED_ADMINBY, JH_JOB, ED_LOC, ED_DIV, WRKEMP) "
                    strSQL = strSQL & "VALUES (" & rsDATA("ED_EMPNBR") & ", '" & rsDATA("ED_DEPTNO") & "', '" & rsDATA("ED_GLNO") & "', '" & rsDATA("ED_PT") & "', '" & rsDATA("ED_ADMINBY") & "', '" & rsDATA("JH_JOB") & "', '" & rsDATA("ED_LOC") & "', '" & rsDATA("ED_DIV") & "', '" & glbUserID & "')"
                    gdbAdoIhr001W.BeginTrans
                    gdbAdoIhr001W.Execute strSQL
                    gdbAdoIhr001W.CommitTrans
                    rsDATA.MoveNext
                Loop Until rsDATA.EOF
            End If
            rsDATA.Close
            'Join the Terminated Eployees to Job Status and count.
            strSQL = "SELECT Count(HRMANTERM.ED_EMPNBR) AS EMPCNT, HRMANTERM.ED_DEPTNO, HRMANTERM.ED_GLNO, HRMANTERM.ED_PT, HRMANTERM.ED_ADMINBY, HRJOB.JB_STATUS, HRMANTERM.ED_DIV " ',HRMANTERM.ED_LOC  "
            If glbOracle Then
                strSQL = strSQL & "From  HRMANTERM, hrjob "
                strSQL = strSQL & "WHERE HRMANTERM.JH_JOB = hrjob.JB_CODE "
            Else
                strSQL = strSQL & "From  HRMANTERM INNER JOIN hrjob ON HRMANTERM.JH_JOB = hrjob.JB_CODE "
            End If
            strSQL = strSQL & "GROUP BY HRMANTERM.ED_DEPTNO, HRMANTERM.ED_GLNO, HRMANTERM.ED_PT, HRMANTERM.ED_ADMINBY, HRJOB.JB_STATUS, HRMANTERM.ED_DIV " ', HRMANTERM.ED_LOC "
            strSQL = strSQL & "HAVING (hrjob.JB_STATUS<>'NA')"
            rsDATA.Open strSQL, gdbAdoIhr001W, adOpenStatic, adLockOptimistic, adCmdText
            If rsDATA.EOF = False And rsDATA.BOF = False Then
                Do
                    Select Case rsDATA("JB_STATUS")
                    Case "L"
                        strJobStatus = "S"
                    Case Else
                        strJobStatus = "A"
                    End Select
                    strSQL = "SELECT ACTUAL_FT_" & strJobStatus & ", ACTUAL_TMP_" & strJobStatus & ", ACTUAL_OTHER_" & strJobStatus & " FROM HRBUDGET "
                    strSQL = strSQL & " WHERE BD_DEPT='" & rsDATA("ED_DEPTNO") & "' AND GL_NUMBER='" & rsDATA("ED_GLNO") & "' AND BD_ADMINBY='" & rsDATA("ED_ADMINBY") & "' AND BD_DIV='" & rsDATA("ED_DIV") & "' " 'AND BD_LOCATION='" & rsDATA("ED_LOC") & "' "
                    strSQL = strSQL & "AND BUDGET_YEAR=" & txtYear.Text & " AND BUDGET_MONTH=" & c%
                    strSQL = strSQL & " AND BD_FREEZE=0"
                    rsTemp.Open strSQL, gdbAdoIhr001, adOpenDynamic, adLockOptimistic, adCmdText
                    If rsTemp.EOF = False And rsTemp.BOF = False Then
                        If rsDATA("ED_PT") = "FT" Then
                            rsTemp("ACTUAL_FT_" & strJobStatus) = rsTemp("ACTUAL_FT_" & strJobStatus) + rsDATA("EMPCNT")
    '                        strSQL = "UPDATE HRBUDGET SET ACTUAL_FT_" & strJobStatus & " = ACTUAL_FT_" & strJobStatus & " + " & rsDATA("EMPCNT")
    '                        strSQL = strSQL & " WHERE BD_DEPT='" & rsDATA("ED_DEPTNO") & "' AND GL_NUMBER='" & rsDATA("ED_GLNO") & "' AND BD_ADMINBY='" & rsDATA("ED_ADMINBY") & "' AND BD_LOCATION='" & rsDATA("ED_LOC") & "' "
    '                        strSQL = strSQL & "AND BUDGET_YEAR=" & txtYear.Text & " AND BUDGET_MONTH=" & c%
                        ElseIf rsDATA("ED_PT") = "TMP" Then
                            rsTemp("ACTUAL_TMP_" & strJobStatus) = rsTemp("ACTUAL_TMP_" & strJobStatus) + rsDATA("EMPCNT")
    '                        strSQL = "UPDATE HRBUDGET SET ACTUAL_TMP_" & strJobStatus & " = ACTUAL_TMP_" & strJobStatus & " + " & rsDATA("EMPCNT")
    '                        strSQL = strSQL & " WHERE BD_DEPT='" & rsDATA("ED_DEPTNO") & "' AND GL_NUMBER='" & rsDATA("ED_GLNO") & "' AND BD_ADMINBY='" & rsDATA("ED_ADMINBY") & "' AND BD_LOCATION='" & rsDATA("ED_LOC") & "' "
    '                        strSQL = strSQL & "AND BUDGET_YEAR=" & txtYear.Text & " AND BUDGET_MONTH=" & c%
                        ElseIf rsDATA("ED_PT") = "OT" Then
                            rsTemp("ACTUAL_OT_" & strJobStatus) = rsTemp("ACTUAL_OT_" & strJobStatus) + rsDATA("EMPCNT")
    '                        strSQL = "UPDATE HRBUDGET SET ACTUAL_OTHER_" & strJobStatus & " = ACTUAL_OTHER_" & strJobStatus & " + " & rsDATA("EMPCNT")
    '                        strSQL = strSQL & " WHERE BD_DEPT='" & rsDATA("ED_DEPTNO") & "' AND GL_NUMBER='" & rsDATA("ED_GLNO") & "' AND BD_ADMINBY='" & rsDATA("ED_ADMINBY") & "' AND BD_LOCATION='" & rsDATA("ED_LOC") & "' "
    '                        strSQL = strSQL & "AND BUDGET_YEAR=" & txtYear.Text & " AND BUDGET_MONTH=" & c%
                        End If
                        rsTemp.Update
                    End If
                    rsTemp.Close
'                    gdbAdoIhr001.BeginTrans
'                    gdbAdoIhr001.Execute strSQL
'                    gdbAdoIhr001.CommitTrans
                    rsDATA.MoveNext
                Loop Until rsDATA.EOF
            End If
            rsDATA.Close
        End If
    Next c
    
    MDIMain.panHelp(0).FloodPercent = 15
    
    'Get Manpower
    strSQL = "DELETE FROM HRMANWRK WHERE WRKEMP='" & glbUserID & "'"
    gdbAdoIhr001W.BeginTrans
    gdbAdoIhr001W.Execute strSQL
    gdbAdoIhr001W.CommitTrans

    MDIMain.panHelp(0).FloodPercent = 20
    'Gather up the data and put into HRMANWRK for general sections
    strSQL = "INSERT INTO HRMANWRK ( BD_ROW,GL_NUMBER, BD_ADMINBY, BD_ADMINBY_TABL, WRKEMP) "
    strSQL = strSQL & "(SELECT 'FT' as Category, 'Manpower' as manlabel, TB_KEY, TB_NAME, '" & glbUserID & "' FROM HRTABL "
    strSQL = strSQL & "WHERE TB_NAME='EDAB')"
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute strSQL
    gdbAdoIhr001.CommitTrans
    strSQL = "INSERT INTO HRMANWRK ( BD_ROW, GL_NUMBER, BD_ADMINBY, BD_ADMINBY_TABL, WRKEMP) "
    strSQL = strSQL & "(SELECT 'TMP' as Category, 'Manpower' as manlabel, TB_KEY, TB_NAME, '" & glbUserID & "' FROM HRTABL "
    strSQL = strSQL & "WHERE TB_NAME='EDAB')"
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute strSQL
    gdbAdoIhr001.CommitTrans
    strSQL = "INSERT INTO HRMANWRK (BD_ROW, GL_NUMBER, BD_ADMINBY, BD_ADMINBY_TABL, WRKEMP) "
    strSQL = strSQL & "(SELECT 'FT' as Category, 'Overtime' as manlabel, TB_KEY, TB_NAME, '" & glbUserID & "' FROM HRTABL "
    strSQL = strSQL & "WHERE TB_NAME='EDAB')"
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute strSQL
    gdbAdoIhr001.CommitTrans
    strSQL = "INSERT INTO HRMANWRK (BD_ROW, GL_NUMBER, BD_ADMINBY, BD_ADMINBY_TABL, WRKEMP) "
    strSQL = strSQL & "(SELECT 'TMP' as Category, 'Overtime' as manlabel, TB_KEY, TB_NAME, '" & glbUserID & "' FROM HRTABL "
    strSQL = strSQL & "WHERE TB_NAME='EDAB')"
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute strSQL
    gdbAdoIhr001.CommitTrans
    strSQL = "INSERT INTO HRMANWRK (BD_ROW, GL_NUMBER, BD_ADMINBY, BD_ADMINBY_TABL, WRKEMP) "
    strSQL = strSQL & "(SELECT 'FT' as Category, 'Safety Recordables' as manlabel, TB_KEY, TB_NAME, '" & glbUserID & "' FROM HRTABL "
    strSQL = strSQL & "WHERE TB_NAME='EDAB')"
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute strSQL
    gdbAdoIhr001.CommitTrans
    strSQL = "INSERT INTO HRMANWRK (BD_ROW, GL_NUMBER, BD_ADMINBY, BD_ADMINBY_TABL, WRKEMP) "
    strSQL = strSQL & "(SELECT 'TMP' as Category, 'Safety Recordables' as manlabel, TB_KEY, TB_NAME, '" & glbUserID & "' FROM HRTABL "
    strSQL = strSQL & "WHERE TB_NAME='EDAB')"
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute strSQL
    gdbAdoIhr001.CommitTrans
    
    If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
        c = month(dlpDateRange(0).Text)
        If c >= 4 Then
            sFrom = c - 3
        Else
            sFrom = c + 9
        End If
        If month(dlpDateRange(1).Text) >= 4 Then
            sTo = month(dlpDateRange(1).Text) - 3
        Else
            sTo = month(dlpDateRange(1).Text) + 9
        End If
    Else
        c = 4
        sFrom = 1
        sTo = 12
    End If
    
    For x = sFrom To sTo
        
        'Manpower
        strSQL = "SELECT sum(ACTUAL_FT_A) as AFT, sum(ACTUAL_FT_S) as SFT, sum(ACTUAL_TMP_A) as ATMP, sum(ACTUAL_TMP_S) as STMP, BD_ADMINBY, BD_ADMINBY_TABL, Budget_Month "
        strSQL = strSQL & "FROM HRBUDGET WHERE BUDGET_YEAR=" & txtYear.Text & " AND MONTH_SEQ=" & x
        strSQL = strSQL & " GROUP BY BD_ADMINBY, BD_ADMINBY_TABL, Budget_Month "
        rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
        If rsDATA.EOF = False And rsDATA.BOF = False Then
            Do
                strSQL = "UPDATE HRMANWRK SET seq" & x & "='" & CStr(nz(rsDATA("AFT"), 0) + nz(rsDATA("SFT"), 0)) & "' "
                strSQL = strSQL & "WHERE BD_Row='FT' AND GL_NUMBER='Manpower' AND BD_ADMINBY='" & rsDATA("BD_ADMINBY") & "' AND WrkEmp='" & glbUserID & "'"
                gdbAdoIhr001.BeginTrans
                gdbAdoIhr001.Execute strSQL
                gdbAdoIhr001.CommitTrans
                strSQL = "UPDATE HRMANWRK SET seq" & x & "='" & CStr(nz(rsDATA("ATMP"), 0) + nz(rsDATA("STMP"), 0)) & "' "
                strSQL = strSQL & "WHERE BD_Row='TMP' AND GL_NUMBER='Manpower' AND BD_ADMINBY='" & rsDATA("BD_ADMINBY") & "' AND WrkEmp='" & glbUserID & "'"
                gdbAdoIhr001.BeginTrans
                gdbAdoIhr001.Execute strSQL
                gdbAdoIhr001.CommitTrans
                rsDATA.MoveNext
            Loop Until rsDATA.EOF
        End If
        rsDATA.Close
        
        If c% < x% Then 'If month is less than the sequence then it must be next year
            DTE = getEOM(c%) & "/" & c% & "/" & (txtYear.Text + 1)
            DTE1 = "01/" & c% & "/" & (txtYear.Text + 1)
        Else
            DTE = getEOM(c%) & "/" & c% & "/" & txtYear.Text
            DTE1 = "01/" & c% & "/" & (txtYear.Text)
        End If
        'Overtime
        strSQL = "SELECT SUM(HR_ATTENDANCE.AD_HRS) AS OTHrs, HREMP.ED_ADMINBY_TABL, HREMP.ED_ADMINBY, ED_PT "
        strSQL = strSQL & "FROM HR_ATTENDANCE INNER JOIN HREMP ON HR_ATTENDANCE.AD_EMPNBR = HREMP.ED_EMPNBR "
        strSQL = strSQL & "WHERE (HR_ATTENDANCE.AD_REASON = 'OT') AND (HR_ATTENDANCE.AD_DOA BETWEEN " & Date_SQL(DTE1)
        strSQL = strSQL & " AND " & Date_SQL(DTE) & ") AND (HREMP.ED_PT='FT' OR HREMP.ED_PT='TMP') "
        strSQL = strSQL & "GROUP BY HREMP.ED_ADMINBY_TABL, HREMP.ED_ADMINBY, ED_PT"
        rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
        If rsDATA.EOF = False And rsDATA.BOF = False Then
            Do
                Select Case rsDATA("ED_PT")
                Case "FT"
                    strSQL = "UPDATE HRMANWRK SET seq" & x & "='" & CStr(nz(rsDATA("OTHrs"), 0)) & "' "
                    strSQL = strSQL & "WHERE BD_Row='FT' AND GL_NUMBER='Overtime' AND BD_ADMINBY='" & rsDATA("ED_ADMINBY") & "' AND WrkEmp='" & glbUserID & "'"
                Case "TMP"
                    strSQL = "UPDATE HRMANWRK SET seq" & x & "='" & CStr(nz(rsDATA("OTHrs"), 0)) & "' "
                    strSQL = strSQL & "WHERE BD_Row='TMP' AND GL_NUMBER='Overtime' AND BD_ADMINBY='" & rsDATA("ED_ADMINBY") & "' AND WrkEmp='" & glbUserID & "'"
                End Select
                gdbAdoIhr001.BeginTrans
                gdbAdoIhr001.Execute strSQL
                gdbAdoIhr001.CommitTrans
                rsDATA.MoveNext
            Loop Until rsDATA.EOF
        End If
        rsDATA.Close
        
        'Safety Recordables
        strSQL = "SELECT Count(HR_OCC_HEALTH_SAFETY.EC_EMPNBR) as EmpCnt, HREMP.ED_PT, HREMP.ED_ADMINBY_TABL, HREMP.ED_ADMINBY "
        strSQL = strSQL & "FROM HR_OCC_HEALTH_SAFETY INNER JOIN HREMP ON HR_OCC_HEALTH_SAFETY.EC_EMPNBR = HREMP.ED_EMPNBR "
        strSQL = strSQL & "WHERE (HR_OCC_HEALTH_SAFETY.EC_HAZARD = 'A' OR HR_OCC_HEALTH_SAFETY.EC_HAZARD = 'B' OR HR_OCC_HEALTH_SAFETY.EC_HAZARD = 'C') AND (HR_OCC_HEALTH_SAFETY.EC_TYPE = 'MA' OR "
        strSQL = strSQL & "HR_OCC_HEALTH_SAFETY.EC_TYPE = 'LT' OR HR_OCC_HEALTH_SAFETY.EC_TYPE = 'FA' OR HR_OCC_HEALTH_SAFETY.EC_TYPE = 'NM') AND (HR_OCC_HEALTH_SAFETY.EC_OCCDATE BETWEEN " & Date_SQL(DTE1)
        strSQL = strSQL & " AND " & Date_SQL(DTE) & ")"
        strSQL = strSQL & "GROUP BY HREMP.ED_PT, HREMP.ED_ADMINBY_TABL, HREMP.ED_ADMINBY"
        rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
        If rsDATA.EOF = False And rsDATA.BOF = False Then
            Do
                Select Case rsDATA("ED_PT")
                Case "FT"
                    strSQL = "UPDATE HRMANWRK SET seq" & x & "='" & CStr(nz(rsDATA("empcnt"), 0)) & "' "
                    strSQL = strSQL & "WHERE BD_Row='FT' AND GL_NUMBER='Safety Recordables' AND BD_ADMINBY='" & rsDATA("ED_ADMINBY") & "' AND WrkEmp='" & glbUserID & "'"
                Case "TMP"
                    strSQL = "UPDATE HRMANWRK SET seq" & x & "='" & CStr(nz(rsDATA("empcnt"), 0)) & "' "
                    strSQL = strSQL & "WHERE BD_Row='TMP' AND GL_NUMBER='Safety Recordables' AND BD_ADMINBY='" & rsDATA("ED_ADMINBY") & "' AND WrkEmp='" & glbUserID & "'"
                End Select
                gdbAdoIhr001.BeginTrans
                gdbAdoIhr001.Execute strSQL
                gdbAdoIhr001.CommitTrans
                rsDATA.MoveNext
            Loop Until rsDATA.EOF
        End If
        rsDATA.Close
         MDIMain.panHelp(0).FloodPercent = MDIMain.panHelp(0).FloodPercent + 6
         c = c + 1
         If c > 12 Then c = 1
         
    Next x
     MDIMain.panHelp(0).FloodPercent = 100
        
ExH:
    Set rsDATA = Nothing
    Screen.MousePointer = DEFAULT
    Exit Sub
EH:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "selAttWrk", "Manpower Report", "Select")
    Resume ExH
  
        
End Sub

Public Function getEOM(Mnt As Integer) As Integer
   Dim myDate As Date
   Dim NextMonth As Date, EndOfMonth As Date
   myDate = Format("1/" & Mnt & "/2005", "dd/mm/yyyy")
   NextMonth = DateAdd("m", 1, myDate)
   EndOfMonth = NextMonth - DatePart("d", NextMonth)
   getEOM = Day(EndOfMonth)

End Function

Private Function nz(var As Variant, rep As Variant) As Variant
    If IsNull(var) Then
        nz = rep
    Else
        nz = var
    End If
End Function

Private Sub Cri_Notes()
On Error GoTo EH

Dim rs As New ADODB.Recordset
Dim strSQL As String

strSQL = "DELETE FROM HRQNotesWrk WHERE WrkEmp='" & glbUserID & "'"
gdbAdoIhr001W.BeginTrans
gdbAdoIhr001W.Execute strSQL
gdbAdoIhr001W.CommitTrans

rs.Open "HRQNotesWrk", gdbAdoIhr001W, adOpenDynamic, adLockOptimistic, adCmdTable
rs.AddNew
    rs("NT_Current") = Left(txtCurrent.Text, 254)
    rs("NT_Upcoming") = txtUpcoming.Text
    rs("WrkEmp") = glbUserID
rs.Update
    
rs.Close

ExH:
    Set rs = Nothing
    Screen.MousePointer = DEFAULT
    Exit Sub
EH:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "selAttWrk", "Manpower Report", "Select")
    Resume ExH

End Sub

