VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRManpower 
   Caption         =   "Manpower Plan"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   9975
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtBudgetYear 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   1920
      Width           =   975
   End
   Begin Threed.SSCheck chkTemporary 
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Tag             =   "Include Temporary"
      Top             =   1320
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Temporary"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Value           =   -1  'True
   End
   Begin Threed.SSCheck chkFullTime 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Tag             =   "Include Full Time"
      Top             =   1320
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Full Time"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Value           =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpGL 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Tag             =   "Specific GL# Desired"
      Top             =   720
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   0
      LookupType      =   3
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Tag             =   "Specific Division Desired"
      Top             =   240
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   15
      LookupType      =   2
      MultiSelect     =   -1  'True
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   6480
      Top             =   2760
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
   Begin VB.Label lblDateRange 
      Caption         =   "Budget Year"
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
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblGL 
      Caption         =   "G/L"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblDept 
      Caption         =   "Department"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmRManpower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*******************************************************
'*                                                     *
'*      Form: frmRManpower                             *
'*                                                     *
'*           Created: 14/Jul/05    By: Bryan           *
'*           Modified:             By:                 *
'*                                                     *
'*           Comments: To report budgeted Manpower data*
'*                                                     *
'*******************************************************

Private Sub chkFullTime_GotFocus()
     Call SetPanHelp(Me.ActiveControl)
End Sub


Private Sub chkTemporary_GotFocus()
     Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub clpDept_GotFocus()
     Call SetPanHelp(Me.ActiveControl)
End Sub


Private Sub clpGL_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub dlpDateRange_GotFocus(Index As Integer)
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub Form_Activate()
    Call SET_UP_MODE
End Sub

Private Sub Form_Load()
    MDIMain.lstPanel.Visible = False
    MDIMain.lstView.Visible = False
    glbOnTop = "FRMRMANPOWER"
    Screen.MousePointer = HOURGLASS
    
    Call setRptCaption(Me)
    Call setCaption(lblGL)
    Call INI_Controls(Me)

    Screen.MousePointer = DEFAULT
    If glbCompSerial = "S/W - 2369W" Then
        Me.chkTemporary.Caption = "Temporary"
    Else
        'Ticket #21100 - should not be 'Other' but be 'Full Time' only. Instead the Temporary checkbox
        'should be 'Other'.
        'Me.chkFullTime.Caption = "Other"
        Me.chkTemporary.Caption = "Other"
    End If
    
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

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim X%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm("Manpower Plan", Me) Then Exit Sub
    Call set_PrintState(False)
    X% = Cri_SetAll()
    Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    Call set_PrintState(True)
End If
Exit Sub

PrntErr:
MsgBox "CRW ERROR : " & Chr(10) & "[" & str(Err) & "] : " & Me.vbxCrystal.LastErrorString, , "info:HR"
Resume Next

End Sub



Public Sub cmdView_Click()
Dim X%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    Call set_PrintState(False)
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    Screen.MousePointer = HOURGLASS
    X% = Cri_SetAll()
    Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    Call set_PrintState(True)
End If
Exit Sub

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
MsgBox "CRW ERROR : " & Chr(10) & "[" & str(Err) & "] : " & Me.vbxCrystal.LastErrorString, , "info:HR"

Resume Next

End Sub

Private Function CriCheck()
Dim X%

    CriCheck = False
    
    If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
        MsgBox lStr("If Department Entered - it must be known"), , "info:HR"
         clpDept.SetFocus
        Exit Function
    End If
    
    If Len(clpGL.Text) > 0 And clpGL.Caption = "Unassigned" Then
        MsgBox lStr("If G/L Entered - it must be known"), , "info:HR"
        clpGL.SetFocus
        Exit Function
    End If

    If Len(txtBudgetYear.Text) = 0 Then
        MsgBox "Budget Year is required", , "info:HR"
        txtBudgetYear.SetFocus
        Exit Function
    End If
    
    CriCheck = True
End Function

Private Function Cri_SetAll()
Dim X%, strRName$, xNoFiles, dscGroup$
Dim dtYYY#, dtMM#, dtDD#


Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

' call cri models set both glbiONeWhere and strSelCri
Call Cri_GL
Call Cri_Dates

'Ticket #24306 - Open it up - they are using this
'Ticket #21876
'If glbCompSerial <> "S/N - 2369W" Then
    Call SELATTWRK
'End If

Call Cri_Sorts

MDIMain.panHelp(0).FloodPercent = 100

If Len(glbstrSelCri) > 0 Then
    glbstrSelCri = glbstrSelCri & "AND {HRMANWRK.WRKEMP}='" & glbUserID & "'"
Else
    glbstrSelCri = "{HRMANWRK.WRKEMP}='" & glbUserID & "'"
End If
Me.vbxCrystal.SelectionFormula = glbstrSelCri

If glbCompSerial = "S/N - 2369W" Then
    strRName$ = glbIHRREPORTS & "rzmanpow.rpt"
Else
    strRName$ = glbIHRREPORTS & "rzmanplan.rpt"
End If
Me.vbxCrystal.ReportFileName = strRName$

If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    Me.vbxCrystal.DataFiles(0) = glbIHRDB
    Me.vbxCrystal.DataFiles(1) = glbIHRDB
    Me.vbxCrystal.DataFiles(2) = glbIHRDBW
    Me.vbxCrystal.DataFiles(3) = glbIHRDB
End If

Me.vbxCrystal.SubreportToChange = "rzmanpow1.rpt"
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    Me.vbxCrystal.DataFiles(0) = glbIHRDBW
    Me.vbxCrystal.DataFiles(1) = glbIHRDB
End If
Me.vbxCrystal.SubreportToChange = ""

Me.vbxCrystal.SubreportToChange = "rzmanpow2.rpt"
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    Me.vbxCrystal.DataFiles(0) = glbIHRDBW
End If

Me.vbxCrystal.SubreportToChange = ""

' window title if appropriate
Me.vbxCrystal.WindowTitle = "Manpower Plan"

Cri_SetAll = True

Exit Function

modSetCriteria_Err:
    Screen.MousePointer = DEFAULT
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    MDIMain.panHelp(0).FloodPercent = 0
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Cri_SetAll", "Manpower Plan", "Select")
    Cri_SetAll = False
    Resume Next
End Function

Private Sub Cri_Div()
    Dim DivCri As String
    
    If Len(clpDept.Text) > 0 Then
        DivCri = "({HRBUDGET.BD_DEPT} = '" & clpDept.Text & "')"
    End If
    
    If Len(DivCri) >= 1 Then
        If Not glbiOneWhere Then
            glbstrSelCri = DivCri
        Else
            glbstrSelCri = glbstrSelCri & " AND " & DivCri
        End If
        glbiOneWhere = True
    End If
End Sub

Private Sub Cri_GL()
    Dim GLCri As String
    
    If Len(clpGL.Text) > 0 Then
        GLCri = "{HRGL.GL_NO} IN ['" & getCodes(clpGL.Text) & "'] "
    End If
    
    If Len(GLCri) >= 1 Then
        If glbiOneWhere Then
            glbstrSelCri = glbstrSelCri & " AND " & GLCri
        Else
            glbstrSelCri = GLCri
        End If
        glbiOneWhere = True
    End If
End Sub

Private Sub Cri_Dates()
 Dim TempCri As String
Dim dtYYY#, dtMM#, dtDD#
Dim X%
Dim EECri As String, LocCri As String

If Len(txtBudgetYear.Text) = 0 Then Exit Sub

TempCri = "{HRMANWRK.BD_YEAR} = " & txtBudgetYear.Text


Cri_FTDatst:
If Len(TempCri) > 0 Then
    If Len(glbstrSelCri) > 0 Then
      glbstrSelCri = glbstrSelCri & " AND " & TempCri
    Else
      glbstrSelCri = TempCri
    End If
    glbiOneWhere = True
End If
End Sub

Private Sub SELATTWRK()
    On Error GoTo Eh
    Dim strSQL As String
    Dim rsDATA As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim DTE As String
    Dim strDiv As String
    Dim strGL As String
    Dim strJobStatus As String
    Dim C% 'Month
    Dim X% 'Sequence
    Dim FTA, FTS, TMA, TMS

    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(2).Caption = ""
    Screen.MousePointer = HOURGLASS
    
   
    If Len(clpDept.Text) > 0 Then
        strDiv = " WHERE BD_DEPT='" & clpDept.Text & "'"
    Else
        strDiv = ""
    End If
    
    strSQL = "UPDATE HRBUDGET SET ACTUAL_FT_A = NULL, ACTUAL_TMP_A=NULL, ACTUAL_OTHER_A=NULL, ACTUAL_FT_S = NULL, ACTUAL_TMP_S=NULL, ACTUAL_OTHER_S=NULL "
    strSQL = strSQL & " WHERE BD_FREEZE=0 and BUDGET_YEAR=" & txtBudgetYear.Text
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute strSQL
    gdbAdoIhr001.CommitTrans
    
    For C% = 1 To 12
        X% = 0
        strSQL = "SELECT MONTH_SEQ FROM HRBUDGET WHERE BUDGET_MONTH=" & C%
        strSQL = strSQL & " AND BUDGET_YEAR=" & txtBudgetYear.Text
        rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic
        If rsDATA.EOF = False And rsDATA.BOF = False Then
            X% = rsDATA("month_seq")
        End If
        rsDATA.Close
            
        If X% > 0 Then
            If C% < X% Then 'If month is less than the sequence then it must be next year
                DTE = getEOM(C%) & "/" & MonthName(C%, True) & "/" & (txtBudgetYear.Text + 1)
            Else
                DTE = getEOM(C%) & "/" & MonthName(C%, True) & "/" & txtBudgetYear.Text
            End If
            
            If CDate(DTE) < Date Then
                'Find Actual Associates for this month(c)
                
                strSQL = "SELECT * FROM HRBUDGET WHERE BUDGET_YEAR=" & txtBudgetYear.Text & " And BUDGET_MONTH = " & C%
                strSQL = strSQL & " AND BD_FREEZE=0"
                rsTemp.Open strSQL, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
                If rsTemp.EOF = False And rsTemp.EOF = False Then
                    Do
                        FTA = Null: FTS = Null: TMA = Null: TMS = Null
                        'Ticket #21876
                        If rsTemp("BD_FTE") = 0 Or IsNull(rsTemp("BD_FTE")) Then
                            strSQL = "SELECT Count(HREMP.ED_EMPNBR) AS EMPCNT, HREMP.ED_PT, HRJOB.JB_STATUS "
                        Else
                            strSQL = "SELECT Sum(HR_JOB_HISTORY.JH_FTENUM) AS EMPCNT, HREMP.ED_PT, HRJOB.JB_STATUS "
                        End If
                        If glbOracle Then
                            strSQL = strSQL & "FROM HREMP, HR_JOB_HISTORY, HRJOB "
                            strSQL = strSQL & "WHERE HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE AND (HREMP.ED_DOH <= " & Date_SQL(DTE) & ") AND (HR_JOB_HISTORY.JH_CURRENT<>0)"
                        Else
                            strSQL = strSQL & "FROM  (HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
                            strSQL = strSQL & "WHERE (HREMP.ED_DOH <= " & Date_SQL(DTE) & ")  AND (HR_JOB_HISTORY.JH_CURRENT<>0) "
                        End If
                        If Len(rsTemp("BD_DEPT")) > 0 Then
                            strSQL = strSQL & "AND ED_DEPTNO='" & rsTemp("BD_DEPT") & "' "
                        End If
                        If Len(rsTemp("GL_NUMBER")) > 0 Then
                            strSQL = strSQL & "AND ED_GLNO='" & rsTemp("GL_NUMBER") & "' "
                        End If
                        If Len(rsTemp("BD_ADMINBY")) > 0 Then
                            strSQL = strSQL & "AND ED_ADMINBY='" & rsTemp("BD_ADMINBY") & "' "
                        End If
                        If Len(rsTemp("BD_DIV")) > 0 Then
                            strSQL = strSQL & "AND ED_DIV='" & rsTemp("BD_DIV") & "' "
                        End If
                        strSQL = strSQL & "GROUP BY  HREMP.ED_PT, HRJOB.JB_STATUS "
                        strSQL = strSQL & "HAVING HRJOB.JB_STATUS <> 'NA' "
                        
                        'TS Tech - FT and TMP only
                        If glbCompSerial = "S/N - 2369W" Then
                            strSQL = strSQL & " AND (HREMP.ED_PT='FT' Or HREMP.ED_PT='TMP') "
                        End If
                        
                        rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
                        If rsDATA.EOF = False And rsDATA.BOF = False Then
                            Do
                                'TS Tech - Count of FT and TMP
                                If glbCompSerial = "S/N - 2369W" And rsDATA("JB_STATUS") = "L" Then
                                    If rsDATA("ED_PT") = "FT" Then
                                        If IsNull(FTS) Then FTS = 0
                                        FTS = FTS + rsDATA("EMPCNT")
                                    Else
                                        If IsNull(TMS) Then TMS = 0
                                        TMS = TMS + rsDATA("EMPCNT")
                                    End If
                                Else
                                    'All other clients - Count of FT and other Categories altogher
                                    If rsDATA("ED_PT") = "FT" Then
                                        If IsNull(FTA) Then FTA = 0
                                        FTA = FTA + IIf(IsNull(rsDATA("EMPCNT")), 0, rsDATA("EMPCNT"))
                                    Else
                                        If IsNull(TMA) Then TMA = 0
                                        TMA = TMA + IIf(IsNull(rsDATA("EMPCNT")), 0, rsDATA("EMPCNT"))
                                    End If
                                End If
                                rsDATA.MoveNext
                            Loop Until rsDATA.EOF
                        End If
                        rsDATA.Close
                    
   
                        'Changed by Bryan 07/Mar/2006 Ticket#10493
                        'Will select terminated employees terminated on the last day of the month.
                        '**********************
                        If C% < X% Then
                            DTE = getEOM(C%) & "/" & MonthName(C%, True) & "/" & (txtBudgetYear.Text + 1)
                        Else
                            DTE = getEOM(C%) & "/" & MonthName(C%, True) & "/" & txtBudgetYear
                        End If
                        
                        If CDate(DTE) < Date Then
                            'Clear the temp table for term employees
                            strSQL = "DELETE FROM HRMANTERM WHERE WRKEMP='" & glbUserID & "'"
                            gdbAdoIhr001W.BeginTrans
                            gdbAdoIhr001W.Execute strSQL
                            gdbAdoIhr001W.CommitTrans
                            
                            'Find Actual Employees from the terminated for this month(c)
                            strSQL = "SELECT TERM_HREMP.ED_EMPNBR, TERM_HREMP.ED_DEPTNO, TERM_HREMP.ED_GLNO, TERM_HREMP.ED_PT, TERM_HREMP.ED_ADMINBY, Term_JOB_HISTORY.JH_JOB, TERM_HREMP.ED_LOC, TERM_HREMP.ED_DIV, Term_JOB_HISTORY.JH_FTENUM  "
                            If glbOracle Then
                                strSQL = strSQL & "FROM  TERM_HREMP, TERM_HRTRMEMP, Term_JOB_HISTORY "
                                strSQL = strSQL & "WHERE TERM_HREMP.ED_EMPNBR = TERM_HRTRMEMP.TERM_SEQ AND TERM_HREMP.TERM_SEQ = Term_JOB_HISTORY.JH_EMPNBR AND (TERM_HREMP.ED_DOH <= " & Date_SQL(DTE) & ") AND (TERM_HRTRMEMP.TERM_DOT = " & Date_SQL(DTE) & ") "
                            Else
                                strSQL = strSQL & "FROM  ((TERM_HREMP INNER JOIN TERM_HRTRMEMP ON TERM_HREMP.TERM_SEQ = TERM_HRTRMEMP.TERM_SEQ) INNER JOIN Term_JOB_HISTORY ON TERM_HREMP.TERM_SEQ = Term_JOB_HISTORY.TERM_SEQ)"
                                strSQL = strSQL & "WHERE (TERM_HREMP.ED_DOH <= " & Date_SQL(DTE) & ") AND (TERM_HRTRMEMP.TERM_DOT = " & Date_SQL(DTE) & ") "
                            End If
                            strSQL = strSQL & "AND Term_JOB_HISTORY.JH_CURRENT<>0"
                            If Len(rsTemp("BD_DEPT")) > 0 Then
                                strSQL = strSQL & "AND ED_DEPTNO='" & rsTemp("BD_DEPT") & "' "
                            End If
                            If Len(rsTemp("GL_NUMBER")) > 0 Then
                                strSQL = strSQL & "AND ED_GLNO='" & rsTemp("GL_NUMBER") & "' "
                            End If
                            If Len(rsTemp("BD_ADMINBY")) > 0 Then
                                strSQL = strSQL & "AND ED_ADMINBY='" & rsTemp("BD_ADMINBY") & "' "
                            End If
                            If Len(rsTemp("BD_DIV")) > 0 Then
                                strSQL = strSQL & "AND ED_DIV='" & rsTemp("BD_DIV") & "' "
                            End If
                            
                            'TS Tech - FT and TMP only
                            If glbCompSerial = "S/N - 2369W" Then
                                strSQL = strSQL & "AND (TERM_HREMP.ED_PT='FT' Or TERM_HREMP.ED_PT='TMP') "
                            End If
                            
                            rsDATA.Open strSQL, gdbAdoIhr001X, adOpenStatic, adLockOptimistic
                            If rsDATA.EOF = False And rsDATA.BOF = False Then
                                Do
                                    'Insert Term employees into temp. table
                                    strSQL = "INSERT INTO HRMANTERM (ED_EMPNBR, ED_DEPTNO, ED_GLNO, ED_PT, ED_ADMINBY, JH_JOB, ED_LOC, ED_DIV, JH_FTENUM, WRKEMP) "
                                    'strSQL = strSQL & "VALUES (" & rsDATA("ED_EMPNBR") & ", '" & rsDATA("ED_DEPTNO") & "', '" & rsDATA("ED_GLNO") & "', '" & rsDATA("ED_PT") & "', '" & rsDATA("ED_ADMINBY") & "', '" & rsDATA("JH_JOB") & "', '" & rsDATA("ED_LOC") & "', '" & rsDATA("ED_DIV") & "', " & rsDATA("JH_FTENUM") & ", '" & glbUserID & "')"
                                    'Ticket #25050 Franks 02/07/2014 to fix null issue of JH_FTENUM
                                    strSQL = strSQL & "VALUES (" & rsDATA("ED_EMPNBR") & ", '" & rsDATA("ED_DEPTNO") & "', '" & rsDATA("ED_GLNO") & "', '" & rsDATA("ED_PT") & "', '" & rsDATA("ED_ADMINBY") & "', '" & rsDATA("JH_JOB") & "', '" & rsDATA("ED_LOC") & "', '" & rsDATA("ED_DIV") & "', " & IIf(IsNull(rsDATA("JH_FTENUM")), 0, rsDATA("JH_FTENUM")) & ", '" & glbUserID & "')"
                                    gdbAdoIhr001W.BeginTrans
                                    gdbAdoIhr001W.Execute strSQL
                                    gdbAdoIhr001W.CommitTrans
                                    
                                    rsDATA.MoveNext
                                Loop Until rsDATA.EOF
                            End If
                            rsDATA.Close
                            
                            'Join the Terminated Eployees to Job Status and count.
                            If rsTemp("BD_FTE") = 0 Or IsNull(rsTemp("BD_FTE")) Then
                                strSQL = "SELECT Count(HRMANTERM.ED_EMPNBR) AS EMPCNT, HRMANTERM.ED_PT,  HRJOB.JB_STATUS "
                            Else
                                strSQL = "SELECT Sum(HRMANTERM.JH_FTENUM) AS EMPCNT, HRMANTERM.ED_PT,  HRJOB.JB_STATUS "
                            End If
                            If glbOracle Then
                                strSQL = strSQL & "FROM  HRMANTERM, hrjob "
                                strSQL = strSQL & "WHERE HRMANTERM.JH_JOB = hrjob.JB_CODE AND HRMANTERM.WRKEMP = '" & glbUserID & "'"
                            Else
                                strSQL = strSQL & "FROM  HRMANTERM INNER JOIN hrjob ON HRMANTERM.JH_JOB = hrjob.JB_CODE WHERE HRMANTERM.WRKEMP = '" & glbUserID & "'"
                            End If
                            strSQL = strSQL & "GROUP BY HRMANTERM.ED_PT,  HRJOB.JB_STATUS "
                            strSQL = strSQL & "HAVING (hrjob.JB_STATUS<>'NA')"
                            rsDATA.Open strSQL, gdbAdoIhr001W, adOpenStatic, adLockOptimistic
                            If rsDATA.EOF = False And rsDATA.BOF = False Then
                                Do
                                    'TS Tech
                                    If glbCompSerial = "S/N - 2369W" And rsDATA("JB_STATUS") = "L" Then
                                        If rsDATA("ED_PT") = "FT" Then
                                            If IsNull(FTS) Then FTS = 0
                                            FTS = FTS + rsDATA("EMPCNT")
                                        Else
                                            If IsNull(TMS) Then TMS = 0
                                            TMS = TMS + rsDATA("EMPCNT")
                                        End If
                                    Else
                                        'Other clients
                                        If rsDATA("ED_PT") = "FT" Then
                                            If IsNull(FTA) Then FTA = 0
                                            FTA = FTA + IIf(IsNull(rsDATA("EMPCNT")), 0, rsDATA("EMPCNT"))
                                        Else
                                            If IsNull(TMA) Then TMA = 0
                                            TMA = TMA + IIf(IsNull(rsDATA("EMPCNT")), 0, rsDATA("EMPCNT"))
                                        End If
                                    End If
                                    
                                    rsDATA.MoveNext
                                Loop Until rsDATA.EOF
                            End If
                            rsDATA.Close
                        End If
                        
                        rsTemp("ACTUAL_FT_A") = FTA
                        rsTemp("ACTUAL_FT_S") = FTS
                        rsTemp("ACTUAL_TMP_A") = TMA
                        rsTemp("ACTUAL_TMP_s") = TMS

                        rsTemp.Update
                        rsTemp.MoveNext
                    Loop Until rsTemp.EOF
                End If
                rsTemp.Close
            End If
        End If
        
    Next C
             
    MDIMain.panHelp(0).FloodPercent = 15
    
      
    strSQL = "DELETE FROM HRMANWRK WHERE WRKEMP='" & glbUserID & "'"
    gdbAdoIhr001W.BeginTrans
    gdbAdoIhr001W.Execute strSQL
    gdbAdoIhr001W.CommitTrans

    MDIMain.panHelp(0).FloodPercent = 20
    
    'Gather up the data and put into HRMANWRK for general sections
    strSQL = "SELECT DISTINCT GL_NUMBER, BD_ADMINBY, BD_ADMINBY_TABL, BD_DIV FROM HRBUDGET "
    strSQL = strSQL & "WHERE BUDGET_YEAR=" & txtBudgetYear.Text
    If Len(clpDept.Text) > 0 Then
        strSQL = strSQL & " AND BD_DEPT='" & clpDept.Text & "'"
    End If
    If Len(clpGL.Text) > 0 Then
        strSQL = strSQL & " AND GL_NUMBER IN ('" & getCodes(clpGL.Text) & "') "
    End If
    
    rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If rsDATA.EOF = False And rsDATA.BOF = False Then
        Do
            If chkFullTime.Value = True Then
                strSQL = "INSERT INTO HRMANWRK(GL_NUMBER, BD_ADMINBY, BD_ADMINBY_TABL, BD_DIV, BD_YEAR, WRKEMP, BD_ROW) VALUES ('" & rsDATA("GL_NUMBER") & "', '"
                strSQL = strSQL & rsDATA("BD_ADMINBY") & "','" & rsDATA("BD_ADMINBY_TABL") & "','" & rsDATA("BD_DIV") & "', " & txtBudgetYear.Text & ", '" & glbUserID & "', 'BFT')"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
            
            If chkTemporary.Value = True Then
                strSQL = "INSERT INTO HRMANWRK(GL_NUMBER, BD_ADMINBY, BD_ADMINBY_TABL, BD_DIV, BD_YEAR, WRKEMP, BD_ROW) VALUES ('" & rsDATA("GL_NUMBER") & "', '"
                strSQL = strSQL & rsDATA("BD_ADMINBY") & "','" & rsDATA("BD_ADMINBY_TABL") & "','" & rsDATA("BD_DIV") & "', " & txtBudgetYear.Text & ", '" & glbUserID & "', 'BTM')"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
            
            If chkFullTime.Value = True Then
                strSQL = "INSERT INTO HRMANWRK(GL_NUMBER, BD_ADMINBY, BD_ADMINBY_TABL, BD_DIV, BD_YEAR, WRKEMP, BD_ROW) VALUES ('" & rsDATA("GL_NUMBER") & "', '"
                strSQL = strSQL & rsDATA("BD_ADMINBY") & "','" & rsDATA("BD_ADMINBY_TABL") & "','" & rsDATA("BD_DIV") & "', " & txtBudgetYear.Text & ", '" & glbUserID & "', 'AFT')"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
            
            If chkTemporary.Value = True Then
                strSQL = "INSERT INTO HRMANWRK(GL_NUMBER, BD_ADMINBY, BD_ADMINBY_TABL, BD_DIV, BD_YEAR, WRKEMP, BD_ROW) VALUES ('" & rsDATA("GL_NUMBER") & "', '"
                strSQL = strSQL & rsDATA("BD_ADMINBY") & "','" & rsDATA("BD_ADMINBY_TABL") & "','" & rsDATA("BD_DIV") & "', " & txtBudgetYear.Text & ", '" & glbUserID & "', 'ATM')"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
            
            strSQL = "INSERT INTO HRMANWRK(GL_NUMBER, BD_ADMINBY, BD_ADMINBY_TABL, BD_DIV, BD_YEAR, WRKEMP, BD_ROW) VALUES ('" & rsDATA("GL_NUMBER") & "', '"
            strSQL = strSQL & rsDATA("BD_ADMINBY") & "','" & rsDATA("BD_ADMINBY_TABL") & "','" & rsDATA("BD_DIV") & "', " & txtBudgetYear.Text & ", '" & glbUserID & "', 'VAR')"
            gdbAdoIhr001W.BeginTrans
            gdbAdoIhr001W.Execute strSQL
            gdbAdoIhr001W.CommitTrans
            
            rsDATA.MoveNext
        Loop Until rsDATA.EOF
    End If
    rsDATA.Close
    
    MDIMain.panHelp(0).FloodPercent = 25
    
    'TS Tech Ticket#11088
    If glbCompSerial = "S/N - 2369W" Then
        strSQL = "SELECT HR_JOB_HISTORY.JH_JOB, HREMP.ED_GLNO, COUNT(HREMP.ED_EMPNBR) AS cntDIR "
        strSQL = strSQL & " FROM HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
        strSQL = strSQL & " Where (HR_JOB_HISTORY.JH_CURRENT <> 0) "
        strSQL = strSQL & " GROUP BY HR_JOB_HISTORY.JH_JOB, HREMP.ED_GLNO "
        strSQL = strSQL & " HAVING (HR_JOB_HISTORY.JH_JOB = 'MGR') OR (HR_JOB_HISTORY.JH_JOB = 'CO-ORD') "
        rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic
        If rsDATA.EOF = False And rsDATA.BOF = False Then
            Do
                strSQL = " UPDATE HRMANWRK SET "
                If rsDATA("JH_JOB") = "MGR" Then
                    strSQL = strSQL & "BD_MGR=" & rsDATA("cntDIR")
                ElseIf rsDATA("JH_JOB") = "CO-ORD" Then
                    strSQL = strSQL & "BD_COOR=" & rsDATA("cntDIR")
                End If
                strSQL = strSQL & " WHERE GL_NUMBER='" & rsDATA("ED_GLNO") & "' AND BD_YEAR=" & txtBudgetYear.Text
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
                
                rsDATA.MoveNext
            Loop Until rsDATA.EOF
        End If
        rsDATA.Close
    End If
    
    Dim TMP#
    Dim BFT As Variant, BTM As Variant, AFT As Variant, ATM As Variant
    
    strSQL = "SELECT SUM(BUDGET_FT_A) as BFTA,  SUM(BUDGET_FT_S) AS BFTS, SUM(BUDGET_TMP_A) as BTMA, SUM(BUDGET_TMP_S) AS BTMS,  "
    strSQL = strSQL & " SUM(ACTUAL_FT_A) as AFTA,  SUM(ACTUAL_FT_S) AS AFTS, SUM(ACTUAL_TMP_A) AS ATMA, SUM(ACTUAL_TMP_S) AS ATMS, "
    strSQL = strSQL & " GL_NUMBER, BD_ADMINBY, MONTH_SEQ, BD_DIV "
    strSQL = strSQL & " FROM HRBUDGET WHERE BUDGET_YEAR=" & txtBudgetYear.Text
    strSQL = strSQL & " GROUP BY GL_NUMBER, BD_ADMINBY, MONTH_SEQ, BD_DIV"
    rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If rsDATA.EOF = False And rsDATA.BOF = False Then
        Do
            BFT = Null: BTM = Null:  AFT = Null: ATM = Null
            
            If IsNull(rsDATA("BFTA")) = False Or IsNull(rsDATA("BFTS")) = False Then
                BFT = nz(rsDATA("bfta"), 0) + nz(rsDATA("BFTS"), 0)
                strSQL = "UPDATE HRMANWRK SET SEQ" & rsDATA("month_seq") & "='" & BFT & "' "
                strSQL = strSQL & "WHERE GL_NUMBER='" & rsDATA("GL_NUMBER") & "' AND BD_ADMINBY='" & rsDATA("BD_ADMINBY") & "' AND BD_Row='BFT' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                strSQL = strSQL & " AND BD_DIV='" & rsDATA("BD_DIV") & "'"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
            If IsNull(rsDATA("BTMA")) = False Or IsNull(rsDATA("BTMS")) = False Then
                BTM = nz(rsDATA("bTMa"), 0) + nz(rsDATA("BTMS"), 0)
                strSQL = "UPDATE HRMANWRK SET SEQ" & rsDATA("month_seq") & "='" & BTM & "' "
                strSQL = strSQL & "WHERE GL_NUMBER='" & rsDATA("GL_NUMBER") & "' AND BD_ADMINBY='" & rsDATA("BD_ADMINBY") & "' AND BD_Row='BTM' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                strSQL = strSQL & " AND BD_DIV='" & rsDATA("BD_DIV") & "'"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
            If IsNull(rsDATA("AFTA")) = False Or IsNull(rsDATA("AFTS")) = False Then
                AFT = nz(rsDATA("AFTa"), 0) + nz(rsDATA("AFTS"), 0)
                strSQL = "UPDATE HRMANWRK SET SEQ" & rsDATA("month_seq") & "='" & AFT & "' "
                strSQL = strSQL & "WHERE GL_NUMBER='" & rsDATA("GL_NUMBER") & "' AND BD_ADMINBY='" & rsDATA("BD_ADMINBY") & "' AND BD_Row='AFT' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                strSQL = strSQL & " AND BD_DIV='" & rsDATA("BD_DIV") & "'"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
            If IsNull(rsDATA("ATMA")) = False Or IsNull(rsDATA("ATMS")) = False Then
                ATM = nz(rsDATA("ATMa"), 0) + nz(rsDATA("ATMS"), 0)
                strSQL = "UPDATE HRMANWRK SET SEQ" & rsDATA("month_seq") & "='" & ATM & "' "
                strSQL = strSQL & "WHERE GL_NUMBER='" & rsDATA("GL_NUMBER") & "' AND BD_ADMINBY='" & rsDATA("BD_ADMINBY") & "' AND BD_Row='ATM' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                strSQL = strSQL & " AND BD_DIV='" & rsDATA("BD_DIV") & "'"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
                        
            If (IsNull(IIf(chkFullTime.Value = True, AFT, Null)) And IsNull(IIf(chkTemporary.Value = True, ATM, Null))) _
                Or (IsNull(IIf(chkFullTime.Value = True, BFT, Null)) And IsNull(IIf(chkTemporary.Value = True, BTM, Null))) Then
            Else
                If chkFullTime.Value = False Then
                    TMP# = (nz(ATM, 0)) - (nz(BTM, 0))
                ElseIf chkTemporary.Value = False Then
                    TMP# = (nz(AFT, 0)) - (nz(BFT, 0))
                Else
                    TMP# = (nz(AFT, 0) + nz(ATM, 0)) - (nz(BFT, 0) + nz(BTM, 0))
                End If
                If glbCompSerial = "S/N - 2369W" Then
                    strSQL = "UPDATE HRMANWRK SET SEQ" & rsDATA("month_seq") & "='" & FNegs(TMP#) & "' "
                Else
                    strSQL = "UPDATE HRMANWRK SET SEQ" & rsDATA("month_seq") & "='" & FNegs(TMP#, "0") & "' "
                End If
                strSQL = strSQL & "WHERE GL_NUMBER='" & rsDATA("GL_NUMBER") & "' AND BD_ADMINBY='" & rsDATA("BD_ADMINBY") & "' AND BD_Row='VAR' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                strSQL = strSQL & " AND BD_DIV='" & rsDATA("BD_DIV") & "'"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
            rsDATA.MoveNext
        Loop Until rsDATA.EOF
    End If
    rsDATA.Close
    MDIMain.panHelp(0).FloodPercent = 30
    

    MDIMain.panHelp(0).FloodPercent = 35
    
    strSQL = "SELECT seq1, seq2, seq3, seq4, seq5, seq6, seq7, seq8, seq9, seq10, seq11, seq12, GL_Number, BD_ADMINBY, BD_ROW, BD_DIV "
    strSQL = strSQL & "FROM HRMANWRK WHERE (BD_Row='VAR' or BD_Row='BFT' or BD_Row='BTM' or BD_Row='AFT' or BD_Row='ATM' ) AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='"
    strSQL = strSQL & glbUserID & "' GROUP BY seq1, seq2, seq3, seq4, seq5, seq6, seq7, seq8, seq9, seq10, seq11, seq12, GL_Number, BD_ADMINBY, BD_ROW, BD_DIV "
    
    Dim total#, CNT#
    
    rsDATA.Open strSQL, gdbAdoIhr001W, adOpenStatic, adLockOptimistic
    If rsDATA.EOF = False And rsDATA.BOF = False Then
        Do
            total# = 0
            CNT# = 0
            If Not IsNull(rsDATA("seq1")) Then
                total# = total# + CDbl(rsDATA("seq1"))
                CNT# = CNT# + 1
            End If
            If Not IsNull(rsDATA("seq2")) Then
                total# = total# + CDbl(rsDATA("seq2"))
                CNT# = CNT# + 1
            End If
            If Not IsNull(rsDATA("seq3")) Then
                total# = total# + CDbl(rsDATA("seq3"))
                CNT# = CNT# + 1
            End If
            If Not IsNull(rsDATA("seq4")) Then
                total# = total# + CDbl(rsDATA("seq4"))
                CNT# = CNT# + 1
            End If
            If Not IsNull(rsDATA("seq5")) Then
                total# = total# + CDbl(rsDATA("seq5"))
                CNT# = CNT# + 1
            End If
            If Not IsNull(rsDATA("seq6")) Then
                total# = total# + CDbl(rsDATA("seq6"))
                CNT# = CNT# + 1
            End If
            If Not IsNull(rsDATA("seq7")) Then
                total# = total# + CDbl(rsDATA("seq7"))
                CNT# = CNT# + 1
            End If
            If Not IsNull(rsDATA("seq8")) Then
                total# = total# + CDbl(rsDATA("seq8"))
                CNT# = CNT# + 1
            End If
            If Not IsNull(rsDATA("seq9")) Then
                total# = total# + CDbl(rsDATA("seq9"))
                CNT# = CNT# + 1
            End If
            If Not IsNull(rsDATA("seq10")) Then
                total# = total# + CDbl(rsDATA("seq10"))
                CNT# = CNT# + 1
            End If
            If Not IsNull(rsDATA("seq11")) Then
                total# = total# + CDbl(rsDATA("seq11"))
                CNT# = CNT# + 1
            End If
            If Not IsNull(rsDATA("seq12")) Then
                total# = total# + CDbl(rsDATA("seq12"))
                CNT# = CNT# + 1
            End If
            If CNT# > 0 Then
                total# = total# / CNT#
                If glbCompSerial = "S/N - 2369W" Then
                    strSQL = "UPDATE HRMANWRK SET AvgVal='" & FNegs(total#) & "' WHERE GL_Number='" & rsDATA("GL_NUMBER") & "' AND BD_ADMINBY='" & rsDATA("BD_ADMINBY") & "' AND BD_ROW='" & rsDATA("BD_ROW") & "' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                Else
                    strSQL = "UPDATE HRMANWRK SET AvgVal='" & FNegs(total#, "0") & "' WHERE GL_Number='" & rsDATA("GL_NUMBER") & "' AND BD_ADMINBY='" & rsDATA("BD_ADMINBY") & "' AND BD_ROW='" & rsDATA("BD_ROW") & "' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                End If
                strSQL = strSQL & " AND BD_DIV='" & rsDATA("BD_DIV") & "'"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
            
            rsDATA.MoveNext
        Loop Until rsDATA.EOF
    End If
    rsDATA.Close
    
    Call selattwrk2
    Call selattwrk3
    
            
exH:
    MDIMain.panHelp(0).FloodPercent = 100
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    Set rsDATA = Nothing
    Screen.MousePointer = DEFAULT
    Exit Sub
Eh:
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "selAttWrk", "Manpower Report", "Select")
    Resume exH
    
End Sub

Private Sub selattwrk2()
    On Error GoTo Eh
    Dim strSQL As String
    Dim strCri As String
    Dim rsDATA As New ADODB.Recordset
    Dim CntAr() As String
    Dim X%
    Dim errNote As String
    
    If Len(clpDept.Text) > 0 Then
        strCri = " BD_DEPT='" & clpDept.Text & "' "
    Else
        strCri = ""
    End If
    
    If Len(clpGL.Text) > 0 Then
        If Len(strCri) > 0 Then
            strCri = strCri & " AND GL_NUMBER IN ('" & getCodes(clpGL.Text) & "') "
        Else
            strCri = " GL_NUMBER IN ('" & getCodes(clpGL.Text) & "') "
        End If
    End If
    
    strSQL = "SELECT DISTINCT BD_ADMINBY, BD_ADMINBY_TABL FROM HRBUDGET"
    If Len(strCri) > 0 Then
        strSQL = strSQL & " WHERE " & strCri
    End If
    errNote = "Setup"
    
    rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If rsDATA.EOF = False And rsDATA.BOF = False Then
        ReDim CntAr(rsDATA.RecordCount, 10)
        Do
            If Not IsNull(rsDATA("BD_AdminBY")) Then
                CntAr(X%, 0) = rsDATA("BD_AdminBY")
            End If
            
            X% = X% + 1
            
            If chkFullTime.Value = True Then
                strSQL = "INSERT INTO HRMANWRK( BD_ADMINBY_TABL, BD_ADMINBY, BD_YEAR, WRKEMP, BD_ROW) VALUES ('"
                strSQL = strSQL & rsDATA("BD_ADMINBY_TABL") & "', '" & rsDATA("BD_ADMINBY") & "', " & txtBudgetYear.Text & ", '" & glbUserID & "', 'BFT1')"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
            
            If chkTemporary.Value = True Then
                strSQL = "INSERT INTO HRMANWRK( BD_ADMINBY_TABL, BD_ADMINBY, BD_YEAR, WRKEMP, BD_ROW) VALUES ('"
                strSQL = strSQL & rsDATA("BD_ADMINBY_TABL") & "', '" & rsDATA("BD_ADMINBY") & "', " & txtBudgetYear.Text & ", '" & glbUserID & "', 'BTM1')"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
            
            If chkFullTime.Value = True Then
                strSQL = "INSERT INTO HRMANWRK( BD_ADMINBY_TABL, BD_ADMINBY, BD_YEAR, WRKEMP, BD_ROW) VALUES ('"
                strSQL = strSQL & rsDATA("BD_ADMINBY_TABL") & "', '" & rsDATA("BD_ADMINBY") & "', " & txtBudgetYear.Text & ", '" & glbUserID & "', 'AFT1')"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
            
            If chkTemporary.Value = True Then
                strSQL = "INSERT INTO HRMANWRK( BD_ADMINBY_TABL, BD_ADMINBY, BD_YEAR, WRKEMP, BD_ROW) VALUES ('"
                strSQL = strSQL & rsDATA("BD_ADMINBY_TABL") & "', '" & rsDATA("BD_ADMINBY") & "', " & txtBudgetYear.Text & ", '" & glbUserID & "', 'ATM1')"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
            
            strSQL = "INSERT INTO HRMANWRK( BD_ADMINBY_TABL, BD_ADMINBY, BD_YEAR, WRKEMP, BD_ROW) VALUES ('"
            strSQL = strSQL & rsDATA("BD_ADMINBY_TABL") & "', '" & rsDATA("BD_ADMINBY") & "', " & txtBudgetYear.Text & ", '" & glbUserID & "', 'VAR1')"
            gdbAdoIhr001W.BeginTrans
            gdbAdoIhr001W.Execute strSQL
            gdbAdoIhr001W.CommitTrans
            
            strSQL = "INSERT INTO HRMANWRK( BD_ADMINBY_TABL, BD_ADMINBY, BD_YEAR, WRKEMP, BD_ROW) VALUES ('"
            strSQL = strSQL & rsDATA("BD_ADMINBY_TABL") & "', '" & rsDATA("BD_ADMINBY") & "', " & txtBudgetYear.Text & ", '" & glbUserID & "', 'TB1')"
            gdbAdoIhr001W.BeginTrans
            gdbAdoIhr001W.Execute strSQL
            gdbAdoIhr001W.CommitTrans
            
            strSQL = "INSERT INTO HRMANWRK( BD_ADMINBY_TABL, BD_ADMINBY, BD_YEAR, WRKEMP, BD_ROW) VALUES ('"
            strSQL = strSQL & rsDATA("BD_ADMINBY_TABL") & "', '" & rsDATA("BD_ADMINBY") & "', " & txtBudgetYear.Text & ", '" & glbUserID & "', 'TA1')"
            gdbAdoIhr001W.BeginTrans
            gdbAdoIhr001W.Execute strSQL
            gdbAdoIhr001W.CommitTrans
            
            strSQL = "INSERT INTO HRMANWRK( BD_ADMINBY_TABL, BD_ADMINBY, BD_YEAR, WRKEMP, BD_ROW) VALUES ('"
            strSQL = strSQL & rsDATA("BD_ADMINBY_TABL") & "', '" & rsDATA("BD_ADMINBY") & "', " & txtBudgetYear.Text & ", '" & glbUserID & "', 'VFT1')"
            gdbAdoIhr001W.BeginTrans
            gdbAdoIhr001W.Execute strSQL
            gdbAdoIhr001W.CommitTrans
            
            strSQL = "INSERT INTO HRMANWRK( BD_ADMINBY_TABL, BD_ADMINBY, BD_YEAR, WRKEMP, BD_ROW) VALUES ('"
            strSQL = strSQL & rsDATA("BD_ADMINBY_TABL") & "', '" & rsDATA("BD_ADMINBY") & "', " & txtBudgetYear.Text & ", '" & glbUserID & "', 'VTM1')"
            gdbAdoIhr001W.BeginTrans
            gdbAdoIhr001W.Execute strSQL
            gdbAdoIhr001W.CommitTrans
            
            rsDATA.MoveNext
        Loop Until rsDATA.EOF
    End If
    rsDATA.Close
    
    MDIMain.panHelp(0).FloodPercent = 40
    
    errNote = "Adminby"
    For X% = 0 To UBound(CntAr, 1)
        CntAr(X%, 1) = "0"
        CntAr(X%, 2) = "0"
        CntAr(X%, 3) = "0"
        CntAr(X%, 4) = "0"
        CntAr(X%, 5) = "0"
        CntAr(X%, 6) = "0"
        CntAr(X%, 7) = "0"
        CntAr(X%, 8) = "0"
        CntAr(X%, 9) = "0"
        CntAr(X%, 10) = "0"
    Next X%
   
    Dim TMP#
    Dim BFT As Variant, BTM As Variant, AFT As Variant, ATM As Variant
    
    strSQL = "SELECT SUM(BUDGET_FT_A) as BFTA,  SUM(BUDGET_FT_S) AS BFTS, SUM(BUDGET_TMP_A) as BTMA, SUM(BUDGET_TMP_S) AS BTMS,   SUM(ACTUAL_FT_A) as AFTA,  SUM(ACTUAL_FT_S) AS AFTS, SUM(ACTUAL_TMP_A) AS ATMA, SUM(ACTUAL_TMP_S) AS ATMS,  BD_ADMINBY, MONTH_SEQ "
    strSQL = strSQL & "FROM HRBUDGET WHERE BUDGET_YEAR=" & txtBudgetYear.Text
    If Len(strCri) > 0 Then strSQL = strSQL & " AND " & strCri
    strSQL = strSQL & " GROUP BY BD_ADMINBY, MONTH_SEQ "
    rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If rsDATA.EOF = False And rsDATA.BOF = False Then
        Do
            TMP# = 0
            For X% = 0 To UBound(CntAr, 1)
                If CntAr(X%, 0) = rsDATA("BD_ADMINBY") Then Exit For
            Next X%
            
            BFT = Null: BTM = Null:  AFT = Null: ATM = Null
            If IsNull(rsDATA("BFTA")) = False Or IsNull(rsDATA("BFTS")) = False Then
                BFT = nz(rsDATA("bfta"), 0) + nz(rsDATA("BFTS"), 0)
                strSQL = "UPDATE HRMANWRK SET SEQ" & rsDATA("month_seq") & "='" & BFT & "' "
                strSQL = strSQL & "WHERE BD_ADMINBY='" & rsDATA("BD_ADMINBY") & "' AND BD_Row='BFT1' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
                CntAr(X%, 1) = CLng(CntAr(X%, 1)) + 1
            End If
            If IsNull(rsDATA("BTMA")) = False Or IsNull(rsDATA("BTMS")) = False Then
                BTM = nz(rsDATA("bTMa"), 0) + nz(rsDATA("BTMS"), 0)
                strSQL = "UPDATE HRMANWRK SET SEQ" & rsDATA("month_seq") & "='" & BTM & "' "
                strSQL = strSQL & "WHERE BD_ADMINBY='" & rsDATA("BD_ADMINBY") & "'  AND BD_Row='BTM1' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
                CntAr(X%, 2) = CLng(CntAr(X%, 2)) + 1
            End If
            If IsNull(rsDATA("AFTA")) = False Or IsNull(rsDATA("AFTS")) = False Then
                AFT = nz(rsDATA("Afta"), 0) + nz(rsDATA("AFTS"), 0)
                strSQL = "UPDATE HRMANWRK SET SEQ" & rsDATA("month_seq") & "='" & AFT & "' "
                strSQL = strSQL & "WHERE BD_ADMINBY='" & rsDATA("BD_ADMINBY") & "'  AND BD_Row='AFT1' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
                CntAr(X%, 4) = CLng(CntAr(X%, 4)) + 1
            End If
            If IsNull(rsDATA("ATMA")) = False Or IsNull(rsDATA("ATMS")) = False Then
                ATM = nz(rsDATA("ATMa"), 0) + nz(rsDATA("ATMS"), 0)
                strSQL = "UPDATE HRMANWRK SET SEQ" & rsDATA("month_seq") & "='" & ATM & "' "
                strSQL = strSQL & "WHERE BD_ADMINBY='" & rsDATA("BD_ADMINBY") & "'  AND BD_Row='ATM1' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
                CntAr(X%, 5) = CLng(CntAr(X%, 5)) + 1
            End If
                        
            If IsNull(BFT) = False Or IsNull(BTM) = False Then
                TMP# = nz(BFT, 0) + nz(BTM, 0)
                If glbCompSerial = "S/N - 2369W" Then
                    strSQL = "UPDATE HRMANWRK SET SEQ" & rsDATA("month_seq") & "='" & FNegs(TMP#) & "' "
                Else
                    strSQL = "UPDATE HRMANWRK SET SEQ" & rsDATA("month_seq") & "='" & FNegs(TMP#, "0") & "' "
                End If
                strSQL = strSQL & "WHERE BD_ADMINBY='" & rsDATA("BD_ADMINBY") & "'  AND BD_Row='TB1' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
                CntAr(X%, 7) = CLng(CntAr(X%, 7)) + 1
            End If
            
            If IsNull(AFT) = False Or IsNull(ATM) = False Then
                TMP# = nz(AFT, 0) + nz(ATM, 0)
                If glbCompSerial = "S/N - 2369W" Then
                    strSQL = "UPDATE HRMANWRK SET SEQ" & rsDATA("month_seq") & "='" & FNegs(TMP#) & "' "
                Else
                    strSQL = "UPDATE HRMANWRK SET SEQ" & rsDATA("month_seq") & "='" & FNegs(TMP#, "0") & "' "
                End If
                strSQL = strSQL & "WHERE BD_ADMINBY='" & rsDATA("BD_ADMINBY") & "'  AND BD_Row='TA1' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
                CntAr(X%, 8) = CLng(CntAr(X%, 8)) + 1
            End If
            
            If IsNull(AFT) = False And IsNull(BFT) = False Then
                TMP# = nz(AFT, 0) - nz(BFT, 0)
                If glbCompSerial = "S/N - 2369W" Then
                    strSQL = "UPDATE HRMANWRK SET SEQ" & rsDATA("month_seq") & "='" & FNegs(TMP#) & "' "
                Else
                    strSQL = "UPDATE HRMANWRK SET SEQ" & rsDATA("month_seq") & "='" & FNegs(TMP#, "0") & "' "
                End If
                strSQL = strSQL & "WHERE BD_ADMINBY='" & rsDATA("BD_ADMINBY") & "'  AND BD_Row='VFT1' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
                CntAr(X%, 9) = CLng(CntAr(X%, 9)) + 1
            End If
            
            If IsNull(ATM) = False And IsNull(BTM) = False Then
                TMP# = nz(ATM, 0) - nz(BTM, 0)
                If glbCompSerial = "S/N - 2369W" Then
                    strSQL = "UPDATE HRMANWRK SET SEQ" & rsDATA("month_seq") & "='" & FNegs(TMP#) & "' "
                Else
                    strSQL = "UPDATE HRMANWRK SET SEQ" & rsDATA("month_seq") & "='" & FNegs(TMP#, "0") & "' "
                End If
                strSQL = strSQL & "WHERE BD_ADMINBY='" & rsDATA("BD_ADMINBY") & "'  AND BD_Row='VTM1' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
                CntAr(X%, 10) = CLng(CntAr(X%, 10)) + 1
            End If
                
            If (IsNull(IIf(chkFullTime.Value = True, AFT, Null)) And IsNull(IIf(chkTemporary.Value = True, ATM, Null))) _
                Or (IsNull(IIf(chkFullTime.Value = True, BFT, Null)) And IsNull(IIf(chkTemporary.Value = True, BTM, Null))) Then
            Else
                If chkFullTime.Value = False Then
                    TMP# = (nz(ATM, 0)) - (nz(BTM, 0))
                ElseIf chkTemporary.Value = False Then
                    TMP# = (nz(AFT, 0)) - (nz(BFT, 0))
                Else
                    TMP# = (nz(AFT, 0) + nz(ATM, 0)) - (nz(BFT, 0) + nz(BTM, 0))
                End If
                If glbCompSerial = "S/N - 2369W" Then
                    strSQL = "UPDATE HRMANWRK SET SEQ" & rsDATA("month_seq") & "='" & FNegs(TMP#) & "' "
                Else
                    strSQL = "UPDATE HRMANWRK SET SEQ" & rsDATA("month_seq") & "='" & FNegs(TMP#, "0") & "' "
                End If
                strSQL = strSQL & "WHERE BD_ADMINBY='" & rsDATA("BD_ADMINBY") & "'  AND BD_Row='VAR1' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
            
            rsDATA.MoveNext
        Loop Until rsDATA.EOF
    End If
    rsDATA.Close
    
    MDIMain.panHelp(0).FloodPercent = 50
    
    errNote = "Adding"
    strSQL = "SELECT SUM(BUDGET_FT_A) as BFTA,  SUM(BUDGET_FT_S) AS BFTS, SUM(BUDGET_TMP_A) as BTMA, SUM(BUDGET_TMP_S) AS BTMS,  SUM(ACTUAL_FT_A) as AFTA,  SUM(ACTUAL_FT_S) AS AFTS, SUM(ACTUAL_TMP_A) AS ATMA, SUM(ACTUAL_TMP_S) AS ATMS, BD_ADMINBY "
    strSQL = strSQL & "FROM HRBUDGET WHERE BUDGET_YEAR=" & txtBudgetYear.Text
    If Len(strCri) > 0 Then strSQL = strSQL & " AND " & strCri
    strSQL = strSQL & " GROUP BY BD_ADMINBY"
    rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If rsDATA.EOF = False And rsDATA.BOF = False Then
        Do
            For X% = 0 To UBound(CntAr, 1)
                If CntAr(X%, 0) = rsDATA("BD_ADMINBY") Then Exit For
            Next X%
            BFT = Null: BTM = Null:  AFT = Null: ATM = Null
            If IsNull(rsDATA("BFTA")) = False Or IsNull(rsDATA("BFTS")) = False Then
                BFT = nz(rsDATA("BFta"), 0) + nz(rsDATA("BFTS"), 0)
                If glbCompSerial = "S/N - 2369W" Then
                    strSQL = "UPDATE HRMANWRK SET AVGVAL='" & FNegs(BFT / CntAr(X%, 1)) & "' "
                Else
                    strSQL = "UPDATE HRMANWRK SET AVGVAL='" & FNegs(BFT / CntAr(X%, 1), "0") & "' "
                End If
                strSQL = strSQL & "WHERE BD_ADMINBY='" & rsDATA("BD_ADMINBY") & "' AND BD_Row='BFT1' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
            If IsNull(rsDATA("BTMA")) = False Or IsNull(rsDATA("BTMS")) = False Then
                BTM = nz(rsDATA("BTMa"), 0) + nz(rsDATA("BTMS"), 0)
                If glbCompSerial = "S/N - 2369W" Then
                    strSQL = "UPDATE HRMANWRK SET AVGVAL='" & FNegs(BTM / CntAr(X%, 2)) & "' "
                Else
                    strSQL = "UPDATE HRMANWRK SET AVGVAL='" & FNegs(BTM / CntAr(X%, 2), "0") & "' "
                End If
                strSQL = strSQL & "WHERE BD_ADMINBY='" & rsDATA("BD_ADMINBY") & "' AND BD_Row='BTM1' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
            If IsNull(rsDATA("AFTA")) = False Or IsNull(rsDATA("AFTS")) = False Then
                AFT = nz(rsDATA("AFta"), 0) + nz(rsDATA("AFTS"), 0)
                If glbCompSerial = "S/N - 2369W" Then
                    strSQL = "UPDATE HRMANWRK SET AVGVAL='" & FNegs(AFT / CntAr(X%, 4)) & "' "
                Else
                    strSQL = "UPDATE HRMANWRK SET AVGVAL='" & FNegs(AFT / CntAr(X%, 4), "0") & "' "
                End If
                strSQL = strSQL & "WHERE BD_ADMINBY='" & rsDATA("BD_ADMINBY") & "' AND BD_Row='AFT1' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
            If IsNull(rsDATA("ATMA")) = False Or IsNull(rsDATA("ATMS")) = False Then
                ATM = nz(rsDATA("ATMa"), 0) + nz(rsDATA("ATMS"), 0)
                If glbCompSerial = "S/N - 2369W" Then
                    strSQL = "UPDATE HRMANWRK SET AVGVAL='" & FNegs(ATM / CntAr(X%, 5)) & "' "
                Else
                    strSQL = "UPDATE HRMANWRK SET AVGVAL='" & FNegs(ATM / CntAr(X%, 5), "0") & "' "
                End If
                strSQL = strSQL & "WHERE BD_ADMINBY='" & rsDATA("BD_ADMINBY") & "' AND BD_Row='ATM1' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
                        
            If IsNull(BFT) = False Or IsNull(BTM) = False Then
                TMP# = nz(BFT, 0) + nz(BTM, 0)
                If glbCompSerial = "S/N - 2369W" Then
                    strSQL = "UPDATE HRMANWRK SET AVGVAL='" & FNegs(TMP# / CntAr(X%, 7)) & "' "
                Else
                    strSQL = "UPDATE HRMANWRK SET AVGVAL='" & FNegs(TMP# / CntAr(X%, 7), "0") & "' "
                End If
                strSQL = strSQL & "WHERE BD_ADMINBY='" & rsDATA("BD_ADMINBY") & "' AND BD_Row='TB1' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
            
            If IsNull(AFT) = False Or IsNull(ATM) = False Then
                TMP# = nz(AFT, 0) + nz(ATM, 0)
                If glbCompSerial = "S/N - 2369W" Then
                    strSQL = "UPDATE HRMANWRK SET AVGVAL='" & FNegs(TMP# / CntAr(X%, 8)) & "' "
                Else
                    strSQL = "UPDATE HRMANWRK SET AVGVAL='" & FNegs(TMP# / CntAr(X%, 8), "0") & "' "
                End If
                strSQL = strSQL & "WHERE BD_ADMINBY='" & rsDATA("BD_ADMINBY") & "' AND BD_Row='TA1' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
            
            rsDATA.MoveNext
        Loop Until rsDATA.EOF
    End If
    errNote = "Totals"
    rsDATA.Close
    
    MDIMain.panHelp(0).FloodPercent = 60
    
    Dim str As String
    
    For X% = 0 To 2
        Select Case X%
        Case 0
            str = "VAR1"
        Case 1
            str = "VFT1"
        Case 2
            str = "VTM1"
        End Select
        
        strSQL = "SELECT seq1, seq2, seq3, seq4, seq5, seq6, seq7, seq8, seq9, seq10, seq11, seq12, BD_ADMINBY "
        strSQL = strSQL & "FROM HRMANWRK WHERE BD_Row = ('" & str & "') AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='"
        strSQL = strSQL & glbUserID & "' GROUP BY seq1, seq2, seq3, seq4, seq5, seq6, seq7, seq8, seq9, seq10, seq11, seq12, BD_ADMINBY "
        
        Dim total#, CNT#
        
        rsDATA.Open strSQL, gdbAdoIhr001W, adOpenStatic, adLockOptimistic
        If rsDATA.EOF = False And rsDATA.BOF = False Then
            Do
                total# = 0
                CNT# = 0
                If Not IsNull(rsDATA("seq1")) Then
                    total# = total# + CDbl(rsDATA("seq1"))
                    CNT# = CNT# + 1
                End If
                If Not IsNull(rsDATA("seq2")) Then
                    total# = total# + CDbl(rsDATA("seq2"))
                    CNT# = CNT# + 1
                End If
                If Not IsNull(rsDATA("seq3")) Then
                    total# = total# + CDbl(rsDATA("seq3"))
                    CNT# = CNT# + 1
                End If
                If Not IsNull(rsDATA("seq4")) Then
                    total# = total# + CDbl(rsDATA("seq4"))
                    CNT# = CNT# + 1
                End If
                If Not IsNull(rsDATA("seq5")) Then
                    total# = total# + CDbl(rsDATA("seq5"))
                    CNT# = CNT# + 1
                End If
                If Not IsNull(rsDATA("seq6")) Then
                    total# = total# + CDbl(rsDATA("seq6"))
                    CNT# = CNT# + 1
                End If
                If Not IsNull(rsDATA("seq7")) Then
                    total# = total# + CDbl(rsDATA("seq7"))
                    CNT# = CNT# + 1
                End If
                If Not IsNull(rsDATA("seq8")) Then
                    total# = total# + CDbl(rsDATA("seq8"))
                    CNT# = CNT# + 1
                End If
                If Not IsNull(rsDATA("seq9")) Then
                    total# = total# + CDbl(rsDATA("seq9"))
                    CNT# = CNT# + 1
                End If
                If Not IsNull(rsDATA("seq10")) Then
                    total# = total# + CDbl(rsDATA("seq10"))
                    CNT# = CNT# + 1
                End If
                If Not IsNull(rsDATA("seq11")) Then
                    total# = total# + CDbl(rsDATA("seq11"))
                    CNT# = CNT# + 1
                End If
                If Not IsNull(rsDATA("seq12")) Then
                    total# = total# + CDbl(rsDATA("seq12"))
                    CNT# = CNT# + 1
                End If
                If CNT# > 0 Then
                    total# = total# / CNT#
                    If glbCompSerial = "S/N - 2369W" Then
                        strSQL = "UPDATE HRMANWRK SET AvgVal='" & FNegs(total#) & "' WHERE BD_ADMINBY='" & rsDATA("BD_ADMINBY") & "' AND BD_ROW='" & str & "' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                    Else
                        strSQL = "UPDATE HRMANWRK SET AvgVal='" & FNegs(total#, "0") & "' WHERE BD_ADMINBY='" & rsDATA("BD_ADMINBY") & "' AND BD_ROW='" & str & "' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                    End If
                    gdbAdoIhr001W.BeginTrans
                    gdbAdoIhr001W.Execute strSQL
                    gdbAdoIhr001W.CommitTrans
                End If
                
                rsDATA.MoveNext
            Loop Until rsDATA.EOF
        End If
        rsDATA.Close
    Next X%
    errNote = "Variation"
exH:
    Set rsDATA = Nothing
    Exit Sub
Eh:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "selAttWrk2", errNote, "Select")
    Resume exH
End Sub

Private Sub selattwrk3()
    On Error GoTo Eh
    'Key Statistics
    Dim strSQL As String
    Dim strCri As String
    Dim rsDATA As New ADODB.Recordset
    Dim rsTerm As New ADODB.Recordset
    Dim varTMP As Variant
    
    If Len(clpDept.Text) > 0 Then
        strCri = " ED_DEPTNO='" & clpDept.Text & "' "
    Else
        strCri = ""
    End If
    
    If Len(clpGL.Text) > 0 Then
        If Len(strCri) > 0 Then
            strCri = strCri & " AND ED_GLNO IN ('" & getCodes(clpGL.Text) & "') "
        Else
            strCri = " ED_GLNO IN ('" & getCodes(clpGL.Text) & "') "
        End If
    End If

'Setup Key Statistics Rows in working table
            If chkFullTime.Value = True Then
                strSQL = "INSERT INTO HRMANWRK(GL_NUMBER, BD_YEAR, WRKEMP, BD_ROW) VALUES ('Stat', "
                strSQL = strSQL & txtBudgetYear.Text & ", '" & glbUserID & "', 'NHF')"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
            
            If chkTemporary.Value = True Then
                strSQL = "INSERT INTO HRMANWRK(GL_NUMBER, BD_YEAR, WRKEMP, BD_ROW) VALUES ('Stat', "
                strSQL = strSQL & txtBudgetYear.Text & ", '" & glbUserID & "', 'NHT')"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
            
            strSQL = "INSERT INTO HRMANWRK(GL_NUMBER, BD_YEAR, WRKEMP, BD_ROW) VALUES ('Stat', "
            strSQL = strSQL & txtBudgetYear.Text & ", '" & glbUserID & "', 'NHS')"
            gdbAdoIhr001W.BeginTrans
            gdbAdoIhr001W.Execute strSQL
            gdbAdoIhr001W.CommitTrans
            
            If chkFullTime.Value = True Then
                strSQL = "INSERT INTO HRMANWRK(GL_NUMBER, BD_YEAR, WRKEMP, BD_ROW) VALUES ('Stat', "
                strSQL = strSQL & txtBudgetYear.Text & ", '" & glbUserID & "', 'TEF')"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
            
            If chkTemporary.Value = True Then
                strSQL = "INSERT INTO HRMANWRK(GL_NUMBER, BD_YEAR, WRKEMP, BD_ROW) VALUES ('Stat', "
                strSQL = strSQL & txtBudgetYear.Text & ", '" & glbUserID & "', 'TET')"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
            
            strSQL = "INSERT INTO HRMANWRK(GL_NUMBER, BD_YEAR, WRKEMP, BD_ROW) VALUES ('Stat', "
            strSQL = strSQL & txtBudgetYear.Text & ", '" & glbUserID & "', 'TES')"
            gdbAdoIhr001W.BeginTrans
            gdbAdoIhr001W.Execute strSQL
            gdbAdoIhr001W.CommitTrans

            If chkFullTime.Value = True Then
                strSQL = "INSERT INTO HRMANWRK(GL_NUMBER, BD_YEAR, WRKEMP, BD_ROW) VALUES ('Stat', "
                strSQL = strSQL & txtBudgetYear.Text & ", '" & glbUserID & "', 'TUF')"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
            
            If chkTemporary.Value = True Then
                strSQL = "INSERT INTO HRMANWRK(GL_NUMBER, BD_YEAR, WRKEMP, BD_ROW) VALUES ('Stat', "
                strSQL = strSQL & txtBudgetYear.Text & ", '" & glbUserID & "', 'TUT')"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
            
            strSQL = "INSERT INTO HRMANWRK(GL_NUMBER, BD_YEAR, WRKEMP, BD_ROW) VALUES ('Stat', "
            strSQL = strSQL & txtBudgetYear.Text & ", '" & glbUserID & "', 'TUS')"
            gdbAdoIhr001W.BeginTrans
            gdbAdoIhr001W.Execute strSQL
            gdbAdoIhr001W.CommitTrans
            
            If chkFullTime.Value = True Then
                strSQL = "INSERT INTO HRMANWRK(GL_NUMBER, BD_YEAR, WRKEMP, BD_ROW) VALUES ('Stat', "
                strSQL = strSQL & txtBudgetYear.Text & ", '" & glbUserID & "', 'ABF')"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
            
            If chkTemporary.Value = True Then
                strSQL = "INSERT INTO HRMANWRK(GL_NUMBER, BD_YEAR, WRKEMP, BD_ROW) VALUES ('Stat', "
                strSQL = strSQL & txtBudgetYear.Text & ", '" & glbUserID & "', 'ABT')"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
                        
            strSQL = "INSERT INTO HRMANWRK(GL_NUMBER, BD_YEAR, WRKEMP, BD_ROW) VALUES ('Stat', "
            strSQL = strSQL & txtBudgetYear.Text & ", '" & glbUserID & "', 'ABS')"
            gdbAdoIhr001W.BeginTrans
            gdbAdoIhr001W.Execute strSQL
            gdbAdoIhr001W.CommitTrans
            
            strSQL = "INSERT INTO HRMANWRK(GL_NUMBER, BD_YEAR, WRKEMP, BD_ROW) VALUES ('Stat', "
            strSQL = strSQL & txtBudgetYear.Text & ", '" & glbUserID & "', 'TOTS')"
            gdbAdoIhr001W.BeginTrans
            gdbAdoIhr001W.Execute strSQL
            gdbAdoIhr001W.CommitTrans
            
            strSQL = "INSERT INTO HRMANWRK(GL_NUMBER, BD_YEAR, WRKEMP, BD_ROW) VALUES ('Stat', "
            strSQL = strSQL & txtBudgetYear.Text & ", '" & glbUserID & "', 'TOTM')"
            gdbAdoIhr001W.BeginTrans
            gdbAdoIhr001W.Execute strSQL
            gdbAdoIhr001W.CommitTrans
            
            strSQL = "INSERT INTO HRMANWRK(GL_NUMBER, BD_YEAR, WRKEMP, BD_ROW) VALUES ('Stat', "
            strSQL = strSQL & txtBudgetYear.Text & ", '" & glbUserID & "', 'TOTV')"
            gdbAdoIhr001W.BeginTrans
            gdbAdoIhr001W.Execute strSQL
            gdbAdoIhr001W.CommitTrans
            
            strSQL = "INSERT INTO HRMANWRK(GL_NUMBER, BD_YEAR, WRKEMP, BD_ROW) VALUES ('Stat', "
            strSQL = strSQL & txtBudgetYear.Text & ", '" & glbUserID & "', 'TOTF')"
            gdbAdoIhr001W.BeginTrans
            gdbAdoIhr001W.Execute strSQL
            gdbAdoIhr001W.CommitTrans
            
            strSQL = "INSERT INTO HRMANWRK(GL_NUMBER, BD_YEAR, WRKEMP, BD_ROW) VALUES ('Stat', "
            strSQL = strSQL & txtBudgetYear.Text & ", '" & glbUserID & "', 'TOTT')"
            gdbAdoIhr001W.BeginTrans
            gdbAdoIhr001W.Execute strSQL
            gdbAdoIhr001W.CommitTrans
            
            strSQL = "INSERT INTO HRMANWRK(GL_NUMBER, BD_YEAR, WRKEMP, BD_ROW) VALUES ('Stat', "
            strSQL = strSQL & txtBudgetYear.Text & ", '" & glbUserID & "', 'TOTA')"
            gdbAdoIhr001W.BeginTrans
            gdbAdoIhr001W.Execute strSQL
            gdbAdoIhr001W.CommitTrans
            
            strSQL = "INSERT INTO HRMANWRK(GL_NUMBER, BD_YEAR, WRKEMP, BD_ROW) VALUES ('Stat', "
            strSQL = strSQL & txtBudgetYear.Text & ", '" & glbUserID & "', 'VALA')"
            gdbAdoIhr001W.BeginTrans
            gdbAdoIhr001W.Execute strSQL
            gdbAdoIhr001W.CommitTrans

    MDIMain.panHelp(0).FloodPercent = 70
    
    
    Dim TMP#, TFT#, TTM#
    Dim C%, X%, n%
    Dim DTE As String
    
    strSQL = "SELECT BUDGET_MONTH FROM HRBUDGET WHERE MONTH_SEQ=1"
    strSQL = strSQL & " AND BUDGET_YEAR=" & txtBudgetYear.Text
    rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If rsDATA.EOF = False And rsDATA.BOF = False Then
        X% = rsDATA("BUDGET_MONTH")
    End If
    
    rsDATA.Close
    
    For C% = 1 To 12
        TMP# = 0
        
        If X% > 0 Then
            If C% > X% Then 'If month is less than the sequence then it must be next year
                DTE = getEOM(X%) & "/" & MonthName(X%, True) & "/" & (txtBudgetYear.Text + 1)
            Else
                DTE = getEOM(X%) & "/" & MonthName(X%, True) & "/" & txtBudgetYear.Text
            End If
            
            Dim Tdate As String
            
            Tdate = getEOM(month(Date)) & "/"
            Tdate = Tdate & MonthName(month(Date), True) & "/"
            Tdate = Tdate & Year(Date)
            
            'no future new hires
            If CDate(DTE) < CDate(Tdate) Then
                'Get New Hires FT
                strSQL = "SELECT Count(ED_EMPNBR) AS EMPCNT, ED_PT "
                strSQL = strSQL & "FROM HREMP "
                strSQL = strSQL & "WHERE (ED_SENDTE >= " & Date_SQL("1/" & MonthName(X%, True) & "/" & Year(DTE)) & ") AND (ED_SENDTE <= " & Date_SQL(DTE) & ") "
                If Len(strCri) > 0 Then strSQL = strSQL & " AND " & Replace(Replace(strCri, "BD_DEPT", "ED_DEPTNO"), "GL_NUMBER", "ED_GLNO") 'strCri
                strSQL = strSQL & "GROUP BY ED_PT "
                strSQL = strSQL & "HAVING ED_PT='FT'"
        
                rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic
                If rsDATA.EOF = False And rsDATA.BOF = False Then
                    Do
                        TMP# = TMP# + rsDATA("EMPCNT")
                        If rsDATA("ED_PT") = "FT" Then
                            strSQL = "UPDATE HRMANWRK SET SEQ" & C% & "='" & FNegs(rsDATA("EMPCNT")) & "' "
                            strSQL = strSQL & "WHERE GL_NUMBER='Stat' AND BD_Row='NHF' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
'                        ElseIf rsDATA("ED_PT") = "TMP" Then
'                            strSQL = "UPDATE HRMANWRK SET SEQ" & c% & "='" & FNegs(rsDATA("EMPCNT")) & "' "
'                            strSQL = strSQL & "WHERE GL_NUMBER='Stat' AND BD_Row='NHT' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                        End If
                        gdbAdoIhr001W.BeginTrans
                        gdbAdoIhr001W.Execute strSQL
                        gdbAdoIhr001W.CommitTrans
                        rsDATA.MoveNext
                    Loop Until rsDATA.EOF
                End If
                rsDATA.Close
                
                'Get New Hires TMP
                strSQL = "SELECT Count(ED_EMPNBR) AS EMPCNT, ED_PT "
                strSQL = strSQL & "From HREMP "
                strSQL = strSQL & "Where (ED_DOH >= " & Date_SQL("1/" & MonthName(X%, True) & "/" & Year(DTE)) & ") AND (ED_DOH <= " & Date_SQL(DTE) & ") "
                If Len(strCri) > 0 Then strSQL = strSQL & " AND " & Replace(Replace(strCri, "BD_DEPT", "ED_DEPTNO"), "GL_NUMBER", "ED_GLNO") 'strCri
                strSQL = strSQL & "GROUP BY ED_PT "
                strSQL = strSQL & "HAVING ED_PT<>'FT'"
                
                If glbCompSerial = "S/N - 2369W" Then
                    strSQL = strSQL & " AND (ED_PT='TMP')"
                End If
                
                rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic
                If rsDATA.EOF = False And rsDATA.BOF = False Then
                    Do
                        TMP# = TMP# + rsDATA("EMPCNT")
                        If rsDATA("ED_PT") <> "FT" Then
                            strSQL = "UPDATE HRMANWRK SET SEQ" & C% & "='" & FNegs(rsDATA("EMPCNT")) & "' "
                            strSQL = strSQL & "WHERE GL_NUMBER='Stat' AND BD_Row='NHT' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                        End If
                        gdbAdoIhr001W.BeginTrans
                        gdbAdoIhr001W.Execute strSQL
                        gdbAdoIhr001W.CommitTrans
                        rsDATA.MoveNext
                    Loop Until rsDATA.EOF
                End If
                rsDATA.Close
                
                
                'Added Terminated FT 02/Feb/2005 Bryan Ticket#10222
                strSQL = "SELECT Count(ED_EMPNBR) AS EMPCNT, ED_PT "
                strSQL = strSQL & "From Term_HREMP "
                strSQL = strSQL & "Where (ED_SENDTE >= " & Date_SQL("1/" & MonthName(X%, True) & "/" & Year(DTE)) & ") AND (ED_SENDTE <= " & Date_SQL(DTE) & ") "
                If Len(strCri) > 0 Then strSQL = strSQL & " AND " & Replace(Replace(strCri, "BD_DEPT", "ED_DEPTNO"), "GL_NUMBER", "ED_GLNO") 'strCri
                strSQL = strSQL & "GROUP BY ED_PT "
                strSQL = strSQL & "HAVING (ED_PT='FT')"

                rsDATA.Open strSQL, gdbAdoIhr001X, adOpenStatic, adLockOptimistic
                If rsDATA.EOF = False And rsDATA.BOF = False Then
                    Do
                        TMP# = TMP# + rsDATA("EMPCNT")
                        If rsDATA("ED_PT") = "FT" Then
                            strSQL = "SELECT SEQ" & C% & " FROM HRMANWRK "
                            strSQL = strSQL & "WHERE GL_NUMBER='Stat' AND BD_Row='NHF' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                        End If
                        rsTerm.Open strSQL, gdbAdoIhr001W, adOpenKeyset, adLockOptimistic, adCmdText
                        rsTerm("SEQ" & C%) = FNegs(CLng(nz(rsTerm("SEQ" & C%), 0)) + rsDATA("EMPCNT"))
                        rsTerm.Update
                        rsTerm.Close
                        rsDATA.MoveNext
                    Loop Until rsDATA.EOF
                End If
                rsDATA.Close
                
                'Added Terminated TMP 02/Feb/2005 Bryan Ticket#10222
                strSQL = "SELECT Count(ED_EMPNBR) AS EMPCNT, ED_PT "
                strSQL = strSQL & "From Term_HREMP "
                strSQL = strSQL & "Where (ED_DOH >= " & Date_SQL("1/" & MonthName(X%, True) & "/" & Year(DTE)) & ") AND (ED_DOH <= " & Date_SQL(DTE) & ") "
                If Len(strCri) > 0 Then strSQL = strSQL & " AND " & Replace(Replace(strCri, "BD_DEPT", "ED_DEPTNO"), "GL_NUMBER", "ED_GLNO") 'strCri
                strSQL = strSQL & "GROUP BY ED_PT "
                strSQL = strSQL & "HAVING (ED_PT<>'FT')"
                
                If glbCompSerial = "S/N - 2369W" Then
                    strSQL = strSQL & " AND (ED_PT='TMP')"
                End If

                rsDATA.Open strSQL, gdbAdoIhr001X, adOpenStatic, adLockOptimistic
                If rsDATA.EOF = False And rsDATA.BOF = False Then
                    Do
                        TMP# = TMP# + rsDATA("EMPCNT")
                        If rsDATA("ED_PT") <> "FT" Then
                            strSQL = "SELECT SEQ" & C% & " FROM HRMANWRK "
                            strSQL = strSQL & "WHERE GL_NUMBER='Stat' AND BD_Row='NHT' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                        End If
                        rsTerm.Open strSQL, gdbAdoIhr001W, adOpenKeyset, adLockOptimistic, adCmdText
                        rsTerm("SEQ" & C%) = FNegs(CLng(nz(rsTerm("SEQ" & C%), 0)) + rsDATA("EMPCNT"))
                        rsTerm.Update
                        rsTerm.Close
                        rsDATA.MoveNext
                    Loop Until rsDATA.EOF
                End If
                rsDATA.Close
                
                If TMP# > 0 Then
                    strSQL = "UPDATE HRMANWRK SET SEQ" & C% & "='" & FNegs(TMP#) & "' "
                    strSQL = strSQL & "WHERE GL_NUMBER='Stat' AND BD_Row='NHS' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                    gdbAdoIhr001W.BeginTrans
                    gdbAdoIhr001W.Execute strSQL
                    gdbAdoIhr001W.CommitTrans
                End If
                
                MDIMain.panHelp(0).FloodPercent = 80
                
                TMP# = 0
                TTM# = 0
                TFT# = 0
                
                'Get  Terminated
                strSQL = "SELECT Count(TERM_HREMP.ED_EMPNBR) AS EMPCNT, TERM_HREMP.ED_PT  "
                If glbOracle Then
                    strSQL = strSQL & "From TERM_HREMP, TERM_HRTRMEMP "
                    strSQL = strSQL & "Where TERM_HREMP.TERM_SEQ = TERM_HRTRMEMP.TERM_SEQ AND (TERM_HRTRMEMP.TERM_DOT >= " & Date_SQL("1/" & MonthName(X%, True) & "/" & Year(DTE)) & " AND TERM_HRTRMEMP.TERM_DOT <= " & Date_SQL(DTE) & ") "
                Else
                    strSQL = strSQL & "From TERM_HREMP INNER JOIN TERM_HRTRMEMP ON TERM_HREMP.TERM_SEQ = TERM_HRTRMEMP.TERM_SEQ "
                    strSQL = strSQL & "Where (TERM_HRTRMEMP.TERM_DOT >= " & Date_SQL("1/" & MonthName(X%, True) & "/" & Year(DTE)) & " AND TERM_HRTRMEMP.TERM_DOT <= " & Date_SQL(DTE) & ") "
                End If
                If Len(strCri) > 0 Then strSQL = strSQL & " AND " & Replace(Replace(strCri, "BD_DEPT", "ED_DEPTNO"), "GL_NUMBER", "ED_GLNO") 'strCri
                strSQL = strSQL & "GROUP BY TERM_HREMP.ED_PT "
                If glbCompSerial = "S/N - 2369W" Then
                    strSQL = strSQL & "HAVING (TERM_HREMP.ED_PT='FT' Or TERM_HREMP.ED_PT='TMP') "
                End If
        
                rsDATA.Open strSQL, gdbAdoIhr001X, adOpenStatic, adLockOptimistic
                If rsDATA.EOF = False And rsDATA.BOF = False Then
                    Do
                        TMP# = TMP# + rsDATA("EMPCNT")
                        If rsDATA("ED_PT") = "FT" Then
                            strSQL = "UPDATE HRMANWRK SET SEQ" & C% & "='" & FNegs(rsDATA("EMPCNT")) & "' "
                            strSQL = strSQL & "WHERE GL_NUMBER='Stat' AND BD_Row='TEF' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                            TFT# = rsDATA("EMPCNT")
                        ElseIf rsDATA("ED_PT") <> "FT" Then
                            strSQL = "UPDATE HRMANWRK SET SEQ" & C% & "='" & FNegs(rsDATA("EMPCNT")) & "' "
                            strSQL = strSQL & "WHERE GL_NUMBER='Stat' AND BD_Row='TET' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                            TTM# = rsDATA("EMPCNT")
                        End If
                        gdbAdoIhr001W.BeginTrans
                        gdbAdoIhr001W.Execute strSQL
                        gdbAdoIhr001W.CommitTrans
                        rsDATA.MoveNext
                    Loop Until rsDATA.EOF
                    strSQL = "UPDATE HRMANWRK SET SEQ" & C% & "='" & FNegs(TMP#) & "' "
                    strSQL = strSQL & "WHERE GL_NUMBER='Stat' AND BD_Row='TES' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                    gdbAdoIhr001W.BeginTrans
                    gdbAdoIhr001W.Execute strSQL
                    gdbAdoIhr001W.CommitTrans
                End If
                rsDATA.Close
                
                'get Turnover FT
                Dim tmpFT#
                Dim tmpTMP#
                tmpTMP# = 0
                tmpFT# = 0
                TMP# = 0
                
                strSQL = "SELECT sum(ACTUAL_FT_A)  as TOTF,  sum(ACTUAL_FT_S)  as TOTFs "
                strSQL = strSQL & "FROM HRBUDGET WHERE BUDGET_YEAR = " & txtBudgetYear.Text
                strSQL = strSQL & " AND MONTH_SEQ=" & C%
                'If Len(strCri) > 0 Then strSQL = strSQL & " AND " & strCri
                'Ticket #20626 Franks 07/14/2011 fix error "Invalid column 'ED_GLNO'"
                If Len(strCri) > 0 Then strSQL = strSQL & " AND " & Replace(Replace(strCri, "ED_DEPTNO", "BD_DEPT"), "ED_GLNO", "GL_NUMBER")

                If rsDATA.State <> 0 Then rsDATA.Close
                rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic
                If rsDATA.EOF = False And rsDATA.BOF = False Then
                    Do
                        If IsNull(rsDATA("TOTF")) = False Or IsNull(rsDATA("TOTFs")) = False Then
                            tmpFT# = nz(rsDATA("TOTF"), 0) + nz(rsDATA("TOTFs"), 0)
                            If tmpFT# > 0 Then
                                strSQL = "UPDATE HRMANWRK  SET SEQ" & C% & "='" & FNegs((TFT# / tmpFT#) * 100, "%") & "' "
                                strSQL = strSQL & "WHERE GL_NUMBER='Stat' AND BD_Row='TUF' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                            End If
                        End If
                        gdbAdoIhr001W.BeginTrans
                        gdbAdoIhr001W.Execute strSQL
                        gdbAdoIhr001W.CommitTrans
                        rsDATA.MoveNext
                    Loop Until rsDATA.EOF
                End If
                rsDATA.Close
                
                'Turnover TMP
                strSQL = "SELECT sum(ACTUAL_TMP_A)  as TOTT, sum(ACTUAL_TMP_S)  as TOTTs "
                strSQL = strSQL & "FROM HRBUDGET WHERE BUDGET_YEAR = " & txtBudgetYear.Text
                strSQL = strSQL & " AND MONTH_SEQ=" & C%
                'If Len(strCri) > 0 Then strSQL = strSQL & " AND " & strCri
                'Ticket #20626 Franks 07/14/2011 fix error "Invalid column 'ED_GLNO'"
                If Len(strCri) > 0 Then strSQL = strSQL & " AND " & Replace(Replace(strCri, "ED_DEPTNO", "BD_DEPT"), "ED_GLNO", "GL_NUMBER")
                
                If rsDATA.State <> 0 Then rsDATA.Close
                rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
                If rsDATA.EOF = False And rsDATA.BOF = False Then
                    Do

                        If IsNull(rsDATA("TOTT")) = False Or IsNull(rsDATA("TOTTs")) = False Then
                            tmpTMP# = nz(rsDATA("TOTT"), 0) + nz(rsDATA("TOTTs"), 0)
                            If tmpTMP# > 0 Then
                                strSQL = "UPDATE HRMANWRK  SET SEQ" & C% & "='" & FNegs((TTM# / tmpTMP) * 100, "%") & "' "
                                strSQL = strSQL & "WHERE GL_NUMBER='Stat' AND BD_Row='TUT' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                            End If
                        End If
                        gdbAdoIhr001W.BeginTrans
                        gdbAdoIhr001W.Execute strSQL
                        gdbAdoIhr001W.CommitTrans
                        rsDATA.MoveNext
                    Loop Until rsDATA.EOF
                End If
                rsDATA.Close
                If tmpFT + tmpTMP > 0 Then
                    strSQL = "UPDATE HRMANWRK SET SEQ" & C% & "='" & FNegs(((TFT# + TTM#) / (tmpFT + tmpTMP)) * 100, "%") & "' "
                    strSQL = strSQL & "WHERE GL_NUMBER='Stat' AND BD_Row='TUS' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                    gdbAdoIhr001W.BeginTrans
                    gdbAdoIhr001W.Execute strSQL
                    gdbAdoIhr001W.CommitTrans
                End If
            End If
        End If
        If X% + 1 > 12 Then
            X% = 1
        Else
            X% = X% + 1
        End If
    Next C%
    
    If Len(clpDept.Text) > 0 Then
        strCri = " BD_DEPT='" & clpDept.Text & "' "
    Else
        strCri = ""
    End If
    
    If Len(clpGL.Text) > 0 Then
        If Len(strCri) > 0 Then
            strCri = strCri & " AND GL_NUMBER IN ('" & getCodes(clpGL.Text) & "') "
        Else
            strCri = " GL_NUMBER IN ('" & getCodes(clpGL.Text) & "') "
        End If
    End If
    
            strSQL = "SELECT SUM(TOTAL_SALES) AS TOTS, SUM(TOTAL_MATERIAL_COST) AS TOTM, SUM(TOTAL_VALUE_ADDED) AS TOTV, SUM(VALUE_ADDED_ASSOC) AS VALA, MONTH_SEQ "
            strSQL = strSQL & "FROM HRBUDGET WHERE BUDGET_YEAR = " & txtBudgetYear.Text
            If Len(strCri) > 0 Then strSQL = strSQL & " AND " & strCri
            strSQL = strSQL & " GROUP BY MONTH_SEQ"
            rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic
            If rsDATA.EOF = False And rsDATA.BOF = False Then
                Do
                    strSQL = "UPDATE HRMANWRK  SET SEQ" & rsDATA("MONTH_SEQ") & "='" & rsDATA("TOTS") & "' "
                    strSQL = strSQL & "WHERE GL_NUMBER='Stat' AND BD_Row='TOTS' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                    gdbAdoIhr001W.BeginTrans
                    gdbAdoIhr001W.Execute strSQL
                    gdbAdoIhr001W.CommitTrans
                    
                    strSQL = "UPDATE HRMANWRK  SET SEQ" & rsDATA("MONTH_SEQ") & "='" & rsDATA("TOTM") & "' "
                    strSQL = strSQL & "WHERE GL_NUMBER='Stat' AND BD_Row='TOTM' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                    gdbAdoIhr001W.BeginTrans
                    gdbAdoIhr001W.Execute strSQL
                    gdbAdoIhr001W.CommitTrans
                    
                    strSQL = "UPDATE HRMANWRK  SET SEQ" & rsDATA("MONTH_SEQ") & "='" & rsDATA("TOTV") & "' "
                    strSQL = strSQL & "WHERE GL_NUMBER='Stat' AND BD_Row='TOTV' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                    gdbAdoIhr001W.BeginTrans
                    gdbAdoIhr001W.Execute strSQL
                    gdbAdoIhr001W.CommitTrans
                    
                    strSQL = "UPDATE HRMANWRK  SET SEQ" & rsDATA("MONTH_SEQ") & "='" & rsDATA("VALA") & "' "
                    strSQL = strSQL & "WHERE GL_NUMBER='Stat' AND BD_Row='VALA' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                    gdbAdoIhr001W.BeginTrans
                    gdbAdoIhr001W.Execute strSQL
                    gdbAdoIhr001W.CommitTrans
                    rsDATA.MoveNext
                Loop Until rsDATA.EOF
            End If
            rsDATA.Close
        
            'get Totals
            strSQL = "SELECT sum(ACTUAL_FT_A)  as TOTF, sum(ACTUAL_TMP_A)  as TOTT,  sum(ACTUAL_FT_S)  as TOTFs, sum(ACTUAL_TMP_S)  as TOTTs,  MONTH_SEQ "
            strSQL = strSQL & "FROM HRBUDGET WHERE BUDGET_YEAR = " & txtBudgetYear.Text
            If Len(strCri) > 0 Then strSQL = strSQL & " AND " & strCri
            strSQL = strSQL & " GROUP BY MONTH_SEQ"
            rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic
            If rsDATA.EOF = False And rsDATA.BOF = False Then
                Do
                    If IsNull(rsDATA("TOTF")) = False Or IsNull(rsDATA("TOTFs")) = False Then
                        TMP# = nz(rsDATA("TOTF"), 0) + nz(rsDATA("TOTFs"), 0)
                        strSQL = "UPDATE HRMANWRK  SET SEQ" & rsDATA("MONTH_SEQ") & "='" & TMP# & "' "
                        strSQL = strSQL & "WHERE GL_NUMBER='Stat' AND BD_Row='TOTF' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                        gdbAdoIhr001W.BeginTrans
                        gdbAdoIhr001W.Execute strSQL
                        gdbAdoIhr001W.CommitTrans
                    End If
                    If IsNull(rsDATA("TOTT")) = False Or IsNull(rsDATA("TOTTs")) = False Then
                        TMP# = nz(rsDATA("TOTT"), 0) + nz(rsDATA("TOTTs"), 0)
                        strSQL = "UPDATE HRMANWRK  SET SEQ" & rsDATA("MONTH_SEQ") & "='" & TMP# & "' "
                        strSQL = strSQL & "WHERE GL_NUMBER='Stat' AND BD_Row='TOTT' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                        gdbAdoIhr001W.BeginTrans
                        gdbAdoIhr001W.Execute strSQL
                     gdbAdoIhr001W.CommitTrans
                    End If
                    If IsNull(rsDATA("TOTF")) = False Or IsNull(rsDATA("TOTT")) = False Or IsNull(rsDATA("TOTFs")) = False Or IsNull(rsDATA("TOTTs")) = False Then
                        TMP# = nz(rsDATA("TOTF"), 0) + nz(rsDATA("TOTT"), 0) + nz(rsDATA("TOTFs"), 0) + nz(rsDATA("TOTTs"), 0)
                        strSQL = "UPDATE HRMANWRK  SET SEQ" & rsDATA("MONTH_SEQ") & "='" & TMP# & "' "
                        strSQL = strSQL & "WHERE GL_NUMBER='Stat' AND BD_Row='TOTA' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                        gdbAdoIhr001W.BeginTrans
                        gdbAdoIhr001W.Execute strSQL
                        gdbAdoIhr001W.CommitTrans
                    End If
                    rsDATA.MoveNext
                Loop Until rsDATA.EOF
            End If
            rsDATA.Close

    'Abesnteeism
    strSQL = "SELECT SUM(ABSENT_HOURS_FT) AS AFT, SUM(ABSENT_HOURS_TMP) AS ATM, SUM(SCHED_HOURS_FT) AS SFT, SUM(SCHED_HOURS_TMP) AS STM, MONTH_SEQ "
    strSQL = strSQL & "FROM HRBUDGET WHERE BUDGET_YEAR=" & txtBudgetYear.Text
    If Len(strCri) > 0 Then strSQL = strSQL & " AND " & strCri
    strSQL = strSQL & " GROUP BY MONTH_SEQ"
    rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If rsDATA.EOF = False And rsDATA.BOF = False Then
        Do
            If Not IsNull(rsDATA("SFT")) Then
                If rsDATA("SFT") > 0 Then
                    strSQL = "UPDATE HRMANWRK SET SEQ" & rsDATA("month_seq") & "='" & FNegs((rsDATA("AFT") / rsDATA("SFT")) * 100, "%") & "' "
                    strSQL = strSQL & "WHERE GL_NUMBER='Stat' AND BD_Row='ABF' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                    gdbAdoIhr001W.BeginTrans
                    gdbAdoIhr001W.Execute strSQL
                    gdbAdoIhr001W.CommitTrans
                End If
            End If
            If Not IsNull(rsDATA("STM")) Then
                If rsDATA("STM") > 0 Then
                    strSQL = "UPDATE HRMANWRK SET SEQ" & rsDATA("month_seq") & "='" & FNegs((rsDATA("ATM") / rsDATA("STM")) * 100, "%") & "' "
                    strSQL = strSQL & "WHERE GL_NUMBER='Stat' AND BD_Row='ABT' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                    gdbAdoIhr001W.BeginTrans
                    gdbAdoIhr001W.Execute strSQL
                    gdbAdoIhr001W.CommitTrans
                End If
            End If
            varTMP = Null
            '***************
'            If IsNull(rsDATA("AFT")) = False And IsNull(rsDATA("SFT")) = False Then
'                If rsDATA("SFT") > 0 Then
'                    varTMP = (nz(rsDATA("AFT"), 0) / nz(rsDATA("SFT"), 1))
'                End If
'            End If
'            If IsNull(rsDATA("ATM")) = False And IsNull(rsDATA("STM")) = False Then
'                If rsDATA("STM") > 0 Then
'                    varTMP = varTMP + (nz(rsDATA("ATM"), 0) / nz(rsDATA("STM"), 1))
'                End If
'            End If



            varTMP = (nz(rsDATA("AFT"), 0) + nz(rsDATA("ATM"), 0))
            If nz(rsDATA("STM"), 0) + nz(rsDATA("SFT"), 0) > 0 Then
                varTMP = varTMP / (nz(rsDATA("STM"), 0) + nz(rsDATA("SFT"), 0))
            End If

            If Not IsNull(varTMP) Then
                strSQL = "UPDATE HRMANWRK SET SEQ" & rsDATA("month_seq") & "='" & FNegs(varTMP * 100, "%") & "' "
                strSQL = strSQL & "WHERE GL_NUMBER='Stat' AND BD_Row='ABS' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
            
            rsDATA.MoveNext
        Loop Until rsDATA.EOF
    End If
    rsDATA.Close

    'Average count YTD
    strSQL = "SELECT seq1, seq2, seq3, seq4, seq5, seq6, seq7, seq8, seq9, seq10, seq11, seq12, BD_ROW "
    strSQL = strSQL & "FROM HRMANWRK WHERE BD_Row IN ('NHF','NHT','NHS','TEF','TET','TES','TUF','TUT','TUS','ABF','ABT','ABS', 'TOTS','TOTM','TOTV','TOTA','VALA') "
    strSQL = strSQL & "AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='"
    strSQL = strSQL & glbUserID & "' GROUP BY seq1, seq2, seq3, seq4, seq5, seq6, seq7, seq8, seq9, seq10, seq11, seq12, BD_ROW "
    Dim total#, CNT#
    rsDATA.Open strSQL, gdbAdoIhr001W, adOpenStatic, adLockOptimistic
    If rsDATA.EOF = False And rsDATA.BOF = False Then
        Do
            total# = 0
            CNT# = 0
            If Len(rsDATA("seq1")) > 0 Then
                total# = total# + CDbl(rsDATA("seq1"))
                CNT# = CNT# + 1
            End If
            If Len(rsDATA("seq2")) > 0 Then
                total# = total# + CDbl(rsDATA("seq2"))
                CNT# = CNT# + 1
            End If
            If Len(rsDATA("seq3")) > 0 Then
                total# = total# + CDbl(rsDATA("seq3"))
                CNT# = CNT# + 1
            End If
            If Len(rsDATA("seq4")) > 0 Then
                total# = total# + CDbl(rsDATA("seq4"))
                CNT# = CNT# + 1
            End If
            If Len(rsDATA("seq5")) > 0 Then
                total# = total# + CDbl(rsDATA("seq5"))
                CNT# = CNT# + 1
            End If
            If Len(rsDATA("seq6")) > 0 Then
                total# = total# + CDbl(rsDATA("seq6"))
                CNT# = CNT# + 1
            End If
            If Len(rsDATA("seq7")) > 0 Then
                total# = total# + CDbl(rsDATA("seq7"))
                CNT# = CNT# + 1
            End If
            If Len(rsDATA("seq8")) > 0 Then
                total# = total# + CDbl(rsDATA("seq8"))
                CNT# = CNT# + 1
            End If
            If Len(rsDATA("seq9")) > 0 Then
                total# = total# + CDbl(rsDATA("seq9"))
                CNT# = CNT# + 1
            End If
            If Len(rsDATA("seq10")) > 0 Then
                total# = total# + CDbl(rsDATA("seq10"))
                CNT# = CNT# + 1
            End If
            If Len(rsDATA("seq11")) > 0 Then
                total# = total# + CDbl(rsDATA("seq11"))
                CNT# = CNT# + 1
            End If
            If Len(rsDATA("seq12")) > 0 Then
                total# = total# + CDbl(rsDATA("seq12"))
                CNT# = CNT# + 1
            End If
            If CNT# > 0 Then
                total# = total# / CNT#
                strSQL = "UPDATE HRMANWRK SET AvgVal='" & FNegs(total#, "0") & "' WHERE BD_ROW='" & rsDATA("BD_ROW") & "' AND BD_YEAR=" & txtBudgetYear.Text & " AND WRKEMP='" & glbUserID & "'"
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute strSQL
                gdbAdoIhr001W.CommitTrans
            End If
            rsDATA.MoveNext
        Loop Until rsDATA.EOF
    End If
    
exH:
    Set rsDATA = Nothing
    Exit Sub
Eh:
    
    'Debug.Print Err.Description
    'Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err

    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "selAttWrk3", "Manpower Report", "Select")
    Resume exH
End Sub

Private Function Cri_Sorts()
    Dim X%, C%
    Dim MonthLabels(12) As String
    Dim DTE As String
    Dim strSQL As String
    Dim rsDATA As New ADODB.Recordset
    
    X% = 0
    
    strSQL = "SELECT DISTINCT BUDGET_MONTH, MONTH_SEQ FROM HRBUDGET WHERE BUDGET_YEAR=" & txtBudgetYear.Text
    strSQL = strSQL & " ORDER BY MONTH_SEQ ASC"
    rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If rsDATA.EOF = False And rsDATA.BOF = False Then
        X% = rsDATA("BUDGET_MONTH")
        C% = rsDATA("month_seq")
    End If
    rsDATA.Close
    If X% > 0 Then
        DTE = Format("01/" & MonthName(X%, True) & "/" & txtBudgetYear.Text, "MMMM dd, yyyy")
        If X% > 1 And C% = 1 Then 'If the first month isn't jan then the last won't be in this year
            DTE = DTE & " - " & Format(getEOM(X% - 1) & "/" & MonthName((X% - 1), True) & "/" & (txtBudgetYear.Text + 1), "MMMM dd, yyyy")
        ElseIf X% > 1 And C% < X% Then
            DTE = DTE & " - " & Format(getEOM(X% - C% + 1) & "/" & MonthName(X% - C% + 1), True & "/" & (txtBudgetYear.Text + 1), "MMMM dd, yyyy")
        Else
            DTE = DTE & " - " & Format(getEOM(12) & "/DEC/" & (txtBudgetYear.Text), "MMMM dd, yyyy")
        End If
        Me.vbxCrystal.Formulas(52) = "FTDates='" & DTE & "'"
        
        For C% = 1 To 12
            MonthLabels(C%) = MonthName(X%, True)
            If C% <= X% Then
                MonthLabels(C%) = MonthLabels(C%) & "    " & txtBudgetYear.Text
            Else
                MonthLabels(C%) = MonthLabels(C%) & "    " & (txtBudgetYear.Text + 1)
            End If
            If X% + 1 <= 12 Then
                X% = X% + 1
            Else
                X% = 1
            End If
        Next C%
    Else
        For C% = 1 To 12
            MonthLabels(C%) = MonthName(C%, True) & "    " & txtBudgetYear.Text
        Next C%
    End If
    
    Me.vbxCrystal.Formulas(53) = "lblDepartment='" & lStr("G/L") & "'"
    Me.vbxCrystal.Formulas(54) = "lblMonth1='" & MonthLabels(1) & "'"
    Me.vbxCrystal.Formulas(55) = "lblMonth10='" & MonthLabels(10) & "'"
    Me.vbxCrystal.Formulas(56) = "lblMonth11='" & MonthLabels(11) & "'"
    Me.vbxCrystal.Formulas(57) = "lblMonth12='" & MonthLabels(12) & "'"
    Me.vbxCrystal.Formulas(58) = "lblMonth2='" & MonthLabels(2) & "'"
    Me.vbxCrystal.Formulas(59) = "lblMonth3='" & MonthLabels(3) & "'"
    Me.vbxCrystal.Formulas(60) = "lblMonth4='" & MonthLabels(4) & "'"
    Me.vbxCrystal.Formulas(61) = "lblMonth5='" & MonthLabels(5) & "'"
    Me.vbxCrystal.Formulas(62) = "lblMonth6='" & MonthLabels(6) & "'"
    Me.vbxCrystal.Formulas(63) = "lblMonth7='" & MonthLabels(7) & "'"
    Me.vbxCrystal.Formulas(64) = "lblMonth8='" & MonthLabels(8) & "'"
    Me.vbxCrystal.Formulas(65) = "lblMonth9='" & MonthLabels(9) & "'"
    Me.vbxCrystal.Formulas(66) = "fldBudgetYear='Budget Year: " & txtBudgetYear.Text & "'"
    
End Function
Public Function getEOM(Mnt As Variant) As Integer
   Dim myDate As Date
   Dim myMonth As String
   
   Dim NextMonth As Date, EndOfMonth As Date
   If IsNumeric(Mnt) Then
        myMonth = MonthName(Mnt, True)
    Else
        myMonth = Mnt
    End If
   
   myDate = Format("1/" & myMonth & "/2005", "dd/mmm/yyyy")
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

Private Function FNegs(num As Variant, Optional frmAT As String) As String
    Dim retVal As String
    
    If frmAT = "%" Then
        If num < 0 Then
            retVal = Format(Abs(num), "(##0.00)")
        Else
            retVal = Format(num, "##0.00")
        End If
    ElseIf frmAT = "0" Then
        If num < 0 Then
            retVal = Format(Abs(num), "(##0.0)")
        Else
            retVal = Format(num, "##0.0")
        End If
    Else
        If num < 0 Then
            retVal = Format(Abs(num), "(##0)")
        Else
            retVal = Format(num, "##0")
        End If
    End If
        
    FNegs = retVal
        
End Function

