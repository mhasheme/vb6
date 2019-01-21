VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRDailyMan 
   Caption         =   "Daily Manpower Update"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   9210
   WindowState     =   2  'Maximized
   Begin INFOHR_Controls.CodeLookup clpDiv 
      Height          =   285
      Left            =   1635
      TabIndex        =   1
      Top             =   600
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   1635
      TabIndex        =   0
      Tag             =   "00-Administered By"
      Top             =   225
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1635
      TabIndex        =   2
      Top             =   960
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   1635
      TabIndex        =   3
      Tag             =   "00-Location - Code"
      Top             =   1320
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   6480
      Top             =   720
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
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Label lblDiv 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   645
      Width           =   555
   End
   Begin VB.Label lblLocation 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1365
      Width           =   615
   End
   Begin VB.Label lblAdmin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Administered By"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   270
      Width           =   1125
   End
   Begin VB.Label lblDept 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1005
      Width           =   990
   End
End
Attribute VB_Name = "frmRDailyMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim FTA, FTS, TMA, TMS
'*******************************************************
'*                                                     *
'*      Form: frmRManpower                             *
'*                                                     *
'*           Created: 09/Sep/05    By: Bryan           *
'*           Modified:             By:                 *
'*                                                     *
'*           Comments: To report budgeted Manpower data*
'*                                                     *
'*******************************************************




Private Sub clpCode_GotFocus(Index As Integer)
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub clpDept_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub



Private Sub clpDiv_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub Form_Activate()
    Call SET_UP_MODE
End Sub

Private Sub Form_Load()
glbOnTop = "FRMRDAILYMAN"
    MDIMain.lstPanel.Visible = False
    MDIMain.lstView.Visible = False
    
    Screen.MousePointer = HOURGLASS
    
    Call setRptCaption(Me)
    If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6
    Call INI_Controls(Me)
    Screen.MousePointer = DEFAULT
    
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
Dim x%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm("Manpower Plan", Me) Then Exit Sub
    Call set_PrintState(False)
    x% = Cri_SetAll()
    Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    Call set_PrintState(True)
End If
Exit Sub

PrntErr:
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString, , "info:HR"
Resume Next

End Sub



Public Sub cmdView_Click()
Dim x%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    Call set_PrintState(False)
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    Screen.MousePointer = HOURGLASS
    x% = Cri_SetAll()
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
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString, , "info:HR"

Resume Next

End Sub

Private Function CriCheck()

    CriCheck = False
    
    If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
        MsgBox lStr("If Department Entered - it must be known"), , "info:HR"
         clpDept.SetFocus
        Exit Function
    End If
    
    If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
        MsgBox lStr("If Department Entered - it must be known"), , "info:HR"
         clpDiv.SetFocus
        Exit Function
    End If
    
    If Len(clpCode(0).Text) > 0 And clpCode(0).Caption = "Unassigned" Then
        MsgBox lStr("If Department Entered - it must be known"), , "info:HR"
         clpCode(0).SetFocus
        Exit Function
    End If
    
    If Len(clpCode(1).Text) > 0 And clpCode(1).Caption = "Unassigned" Then
        MsgBox lStr("If Department Entered - it must be known"), , "info:HR"
         clpCode(1).SetFocus
        Exit Function
    End If
    
    CriCheck = True
End Function

Private Function Cri_SetAll()
Dim x%, strRName$, xNoFiles, dscGroup$
Dim dtYYY#, dtMM#, dtDD#


Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

' call cri models set both glbiONeWhere and strSelCri
If CriCheck = False Then Exit Function
Call Cri_Div
Call Cri_Dept
Call Cri_AdminBY
Call Cri_Loc
Call SELATTWRK
Call Cri_Sorts

If Len(glbstrSelCri) > 0 Then
    glbstrSelCri = glbstrSelCri & " AND {HRDMANWRK.WRKEMP}='" & glbUserID & "'"
Else
    glbstrSelCri = "{HRDMANWRK.WRKEMP}='" & glbUserID & "'"
End If
Me.vbxCrystal.SelectionFormula = glbstrSelCri
strRName$ = glbIHRREPORTS & "rzDailyman.rpt"
Me.vbxCrystal.ReportFileName = strRName$

If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    Me.vbxCrystal.DataFiles(0) = glbIHRDB
    Me.vbxCrystal.DataFiles(1) = glbIHRDB
    Me.vbxCrystal.DataFiles(2) = glbIHRDBW
    Me.vbxCrystal.DataFiles(3) = glbIHRDB
    Me.vbxCrystal.DataFiles(4) = glbIHRDB
    Me.vbxCrystal.DataFiles(5) = glbIHRDB
End If

' window title if appropriate
Me.vbxCrystal.WindowTitle = "Daily Manpower Update"

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

Private Sub Cri_Dept()
    Dim DivCri As String
    
    If Len(clpDept.Text) > 0 Then
        DivCri = "({HRDMANWRK.BD_DEPT} = '" & clpDept.Text & "')"
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
Private Sub Cri_AdminBY()
    Dim DivCri As String
    
    If Len(clpCode(0).Text) > 0 Then
        DivCri = "({HRDMANWRK.BD_ADMINBY} = '" & clpCode(0).Text & "')"
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
Private Sub Cri_Div()
    Dim DivCri As String
    
    If Len(clpDiv.Text) > 0 Then
        DivCri = "({HRDMANWRK.BD_DIV} = '" & clpDiv.Text & "')"
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

Private Sub Cri_Loc()
    Dim DivCri As String
    
    If Len(clpCode(1).Text) > 0 Then
        DivCri = "({HRDMANWRK.BD_LOC} = '" & clpCode(1).Text & "')"
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

Private Sub SELATTWRK()
    Dim strSQL As String, SQLQ As String
    Dim strJobStatus As String
    Dim c%, x%
    Dim rsDATA As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim DTE As Long
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(2).Caption = ""
    Screen.MousePointer = HOURGLASS
    
    'get budget year
    SQLQ = ""
    If Not glbOracle Then
        SQLQ = "top 1"
    End If
    strSQL = "SELECT " & SQLQ & " Month_SEQ, Budget_Month FROM HRBUDGET WHERE BUDGET_YEAR=" & Year(Date)
    If glbOracle Then
        strSQL = strSQL & " AND ROWNUM = 1"
    End If
    strSQL = strSQL & " ORDER BY Month_Seq ASC"
    rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If rsDATA.EOF = False And rsDATA.BOF = False Then
        If rsDATA("month_seq") = rsDATA("Budget_Month") Then 'if the budget year is from jan to dec
            DTE = Year(Date)
        ElseIf rsDATA("month_seq") < rsDATA("Budget_Month") Then
            'the current date might be this year, last budget year
            If month(Date) < rsDATA("Budget_Month") Then
                DTE = Year(Date) - 1
            Else
                DTE = Year(Date)
            End If
        End If
    Else
        DTE = Year(Date)
    End If
    rsDATA.Close
    

        strSQL = "SELECT MONTH_SEQ FROM HRBUDGET WHERE BUDGET_MONTH=" & month(Date)
        strSQL = strSQL & " AND BUDGET_YEAR=" & DTE
        rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
        If rsDATA.EOF = False And rsDATA.BOF = False Then
            x% = rsDATA("MONTH_SEQ")
            c% = month(Date)
        End If
        rsDATA.Close
            
        If x% > 0 Then
            'Refresh working table
            strSQL = "DELETE FROM HRDMANWRK WHERE WRKEMP='" & glbUserID & "'"
            gdbAdoIhr001W.BeginTrans
            gdbAdoIhr001W.Execute strSQL
            gdbAdoIhr001W.CommitTrans
        
            'Insert Budget data into working table
            strSQL = "INSERT INTO HRDMANWRK(BFTA, BTA, BSup, BD_DIV, BD_ADMINBY, BD_ADMINBY_TABL, BD_DEPT,  WRKEMP) "
            strSQL = strSQL & "SELECT SUM(BUDGET_FT_A) as BFTA, SUM(BUDGET_TMP_A) AS BTMA, SUM(BUDGET_FT_S) AS BSup, "
            strSQL = strSQL & "BD_DIV,  BD_ADMINBY, BD_ADMINBY_TABL, BD_DEPT,  '" & glbUserID & "' as WrkID "
            strSQL = strSQL & "From HRBUDGET "
            strSQL = strSQL & "Where (BUDGET_YEAR = " & DTE & ") And (BUDGET_MONTH = " & month(Date) & ")"
            strSQL = strSQL & " GROUP BY BD_DIV, BD_ADMINBY, BD_ADMINBY_TABL, BD_DEPT"

            gdbAdoIhr001W.BeginTrans
            gdbAdoIhr001W.Execute strSQL
            gdbAdoIhr001W.CommitTrans
       
            'Insert Actual data into Working table
                strSQL = "SELECT * FROM HRDMANWRK WHERE WRKEMP='" & glbUserID & "'"
                rsTemp.Open strSQL, gdbAdoIhr001, adOpenDynamic, adLockOptimistic, adCmdText
                If rsTemp.EOF = False And rsTemp.EOF = False Then
                    Do
                        FTA = Null: FTS = Null: TMA = Null: TMS = Null
                        strSQL = "SELECT Count(HREMP.ED_EMPNBR) AS EMPCNT, HREMP.ED_PT, HRJOB.JB_STATUS "
                        If glbOracle Then
                            strSQL = strSQL & "From HREMP, HR_JOB_HISTORY, HRJOB "
                            strSQL = strSQL & "Where HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE AND (HREMP.ED_DOH <= " & Date_SQL(DTE) & ") AND (HR_JOB_HISTORY.JH_CURRENT<>0)"
                        Else
                            strSQL = strSQL & "From  (HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
                            strSQL = strSQL & "Where (HREMP.ED_DOH <= " & Date_SQL(Date) & ")  AND (HR_JOB_HISTORY.JH_CURRENT<>0) "
                        End If
                        If Len(rsTemp("BD_DIV")) > 0 Then
                            strSQL = strSQL & "AND ED_DIV='" & rsTemp("BD_DIV") & "' "
                        End If
'                        If Len(RsTemp("BD_LOCATION")) > 0 Then
'                            strSQL = strSQL & "AND ED_GLNO='" & RsTemp("ED_LOC") & "' "
'                        End If
                        If Len(rsTemp("BD_ADMINBY")) > 0 Then
                            strSQL = strSQL & "AND ED_ADMINBY='" & rsTemp("BD_ADMINBY") & "' "
                        End If
                        If Len(rsTemp("BD_DEPT")) > 0 Then
                            strSQL = strSQL & "AND ED_DEPTNO='" & rsTemp("BD_DEPT") & "' "
                        End If
                        
                        strSQL = strSQL & "GROUP BY  HREMP.ED_PT, HRJOB.JB_STATUS "
                        strSQL = strSQL & "HAVING HRJOB.JB_STATUS <> 'NA' "
                        If glbCompSerial = "S/N - 2369W" Then
                            strSQL = strSQL & " AND (HREMP.ED_PT='FT' Or HREMP.ED_PT='TMP') "
                        End If
                        
                        rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
                        If rsDATA.EOF = False And rsDATA.BOF = False Then
                            Do
                                If glbCompSerial = "S/N - 2369W" And rsDATA("JB_STATUS") = "L" Then
                                    If rsDATA("ED_PT") = "FT" Then
                                        If IsNull(FTS) Then FTS = 0
                                        FTS = FTS + rsDATA("EMPCNT")
                                    Else
                                        If IsNull(TMS) Then TMS = 0
                                        TMS = TMS + rsDATA("EMPCNT")
                                    End If
                                 Else
                                    If rsDATA("ED_PT") = "FT" Then
                                        If IsNull(FTA) Then FTA = 0
                                        FTA = FTA + rsDATA("EMPCNT")
                                    Else
                                        If IsNull(TMA) Then TMA = 0
                                        TMA = TMA + rsDATA("EMPCNT")
                                    End If
                                End If
                                rsDATA.MoveNext
                            Loop Until rsDATA.EOF
                        End If
                        rsDATA.Close
                    
   
            'Changed by Bryan 07/Mar/2006 Ticket#10493
            'Will select terminated employees terminated on the last day of the month.
            '**********************
                Dim TermDTE As Date
                        If c% < x% Then
                            TermDTE = getEOM(c%) & "/" & MonthName(c%, True) & "/" & (DTE + 1)
                        Else
                            TermDTE = getEOM(c%) & "/" & MonthName(c%, True) & "/" & DTE
                        End If
                        If CDate(TermDTE) < Date Then
                            strSQL = "DELETE FROM HRMANTERM WHERE WRKEMP='" & glbUserID & "'"
                            gdbAdoIhr001W.BeginTrans
                            gdbAdoIhr001W.Execute strSQL
                            gdbAdoIhr001W.CommitTrans
                            'Find Actual Employees from the terminated for this month(c)
                            strSQL = "SELECT TERM_HREMP.ED_EMPNBR, TERM_HREMP.ED_DEPTNO, TERM_HREMP.ED_GLNO, TERM_HREMP.ED_PT, TERM_HREMP.ED_ADMINBY, Term_JOB_HISTORY.JH_JOB, TERM_HREMP.ED_LOC, TERM_HREMP.ED_DIV  "
                            If glbOracle Then
                                strSQL = strSQL & "From  TERM_HREMP, TERM_HRTRMEMP, Term_JOB_HISTORY "
                                strSQL = strSQL & "Where TERM_HREMP.ED_EMPNBR = TERM_HRTRMEMP.TERM_SEQ AND TERM_HREMP.TERM_SEQ = Term_JOB_HISTORY.JH_EMPNBR AND (TERM_HREMP.ED_DOH <= " & Date_SQL(TermDTE) & ") AND (TERM_HRTRMEMP.TERM_DOT = " & Date_SQL(TermDTE) & ") "
                            Else
                                strSQL = strSQL & "From  ((TERM_HREMP INNER JOIN TERM_HRTRMEMP ON TERM_HREMP.TERM_SEQ = TERM_HRTRMEMP.TERM_SEQ) INNER JOIN Term_JOB_HISTORY ON TERM_HREMP.TERM_SEQ = Term_JOB_HISTORY.TERM_SEQ)"
                                strSQL = strSQL & "Where (TERM_HREMP.ED_DOH <= " & Date_SQL(TermDTE) & ") AND (TERM_HRTRMEMP.TERM_DOT = " & Date_SQL(TermDTE) & ") "
                            End If
                            strSQL = strSQL & "AND Term_JOB_HISTORY.JH_CURRENT<>0"
                            If Len(rsTemp("BD_DEPT")) > 0 Then
                                strSQL = strSQL & "AND ED_DEPTNO='" & rsTemp("BD_DEPT") & "' "
                            End If
                            If Len(rsTemp("BD_ADMINBY")) > 0 Then
                                strSQL = strSQL & "AND ED_ADMINBY='" & rsTemp("BD_ADMINBY") & "' "
                            End If
                            If Len(rsTemp("BD_DIV")) > 0 Then
                                strSQL = strSQL & "AND ED_DIV='" & rsTemp("BD_DIV") & "' "
                            End If
                            If glbCompSerial = "S/N - 2369W" Then
                                strSQL = strSQL & "AND (TERM_HREMP.ED_PT='FT' Or TERM_HREMP.ED_PT='TMP') "
                            End If
                            
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
                            strSQL = "SELECT Count(HRMANTERM.ED_EMPNBR) AS EMPCNT, HRMANTERM.ED_PT,  HRJOB.JB_STATUS "
                            If glbOracle Then
                                strSQL = strSQL & "From  HRMANTERM, hrjob "
                                strSQL = strSQL & "WHERE HRMANTERM.JH_JOB = hrjob.JB_CODE AND HRMANTERM.WRKEMP = '" & glbUserID & "'"
                            Else
                                strSQL = strSQL & "From  HRMANTERM INNER JOIN hrjob ON HRMANTERM.JH_JOB = hrjob.JB_CODE WHERE HRMANTERM.WRKEMP = '" & glbUserID & "'"
                            End If
                            strSQL = strSQL & "GROUP BY HRMANTERM.ED_PT,  HRJOB.JB_STATUS "
                            strSQL = strSQL & "HAVING (hrjob.JB_STATUS<>'NA')"
                            rsDATA.Open strSQL, gdbAdoIhr001W, adOpenStatic, adLockOptimistic, adCmdText
                            If rsDATA.EOF = False And rsDATA.BOF = False Then
                                Do
                                    If glbCompSerial = "S/N - 2369W" And rsDATA("JB_STATUS") = "L" Then
                                        If rsDATA("ED_PT") = "FT" Then
                                            If IsNull(FTS) Then FTS = 0
                                            FTS = FTS + rsDATA("EMPCNT")
                                        Else
                                            If IsNull(TMS) Then TMS = 0
                                            TMS = TMS + rsDATA("EMPCNT")
                                        End If
                                     Else
                                        If rsDATA("ED_PT") = "FT" Then
                                            If IsNull(FTA) Then FTA = 0
                                            FTA = FTA + rsDATA("EMPCNT")
                                        Else
                                            If IsNull(TMA) Then TMA = 0
                                            TMA = TMA + rsDATA("EMPCNT")
                                        End If
                                    End If
                                    rsDATA.MoveNext
                                Loop Until rsDATA.EOF
                            End If
                            rsDATA.Close
                        End If
                        rsTemp("AFTA") = FTA
                        rsTemp("ASup") = FTS
                        rsTemp("ATA") = TMA
                        

                        rsTemp.Update
                        rsTemp.MoveNext
                    Loop Until rsTemp.EOF
                End If
                rsTemp.Close
        End If
  
    
    MDIMain.panHelp(0).FloodPercent = 75
    'Get Leave of Absence
    strSQL = "SELECT Count(HREMP.ED_EMPNBR) AS EMPCNT, HREMP.ED_DIV, HREMP.ED_DEPTNO, HREMP.ED_ADMINBY, HREMP.ED_LOC "
    If glbOracle Then
        strSQL = strSQL & "From HREMP, HR_JOB_HISTORY, HRJOB "
        strSQL = strSQL & "Where HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE AND (HREMP.ED_DOH <= " & Date_SQL(Date) & ") "
        SQLQ = "(SELECT HREMP.ED_EMPNBR, HRTABL.TB_USR3 FROM HREMP, HRTABL WHERE HREMP.ED_EMP = HRTABL.TB_KEY AND (HRTABL.TB_NAME = 'EDEM') AND (HRTABL.TB_USR3 = 1)) "
    Else
        strSQL = strSQL & "From  (HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
        strSQL = strSQL & "Where (HREMP.ED_DOH <= " & Date_SQL(Date) & ") "
        strSQL = strSQL & "AND HREMP.ED_EMPNBR IN (SELECT HREMP.ED_EMPNBR FROM HREMP INNER JOIN HRTABL ON HREMP.ED_EMP = HRTABL.TB_KEY WHERE (HRTABL.TB_NAME = 'EDEM') AND (HRTABL.TB_USR3 = 1)) "
    End If

    strSQL = strSQL & "AND (HRJOB.JB_STATUS<>'NA') AND HREMP.ED_PT='FT' "
    strSQL = strSQL & "GROUP BY HREMP.ED_DEPTNO, HREMP.ED_PT, HREMP.ED_ADMINBY, HREMP.ED_LOC, HREMP.ED_DIV "
    
    
    rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If rsDATA.EOF = False And rsDATA.BOF = False Then
        Do
            strSQL = "UPDATE HRDMANWRK SET ABS = " & CInt(rsDATA("EMPCNT"))
            strSQL = strSQL & " WHERE BD_DEPT='" & rsDATA("ED_DEPTNO") & "' AND BD_ADMINBY='" & rsDATA("ED_ADMINBY") & "' AND BD_LOC"
            If Not IsNull(rsDATA("ED_LOC")) Then
                strSQL = strSQL & "='" & rsDATA("ED_LOC") & "'"
            Else
                strSQL = strSQL & " IS NULL"
            End If
            gdbAdoIhr001W.BeginTrans
            gdbAdoIhr001W.Execute strSQL
            gdbAdoIhr001W.CommitTrans
            rsDATA.MoveNext
        Loop Until rsDATA.EOF
    End If
    rsDATA.Close
    MDIMain.panHelp(0).FloodPercent = 80
    'Get Vacation
    strSQL = "SELECT Count(HREMP.ED_EMPNBR) AS EMPCNT, HREMP.ED_DEPTNO, HREMP.ED_ADMINBY, HRJOB.JB_STATUS, HREMP.ED_LOC "
    If glbOracle Then
        strSQL = strSQL & "From HREMP, HR_JOB_HISTORY, HRJOB "
        strSQL = strSQL & "Where HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE AND (HREMP.ED_DOH <= " & Date_SQL(Date) & ") "
    Else
        strSQL = strSQL & "From  (HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
        strSQL = strSQL & "Where (HREMP.ED_DOH <= " & Date_SQL(Date) & ") "
    End If
    strSQL = strSQL & "AND HREMP.ED_EMPNBR IN (SELECT DISTINCT AD_EMPNBR FROM HR_ATTENDANCE WHERE AD_REASON='VAC' AND AD_DOA=" & Date_SQL(Date) & ") "
    strSQL = strSQL & "AND HREMP.ED_PT='FT' AND (HRJOB.JB_STATUS<>'NA')"
    strSQL = strSQL & "GROUP BY HREMP.ED_DEPTNO, HREMP.ED_ADMINBY, HRJOB.JB_STATUS, HREMP.ED_LOC "
    
    
    rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If rsDATA.EOF = False And rsDATA.BOF = False Then
        Do
            strSQL = "UPDATE HRDMANWRK SET VAC = " & CInt(rsDATA("EMPCNT"))
            strSQL = strSQL & " WHERE BD_DEPT='" & rsDATA("ED_DEPTNO") & "' AND BD_ADMINBY='" & rsDATA("ED_ADMINBY") & "' AND BD_LOC"
            If Not IsNull(rsDATA("ED_LOC")) Then
                strSQL = strSQL & "='" & rsDATA("ED_LOC") & "'"
            Else
                strSQL = strSQL & " IS NULL"
            End If

            gdbAdoIhr001W.BeginTrans
            gdbAdoIhr001W.Execute strSQL
            gdbAdoIhr001W.CommitTrans
            rsDATA.MoveNext
        Loop Until rsDATA.EOF
    End If
    rsDATA.Close
    MDIMain.panHelp(0).FloodPercent = 90
    
    'Get Absent.
    strSQL = "SELECT Count(HREMP.ED_EMPNBR) AS EMPCNT, HREMP.ED_DEPTNO, HREMP.ED_ADMINBY, HRJOB.JB_STATUS, HREMP.ED_LOC, HREMP.ED_PT "
    If glbOracle Then
        strSQL = strSQL & "From HREMP, HR_JOB_HISTORY, HRJOB "
        strSQL = strSQL & "Where HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE AND (HREMP.ED_DOH <= " & Date_SQL(Date) & ") "
    Else
        strSQL = strSQL & "From  (HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
        strSQL = strSQL & "Where (HREMP.ED_DOH <= " & Date_SQL(Date) & ") "
    End If
    strSQL = strSQL & "AND HREMP.ED_EMPNBR IN (SELECT DISTINCT AD_EMPNBR FROM HR_ATTENDANCE WHERE AD_REASON<>'VAC' AND AD_DOA=" & Date_SQL(Date) & ") "
    strSQL = strSQL & "AND (HRJOB.JB_STATUS<>'NA')"
    strSQL = strSQL & "GROUP BY HREMP.ED_DEPTNO, HREMP.ED_ADMINBY, HRJOB.JB_STATUS, HREMP.ED_LOC, HREMP.ED_PT "
    strSQL = strSQL & "HAVING (HREMP.ED_PT='FT' Or HREMP.ED_PT='TMP' Or HREMP.ED_PT='OT') "
    
    rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If rsDATA.EOF = False And rsDATA.BOF = False Then
        Do
            If rsDATA("ED_PT") = "FT" Then
                strSQL = "UPDATE HRDMANWRK SET ABSF = " & CInt(rsDATA("EMPCNT"))
                strSQL = strSQL & " WHERE BD_DEPT='" & rsDATA("ED_DEPTNO") & "' AND BD_ADMINBY='" & rsDATA("ED_ADMINBY") & "' AND BD_LOC"
                If Not IsNull(rsDATA("ED_LOC")) Then
                    strSQL = strSQL & "='" & rsDATA("ED_LOC") & "'"
                Else
                    strSQL = strSQL & " IS NULL"
                End If

            ElseIf rsDATA("ED_PT") = "TMP" Then
                strSQL = "UPDATE HRDMANWRK SET ABST = " & CInt(rsDATA("EMPCNT"))
                strSQL = strSQL & " WHERE BD_DEPT='" & rsDATA("ED_DEPTNO") & "' AND BD_ADMINBY='" & rsDATA("ED_ADMINBY") & "' AND BD_LOC"
                If Not IsNull(rsDATA("ED_LOC")) Then
                    strSQL = strSQL & "='" & rsDATA("ED_LOC") & "'"
                Else
                    strSQL = strSQL & " IS NULL"
                End If
            End If
            gdbAdoIhr001W.BeginTrans
            gdbAdoIhr001W.Execute strSQL
            gdbAdoIhr001W.CommitTrans
            rsDATA.MoveNext
        Loop Until rsDATA.EOF
    End If
    rsDATA.Close
    MDIMain.panHelp(0).FloodPercent = 100
            
exH:
    Set rsDATA = Nothing
    Screen.MousePointer = DEFAULT
    MDIMain.panHelp(0).FloodPercent = 0
    Exit Sub
EH:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "selAttWrk", "Manpower Report", "Select")
    Resume exH

End Sub


Private Function Cri_Sorts()

'Set up Labels
If glbCompSerial = "S/N - 2369W" Then 'TS Tech
    Me.vbxCrystal.Formulas(12) = "lblSuper='ATL/TL/COOR/MGR.'"
Else
    Me.vbxCrystal.Formulas(12) = "lblSuper=''"
End If
Me.vbxCrystal.Formulas(11) = "lblDept='" & lStr("Department") & "'"
Me.vbxCrystal.Formulas(10) = "lblAdmin='" & lStr("Administered By") & "'"
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
