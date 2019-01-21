VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRExcelRpt 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Health & Safety report"
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
   Begin VB.Frame fraCBAmount 
      Caption         =   "Bonus"
      Height          =   1095
      Left            =   6240
      TabIndex        =   25
      Top             =   240
      Width           =   3495
      Begin VB.TextBox txtAmount 
         Height          =   285
         Left            =   1200
         TabIndex        =   26
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Amount ($):"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   975
      End
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
      TabIndex        =   18
      Top             =   3720
      Width           =   6075
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   1
         Left            =   3750
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   19
         Top             =   0
         Width           =   1095
      End
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   3
      Tag             =   "00-Enter Status Code"
      Top             =   1650
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDEM"
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Tag             =   "EDPT-Category"
      Top             =   1980
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDPT"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   2
      Tag             =   "00-Enter Location Code"
      Top             =   990
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   660
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   330
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   1800
      TabIndex        =   7
      Tag             =   "00-Enter Administered By Code"
      Top             =   2970
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   1800
      TabIndex        =   8
      Tag             =   "00-Enter Section Code"
      Top             =   3300
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   6
      Tag             =   "00-Enter Region Code"
      Top             =   2640
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Tag             =   "10-Enter Employee Number"
      Top             =   2310
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   6835
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   23
      Tag             =   "00-Enter G/L Code"
      Top             =   1320
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   10
      LookupType      =   3
   End
   Begin VB.Label lblGL 
      BackStyle       =   0  'Transparent
      Caption         =   "G/L"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblPT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   2010
      Width           =   630
   End
   Begin VB.Label lblEENum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   2340
      Width           =   1290
   End
   Begin VB.Label lblSection 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   3300
      Width           =   540
   End
   Begin VB.Label lblRegion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   2670
      Width           =   510
   End
   Begin VB.Label lblAdmin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Administered By"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   1125
   End
   Begin VB.Label lblLocation 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   1020
      Width           =   615
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   450
   End
   Begin VB.Label lblDept 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   690
      Width           =   825
   End
   Begin VB.Label lblDiv 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   555
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
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmRExcelRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim HREMP As String
Dim strRpt As String
Dim rsRPT As New ADODB.Recordset
Dim fglbFileName
Dim fglbDateTable
Dim fglbDateField
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long


Private Sub cmdClose_Click()
Unload Me
End Sub



Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm(frmRExcelRpt.Caption & " Criteria", Me) Then Exit Sub
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
    MDIMain.Timer1.Enabled = True
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
    Call set_PrintState(True)
End If
Exit Sub

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", HREMP, "SELECT")
Resume Next

End Sub





Private Sub Cri_Code(intIdx%)
Dim CodeCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

If Len(clpCode(intIdx%)) > 0 Then
    Select Case intIdx%
    Case 0: strCd$ = "HREMP.ED_LOC"
    Case 1: strCd$ = "HREMP.ED_GLNO"
    Case 2: strCd$ = "HREMP.ED_EMP"
    Case 3: strCd$ = "HREMP.ED_REGION"
    Case 4: strCd$ = "HREMP.ED_ADMINBY"
    Case 5: strCd$ = "HREMP.ED_SECTION"  'Lucy June 29, 2000
    End Select
    'CodeCri = "({" & strCd$ & "} = '" & clpCode(intIdx%) & "')"
    CodeCri = "(" & strCd$ & " in  ('" & Replace(clpCode(intIdx%).Text, ",", "','") & "'))"
End If


If Len(CodeCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = CodeCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & CodeCri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_Div()

Dim DivCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level


If Len(clpDiv.Text) > 0 Then
    DivCri = "(HREMP.ED_DIV in ('" & Replace(clpDiv.Text, ",", "','") & "'))"
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

Private Sub Cri_EE()
Dim EECri As String

If Len(elpEEID.Text) > 0 Then
    EECri = "HREMP.ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
End If


If Len(EECri) >= 1 Then
    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_PT()
Dim EECri As String, OneSet%, x%

If Len(clpPT.Text) < 1 Then Exit Sub
EECri = "HREMP.ED_PT in ('" & Replace(clpPT.Text, ",", "','") & "')"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Function Cri_SetAll()
Dim x%, strRName$

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

Call glbCri_DeptUN(clpDept.Text)
glbstrSelCri = Replace(glbstrSelCri, "{", "")
glbstrSelCri = Replace(glbstrSelCri, "}", "")


Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
Call Cri_PT
Call Cri_EE
For x% = 0 To 5
    Call Cri_Code(x%)
Next x%

If frmDate.Visible Then Cri_FTDates

Select Case strRpt
Case "sn2369HS"
    Cri_FTDates
    Call XLSwriterHS
Case "rzChalBonus"
    Call XLSwriterCB
End Select
    
Cri_SetAll = True

Screen.MousePointer = DEFAULT

Exit Function


modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Cri_SetAll", "", "Select")
Cri_SetAll = False
Resume Next

End Function

Private Function CriCheck()
Dim x%

CriCheck = False

If Not clpDiv.ListChecker Then
'If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
    'MsgBox lStr("If Division Entered - it must be known")
    'clpDiv.SetFocus
    Exit Function
End If

If glbCompSerial = "S/N - 2347W" And InStr(1, fglbFileName, "SN234718.rpt") > 0 Then
    If Len(Trim(clpDiv.Text)) = 0 Then
        MsgBox lStr("Division cannot be left blank")
        clpDiv.SetFocus
        Exit Function
    End If
End If

If Not clpDept.ListChecker Then
'If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    'MsgBox "If Department Entered - it must be known"
    'clpDept.SetFocus
    Exit Function
End If

For x% = 0 To 5
    If x% <> 1 Then
        If Not clpCode(x).ListChecker Then Exit Function
    End If
Next x%

For x% = 0 To 1
 If Len(dlpDateRange(x%).Text) > 0 Then
    If Not IsDate(dlpDateRange(x%).Text) Then
        MsgBox "Not a valid date"
        dlpDateRange(x%).Text = ""
        dlpDateRange(x%).SetFocus
        Exit Function
    End If
 End If
Next x%

If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 Then
    'If Len(clpPT) > 0 And clpPT.Caption = "Unassigned" Then
        'MsgBox lStr("Category code must be valid")
        'clpPT.SetFocus
        Exit Function
    'End If
End If

If Not elpEEID.ListChecker Then
    Exit Function
End If

CriCheck = True
End Function

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
glbOnTop = "FRMREXCELRPT"


Call setRptCaption(Me)
lblGL.Caption = lStr("G/L")
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6
Call INI_Controls(Me)

Screen.MousePointer = DEFAULT


If strRpt = "rzChalBonus" Then
    fraCBAmount.Visible = True
    If glbCompSerial = "S/N - 2369W" Then txtAmount.Text = "1000"
    
Else
    fraCBAmount.Visible = True
End If
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

Private Sub Cri_FTDates()
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim x%

    If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
        TempCri = "(HR_OCC_HEALTH_SAFETY.EC_OCCDATE BETWEEN"
        TempCri = TempCri & Date_SQL(dlpDateRange(0).Text)
        TempCri = TempCri & " AND " & Date_SQL(dlpDateRange(1).Text) & ") "
        GoTo Cri_FTDatst
    End If

    For x% = 0 To 1
        If Len(dlpDateRange(x).Text) > 0 Then
            TempCri = "(HR_OCC_HEALTH_SAFETY.EC_OCCDATE"
            If x% = 0 Then
                TempCri = TempCri & " >= "
            Else
                TempCri = TempCri & " <= "
            End If
            TempCri = TempCri & Date_SQL(dlpDateRange(x).Text) & ") "
            GoTo Cri_FTDatst
        End If
    Next x%

Cri_FTDatst:
    If Len(TempCri) >= 1 Then
        If Not glbiOneWhere Then
            glbstrSelCri = TempCri
        Else
            glbstrSelCri = glbstrSelCri & " AND " & TempCri
        End If
        glbiOneWhere = True
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

Private Sub XLSwriterHS()
On Error GoTo EH

Dim rsDATA As New ADODB.Recordset
Dim rsCounter As New ADODB.Recordset
Dim exApp As Excel.Application
Dim exBook As Excel.Workbook
Dim exSheet As Excel.Worksheet
Dim c As Long
Dim strTemp As String
Dim strSQL As String
Dim xlsFileTmp As String
Dim xlsFileMat As String
Dim xRow As Long
Dim NewDateFormat As String

    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "sn2369HSTmp.xls"
    'Ticket# 8293
    xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "sn2369HS" & Trim(glbUserID) & ".xls"

    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat

    NewDateFormat = UCase(glbsDateFormat)
    If InStr(1, NewDateFormat, "YYYY") = 0 Then
        NewDateFormat = Replace(NewDateFormat, "YY", "YYYY")
    End If


    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(0).Caption = "Please wait..."

    FileCopy xlsFileTmp, xlsFileMat

    strSQL = "SELECT HR_OCC_HEALTH_SAFETY.EC_OCCDATE, HREMP.ED_SURNAME, HREMP.ED_FNAME, HRGL.GL_DESCR AS GLDept, "
    strSQL = strSQL & "HR_OCC_HEALTH_SAFETY.EC_SHIFT, HR_OCC_HEALTH_SAFETY.EC_OCCTM, tblItype.TB_DESC AS IType, tblBody.TB_DESC AS BodySite, "
    strSQL = strSQL & "HR_OCC_HEALTH_SAFETY.EC_HAZARD, HREMP.ED_PT, tblLocation.TB_DESC AS Location, tblEquip.TB_DESC AS Equipment, "
    strSQL = strSQL & "tblCause.TB_DESC AS Cause, HR_OCC_HEALTH_SAFETY.EC_SINJ, HR_OCC_HEALTH_SAFETY.EC_TYPE, "
    strSQL = strSQL & "HR_OCC_HEALTH_SAFETY.EC_CLASS, HR_OCC_HEALTH_SAFETY.EC_CASE "
    strSQL = strSQL & "FROM  HR_OCC_HEALTH_SAFETY INNER JOIN "
    strSQL = strSQL & "HREMP ON HR_OCC_HEALTH_SAFETY.EC_EMPNBR = HREMP.ED_EMPNBR INNER JOIN "
    strSQL = strSQL & "HRGL ON HREMP.ED_GLNO = HRGL.GL_NO INNER JOIN "
    strSQL = strSQL & "HRTABL tblItype ON HR_OCC_HEALTH_SAFETY.EC_CODE_TABL = tblItype.TB_NAME AND "
    strSQL = strSQL & "HR_OCC_HEALTH_SAFETY.EC_CODE = tblItype.TB_KEY INNER JOIN "
    strSQL = strSQL & "HRTABL tblBody ON HR_OCC_HEALTH_SAFETY.EC_PBODY_TABL = tblBody.TB_NAME AND "
    strSQL = strSQL & "HR_OCC_HEALTH_SAFETY.EC_PBODY = tblBody.TB_KEY INNER JOIN "
    strSQL = strSQL & "HRTABL tblLocation ON HR_OCC_HEALTH_SAFETY.EC_LOC_TABL = tblLocation.TB_NAME AND "
    strSQL = strSQL & "HR_OCC_HEALTH_SAFETY.EC_LOC = tblLocation.TB_KEY INNER JOIN "
    strSQL = strSQL & "HRTABL tblEquip ON HR_OCC_HEALTH_SAFETY.EC_EQUIP_TABL = tblEquip.TB_NAME AND "
    strSQL = strSQL & "HR_OCC_HEALTH_SAFETY.EC_EQUIP = tblEquip.TB_KEY INNER JOIN "
    strSQL = strSQL & "HRTABL tblCause ON HR_OCC_HEALTH_SAFETY.EC_CAUSECD_TABL = tblCause.TB_NAME AND "
    strSQL = strSQL & "HR_OCC_HEALTH_SAFETY.EC_CAUSECD = tblCause.TB_KEY "
    strSQL = strSQL & "WHERE " & glbstrSelCri
    strSQL = strSQL & " ORDER BY HR_OCC_HEALTH_SAFETY.EC_OCCDATE ASC"
    
    rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If rsDATA.EOF = False And rsDATA.BOF = False Then

        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)
    
        
        exSheet.Cells(1, 4) = "Health & Safety Report"
        exSheet.Cells(1, 1) = "Date: " & Format(Now, "mmm dd, yyyy")
        exSheet.Cells(2, 1) = "Time: " & Time$
        If Not (IsDate(dlpDateRange(0).Text) Or IsDate(dlpDateRange(1).Text)) Then
            exSheet.Cells(2, 5) = "No date entered"
        Else
            strTemp = ""
            If IsDate(dlpDateRange(0).Text) Then
                strTemp = "From Date: " & Format(dlpDateRange(0).Text, "mmm dd, yyyy") & "  "
            End If
            If IsDate(dlpDateRange(1).Text) Then
                strTemp = strTemp & "To Date: " & Format(dlpDateRange(1).Text, "mmm dd, yyyy")
            End If
            exSheet.Cells(2, 5) = strTemp
        End If
    
        xRow = 6
        c = 1
        
        Do
            exSheet.Cells(xRow, 2) = c 'No.
            exSheet.Cells(xRow, 3) = Format(rsDATA("EC_OCCDATE"), "DD/MMM/YYYY")
            exSheet.Cells(xRow, 4) = rsDATA("ED_Surname") & ", " & rsDATA("ED_FName")
            exSheet.Cells(xRow, 5) = rsDATA("GLDept")
            exSheet.Cells(xRow, 6) = rsDATA("EC_Shift")
            exSheet.Cells(xRow, 7) = Format(rsDATA("EC_OCCTM"), "AM/PM")
            exSheet.Cells(xRow, 8) = rsDATA("IType") & " " & rsDATA("BodySite")
            If rsDATA("ED_PT") = "FT" Then
                exSheet.Cells(xRow, 9) = rsDATA("EC_Hazard")
            ElseIf rsDATA("ED_PT") = "TMP" Then
                exSheet.Cells(xRow, 10) = rsDATA("EC_Hazard")
            End If
            exSheet.Cells(xRow, 11) = rsDATA("Location")
            exSheet.Cells(xRow, 12) = rsDATA("Equipment")
            exSheet.Cells(xRow, 13) = rsDATA("Cause")
            If rsDATA("EC_SINJ") = "I" Then
                exSheet.Cells(xRow, 15) = 1
            End If
            If rsDATA("ED_PT") = "FT" And (rsDATA("EC_TYPE") = "FA" Or rsDATA("EC_TYPE") = "MA") Then
                exSheet.Cells(xRow, 16) = 1
            ElseIf rsDATA("ED_PT") = "TMP" And (rsDATA("EC_TYPE") = "FA" Or rsDATA("EC_TYPE") = "MA") Then
                exSheet.Cells(xRow, 17) = 1
            End If
            If rsDATA("ED_PT") = "FT" And rsDATA("EC_CLASS") = "ERGO" Then
                exSheet.Cells(xRow, 18) = 1
            ElseIf rsDATA("ED_PT") = "TMP" And rsDATA("EC_CLASS") = "ERGO" Then
                exSheet.Cells(xRow, 19) = 1
            End If
            If rsDATA("EC_TYPE") = "LT" Then
                exSheet.Cells(xRow, 20) = 1
            End If
            
            strSQL = "SELECT HRTABL.TB_DESC AS Counter FROM HR_OHS_CORRECTIVE INNER JOIN "
            strSQL = strSQL & "HRTABL ON HR_OHS_CORRECTIVE.CR_Code_TABL = HRTABL.TB_NAME AND "
            strSQL = strSQL & "HR_OHS_CORRECTIVE.CR_Code = HRTABL.TB_KEY "
            strSQL = strSQL & "WHERE HR_OHS_CORRECTIVE.CR_CASE=" & rsDATA("EC_CASE")
            rsCounter.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
            If rsCounter.EOF = False And rsCounter.BOF = False Then
                strTemp = ""
                Do
                    strTemp = strTemp & rsCounter("Counter") & ", "
                    rsCounter.MoveNext
                Loop Until rsCounter.EOF
                If Right(strTemp, 2) = ", " Then
                    strTemp = Left(strTemp, Len(strTemp) - 2)
                End If
                exSheet.Cells(xRow, 14) = strTemp
            End If
            rsCounter.Close
            rsDATA.MoveNext
            c = c + 1
            xRow = xRow + 1
        Loop Until rsDATA.EOF
    End If
    rsDATA.Close
    
    'Save new Excel file as XLS
    exBook.Save
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Call Pause(1)
    'launch Excel file
    'Shell "Start " & GetShortName(xlsFileMat)
    If Not LanchXlsW98(xlsFileMat) Then
        Shell "cmd /c " & GetShortName(xlsFileMat)
    End If
            
exH:
    If Not exBook Is Nothing Then
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing
    End If
    Exit Sub
EH:

    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Screen.MousePointer = DEFAULT
    
    If Err = 1004 Then
        Resume Next
    End If
    
    If Err = 75 Then
        MsgBox Err.Description & Chr(10) & "Please close Excel Files."
        Resume exH
    End If
    If Err = 70 Then
        MsgBox Err.Description & Chr(10) & "Please close Excel Files."
        Resume exH
    End If

    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "XLSwriterHS", "", "Select")
    Resume exH
End Sub

Function LanchXlsW98(xFileName)
On Error GoTo Error_Deal
    LanchXlsW98 = False
    Shell "Start " & GetShortName(xFileName)
    LanchXlsW98 = True
Exit Function
Error_Deal:

End Function

Public Function GetShortName(ByVal sLongFileName As String) As String
    Dim lRetVal As Long, sShortPathName As String, iLen As Integer
    'Set up buffer area for API function call return
    sShortPathName = Space(255)
    iLen = Len(sShortPathName)

    'Call the function
    lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
    'Strip away unwanted characters.
    GetShortName = Left(sShortPathName, lRetVal)
End Function

Public Property Get Rptname() As String
    Rptname = strRpt
End Property

Public Property Let Rptname(RPT As String)
    strRpt = RPT
End Property

Private Sub XLSwriterCB()
On Error GoTo EH

Dim rsDATA As New ADODB.Recordset
Dim rsDays As New ADODB.Recordset
Dim exApp As Excel.Application
Dim exBook As Excel.Workbook
Dim exSheet As Excel.Worksheet
Dim c As Long
Dim strTemp As String
Dim strSQL As String
Dim xlsFileTmp As String
Dim xlsFileMat As String
Dim xRow As Long
Dim NewDateFormat As String
Dim xEmpnbr As Long

    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "rzChalBonusTmp.xls"
    'Ticket# 8293
    xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "rzChalBonus" & Trim(glbUserID) & ".xls"

    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat

    NewDateFormat = UCase(glbsDateFormat)
    If InStr(1, NewDateFormat, "YYYY") = 0 Then
        NewDateFormat = Replace(NewDateFormat, "YY", "YYYY")
    End If


    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(0).Caption = "Please wait..."

    FileCopy xlsFileTmp, xlsFileMat

'************************
    strSQL = "SELECT HREMP.ED_EMPNBR, HR_JOB_HISTORY.JH_SDATE, HR_JOB_HISTORY.JH_CURRENT, HRJOB.JB_DESCR,  HREMP.ED_FNAME , HREMP.ED_SURNAME "
    strSQL = strSQL & "FROM HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR INNER JOIN "
    strSQL = strSQL & "HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0)"
    If Len(glbstrSelCri) > 0 Then
        strSQL = strSQL & "AND " & glbstrSelCri
    End If
    rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If rsDATA.EOF = False And rsDATA.BOF = False Then

        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)
    
        
        exSheet.Cells(1, 3) = "Bonus Report" '"Challenge Bonus Report"
        exSheet.Cells(1, 1) = "Date: " & Format(Now, "mmm dd, yyyy")
        exSheet.Cells(2, 1) = "Time: " & Time$
        If Not (IsDate(dlpDateRange(0).Text) Or IsDate(dlpDateRange(1).Text)) Then
            exSheet.Cells(2, 5) = "No date entered"
        Else
            strTemp = ""
            If IsDate(dlpDateRange(0).Text) Then
                strTemp = "From Date: " & Format(dlpDateRange(0).Text, "mmm dd, yyyy") & "  "
            End If
            If IsDate(dlpDateRange(1).Text) Then
                strTemp = strTemp & "To Date: " & Format(dlpDateRange(1).Text, "mmm dd, yyyy")
            End If
            exSheet.Cells(2, 5) = strTemp
        End If
    
        xRow = 6
        c = 1
        
        Do
            xEmpnbr = rsDATA("ED_EMPNBR")
            'exSheet.Cells(xRow, 1) = c 'No.
            exSheet.Cells(xRow, 1) = rsDATA("ED_SURNAME") & ", " & rsDATA("ED_FNAME")
            exSheet.Cells(xRow, 2) = xEmpnbr
            exSheet.Cells(xRow, 3) = rsDATA("JH_SDATE")
            If IsDate(dlpDateRange(0).Text) Then
                If DateDiff("d", dlpDateRange(0).Text, rsDATA("JH_SDATE")) > 0 Then
                    exSheet.Cells(xRow, 4) = rsDATA("JH_SDATE")
                Else
                    exSheet.Cells(xRow, 4) = "Full Share"
                End If
            Else
                exSheet.Cells(xRow, 4) = "Full Share"
            End If
            '{HR_JOB_HISTORY.JH_WHRS} * 52 / {HR_JOB_HISTORY.JH_DHRS}
            strSQL = "SELECT JH_WHRS, JH_DHRS FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR=" & rsDATA("ED_EMPNBR")
            rsDays.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
            If rsDays.EOF = False And rsDays.BOF = False Then
                If rsDays("JH_DHRS") > 0 Then
                    If IsDate(dlpDateRange(0).Text) Then
                        exSheet.Cells(xRow, 7) = rsDays("JH_WHRS") * Abs(DateDiff("ww", dlpDateRange(0).Text, dlpDateRange(1).Text)) / rsDays("JH_DHRS")
                    Else
                        exSheet.Cells(xRow, 7) = 0
                    End If
                Else
                    exSheet.Cells(xRow, 7) = 0
                End If
            End If
            rsDays.Close
            
            strSQL = "SELECT count(AD_DOA) as Missed FROM HR_ATTENDANCE WHERE AD_DOA BETWEEN " & Date_SQL(dlpDateRange(0).Text) & " AND " & Date_SQL(dlpDateRange(1).Text)
            strSQL = strSQL & "AND AD_REASON NOT IN ('VAC', 'BT', 'GRA') AND AD_EMPNBR=" & xEmpnbr
            rsDays.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
            If rsDays.EOF = False And rsDays.BOF = False Then
                exSheet.Cells(xRow, 6) = exSheet.Cells(xRow, 7) - rsDays("Missed")
            End If
            rsDays.Close
            
            If exSheet.Cells(xRow, 6) = 0 Then
                exSheet.Cells(xRow, 5) = 0
            Else
                exSheet.Cells(xRow, 5) = 1
            End If
            
            If Len(txtAmount) > 0 And IsNumeric(txtAmount) Then
                exSheet.Cells(xRow, 8) = txtAmount
            Else
                exSheet.Cells(xRow, 8) = 1000
            End If
            If exSheet.Cells(xRow, 7) > 0 Then
                exSheet.Cells(xRow, 9) = exSheet.Cells(xRow, 8) * exSheet.Cells(xRow, 6) / exSheet.Cells(xRow, 7)
            Else
                exSheet.Cells(xRow, 9) = 0
            End If
            
            strSQL = "SELECT COUNT(CL_EMPNBR) as Counsel FROM HR_COUNSEL WHERE CL_EMPNBR=" & xEmpnbr
            rsDays.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
            If rsDays.EOF = False And rsDays.BOF = False Then
                exSheet.Cells(xRow, 10) = rsDays("Counsel")
                Select Case rsDays("Counsel")
                Case 0
                    exSheet.Cells(xRow, 11) = 0
                    exSheet.Cells(xRow, 12) = 0
                Case 1
                    exSheet.Cells(xRow, 11) = 10
                    exSheet.Cells(xRow, 12) = exSheet.Cells(xRow, 9) * 0.1
                Case 2
                    exSheet.Cells(xRow, 11) = 30
                    exSheet.Cells(xRow, 12) = exSheet.Cells(xRow, 9) * 0.3
                Case 3
                    exSheet.Cells(xRow, 11) = 50
                    exSheet.Cells(xRow, 12) = exSheet.Cells(xRow, 9) * 0.5
                End Select
            Else
                exSheet.Cells(xRow, 10) = 0
                exSheet.Cells(xRow, 11) = 0
                exSheet.Cells(xRow, 12) = 0
            End If
            rsDays.Close
            exSheet.Cells(xRow, 13) = exSheet.Cells(xRow, 9) - exSheet.Cells(xRow, 12)
            exSheet.Cells(xRow, 14) = rsDATA("JB_DESCR")
            rsDATA.MoveNext
            c = c + 1
            xRow = xRow + 1
        Loop Until rsDATA.EOF
    End If
    rsDATA.Close
    
    'Save new Excel file as XLS
    exBook.Save
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Call Pause(1)
    'launch Excel file
    'Shell "Start " & GetShortName(xlsFileMat)
    If Not LanchXlsW98(xlsFileMat) Then
        Shell "cmd /c " & GetShortName(xlsFileMat)
    End If
    
    
    
exH:
    If Not exBook Is Nothing Then
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing
    End If
    Exit Sub
EH:

    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Screen.MousePointer = DEFAULT
    
    If Err = 1004 Then
        Resume Next
    End If
    
    If Err = 75 Then
        MsgBox Err.Description & Chr(10) & "Please close Excel Files."
        Resume exH
    End If
    If Err = 70 Then
        MsgBox Err.Description & Chr(10) & "Please close Excel Files."
        Resume exH
    End If

    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "XLSwriterCB", "", "Select")
    Resume exH
End Sub

Private Sub SELATTWRK()
    On Error GoTo EH
    
    Dim strSQL As String
    Dim rsDATA As ADODB.Recordset
    
    strSQL = "DELETE FROM HRCBWRK WHERE WRKEMP='" & glbUserID & "'"
    gdbAdoIhr001W.BeginTrans
    gdbAdoIhr001W.Execute strSQL
    gdbAdoIhr001W.CommitTrans
    
    strSQL = "INSERT INTO HRCBWRK (CB_EMPNBR, WRKEMP) SELECT ED_EMPNBR, '" & glbUserID & "' FROM HREMP WHERE ED_PT = 'FT'"
    gdbAdoIhr001W.BeginTrans
    gdbAdoIhr001W.Execute strSQL
    gdbAdoIhr001W.CommitTrans
    'Get Attendance
    strSQL = "SELECT HREMP.ED_EMPNBR, count(HR_ATTENDANCE.AD_EMPNBR) as dayMissed "
    strSQL = strSQL & "FROM HREMP INNER JOIN HR_ATTENDANCE ON HREMP.ED_EMPNBR = HR_ATTENDANCE.AD_EMPNBR "
    strSQL = strSQL & "WHERE (HREMP.ED_PT = 'FT') "
    strSQL = strSQL & " And (HR_ATTENDANCE.AD_DOA >= " & Date_SQL(dlpDateRange(0)) & ") AND (HR_ATTENDANCE.AD_DOA <= " & Date_SQL(dlpDateRange(1)) & ")"
    strSQL = strSQL & "GROUP BY HREMP.ED_EMPNBR "
    strSQL = strSQL & "ORDER BY HREMP.ED_EMPNBR "
    rsDATA.Open strSQL, gdbAdoIhr001W, adOpenStatic, adLockOptimistic, adCmdText
    If rsDATA.EOF = False And rsDATA.BOF = False Then
        Do
            strSQL = "UPDATE HRCBWRK SET CB_MISSED=" & rsDATA("dayMissed") & " WHERE CB_EMPNBR=" & rsDATA("ED_EMPNBR") & " AND WRKEMP='" & glbUserID & "'"
            gdbAdoIhr001W.BeginTrans
            gdbAdoIhr001W.Execute strSQL
            gdbAdoIhr001W.CommitTrans
            rsDATA.MoveNext
        Loop Until rsDATA.EOF
    End If
    rsDATA.Close
    
    'Get Counselling
    strSQL = "SELECT COUNT(HR_COUNSEL.CL_ID) AS Counsel, HREMP.ED_EMPNBR"
    strSQL = strSQL & "FROM HREMP RIGHT OUTER JOIN"
    strSQL = strSQL & "HR_COUNSEL ON HREMP.ED_EMPNBR = HR_COUNSEL.CL_EMPNBR"
    strSQL = strSQL & "WHERE (HREMP.ED_PT = 'FT') AND (HR_COUNSEL.CL_COUDATE BETWEEN " & Date_SQL(dlpDateRange(0)) & " AND"
    strSQL = strSQL & Date_SQL(dlpDateRange(1)) & ")"
    strSQL = strSQL & "GROUP BY HREMP.ED_EMPNBR"
    rsDATA.Open strSQL, gdbAdoIhr001W, adOpenStatic, adLockOptimistic, adCmdText
    If rsDATA.EOF = False And rsDATA.BOF = False Then
        Do
            strSQL = "UPDATE HRCBWRK SET CB_COUNSEL=" & rsDATA("Counsel") & " WHERE CB_EMPNBR=" & rsDATA("ED_EMPNBR") & " AND WRKEMP='" & glbUserID & "'"
            gdbAdoIhr001W.BeginTrans
            gdbAdoIhr001W.Execute strSQL
            gdbAdoIhr001W.CommitTrans
            rsDATA.MoveNext
        Loop Until rsDATA.EOF
    End If
    rsDATA.Close
exH:
    Set rsDATA = Nothing
    Exit Sub
EH:
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Screen.MousePointer = DEFAULT
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "XLSwriterHS", "", "Select")
    Resume exH
End Sub

