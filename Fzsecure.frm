VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRSecure 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Security Master Report"
   ClientHeight    =   5595
   ClientLeft      =   210
   ClientTop       =   1395
   ClientWidth     =   9765
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5595
   ScaleWidth      =   9765
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbSecTemplate 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2110
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Tag             =   "40-Security Template"
      Top             =   1845
      Width           =   2325
   End
   Begin VB.TextBox txtDateRange 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   5520
      MaxLength       =   12
      TabIndex        =   9
      Tag             =   "40-CT/OT Hours upto and including this date"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtDateRange 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   5520
      MaxLength       =   12
      TabIndex        =   8
      Tag             =   "40-CT/OT Hours from and including this date forward"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox comGroup 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Tag             =   "Second level of grouping records"
      Top             =   3540
      Width           =   2325
   End
   Begin VB.ComboBox comGroup 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Tag             =   "First Level of grouping records"
      Top             =   3210
      Width           =   2325
   End
   Begin Threed.SSOption optGrouping 
      Height          =   255
      Index           =   0
      Left            =   2085
      TabIndex        =   5
      Tag             =   "There are no report grouping option for this report"
      Top             =   2460
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "   Security && Code Matrix"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   750
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "n/a"
      MaxLength       =   0
      LookupType      =   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   390
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "n/a"
      MaxLength       =   0
      LookupType      =   1
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Tag             =   "10-Enter Employee Number"
      Top             =   1110
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.EmployeeLookup elpUSERID 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Tag             =   "10-User ID"
      Top             =   1470
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   6120
      Top             =   4440
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
   Begin Threed.SSOption optGrouping 
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   20
      Tag             =   "Report Grouping option available"
      Top             =   2460
      Width           =   2655
      _Version        =   65536
      _ExtentX        =   4683
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "   Employee/Department Matrix"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption optGrouping 
      Height          =   255
      Index           =   2
      Left            =   7320
      TabIndex        =   21
      Tag             =   "Report Grouping option available"
      Top             =   2460
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   " User Listing"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Security Template"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   1905
      Width           =   1275
   End
   Begin VB.Label lblEENum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   1515
      Width           =   540
   End
   Begin VB.Label lblGrp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Final Sort"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   3570
      Width           =   660
   End
   Begin VB.Label lblGrp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Grouping #1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   3255
      Width           =   885
   End
   Begin VB.Label lblRepGrp 
      Caption         =   "Report Grouping"
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
      TabIndex        =   15
      Top             =   3015
      Width           =   1575
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
      TabIndex        =   14
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Type of Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   13
      Top             =   2475
      Width           =   1065
   End
   Begin VB.Label lblEENum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   1155
      Width           =   1290
   End
   Begin VB.Label lblDept 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   795
      Width           =   825
   End
   Begin VB.Label lblDiv 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   435
      Width           =   555
   End
End
Attribute VB_Name = "frmRSecure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim X%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm("Security Master Report Criteria", Me) Then Exit Sub
    Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    X% = Cri_SetAll()
    Me.vbxCrystal.Destination = 1
    Me.vbxCrystal.WindowTitle = "Security Master Report"
    MDIMain.Timer1.Enabled = False
    'Me.vbxCrystal.Password = gstrAccPWord$
    'Me.vbxCrystal.UserName = gstrAccUID$
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    Call set_PrintState(True)
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
End If

Exit Sub

PrntErr:
MsgBox "Error Printing - check your Windows Printer setup"
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Public Sub cmdView_Click()
Dim X%
Dim strWHand As String

If CriCheck() Then
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    Call set_PrintState(False)
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    X% = Cri_SetAll()
    Me.vbxCrystal.Destination = 0
    Me.vbxCrystal.WindowTitle = "Security Master Report"
    MDIMain.Timer1.Enabled = False
    'Me.vbxCrystal.Password = gstrAccPWord$
    'Me.vbxCrystal.UserName = gstrAccUID$
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    Call set_PrintState(True)
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
End If
End Sub

Private Sub cmbSecTemplate_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub comGroup_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub comGrpLoad()
    
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Region")
    If glbLinamar Then ' Frank May 2,2001
        comGroup(0).AddItem ("Home Line")
    End If
    comGroup(0).AddItem "Template"
    comGroup(0).AddItem "(none)"
    comGroup(1).AddItem "Employee Name"
    comGroup(0).ListIndex = 0
    comGroup(1).ListIndex = 0
    
End Sub

Private Sub Cri_Div()
Dim DivCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level

If Len(clpDiv.Text) > 0 Then
    DivCri = "({HREMP.ED_DIV} in ['" & Replace(clpDiv.Text, ",", "','") & "'])"
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
    EECri = "{HREMP.ED_EMPNBR} IN [" & getEmpnbr(elpEEID.Text) & "] "
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

Private Sub Cri_USERID()
Dim EECri As String

If Len(elpUSERID.Text) > 0 Then
    EECri = "LowerCase({HR_SECURE_BASIC.USERID}) IN ['" & Replace(Replace(Trim(LCase(elpUSERID.Text)), "'", "''"), ",", "','") & "'] "
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

Private Sub Cri_Sec_Template()
Dim EECri As String

If Len(cmbSecTemplate) > 0 Then
    'EECri = "{HR_SECURE_BASIC.SECURE_TEMPLATE} IN ['" & Replace(Trim(cmbSecTemplate.Text), ",", "','") & "'] "
    EECri = "{HR_SECURE_BASIC.SECURE_TEMPLATE} = '" & cmbSecTemplate.Text & "'"
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

Private Function Cri_SetAll()
 Dim X%

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

' call cri models set both glbiONeWhere(set in report)
' and strSelCri(string selection criteria
'Call glbCri_Dept(Me)  'laura nov 22, 1997
Call glbCri_DeptUN(clpDept.Text)

Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
Call Cri_EE
Call Cri_Sec_Template

'X% = Cri_Sorts()
If optGrouping(0) Then
    'Call SECWRK
    'Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzsecurd.rpt"
    'Me.vbxCrystal.SelectionFormula = "{HRSECWRK.WRKEMP}='" & glbUserID & "'"
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzsecurity.rpt"
    Call Cri_USERID
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
ElseIf optGrouping(1) Then
    If comGroup(0) = "(none)" Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzsecus1.rpt"
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzsecurs.rpt"
    End If
    X% = Cri_Sorts()
    Call Cri_USERID
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
Else
    If comGroup(0) = "(none)" Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzSecTmpl1.rpt"
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzSecTmpl.rpt"
    End If
    X% = Cri_Sorts()
    Call Cri_USERID
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
    If optGrouping(0) Then
        Me.vbxCrystal.SubreportToChange = "rgtabsec.rpt"
        Me.vbxCrystal.Connect = RptODBC_SQL
        Me.vbxCrystal.SubreportToChange = ""
    End If
Else
    If Not optGrouping(0) Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For X% = 0 To 4
            Me.vbxCrystal.DataFiles(X%) = glbIHRDB
        Next
        Me.vbxCrystal.DataFiles(5) = glbIHRDBW
        Me.vbxCrystal.DataFiles(6) = glbIHRDB
        Me.vbxCrystal.DataFiles(7) = glbIHRDB
        
        If optGrouping(0) Then
            Me.vbxCrystal.SubreportToChange = "rgtabsec.rpt"
            Me.vbxCrystal.Connect = "PWD=petman;"
            For X% = 0 To 1
                Me.vbxCrystal.DataFiles(X%) = glbIHRDB
            Next
            Me.vbxCrystal.SubreportToChange = ""
        End If
    End If
End If
Cri_SetAll = True

Screen.MousePointer = DEFAULT
Exit Function


modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "FR Security Time", "Security Report", "Select")
Cri_SetAll = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim X%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

GrpIdx% = comGroup(0).ListIndex
If Not glbLinamar Then ' Frank May 2, 2001
    Select Case GrpIdx%
        Case 0: grpField$ = "{HR_DIVISION.Division_Name}"
        Case 1: grpField$ = "{HRDEPT.DF_NAME}"
        Case 2: grpField$ = "{tblRegion.TB_DESC}"
        Case 3: grpField$ = "{HR_SECURE_BASIC.SECURE_TEMPLATE}"
        Case 4: grpField$ = "(none)"
    End Select
Else
    Select Case GrpIdx%
        Case 0: grpField$ = "{HR_DIVISION.Division_Name}"
        Case 1: grpField$ = "{HRDEPT.DF_NAME}"
        Case 2: grpField$ = "{tblRegion.TB_DESC}"
        Case 3: grpField$ = "{LN_HOMES.TB_DESC}"
        Case 4: grpField$ = "{HR_SECURE_BASIC.SECURE_TEMPLATE}"
        Case 5: grpField$ = "(none)"
    End Select
End If

If comGroup(0) <> "(none)" Then
    dscGroup$ = comGroup(0).Text
    dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
    Me.vbxCrystal.Formulas(0) = dscGroup$
    
    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
    End Select
    grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(1) = grpCond$

    strSFormat$ = "GH1;T;T;F;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(0) = strSFormat$
Else
    grpCond$ = "GROUP" & CStr(1) & ";" & "{@EFullName}" & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$
    'Me.vbxCrystal.SectionFormat(0) = "HEADER;F;F;F;T;X;X;X;X"
    'Me.vbxCrystal.SectionFormat(1) = "GF1;F;F;F;T;X;X;X;X"
End If

If Not optGrouping(0) Then
    Call setRptLabel(Me, 0)
    Me.vbxCrystal.Formulas(18) = "lblSection='" & lStr("Section") & "'"
    Me.vbxCrystal.Formulas(16) = "lblAdmin='" & lStr("Administered By") & "'"
End If

Cri_Sorts = z% ' next section number to format

End Function

Private Function CriCheck()
Dim I%

CriCheck = False

'Hemu - 05/14/2003 Begin
If Not clpDiv.ListChecker Then
'If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
    'MsgBox "If " & lblDiv.Caption & " Entered - it must be known"
    'clpDiv.SetFocus
    Exit Function
End If
'Hemu - 05/14/2003 End

If Not clpDept.ListChecker Then
'If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    'MsgBox "If Department Entered - it must be known"
     'clpDept.SetFocus
    Exit Function
End If

For I% = 0 To 1
    If Len(txtDateRange(I%)) > 0 Then
        If Not IsDate(txtDateRange(I%)) Then
            MsgBox "Not a valid date"
            txtDateRange(I%) = ""
            txtDateRange(I%).SetFocus
            Exit Function
        End If
    End If
Next I%

If Not elpEEID.ListChecker Then
    Exit Function
End If

If Not elpUSERID.ListChecker Then
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
glbOnTop = Me.name

Screen.MousePointer = HOURGLASS

Call comGrpLoad

Call setRptCaption(Me)

'Ticket #20585 - Template based security profile
Call Populate_Security_Template

Screen.MousePointer = DEFAULT

lblRepGrp.Visible = False   'js-01Apr99
lblGrp(0).Visible = False   '
lblGrp(3).Visible = False   '
comGroup(0).Visible = False '
comGroup(1).Visible = False '

Call INI_Controls(Me)

elpUSERID.LookupType = 2

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

Private Sub optGrouping_Click(Index As Integer, Value As Integer)

If optGrouping(1).Value = True Or optGrouping(2).Value = True Then
    lblRepGrp.Visible = True    'js-01Apr99
    lblGrp(0).Visible = True    '
    lblGrp(3).Visible = True    '
    comGroup(0).Visible = True  '
    comGroup(1).Visible = True  '
    'panReportG.Visible = True  '
Else
    lblRepGrp.Visible = False   'js-01Apr99
    lblGrp(0).Visible = False   '
    lblGrp(3).Visible = False   '
    comGroup(0).Visible = False '
    comGroup(1).Visible = False '
    'panReportG.Visible = False '
End If
End Sub

Private Sub optGrouping_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtDateRange_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub SECWRK()

Dim SQLQ, xField As String, X
Dim xQue
Dim rsSEC As New ADODB.Recordset
Dim rsFun As New ADODB.Recordset
Dim rsSECWrk As New ADODB.Recordset
Dim prec%, lngRecs&
MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 5
    
SQLQ = "select * from HR_SECURE_BASIC where (EMPNBR is null "
SQLQ = SQLQ & " or EMPNBR in (SELECT ED_EMPNBR FROM HREMP "
SQLQ = SQLQ & " where " & Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")") & "))"
If Len(elpUSERID.Text) > 0 Then SQLQ = SQLQ & " and Lower(USERID) IN ('" & Replace(Trim(Replace(LCase(elpUSERID.Text), "'", "''")), ",", "','") & "')"
rsSEC.Open SQLQ, gdbAdoIhr001, adOpenStatic
lngRecs& = rsSEC.RecordCount

gdbAdoIhr001W.BeginTrans
gdbAdoIhr001W.Execute "delete from HRSECWRK WHERE WRKEMP='" & glbUserID & "'"
gdbAdoIhr001W.CommitTrans
rsSECWrk.Open "HRSECWRK", gdbAdoIhr001W, adOpenStatic, adLockPessimistic, adCmdTableDirect
MDIMain.panHelp(0).FloodPercent = 10
Do Until rsSEC.EOF
    MDIMain.panHelp(0).FloodPercent = (prec% / lngRecs&) * 90 + 10
    rsSECWrk.AddNew
    rsSECWrk("EMPNBR") = rsSEC("EMPNBR")
    rsSECWrk("USERID") = rsSEC("USERID")
    xQue = False
    For X = 1 To rsSECWrk.Fields.count - 1
        xField = rsSECWrk.Fields(X).name
        xQue = True
        If UCase(xField) = UCase("Basic_Inquiry") Then xQue = True
        If UCase(xField) = "PS_CHGDATE" Then xQue = False
        If xQue Then
            SQLQ = "SELECT * FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(rsSEC("USERID"), "'", "''")
            SQLQ = SQLQ & "' and " & Upper_SQL(Field_SQL("FUNCTION")) & "='" & IIf(glbOracle, UCase(xField), xField) & "'"
            rsFun.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsFun.EOF Then rsSECWrk(xField) = rsFun("ACCESSABLE")
            rsFun.Close
        End If
    Next
    rsSECWrk("WRKEMP") = glbUserID
    rsSECWrk.Update
    rsSEC.MoveNext
    prec% = prec% + 1
    DoEvents
Loop

MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""

End Sub

Private Sub Populate_Security_Template()
    Dim SQLQ As String
    Dim rsSecBasic As New ADODB.Recordset
    
    'Ticket #20585 - Template based Security. Populate the combo box with Templates.
    
    'Clear the combo box
    cmbSecTemplate.Clear
    
    'Not all users will have a template or needs to have a template
    cmbSecTemplate.AddItem ""
    
    'Add Default value "TEMPLATE". This value is to be chosen when templates are created
    cmbSecTemplate.AddItem "TEMPLATE"
    
    
    'Populate User Template combo box list
    SQLQ = "SELECT USERID, SECURE_TEMPLATE FROM HR_SECURE_BASIC WHERE SECURE_TEMPLATE = 'TEMPLATE'"
    SQLQ = SQLQ & " ORDER BY SECURE_TEMPLATE"
    rsSecBasic.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsSecBasic.EOF
        If Not IsNull(rsSecBasic("USERID")) Then
            cmbSecTemplate.AddItem rsSecBasic("USERID")
        End If
        rsSecBasic.MoveNext
    Loop
    rsSecBasic.Close
    Set rsSecBasic = Nothing
    
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

