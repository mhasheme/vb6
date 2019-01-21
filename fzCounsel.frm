VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRCounsel 
   Caption         =   "Counseling Report"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10005
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   10005
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkCDate 
      Caption         =   "Show Blank Counselling Date Only"
      Height          =   255
      Left            =   3360
      TabIndex        =   19
      Tag             =   "Show Comments"
      Top             =   5880
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2475
      MaxLength       =   4
      TabIndex        =   14
      Tag             =   "00-Employee Position Shift"
      Top             =   4800
      Visible         =   0   'False
      Width           =   450
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   7
      Left            =   2160
      TabIndex        =   13
      Tag             =   "00-Enter Administered By Code"
      Top             =   4464
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
      Left            =   2160
      TabIndex        =   11
      Tag             =   "00-Enter Region Code"
      Top             =   3800
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   2160
      TabIndex        =   8
      Tag             =   "00-Attendance Codes"
      Top             =   3136
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "CETY"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   2160
      TabIndex        =   7
      Tag             =   "00-Attendance Codes"
      Top             =   2804
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "CERE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   2160
      TabIndex        =   4
      Tag             =   "00-Enter Status Code"
      Top             =   1808
      Width           =   6960
      _ExtentX        =   12277
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDEM"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Tag             =   "EDPT-Category"
      Top             =   2140
      Width           =   6960
      _ExtentX        =   12277
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDPT"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   3
      Tag             =   "00-Enter Union Code"
      Top             =   1476
      Width           =   6960
      _ExtentX        =   12277
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDOR"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   2
      Tag             =   "00-Enter Location Code"
      Top             =   1144
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   812
      Width           =   6960
      _ExtentX        =   12277
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
      Left            =   2160
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   480
      Width           =   6960
      _ExtentX        =   12277
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "n/a"
      MaxLength       =   0
      LookupType      =   1
      MultiSelect     =   -1  'True
   End
   Begin VB.CheckBox chkComments 
      Caption         =   "Show Comments"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Tag             =   "Show Comments"
      Top             =   5880
      Value           =   1  'Checked
      Width           =   2235
   End
   Begin VB.ComboBox comGroup 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   1980
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Tag             =   "First Level of grouping records"
      Top             =   6750
      Width           =   2325
   End
   Begin VB.ComboBox comGroup 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   1980
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Tag             =   "Second level of grouping records"
      Top             =   7065
      Width           =   2325
   End
   Begin VB.ComboBox comGroup 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   2
      Left            =   1980
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Tag             =   "Third level of grouping records"
      Top             =   7380
      Width           =   2325
   End
   Begin VB.ComboBox comGroup 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   3
      Left            =   1980
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Tag             =   "Final sorting of records - no totals"
      Top             =   7695
      Width           =   2325
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   6
      Left            =   2160
      TabIndex        =   12
      Tag             =   "00-Enter Section Code"
      Top             =   4132
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   4080
      TabIndex        =   10
      Tag             =   "40-Date upto and including this date forward"
      Top             =   3468
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   9
      Tag             =   "40-Date from and including this date forward"
      Top             =   3468
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   2160
      TabIndex        =   6
      Tag             =   "10-Enter Employee Number"
      Top             =   2472
      Width           =   6960
      _ExtentX        =   12277
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   6640
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   7080
      Top             =   7800
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
   Begin INFOHR_Controls.EmployeeLookup elpRept 
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   15
      Tag             =   "10-Reporting Authority 1"
      Top             =   5140
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   503
      ShowDescription =   0   'False
      RefreshDescriptionWhen=   2
   End
   Begin INFOHR_Controls.EmployeeLookup elpRept 
      Height          =   285
      Index           =   1
      Left            =   3930
      TabIndex        =   16
      Tag             =   "10-Reporting Authority 2"
      Top             =   5140
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   503
      RefreshDescriptionWhen=   2
   End
   Begin INFOHR_Controls.EmployeeLookup elpRept 
      Height          =   285
      Index           =   2
      Left            =   5715
      TabIndex        =   17
      Tag             =   "10-Reporting Authority 3"
      Top             =   5140
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      ShowDescription =   0   'False
      RefreshDescriptionWhen=   2
   End
   Begin VB.Label lblRep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reporting Authority:"
      Height          =   195
      Left            =   120
      TabIndex        =   44
      Top             =   5185
      Width           =   1395
   End
   Begin VB.Label lblShift 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Shift"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   43
      Top             =   4845
      Visible         =   0   'False
      Width           =   645
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
      TabIndex        =   42
      Top             =   2185
      Width           =   630
   End
   Begin VB.Label lblReasonCode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Reason Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   41
      Top             =   2849
      Width           =   975
   End
   Begin VB.Label lblAttendCrit 
      BackStyle       =   0  'Transparent
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
      TabIndex        =   40
      Top             =   120
      Width           =   1575
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
      TabIndex        =   39
      Top             =   525
      Width           =   555
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
      TabIndex        =   38
      Top             =   857
      Width           =   825
   End
   Begin VB.Label lblUnion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Union"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   37
      Top             =   1521
      Width           =   420
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   36
      Top             =   1853
      Width           =   450
   End
   Begin VB.Label lblEENum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   35
      Top             =   2517
      Width           =   1290
   End
   Begin VB.Label lblTypeCode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Type Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   34
      Top             =   3181
      Width           =   780
   End
   Begin VB.Label lblFromTo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "From / To Counseling Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   33
      Top             =   3513
      Width           =   1920
   End
   Begin VB.Label lblReprtGrping 
      BackStyle       =   0  'Transparent
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
      TabIndex        =   32
      Top             =   6525
      Width           =   1695
   End
   Begin VB.Label lblGrp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Grouping #1"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   31
      Top             =   6765
      Width           =   885
   End
   Begin VB.Label lblGrp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Grouping #2"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   30
      Top             =   7095
      Width           =   885
   End
   Begin VB.Label lblGrp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Grouping #3"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   29
      Top             =   7410
      Width           =   885
   End
   Begin VB.Label lblGrp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Final Sort"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   28
      Top             =   7725
      Width           =   660
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
      TabIndex        =   27
      Top             =   1189
      Width           =   615
   End
   Begin VB.Label lblRegion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   26
      Top             =   3845
      Width           =   510
   End
   Begin VB.Label lblAdmin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Administered By"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   25
      Top             =   4509
      Width           =   1125
   End
   Begin VB.Label lblSection 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   4177
      Width           =   540
   End
End
Attribute VB_Name = "frmRCounsel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkComments_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim X%
Dim strMessage As String

On Error GoTo PrntErr

If glbCompSerial = "S/N - 2376W" Then ' George added for Assembling of 1st Nations #9535
    strMessage = "Assets Report Criteria"
Else
    strMessage = lStr("Counseling Report Criteria")
End If
If CriCheck() Then
    Call set_PrintState(False)
    If Not PrtForm(strMessage, Me) Then Exit Sub
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    X% = Cri_SetAll()
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
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString
Resume Next
Screen.MousePointer = DEFAULT
End Sub

Public Sub cmdView_Click()
Dim X%, SQLQ
Dim strWHand As String, MyQuery As QueryDef
On Error GoTo CRW_Err

If CriCheck() Then
    Call set_PrintState(False)
    Screen.MousePointer = HOURGLASS
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False

    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

    X% = Cri_SetAll()
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
'MsgBox Error$(Err)
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString
If MDIMain.panHelp(0).FloodType > 0 Then MDIMain.panHelp(0).FloodType = 0  'ADDED BY RAUBREY 4/10/97
Resume Next
Screen.MousePointer = DEFAULT

End Sub

Private Sub comGroup_GotFocus(Index As Integer)
    Call SetPanHelp(Me.ActiveControl)
End Sub

Public Sub comGrpLoad()
    comGroup(0).Clear
    comGroup(1).Clear
    comGroup(2).Clear
    comGroup(3).Clear
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")
    comGroup(0).AddItem "Employee Name"
    comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
    If glbLinamar Then
        comGroup(0).AddItem "Employment Type"
    End If
    If Not glbMulti Then comGroup(0).AddItem "Shift"
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem "Reporting Authority #1"
    comGroup(0).AddItem "(none)"
    comGroup(1).AddItem "Employee Name"
    comGroup(1).AddItem "Type Code"
    comGroup(1).AddItem "(none)"
    comGroup(2).AddItem "Type Code"
    comGroup(2).AddItem "(none)"
    If glbCompSerial = "S/N - 2376W" Then ' George added for Assembling of 1st Nations #9535
        comGroup(3).AddItem "Issuing Date"
    Else
        comGroup(3).AddItem lStr("Counseling Date")
    End If
    comGroup(3).AddItem "(none)"
    comGroup(0).ListIndex = 0
    comGroup(1).ListIndex = 0
    comGroup(2).ListIndex = 0
    comGroup(3).ListIndex = 0
    comGroup(3).Enabled = False
End Sub

Private Sub Cri_Code(intIdx%)
Dim CodeCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

If Len(clpCode(intIdx%).Text) > 0 Then
    Select Case intIdx%
        Case 0: strCd$ = "HREMP.ED_LOC"
        Case 1: strCd$ = "HREMP.ED_ORG"
        Case 2: strCd$ = "HREMP.ED_EMP"
        Case 3: strCd$ = "HR_COUNSEL.CL_REASON"
        Case 4: strCd$ = "HR_COUNSEL.CL_TYPE"
        Case 5: strCd$ = "HREMP.ED_REGION"
        Case 6: strCd$ = "HREMP.ED_SECTION"
        Case 7: strCd$ = "HREMP.ED_ADMINBY"
    End Select
        CodeCri = "({" & strCd$ & "} in  ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
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
        glbstrSelCri = glbstrSelCri & " AND (" & EECri & ") "
    Else
        glbstrSelCri = EECri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_BlankCDates()
Dim EECri As String, OneSet%, X%

If Not (chkCDate.Value = 1) Then
    Exit Sub
Else
    EECri = " isnull({HR_COUNSEL.CL_COUDATE}) "
End If

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_FTDates()
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim X%

If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    TempCri = "({HR_COUNSEL.CL_COUDATE} "
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    GoTo Cri_FTDatst
End If

For X% = 0 To 1
    If Len(dlpDateRange(X%).Text) > 0 Then
        TempCri = "({HR_COUNSEL.CL_COUDATE} "
        If X% = 0 Then
            TempCri = TempCri & " >= "
        Else
            TempCri = TempCri & " <= "
        End If
        dtYYY% = Year(dlpDateRange(X%).Text)
        dtMM% = month(dlpDateRange(X%).Text)
        dtDD% = Day(dlpDateRange(X%).Text)
        TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        GoTo Cri_FTDatst
    End If
Next X%

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

Private Sub Cri_PT()
Dim EECri As String, OneSet%, X%

If Len(clpPT.Text) < 1 Then Exit Sub
EECri = "{HREMP.ED_PT} in ['" & Replace(clpPT.Text, ",", "','") & "']"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_RepAuth()
Dim TempCri As String
Dim EECri As String, LocCri As String
Dim I, xTemp As Boolean
    
    xTemp = False
    EECri = ""

    If Len(Trim(elpRept(0).Text)) > 0 Then
        EECri = EECri & "{HR_JOB_HISTORY.JH_REPTAU} = " & Trim(elpRept(0).Text) & " "
        xTemp = True
    End If
    If Len(Trim(elpRept(1).Text)) > 0 Then
        If xTemp Then
            EECri = EECri & "and {HR_JOB_HISTORY.JH_REPTAU2} = " & Trim(elpRept(1).Text) & " "
        Else
            EECri = EECri & "{HR_JOB_HISTORY.JH_REPTAU2} = " & Trim(elpRept(1).Text) & " "
        End If
        xTemp = True
    End If
    If Len(Trim(elpRept(2).Text)) > 0 Then
        If xTemp Then
            EECri = EECri & "and {HR_JOB_HISTORY.JH_REPTAU3} = " & Trim(elpRept(2).Text) & " "
        Else
            EECri = EECri & "{HR_JOB_HISTORY.JH_REPTAU3} = " & Trim(elpRept(2).Text) & " "
        End If
        xTemp = True
    End If
        
    
    If Len(EECri) > 0 Then
        If glbiOneWhere Then
          glbstrSelCri = glbstrSelCri & " AND " & EECri
        Else
          glbstrSelCri = EECri
        End If
    Else
        Exit Sub
    End If
    glbiOneWhere = True
    
End Sub

Private Function Cri_SetAll()
Dim X%

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

'Call glbCri_Dept(Me)  'laura nov 21, 1997
Call glbCri_DeptUN(clpDept.Text)
Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
For X% = 0 To 7
    Call Cri_Code(X%)
Next X%
Call Cri_PT
Call Cri_Shift
Call Cri_EE
Call Cri_FTDates
Call Cri_RepAuth

If glbBurlTech Then
    Call Cri_BlankCDates
End If
' report name

If glbCompSerial = "S/N - 2376W" Then ' George added for Assembling of 1st Nations #9535
    vbxCrystal.ReportFileName = glbIHRREPORTS & "rzcounsl_FN.rpt"
    Me.vbxCrystal.WindowTitle = "Assets Report"
Else
    vbxCrystal.ReportFileName = glbIHRREPORTS & "rzcounsl.rpt"
    Me.vbxCrystal.WindowTitle = lStr("Counseling Report")
End If

' set to sorting/grouping criteria
X% = Cri_Sorts()   ' returns number of sections formated

'Release 8.0 - Ticket #22682: View Own security
'If View Own not checked then do not retrieve Counselling information of the User/Employee No
If Len(glbUserEmpNo) > 0 And glbUserEmpNo <> 0 And Not gSec_Counsel_ViewOwn Then
    'Do not show user's Counselling records based on the Employee # associated to the User.
    If Len(glbstrSelCri) > 0 Then
        glbstrSelCri = glbstrSelCri & " AND {HR_COUNSEL.CL_EMPNBR} <> " & glbUserEmpNo
    Else
        glbstrSelCri = glbstrSelCri & " {HR_COUNSEL.CL_EMPNBR} <> " & glbUserEmpNo
    End If
End If

'set location for database tables
If Len(glbstrSelCri) >= 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If

Me.vbxCrystal.Connect = RptODBC_SQL


Cri_SetAll = True
Screen.MousePointer = DEFAULT
Exit Function


modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "FR Comp Time", "Comp Report", "Select")
Cri_SetAll = False
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub Cri_Shift()
Dim EECri As String, OneSet%, X%

If Len(txtShift.Text) < 1 Then Exit Sub
EECri = "{HREMP.ED_SHIFT}= '" & txtShift.Text & "'"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True
End Sub

Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim X%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$, DOA
Dim dscGroup$, GrpIdx%, SavGrp1, SavGrp2, SavGrp3
z% = 0
Cri_Sorts = 0


If dlpDateRange(0).Text <> "" And dlpDateRange(1).Text <> "" Then
    strSFormat$ = "As of " & dlpDateRange(0).Text & " through " & dlpDateRange(1).Text
    Me.vbxCrystal.Formulas(2) = "Daterange = '" & strSFormat$ & "'"
    
Else
    strSFormat$ = "No date entered"
    Me.vbxCrystal.Formulas(2) = "Daterange = '" & strSFormat$ & "'"
End If

If chkComments.Value = 1 Then
    Me.vbxCrystal.Formulas(3) = "ShowComments = True"
Else
    Me.vbxCrystal.Formulas(3) = "ShowComments = False"
End If
'Counseling Label
Me.vbxCrystal.Formulas(4) = "lblRptTitle = '" & lStr("Counseling") & " Report'"
Me.vbxCrystal.Formulas(5) = "lblCouDate = '" & lStr("Counseling") & " Date'"
Me.vbxCrystal.Formulas(6) = "lblCouBy = '" & lStr("Counseling") & " By'"


grpField$ = getEGroup(comGroup(0).Text)
If grpField$ = "(none)" Then grpField$ = "{HREMP.ED_COMPNO}"

SavGrp1 = grpField$
If Not (grpField$ = "{HREMP.ED_COMPNO}") Then 'If GrpIdx% < 5 Then
  Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = '" & comGroup(0).Text & "'"
  Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = " & grpField$
Else
  Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = ''"
  Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = ''"
  Me.vbxCrystal.SectionFormat(z%) = "GF1;F;X;X;X;X;X;X"
  z% = z% + 1
End If
vbxCrystal.GroupCondition(0) = "GROUP1;" & grpField$ & ";ANYCHANGE;A"

' SECOND primary grouping ------------------------- **
GrpIdx% = comGroup(1).ListIndex
Select Case GrpIdx%
    Case 0: grpField$ = "{@EFullName}"
    Case 1: grpField$ = "{tblType.TB_DESC}"
    Case 2: grpField$ = SavGrp1
End Select
SavGrp2 = grpField$

Me.vbxCrystal.GroupCondition(1) = "GROUP2;" & grpField$ & ";ANYCHANGE;A"

If comGroup(1).ListIndex > 1 Then
    Me.vbxCrystal.SectionFormat(z%) = "GF2;F;X;X;X;X;X;X"
Else
    Me.vbxCrystal.SectionFormat(z%) = "GF2;T;X;X;X;X;X;X"
End If

z% = z% + 1

' 3TH primary grouping ------------------------- **8
GrpIdx% = comGroup(2).ListIndex
Select Case GrpIdx%
    Case 0: grpField$ = "{tblType.TB_DESC}"
    Case 1: grpField$ = SavGrp2
End Select
    Me.vbxCrystal.GroupCondition(2) = "GROUP3;" & grpField$ & ";ANYCHANGE;A"
    If comGroup(2).ListIndex > 0 Then
        Me.vbxCrystal.SectionFormat(z%) = "GF3;F;X;X;X;X;X;X"
    Else
        Me.vbxCrystal.SectionFormat(z%) = "GF3;T;X;X;X;X;X;X"
    End If

    Me.vbxCrystal.GroupCondition(3) = "GROUP4;" & "{HR_COUNSEL.CL_COUDATE}" & ";ANYCHANGE;A"
    'Me.vbxCrystal.GroupCondition(3) = "GROUP4;" & grpField$ & ";ANYCHANGE;A"
    
    If comGroup(2).ListIndex > 0 Then
        Me.vbxCrystal.SectionFormat(z% + 1) = "GF4;F;X;X;X;X;X;X"
    Else
        Me.vbxCrystal.SectionFormat(z% + 1) = "GF4;T;X;X;X;X;X;X"
    End If
z% = z% + 1
Me.vbxCrystal.SectionFormat(z%) = "GF3;F;X;X;X;X;X;X"
z% = z% + 1

' danielk - 04/10/2003 - don't know why this code is here, it's making the grouping not work
    'If comGroup(1).ListIndex = 0 Then
    '    Me.vbxCrystal.GroupCondition(2) = "GROUP3;{@EFullName};ANYCHANGE;A"
    '    Me.vbxCrystal.SectionFormat(2) = "GF3;T;X;X;X;X;X;X"
    '
    '    Me.vbxCrystal.GroupCondition(3) = "GROUP4;{HREMP.ED_COMPNO};ANYCHANGE;A"
    '    Me.vbxCrystal.SectionFormat(3) = "GF4;F;X;X;X;X;X;X"
    'End If
    'If comGroup(1).ListIndex = 1 Then
    '    Me.vbxCrystal.GroupCondition(2) = "GROUP3;{tblType.TB_DESC};ANYCHANGE;A"
    '    Me.vbxCrystal.SectionFormat(2) = "GF3;T;X;X;X;X;X;X"
    '
    '    Me.vbxCrystal.GroupCondition(3) = "GROUP4;{@EFullName};ANYCHANGE;A"
    '    Me.vbxCrystal.SectionFormat(3) = "GF4;F;X;X;X;X;X;X"
    'End If
    'If comGroup(1).ListIndex = 2 Then
    '    Me.vbxCrystal.GroupCondition(1) = "GROUP2;{HREMP.ED_COMPNO};ANYCHANGE;A"
    '    Me.vbxCrystal.SectionFormat(1) = "GF2;F;X;X;X;X;X;X"
    '
    '    Me.vbxCrystal.GroupCondition(2) = "GROUP3;{HREMP.ED_COMPNO};ANYCHANGE;A"
    '    Me.vbxCrystal.SectionFormat(2) = "GF3;F;X;X;X;X;X;X"
    '
    '    Me.vbxCrystal.GroupCondition(3) = "GROUP4;{@EFullName};ANYCHANGE;A"
    '    Me.vbxCrystal.SectionFormat(3) = "GF4;F;X;X;X;X;X;X"
    'End If
' danielk - 04/10/2003 - end


Cri_Sorts = z% ' next section number to format

End Function

Private Function CriCheck()
Dim X%

CriCheck = False

If Not clpDiv.ListChecker Then
'If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
    'MsgBox lStr("If Division Entered - it must be known")
    'clpDiv.SetFocus
    Exit Function
End If

If Not clpDept.ListChecker Then
'If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    'MsgBox "If Department Entered - it must be known"
    'clpDept.SetFocus
    Exit Function
End If


For X% = 0 To 7
    If Not clpCode(X).ListChecker Then Exit Function
Next X%

For X% = 0 To 1
 If Len(dlpDateRange(X%).Text) > 0 Then
    If Not IsDate(dlpDateRange(X%).Text) Then
        MsgBox "Not a valid date"
        dlpDateRange(X%).Text = ""
        dlpDateRange(X%).SetFocus
        Exit Function
    End If
 End If
Next X%

If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
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

glbOnTop = "FRMRCOUNSEL"

Screen.MousePointer = HOURGLASS

If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If

Call comGrpLoad
Call setRptCaption(Me)

If glbCompSerial = "S/N - 2227W" Then clpCode(5).MaxLength = 6

If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6

If glbLinamar Then
    clpCode(5).MaxLength = 8
End If

If glbBurlTech Then
    chkCDate.Visible = True
End If

Call INI_Controls(Me)

Screen.MousePointer = DEFAULT

If glbCompSerial = "S/N - 2376W" Then ' George added for Assembling of 1st Nations #9535
    lblFromTo.Caption = "From / To Issuing Date"
    lblReasonCode = "Item Code"
    clpCode(3).TABLTitle = "Item Codes"
    clpCode(4).TABLTitle = "Type Codes"
End If

frmRCounsel.Caption = lStr(frmRCounsel.Caption)
lblFromTo.Caption = "From / To " & lStr("Counseling") & " Date"

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

Private Sub txtShift_GotFocus()
Call SetPanHelp(ActiveControl)
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
