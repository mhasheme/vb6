VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmRHistory1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   7710
   ClientLeft      =   435
   ClientTop       =   870
   ClientWidth     =   10050
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7710
   ScaleWidth      =   10050
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkLastDay 
      Caption         =   "Show Last Day"
      Height          =   285
      Left            =   2160
      TabIndex        =   20
      Top             =   8100
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6030
      MaxLength       =   2
      TabIndex        =   19
      Tag             =   "00-Employee Position Shift"
      Top             =   7290
      Visible         =   0   'False
      Width           =   450
   End
   Begin INFOHR_Controls.CodeLookup clpJob 
      Height          =   285
      Left            =   1860
      TabIndex        =   7
      Tag             =   "00-Enter Position Code "
      Top             =   7650
      Visible         =   0   'False
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   6
      LookupType      =   5
   End
   Begin VB.Frame frmDesc 
      Caption         =   "Languages Description"
      Height          =   1485
      Left            =   8160
      TabIndex        =   40
      Top             =   6360
      Visible         =   0   'False
      Width           =   3195
      Begin VB.Label lblCodeDesc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Unassigned"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   5
         Left            =   300
         TabIndex        =   46
         Top             =   800
         Width           =   840
      End
      Begin VB.Label lblCodeDesc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Unassigned"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   4
         Left            =   300
         TabIndex        =   45
         Top             =   520
         Width           =   840
      End
      Begin VB.Label lblCodeDesc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Unassigned"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   6
         Left            =   300
         TabIndex        =   44
         Top             =   1100
         Width           =   840
      End
      Begin VB.Label lblCodeDesc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Unassigned"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   3
         Left            =   300
         TabIndex        =   43
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.ComboBox comGroup 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   2010
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Tag             =   "Final sorting of records"
      Top             =   6435
      Width           =   2325
   End
   Begin VB.ComboBox comGroup 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   2010
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Tag             =   "First Level of grouping records"
      Top             =   6120
      Width           =   2325
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1740
      TabIndex        =   4
      Tag             =   "00-Enter Status Code"
      Top             =   1680
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDEM"
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   1740
      TabIndex        =   5
      Tag             =   "EDPT-Category"
      Top             =   2010
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDPT"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   1740
      TabIndex        =   3
      Tag             =   "00-Enter Union Code"
      Top             =   1350
      Width           =   4005
      _ExtentX        =   7064
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
      Left            =   1740
      TabIndex        =   2
      Tag             =   "00-Enter Location Code"
      Top             =   1020
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1740
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   690
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
      Left            =   1740
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   360
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   8
      Left            =   1740
      TabIndex        =   12
      Tag             =   "00-Enter Administered By Code"
      Top             =   3300
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   9
      Left            =   1740
      TabIndex        =   13
      Tag             =   "00-Enter Section Code"
      Top             =   3630
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   7
      Left            =   1740
      TabIndex        =   11
      Tag             =   "00-Enter Region Code"
      Top             =   2970
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   3540
      TabIndex        =   10
      Tag             =   "40-Date upto and including this date forward"
      Top             =   2640
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   0
      Left            =   1740
      TabIndex        =   9
      Tag             =   "40-Date from and including this date forward"
      Top             =   2640
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1740
      TabIndex        =   6
      Tag             =   "10-Enter Employee Number"
      Top             =   2340
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.EmployeeLookup elpRept 
      Height          =   285
      Index           =   0
      Left            =   1860
      TabIndex        =   14
      Tag             =   "10-Reporting Authority 1"
      Top             =   6960
      Visible         =   0   'False
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   503
      ShowDescription =   0   'False
      RefreshDescriptionWhen=   2
   End
   Begin INFOHR_Controls.EmployeeLookup elpRept 
      Height          =   285
      Index           =   1
      Left            =   3690
      TabIndex        =   15
      Tag             =   "10-Reporting Authority 2"
      Top             =   6960
      Visible         =   0   'False
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   503
      RefreshDescriptionWhen=   2
   End
   Begin INFOHR_Controls.EmployeeLookup elpRept 
      Height          =   285
      Index           =   2
      Left            =   5700
      TabIndex        =   16
      Tag             =   "10-Reporting Authority 3"
      Top             =   6990
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      ShowDescription =   0   'False
      RefreshDescriptionWhen=   2
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   6660
      Top             =   6300
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
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   3
      Left            =   3690
      TabIndex        =   18
      Tag             =   "40-Date upto and including this date forward"
      Top             =   7305
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   2
      Left            =   1860
      TabIndex        =   17
      Tag             =   "40-Date from and including this date forward"
      Top             =   7305
      Visible         =   0   'False
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpGrid 
      Height          =   285
      Left            =   8160
      TabIndex        =   8
      Top             =   7650
      Visible         =   0   'False
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "JBGD"
   End
   Begin VB.Label lblGrid 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Grid Category"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6960
      TabIndex        =   48
      Top             =   7680
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label lblEmplStFrpmTo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Status From / To Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   47
      Top             =   7320
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label FName 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
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
      Left            =   4920
      TabIndex        =   42
      Top             =   6960
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label lblShift 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Shift"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5580
      TabIndex        =   41
      Top             =   7320
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
      TabIndex        =   39
      Top             =   2010
      Width           =   630
   End
   Begin VB.Label lblRep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reporting Authority:"
      Height          =   195
      Left            =   240
      TabIndex        =   38
      Top             =   6960
      Visible         =   0   'False
      Width           =   1395
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
      TabIndex        =   37
      Top             =   3630
      Width           =   540
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
      TabIndex        =   36
      Top             =   3300
      Width           =   1125
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
      TabIndex        =   35
      Top             =   2970
      Width           =   510
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
      TabIndex        =   34
      Top             =   1020
      Width           =   615
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
      TabIndex        =   33
      Top             =   6465
      Width           =   660
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
      TabIndex        =   32
      Top             =   6150
      Width           =   885
   End
   Begin VB.Label lblRepGrp 
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
      TabIndex        =   31
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label lblSelCri 
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
      TabIndex        =   30
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblFromTo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "From / To Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   29
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblPosition 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   28
      Top             =   7650
      Visible         =   0   'False
      Width           =   975
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
      TabIndex        =   27
      Top             =   2340
      Width           =   1290
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
      TabIndex        =   26
      Top             =   1680
      Width           =   450
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
      TabIndex        =   25
      Top             =   1350
      Width           =   420
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
      TabIndex        =   24
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
      TabIndex        =   23
      Top             =   360
      Width           =   555
   End
End
Attribute VB_Name = "frmRHistory1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReportSel, SQLQ
Dim RSTABL As New ADODB.Recordset
Dim rsJOB As New ADODB.Recordset


Private Sub clpCode_LostFocus(Index As Integer)
On Error Resume Next: lblCodeDesc(Index).Caption = clpCode(Index).Caption

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr

If CriCheck() Then
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
    Screen.MousePointer = DEFAULT
End If
Exit Sub

PrntErr:
MsgBox "CRW ERROR : " & Chr(10) & "[" & str(Err) & "] : " & Me.vbxCrystal.LastErrorString
Resume Next
Screen.MousePointer = DEFAULT
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
MsgBox "CRW ERROR : " & Chr(10) & "[" & str(Err) & "] : " & Me.vbxCrystal.LastErrorString
Resume Next
Screen.MousePointer = DEFAULT
End Sub


Private Sub comGroup_GotFocus(Index As Integer)
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()

comGroup(0).AddItem lStr("Division")
comGroup(0).AddItem lStr("Department")
comGroup(0).AddItem lStr("Location")  'Jaddy jun 16,1999
comGroup(0).AddItem lStr("Union")
comGroup(0).AddItem "Employee Name"
comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
If glbLinamar Then ' Frank May 2,2001
    comGroup(0).AddItem "Employment Type"
    comGroup(0).AddItem ("Home Line")
End If
If Not glbMulti Then comGroup(0).AddItem "Shift"
comGroup(0).AddItem lStr("Region")
comGroup(0).AddItem "(none)"

comGroup(0).ListIndex = 0
comGroup(1).AddItem "Employee Name"
comGroup(1).ListIndex = 0


End Sub



Private Sub Cri_Assoc()
Dim EECri As String
If Len(clpCode(1).Text) <= 0 Then Exit Sub

If ReportSel = "POS" Then
    If glbMulti And FormEmplPosition Then
        'EECri = "{HR_JOB_HISTORY.JH_ORG} = '" & clpCode(1).Text & "' "
        EECri = "{HR_JOB_HISTORY.JH_ORG} in  ['" & Replace(clpCode(1).Text, ",", "','") & "']"
    Else
        'EECri = "{HREMP.ED_ORG} = '" & clpCode(1).Text & "' "
        EECri = "{HREMP.ED_ORG} in  ['" & Replace(clpCode(1).Text, ",", "','") & "']"
    End If
    
    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
    
Else
    'EECri = "ED_ORG = '" & clpCode(1).Text & "' "
    'If glbSQL Or glbOracle Then
    '    EECri = "ED_ORG in  ['" & Replace(clpCode(1).Text, ",", "','") & "']"
    'Else
        EECri = "ED_ORG in  ('" & Replace(clpCode(1).Text, ",", "','") & "')"
    'End If
    If Len(SQLQ) > 1 Then
        SQLQ = SQLQ & " AND " & EECri
    Else
        SQLQ = EECri
    End If
End If


End Sub

Private Sub Cri_Dept()
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim DeptCri As String
DeptCri = ""
If ReportSel = "POS" Then
    Call glbCri_DeptUN(clpDept.Text)
Else
    
    If Len(clpDept.Text) > 0 Then
        DeptCri = glbSeleDeptUn & " AND (ED_DEPTNO = '" & clpDept.Text & "')"
    Else
        DeptCri = glbSeleDeptUn
    End If
    If Len(DeptCri) >= 1 Then
        If Len(SQLQ) > 1 Then
            SQLQ = SQLQ & " AND " & DeptCri
        Else
            SQLQ = DeptCri
        End If
    End If
End If

End Sub

Private Sub Cri_Div()
Dim DivCri As String
If Len(clpDiv.Text) <= 0 Then Exit Sub
If ReportSel = "POS" Then
    DivCri = "({HREMP.ED_DIV} = '" & clpDiv.Text & "')"
    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & DivCri
    Else
        glbstrSelCri = DivCri
    End If
Else
    DivCri = "(ED_DIV = '" & clpDiv.Text & "' )"
    If Len(SQLQ) > 1 Then
        SQLQ = SQLQ & " AND " & DivCri
    Else
        SQLQ = DivCri
    End If
End If

End Sub

Private Sub Cri_EE()
Dim EECri As String
If Len(elpEEID.Text) <= 0 Then Exit Sub
If ReportSel = "POS" Then
    EECri = "{HREMP.ED_EMPNBR} IN [" & getEmpnbr(elpEEID.Text) & "] "
    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
Else
    EECri = "ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
    If Len(SQLQ) > 1 Then
        SQLQ = SQLQ & " AND " & EECri
    Else
        SQLQ = EECri
    End If
End If

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
        If Len(glbstrSelCri) > 0 Then
          glbstrSelCri = glbstrSelCri & " AND " & EECri
        Else
          glbstrSelCri = EECri
        End If
    End If

End Sub


Private Sub Cri_FTDates()
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim x%
Dim EECri As String, LocCri As String

If Len(dlpDateRange(0).Text) = 0 And Len(dlpDateRange(1).Text) = 0 Then Exit Sub
If ReportSel = "PRO" Then
    If Len(dlpDateRange(0).Text) > 0 Then
        LocCri = "(ED_DOH >=" & Date_SQL(dlpDateRange(0).Text) & ")"
        If Len(SQLQ) > 1 Then
            SQLQ = SQLQ & " AND " & LocCri
        Else
            SQLQ = LocCri
        End If
    End If
    If Len(dlpDateRange(1).Text) > 0 Then
        LocCri = "(ED_DOH <=" & Date_SQL(dlpDateRange(1).Text) & ")"
        If Len(SQLQ) > 1 Then
            SQLQ = SQLQ & " AND " & LocCri
        Else
            SQLQ = LocCri
        End If
    End If
    Exit Sub
End If
If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    TempCri = "({HREMP.ED_DOH} "
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = Month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = Month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    GoTo Cri_FTDatst
ElseIf Len(dlpDateRange(0).Text) > 0 Then    ' Daniel - 10/20/1999
  TempCri = "({HREMP.ED_DOH} "         ' Added section to enable entering only From date, no To date.
  dtYYY% = Year(dlpDateRange(0).Text)
  dtMM% = Month(dlpDateRange(0).Text)
  dtDD% = Day(dlpDateRange(0).Text)
  TempCri = TempCri & " >= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "    'Hemu - 07/02/2003, Included '='
  GoTo Cri_FTDatst
ElseIf Len(dlpDateRange(1).Text) > 0 Then    ' Daniel - 10/20/1999
  TempCri = "({HREMP.ED_DOH} "         ' Added section to enable entering only To date, no From date.
  dtYYY% = Year(dlpDateRange(1).Text)
  dtMM% = Month(dlpDateRange(1).Text)
  dtDD% = Day(dlpDateRange(1).Text)
  TempCri = TempCri & " <= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "    'Hemu - 07/02/2003, Included '='
  GoTo Cri_FTDatst
End If

For x% = 0 To 1
    If Len(dlpDateRange(0).Text) > 0 Then
        TempCri = "({HREMP.ED_DOH}  "
        If x% = 0 Then
            TempCri = TempCri & " >= "
        Else
            TempCri = TempCri & " <= "
        End If
        dtYYY% = Year(dlpDateRange(0).Text)
        dtMM% = Month(dlpDateRange(0).Text)
        dtDD% = Day(dlpDateRange(0).Text)
        TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        GoTo Cri_FTDatst
    End If
Next x%



Cri_FTDatst:
' Daniel - 10/20/1999 - Changed code to enable date AND other criteria simultaneously.
If Len(TempCri) > 0 Then
    If Len(glbstrSelCri) > 0 Then
      glbstrSelCri = glbstrSelCri & " AND " & TempCri
    Else
      glbstrSelCri = TempCri
    End If
End If
End Sub

Private Sub Cri_Lang1()
Dim EECri As String, OneSet%, x%
Dim strCx  As String
Dim strCa$, strC2$

OneSet% = False
strCa$ = "HREMP.ED_LANG1"
strC2$ = "HREMP.ED_LANG2"


For x% = 3 To 6
    If Len(clpCode(x%).Text) > 0 Then
        OneSet% = OneSet% + 1
    End If
Next x%

If OneSet% = 0 Then
    EECri = EECri & "({" & strCa$ & "}<> '')"
    EECri = EECri & " OR " & "({" & strC2$ & "}) <> ''"
    EECri = EECri
    If glbstrSelCri <> "" Then
        glbstrSelCri = glbstrSelCri & " AND " & "(" & EECri & ")"
    Else
        glbstrSelCri = EECri
    End If
    glbiOneWhere = True
    Exit Sub
End If


For x% = 3 To 6
  If Len(clpCode(x%).Text) > 0 Then
    EECri = EECri & "({" & strCa$ & "} = '" & clpCode(x%).Text & "')"
    EECri = EECri & " OR "
  End If
Next x%

For x% = 3 To 6
  If Len(clpCode(x%).Text) > 0 Then
    EECri = EECri & "({" & strC2$ & "} = '" & clpCode(x%).Text & "')"
    OneSet% = OneSet% - 1
    If OneSet% > 0 Then
      EECri = EECri & " OR "
    Else
      EECri = EECri
    End If
  End If
Next x%

If Len(glbstrSelCri) > 1 Then
    glbstrSelCri = glbstrSelCri & " AND " & "(" & EECri & ")"
Else
    glbstrSelCri = EECri
End If


glbiOneWhere = True


End Sub

Private Sub Cri_Position()
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim PosCri As String
If Len(clpJob.Text) <= 0 Then Exit Sub
PosCri = "({HR_JOB_HISTORY.JH_JOB} = '" & clpJob.Text & "')"
If Len(glbstrSelCri) > 1 Then
    glbstrSelCri = glbstrSelCri & " AND " & PosCri
Else
    glbstrSelCri = PosCri
End If

End Sub
Private Sub Cri_Grid()
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim GirdCri As String
If Len(clpGrid.Text) <= 0 Then Exit Sub
GirdCri = "({HR_JOB_HISTORY.JH_GRID} = '" & clpGrid.Text & "')"
If Len(glbstrSelCri) > 1 Then
    glbstrSelCri = glbstrSelCri & " AND " & GirdCri
Else
    glbstrSelCri = GirdCri
End If

End Sub

Private Sub Cri_PT()
Dim EECri As String
If Len(clpPT.Text) < 1 Then Exit Sub
If ReportSel = "POS" Then
    If (glbMulti And FormEmplPosition) Or frmRPosition.Caption = "Category/Status Report" Then
        EECri = "{HR_JOB_HISTORY.JH_PT}= '" & clpPT.Text & "'"
    Else
        EECri = "{HREMP.ED_PT}= '" & clpPT.Text & "'"
    End If
    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
Else
    EECri = "ED_PT = '" & clpPT.Text & "' "
    If Len(SQLQ) > 1 Then
        SQLQ = SQLQ & " AND " & EECri
    Else
        SQLQ = EECri
    End If
End If
End Sub

Private Function Cri_SetAll()
Dim x%, strRName$

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""
SQLQ = ""

' call cri models set both glbiONeWhere and strSelCri
If ReportSel = "PRO" Then
            'Laura nov 3, 1997
    Call glbCri_DeptUN(clpDept.Text)
    SQLQ = glbstrSelCri
    Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
    Call Cri_Assoc
    Call Cri_Code(0)  'Jaddy jun 16,1999
    Call Cri_Code(7)  'Jaddy jun 16,1999
    Call Cri_Code(8)  'Jaddy jun 16,1999
    ' dkostka - 07/05/2001 - 'Section' selection criteria was being ignored, fixed.
    Call Cri_Code(9)
        
        Call Cri_Position

    Call Cri_Status
    Call Cri_PT
    Call Cri_Shift
    Call Cri_EE
    Call Cri_FTDates
    Call Cri_EmpStatFTDates
    Call EmpWrk
    x% = Cri_Sorts()
    glbstrSelCri = IIf(Len(glbstrSelCri) > 0, glbstrSelCri & " AND ", glbstrSelCri) & " {HREMPWRK.TT_WRKEMP}='" & glbUserID & "'"
    If Len(glbstrSelCri) >= 0 Then
        Me.vbxCrystal.SelectionFormula = glbstrSelCri
    End If
    strRName$ = glbIHRREPORTS & "rzProfil.rpt"  'Hemu - EMPHIS
    Me.vbxCrystal.ReportFileName = strRName$
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRDBW
        'Changed by Frank Apr 5,2002 for the 20533 error, "cannot open database"
        'If the Databases are not in as same folder as reports are
        'For For X% = 1 To 9
        For x% = 1 To 12
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next x%
    End If
    Me.vbxCrystal.WindowTitle = "Employee Profile Report"
    Exit Function
'ElseIf FormEmplPosition% = True Then    'laura nov 3, 1997
'        '~~~~   laura nov 3, 1997
'    'Call glbCri_Dept(Me)  'laura nov 22, 1997
'    Call glbCri_DeptUN(clpDept.Text)
'    Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
'    Call Cri_Assoc
'    Call Cri_Code(0)  'Jaddy jun 16,1999
'    Call Cri_Code(7)  'Jaddy jun 16,1999
'    Call Cri_Code(8)  'Jaddy jun 16,1999
'    Call Cri_Code(9)  'Frank Nov 18,2002 for Section was being ignored
'    Call Cri_Status
'    Call Cri_PT
'    Call Cri_Shift
'    Call Cri_EE
'        '~~~~~~
'    Call Cri_Position
'    Call Cri_Grid
'    Call Cri_FTDates    'laura 03/25/98
'    Call Cri_RepAuth
'    Call Cri_EmpStatFTDates
'        ' report name
'
'    If frmRPosition.Caption = "Category/Status Report" Then
'        If comGroup(0) = "(none)" Then
'            Me.vbxCrystal.Formulas(1) = "descGroup1 = '(none)'"
'            Me.vbxCrystal.GroupCondition(1) = "GROUP1;{HRPARCO.PC_CO};ANYCHANGE;A"
'        End If
'        strRName$ = glbIHRREPORTS & "sn2343.rpt"
'        Me.vbxCrystal.WindowTitle = "Category/Status Report"
'    Else
'        If comGroup(0) <> "(none)" Then
'            strRName$ = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "positn.rpt"
'            If glbMulti Then
'                Me.vbxCrystal.GroupCondition(3) = "GROUP3;{@EFullName};ANYCHANGE;A"
'            End If
'        Else
'            strRName$ = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "posit1.rpt"
'        End If
'
'
'        Me.vbxCrystal.WindowTitle = "Alphabetical List of Employee/Positions"
'    End If
'
'    Me.vbxCrystal.ReportFileName = strRName$
'        ' set to sorting/grouping criteria
'    x% = Cri_Sorts()   ' returns number of sections formated
'        'set location for database tables
'
'    If glbNoNONE And glbNoEXEC Then  'Hemu -EXE
'        If Len(glbstrSelCri) >= 0 Then
'            Me.vbxCrystal.SelectionFormula = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR ({HREMP.ED_ORG }<> 'NONE' AND {HREMP.ED_ORG }<> 'EXEC'))"
'        Else
'            Me.vbxCrystal.SelectionFormula = "(isnull({HREMP.ED_ORG }) OR ({HREMP.ED_ORG } <> 'NONE' AND {HREMP.ED_ORG }<> 'EXEC'))"
'        End If
'    ElseIf glbNoNONE Then
'        If Len(glbstrSelCri) >= 0 Then
'            Me.vbxCrystal.SelectionFormula = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG }<> 'NONE')"
'        Else
'            Me.vbxCrystal.SelectionFormula = "(isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG } <> 'NONE')"
'        End If
'    ElseIf glbNoEXEC Then    'Hemu -EXE
'        If Len(glbstrSelCri) >= 0 Then
'            Me.vbxCrystal.SelectionFormula = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG }<> 'EXEC')"    'Hemu -EXE
'        Else
'            Me.vbxCrystal.SelectionFormula = "(isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG } <> 'EXEC')"   'Hemu -EXE
'        End If
'    Else
'        If Len(glbstrSelCri) >= 0 Then
'            Me.vbxCrystal.SelectionFormula = glbstrSelCri
'        End If
'    End If
'
'    Me.vbxCrystal.Connect = RptODBC_SQL
    
End If
Cri_SetAll = True

Screen.MousePointer = DEFAULT
Exit Function


modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "FR Comp Time", "Comp Report", "Select")
Cri_SetAll = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim x%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%
'for labels - sort by name always
' imbeded in report

Cri_Sorts = 0
' first set primary grouping
z% = 0
x% = 0
grpField$ = getEGroup(comGroup(0).Text)
Y% = x% + 1
If ReportSel = "PRO" Then
    If comGroup(0) = "(none)" Then grpField$ = "{@EFullName}"
    Me.vbxCrystal.Formulas(9) = "showSIN = " & IIf(gSec_Show_SIN_SSN = 0, False, True) & " "
    Me.vbxCrystal.Formulas(35) = "showDOB = " & IIf(gSec_Show_DOB = 0, False, True) & " "
    Me.vbxCrystal.Formulas(36) = "showADDRESS = " & IIf(gSec_Show_ADDRESS = 0, False, True) & " "
    Call setRptLabel(Me, 1)
Else
'    If FormEmplPosition% = True Then
'        Me.vbxCrystal.Formulas(9) = "lblOHireDate='" & lStr("Original Hire Date") & "'"
'        Me.vbxCrystal.Formulas(38) = "lblPT='" & lStr("Category") & "'"
'        Call setRptLabel(Me, 0)
'    End If
    If comGroup(0) = "(none)" Then Exit Function
    dscGroup$ = comGroup(0).Text
    dscGroup$ = "descGroup1 = '" & dscGroup$ & "'"
    Me.vbxCrystal.Formulas(x%) = dscGroup$
End If
grpCond$ = "GROUP" & CStr(Y%) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(x%) = grpCond$

If ReportSel = "PRO" Then
    Exit Function
Else
    strVis$ = "T;"
    strFVis$ = "T;"
    strPage$ = "T;"
    strSFormat$ = "GH" & CStr(Y%) & ";" & strVis$ & strPage$ & "X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(z%) = strSFormat$
    z% = z% + 1
    strSFormat$ = "GF" & CStr(Y%) & ";" & strFVis$ & "X;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(z%) = strSFormat$
    z% = z% + 1
End If


Cri_Sorts = z% ' next section number to format

End Function

Private Sub Cri_Status()
Dim EECri As String, LocCri As String
If Len(clpCode(2).Text) <= 0 Then Exit Sub

If ReportSel = "PRO" Then
    LocCri = "(ED_EMP = '" & clpCode(2).Text & "' )"
    If Len(SQLQ) > 1 Then
        SQLQ = SQLQ & " AND " & LocCri
    Else
        SQLQ = LocCri
    End If
Else
    If Len(clpCode(2).Text) > 0 Then
        If frmRPosition.Caption = "Category/Status Report" Then
            EECri = "{HR_JOB_HISTORY.JH_EMP}= '" & clpCode(2).Text & "'"
        Else
            EECri = "{HREMP.ED_EMP} = '" & clpCode(2).Text & "' "
        End If
    End If
    If Len(EECri) >= 1 Then
        If Len(glbstrSelCri) > 1 Then '21July99-js-added check for other selection criteria
        'If glbiOneWhere Then         '           -commented out
            glbstrSelCri = glbstrSelCri & " AND " & EECri
        Else
            glbstrSelCri = EECri
        End If
        glbiOneWhere = True
    End If
End If
End Sub
Private Sub Cri_Code(intIdx%)
Dim CodeCri As String
Dim strCd$
If Len(clpCode(intIdx%).Text) > 0 Then
    If intIdx% = 0 Then strCd$ = "HREMP.ED_LOC"
    If intIdx% = 7 Then strCd$ = "HREMP.ED_REGION"
    If intIdx% = 8 Then strCd$ = "HREMP.ED_ADMINBY"
    If intIdx% = 9 Then strCd$ = "HREMP.ED_SECTION"  'Lucy July 4, 2000

    If ReportSel = "POS" Then
            CodeCri = "({" & strCd$ & "} in  ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
        If glbLinamar And (strCd$ = "HREMP.ED_REGION" Or strCd$ = "HREMP.ED_SECTION") Then
            CodeCri = "(({" & strCd$ & "} = '" & clpDiv.Text & clpCode(intIdx%).Text & "') or ({" & strCd$ & "} = 'ALL" & clpCode(intIdx%).Text & "') )"
        End If
        If Len(glbstrSelCri) > 1 Then
            glbstrSelCri = glbstrSelCri & " AND " & CodeCri
        Else
            glbstrSelCri = CodeCri
        End If
    Else
        CodeCri = "(" & strCd$ & " = '" & clpCode(intIdx%).Text & "')"
        If glbLinamar And (strCd$ = "HREMP.ED_REGION" Or strCd$ = "HREMP.ED_SECTION") Then
            CodeCri = "(" & strCd$ & " = '" & clpDiv.Text & clpCode(intIdx%).Text & "')"
        End If
        If Len(SQLQ) > 1 Then
            SQLQ = SQLQ & " AND " & CodeCri
        Else
            SQLQ = CodeCri
        End If
    End If
End If

End Sub

Private Function CriCheck()
Dim x%, I
CriCheck = False
If Me.Caption = "Employee Profile Report" Then
  ReportSel = "PRO"
Else
  ReportSel = "POS"
End If

If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
    MsgBox lStr("If Division Entered - it must be known")
     clpDiv.SetFocus
    Exit Function
End If

If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    MsgBox "If Department Entered - it must be known"
     clpDept.SetFocus
    Exit Function
End If


For x% = 0 To 9
    If Not clpCode(x).ListChecker Then Exit Function

Next x%

If Len(clpJob.Text) > 0 And clpJob.Caption = "Unassigned" Then
    MsgBox "If code entered it must be known"
     clpJob.SetFocus
    Exit Function
End If

If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    MsgBox lStr("Category code must be valid")
     clpPT.SetFocus
    Exit Function
End If

If ReportSel = "PRO" Then
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
End If

If frmRPosition.Caption = "Employee/Position Report" Then
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
End If
If IsDate(dlpDateRange(0)) And IsDate(dlpDateRange(1)) Then
    If DaysBetween(dlpDateRange(0), dlpDateRange(1)) < 0 Then                               'Serbo
        MsgBox "To Date can't be prior to From Date!"                       '
        Me.dlpDateRange(0).SetFocus                                         '
        Exit Function                                                       '
    End If
End If
If IsDate(dlpDateRange(2)) And IsDate(dlpDateRange(3)) Then
    If DaysBetween(dlpDateRange(2), dlpDateRange(3)) < 0 Then                               'Serbo
        MsgBox "To Date can't be prior to From Date!"                       '
        Me.dlpDateRange(2).SetFocus                                         '
        Exit Function                                                       '
    End If
End If


For I = 0 To 2
    If elpRept(I).Caption = "Enter Valid Employee #" Then
        MsgBox "If Reporting Authority Entered - they must exist"
        elpRept(I).SetFocus
        Exit Function
    End If
Next
If Not elpEEID.ListChecker Then
    Exit Function
End If

CriCheck = True
End Function



Private Sub EmpWrk()
Dim SQLX
Dim xEmplist
Dim xDate1, xDate2
Dim rsEMP As New ADODB.Recordset
On Error GoTo ERR_EmpWrk
gdbAdoIhr001.CommandTimeout = 300
gdbAdoIhr001W.CommandTimeout = 300
If Len(dlpDateRange(0).Text) = 0 Then
  xDate1 = DateAdd("yyyy", -100, Date)
Else
  xDate1 = dlpDateRange(0).Text
End If
If Len(dlpDateRange(1).Text) = 0 Then
  'xDate2 = DateAdd("yyyy", -100, Date)     'Jaddy 10/27/99
  xDate2 = DateAdd("yyyy", 50, Date)     'Jaddy 10/27/99
Else
  xDate2 = dlpDateRange(1).Text
End If
MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 0
gdbAdoIhr001.BeginTrans
SQLX = "DELETE FROM HREMPWRK " & in_SQL(glbIHRDBW) & "WHERE TT_WRKEMP='" & glbUserID & "'"
gdbAdoIhr001.Execute SQLX
gdbAdoIhr001.CommitTrans

SQLX = "SELECT ED_EMPNBR FROM HREMP "
If Len(clpJob.Text) > 0 Or Len(clpGrid.Text) > 0 Then
    If glbOracle Then
        SQLX = SQLX & ",HR_JOB_HISTORY WHERE HREMP.ED_EMPNBR(+)=HR_JOB_HISTORY.JH_EMPNBR "
    Else
        SQLX = SQLX & " LEFT JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR=HR_JOB_HISTORY.JH_EMPNBR WHERE (1=1)"
    End If
    
    If Len(SQLQ) > 1 Then
        SQLX = SQLX & " AND " & Replace(Replace(Replace(Replace(SQLQ, "{", ""), "}", ""), "[", "("), "]", ")")
    End If
    
    If Len(clpJob.Text) > 0 Then SQLX = SQLX & " AND JH_JOB='" & clpJob.Text & "'"
    If Len(clpGrid.Text) > 0 Then SQLX = SQLX & " AND JH_GRID='" & clpGrid.Text & "'"
    
Else
    If Len(SQLQ) > 1 Then
        SQLX = SQLX & " WHERE " & Replace(Replace(Replace(Replace(SQLQ, "{", ""), "}", ""), "[", "("), "]", ")")
    Else
        SQLX = SQLX & " WHERE (1=1)"
    End If
End If
If glbNoNONE Then
    SQLX = SQLX & " AND (ED_ORG IS NULL OR ED_ORG <> 'NONE') "
End If
If glbNoEXEC Then       'Hemu -EXE
    SQLX = SQLX & " AND (ED_ORG IS NULL OR ED_ORG <> 'EXEC') "  'Hemu -EXE
End If
rsEMP.Open SQLX, gdbAdoIhr001, adOpenStatic
If rsEMP.EOF And rsEMP.BOF Then
    GoTo rr
    Exit Sub
End If

MDIMain.panHelp(0).FloodPercent = 5
xEmplist = "(" & SQLX & ")"
'Do Until rsEMP.EOF
'    xEmplist = xEmplist & "," & rsEMP("ED_EMPNBR")
'    rsEMP.MoveNext
'    If rsEMP.EOF Then Exit Do
'Loop
'rsEMP.Close
'xEmplist = "(" & Mid(xEmplist, 2) & ")"
Call glbEmpWrk(xEmplist, xDate1, xDate2)
rr:
gdbAdoIhr001.CommandTimeout = 300
gdbAdoIhr001W.CommandTimeout = 300

MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""

Exit Sub
ERR_EmpWrk:
If Err = 13 Then
  FName.Visible = True
  MsgBox "SYSTEM ERROR : 13 - Type MisMatch"
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Create", "EMPWRK", "WORK FILE")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub



Private Sub Form_Activate()
Call SET_UP_MODE
Screen.MousePointer = HOURGLASS
If frmRPosition.Caption = "Category/Status Report" Then 'for  "S/N - 2343W"   'ottawa ccac
    lblPT.ForeColor = &HC000C0
    lblStatus.ForeColor = &HC000C0
End If
Screen.MousePointer = DEFAULT
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

'If FormLanguages% = True Then
'
'
'  lblPosition.Visible = False
'  clpJob.Visible = False
'
'  lblEmplStFrpmTo.Visible = False
'    dlpDateRange(0).Visible = False
'    dlpDateRange(1).Visible = False
'    dlpDateRange(2).Visible = False
'    dlpDateRange(3).Visible = False
'    lblFromTo.Visible = False
'
'Else
'    If glbMultiGrid Then
'        lblGrid.Visible = True
'        clpGrid.Visible = True
'    End If
'End If
'~~~~~~~~
If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If
Call setRptCaption(Me)
Call comGrpLoad
If glbLinamar Then clpCode(7).MaxLength = 8
If glbCompSerial = "S/N - 2227W" Then clpCode(7).MaxLength = 6
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
Set frmRPosition = Nothing  'carmen apr 2000
End Sub


Private Sub txtShift_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Cri_Shift()
Dim EECri As String, OneSet%, x%

If Len(txtShift.Text) < 1 Then Exit Sub

If ReportSel = "POS" Then

    EECri = "{HREMP.ED_SHIFT}= '" & txtShift.Text & "'"

    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
Else
    EECri = "ED_SHIFT = '" & txtShift.Text & "' "
    If Len(SQLQ) > 1 Then
        SQLQ = SQLQ & " AND " & EECri
    Else
        SQLQ = EECri
    End If
End If

glbiOneWhere = True
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


Private Sub Cri_EmpStatFTDates()
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%, x%
Dim FromDate, ToDate, SQLQ
Dim RsHRPARCO As New ADODB.Recordset

If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
    TempCri = "({HREMP.ED_SFDATE} "
    dtYYY% = Year(dlpDateRange(2).Text)
    dtMM% = Month(dlpDateRange(2).Text)
    dtDD% = Day(dlpDateRange(2).Text)
    TempCri = TempCri & " >= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) and "
    
    dtYYY% = Year(dlpDateRange(3).Text)
    dtMM% = Month(dlpDateRange(3).Text)
    dtDD% = Day(dlpDateRange(3).Text)
    TempCri = TempCri & " ({HREMP.ED_STDATE} <= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    GoTo Cri_FTDatst
End If

If Len(dlpDateRange(2).Text) > 0 Or Len(dlpDateRange(3).Text) > 0 Then
    If Len(dlpDateRange(2).Text) > 0 Then
        TempCri = "({HREMP.ED_SFDATE} "
        TempCri = TempCri & " >= "
        dtYYY% = Year(dlpDateRange(2).Text)
        dtMM% = Month(dlpDateRange(2).Text)
        dtDD% = Day(dlpDateRange(2).Text)
        TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        GoTo Cri_FTDatst
    End If
    If Len(dlpDateRange(3).Text) > 0 Then
        TempCri = TempCri & "({HREMP.ED_STDATE}  "
        TempCri = TempCri & " <= "
        dtYYY% = Year(dlpDateRange(3).Text)
        dtMM% = Month(dlpDateRange(3).Text)
        dtDD% = Day(dlpDateRange(3).Text)
        TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        GoTo Cri_FTDatst
    End If
Else
    GoTo Cri_FTDatst
End If



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

