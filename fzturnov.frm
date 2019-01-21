VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmRTurnov 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Employee Turnover"
   ClientHeight    =   7275
   ClientLeft      =   675
   ClientTop       =   1635
   ClientWidth     =   11145
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000040&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7275
   ScaleWidth      =   11145
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkExclCONP 
      Caption         =   "Exclude Employment Status of CONP"
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
      Left            =   120
      TabIndex        =   19
      Tag             =   "Check to Exclude Employees with CONP Employment Status"
      Top             =   6000
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.CheckBox chkExclRET 
      Caption         =   "Exclude Employment Status of RET"
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
      Left            =   120
      TabIndex        =   20
      Tag             =   "Check to Exclude Employees with RET Employment Status"
      Top             =   6300
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.Frame frmQuarter 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      TabIndex        =   36
      Top             =   5280
      Visible         =   0   'False
      Width           =   6975
      Begin VB.ComboBox comQuarter 
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
         ItemData        =   "fzturnov.frx":0000
         Left            =   840
         List            =   "fzturnov.frx":0002
         TabIndex        =   15
         Tag             =   "00-Country"
         Top             =   0
         Width           =   840
      End
      Begin VB.TextBox txtDFrom 
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
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   16
         Tag             =   "00-Employee Position Shift"
         Top             =   0
         Width           =   1170
      End
      Begin VB.TextBox txtTFrom 
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
         Left            =   3720
         MaxLength       =   10
         TabIndex        =   17
         Tag             =   "00-Employee Position Shift"
         Top             =   0
         Width           =   1170
      End
      Begin VB.CheckBox chkCurYear 
         Caption         =   "Current Year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   18
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lblQuarter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Quarter"
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
         Left            =   0
         TabIndex        =   39
         Top             =   0
         Width           =   645
      End
      Begin VB.Label lblDFrom 
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   38
         Top             =   0
         Width           =   495
      End
      Begin VB.Label lblTFrom 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   37
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.CheckBox chkShowEmp 
      Caption         =   "Show Employee List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   21
      Tag             =   "Check to show Employee Names"
      Top             =   4440
      Width           =   2055
   End
   Begin VB.TextBox txtShift 
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
      Height          =   285
      Left            =   1890
      MaxLength       =   4
      TabIndex        =   10
      Tag             =   "00-Employee Position Shift"
      Top             =   3680
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.TextBox txtYear 
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
      Height          =   285
      Left            =   960
      MaxLength       =   4
      TabIndex        =   12
      Tag             =   "61- Enter Year"
      Top             =   4425
      Width           =   855
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   4
      Tag             =   "00-Enter Status Code"
      Top             =   1680
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDEM"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Tag             =   "EDPT-Category"
      Top             =   2010
      Width           =   7515
      _ExtentX        =   13256
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
      Left            =   1560
      TabIndex        =   3
      Tag             =   "00-Enter Union Code"
      Top             =   1350
      Width           =   7515
      _ExtentX        =   13256
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
      Left            =   1560
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
      Left            =   1560
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   690
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
      Left            =   1560
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   360
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   1560
      TabIndex        =   8
      Tag             =   "00-Enter Section Code"
      Top             =   3000
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   1560
      TabIndex        =   7
      Tag             =   "00-Enter Administered By Code"
      Top             =   2670
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1560
      TabIndex        =   6
      Tag             =   "00-Enter Region Code"
      Top             =   2340
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Tag             =   "10-Enter Employee Number"
      Top             =   3330
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   9960
      Top             =   6840
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
   Begin INFOHR_Controls.CodeLookup clpJOB 
      Height          =   285
      Left            =   1560
      TabIndex        =   11
      Tag             =   "00-Enter Position Code"
      Top             =   4005
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "n/a"
      MaxLength       =   0
      LookupType      =   5
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.DateLookup dlpToDate 
      Height          =   285
      Left            =   3000
      TabIndex        =   14
      Tag             =   "40-Date upto and including this date"
      Top             =   4800
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpFromDate 
      Height          =   285
      Left            =   960
      TabIndex        =   13
      Tag             =   "40-Date from and including this date forward"
      Top             =   4800
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin VB.Label lblTDate 
      Caption         =   "To"
      Height          =   255
      Left            =   2640
      TabIndex        =   41
      Top             =   4815
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblFDate 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   4815
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblJOB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   35
      Top             =   4005
      Width           =   975
   End
   Begin VB.Label lblShift 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Shift"
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
      TabIndex        =   34
      Top             =   3675
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label lblPT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
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
      TabIndex        =   33
      Top             =   2010
      Width           =   630
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
      Index           =   0
      Left            =   120
      TabIndex        =   32
      Top             =   3330
      Width           =   1290
   End
   Begin VB.Label lblSection 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
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
      TabIndex        =   31
      Top             =   3030
      Width           =   540
   End
   Begin VB.Label lblAdmin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Administered By"
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
      TabIndex        =   30
      Top             =   2670
      Width           =   1125
   End
   Begin VB.Label lblRegion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
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
      TabIndex        =   29
      Top             =   2340
      Width           =   510
   End
   Begin VB.Label lblLocation 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
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
      TabIndex        =   28
      Top             =   1020
      Width           =   615
   End
   Begin VB.Label lblYear 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   4455
      Width           =   495
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
      TabIndex        =   26
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
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
      TabIndex        =   25
      Top             =   1680
      Width           =   450
   End
   Begin VB.Label lblUnion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Union"
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
      TabIndex        =   24
      Top             =   1350
      Width           =   420
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
      TabIndex        =   23
      Top             =   690
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
      TabIndex        =   22
      Top             =   360
      Width           =   555
   End
End
Attribute VB_Name = "frmRTurnov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ReportSel, SQLQ

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr

If CriCheck() Then
  If FormEmplPosition% = True Then
    If Not PrtForm("Employee/Position Report Criteria", Me) Then Exit Sub
  ElseIf FormLanguages% = True Then
    If Not PrtForm("Languages Report Criteria", Me) Then Exit Sub    'laura nov 3, 1997
  Else
  End If
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    Call set_PrintState(False)
    x% = Cri_SetAll()
    Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    'cmdPrint.Enabled = True
    'cmdView.Enabled = True
      Call set_PrintState(True)
    Screen.MousePointer = DEFAULT
End If
Exit Sub

PrntErr:
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString
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
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString
Resume Next
Screen.MousePointer = DEFAULT
End Sub

Private Function Cri_SetAll()
Dim x%, strRName$
Dim xDateRange
Dim xQuarter

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = True
glbstrSelCri = " {Employee_Turnover.WRKEMP}='" & glbUserID & "'"
SQLQ = ""

Call Cri_Year

If chkShowEmp Then Me.vbxCrystal.Formulas(0) = "lblEmp='Employee Nb./Name'"

If glbCompSerial <> "S/N - 2214W" Then 'Casey House then
    xDateRange = "(" & dlpFromDate.Text & " - " & dlpToDate.Text & ")"
    Me.vbxCrystal.Formulas(1) = "DateRange='" & xDateRange & "'"
End If

If glbCompSerial = "S/N - 2214W" Then 'Casey House
    If comQuarter.Text <> "" Then
        xQuarter = "Quarter: " & comQuarter.Text
        Me.vbxCrystal.Formulas(2) = "quarter='" & xQuarter & "'"
    End If
    xDateRange = "(" & txtDFrom.Text & " - " & txtTFrom.Text & ")"
    Me.vbxCrystal.Formulas(1) = "DateRange='" & xDateRange & "'"
End If

' report name
If glbCompSerial = "S/N - 2443W" Then 'Walters Inc Ticket #26271 Franks 11/13/2014
    strRName$ = glbIHRREPORTS & "rzemptrn2.rpt"
    vbxCrystal.Formulas(80) = "lblRegion='" & lStr("Region") & "'"
Else
    strRName$ = glbIHRREPORTS & "rzemptrn.rpt"
End If

Call setRptLabel(Me, 0)
Me.vbxCrystal.ReportFileName = strRName$
  
'Ticket #29660 - Contract Employees Enhancement
If glbWFC Then
    If chkExclCONP.Visible And chkExclRET.Visible = True Then
        If chkExclCONP Then
            If Len(glbstrSelCri) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND {HREMP.ED_EMP} <> 'CONP'"
            Else
                glbstrSelCri = "{HREMP.ED_EMP} <> 'CONP'"
            End If
        End If
        If chkExclRET Then
            If Len(glbstrSelCri) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND {HREMP.ED_EMP} <> 'RET'"
            Else
                glbstrSelCri = "{HREMP.ED_EMP} <> 'RET'"
            End If
        End If
    End If
End If
  
'set location for database tables
If Len(glbstrSelCri) >= 0 Then
  Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    For x% = 0 To 2
      Me.vbxCrystal.DataFiles(x%) = glbIHRDB
    Next x%
    
    Me.vbxCrystal.DataFiles(3) = glbIHRDBW
    
    For x% = 4 To 5
      Me.vbxCrystal.DataFiles(x%) = glbIHRDB
    Next x%

 
  ' set security for database
   'Me.vbxCrystal.Password = gstrAccPWord$
   'Me.vbxCrystal.UserName = gstrAccUID$
End If
  ' window title if appropriate
Me.vbxCrystal.WindowTitle = "List of Employee Turnovers"


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

Private Sub Cri_Year()
Dim SQLF, SQLW, sqlI, SQLQ
Dim Tdate As Variant
Dim Fdate As Variant

If Trim(txtYear.Text) <> "" Then
    Fdate = GetMonth("Jan") & " 1, " & txtYear
    Tdate = IIf(glbFrench, "déc", "Dec") & " 31, " & txtYear
End If

If glbCompSerial = "S/N - 2214W" Then 'Casey House
    Fdate = txtDFrom
    Tdate = txtTFrom
Else
    If IsDate(dlpFromDate.Text) And IsDate(dlpToDate.Text) Then
        Fdate = dlpFromDate.Text
        Tdate = dlpToDate.Text
        
        If txtYear.Text = "" Then
            txtYear.Text = Year(dlpFromDate.Text)
        End If
    ElseIf txtYear.Text <> "" Then
        dlpFromDate.Text = "01/01/" & txtYear
        dlpToDate.Text = "12/31/" & txtYear
    End If
End If
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Screen.MousePointer = 11

gdbAdoIhr001W.BeginTrans
gdbAdoIhr001W.Execute "DELETE FROM Employee_Turnover WHERE WRKEMP='" & glbUserID & "'"
gdbAdoIhr001W.CommitTrans

SQLW = glbSeleDeptUn
If clpDiv.Text <> "" Then SQLW = SQLW & " AND ED_DIV in ('" & Replace(clpDiv.Text, ",", "','") & "')"
If clpDept.Text <> "" Then SQLW = SQLW & " AND ED_DEPTNO in ('" & Replace(clpDept.Text, ",", "','") & "')"
If clpCode(0).Text <> "" Then SQLW = SQLW & " AND ED_LOC='" & clpCode(0).Text & "'"

'If clpCode(1).Text <> "" Then SQLW = SQLW & " AND ED_ORG='" & clpCode(1).Text & "'"
If clpCode(1).Text <> "" Then SQLW = SQLW & " AND ED_ORG IN ('" & Replace(clpCode(1).Text, ",", "','") & "')"

If clpCode(2).Text <> "" Then SQLW = SQLW & " AND ED_EMP in ('" & Replace(clpCode(2).Text, ",", "','") & "')"
If clpPT.Text <> "" Then SQLW = SQLW & " AND ED_PT in ('" & Replace(clpPT.Text, ",", "','") & "')"
If txtShift.Text <> "" Then SQLW = SQLW & " AND ED_SHIFT='" & txtShift.Text & "'"
If clpCode(3).Text <> "" Then SQLW = SQLW & " AND ED_REGION='" & IIf(glbLinamar, clpDiv.Text, "") & clpCode(3).Text & "'"
If clpCode(4).Text <> "" Then SQLW = SQLW & " AND ED_ADMINBY='" & clpCode(4).Text & "'"
If clpCode(5).Text <> "" Then SQLW = SQLW & " AND ED_SECTION='" & IIf(glbLinamar, clpDiv.Text, "") & clpCode(5).Text & "'"
If Len(elpEEID.Text) > 0 Then SQLW = SQLW & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "

'Position Codes - Begin Ticket# 10209
If Len(clpJob.Text) > 0 Then
    Dim JobCodeCri
    If glbOracle Then
        JobCodeCri = "['" & getCodes(clpJob.Text) & "']"
    Else
        JobCodeCri = "('" & getCodes(clpJob.Text) & "')"
    End If
    SQLW = SQLW & " AND ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_JOB IN " & JobCodeCri & ") "
End If
'Position Codes - End

If glbCompSerial = "S/N - 2443W" Then 'Walters Inc Ticket #26271 Franks 11/13/2014 use ED_REGION as dept
    SQLW = SQLW & " AND NOT (ED_REGION IS NULL) "
End If

'Adding New Hires count in Active table
sqlI = " INSERT INTO Employee_Turnover(company, yr, division, dept, empnbr," & Field_SQL("name") & ",tot_emp, female, male,newhires, terms, WRKEMP) "
sqlI = sqlI & in_SQL(glbIHRDBW)

If glbCompSerial = "S/N - 2443W" Then 'Walters Inc Ticket #26271 Franks 11/13/2014 use ED_REGION as dept
    SQLF = "SELECT ED_COMPNO AS company," & txtYear & " AS yr,ED_DIV AS division,ED_REGION AS dept,"
Else
    SQLF = "SELECT ED_COMPNO AS company," & txtYear & " AS yr,ED_DIV AS division,ED_DEPTNO AS dept,"
End If
If glbOracle Then
    SQLF = SQLF & "ED_EMPNBR AS empnbr,(ED_SURNAME||', '||ED_FNAME) as " & Field_SQL("name") & ",1 AS tot_emp,"
Else
    SQLF = SQLF & "ED_EMPNBR AS empnbr,LEFT((ED_SURNAME+', '+ED_FNAME),30) as " & Field_SQL("name") & ",1 AS tot_emp,"
End If
If glbSQL Or glbOracle Then
    SQLF = SQLF & "(CASE WHEN ED_SEX='F' THEN 1 ELSE 0 END) AS female, (CASE WHEN ED_SEX='M' THEN 1 ELSE 0 END ) AS MALE,"
Else
    SQLF = SQLF & "IIF(ED_SEX='F',1,0) AS FEMALE, IIF(ED_SEX='M',1,0) AS male,"
End If

SQLQ = sqlI & SQLF
If glbSQL Or glbOracle Then
    SQLQ = SQLQ & "(CASE WHEN ED_DOH>=" & Date_SQL(Fdate) & " AND ED_DOH<=" & Date_SQL(Tdate) & " THEN 1 ELSE 0 END) AS newhires,"
Else
    SQLQ = SQLQ & "iif(ED_DOH>=" & Date_SQL(Fdate) & " AND ED_DOH<=" & Date_SQL(Tdate) & " ,1 ,0 ) AS newhires,"
End If
SQLQ = SQLQ & "0 AS terms "
SQLQ = SQLQ & ",'" & glbUserID & "' AS WRKEMP "
SQLQ = SQLQ & "FROM HREMP "
SQLQ = SQLQ & "WHERE " & SQLW
SQLQ = SQLQ & " AND ED_DOH<=" & Date_SQL(Tdate)

gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans


'Adding Employees who were Activee but now in Term table
sqlI = " INSERT INTO Employee_Turnover(company, yr, division, dept, empnbr," & Field_SQL("name") & ",tot_emp, female, male,newhires, terms, WRKEMP) "
sqlI = sqlI & in_SQL(glbIHRDBW)

If glbCompSerial = "S/N - 2443W" Then 'Walters Inc Ticket #26271 Franks 11/13/2014 use ED_REGION as dept
    SQLF = "SELECT ED_COMPNO AS company," & txtYear & " AS yr,ED_DIV AS division,ED_REGION AS dept,"
Else
    SQLF = "SELECT ED_COMPNO AS company," & txtYear & " AS yr,ED_DIV AS division,ED_DEPTNO AS dept,"
End If
If glbOracle Then
    SQLF = SQLF & "ED_EMPNBR AS empnbr,(ED_SURNAME||', '||ED_FNAME) as " & Field_SQL("name") & ",1 AS tot_emp,"
Else
    SQLF = SQLF & "ED_EMPNBR AS empnbr,LEFT((ED_SURNAME+', '+ED_FNAME),30) as " & Field_SQL("name") & ",1 AS tot_emp,"
End If
If glbSQL Or glbOracle Then
    SQLF = SQLF & "(CASE WHEN ED_SEX='F' THEN 1 ELSE 0 END) AS female, (CASE WHEN ED_SEX='M' THEN 1 ELSE 0 END ) AS MALE,"
Else
    SQLF = SQLF & "IIF(ED_SEX='F',1,0) AS FEMALE, IIF(ED_SEX='M',1,0) AS male,"
End If

SQLQ = sqlI & SQLF
If glbSQL Or glbOracle Then
    SQLQ = SQLQ & "(CASE WHEN ED_DOH>=" & Date_SQL(Fdate) & " AND ED_DOH<=" & Date_SQL(Tdate) & " THEN 1 ELSE 0 END) AS newhires,"
Else
    SQLQ = SQLQ & "iif(ED_DOH>=" & Date_SQL(Fdate) & " AND ED_DOH<=" & Date_SQL(Tdate) & " ,1 ,0 ) AS newhires,"
End If
SQLQ = SQLQ & "0 AS terms "
SQLQ = SQLQ & ",'" & glbUserID & "' AS WRKEMP "
If glbOracle Then
    SQLQ = SQLQ & " FROM Term_HRTRMEMP, Term_HREMP WHERE Term_HRTRMEMP.TERM_SEQ=Term_HREMP.TERM_SEQ "
    SQLQ = SQLQ & " AND " & SQLW
Else
    SQLQ = SQLQ & " FROM Term_HRTRMEMP INNER JOIN Term_HREMP ON Term_HRTRMEMP.TERM_SEQ=Term_HREMP.TERM_SEQ "
    SQLQ = SQLQ & " WHERE " & SQLW
End If
SQLQ = SQLQ & " AND ED_DOH<=" & Date_SQL(Tdate)             'Still hired before the To Date
SQLQ = SQLQ & " AND Term_DOT>" & Date_SQL(Tdate)            'Employees terminated after the To Date - this is to get the list of employees who were still active
SQLQ = SQLQ & " AND (Term_DOR is null or Term_DOR ='')"     'Who have not been already rehired as they will appear in the above recordset of Active table Active Employees

gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans

If glbCompSerial = "S/N - 2214W" Then 'Casey House
    If chkCurYear Then
        Call CaseyCurYearTerm
    End If
End If

'Adding New Hires count who are now Terminated in Term table
SQLQ = sqlI & SQLF
If glbSQL Or glbOracle Then
    SQLQ = SQLQ & "(CASE WHEN ED_DOH>=" & Date_SQL(Fdate) & " AND ED_DOH<=" & Date_SQL(Tdate) & " THEN 1 ELSE 0 END) AS newhires,"
Else
    SQLQ = SQLQ & "iif(ED_DOH>=" & Date_SQL(Fdate) & " AND ED_DOH<=" & Date_SQL(Tdate) & " ,1 ,0 ) AS newhires,"
End If
SQLQ = SQLQ & "0 AS terms "
SQLQ = SQLQ & ",'" & glbUserID & "' AS WRKEMP "
If glbOracle Then
    SQLQ = SQLQ & " FROM Term_HRTRMEMP, Term_HREMP WHERE Term_HRTRMEMP.TERM_SEQ=Term_HREMP.TERM_SEQ "
    SQLQ = SQLQ & " AND " & SQLW
Else
    SQLQ = SQLQ & " FROM Term_HRTRMEMP INNER JOIN Term_HREMP ON Term_HRTRMEMP.TERM_SEQ=Term_HREMP.TERM_SEQ "
    SQLQ = SQLQ & " WHERE " & SQLW
End If


SQLQ = SQLQ & " AND ED_EMPNBR NOT IN (SELECT empnbr FROM Employee_Turnover)"


'Ticket #28551 - To include terminated employees who were terminated between the date range
'SQLQ = SQLQ & " AND ED_DOH<=" & Date_SQL(Tdate)
'Ticket #28551 - Adding the start ED_DOH range as well as employee hired outside the period were showing up
'SQLQ = SQLQ & " AND (ED_DOH<=" & Date_SQL(Tdate)
SQLQ = SQLQ & " AND (ED_DOH>=" & Date_SQL(Fdate) & " AND ED_DOH<=" & Date_SQL(Tdate) & ")"
'Ticket #28551 - The following line is causing an issue where if a same employee is hired and terminated within same date range, it was
'excluding the employee from the New Hire count.
'SQLQ = SQLQ & " AND Term_DOT>" & Date_SQL(Tdate)
'Ticket #28551 - To include terminated employees who were terminated between the date range
'Ticket #28551 - After adding the start ED_DOH range above, if hired and terminated in the same month and year were not showing up so added the OR clause and also
'                with that I had to add the whole selection criteria in SQLW as well.
'SQLQ = SQLQ & " AND Term_DOT>=" & Date_SQL(Fdate) & " AND Term_DOT<=" & Date_SQL(Tdate) & ")"
SQLQ = SQLQ & " OR ( " & SQLW & " AND Term_DOT>=" & Date_SQL(Fdate) & " AND Term_DOT<=" & Date_SQL(Tdate) & ")"

SQLQ = Replace(SQLQ, "HR_JOB_HISTORY", "Term_JOB_HISTORY")
gdbAdoIhr001X.BeginTrans
gdbAdoIhr001X.Execute SQLQ
gdbAdoIhr001X.CommitTrans

'Add Terms count from the Term table
Call Terminated_Count(Fdate, Tdate, SQLW)
'Commenting the code below because I added the above call which is doing what the commented code below does but little
'differently.
'SQLQ = sqlI & SQLF
'SQLQ = SQLQ & "0 AS newhires, 1 as terms "
'SQLQ = SQLQ & ",'" & glbUserID & "' AS WRKEMP "
'If glbOracle Then
'    SQLQ = SQLQ & " FROM Term_HRTRMEMP, Term_HREMP WHERE Term_HRTRMEMP.TERM_SEQ=Term_HREMP.TERM_SEQ "
'    SQLQ = SQLQ & " AND " & SQLW
'Else
'    SQLQ = SQLQ & " FROM Term_HRTRMEMP INNER JOIN Term_HREMP ON Term_HRTRMEMP.TERM_SEQ=Term_HREMP.TERM_SEQ "
'    SQLQ = SQLQ & " WHERE " & SQLW
'End If
'SQLQ = SQLQ & " AND Term_DOT>=" & Date_SQL(Fdate)
'SQLQ = SQLQ & " AND Term_DOT<=" & Date_SQL(Tdate)
'
'SQLQ = Replace(SQLQ, "HR_JOB_HISTORY", "Term_JOB_HISTORY")
'gdbAdoIhr001X.BeginTrans
'gdbAdoIhr001X.Execute SQLQ
'gdbAdoIhr001X.CommitTrans

End Sub

Private Sub CaseyCurYearTerm()
Dim rsCEmp As New ADODB.Recordset
Dim rsCWrk As New ADODB.Recordset
Dim xStr As String
    xStr = "SELECT * FROM Employee_Turnover WHERE WRKEMP='" & glbUserID & "'"
    rsCWrk.Open xStr, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsCWrk.EOF
        xStr = "SELECT ED_EMPNBR,ED_PT FROM HREMP WHERE ED_PT='TR' AND ED_EMPNBR=" & rsCWrk("empnbr") & " "
        rsCEmp.Open xStr, gdbAdoIhr001, adOpenStatic
        If Not rsCEmp.EOF Then
            rsCWrk("terms") = 1
            rsCWrk.Update
        End If
        rsCEmp.Close
        rsCWrk.MoveNext
    Loop
    rsCWrk.Close
End Sub

Private Sub Terminated_Count(Fdate, Tdate, SQLW)
Dim rsTerm As New ADODB.Recordset
Dim rsEmpTrnOWrk As New ADODB.Recordset
Dim SQLQ As String

    If glbOracle Then
        SQLQ = "SELECT * FROM Term_HRTRMEMP, Term_HREMP WHERE Term_HRTRMEMP.TERM_SEQ=Term_HREMP.TERM_SEQ "
        SQLQ = SQLQ & " AND " & SQLW
    Else
        SQLQ = "SELECT * FROM Term_HRTRMEMP INNER JOIN Term_HREMP ON Term_HRTRMEMP.TERM_SEQ=Term_HREMP.TERM_SEQ "
        SQLQ = SQLQ & " WHERE " & SQLW
    End If
    SQLQ = SQLQ & " AND Term_DOT>=" & Date_SQL(Fdate)
    SQLQ = SQLQ & " AND Term_DOT<=" & Date_SQL(Tdate)
    
    SQLQ = Replace(SQLQ, "HR_JOB_HISTORY", "Term_JOB_HISTORY")
    
    rsTerm.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsTerm.EOF
        SQLQ = "SELECT * FROM Employee_Turnover WHERE WRKEMP='" & glbUserID & "'"
        SQLQ = SQLQ & " AND empnbr = " & rsTerm("ED_EMPNBR")
        rsEmpTrnOWrk.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsEmpTrnOWrk.EOF Then
            rsEmpTrnOWrk("terms") = 1
            rsEmpTrnOWrk.Update
        Else
            rsEmpTrnOWrk.AddNew
            rsEmpTrnOWrk("company") = rsTerm("ED_COMPNO")
            rsEmpTrnOWrk("yr") = txtYear
            rsEmpTrnOWrk("division") = rsTerm("ED_DIV")
            If glbCompSerial = "S/N - 2443W" Then 'Walters Inc Ticket #26271 Franks 11/13/2014 use ED_REGION as dept
                rsEmpTrnOWrk("dept") = rsTerm("ED_REGION")
            Else
                rsEmpTrnOWrk("dept") = rsTerm("ED_DEPTNO")
            End If
            rsEmpTrnOWrk("empnbr") = rsTerm("ED_EMPNBR")
            rsEmpTrnOWrk("name") = Left(rsTerm("ED_SURNAME") & ", " & rsTerm("ED_FNAME"), 30)
            rsEmpTrnOWrk("tot_emp") = 1
            rsEmpTrnOWrk("female") = IIf(rsTerm("ED_SEX") = "F", 1, 0)
            rsEmpTrnOWrk("male") = IIf(rsTerm("ED_SEX") = "M", 1, 0)
            rsEmpTrnOWrk("newhires") = 0
            rsEmpTrnOWrk("terms") = 1
            rsEmpTrnOWrk("WRKEMP") = glbUserID
            rsEmpTrnOWrk.Update
        End If
        rsEmpTrnOWrk.Close
        Set rsEmpTrnOWrk = Nothing
        
        rsTerm.MoveNext
    Loop
    rsTerm.Close
    Set rsTerm = Nothing
    
'sqlI = "INSERT INTO Employee_Turnover(company, yr, division, dept, empnbr," & Field_SQL("name") & ",tot_emp, female, male,newhires, terms, WRKEMP) "
'sqlI = sqlI & in_SQL(glbIHRDBW)
'
'If glbCompSerial = "S/N - 2443W" Then 'Walters Inc Ticket #26271 Franks 11/13/2014 use ED_REGION as dept
'    SQLF = "SELECT ED_COMPNO AS company," & txtYear & " AS yr,ED_DIV AS division,ED_REGION AS dept,"
'Else
'    SQLF = "SELECT ED_COMPNO AS company," & txtYear & " AS yr,ED_DIV AS division,ED_DEPTNO AS dept,"
'End If
'If glbOracle Then
'    SQLF = SQLF & "ED_EMPNBR AS empnbr,(ED_SURNAME||', '||ED_FNAME) as " & Field_SQL("name") & ",1 AS tot_emp,"
'Else
'    SQLF = SQLF & "ED_EMPNBR AS empnbr,LEFT((ED_SURNAME+', '+ED_FNAME),30) as " & Field_SQL("name") & ",1 AS tot_emp,"
'End If
'If glbSQL Or glbOracle Then
'    SQLF = SQLF & "(CASE WHEN ED_SEX='F' THEN 1 ELSE 0 END) AS female, (CASE WHEN ED_SEX='M' THEN 1 ELSE 0 END ) AS MALE,"
'Else
'    SQLF = SQLF & "IIF(ED_SEX='F',1,0) AS FEMALE, IIF(ED_SEX='M',1,0) AS male,"
'End If
'
'SQLQ = sqlI & SQLF
'SQLQ = SQLQ & "0 AS newhires, 1 as terms "
'SQLQ = SQLQ & ",'" & glbUserID & "' AS WRKEMP "
'If glbOracle Then
'    SQLQ = SQLQ & " FROM Term_HRTRMEMP, Term_HREMP WHERE Term_HRTRMEMP.TERM_SEQ=Term_HREMP.TERM_SEQ "
'    SQLQ = SQLQ & " AND " & SQLW
'Else
'    SQLQ = SQLQ & " FROM Term_HRTRMEMP INNER JOIN Term_HREMP ON Term_HRTRMEMP.TERM_SEQ=Term_HREMP.TERM_SEQ "
'    SQLQ = SQLQ & " WHERE " & SQLW
'End If
'SQLQ = SQLQ & " AND Term_DOT>=" & Date_SQL(Fdate)
'SQLQ = SQLQ & " AND Term_DOT<=" & Date_SQL(Tdate)
'
'SQLQ = Replace(SQLQ, "HR_JOB_HISTORY", "Term_JOB_HISTORY")

End Sub

Private Function CriCheck()
Dim x%

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

For x% = 0 To 5
    If Not clpCode(x).ListChecker Then Exit Function
Next x%

If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

If txtYear = "" And (Trim(dlpFromDate.Text) = "" And Trim(dlpToDate.Text) = "") Then
    MsgBox "You have to enter a year or the date range!"
    txtYear = ""
    txtYear.SetFocus
    Exit Function
Else
    If (Trim(dlpFromDate.Text) <> "" And Trim(dlpToDate.Text) <> "") Then
        If Not IsDate(dlpFromDate.Text) Then
            MsgBox "Invalid From Date"
            dlpFromDate.SetFocus
            Exit Function
        ElseIf Not IsDate(dlpToDate.Text) Then
            MsgBox "Invalid To Date!"
            dlpToDate.SetFocus
            Exit Function
        End If
    ElseIf Trim(dlpFromDate.Text) = "" And Trim(dlpToDate.Text) <> "" Then
        MsgBox "Enter both the From Date and To Date range"
        dlpFromDate.SetFocus
        Exit Function
    ElseIf Trim(dlpFromDate.Text) <> "" And Trim(dlpToDate.Text) = "" Then
        MsgBox "Enter both the From Date and To Date range"
        dlpToDate.SetFocus
        Exit Function
    Else
        'Frank May 12,2003
        If Not IsNumeric(txtYear) Then
            MsgBox "Invalid year!"
            txtYear = ""
            txtYear.SetFocus
            Exit Function
        Else
            If Len(txtYear) <> 4 Then
                MsgBox "Invalid year!"
                txtYear = ""
                txtYear.SetFocus
                Exit Function
            End If
        End If
    End If
End If

If Not elpEEID.ListChecker Then
    Exit Function
End If

CriCheck = True
End Function

Private Sub chkExclCONP_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkExclRET_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkShowEmp_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comQuarter_Click()
    If txtYear = "" Then
        MsgBox "You have to enter a year!"
        txtYear = ""
        txtYear.SetFocus
        Exit Sub
    Else
        If Not IsNumeric(txtYear) Then
            MsgBox "Invalid year!"
            txtYear = ""
            txtYear.SetFocus
            Exit Sub
        Else
            If Len(txtYear) <> 4 Then
                MsgBox "Invalid year!"
                txtYear = ""
                txtYear.SetFocus
                Exit Sub
            End If
        End If
    End If
    If comQuarter = "" Then
        txtDFrom = CVDate(GetMonth("Apr") & " 1," & txtYear)
        txtTFrom = DateAdd("m", 12, txtDFrom)
        txtTFrom = DateAdd("d", -1, txtTFrom)
    End If
    If comQuarter = "Q1" Then
        txtDFrom = CVDate(GetMonth("Apr") & " 1," & txtYear)
        txtTFrom = DateAdd("m", 3, txtDFrom)
        txtTFrom = DateAdd("d", -1, txtTFrom)
    End If
    If comQuarter = "Q2" Then
        txtDFrom = CVDate(GetMonth("Jul") & " 1," & txtYear)
        txtTFrom = DateAdd("m", 3, txtDFrom)
        txtTFrom = DateAdd("d", -1, txtTFrom)
    End If
    If comQuarter = "Q3" Then
        txtDFrom = CVDate(GetMonth("Oct") & " 1," & txtYear)
        txtTFrom = DateAdd("m", 3, txtDFrom)
        txtTFrom = DateAdd("d", -1, txtTFrom)
    End If
    If comQuarter = "Q4" Then
        txtDFrom = CVDate(GetMonth("Jan") & " 1," & txtYear + 1)
        txtTFrom = DateAdd("m", 3, txtDFrom)
        txtTFrom = DateAdd("d", -1, txtTFrom)
    End If
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
glbOnTop = Me.name

Screen.MousePointer = HOURGLASS
If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If
Call setRptCaption(Me)
If glbLinamar Then clpCode(3).MaxLength = 8

If glbCompSerial = "S/N - 2227W" Then clpCode(3).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6

Call INI_Controls(Me)

If glbCompSerial = "S/N - 2214W" Then 'Casey House
    frmQuarter.Top = 4800
    frmQuarter.Visible = True
    lblFDate.Visible = False
    lblTDate.Visible = False
    dlpFromDate.Visible = False
    dlpToDate.Visible = False
    
    comQuarter.AddItem ""
    comQuarter.AddItem "Q1"
    comQuarter.AddItem "Q2"
    comQuarter.AddItem "Q3"
    comQuarter.AddItem "Q4"
Else
    frmQuarter.Visible = False
    lblFDate.Visible = True
    lblTDate.Visible = True
    dlpFromDate.Visible = True
    dlpToDate.Visible = True
End If

'Ticket #29660 - Contract Employees Enhancement
If glbWFC Then
    chkExclCONP.Visible = True
    chkExclRET.Visible = True
Else
    chkExclCONP.Visible = False
    chkExclRET.Visible = False
End If

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

Private Sub txtShift_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtYear_Change()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtYear_GotFocus()
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

Private Sub txtYear_LostFocus()
If glbCompSerial = "S/N - 2214W" Then 'Casey House
    If IsNumeric(txtYear) Then
        If Len(txtYear) = 4 Then
        txtDFrom = CVDate(GetMonth("Apr") & " 1," & txtYear)
        txtTFrom = DateAdd("m", 12, txtDFrom)
        txtTFrom = DateAdd("d", -1, txtTFrom)
        End If
    End If
End If
End Sub
