VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmRMaster 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Employee/Comments"
   ClientHeight    =   7605
   ClientLeft      =   690
   ClientTop       =   630
   ClientWidth     =   10755
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
   ScaleHeight     =   7605
   ScaleWidth      =   10755
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
      TabIndex        =   15
      Tag             =   "Check to Exclude Employees with CONP Employment Status"
      Top             =   5040
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
      TabIndex        =   16
      Tag             =   "Check to Exclude Employees with RET Employment Status"
      Top             =   5340
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2085
      MaxLength       =   4
      TabIndex        =   14
      Tag             =   "00-Employee Position Shift"
      Top             =   4590
      Visible         =   0   'False
      Width           =   450
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1770
      TabIndex        =   8
      Tag             =   "00-Comment Type Code"
      Top             =   2940
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "ECOM"
      MaxLength       =   0
      MultiSelect     =   -1  'True
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
      Index           =   2
      Left            =   2010
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Tag             =   "Final Sort of Records"
      Top             =   6840
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
      Index           =   1
      Left            =   2010
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Tag             =   "Second level of grouping records"
      Top             =   6525
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
      Left            =   2010
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Tag             =   "First Level of grouping records"
      Top             =   6210
      Width           =   2325
   End
   Begin VB.ComboBox comEmpComm 
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
      Left            =   2070
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Tag             =   "10-Employee Comments Desired"
      Top             =   2580
      Width           =   870
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1770
      TabIndex        =   4
      Tag             =   "00-Enter Status Code"
      Top             =   1590
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
      Left            =   1770
      TabIndex        =   5
      Tag             =   "EDPT-Category"
      Top             =   1920
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
      Left            =   1770
      TabIndex        =   3
      Tag             =   "00-Enter Union Code"
      Top             =   1260
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
      Left            =   1770
      TabIndex        =   2
      Tag             =   "00-Enter Location Code"
      Top             =   930
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1770
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   600
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
      Left            =   1770
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   270
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
      Left            =   1770
      TabIndex        =   12
      Tag             =   "00-Enter Administered By Code"
      Top             =   3930
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   6
      Left            =   1770
      TabIndex        =   13
      Tag             =   "00-Enter Section Code"
      Top             =   4260
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   1770
      TabIndex        =   11
      Tag             =   "00-Enter Region Code"
      Top             =   3600
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   3810
      TabIndex        =   10
      Tag             =   "40-Date upto and including this date forward"
      Top             =   3270
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   0
      Left            =   1770
      TabIndex        =   9
      Tag             =   "40-Date from and including this date forward"
      Top             =   3270
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1770
      TabIndex        =   6
      Tag             =   "10-Enter Employee Number"
      Top             =   2250
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   7080
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
      TabIndex        =   38
      Top             =   4590
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
      TabIndex        =   37
      Top             =   1920
      Width           =   630
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
      TabIndex        =   36
      Top             =   4260
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
      TabIndex        =   35
      Top             =   3900
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
      TabIndex        =   34
      Top             =   3570
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
      TabIndex        =   33
      Top             =   930
      Width           =   615
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
      TabIndex        =   32
      Top             =   6870
      Width           =   660
   End
   Begin VB.Label lblGrp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Grouping #2"
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
      TabIndex        =   31
      Top             =   6555
      Width           =   885
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
      TabIndex        =   30
      Top             =   6240
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
      TabIndex        =   29
      Top             =   5940
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
      Left            =   -30
      TabIndex        =   28
      Top             =   30
      Width           =   1575
   End
   Begin VB.Label lblFromDate 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Effective Date Range"
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
      TabIndex        =   27
      Top             =   3240
      Width           =   1545
   End
   Begin VB.Label lblComType 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Comment Type"
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
      TabIndex        =   26
      Top             =   2910
      Width           =   1065
   End
   Begin VB.Label lblEmpComm 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Comments"
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
      Top             =   2580
      Width           =   1935
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
      TabIndex        =   24
      Top             =   2250
      Width           =   1290
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
      TabIndex        =   23
      Top             =   1590
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
      TabIndex        =   22
      Top             =   1260
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
      TabIndex        =   21
      Top             =   600
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
      TabIndex        =   20
      Top             =   270
      Width           =   555
   End
End
Attribute VB_Name = "frmRMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim glbstrSecCri As String

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm("Employee/Comments Report Criteria", Me) Then Exit Sub
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
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Public Sub cmdView_Click()
Dim x%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False

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
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
End If
Exit Sub

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "ATTEND", "SELECT")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub ComEComm()
If gSec_Rpt_Profiles Then comEmpComm.AddItem "No"
If gSec_Rpt_Profiles And gSec_Inq_Comments Then comEmpComm.AddItem "Yes"
If gSec_Inq_Comments Then comEmpComm.AddItem "Only"
comEmpComm.ListIndex = 0

End Sub

Private Sub chkExclCONP_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkExclRET_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub clpCode_LostFocus(Index As Integer)
'Added by Bryan 14/Oct/05 Ticket#9424
    If Index = 3 And Len(clpCode(3).Text) > 0 Then
        Dim rs As New ADODB.Recordset
        Dim strSQL As String
        Dim xTemplate As String
        
        '????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
        xTemplate = ""
        xTemplate = Get_Template(glbUserID)
        
        If xTemplate = "" Or xTemplate = "TEMPLATE" Then
            strSQL = "SELECT ACCESSABLE FROM HR_SECURE_COMMENTS WHERE USERID='" & Replace(glbUserID, "'", "''") & "'"
        Else
            '????Ticket #24808 -  Retrieve template's security profile
            strSQL = "SELECT ACCESSABLE FROM HR_SECURE_COMMENTS WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
        End If
        strSQL = strSQL & " AND CODENAME IN ('" & Replace(clpCode(3).Text, ",", "','") & "') AND TB_NAME='ECOM'"
        rs.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
        If rs.EOF = False And rs.BOF = False Then
            If rs("ACCESSABLE") = 0 Then
                MsgBox "You do not have Authorization for that Comment Type", vbInformation + vbOKOnly, "Authorization Failure"
                clpCode(3).Text = ""
            End If
        Else
            MsgBox "You do not have Authorization for that Comment Type", vbInformation + vbOKOnly, "Authorization Failure"
            clpCode(3).Text = ""
        End If
        rs.Close
        Set rs = Nothing
    End If
End Sub

Private Sub comEmpComm_Click()
Dim CNT As Integer
  If comEmpComm.Text <> "No" Then
     clpCode(3).Enabled = True
     For CNT = 0 To 1
        dlpDateRange(CNT).Enabled = True
     Next CNT
  Else
     clpCode(3).Enabled = False
     For CNT = 0 To 1
        dlpDateRange(CNT).Enabled = False
     Next CNT
  End If

End Sub

Private Sub comEmpComm_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub comEmpComm_KeyUp(KeyCode As Integer, Shift As Integer)
Dim CNT As Integer
  If comEmpComm.Text <> "No" Then
     clpCode(3).Enabled = True
     For CNT = 0 To 1
        dlpDateRange(CNT).Enabled = True
     Next CNT
  Else
     clpCode(3).Enabled = False
     For CNT = 0 To 1
       dlpDateRange(CNT).Enabled = False
     Next CNT
  End If


End Sub

Private Sub comGroup_GotFocus(Index As Integer)
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()
    
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")
    comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
    comGroup(0).AddItem "Employee Name"
    If glbLinamar Then ' Frank May 2,2001
        comGroup(0).AddItem "Employment Type"
        comGroup(0).AddItem lStr("Region")
        comGroup(0).AddItem ("Home Line")
    End If
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem "(none)"
    
    comGroup(1).AddItem "Employee Name"
    comGroup(1).AddItem ""
    comGroup(2).AddItem "Comment"
    comGroup(0).ListIndex = 0
    comGroup(1).ListIndex = 0
    comGroup(2).ListIndex = 0
    comGroup(2).Enabled = False

End Sub

Private Sub Cri_Assoc()
Dim EECri As String

If Len(clpCode(1).Text) > 0 Then
    'EECri = "{HREMP.ED_ORG} = '" & clpCode(1).Text & "' "
    EECri = "({HREMP.ED_ORG} in  ['" & Replace(clpCode(1).Text, ",", "','") & "'])"
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

Private Sub Cri_Code(intIdx%)
Dim CodeCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

If Len(clpCode(intIdx%).Text) > 0 Then
    If intIdx% = 0 Then strCd$ = "HREMP.ED_LOC"
    If intIdx% = 4 Then strCd$ = "HREMP.ED_REGION"
    If intIdx% = 5 Then strCd$ = "HREMP.ED_ADMINBY"
    If intIdx% = 6 Then strCd$ = "HREMP.ED_SECTION"  'Lucy July 4, 2000
        CodeCri = "({" & strCd$ & "} in  ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
    If glbLinamar And (strCd$ = "HREMP.ED_REGION" Or strCd$ = "HREMP.ED_SECTION") Then
        CodeCri = "(({" & strCd$ & "} = '" & clpDiv.Text & clpCode(intIdx%).Text & "') or ({" & strCd$ & "} = 'ALL" & clpCode(intIdx%).Text & "') )"
    End If
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

Private Sub Cri_Comment()
Dim EECri As String


If Len(clpCode(3).Text) > 0 Then
    EECri = "({HR_COMMENTS.CO_TYPE} IN  ['" & Replace(clpCode(3).Text, ",", "','") & "'])"   ''" & clpCode(3).Text & "' "
End If


If Len(EECri) >= 1 Then
    If glbiOneWhere Then
        glbstrSecCri = glbstrSecCri & " AND " & EECri
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSecCri = EECri
        glbstrSelCri = EECri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_Dates()
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim x%
  If Len(dlpDateRange(0).Text) = 0 And Len(dlpDateRange(1).Text) = 0 Then Exit Sub

TempCri = "({HR_COMMENTS.CO_EDATE} "
If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
  For x% = 0 To 1
    dtYYY% = Year(dlpDateRange(x%).Text)
    dtMM% = month(dlpDateRange(x%).Text)
    dtDD% = Day(dlpDateRange(x%).Text)
    If x% = 0 Then
      TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    Else
      TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    End If
  Next x%
Else
  If Len(dlpDateRange(0).Text) > 0 Then
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    TempCri = TempCri & " >= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
  End If
  If Len(dlpDateRange(1).Text) > 0 Then
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    TempCri = TempCri & " <= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
  End If

End If


GoTo Cri_Datst

'If Len(txtDateRange(0).Text) > 0 And Len(txtDateRange(1).Text) > 0 Then
'    TempCri = "({HR_COMMENTS.CO_EDATE} "
'    dtYYY% = Year(txtDateRange(0).Text)
'    dtMM% = Month(txtDateRange(0).Text)
'    dtDD% = Day(txtDateRange(0).Text)
'    TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
'    dtYYY% = Year(txtDateRange(1).Text)
'    dtMM% = Month(txtDateRange(1).Text)
'    dtDD% = Day(txtDateRange(1).Text)
'    TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
'    GoTo Cri_Datst
'End If



Cri_Datst:
If Len(TempCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSecCri = TempCri
        glbstrSelCri = TempCri
    Else
        glbstrSecCri = glbstrSecCri & " AND " & TempCri
        glbstrSelCri = glbstrSelCri & " AND " & TempCri
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

EECri = "{HREMP.ED_PT} in ['" & Replace(clpPT.Text, ",", "','") & "']"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True


End Sub

Private Function Cri_SetAll()
Dim x%, strRName$
Dim RPT
Dim C As Integer

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""
glbstrSecCri = ""

' call cri models set both glbiONeWhere and strSelCri
'Call glbCri_Dept(Me)  'laura nov 22, 1997
Call glbCri_DeptUN(clpDept.Text)

Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
Call Cri_Assoc
Call Cri_Status
Call Cri_PT
Call Cri_Shift
Call Cri_EE
Call Cri_Code(0)
Call Cri_Code(4)
Call Cri_Code(5)
' dkostka - 07/05/2001 - 'Section' selection criteria was being ignored, fixed.
Call Cri_Code(6)

If comEmpComm.Text <> "No" Then
  Call Cri_Comment
  Call Cri_Sec
  Call Cri_Dates
End If
' report name


' set to sorting/grouping criteria
x% = Cri_Sorts()   ' returns number of sections formated

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
    If comEmpComm.Text = "Yes" And Len(clpCode(3).Text) > 0 Then
        Me.vbxCrystal.SelectionFormula = glbstrSelCri & glbstrSecCri
    Else
        Me.vbxCrystal.SelectionFormula = glbstrSelCri
    End If
End If


'Release 8.0 - Ticket #22682: View Own security
If comEmpComm.Text <> "No" Then
    'If View Own not checked then do not retrieve Comments of the User/Employee No
    If Len(glbUserEmpNo) > 0 And glbUserEmpNo <> 0 And Not gSec_Comments_ViewOwn Then
        'Do not show user's Comments records based on the Employee # associated to the User.
        If Len(glbstrSelCri) > 0 Then
            glbstrSelCri = glbstrSelCri & " AND ({HR_COMMENTS.CO_EMPNBR} <> " & glbUserEmpNo & ")"
        Else
            glbstrSelCri = glbstrSelCri & " {HR_COMMENTS.CO_EMPNBR} <> " & glbUserEmpNo
        End If
    
        If Len(glbstrSecCri) > 0 Then
            glbstrSecCri = glbstrSecCri & " AND ({HR_COMMENTS.CO_EMPNBR} <> " & glbUserEmpNo & ")"
        Else
            glbstrSecCri = glbstrSecCri & " {HR_COMMENTS.CO_EMPNBR} <> " & glbUserEmpNo
        End If
    
        If comEmpComm.Text = "Only" Then
            glbstrSelCri = glbstrSelCri & " AND ({HREMP.ED_EMPNBR} <> " & glbUserEmpNo & ")"
            glbstrSecCri = glbstrSecCri & " AND ({HREMP.ED_EMPNBR} <> " & glbUserEmpNo & ")"
        End If
    End If
End If

Select Case comEmpComm.Text
    Case "Yes"
        strRName$ = glbIHRREPORTS & "rzmaster.rpt"
        RPT = 5
    Case "No"
        strRName$ = glbIHRREPORTS & "rzmastrn.rpt"
        RPT = 6
    Case "Only"
        If comGroup(0) = "(none)" Then
            strRName$ = glbIHRREPORTS & "rzmastr1N.rpt"
        Else
            strRName$ = glbIHRREPORTS & "rzmastr1.rpt"
        End If
        RPT = 7
End Select

Me.vbxCrystal.ReportFileName = strRName$

If glbOracle Or glbSQL Or RPT = 6 Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    If RPT = 7 Then
        Me.vbxCrystal.Connect = "PWD=petman;"
        For C = 0 To 11
            vbxCrystal.DataFiles(C) = glbIHRDB
        Next C
    Else
        Me.vbxCrystal.Connect = RptODBC_SQL
    End If
End If
'Comments security
If RPT = 5 And Len(clpCode(3).Text) = 0 Then
    'Ticket #18678 - Report freezes so had to put the condition only to pass the selection formula when the
    'Comments Code is not entered. This seems to resolve the issue.
    Me.vbxCrystal.SubreportToChange = "Comments"
    Me.vbxCrystal.SelectionFormula = glbstrSecCri
    Me.vbxCrystal.SubreportToChange = ""
End If
' window title if appropriate
Me.vbxCrystal.WindowTitle = lStr("Employee/Comments Report")


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

Private Sub Cri_Shift()
Dim EECri As String, OneSet%, x%

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
Dim x%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

'for labels - sort by name always
' imbeded in report

Cri_Sorts = 0

grpField$ = getEGroup(comGroup(0).Text)

If grpField$ = "(none)" And comEmpComm.Text <> "Only" Then grpField$ = "{@EFullName}"

dscGroup$ = comGroup(0).Text
dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"

Me.vbxCrystal.Formulas(0) = dscGroup$

Me.vbxCrystal.Formulas(1) = "showSIN = " & IIf(gSec_Show_SIN_SSN = 0, False, True) & " "
Me.vbxCrystal.Formulas(2) = "showDOB = " & IIf(gSec_Show_DOB = 0, False, True) & " "
Me.vbxCrystal.Formulas(3) = "showADDRESS = " & IIf(gSec_Show_ADDRESS = 0, False, True) & " "
Me.vbxCrystal.Formulas(4) = "showMarital = " & IIf(gSec_Show_Marital = 0, False, True) & " "

If comEmpComm.Text = "Only" And grpField$ = "(none)" Then
    'do not pass any grouping
Else
    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$
End If

If comEmpComm.Text <> "Only" Then
    Call setRptLabel(Me, 1)
End If

Cri_Sorts = z% ' next section number to format

End Function

Private Sub Cri_Status()
Dim EECri As String

If Len(clpCode(2).Text) > 0 Then
    EECri = "{HREMP.ED_EMP} in ['" & Replace(clpCode(2).Text, ",", "','") & "']"
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

For x% = 0 To 6
    If Not clpCode(x).ListChecker Then Exit Function
Next x%

 If Len(dlpDateRange(0).Text) > 0 Then
    If Not IsDate(dlpDateRange(0).Text) Then
        MsgBox "Not a valid date"
        dlpDateRange(0).Text = ""
        dlpDateRange(0).SetFocus
        Exit Function
    End If
 End If
 If Len(dlpDateRange(1).Text) > 0 Then
    If Not IsDate(dlpDateRange(1).Text) Then
        MsgBox "Not a valid date"
        dlpDateRange(1).Text = ""
        dlpDateRange(1).SetFocus
        Exit Function
    End If
 End If
    'check to ensure that the from date is <= the to date
 If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
      If CVDate(dlpDateRange(0).Text) > CVDate(dlpDateRange(1).Text) Then
        MsgBox "Not a valid date range"
        dlpDateRange(1).Text = ""
        dlpDateRange(1).SetFocus
        Exit Function
      End If
 End If

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

glbOnTop = "FRMRMASTER"

Screen.MousePointer = HOURGLASS

If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If

Call ComEComm
Call comGrpLoad
Call setRptCaption(Me)

If glbLinamar Then clpCode(4).MaxLength = 8
If glbCompSerial = "S/N - 2227W" Then clpCode(4).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6

Call INI_Controls(Me)

lblEmpComm.Caption = lStr(lblEmpComm.Caption)
Me.Caption = lStr(Me.Caption)

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
Call SetPanHelp(Me.ActiveControl)
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

Private Sub Cri_Sec()
    Dim EECri As String
    Dim strSec As String
    
    strSec = buildSec
    If Len(strSec) >= 1 Then
        EECri = "{HR_COMMENTS.CO_TYPE} " & Replace(Replace(strSec, "(", "["), ")", "]")
    End If
    
If Len(EECri) >= 1 Then
    If comEmpComm.Text = "Only" Then
        If Len(glbstrSelCri) > 0 Then
            glbstrSelCri = glbstrSelCri & " AND " & EECri
        Else
            glbstrSelCri = EECri
        End If
    End If
    If Len(glbstrSecCri) > 0 Then
        glbstrSecCri = glbstrSecCri & " AND " & EECri
    Else
        glbstrSecCri = EECri
    End If
End If
    
End Sub


