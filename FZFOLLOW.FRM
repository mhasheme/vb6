VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRFollow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Follow-Up Report"
   ClientHeight    =   8310
   ClientLeft      =   435
   ClientTop       =   1140
   ClientWidth     =   9900
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
   ScaleHeight     =   8310
   ScaleWidth      =   9900
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkInclTerm 
      Caption         =   "Include Terminated Employees"
      Height          =   315
      Left            =   120
      TabIndex        =   15
      Top             =   5040
      Width           =   2925
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2100
      MaxLength       =   4
      TabIndex        =   14
      Tag             =   "00-Employee Position Shift"
      Top             =   4650
      Visible         =   0   'False
      Width           =   330
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   7
      Tag             =   "00-Enter Follow-Up Code"
      Top             =   2670
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "FURE"
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
      Index           =   3
      Left            =   2130
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Tag             =   "Final sorting of records - no totals"
      Top             =   6915
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
      Index           =   2
      Left            =   2130
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Tag             =   "Third level of grouping records"
      Top             =   6600
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
      Left            =   2130
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Tag             =   "Second level of grouping records"
      Top             =   6285
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
      Left            =   2130
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Tag             =   "First Level of grouping records"
      Top             =   5970
      Width           =   2325
   End
   Begin VB.TextBox txtComp 
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
      Left            =   2115
      MaxLength       =   1
      TabIndex        =   10
      Tag             =   "30-Y - Completed, N - Uncompleted - blank = both"
      Top             =   3330
      Width           =   450
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1800
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
      Left            =   1800
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
      Left            =   1800
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
      Left            =   1800
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
      Left            =   1800
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
      Left            =   1800
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
      Left            =   1800
      TabIndex        =   12
      Tag             =   "00-Enter Administered By Code"
      Top             =   3990
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
      Left            =   1800
      TabIndex        =   13
      Tag             =   "00-Enter Section Code"
      Top             =   4320
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   1800
      TabIndex        =   11
      Tag             =   "00-Enter Region Code"
      Top             =   3660
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   3480
      TabIndex        =   9
      Tag             =   "40-Date upto and including this date forward"
      Top             =   3000
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   8
      Tag             =   "40-Date from and including this date forward"
      Top             =   3000
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1800
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
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   7080
      Top             =   7680
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
      TabIndex        =   39
      Top             =   4650
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
      TabIndex        =   38
      Top             =   2010
      Width           =   630
   End
   Begin VB.Label lblSection 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      TabIndex        =   37
      Top             =   4320
      Width           =   540
   End
   Begin VB.Label lblAdmin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      TabIndex        =   36
      Top             =   3990
      Width           =   1125
   End
   Begin VB.Label lblRegion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      TabIndex        =   35
      Top             =   3660
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
      TabIndex        =   33
      Top             =   6945
      Width           =   660
   End
   Begin VB.Label lblGrp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Grouping #3"
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
      Index           =   2
      Left            =   120
      TabIndex        =   32
      Top             =   6630
      Width           =   885
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
      Top             =   6315
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
      Top             =   6000
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
      Top             =   5760
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
      TabIndex        =   28
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblComp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Completed Flag"
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
      Top             =   3330
      Width           =   1095
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "From / To Date"
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
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lblPosition 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Follow-Up"
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
      Top             =   2670
      Width           =   705
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
      Top             =   2340
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
      TabIndex        =   22
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
      TabIndex        =   21
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
      TabIndex        =   20
      Top             =   360
      Width           =   555
   End
End
Attribute VB_Name = "frmRFollow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flgWrongFUCode As Boolean

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm(lStr("Follow-ups Report Criteria"), Me) Then Exit Sub
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
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
     Call set_PrintState(True)
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

Private Sub clpCode_LostFocus(Index As Integer)
    If Index = 3 And Len(clpCode(3).Text) > 0 Then
        Dim rs As New ADODB.Recordset
        Dim strSQL As String
        Dim xWrongPos, xPos, I
        Dim xList, xShowCell, xCell
        Dim xTemplate As String
        
        '????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
        xTemplate = ""
        xTemplate = Get_Template(glbUserID)
        
        
        xList = clpCode(3).Text
        xWrongPos = 0
        xPos = 0
        Do While Len(xList) <> 0
            xWrongPos = xWrongPos + xPos
            xPos = InStr(xList, ",")
            If xPos = 0 Then
                xShowCell = xList
                xList = ""
            Else
                xShowCell = Left(xList, xPos - 1)
                xList = Mid(xList, xPos + 1)
            End If
            xCell = xShowCell
            
            If xTemplate = "" Or xTemplate = "TEMPLATE" Then
                strSQL = "SELECT ACCESSABLE FROM HR_SECURE_FOLLOW_UP WHERE USERID='" & Replace(glbUserID, "'", "''") & "'"
            Else
                '????Ticket #24808 -  Retrieve template's security profile
                strSQL = "SELECT ACCESSABLE FROM HR_SECURE_FOLLOW_UP WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
            End If
            strSQL = strSQL & " AND CODENAME = '" & xCell & "' AND TB_NAME='FURE'"
            rs.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
            If rs.EOF = False And rs.BOF = False Then
                If rs("ACCESSABLE") = 0 Then
                    flgWrongFUCode = True
                    MsgBox "You do not have Authorization for '" & xCell & "' Follow Up code", vbInformation + vbOKOnly, "Authorization Failure"
                    SendKeys "{HOME}"
                    For I = 1 To xWrongPos
                        SendKeys "{Right}"
                    Next
                    Exit Sub
                End If
            Else
                flgWrongFUCode = True
                MsgBox "You do not have Authorization for '" & xCell & "' Follow Up code", vbInformation + vbOKOnly, "Authorization Failure"
                SendKeys "{HOME}"
                For I = 1 To xWrongPos
                    SendKeys "{Right}"
                Next
                Exit Sub
            End If
            rs.Close
            Set rs = Nothing
        Loop
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
    
    If glbLinamar Then ' Frank May 2,2001
        comGroup(0).AddItem "Employment Type"
        comGroup(0).AddItem ("Home Line")
    End If
    If Not glbMulti Then comGroup(0).AddItem "Shift"
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem "(none)"
    comGroup(1).AddItem "Employee Name"
    comGroup(1).AddItem lStr("Follow-Up Reason")
    comGroup(1).AddItem "(none)"
    comGroup(2).AddItem lStr("Follow-Up Reason")
    comGroup(2).AddItem "(none)"
    comGroup(3).AddItem lStr("Follow-Up Date")
    comGroup(0).ListIndex = 0
    comGroup(1).ListIndex = 0
    comGroup(2).ListIndex = 0
    comGroup(3).Enabled = False
    comGroup(3).ListIndex = 0
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
        Case 3: strCd$ = "HR_FOLLOW_UP.EF_FREAS"
        Case 4: strCd$ = "HREMP.ED_REGION"
        Case 5: strCd$ = "HR_FOLLOW_UP.EF_ADMINBY"
        Case 6: strCd$ = "HREMP.ED_SECTION"
    End Select
    
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

Private Sub Cri_Comp()
Dim CompCri As String, intType%

If Len(txtComp.Text) > 0 Then
    'Franks May 8,2003 -ticket# 4086
    'Access and SQL Server use True and False instead of 1 and 0
    If glbOracle Then
        If txtComp.Text = "Y" Then
            CompCri = "({HR_FOLLOW_UP.EF_COMPLETED} <>0" & ")"
        ElseIf txtComp.Text = "N" Then
            CompCri = "({HR_FOLLOW_UP.EF_COMPLETED} = 0" & ")"
        End If
    Else
        If txtComp.Text = "Y" Then
            CompCri = "({HR_FOLLOW_UP.EF_COMPLETED} = True" & ")"
        ElseIf txtComp.Text = "N" Then
            CompCri = "({HR_FOLLOW_UP.EF_COMPLETED} = False" & ")"
        End If
    End If
End If


If Len(CompCri) >= 1 Then
    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND " & CompCri
    Else
        glbstrSelCri = CompCri
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

Private Sub Cri_FTDates()
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim x%

If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    TempCri = "({HR_FOLLOW_UP.EF_FDATE} "
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

For x% = 0 To 1
    If Len(dlpDateRange(x%).Text) > 0 Then
        TempCri = "({HR_FOLLOW_UP.EF_FDATE}  "
        If x% = 0 Then
            TempCri = TempCri & " >= "
        Else
            TempCri = TempCri & " <= "
        End If
        dtYYY% = Year(dlpDateRange(x%).Text)
        dtMM% = month(dlpDateRange(x%).Text)
        dtDD% = Day(dlpDateRange(x%).Text)
        TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
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

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

' call cri models set both glbiONeWhere and strSelCri
'Call glbCri_Dept(Me)  'laura nov 22, 1997
Call glbCri_DeptUN(clpDept.Text)

Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
Call Cri_PT
Call Cri_Shift
Call Cri_FTDates
Call Cri_Comp
Call Cri_EE
' dkostka - 07/05/2001 - 'Section' selection criteria was being ignored, fixed.
For x% = 0 To 6
    Call Cri_Code(x%)
Next x%

Call Cri_Sec

' report name
If comGroup(0) <> "(none)" Then
    If chkInclTerm Then
        strRName$ = glbIHRREPORTS & "rzfollupTerm.rpt"
    Else
        strRName$ = glbIHRREPORTS & "rzfollup.rpt"
    End If
Else
    If chkInclTerm Then
        strRName$ = glbIHRREPORTS & "rzfollu1Term.rpt"
    Else
        strRName$ = glbIHRREPORTS & "rzfollu1.rpt"
    End If
End If
Me.vbxCrystal.ReportFileName = strRName$


' set to sorting/grouping criteria
x% = Cri_Sorts()   ' returns number of sections formated

'set location for database tables

If glbNoNONE And glbNoEXEC Then 'Hemu -EXE
    glbstrSelCri = glbstrSelCri & " AND NOT (({HREMP.ED_ORG } = 'NONE') AND ({HREMP.ED_ORG } = 'EXEC') AND ({HR_FOLLOW_UP.EF_FREAS} = 'SREV') ) "
ElseIf glbNoNONE Then
    glbstrSelCri = glbstrSelCri & " AND NOT (({HREMP.ED_ORG } = 'NONE') AND ({HR_FOLLOW_UP.EF_FREAS} = 'SREV') ) "
ElseIf glbNoEXEC Then    'Hemu -EXE
    glbstrSelCri = glbstrSelCri & " AND NOT (({HREMP.ED_ORG } = 'EXEC') AND ({HR_FOLLOW_UP.EF_FREAS} = 'SREV') ) "
End If

'Follow Up security
'glbstrSelCri = glbstrSelCri & " AND {HR_SECURE_FOLLOW_UP.USERID} ='" & glbUserID & "' AND {HR_SECURE_FOLLOW_UP.ACCESSABLE} = True"

'Release 8.0 - Ticket #22682: View Own security
'If View Own not checked then do not retrieve follow ups of the User/Employee No
If Len(glbUserEmpNo) > 0 And glbUserEmpNo <> 0 And Not gSec_FollUp_ViewOwn Then
    'Do not show user's Follow Up records based on the Employee # associated to the User.
    If Len(glbstrSelCri) > 0 Then
        glbstrSelCri = glbstrSelCri & " AND {HR_FOLLOW_UP.EF_EMPNBR} <> " & glbUserEmpNo
    Else
        glbstrSelCri = glbstrSelCri & " {HR_FOLLOW_UP.EF_EMPNBR} <> " & glbUserEmpNo
    End If
End If

If Len(glbstrSelCri) >= 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If


Me.vbxCrystal.Connect = RptODBC_SQL


' window title if appropriate
Me.vbxCrystal.WindowTitle = lStr("Follow-ups Report")


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

Me.vbxCrystal.Formulas(1) = "lblRptTitle = '" & lStr("Follow-ups") & " Report'"
Me.vbxCrystal.Formulas(2) = "lblFollowup='" & lStr("Follow-Up") & "'"

' first set primary grouping
grpField$ = getEGroup(comGroup(0).Text)

If comGroup(0) <> "(none)" Then
    dscGroup$ = comGroup(0).Text
    dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
    Me.vbxCrystal.Formulas(0) = dscGroup$
    
    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$
    
    strSFormat$ = "GH1;T;X;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(0) = strSFormat$
    
    strSFormat$ = "GF1;T;X;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(1) = strSFormat$
    
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
        Case 1: grpField$ = "{tblFollow.TB_DESC}"
        Case 2: grpField$ = "(none)"
    End Select
    
    If grpField$ <> "(none)" Then
        If chkInclTerm Then
            grpCond$ = "GROUP" & CStr(3) & ";" & grpField$ & ";ANYCHANGE;A"
        Else
            grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
        End If
        Me.vbxCrystal.GroupCondition(1) = grpCond$
    
        GrpIdx% = comGroup(2).ListIndex
        Select Case GrpIdx%
            Case 0: grpField$ = "{tblFollow.TB_DESC}"
            Case 1: grpField$ = "(none)"
        End Select
        If grpField$ <> "(none)" Then
            If chkInclTerm Then
                grpCond$ = "GROUP" & CStr(4) & ";" & grpField$ & ";ANYCHANGE;A"
            Else
                grpCond$ = "GROUP" & CStr(3) & ";" & grpField$ & ";ANYCHANGE;A"
            End If
            Me.vbxCrystal.GroupCondition(2) = grpCond$
        Else
            If glbCompSerial = "S/N - 2370W" Then 'David Chapman's Ice Cream Limited
                If chkInclTerm Then
                    grpCond$ = "GROUP" & CStr(4) & ";" & "{HR_FOLLOW_UP.EF_FDATE}" & ";ANYCHANGE;A"
                Else
                    grpCond$ = "GROUP" & CStr(3) & ";" & "{HR_FOLLOW_UP.EF_FDATE}" & ";ANYCHANGE;A"
                End If
            Else
                If chkInclTerm Then
                    grpCond$ = "GROUP" & CStr(4) & ";" & "{HR_FOLLOW_UP.EF_FDATE}" & ";ANYCHANGE;D"
                Else
                    grpCond$ = "GROUP" & CStr(3) & ";" & "{HR_FOLLOW_UP.EF_FDATE}" & ";ANYCHANGE;D"
                End If
            End If
            Me.vbxCrystal.GroupCondition(2) = grpCond$
        End If
    End If
    'Here sorting starts when first grouping is "none"
Else
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
        Case 1: grpField$ = "{tblFollow.TB_DESC}"
        Case 2: grpField$ = "(none)"
    End Select
    If grpField$ <> "(none)" Then
        If chkInclTerm Then
            grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
        Else
            grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
        End If
        Me.vbxCrystal.GroupCondition(0) = grpCond$
        GrpIdx% = comGroup(2).ListIndex
        Select Case GrpIdx%
            Case 0: grpField$ = "{tblFollow.TB_DESC}"
            Case 1: grpField$ = "(none)"
        End Select
        If grpField$ <> "(none)" Then
            If chkInclTerm Then
                grpCond$ = "GROUP" & CStr(3) & ";" & grpField$ & ";ANYCHANGE;A"
            Else
                grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
            End If
            Me.vbxCrystal.GroupCondition(1) = grpCond$
            'grpCond$ = "GROUP" & CStr(3) & ";" & "{HR_FOLLOW_UP.EF_FDATE}" & ";ANYCHANGE;D"
            'Me.vbxCrystal.GroupCondition(2) = grpCond$
        End If
    Else
        'Here sorting starts when second grouping is also "none"
        GrpIdx% = comGroup(2).ListIndex
        Select Case GrpIdx%
            Case 0: grpField$ = "{tblFollow.TB_DESC}"
            Case 1: grpField$ = "(none)"
        End Select
        If grpField$ <> "(none)" Then
            If chkInclTerm Then
                grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
            Else
                grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
            End If
            Me.vbxCrystal.GroupCondition(0) = grpCond$
            
            If chkInclTerm Then
                grpCond$ = "GROUP" & CStr(3) & ";" & "{HR_FOLLOW_UP.EF_FDATE}" & ";ANYCHANGE;A"
            Else
                grpCond$ = "GROUP" & CStr(2) & ";" & "{HR_FOLLOW_UP.EF_FDATE}" & ";ANYCHANGE;A"
            End If
            Me.vbxCrystal.GroupCondition(1) = grpCond$
        
        'If third grouping is "none" then final sort is by "Education Year"
        Else
            If glbCompSerial = "S/N - 2370W" Then 'David Chapman's Ice Cream Limited
                If chkInclTerm Then
                    grpCond$ = "GROUP" & CStr(2) & ";" & "{HR_FOLLOW_UP.EF_FDATE}" & ";ANYCHANGE;A"
                Else
                    grpCond$ = "GROUP" & CStr(1) & ";" & "{HR_FOLLOW_UP.EF_FDATE}" & ";ANYCHANGE;A"
                End If
            Else
                If chkInclTerm Then
                    grpCond$ = "GROUP" & CStr(2) & ";" & "{HR_FOLLOW_UP.EF_FDATE}" & ";ANYCHANGE;D"
                Else
                    grpCond$ = "GROUP" & CStr(1) & ";" & "{HR_FOLLOW_UP.EF_FDATE}" & ";ANYCHANGE;D"
                End If
            End If
            Me.vbxCrystal.GroupCondition(0) = grpCond$
        End If
    End If
End If
Cri_Sorts = z% ' next section number to format

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

If Not clpDept.ListChecker Then
'If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    'MsgBox "If Department Entered - it must be known"
    'clpDept.SetFocus
    Exit Function
End If


For x% = 0 To 6
    If Not clpCode(x).ListChecker Then Exit Function
Next x%

'Send LostFocus on Follow Up code so it is validated as per the Follow Up Codes security
flgWrongFUCode = False
Call clpCode_LostFocus(3)
If flgWrongFUCode = True Then Exit Function

If Len(txtComp) > 0 Then
    If txtComp <> "Y" And txtComp <> "N" Then
        MsgBox "If entered completed must be Y or N"
        txtComp.SetFocus
        Exit Function
    End If
End If

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
If IsDate(dlpDateRange(0)) And IsDate(dlpDateRange(1)) Then
    If DaysBetween(dlpDateRange(0), dlpDateRange(1)) < 0 Then                               'Serbo
        MsgBox "To Date can't be prior to From Date!"                       '
        Me.dlpDateRange(0).SetFocus                                         '
        Exit Function                                                       '
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
glbOnTop = "FRMRFOLLOW"

Me.Caption = lStr("Follow-ups Report")

Screen.MousePointer = HOURGLASS
If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If
Call comGrpLoad
Call setRptCaption(Me)
Call setCaption(lblPosition)
If glbLinamar Then clpCode(4).MaxLength = 8
If glbCompSerial = "S/N - 2227W" Then clpCode(4).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6
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

Private Sub txtComp_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
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

Private Sub Cri_Sec()
    Dim EECri As String
    Dim strSec As String
    
    strSec = buildSec_FollowUp
    If Len(strSec) >= 1 Then
        EECri = "{HR_FOLLOW_UP.EF_FREAS} " & Replace(Replace(strSec, "(", "["), ")", "]")
    End If
    
    If Len(EECri) >= 1 Then
        If Len(glbstrSelCri) > 0 Then
            glbstrSelCri = glbstrSelCri & " AND " & EECri
        Else
            glbstrSelCri = EECri
        End If
    End If
    
End Sub

