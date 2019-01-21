VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRWCBCS 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Employee WSIB Cost Summary Report"
   ClientHeight    =   7470
   ClientLeft      =   390
   ClientTop       =   1185
   ClientWidth     =   10080
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
   ScaleHeight     =   7470
   ScaleWidth      =   10080
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtHRat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "EC_HAZARD"
      Height          =   315
      Left            =   3180
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   5162
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox cmbHRat 
      Height          =   315
      Left            =   2355
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Tag             =   "00-Hazard Rating - Choose A,B, or C"
      Top             =   5162
      Width           =   645
   End
   Begin VB.Frame fraDetail 
      Caption         =   "Report Type"
      Height          =   615
      Left            =   120
      TabIndex        =   37
      Top             =   6480
      Width           =   3255
      Begin VB.OptionButton optSummary 
         Caption         =   "Summary"
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
         Left            =   1920
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optDetails 
         Caption         =   "Detail"
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
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.CheckBox chkTerm 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2355
      TabIndex        =   19
      Top             =   6000
      Width           =   255
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2355
      MaxLength       =   4
      TabIndex        =   15
      Tag             =   "00-Employee Position Shift"
      Top             =   4819
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.TextBox txtWCB 
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
      Left            =   2355
      MaxLength       =   12
      TabIndex        =   7
      Tag             =   "00-Enter WSIB Number"
      Top             =   2761
      Width           =   1215
   End
   Begin MSMask.MaskEdBox medDOW 
      Height          =   285
      Index           =   0
      Left            =   3210
      TabIndex        =   16
      Tag             =   "10-Day of week 1=Sunday, 2=Monday..."
      Top             =   4815
      Visible         =   0   'False
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0"
      Mask            =   "#"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   3
      Tag             =   "00-Enter Status Code"
      Top             =   1395
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDOR"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Tag             =   "EDPT-Category"
      Top             =   2070
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
      Left            =   2040
      TabIndex        =   4
      Tag             =   "00-Enter Union Code"
      Top             =   1725
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDEM"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   2
      Tag             =   "00-Enter Location Code"
      Top             =   1050
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   705
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
      Left            =   2040
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
      Index           =   4
      Left            =   2040
      TabIndex        =   13
      Tag             =   "00-Enter Administered By Code"
      Top             =   4140
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
      Left            =   2040
      TabIndex        =   14
      Tag             =   "00-Enter Section Code"
      Top             =   4470
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   12
      Tag             =   "00-Enter Region Code"
      Top             =   3795
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   3990
      TabIndex        =   11
      Tag             =   "40-Date upto and including this date"
      Top             =   3450
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   10
      Tag             =   "40-Date from and including this date forward"
      Top             =   3480
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Tag             =   "10-Enter Employee Number"
      Top             =   2415
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   6720
      Top             =   5400
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
      Index           =   6
      Left            =   2040
      TabIndex        =   18
      Tag             =   "01-Results of Claim - Code"
      Top             =   5535
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECRS"
   End
   Begin INFOHR_Controls.DateLookup dlpStateDate 
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   8
      Tag             =   "40-Statement Date From and including this date forward"
      Top             =   3105
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpStateDate 
      Height          =   285
      Index           =   1
      Left            =   3990
      TabIndex        =   9
      Tag             =   "40-Statement Date To upto and including this date"
      Top             =   3120
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Statement Date (From/To)"
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
      TabIndex        =   41
      Top             =   3150
      Width           =   1860
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Result"
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
      TabIndex        =   40
      Top             =   5580
      Width           =   450
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Hazard Rating"
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
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   38
      Tag             =   "00-Hazard Rationg"
      Top             =   5192
      Width           =   1140
   End
   Begin VB.Label lblTerm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Include Termination"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   36
      Top             =   6030
      Width           =   1845
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
      TabIndex        =   35
      Top             =   4864
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
      TabIndex        =   34
      Top             =   2120
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
      TabIndex        =   33
      Top             =   4521
      Width           =   540
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
      TabIndex        =   32
      Top             =   3835
      Width           =   510
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
      TabIndex        =   31
      Top             =   4178
      Width           =   1125
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
      TabIndex        =   30
      Top             =   1091
      Width           =   615
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
      TabIndex        =   29
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "From /To Date"
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
      Top             =   3492
      Width           =   1050
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WSIB Number"
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
      Top             =   2806
      Width           =   1020
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
      TabIndex        =   26
      Top             =   2463
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
      TabIndex        =   25
      Top             =   1777
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
      Top             =   1434
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
      Top             =   748
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
      Top             =   405
      Width           =   555
   End
End
Attribute VB_Name = "frmRWCBCS"
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
    If frmRWCBCS.Caption = "Employee WSIB Cost Summary Report" Then
        If Not PrtForm("Employee WSIB Cost Summary Report Criteria", Me) Then Exit Sub
    ElseIf frmRWCBCS.Caption = "Total Cost Report" Then
        If Not PrtForm("Total Cost Report Criteria", Me) Then Exit Sub
    End If
    Call set_PrintState(False)
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
On Error GoTo CRW_Err

If CriCheck() Then
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
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
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
End If
Exit Sub

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "OHS", "SELECT")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub Cri_Code(intIdx%)
Dim CodeCri As String
Dim strCd$
Dim tbl As String

If Len(clpCode(intIdx%).Text) > 0 Then
        
    Select Case intIdx%
    Case 0: strCd$ = "HREMP.ED_LOC"
    Case 3: strCd$ = "HREMP.ED_REGION"
    Case 4: strCd$ = "HREMP.ED_ADMINBY"
    Case 5: strCd$ = "HREMP.ED_SECTION"
    Case 1: strCd$ = "HREMP.ED_EMP"
    Case 2: strCd$ = "HREMP.ED_ORG"
    Case 6: strCd$ = "HR_OCC_HEALTH_SAFETY.EC_WCBRES"
    Case Else: strCd$ = "HR_OCC_HEALTH_SAFETY.EC_TYPE"
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

Private Sub Cri_DOW()
Dim EECri As String, OneSet%
Dim X%


For X% = 0 To 1
    If Len(medDOW(X%)) > 0 Then
        OneSet% = OneSet% + 1
    End If
Next X%

If OneSet% = 0 Then Exit Sub


For X% = 0 To 1
    If Len(medDOW(X%)) > 0 Then
        EECri = EECri & medDOW(X%)
        OneSet% = OneSet% - 1
    End If
    If OneSet% > 0 Then
        EECri = EECri & " ,"
    End If
Next X%
EECri = "{@DOW} in [" & EECri & "]"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True


End Sub

Private Sub Cri_EE()
Dim EECri As String

If Len(elpEEID.Text) <= 0 Then Exit Sub

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

Private Sub Cri_StatementDate()

Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim X%
Dim tbl As String

If Len(dlpStateDate(0).Text) <= 0 And Len(dlpStateDate(1).Text) <= 0 Then Exit Sub

If Len(dlpStateDate(0).Text) > 0 And Len(dlpStateDate(1).Text) > 0 Then
    TempCri = "({HROHSCOS.CC_STMTDT} "
    dtYYY% = Year(dlpStateDate(0).Text)
    dtMM% = month(dlpStateDate(0).Text)
    dtDD% = Day(dlpStateDate(0).Text)
    TempCri = TempCri & "in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")"

    dtYYY% = Year(dlpStateDate(1).Text)
    dtMM% = month(dlpStateDate(1).Text)
    dtDD% = Day(dlpStateDate(1).Text)
    TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "

    GoTo Cri_StatDate
    
ElseIf Len(dlpStateDate(0).Text) > 0 Then
    TempCri = "({HROHSCOS.CC_STMTDT} "         ' Added section to enable entering only From date, no To date.
    dtYYY% = Year(dlpStateDate(0).Text)
    dtMM% = month(dlpStateDate(0).Text)
    dtDD% = Day(dlpStateDate(0).Text)
    TempCri = TempCri & " >= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    
    GoTo Cri_StatDate
    
ElseIf Len(dlpStateDate(1).Text) > 0 Then
    TempCri = "({HROHSCOS.CC_STMTDT} "         ' Added section to enable entering only To date, no From date.
    dtYYY% = Year(dlpStateDate(1).Text)
    dtMM% = month(dlpStateDate(1).Text)
    dtDD% = Day(dlpStateDate(1).Text)
    TempCri = TempCri & " <= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    
    GoTo Cri_StatDate
End If

Cri_StatDate:
If Len(TempCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = TempCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & TempCri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_FTDates()

Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim X%
Dim tbl As String

If Len(dlpDateRange(0).Text) <= 0 And Len(dlpDateRange(1).Text) <= 0 Then Exit Sub

If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    TempCri = "({HROHSCOS.CC_FDATE} "
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    TempCri = TempCri & "in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")"

    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "

'Laura if CC_TDATE
    TempCri = TempCri & " AND ({HROHSCOS.CC_TDATE} "
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    TempCri = TempCri & "in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")"
     
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
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

Private Sub Cri_Hazard()
Dim BodyCri As String

If cmbHRat.Text <> "" Then
    BodyCri = "{HR_OCC_HEALTH_SAFETY.EC_HAZARD} = '" & cmbHRat.Text & "'"
Else
    Exit Sub
End If

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & BodyCri
Else
    glbstrSelCri = BodyCri
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

' call cri models set both glbiONeWhere and strSelCri
'Call glbCri_Dept(Me)  'laura nov 22, 1997
Call glbCri_DeptUN(clpDept.Text)

Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
' dkostka - 07/05/2001 - 'Section' selection criteria was being ignored, fixed.
For X% = 0 To 6
    Call Cri_Code(X%)
Next
Call Cri_EE       'laura dec 03, 1997
Call Cri_WCB      'laura dec 03, 1997
Call Cri_PT
Call Cri_Shift
Call Cri_StatementDate
Call Cri_FTDates
Call Cri_Hazard

If chkTerm.Value = vbChecked Then 'Bryan 06/Jul/05 Ticket# 8857
    Call SELATTWRK
    glbstrSelCri = Replace(glbstrSelCri, "HREMP", "HRWSIBWRK")
    glbstrSelCri = Replace(glbstrSelCri, "HROHSCOS", "HRWSIBWRK")
    glbstrSelCri = Replace(glbstrSelCri, "HR_OCC_HEALTH_SAFETY", "HROCCHSWRK")
    If Len(glbstrSelCri) > 0 Then
        glbstrSelCri = glbstrSelCri & " AND ({HRWSIBWRK.WRKEMP} = '" & glbUserID & "' )"
    End If
End If

If Len(glbstrSelCri) >= 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If

'Report Parameters
If Len(dlpStateDate(0).Text) > 0 And Len(dlpStateDate(1).Text) > 0 Then
    Me.vbxCrystal.Formulas(0) = "StatementDate = 'Statement Date From: " & dlpStateDate(0).Text & "  To: " & dlpStateDate(1).Text & "'"
ElseIf Len(dlpStateDate(0).Text) > 0 Then
    Me.vbxCrystal.Formulas(0) = "StatementDate = 'Statement Date From: " & dlpStateDate(0).Text & "'"
ElseIf Len(dlpStateDate(1).Text) > 0 Then
    Me.vbxCrystal.Formulas(0) = "StatementDate = 'Statement Date To: " & dlpStateDate(1).Text & "'"
End If

If Len(dlpDateRange(0).Text) > 0 Then
    Me.vbxCrystal.Formulas(1) = "FTDateRange = 'From Date: " & dlpDateRange(0).Text & "   To Date: " & dlpDateRange(1).Text & "'"
End If

' report name
If optDetails.Value = True Then
    If chkTerm = vbChecked Then 'Bryan 06/Jul/05 Ticket# 8857
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzwcbcs1.rpt"
        If glbSQL Or glbOracle Then
            Me.vbxCrystal.Connect = RptODBC_SQL
        Else
            Me.vbxCrystal.Connect = "PWD=petman;"
            
            For X% = 0 To 2
                Me.vbxCrystal.DataFiles(X%) = glbIHRDB
            Next X%
            Me.vbxCrystal.DataFiles(3) = glbIHRDBW
            Me.vbxCrystal.DataFiles(4) = glbIHRDB
        End If
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzwcbcs.rpt"
        If glbSQL Or glbOracle Then
            Me.vbxCrystal.Connect = RptODBC_SQL
        Else
            Me.vbxCrystal.Connect = "PWD=petman;"
            
            For X% = 0 To 5
                Me.vbxCrystal.DataFiles(X%) = glbIHRDB
            Next X%
        End If
    End If

    Me.vbxCrystal.WindowTitle = "Employee WSIB Cost Report"

ElseIf optSummary.Value = True Then
    If chkTerm = vbChecked Then 'Bryan 06/Jul/05 Ticket# 8857
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzinct1.rpt"
        If glbSQL Or glbOracle Then
            Me.vbxCrystal.Connect = RptODBC_SQL
        Else
            Me.vbxCrystal.Connect = "PWD=petman;"
            
            For X% = 0 To 2
                Me.vbxCrystal.DataFiles(X%) = glbIHRDB
            Next X%
            Me.vbxCrystal.DataFiles(3) = glbIHRDBW
            Me.vbxCrystal.DataFiles(4) = glbIHRDB
        End If
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzinct.rpt"
        If glbSQL Or glbOracle Then
            Me.vbxCrystal.Connect = RptODBC_SQL
        Else
            Me.vbxCrystal.Connect = "PWD=petman;"
            
            For X% = 0 To 5
                Me.vbxCrystal.DataFiles(X%) = glbIHRDB
            Next X%
        End If
    End If
    Me.vbxCrystal.WindowTitle = "Employee WSIB Cost Summary Report"
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

Private Sub Cri_WCB()
Dim EECri As String, OneSet%, X%

If Len(txtWCB.Text) < 1 Then Exit Sub


EECri = "{HROHSCOS.CC_WCBNBR}= '" & txtWCB.Text & "'"


If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

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

For X% = 0 To 5
    If Not clpCode(X).ListChecker Then Exit Function
Next X%

If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

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

If dlpDateRange(0).Text <> "" And dlpDateRange(1).Text = "" Then
    MsgBox "If entering a date, both from and to date must be entered!"
    dlpDateRange(0).Text = ""
    dlpDateRange(0).SetFocus
    Exit Function
End If

If dlpDateRange(1).Text <> "" And dlpDateRange(0).Text = "" Then
    MsgBox "If entering a date, both from and to date must be entered!"
    dlpDateRange(1).Text = ""
    dlpDateRange(1).SetFocus
    Exit Function
End If

If dlpDateRange(0).Text <> "" And dlpDateRange(1).Text <> "" Then
    If CVDate(dlpDateRange(0).Text) > CVDate(dlpDateRange(1).Text) Then
        MsgBox "From Date must be less than To Date!"
        dlpDateRange(1).SetFocus
        Exit Function
    End If
End If

'For X% = 0 To 1
' If Len(medDOW(X%)) > 0 Then
'    If medDOW(X%) < 0 Or medDOW(X%) > 7 Then
'        MsgBox "Day of Week must be between 1 and 7"
'        medDOW(X%).SetFocus
'        Exit Function
'    End If
'  End If
'Next X%

If Not elpEEID.ListChecker Then
    Exit Function
End If

CriCheck = True

End Function

Private Sub clpCode_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbHRat_Change()
'Ticket # 6831 - For Burlington Tech.
If glbCompSerial = "S/N - 2351W" Then
    If cmbHRat = "" Then
        txtHRat = ""
    End If
End If
End Sub

Private Sub cmbHRat_Click()
'Ticket # 6831 - For Burlington Tech.
If glbCompSerial = "S/N - 2351W" Then
    Select Case cmbHRat
        Case "Frequent": txtHRat = "F"
        Case "Occassional": txtHRat = "O"
        Case "Rare": txtHRat = "R"
        Case Else
            txtHRat = ""
    End Select
Else
    txtHRat = cmbHRat
End If
End Sub

Private Sub cmbHRat_GotFocus()
Call SetPanHelp(ActiveControl)
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

'Ticket # 6831 - For Burlington Tech.
If glbCompSerial = "S/N - 2351W" Then
    lblTitle(9).Caption = "Frequency"
    cmbHRat.Width = 1500
    'cmbHRat.Left = 7560
    cmbHRat.AddItem ""
    cmbHRat.AddItem "Frequent"
    cmbHRat.AddItem "Occassional"
    cmbHRat.AddItem "Rare"
Else
    cmbHRat.AddItem ""
    cmbHRat.AddItem "A"
    cmbHRat.AddItem "B"
    cmbHRat.AddItem "C"
End If

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

Private Sub txtShift_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtWCB_GotFocus()
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

Private Sub SELATTWRK()
'Created by Bryan 6/Jul/05 Ticket #8857
Dim xlen, xxx, xx1
Dim db001 As Database
Dim SQLQ
Dim xFieldList

On Error GoTo AttWrkError

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 1
MDIMain.panHelp(0).FloodPercent = 3

gdbAdoIhr001.CommandTimeout = 600

Screen.MousePointer = HOURGLASS

MDIMain.panHelp(0).FloodPercent = 15

gdbAdoIhr001.BeginTrans
SQLQ = "DELETE FROM HRWSIBWRK " & in_SQL(glbIHRDBW) & " WHERE WRKEMP='" & glbUserID & "'"
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans

If Not glbSQL And Not glbOracle Then Call Pause(1)

MDIMain.panHelp(0).FloodPercent = 30

xFieldList = Get_Fields(gdbAdoIhr001W, "HRWSIBWRK", "WRKEMP")
SQLQ = "INSERT INTO HRWSIBWRK (" & xFieldList & ",WRKEMP) "
SQLQ = SQLQ & in_SQL(glbIHRDBW)
SQLQ = SQLQ & " SELECT " & xFieldList & ",'" & glbUserID & "' AS WRKEMP "
If Not glbOracle Then
    'Ticket #16854
    'SQLQ = SQLQ & " FROM Term_HROHSCOS LEFT JOIN Term_HREMP ON Term_HROHSCOS.CC_EMPNBR = Term_HREMP.ED_EMPNBR "
    'Ticket #19554
    'SQLQ = SQLQ & " FROM Term_HROHSCOS LEFT JOIN Term_HREMP ON Term_HROHSCOS.TERM_SEQ = Term_HREMP.TERM_SEQ "
    SQLQ = SQLQ & " FROM HROHSCOS LEFT JOIN HREMP ON HROHSCOS.CC_EMPNBR = HREMP.ED_EMPNBR "
Else
    'Ticket #16854
    'SQLQ = SQLQ & " FROM TERM_HROHSCOS, TERM_HREMP WHERE TERM_HROHSCOS.CC_EMPNBR = TERM_HREMP.ED_EMPNBR(+) "
    SQLQ = SQLQ & " FROM TERM_HROHSCOS, TERM_HREMP WHERE TERM_HROHSCOS.TERM_SEQ = TERM_HREMP.TERM_SEQ(+) "
End If

MDIMain.panHelp(0).FloodPercent = 45

gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ, , adCmdText
gdbAdoIhr001.CommitTrans

If Not glbSQL And Not glbOracle Then Call Pause(1)

MDIMain.panHelp(0).FloodPercent = 50

SQLQ = "INSERT INTO HRWSIBWRK (" & xFieldList & ",WRKEMP) "
SQLQ = SQLQ & in_SQL(glbIHRDBW)
SQLQ = SQLQ & " SELECT " & xFieldList & ",'" & glbUserID & "' AS WRKEMP "
If Not glbOracle Then
    'Ticket #19554
    'SQLQ = SQLQ & " FROM Term_HROHSCOS LEFT JOIN Term_HREMP ON Term_HROHSCOS.CC_EMPNBR = Term_HREMP.ED_EMPNBR "
    SQLQ = SQLQ & " FROM Term_HROHSCOS LEFT JOIN Term_HREMP ON Term_HROHSCOS.TERM_SEQ = Term_HREMP.TERM_SEQ "
Else
    SQLQ = SQLQ & " FROM TERM_HROHSCOS, TERM_HREMP WHERE TERM_HROHSCOS.CC_EMPNBR = TERM_HREMP.ED_EMPNBR(+) "
End If
SQLQ = SQLQ & "WHERE CC_WCBNBR IS NOT NULL"

MDIMain.panHelp(0).FloodPercent = 60

gdbAdoIhr001X.BeginTrans
gdbAdoIhr001X.Execute SQLQ
gdbAdoIhr001X.CommitTrans

If Not glbSQL And Not glbOracle Then Call Pause(1)

'Commented by Bryan 03/Oct/05 Ticket#9419
'xFieldList = Get_Fields(gdbAdoIhr001W, "HROCCHSWRK", "WRKEMP")
'
'SQLQ = "INSERT INTO HROCCHSWRK (" & xFieldList & ",WRKEMP) "
'SQLQ = SQLQ & in_SQL(glbIHRDBW)
'SQLQ = SQLQ & " SELECT " & xFieldList & ",'" & glbUserID & "' AS WRKEMP FROM Term_HR_OCC_HEALTH_SAFETY"
'
'MDIMain.panHelp(0).FloodPercent = 75
'gdbAdoIhr001X.BeginTrans
'gdbAdoIhr001X.Execute SQLQ
'gdbAdoIhr001X.CommitTrans
'If Not glbSQL And Not glbOracle Then Call Pause(1)
'
'SQLQ = "INSERT INTO HROCCHSWRK (" & xFieldList & ",WRKEMP) "
'SQLQ = SQLQ & in_SQL(glbIHRDBW)
'SQLQ = SQLQ & " SELECT " & xFieldList & ",'" & glbUserID & "' AS WRKEMP FROM HR_OCC_HEALTH_SAFETY"
'
'MDIMain.panHelp(0).FloodPercent = 80
'gdbAdoIhr001.BeginTrans
'gdbAdoIhr001.Execute SQLQ
'gdbAdoIhr001.CommitTrans
'If Not glbSQL And Not glbOracle Then Call Pause(1)
'
gdbAdoIhr001.CommandTimeout = 600
MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""
Exit Sub

AttWrkError:
    gdbAdoIhr001.CommandTimeout = 600
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    'If Err.Number = 2147217871 Then MsgBox Err.Description
    ' dkostka - 04/18/01 - Not sure why the previous line was hiding errors,
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Exit Sub

End Sub

