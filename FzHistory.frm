VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmRHistory 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Employee History Report"
   ClientHeight    =   7605
   ClientLeft      =   375
   ClientTop       =   915
   ClientWidth     =   9960
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
   ScaleWidth      =   9960
   WindowState     =   2  'Maximized
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
      TabIndex        =   19
      Tag             =   "Check to Exclude Employees with RET Employment Status"
      Top             =   5820
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3315
   End
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
      TabIndex        =   18
      Tag             =   "Check to Exclude Employees with CONP Employment Status"
      Top             =   5520
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.CheckBox chkUseExcel 
      Caption         =   "Excel Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1860
      TabIndex        =   17
      Top             =   5160
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CheckBox chkUseDesc 
      Caption         =   "Use Descriptions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6360
      TabIndex        =   16
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CheckBox chkIncTerm 
      Caption         =   "Include Terminated"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4080
      TabIndex        =   15
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CheckBox chkShowUser 
      Caption         =   "Show Change by User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1860
      TabIndex        =   14
      Top             =   4800
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1860
      MaxLength       =   2
      TabIndex        =   12
      Tag             =   "00-Employee Position Shift"
      Top             =   4045
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txtChgType 
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
      Left            =   1860
      MaxLength       =   200
      TabIndex        =   13
      Tag             =   "00-Employee Position Shift"
      Top             =   4380
      Width           =   7195
   End
   Begin VB.CheckBox chkComments 
      Caption         =   "Show Comments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9120
      TabIndex        =   20
      Top             =   4800
      Visible         =   0   'False
      Width           =   2055
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
      Left            =   1875
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Tag             =   "Second level of grouping records"
      Top             =   7110
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
      Left            =   1875
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Tag             =   "First Level of grouping records"
      Top             =   6780
      Width           =   2325
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   7200
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   4
      Tag             =   "00-Enter Status Code"
      Top             =   1700
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
      Top             =   2035
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
      Top             =   1365
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
      Top             =   1030
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
      Top             =   695
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
      Index           =   4
      Left            =   1560
      TabIndex        =   8
      Tag             =   "00-Enter Administered By Code"
      Top             =   3040
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
      Left            =   1560
      TabIndex        =   9
      Tag             =   "00-Enter Section Code"
      Top             =   3375
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1560
      TabIndex        =   7
      Tag             =   "00-Enter Region Code"
      Top             =   2705
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Tag             =   "10-Enter Employee Number"
      Top             =   2370
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   3360
      TabIndex        =   11
      Tag             =   "40-Date upto and including this date forward"
      Top             =   3710
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   10
      Tag             =   "40-Date from and including this date forward"
      Top             =   3710
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Left            =   1560
      Picture         =   "FzHistory.frx":0000
      Top             =   4380
      Width           =   240
   End
   Begin VB.Label lblChgType 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Change Type"
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
      Top             =   4425
      Width           =   960
   End
   Begin VB.Label lblFromTo 
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
      TabIndex        =   38
      Top             =   3755
      Width           =   1095
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
      TabIndex        =   37
      Top             =   4090
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
      TabIndex        =   36
      Top             =   2080
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
      TabIndex        =   35
      Top             =   3420
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
      TabIndex        =   34
      Top             =   2750
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
      TabIndex        =   33
      Top             =   3085
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
      TabIndex        =   32
      Top             =   1075
      Width           =   615
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
      Left            =   60
      TabIndex        =   31
      Top             =   6555
      Width           =   1575
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
      Left            =   180
      TabIndex        =   30
      Top             =   7170
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
      Left            =   180
      TabIndex        =   29
      Top             =   6840
      Width           =   885
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
      TabIndex        =   27
      Top             =   2415
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
      TabIndex        =   26
      Top             =   1745
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
      TabIndex        =   25
      Top             =   1410
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
      TabIndex        =   24
      Top             =   740
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
      TabIndex        =   23
      Top             =   405
      Width           =   555
   End
End
Attribute VB_Name = "frmRHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr
 
If CriCheck() Then
    If Not PrtForm("Employee History Report Criteria", Me) Then Exit Sub
    Call set_PrintState(False)
    
    x% = Cri_SetAll()
    
    If glbWFC Then 'Ticket #27553 Franks 09/29/2015
        If chkUseExcel.Value Then
            Screen.MousePointer = DEFAULT
            Call set_PrintState(True)
            Exit Sub
        End If
    End If

    Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    Call set_PrintState(True)
End If

Exit Sub

PrntErr:
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString
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
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    Screen.MousePointer = HOURGLASS
    x% = Cri_SetAll()
    
    If glbWFC Then 'Ticket #27553 Franks 09/29/2015
        If chkUseExcel.Value Then
            Screen.MousePointer = DEFAULT
            Call set_PrintState(True)
            Exit Sub
        End If
    End If
    
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
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub comEmpType_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkExclCONP_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkExclRET_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGroup_GotFocus(Index As Integer)
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comChgTypeLoad()
'    lstChgType.AddItem "All"
'    lstChgType.AddItem lStr("")
'    lstChgType.AddItem lStr("Administered By")
'    lstChgType.AddItem lStr("Benefit Group")
'    lstChgType.AddItem lStr("Category")
'    lstChgType.AddItem lStr("Division")
'    lstChgType.AddItem lStr("Department")
'    lstChgType.AddItem lStr("FTE#")
'    lstChgType.AddItem lStr("FTE# Hours")
'    lstChgType.AddItem lStr("Location")
'    lstChgType.AddItem lStr("Union")
'    lstChgType.AddItem lStr("Region")
'    lstChgType.AddItem lStr("Section")
'    lstChgType.AddItem lStr("Status")
'
'    lstChgType.ListIndex = 0
End Sub

Private Sub comGrpLoad()
    ''comGroup(0).AddItem lStr("Division")
    ''comGroup(0).AddItem lStr("Department")
    ''comGroup(0).AddItem lStr("Location")
    ''comGroup(0).AddItem lStr("Union")
    ''comGroup(0).AddItem lStr("Employee Name")
    ''comGroup(0).AddItem lStr("Region")
    ''comGroup(0).AddItem lStr("Section")
    ''If Not glbMulti Then comGroup(0).AddItem "Shift"
    ''comGroup(0).AddItem "(none)"
    'Ticket #27687 Franks 10/27/2015 - Put the drop down in alphabetical order
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem "Employee Name"
    comGroup(0).AddItem lStr("Location")
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem lStr("Section")
    If Not glbMulti Then comGroup(0).AddItem "Shift"
    comGroup(0).AddItem lStr("Union")
    comGroup(0).AddItem "(none)"
    
    comGroup(1).AddItem "Employee Name"
    comGroup(1).AddItem "Employee Number"
    If Not glbSQL Then
        comGroup(1).AddItem "Date of Hire"
    End If
    comGroup(0).ListIndex = 0
    comGroup(1).ListIndex = 0
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
    Case 3: strCd$ = "HREMP.ED_REGION"
    Case 4: strCd$ = "HREMP.ED_ADMINBY"
    Case 5: strCd$ = "HREMP.ED_SECTION"
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

Private Sub Cri_ChgType()

Dim ChgTypeCri As String

If Len(glbCode) > 1 Then

    'ChgTypeCri = "({HREMPHIS_WRK.EE_HISTYPE} in ('" & glbCode & "'))"
    ChgTypeCri = "({HREMPHIS_WRK.EE_HISTYPE} in ['" & Replace(glbCode, ",", "','") & "'])"
    If InStr(1, ChgTypeCri, "Salary") > 0 Then ChgTypeCri = Replace(ChgTypeCri, "Salary", " ")
End If

If Len(ChgTypeCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = ChgTypeCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & ChgTypeCri
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

Cri_SetAll = False
On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

Call glbCri_DeptUN(clpDept.Text)

Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere

For x% = 0 To 5
    Call Cri_Code(x%)
Next x%
Call Cri_PT
Call Cri_Shift
Call Cri_EE
Call Cri_FTDates
Call Cri_ChgType
' Create query or view in database for display report
'If Not glbSQL And Not glbOracle Then
    Call SETWRK
'End If
    
If glbWFC Then 'Ticket #27553 Franks 09/29/2015
    If chkUseExcel.Value Then
        Call WFCExcelRpt
        Exit Function
    End If
End If
    
' report name
If comGroup(0) <> "(none)" Then
    strRName$ = glbIHRREPORTS & "rzEmpHistory.rpt"
Else
    strRName$ = glbIHRREPORTS & "rzEmpHistory1.rpt"
End If
  
Me.vbxCrystal.ReportFileName = strRName$
' set to sorting/grouping criteria
x% = Cri_Sorts()   ' returns number of sections formated

'Release 8.0 - Ticket #22682: View Own security
'If View Own not checked then do not retrieve Employee History of the User/Employee No
If Len(glbUserEmpNo) > 0 And glbUserEmpNo <> 0 And Not gSec_EmpHis_ViewOwn Then
    'Do not show user's Employee History records based on the Employee # associated to the User.
    If Len(glbstrSelCri) > 0 Then
        glbstrSelCri = glbstrSelCri & " AND {HREMPHIS_WRK.EE_EMPNBR} <> " & glbUserEmpNo
    Else
        glbstrSelCri = glbstrSelCri & " {HREMPHIS_WRK.EE_EMPNBR} <> " & glbUserEmpNo
    End If
End If

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
    'Ticket #16146
    'Me.vbxCrystal.SelectionFormula = glbstrSelCri
    Me.vbxCrystal.SelectionFormula = glbstrSelCri & " AND " & "{HREMPHIS_WRK.EE_WRKEMP}='" & glbUserID & "'"
End If

'Ticket #27553 Franks 09/28/2015
Me.vbxCrystal.Formulas(10) = "ShowUser = " & IIf(chkShowUser.Value, True, False) & " "

If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    For x% = 0 To 6
        Me.vbxCrystal.DataFiles(x%) = glbIHRDB
    Next x%
    Me.vbxCrystal.DataFiles(7) = glbIHRDBW
    Me.vbxCrystal.DataFiles(8) = glbIHRDB
    ' set security for database
    'Me.vbxCrystal.Password = gstrAccPWord$
    'Me.vbxCrystal.UserName = gstrAccUID$
End If

' window title if appropriate
Me.vbxCrystal.WindowTitle = "Employee History Report"

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
'imbeded in report

Cri_Sorts = 0
'first set primary grouping
Y% = 0
grpField$ = getEGroup(comGroup(0).Text)

If comGroup(0) = "(none)" Then
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
        Case 1: grpField$ = "{HREMP.ED_EMPNBR}" 'GROUP ON EMPLOYEE#
    End Select
    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$

    Exit Function
End If
    
Y% = x% + 1
dscGroup$ = comGroup(x%).Text
dscGroup$ = "descGroup" & CStr(Y%) & "= '" & dscGroup$ & "'"
Me.vbxCrystal.Formulas(x%) = dscGroup$

grpCond$ = "GROUP" & CStr(Y%) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(x%) = grpCond$

strSFormat$ = "GH1;T;T;X;X;X;X;X"
Me.vbxCrystal.SectionFormat(z%) = strSFormat$
z% = z% + 1
strSFormat$ = "GF1;T;X;X;X;X;X;X"
Me.vbxCrystal.SectionFormat(z%) = strSFormat$
z% = z% + 1

'GrpIdx% = comGroup(1).ListIndex
'grpField$ = "{@EFullName}"
'grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
'Me.vbxCrystal.GroupCondition(1) = grpCond$

GrpIdx% = comGroup(1).ListIndex
Select Case GrpIdx%
    Case 0: grpField$ = "{@EFullName}"
    Case 1: grpField$ = "{HREMP.ED_EMPNBR}" 'GROUP ON EMPLOYEE#
End Select
grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(1) = grpCond$


Cri_Sorts = z% ' next section number to format

End Function

Private Function Cri_SortsDepend()
Dim grpCond$, grpField$
Dim x%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

'for labels - sort by name always
'imbeded in report

Cri_SortsDepend = 0
'first set primary grouping
Y% = 0
grpField$ = getEGroup(comGroup(0).Text)
If comGroup(0) = "(none)" Then Exit Function
    
Y% = x% + 1
dscGroup$ = comGroup(x%).Text
dscGroup$ = "descGroup" & CStr(Y%) & "= '" & dscGroup$ & "'"
Me.vbxCrystal.Formulas(x%) = dscGroup$

grpCond$ = "GROUP" & CStr(Y%) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(x%) = grpCond$

strSFormat$ = "GH1;T;T;X;X;X;X;X"
Me.vbxCrystal.SectionFormat(z%) = strSFormat$
z% = z% + 1
strSFormat$ = "GF1;T;X;X;X;X;X;X"
Me.vbxCrystal.SectionFormat(z%) = strSFormat$
z% = z% + 1

GrpIdx% = comGroup(1).ListIndex
Select Case GrpIdx%
    Case 0: grpField$ = "{@EFullName}"
End Select
grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(1) = grpCond$

Cri_SortsDepend = z% ' next section number to format

End Function

Private Sub Cri_FTDates()
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim x%
Dim EECri As String, LocCri As String

If Len(dlpDateRange(0).Text) = 0 And Len(dlpDateRange(1).Text) = 0 Then Exit Sub
If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    TempCri = "({HREMPHIS_WRK.EE_CHGDATE} "
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    GoTo Cri_FTDatst
ElseIf Len(dlpDateRange(0).Text) > 0 Then    ' Daniel - 10/20/1999
    TempCri = "({HREMPHIS_WRK.EE_CHGDATE} "         ' Added section to enable entering only From date, no To date.
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    TempCri = TempCri & " >= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "    'Hemu - 07/02/2003, Included '='
    GoTo Cri_FTDatst
ElseIf Len(dlpDateRange(1).Text) > 0 Then    ' Daniel - 10/20/1999
    TempCri = "({HREMPHIS_WRK.EE_CHGDATE} "         ' Added section to enable entering only To date, no From date.
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    TempCri = TempCri & " <= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "    'Hemu - 07/02/2003, Included '='
    GoTo Cri_FTDatst
End If

For x% = 0 To 1
    If Len(dlpDateRange(0).Text) > 0 Then
        TempCri = "({HREMPHIS_WRK.EE_CHGDATE}  "
        If x% = 0 Then
            TempCri = TempCri & " >= "
        Else
            TempCri = TempCri & " <= "
        End If
        dtYYY% = Year(dlpDateRange(0).Text)
        dtMM% = month(dlpDateRange(0).Text)
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

If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

If IsDate(dlpDateRange(0)) And IsDate(dlpDateRange(1)) Then
    If DaysBetween(dlpDateRange(0), dlpDateRange(1)) < 0 Then                               'Serbo
        MsgBox "To Date can't be prior to From Date!"                       '
        Me.dlpDateRange(0).SetFocus                                         '
        Exit Function                                                       '
    End If
End If

For x% = 0 To 5
    If Not clpCode(x).ListChecker Then Exit Function
Next x%

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

glbOnTop = "FRMRHISTORY"

Screen.MousePointer = HOURGLASS

If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If

Call comChgTypeLoad
Call comGrpLoad
Call setRptCaption(Me)

frmRHistory.Caption = "Employee History Report"

If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6

Call INI_Controls(Me)

If glbWFC Then 'Ticket #27553 Franks
    chkUseExcel.Visible = True
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

Private Sub imgIcon_Click()
    txtChgType_DblClick
End Sub

Private Sub imgIcon_DblClick()
    txtChgType_DblClick
End Sub

Private Sub txtChgType_DblClick()
Dim SQLQ As String

On Error GoTo ChangeType_Err

Load frmCHANGETYPE

frmCHANGETYPE.Show 1
txtChgType = glbCode
Exit Sub

ChangeType_Err:
glbFrmCaption$ = lStr("Get Change Types")
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "", "", "")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

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

Private Sub SETWRK()
Dim rsHS As New ADODB.Recordset, rsTmp As New ADODB.Recordset
Dim rsCode As New ADODB.Recordset, rsBODY As New ADODB.Recordset
Dim SQLQ, xNum, xRecNum
Dim xFieldList
xFieldList = Get_Fields(gdbAdoIhr001W, "HREMPHIS_WRK", "KEY_EMPNBR,EE_ID,EE_GROUP, EE_GROUP_DESC,")
'xFieldList = ""

'On Error GoTo AttWrkError
MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 1
MDIMain.panHelp(0).FloodPercent = 3
gdbAdoIhr001.CommandTimeout = 600
Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).FloodPercent = 15

gdbAdoIhr001W.BeginTrans
gdbAdoIhr001W.Execute "DELETE FROM HREMPHIS_WRK WHERE EE_WRKEMP='" & glbUserID & "'"
gdbAdoIhr001W.CommitTrans

Call Pause(1)

MDIMain.panHelp(0).FloodPercent = 30
'for active employees
SQLQ = "INSERT INTO HREMPHIS_WRK  (" & xFieldList & ",KEY_EMPNBR) " & in_SQL(glbIHRDBW)

SQLQ = SQLQ & " Select EE_COMPNO,EE_EMPNBR, "
Select Case gsSystemDb
Case "MS SQL SERVER"
    'SALARY
    SQLQ = SQLQ & " (CASE WHEN EE_SALARY <> 0 THEN cast(EE_SALARY AS varchar(20)) "
    SQLQ = SQLQ & " ELSE ' ' END) AS EE_SALARY, "
    'PER
    SQLQ = SQLQ & " (CASE WHEN EE_SALARY <> 0 THEN EE_SALCD "
    SQLQ = SQLQ & " ELSE '' END) AS EE_SALCD, "
    
    SQLQ = SQLQ & " EE_CHGDATE,EE_LDATE,EE_LTIME,EE_LUSER,EE_DOT, NULL, " 'TERM_SEQ, "
    'TYPE
    SQLQ = SQLQ & " (CASE WHEN EE_NEWDEPT IS NOT NULL THEN '" & lStr("Department") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWDIV IS NOT NULL THEN '" & lStr("Division") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWSTAT IS NOT NULL THEN '" & lStr("Status") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWPT IS NOT NULL THEN '" & lStr("Category") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWORG IS NOT NULL THEN '" & lStr("Union") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWFTE IS NOT NULL THEN '" & lStr("FTE#") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWFTEHR IS NOT NULL THEN '" & lStr("FTE# Hours") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWREGION IS NOT NULL THEN '" & lStr("Region") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWSECTION IS NOT NULL THEN '" & lStr("Section") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWADMINBY IS NOT NULL THEN '" & lStr("Administered By") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWLOC IS NOT NULL THEN '" & lStr("Location") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWBENEGROUP IS NOT NULL THEN '" & lStr("Benefit Group") & " ' "
    If glbWFC Then 'Ticket #21118 Franks 10/31/2011
        SQLQ = SQLQ & " WHEN EE_NEWSMOKER IS NOT NULL THEN '" & ("Smoker") & "' "
        'Ticket #28794 - Opening up Marital Status for everyone
        'SQLQ = SQLQ & " WHEN EE_NEWMSTAT IS NOT NULL THEN '" & ("Marital Status") & "' "
    End If
    'Ticket #28794 - Opening up Marital Status for everyone
    SQLQ = SQLQ & " WHEN EE_NEWMSTAT IS NOT NULL THEN '" & ("Marital Status") & "' "
    
    SQLQ = SQLQ & " WHEN EE_NEWPOSITION IS NOT NULL THEN '" & "Position" & "' " 'Ticket #27553 Franks 09/21/2015
    SQLQ = SQLQ & " WHEN EE_NEWREPORT1 IS NOT NULL THEN '" & lStr("Rept. Authority 1") & "' " 'Ticket #27553 Franks 09/22/2015
    'SQLQ = SQLQ & " WHEN EE_SALARY <> 0 THEN 'Salary' "
    SQLQ = SQLQ & " END) AS EE_HISTYPE, "
    
    
    'OLDVALUE
    SQLQ = SQLQ & " (CASE WHEN EE_NEWDEPT IS NOT NULL THEN EE_OLDDEPT "
    SQLQ = SQLQ & " WHEN EE_NEWDIV IS NOT NULL THEN EE_OLDDIV "
    SQLQ = SQLQ & " WHEN EE_NEWSTAT IS NOT NULL THEN EE_OLDSTAT "
    SQLQ = SQLQ & " WHEN EE_NEWPT IS NOT NULL THEN EE_OLDPT "
    SQLQ = SQLQ & " WHEN EE_NEWORG IS NOT NULL THEN EE_OLDORG "
    SQLQ = SQLQ & " WHEN EE_NEWFTE IS NOT NULL THEN cast(EE_OLDFTE AS varchar(20)) "
    SQLQ = SQLQ & " WHEN EE_NEWFTEHR IS NOT NULL THEN cast(EE_OLDFTEHR AS varchar(20)) "
    SQLQ = SQLQ & " WHEN EE_NEWREGION IS NOT NULL THEN EE_OLDREGION "
    SQLQ = SQLQ & " WHEN EE_NEWSECTION IS NOT NULL THEN EE_OLDSECTION "
    SQLQ = SQLQ & " WHEN EE_NEWADMINBY IS NOT NULL THEN EE_OLDADMINBY "
    SQLQ = SQLQ & " WHEN EE_NEWLOC IS NOT NULL THEN EE_OLDLOC "
    SQLQ = SQLQ & " WHEN EE_NEWBENEGROUP IS NOT NULL THEN EE_OLDBENEGROUP "
    If glbWFC Then '#21118 Franks 10/31/2011
        SQLQ = SQLQ & " WHEN EE_NEWSMOKER IS NOT NULL THEN EE_OLDSMOKER "
        'Ticket #28794 - Opening up Marital Status for everyone
        'SQLQ = SQLQ & " WHEN EE_NEWMSTAT IS NOT NULL THEN EE_OLDMSTAT "
    End If
    'Ticket #28794 - Opening up Marital Status for everyone
    SQLQ = SQLQ & " WHEN EE_NEWMSTAT IS NOT NULL THEN EE_OLDMSTAT "

    SQLQ = SQLQ & " WHEN EE_NEWPOSITION IS NOT NULL THEN EE_OLDPOSITION " 'Ticket #27553 Franks 09/21/2015
    SQLQ = SQLQ & " WHEN EE_NEWREPORT1 IS NOT NULL THEN EE_OLDREPORT1 " 'Ticket #27553 Franks 09/21/2015
    SQLQ = SQLQ & " WHEN EE_SALARY IS NOT NULL THEN '' "
    SQLQ = SQLQ & " END) AS EE_OLDVALUE, "
    'NEWVALUE
    SQLQ = SQLQ & " (CASE WHEN EE_NEWDEPT IS NOT NULL THEN EE_NEWDEPT "
    SQLQ = SQLQ & " WHEN EE_NEWDIV IS NOT NULL THEN EE_NEWDIV "
    SQLQ = SQLQ & " WHEN EE_NEWSTAT IS NOT NULL THEN EE_NEWSTAT "
    SQLQ = SQLQ & " WHEN EE_NEWPT IS NOT NULL THEN EE_NEWPT "
    SQLQ = SQLQ & " WHEN EE_NEWORG IS NOT NULL THEN EE_NEWORG "
    SQLQ = SQLQ & " WHEN EE_NEWFTE IS NOT NULL THEN cast(EE_NEWFTE AS varchar(20)) "
    SQLQ = SQLQ & " WHEN EE_NEWFTEHR IS NOT NULL THEN cast(EE_NEWFTEHR AS varchar(20)) "
    SQLQ = SQLQ & " WHEN EE_NEWREGION IS NOT NULL THEN EE_NEWREGION "
    SQLQ = SQLQ & " WHEN EE_NEWSECTION IS NOT NULL THEN EE_NEWSECTION "
    SQLQ = SQLQ & " WHEN EE_NEWADMINBY IS NOT NULL THEN EE_NEWADMINBY "
    SQLQ = SQLQ & " WHEN EE_NEWLOC IS NOT NULL THEN EE_NEWLOC "
    SQLQ = SQLQ & " WHEN EE_NEWBENEGROUP IS NOT NULL THEN EE_NEWBENEGROUP "
    If glbWFC Then '#21118 Franks 10/28/2011
        SQLQ = SQLQ & " WHEN EE_NEWSMOKER IS NOT NULL THEN EE_NEWSMOKER  "
        'Ticket #28794 - Opening up Marital Status for everyone
        'SQLQ = SQLQ & " WHEN EE_NEWMSTAT IS NOT NULL THEN EE_NEWMSTAT "
    End If
    'Ticket #28794 - Opening up Marital Status for everyone
    SQLQ = SQLQ & " WHEN EE_NEWMSTAT IS NOT NULL THEN EE_NEWMSTAT "
    
    SQLQ = SQLQ & " WHEN EE_NEWPOSITION IS NOT NULL THEN EE_NEWPOSITION " 'Ticket #27553 Franks 09/21/2015
    SQLQ = SQLQ & " WHEN EE_NEWREPORT1 IS NOT NULL THEN EE_NEWREPORT1 " 'Ticket #27553 Franks 09/21/2015
    SQLQ = SQLQ & " WHEN EE_SALARY IS NOT NULL THEN '' "
    SQLQ = SQLQ & " END) AS EE_NEWVALUE "
    SQLQ = SQLQ & ",'" & glbUserID & "' AS EE_WRKEMP "
    SQLQ = SQLQ & ",'1_' + LTRIM(RTRIM(STR(HREMP.ED_EMPNBR)) ) AS KEY_EMPNBR "
Case "ORACLE"
    'SALARY
    SQLQ = SQLQ & " (CASE WHEN EE_SALARY <> 0 THEN TO_CHAR(EE_SALARY) "
    SQLQ = SQLQ & " ELSE ' ' END) AS EE_SALARY, "
    'PER
    SQLQ = SQLQ & " (CASE WHEN EE_SALARY <> 0 THEN EE_SALCD "
    SQLQ = SQLQ & " ELSE '' END) AS EE_SALCD, "
    
    SQLQ = SQLQ & " EE_CHGDATE,EE_LDATE,EE_LTIME,EE_LUSER,EE_DOT,NULL, " ' TERM_SEQ, "
    'TYPE
    SQLQ = SQLQ & " (CASE WHEN EE_NEWDEPT IS NOT NULL THEN '" & lStr("Department") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWDIV IS NOT NULL THEN '" & lStr("Division") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWSTAT IS NOT NULL THEN '" & lStr("Status") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWPT IS NOT NULL THEN '" & lStr("Category") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWORG IS NOT NULL THEN '" & lStr("Union") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWFTE IS NOT NULL THEN '" & lStr("FTE#") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWFTEHR IS NOT NULL THEN '" & lStr("FTE# Hours") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWREGION IS NOT NULL THEN '" & lStr("Region") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWSECTION IS NOT NULL THEN '" & lStr("Section") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWADMINBY IS NOT NULL THEN '" & lStr("Administered By") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWLOC IS NOT NULL THEN '" & lStr("Location") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWBENEGROUP IS NOT NULL THEN '" & lStr("Benefit Group") & " ' "
    'SQLQ = SQLQ & " WHEN EE_SALARY <> 0 THEN 'Salary' "
    SQLQ = SQLQ & " END) AS EE_HISTYPE, "
    
    
    'OLDVALUE
    SQLQ = SQLQ & " (CASE WHEN EE_NEWDEPT IS NOT NULL THEN EE_OLDDEPT "
    SQLQ = SQLQ & " WHEN EE_NEWDIV IS NOT NULL THEN EE_OLDDIV "
    SQLQ = SQLQ & " WHEN EE_NEWSTAT IS NOT NULL THEN EE_OLDSTAT "
    SQLQ = SQLQ & " WHEN EE_NEWPT IS NOT NULL THEN EE_OLDPT "
    SQLQ = SQLQ & " WHEN EE_NEWORG IS NOT NULL THEN EE_OLDORG "
    SQLQ = SQLQ & " WHEN EE_NEWFTE IS NOT NULL THEN TO_CHAR(EE_OLDFTE) "
    SQLQ = SQLQ & " WHEN EE_NEWFTEHR IS NOT NULL THEN TO_CHAR(EE_OLDFTEHR) "
    SQLQ = SQLQ & " WHEN EE_NEWREGION IS NOT NULL THEN EE_OLDREGION "
    SQLQ = SQLQ & " WHEN EE_NEWSECTION IS NOT NULL THEN EE_OLDSECTION "
    SQLQ = SQLQ & " WHEN EE_NEWADMINBY IS NOT NULL THEN EE_OLDADMINBY "
    SQLQ = SQLQ & " WHEN EE_NEWLOC IS NOT NULL THEN EE_OLDLOC "
    SQLQ = SQLQ & " WHEN EE_NEWBENEGROUP IS NOT NULL THEN EE_OLDBENEGROUP "
    SQLQ = SQLQ & " WHEN EE_SALARY IS NOT NULL THEN '' "
    SQLQ = SQLQ & " END) AS EE_OLDVALUE, "
    'NEWVALUE
    SQLQ = SQLQ & " (CASE WHEN EE_NEWDEPT IS NOT NULL THEN EE_NEWDEPT "
    SQLQ = SQLQ & " WHEN EE_NEWDIV IS NOT NULL THEN EE_NEWDIV "
    SQLQ = SQLQ & " WHEN EE_NEWSTAT IS NOT NULL THEN EE_NEWSTAT "
    SQLQ = SQLQ & " WHEN EE_NEWPT IS NOT NULL THEN EE_NEWPT "
    SQLQ = SQLQ & " WHEN EE_NEWORG IS NOT NULL THEN EE_NEWORG "
    SQLQ = SQLQ & " WHEN EE_NEWFTE IS NOT NULL THEN TO_CHAR(EE_NEWFTE) "
    SQLQ = SQLQ & " WHEN EE_NEWFTEHR IS NOT NULL THEN TO_CHAR(EE_NEWFTEHR) "
    SQLQ = SQLQ & " WHEN EE_NEWREGION IS NOT NULL THEN EE_NEWREGION "
    SQLQ = SQLQ & " WHEN EE_NEWSECTION IS NOT NULL THEN EE_NEWSECTION "
    SQLQ = SQLQ & " WHEN EE_NEWADMINBY IS NOT NULL THEN EE_NEWADMINBY "
    SQLQ = SQLQ & " WHEN EE_NEWLOC IS NOT NULL THEN EE_NEWLOC "
    SQLQ = SQLQ & " WHEN EE_NEWBENEGROUP IS NOT NULL THEN EE_NEWBENEGROUP "
    SQLQ = SQLQ & " WHEN EE_SALARY IS NOT NULL THEN '' "
    SQLQ = SQLQ & " END) AS EE_NEWVALUE "
    SQLQ = SQLQ & ",'" & glbUserID & "' AS EE_WRKEMP "
    SQLQ = SQLQ & ",'1_' + LTRIM(STR(HREMP.ED_EMPNBR) ) AS KEY_EMPNBR "
Case Else
    SQLQ = SQLQ & " IIF( EE_SALARY <> 0 , str(EE_SALARY), '') AS SALARY,"
    SQLQ = SQLQ & " IIF( EE_SALARY <> 0 , EE_SALCD , '') AS SALCD,"
    SQLQ = SQLQ & " EE_CHGDATE,EE_LDATE,EE_LTIME,EE_LUSER,EE_DOT,NULL, " 'TERM_SEQ, "
    
    SQLQ = SQLQ & " IIF(EE_NEWDEPT IS NOT NULL , 'Department ' ,"
    SQLQ = SQLQ & " IIF(EE_NEWDIV IS NOT NULL , 'Division ' ,"
    SQLQ = SQLQ & " IIF(EE_NEWSTAT IS NOT NULL , 'Status ' ,"
    SQLQ = SQLQ & " IIF(EE_NEWPT IS NOT NULL , 'FT/PT/SE/TR/OT ' ,"
    SQLQ = SQLQ & " IIF(EE_NEWORG IS NOT NULL , 'Union ' ,"
    SQLQ = SQLQ & " IIF(EE_NEWFTE IS NOT NULL , 'FTE# ' ,"
    SQLQ = SQLQ & " IIF(EE_NEWFTEHR IS NOT NULL , 'FTE# Hours ' ,"
    SQLQ = SQLQ & " IIF(EE_NEWREGION IS NOT NULL , 'Region ' ,"
    SQLQ = SQLQ & " IIF(EE_NEWSECTION IS NOT NULL , 'Section ',"
    SQLQ = SQLQ & " IIF(EE_NEWADMINBY IS NOT NULL , 'Administered By ' ,"
    SQLQ = SQLQ & " IIF(EE_NEWLOC IS NOT NULL , 'Location '  ,"
    SQLQ = SQLQ & " IIF(EE_NEWBENEGROUP IS NOT NULL , 'Benefit Group', '')"
    SQLQ = SQLQ & " )))))))))))"
    SQLQ = SQLQ & " AS EE_HISTYPE,"
    
    SQLQ = SQLQ & " IIF(EE_NEWDEPT IS NOT NULL , EE_OLDDEPT ,"
    SQLQ = SQLQ & " IIF(EE_NEWDIV IS NOT NULL , EE_OLDDIV ,"
    SQLQ = SQLQ & " IIF(EE_NEWSTAT IS NOT NULL , EE_OLDSTAT ,"
    SQLQ = SQLQ & " IIF(EE_NEWPT IS NOT NULL , EE_OLDPT ,"
    SQLQ = SQLQ & " IIF(EE_NEWORG IS NOT NULL , EE_OLDORG,"
    SQLQ = SQLQ & " IIF(EE_NEWFTE IS NOT NULL , str(EE_OLDFTE ),"
    SQLQ = SQLQ & " IIF(EE_NEWFTEHR IS NOT NULL , str(EE_OLDFTEHR),"
    SQLQ = SQLQ & " IIF(EE_NEWREGION IS NOT NULL , EE_OLDREGION,"
    SQLQ = SQLQ & " IIF(EE_NEWSECTION IS NOT NULL , EE_OLDSECTION,"
    SQLQ = SQLQ & " IIF(EE_NEWADMINBY IS NOT NULL , EE_OLDADMINBY ,"
    SQLQ = SQLQ & " IIF( EE_NEWLOC IS NOT NULL , EE_OLDLOC ,"
    SQLQ = SQLQ & " IIF( EE_NEWBENEGROUP IS NOT NULL , EE_OLDBENEGROUP ,"
    SQLQ = SQLQ & " IIF( EE_SALARY IS NOT NULL , ''  ,'')"
    SQLQ = SQLQ & " ))))))))))))"
    SQLQ = SQLQ & "  AS EE_OLDVALUE,"
    
    SQLQ = SQLQ & " IIF( EE_NEWDEPT IS NOT NULL , EE_NEWDEPT  ,"
    SQLQ = SQLQ & " IIF( EE_NEWDIV IS NOT NULL , EE_NEWDIV  ,"
    SQLQ = SQLQ & " IIF( EE_NEWSTAT IS NOT NULL , EE_NEWSTAT  ,"
    SQLQ = SQLQ & " IIF( EE_NEWPT IS NOT NULL , EE_NEWPT  ,"
    SQLQ = SQLQ & " IIF( EE_NEWORG IS NOT NULL , EE_NEWORG  ,"
    SQLQ = SQLQ & " IIF( EE_NEWFTE IS NOT NULL , str(EE_NEWFTE),"
    SQLQ = SQLQ & " IIF( EE_NEWFTEHR IS NOT NULL , str(EE_NEWFTEHR),"
    SQLQ = SQLQ & " IIF( EE_NEWREGION IS NOT NULL , EE_NEWREGION  ,"
    SQLQ = SQLQ & " IIF( EE_NEWSECTION IS NOT NULL , EE_NEWSECTION,"
    SQLQ = SQLQ & " IIF(EE_NEWADMINBY IS NOT NULL , EE_NEWADMINBY  ,"
    SQLQ = SQLQ & " IIF( EE_NEWLOC IS NOT NULL , EE_NEWLOC  ,"
    SQLQ = SQLQ & " IIF( EE_NEWBENEGROUP IS NOT NULL , EE_NEWBENEGROUP  ,"
    SQLQ = SQLQ & " IIF( EE_SALARY IS NOT NULL , '' ,'' )"
    SQLQ = SQLQ & " ))))))))))))"
    SQLQ = SQLQ & " AS EE_NEWVALUE"
    SQLQ = SQLQ & ",'" & glbUserID & "' AS EE_WRKEMP "
    SQLQ = SQLQ & ",'1_' + TRIM(STR(HREMP.ED_EMPNBR) ) AS KEY_EMPNBR "
End Select
'If glbtermopen Then
'    SQLQ = SQLQ & " FROM Term_HREMPHIS "
'    SQLQ = SQLQ & " WHERE Term_SEQ = " & glbTERM_Seq
'Else
'    SQLQ = SQLQ & " FROM HREMPHIS "
'    SQLQ = SQLQ & " WHERE EE_EMPNBR = " & glbLEE_ID
'End If
'fglbSQL = SQLQ

'Ticket #27553 Franks 09/28/2015 - begin
'SQLQ = SQLQ & " FROM HREMPHIS "
SQLQ = SQLQ & " FROM HREMPHIS LEFT JOIN HREMP ON HREMPHIS.EE_EMPNBR = HREMP.ED_EMPNBR "
SQLQ = SQLQ & " WHERE (1=1) "
'Hemu - 09/29/2003 Begin - Taking too long to generate the report, esp. when Emp#
'                          entered
If Len(elpEEID.Text) > 0 Then
    SQLQ = SQLQ & "AND HREMPHIS.EE_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
End If
'Hemu
If Len(clpDiv.Text) > 0 Then
    SQLQ = SQLQ & "AND ED_DIV IN ('" & Replace(clpDiv.Text, ",", "','") & "') "
End If
If Len(clpDept.Text) > 0 Then
    SQLQ = SQLQ & "AND ED_DEPTNO IN ('" & Replace(clpDept.Text, ",", "','") & "') "
End If
If Len(clpCode(0).Text) > 0 Then
    SQLQ = SQLQ & "AND ED_LOC IN ('" & Replace(clpCode(0).Text, ",", "','") & "') "
End If
If Len(clpCode(1).Text) > 0 Then
    SQLQ = SQLQ & "AND ED_ORG IN ('" & Replace(clpCode(1).Text, ",", "','") & "') "
End If
If Len(clpCode(2).Text) > 0 Then
    SQLQ = SQLQ & "AND ED_EMP IN ('" & Replace(clpCode(2).Text, ",", "','") & "') "
End If
If Len(clpPT.Text) > 0 Then
    SQLQ = SQLQ & "AND ED_PT IN ('" & Replace(clpPT.Text, ",", "','") & "') "
End If
If Len(clpCode(3).Text) > 0 Then
    SQLQ = SQLQ & "AND ED_REGION IN ('" & Replace(clpCode(3).Text, ",", "','") & "') "
End If
If Len(clpCode(4).Text) > 0 Then
    SQLQ = SQLQ & "AND ED_ADMINBY IN ('" & Replace(clpCode(4).Text, ",", "','") & "') "
End If
If Len(clpCode(5).Text) > 0 Then
    SQLQ = SQLQ & "AND ED_SECTION IN ('" & Replace(clpCode(5).Text, ",", "','") & "') "
End If
If IsDate(dlpDateRange(0).Text) Then
    SQLQ = SQLQ & "AND HREMPHIS.EE_CHGDATE >= " & Date_SQL(dlpDateRange(0).Text) & " "
End If
If IsDate(dlpDateRange(1).Text) Then
    SQLQ = SQLQ & "AND HREMPHIS.EE_CHGDATE <= " & Date_SQL(dlpDateRange(1).Text) & " "
End If
'Ticket #27553 Franks 09/28/2015 - end

'Ticket #29660 - Contract Employees Enhancement
If glbWFC Then
    If chkExclCONP.Visible And chkExclRET.Visible = True Then
        If chkExclCONP Then
            SQLQ = SQLQ & "AND ED_EMP <> 'CONP'"
        End If
        If chkExclRET Then
            SQLQ = SQLQ & "AND ED_EMP <> 'RET'"
        End If
    End If
End If

gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans
MDIMain.panHelp(0).FloodPercent = 50

'for terminated employees ********************************************
If chkIncTerm.Value Then 'Ticket #27553 Franks 09/28/2015
    SQLQ = "INSERT INTO HREMPHIS_WRK (" & xFieldList & ",KEY_EMPNBR) " & in_SQL(glbIHRDBW)
    
    SQLQ = SQLQ & "Select EE_COMPNO,EE_EMPNBR, "
    Select Case gsSystemDb
    Case "MS SQL SERVER"
        'SALARY
        SQLQ = SQLQ & " (CASE WHEN EE_SALARY <> 0 THEN cast(EE_SALARY AS varchar(20)) "
        SQLQ = SQLQ & " ELSE '' END) AS EE_SALARY, "
        'PER
        SQLQ = SQLQ & " (CASE WHEN EE_SALARY <> 0 THEN EE_SALCD "
        SQLQ = SQLQ & " ELSE '' END) AS EE_SALCD, "
        
        SQLQ = SQLQ & "EE_CHGDATE,EE_LDATE,EE_LTIME,EE_LUSER,EE_DOT,Term_HREMPHIS.TERM_SEQ, "
        'TYPE
        SQLQ = SQLQ & " (CASE WHEN EE_NEWDEPT IS NOT NULL THEN '" & lStr("Department") & " ' "
        SQLQ = SQLQ & " WHEN EE_NEWDIV IS NOT NULL THEN '" & lStr("Division") & " ' "
        SQLQ = SQLQ & " WHEN EE_NEWSTAT IS NOT NULL THEN '" & lStr("Status") & " ' "
        SQLQ = SQLQ & " WHEN EE_NEWPT IS NOT NULL THEN '" & lStr("Category") & " ' "
        SQLQ = SQLQ & " WHEN EE_NEWORG IS NOT NULL THEN '" & lStr("Union") & " ' "
        SQLQ = SQLQ & " WHEN EE_NEWFTE IS NOT NULL THEN '" & lStr("FTE#") & " ' "
        SQLQ = SQLQ & " WHEN EE_NEWFTEHR IS NOT NULL THEN '" & lStr("FTE# Hours") & " ' "
        SQLQ = SQLQ & " WHEN EE_NEWREGION IS NOT NULL THEN '" & lStr("Region") & " ' "
        SQLQ = SQLQ & " WHEN EE_NEWSECTION IS NOT NULL THEN '" & lStr("Section") & " ' "
        SQLQ = SQLQ & " WHEN EE_NEWADMINBY IS NOT NULL THEN '" & lStr("Administered By") & " ' "
        SQLQ = SQLQ & " WHEN EE_NEWLOC IS NOT NULL THEN '" & lStr("Location") & " ' "
        SQLQ = SQLQ & " WHEN EE_NEWBENEGROUP IS NOT NULL THEN '" & lStr("Benefit Group") & " ' "
        If glbWFC Then 'Ticket #21118 Franks 10/31/2011
            SQLQ = SQLQ & " WHEN EE_NEWSMOKER IS NOT NULL THEN '" & ("Smoker") & "' "
            'Ticket #28794 - Opening up Marital Status for everyone
            'SQLQ = SQLQ & " WHEN EE_NEWMSTAT IS NOT NULL THEN '" & ("Marital Status") & "' "
        End If
        'Ticket #28794 - Opening up Marital Status for everyone
        SQLQ = SQLQ & " WHEN EE_NEWMSTAT IS NOT NULL THEN '" & ("Marital Status") & "' "
        
        SQLQ = SQLQ & " WHEN EE_NEWPOSITION IS NOT NULL THEN '" & "Position" & "' " 'Ticket #27553 Franks 09/21/2015
        SQLQ = SQLQ & " WHEN EE_NEWREPORT1 IS NOT NULL THEN '" & lStr("Rept. Authority 1") & "' " 'Ticket #27553 Franks 09/22/2015
        'SQLQ = SQLQ & " WHEN EE_SALARY <> 0 THEN 'Salary' "
        SQLQ = SQLQ & " END) AS EE_HISTYPE, "
    
    
        'OLDVALUE
        SQLQ = SQLQ & " (CASE WHEN EE_NEWDEPT IS NOT NULL THEN EE_OLDDEPT "
        SQLQ = SQLQ & " WHEN EE_NEWDIV IS NOT NULL THEN EE_OLDDIV "
        SQLQ = SQLQ & " WHEN EE_NEWSTAT IS NOT NULL THEN EE_OLDSTAT "
        SQLQ = SQLQ & " WHEN EE_NEWPT IS NOT NULL THEN EE_OLDPT "
        SQLQ = SQLQ & " WHEN EE_NEWORG IS NOT NULL THEN EE_OLDORG "
        SQLQ = SQLQ & " WHEN EE_NEWFTE IS NOT NULL THEN cast(EE_OLDFTE AS varchar(20)) "
        SQLQ = SQLQ & " WHEN EE_NEWFTEHR IS NOT NULL THEN cast(EE_OLDFTEHR AS varchar(20)) "
        SQLQ = SQLQ & " WHEN EE_NEWREGION IS NOT NULL THEN EE_OLDREGION "
        SQLQ = SQLQ & " WHEN EE_NEWSECTION IS NOT NULL THEN EE_OLDSECTION "
        SQLQ = SQLQ & " WHEN EE_NEWADMINBY IS NOT NULL THEN EE_OLDADMINBY "
        SQLQ = SQLQ & " WHEN EE_NEWLOC IS NOT NULL THEN EE_OLDLOC "
        SQLQ = SQLQ & " WHEN EE_NEWBENEGROUP IS NOT NULL THEN EE_OLDBENEGROUP "
        If glbWFC Then '#21118 Franks 10/31/2011
            SQLQ = SQLQ & " WHEN EE_NEWSMOKER IS NOT NULL THEN EE_OLDSMOKER "
            'Ticket #28794 - Opening up Marital Status for everyone
            'SQLQ = SQLQ & " WHEN EE_NEWMSTAT IS NOT NULL THEN EE_OLDMSTAT "
        End If
        'Ticket #28794 - Opening up Marital Status for everyone
        SQLQ = SQLQ & " WHEN EE_NEWMSTAT IS NOT NULL THEN EE_OLDMSTAT "
        
        SQLQ = SQLQ & " WHEN EE_NEWPOSITION IS NOT NULL THEN EE_OLDPOSITION " 'Ticket #27553 Franks 09/21/2015
        SQLQ = SQLQ & " WHEN EE_NEWREPORT1 IS NOT NULL THEN EE_OLDREPORT1 " 'Ticket #27553 Franks 09/21/2015
        SQLQ = SQLQ & " WHEN EE_SALARY IS NOT NULL THEN '' "
        SQLQ = SQLQ & " END) AS EE_OLDVALUE, "
        'NEWVALUE
        SQLQ = SQLQ & " (CASE WHEN EE_NEWDEPT IS NOT NULL THEN EE_NEWDEPT "
        SQLQ = SQLQ & " WHEN EE_NEWDIV IS NOT NULL THEN EE_NEWDIV "
        SQLQ = SQLQ & " WHEN EE_NEWSTAT IS NOT NULL THEN EE_NEWSTAT "
        SQLQ = SQLQ & " WHEN EE_NEWPT IS NOT NULL THEN EE_NEWPT "
        SQLQ = SQLQ & " WHEN EE_NEWORG IS NOT NULL THEN EE_NEWORG "
        SQLQ = SQLQ & " WHEN EE_NEWFTE IS NOT NULL THEN cast(EE_NEWFTE AS varchar(20)) "
        SQLQ = SQLQ & " WHEN EE_NEWFTEHR IS NOT NULL THEN cast(EE_NEWFTEHR AS varchar(20)) "
        SQLQ = SQLQ & " WHEN EE_NEWREGION IS NOT NULL THEN EE_NEWREGION "
        SQLQ = SQLQ & " WHEN EE_NEWSECTION IS NOT NULL THEN EE_NEWSECTION "
        SQLQ = SQLQ & " WHEN EE_NEWADMINBY IS NOT NULL THEN EE_NEWADMINBY "
        SQLQ = SQLQ & " WHEN EE_NEWLOC IS NOT NULL THEN EE_NEWLOC "
        SQLQ = SQLQ & " WHEN EE_NEWBENEGROUP IS NOT NULL THEN EE_NEWBENEGROUP "
        If glbWFC Then '#21118 Franks 10/28/2011
            SQLQ = SQLQ & " WHEN EE_NEWSMOKER IS NOT NULL THEN EE_NEWSMOKER  "
            'Ticket #28794 - Opening up Marital Status for everyone
            'SQLQ = SQLQ & " WHEN EE_NEWMSTAT IS NOT NULL THEN EE_NEWMSTAT "
        End If
        'Ticket #28794 - Opening up Marital Status for everyone
        SQLQ = SQLQ & " WHEN EE_NEWMSTAT IS NOT NULL THEN EE_NEWMSTAT "
        
        SQLQ = SQLQ & " WHEN EE_NEWPOSITION IS NOT NULL THEN EE_NEWPOSITION " 'Ticket #27553 Franks 09/21/2015
        SQLQ = SQLQ & " WHEN EE_NEWREPORT1 IS NOT NULL THEN EE_NEWREPORT1 " 'Ticket #27553 Franks 09/21/2015
        SQLQ = SQLQ & " WHEN EE_SALARY IS NOT NULL THEN '' "
        SQLQ = SQLQ & " END) AS EE_NEWVALUE "
        SQLQ = SQLQ & ",'" & glbUserID & "' AS EE_WRKEMP "
        SQLQ = SQLQ & ",'0_' + LTRIM(RTRIM(STR(Term_HREMP.ED_EMPNBR)) ) + LTRIM(RTRIM(STR(Term_HREMP.TERM_SEQ)) ) AS KEY_EMPNBR "
    Case "ORACLE"
        'SALARY
        SQLQ = SQLQ & " (CASE WHEN EE_SALARY <> 0 THEN TO_CHAR(EE_SALARY) "
        SQLQ = SQLQ & " ELSE '' END) AS EE_SALARY, "
        'PER
        SQLQ = SQLQ & " (CASE WHEN EE_SALARY <> 0 THEN EE_SALCD "
        SQLQ = SQLQ & " ELSE '' END) AS EE_SALCD, "
        
        SQLQ = SQLQ & "EE_CHGDATE,EE_LDATE,EE_LTIME,EE_LUSER,EE_DOT,Term_HREMPHIS.TERM_SEQ, "
        'TYPE
        SQLQ = SQLQ & " (CASE WHEN EE_NEWDEPT IS NOT NULL THEN '" & lStr("Department") & " ' "
        SQLQ = SQLQ & " WHEN EE_NEWDIV IS NOT NULL THEN '" & lStr("Division") & " ' "
        SQLQ = SQLQ & " WHEN EE_NEWSTAT IS NOT NULL THEN '" & lStr("Status") & " ' "
        SQLQ = SQLQ & " WHEN EE_NEWPT IS NOT NULL THEN '" & lStr("Category") & " ' "
        SQLQ = SQLQ & " WHEN EE_NEWORG IS NOT NULL THEN '" & lStr("Union") & " ' "
        SQLQ = SQLQ & " WHEN EE_NEWFTE IS NOT NULL THEN '" & lStr("FTE#") & " ' "
        SQLQ = SQLQ & " WHEN EE_NEWFTEHR IS NOT NULL THEN '" & lStr("FTE# Hours") & " ' "
        SQLQ = SQLQ & " WHEN EE_NEWREGION IS NOT NULL THEN '" & lStr("Region") & " ' "
        SQLQ = SQLQ & " WHEN EE_NEWSECTION IS NOT NULL THEN '" & lStr("Section") & " ' "
        SQLQ = SQLQ & " WHEN EE_NEWADMINBY IS NOT NULL THEN '" & lStr("Administered By") & " ' "
        SQLQ = SQLQ & " WHEN EE_NEWLOC IS NOT NULL THEN '" & lStr("Location") & " ' "
        SQLQ = SQLQ & " WHEN EE_NEWBENEGROUP IS NOT NULL THEN '" & lStr("Benefit Group") & " ' "
        'SQLQ = SQLQ & " WHEN EE_SALARY <> 0 THEN 'Salary' "
        SQLQ = SQLQ & " END) AS EE_HISTYPE, "
    
    
        'OLDVALUE
        SQLQ = SQLQ & " (CASE WHEN EE_NEWDEPT IS NOT NULL THEN EE_OLDDEPT "
        SQLQ = SQLQ & " WHEN EE_NEWDIV IS NOT NULL THEN EE_OLDDIV "
        SQLQ = SQLQ & " WHEN EE_NEWSTAT IS NOT NULL THEN EE_OLDSTAT "
        SQLQ = SQLQ & " WHEN EE_NEWPT IS NOT NULL THEN EE_OLDPT "
        SQLQ = SQLQ & " WHEN EE_NEWORG IS NOT NULL THEN EE_OLDORG "
        SQLQ = SQLQ & " WHEN EE_NEWFTE IS NOT NULL THEN TO_CHAR(EE_OLDFTE) "
        SQLQ = SQLQ & " WHEN EE_NEWFTEHR IS NOT NULL THEN TO_CHAR(EE_OLDFTEHR) "
        SQLQ = SQLQ & " WHEN EE_NEWREGION IS NOT NULL THEN EE_OLDREGION "
        SQLQ = SQLQ & " WHEN EE_NEWSECTION IS NOT NULL THEN EE_OLDSECTION "
        SQLQ = SQLQ & " WHEN EE_NEWADMINBY IS NOT NULL THEN EE_OLDADMINBY "
        SQLQ = SQLQ & " WHEN EE_NEWLOC IS NOT NULL THEN EE_OLDLOC "
        SQLQ = SQLQ & " WHEN EE_NEWBENEGROUP IS NOT NULL THEN EE_OLDBENEGROUP "
        SQLQ = SQLQ & " WHEN EE_SALARY IS NOT NULL THEN '' "
        SQLQ = SQLQ & " END) AS EE_OLDVALUE, "
        'NEWVALUE
        SQLQ = SQLQ & " (CASE WHEN EE_NEWDEPT IS NOT NULL THEN EE_NEWDEPT "
        SQLQ = SQLQ & " WHEN EE_NEWDIV IS NOT NULL THEN EE_NEWDIV "
        SQLQ = SQLQ & " WHEN EE_NEWSTAT IS NOT NULL THEN EE_NEWSTAT "
        SQLQ = SQLQ & " WHEN EE_NEWPT IS NOT NULL THEN EE_NEWPT "
        SQLQ = SQLQ & " WHEN EE_NEWORG IS NOT NULL THEN EE_NEWORG "
        SQLQ = SQLQ & " WHEN EE_NEWFTE IS NOT NULL THEN TO_CHAR(EE_NEWFTE) "
        SQLQ = SQLQ & " WHEN EE_NEWFTEHR IS NOT NULL THEN TO_CHAR(EE_NEWFTEHR) "
        SQLQ = SQLQ & " WHEN EE_NEWREGION IS NOT NULL THEN EE_NEWREGION "
        SQLQ = SQLQ & " WHEN EE_NEWSECTION IS NOT NULL THEN EE_NEWSECTION "
        SQLQ = SQLQ & " WHEN EE_NEWADMINBY IS NOT NULL THEN EE_NEWADMINBY "
        SQLQ = SQLQ & " WHEN EE_NEWLOC IS NOT NULL THEN EE_NEWLOC "
        SQLQ = SQLQ & " WHEN EE_NEWBENEGROUP IS NOT NULL THEN EE_NEWBENEGROUP "
        SQLQ = SQLQ & " WHEN EE_SALARY IS NOT NULL THEN '' "
        SQLQ = SQLQ & " END) AS EE_NEWVALUE "
        SQLQ = SQLQ & ",'" & glbUserID & "' AS EE_WRKEMP "
        SQLQ = SQLQ & ",'0_' + LTRIM(RTRIM(STR(Term_HREMP.ED_EMPNBR)) ) + LTRIM(RTRIM(STR(Term_HREMP.TERM_SEQ)) ) AS KEY_EMPNBR "
    Case Else
        SQLQ = SQLQ & " IIF( EE_SALARY <> 0 , str(EE_SALARY), '') AS SALARY,"
        SQLQ = SQLQ & " IIF( EE_SALARY <> 0 , EE_SALCD , '') AS SALCD,"
        SQLQ = SQLQ & " EE_CHGDATE,EE_LDATE,EE_LTIME,EE_LUSER,EE_DOT,Term_HREMPHIS.TERM_SEQ, "
        
        SQLQ = SQLQ & " IIF(EE_NEWDEPT IS NOT NULL , 'Department ' ,"
        SQLQ = SQLQ & " IIF(EE_NEWDIV IS NOT NULL , 'Division ' ,"
        SQLQ = SQLQ & " IIF(EE_NEWSTAT IS NOT NULL , 'Status ' ,"
        SQLQ = SQLQ & " IIF(EE_NEWPT IS NOT NULL , 'FT/PT/SE/TR/OT ' ,"
        SQLQ = SQLQ & " IIF(EE_NEWORG IS NOT NULL , 'Union ' ,"
        SQLQ = SQLQ & " IIF(EE_NEWFTE IS NOT NULL , 'FTE# ' ,"
        SQLQ = SQLQ & " IIF(EE_NEWFTEHR IS NOT NULL , 'FTE# Hours ' ,"
        SQLQ = SQLQ & " IIF(EE_NEWREGION IS NOT NULL , 'Region ' ,"
        SQLQ = SQLQ & " IIF(EE_NEWSECTION IS NOT NULL , 'Section ',"
        SQLQ = SQLQ & " IIF(EE_NEWADMINBY IS NOT NULL , 'Administered By ' ,"
        SQLQ = SQLQ & " IIF(EE_NEWLOC IS NOT NULL , 'Location '  ,"
        SQLQ = SQLQ & " IIF(EE_NEWBENEGROUP IS NOT NULL , 'Benefit Group', '')"
        SQLQ = SQLQ & " )))))))))))"
        SQLQ = SQLQ & " AS EE_HISTYPE,"
        
        SQLQ = SQLQ & " IIF(EE_NEWDEPT IS NOT NULL , EE_OLDDEPT ,"
        SQLQ = SQLQ & " IIF(EE_NEWDIV IS NOT NULL , EE_OLDDIV ,"
        SQLQ = SQLQ & " IIF(EE_NEWSTAT IS NOT NULL , EE_OLDSTAT ,"
        SQLQ = SQLQ & " IIF(EE_NEWPT IS NOT NULL , EE_OLDPT ,"
        SQLQ = SQLQ & " IIF(EE_NEWORG IS NOT NULL , EE_OLDORG,"
        SQLQ = SQLQ & " IIF(EE_NEWFTE IS NOT NULL , str(EE_OLDFTE ),"
        SQLQ = SQLQ & " IIF(EE_NEWFTEHR IS NOT NULL , str(EE_OLDFTEHR),"
        SQLQ = SQLQ & " IIF(EE_NEWREGION IS NOT NULL , EE_OLDREGION,"
        SQLQ = SQLQ & " IIF(EE_NEWSECTION IS NOT NULL , EE_OLDSECTION,"
        SQLQ = SQLQ & " IIF(EE_NEWADMINBY IS NOT NULL , EE_OLDADMINBY ,"
        SQLQ = SQLQ & " IIF( EE_NEWLOC IS NOT NULL , EE_OLDLOC ,"
        SQLQ = SQLQ & " IIF( EE_NEWBENEGROUP IS NOT NULL , EE_OLDBENEGROUP ,"
        SQLQ = SQLQ & " IIF( EE_SALARY IS NOT NULL , ''  ,'')"
        SQLQ = SQLQ & " ))))))))))))"
        SQLQ = SQLQ & "  AS EE_OLDVALUE,"
        
        SQLQ = SQLQ & " IIF( EE_NEWDEPT IS NOT NULL , EE_NEWDEPT  ,"
        SQLQ = SQLQ & " IIF( EE_NEWDIV IS NOT NULL , EE_NEWDIV  ,"
        SQLQ = SQLQ & " IIF( EE_NEWSTAT IS NOT NULL , EE_NEWSTAT  ,"
        SQLQ = SQLQ & " IIF( EE_NEWPT IS NOT NULL , EE_NEWPT  ,"
        SQLQ = SQLQ & " IIF( EE_NEWORG IS NOT NULL , EE_NEWORG  ,"
        SQLQ = SQLQ & " IIF( EE_NEWFTE IS NOT NULL , str(EE_NEWFTE),"
        SQLQ = SQLQ & " IIF( EE_NEWFTEHR IS NOT NULL , str(EE_NEWFTEHR),"
        SQLQ = SQLQ & " IIF( EE_NEWREGION IS NOT NULL , EE_NEWREGION  ,"
        SQLQ = SQLQ & " IIF( EE_NEWSECTION IS NOT NULL , EE_NEWSECTION,"
        SQLQ = SQLQ & " IIF(EE_NEWADMINBY IS NOT NULL , EE_NEWADMINBY  ,"
        SQLQ = SQLQ & " IIF( EE_NEWLOC IS NOT NULL , EE_NEWLOC  ,"
        SQLQ = SQLQ & " IIF( EE_NEWBENEGROUP IS NOT NULL , EE_NEWBENEGROUP  ,"
        SQLQ = SQLQ & " IIF( EE_SALARY IS NOT NULL , '' ,'' )"
        SQLQ = SQLQ & " ))))))))))))"
        SQLQ = SQLQ & " AS EE_NEWVALUE"
        SQLQ = SQLQ & ",'" & glbUserID & "' AS EE_WRKEMP "
    End Select
    
    'Ticket #27553 Franks 09/28/2015 - begin
    'SQLQ = SQLQ & " FROM Term_HREMPHIS "
    SQLQ = SQLQ & " FROM Term_HREMPHIS LEFT JOIN Term_HREMP ON Term_HREMPHIS.TERM_SEQ = Term_HREMP.TERM_SEQ "
    SQLQ = SQLQ & " WHERE (1=1) "
    'Hemu - 09/29/2003 Begin - Taking too long to generate the report, esp. when Emp#
    '                          entered
    If Len(elpEEID.Text) > 0 Then
        SQLQ = SQLQ & "AND Term_HREMPHIS.EE_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
    End If
    'Hemu
    If Len(clpDiv.Text) > 0 Then
        SQLQ = SQLQ & "AND ED_DIV IN ('" & Replace(clpDiv.Text, ",", "','") & "') "
    End If
    If Len(clpDept.Text) > 0 Then
        SQLQ = SQLQ & "AND ED_DEPTNO IN ('" & Replace(clpDept.Text, ",", "','") & "') "
    End If
    If Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & "AND ED_LOC IN ('" & Replace(clpCode(0).Text, ",", "','") & "') "
    End If
    If Len(clpCode(1).Text) > 0 Then
        SQLQ = SQLQ & "AND ED_ORG IN ('" & Replace(clpCode(1).Text, ",", "','") & "') "
    End If
    If Len(clpCode(2).Text) > 0 Then
        SQLQ = SQLQ & "AND ED_EMP IN ('" & Replace(clpCode(2).Text, ",", "','") & "') "
    End If
    If Len(clpPT.Text) > 0 Then
        SQLQ = SQLQ & "AND ED_PT IN ('" & Replace(clpPT.Text, ",", "','") & "') "
    End If
    If Len(clpCode(3).Text) > 0 Then
        SQLQ = SQLQ & "AND ED_REGION IN ('" & Replace(clpCode(3).Text, ",", "','") & "') "
    End If
    If Len(clpCode(4).Text) > 0 Then
        SQLQ = SQLQ & "AND ED_ADMINBY IN ('" & Replace(clpCode(4).Text, ",", "','") & "') "
    End If
    If Len(clpCode(5).Text) > 0 Then
        SQLQ = SQLQ & "AND ED_SECTION IN ('" & Replace(clpCode(5).Text, ",", "','") & "') "
    End If
    If IsDate(dlpDateRange(0).Text) Then
        SQLQ = SQLQ & "AND Term_HREMPHIS.EE_CHGDATE >= " & Date_SQL(dlpDateRange(0).Text) & " "
    End If
    If IsDate(dlpDateRange(1).Text) Then
        SQLQ = SQLQ & "AND Term_HREMPHIS.EE_CHGDATE <= " & Date_SQL(dlpDateRange(1).Text) & " "
    End If
    'Ticket #27553 Franks 09/28/2015 - end
    
    'Ticket #29660 - Contract Employees Enhancement
    If glbWFC Then
        If chkExclCONP.Visible And chkExclRET.Visible = True Then
            If chkExclCONP Then
                SQLQ = SQLQ & "AND ED_EMP <> 'CONP'"
            End If
            If chkExclRET Then
                SQLQ = SQLQ & "AND ED_EMP <> 'RET'"
            End If
        End If
    End If
    
    gdbAdoIhr001X.BeginTrans
    gdbAdoIhr001X.Execute SQLQ
    gdbAdoIhr001X.CommitTrans

End If '------------------- end of terminated


MDIMain.panHelp(0).FloodPercent = 85
'If chkReport(0) Or optSum(2) Then
    MDIMain.panHelp(0).FloodPercent = 90
    Call Pause(2)
    MDIMain.panHelp(0).FloodPercent = 100
    Call Pause(2)

'End If

'Ticket #16146
SQLQ = "DELETE FROM HREMPHIS_WRK WHERE EE_WRKEMP='" & glbUserID & "' AND (EE_CHGDATE IS NULL)"
gdbAdoIhr001X.Execute SQLQ
Call Pause(1)

MDIMain.panHelp(0).FloodPercent = 100

'Ticket #27553 Franks 09/28/2015
'o   For the "Use Descriptions" check box, Division and Department will be excluded for WFC
If chkUseDesc.Value Then
    Call UseDescPro
End If

If glbWFC And chkUseExcel.Value Then
    If Not comGroup(0).Text = "(none)" Then
        'add group and desc to Excel report
        Call WFCExcRptGroup
    End If
End If

Screen.MousePointer = DEFAULT
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""

End Sub

Private Sub WFCExcRptGroup() 'Ticket #27553 Franks 09/29/2015
Dim rsWrk As New ADODB.Recordset
Dim SQLQ
Dim xStr As String
Dim I As Long
Dim xTot As Long
Dim xCode, xDesc

    SQLQ = "SELECT HREMPHIS_WRK.*,qry_HREMP.* FROM HREMPHIS_WRK "
    SQLQ = SQLQ & "LEFT JOIN qry_HREMP ON HREMPHIS_WRK.KEY_EMPNBR = qry_HREMP.KEY_EMPNBR "
    SQLQ = SQLQ & "WHERE EE_WRKEMP='" & glbUserID & "' "
    SQLQ = SQLQ & "ORDER BY HREMPHIS_WRK.KEY_EMPNBR " 'ED_SURNAME,ED_FNAME,EE_CHGDATE DESC "

    If rsWrk.State <> 0 Then rsWrk.Close
    rsWrk.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsWrk.EOF Then
        'MsgBox "There is no any record in this Selection Criteria"
        rsWrk.Close
        Exit Sub
    End If
    
    
    If rsWrk.State <> 0 Then rsWrk.Close
    rsWrk.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsWrk.EOF Then
        I = 0
        xTot = rsWrk.RecordCount
        Do While Not rsWrk.EOF
            MDIMain.panHelp(0).FloodPercent = (I / xTot) * 100
            I = I + 1
            DoEvents
            
            If comGroup(0).Text = lStr("Division") Then
                If Not IsNull(rsWrk("ED_DIV")) Then
                    xCode = rsWrk("ED_DIV")
                    xDesc = getDivDescPub(xCode)
                End If
            End If
            If comGroup(0).Text = lStr("Department") Then
                If Not IsNull(rsWrk("ED_DEPTNO")) Then
                    xCode = rsWrk("ED_DEPTNO")
                    xDesc = getDeptDescPub(xCode)
                End If
            End If
            If comGroup(0).Text = lStr("Location") Then
                If Not IsNull(rsWrk("ED_LOC")) Then
                    xCode = rsWrk("ED_LOC")
                    xDesc = GetTABLCodePub("EDLC", xCode)
                End If
            End If
            If comGroup(0).Text = lStr("Union") Then
                If Not IsNull(rsWrk("ED_ORG")) Then
                    xCode = rsWrk("ED_ORG")
                    xDesc = GetTABLCodePub("EDOR", xCode)
                End If
            End If
            If comGroup(0).Text = lStr("Region") Then
                If Not IsNull(rsWrk("ED_REGION")) Then
                    xCode = rsWrk("ED_REGION")
                    xDesc = GetTABLCodePub("EDRG", xCode)
                End If
            End If
            If comGroup(0).Text = lStr("Section") Then
                If Not IsNull(rsWrk("ED_SECTION")) Then
                    xCode = rsWrk("ED_SECTION")
                    xDesc = GetTABLCodePub("EDSE", xCode)
                End If
            End If
            
   
            If Len(xDesc) > 0 Then
                rsWrk("EE_GROUP") = Left(xCode, 25)
                rsWrk("EE_GROUP_DESC") = Left(xDesc, 50)
                rsWrk.Update
            End If
            rsWrk.MoveNext
        Loop
    End If
    rsWrk.Close
    MDIMain.panHelp(0).FloodPercent = 100
    
End Sub

Private Sub UseDescPro() 'Ticket #27553 Franks 09/28/2015
Dim rsWrk As New ADODB.Recordset
Dim SQLQ
Dim xStr As String
Dim I As Long
Dim xTot As Long
    
    xStr = "'" & lStr("Administered By") & "',"
    xStr = xStr & "'" & lStr("Benefit Group") & "',"
    If glbWFC Then
        'Division and Department will be excluded for WFC
    Else
        xStr = xStr & "'" & lStr("Department") & "',"
        xStr = xStr & "'" & lStr("Division") & "',"
    End If
    xStr = xStr & "'" & lStr("Category") & "',"
    xStr = xStr & "'" & lStr("Location") & "',"
    xStr = xStr & "'" & lStr("Union") & "',"
    xStr = xStr & "'" & lStr("Region") & "',"
    xStr = xStr & "'" & lStr("Section") & "',"
    xStr = xStr & "'" & lStr("Status") & "',"
    xStr = xStr & "'" & "Position" & "'"
    SQLQ = "SELECT * FROM HREMPHIS_WRK WHERE EE_WRKEMP='" & glbUserID & "' "
    SQLQ = SQLQ & "AND EE_HISTYPE IN (" & xStr & ") "
    
    If rsWrk.State <> 0 Then rsWrk.Close
    rsWrk.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsWrk.EOF Then
        I = 0
        xTot = rsWrk.RecordCount
        Do While Not rsWrk.EOF
            MDIMain.panHelp(0).FloodPercent = (I / xTot) * 100
            I = I + 1
            DoEvents
            If Trim(rsWrk("EE_HISTYPE")) = lStr("Administered By") Then
                If Not IsNull(rsWrk("EE_OLDVALUE")) Then
                    rsWrk("EE_OLDVALUE") = Left(GetTABLCodePub("EDAB", rsWrk("EE_OLDVALUE")), 50)
                End If
                If Not IsNull(rsWrk("EE_NEWVALUE")) Then
                    rsWrk("EE_NEWVALUE") = Left(GetTABLCodePub("EDAB", rsWrk("EE_NEWVALUE")), 50)
                End If
            End If
            If Trim(rsWrk("EE_HISTYPE")) = lStr("Benefit Group") Then
                If Not IsNull(rsWrk("EE_OLDVALUE")) Then
                    rsWrk("EE_OLDVALUE") = Left(GetTABLCodePub("BGMF", rsWrk("EE_OLDVALUE")), 50)
                End If
                If Not IsNull(rsWrk("EE_NEWVALUE")) Then
                    rsWrk("EE_NEWVALUE") = Left(GetTABLCodePub("BGMF", rsWrk("EE_NEWVALUE")), 50)
                End If
            End If
            If Trim(rsWrk("EE_HISTYPE")) = lStr("Department") Then
                If Not IsNull(rsWrk("EE_OLDVALUE")) Then
                    rsWrk("EE_OLDVALUE") = Left(getDeptDescPub(rsWrk("EE_OLDVALUE")), 50)
                End If
                If Not IsNull(rsWrk("EE_NEWVALUE")) Then
                    rsWrk("EE_NEWVALUE") = Left(getDeptDescPub(rsWrk("EE_NEWVALUE")), 50)
                End If
            End If
            If Trim(rsWrk("EE_HISTYPE")) = lStr("Division") Then
                If Not IsNull(rsWrk("EE_OLDVALUE")) Then
                    rsWrk("EE_OLDVALUE") = Left(getDivDescPub(rsWrk("EE_OLDVALUE")), 50)
                End If
                If Not IsNull(rsWrk("EE_NEWVALUE")) Then
                    rsWrk("EE_NEWVALUE") = Left(getDivDescPub(rsWrk("EE_NEWVALUE")), 50)
                End If
            End If
            If Trim(rsWrk("EE_HISTYPE")) = lStr("Category") Then
                If Not IsNull(rsWrk("EE_OLDVALUE")) Then
                    rsWrk("EE_OLDVALUE") = Left(GetTABLCodePub("EDPT", rsWrk("EE_OLDVALUE")), 50)
                End If
                If Not IsNull(rsWrk("EE_NEWVALUE")) Then
                    rsWrk("EE_NEWVALUE") = Left(GetTABLCodePub("EDPT", rsWrk("EE_NEWVALUE")), 50)
                End If
            End If
            If Trim(rsWrk("EE_HISTYPE")) = lStr("Location") Then
                If Not IsNull(rsWrk("EE_OLDVALUE")) Then
                    rsWrk("EE_OLDVALUE") = Left(GetTABLCodePub("EDLC", rsWrk("EE_OLDVALUE")), 50)
                End If
                If Not IsNull(rsWrk("EE_NEWVALUE")) Then
                    rsWrk("EE_NEWVALUE") = Left(GetTABLCodePub("EDLC", rsWrk("EE_NEWVALUE")), 50)
                End If
            End If
            If Trim(rsWrk("EE_HISTYPE")) = lStr("Union") Then
                If Not IsNull(rsWrk("EE_OLDVALUE")) Then
                    rsWrk("EE_OLDVALUE") = Left(GetTABLCodePub("EDOR", rsWrk("EE_OLDVALUE")), 50)
                End If
                If Not IsNull(rsWrk("EE_NEWVALUE")) Then
                    rsWrk("EE_NEWVALUE") = Left(GetTABLCodePub("EDOR", rsWrk("EE_NEWVALUE")), 50)
                End If
            End If
            If Trim(rsWrk("EE_HISTYPE")) = lStr("Region") Then
                If Not IsNull(rsWrk("EE_OLDVALUE")) Then
                    rsWrk("EE_OLDVALUE") = Left(GetTABLCodePub("EDRG", rsWrk("EE_OLDVALUE")), 50)
                End If
                If Not IsNull(rsWrk("EE_NEWVALUE")) Then
                    rsWrk("EE_NEWVALUE") = Left(GetTABLCodePub("EDRG", rsWrk("EE_NEWVALUE")), 50)
                End If
            End If
            If Trim(rsWrk("EE_HISTYPE")) = lStr("Section") Then
                If Not IsNull(rsWrk("EE_OLDVALUE")) Then
                    rsWrk("EE_OLDVALUE") = Left(GetTABLCodePub("EDSE", rsWrk("EE_OLDVALUE")), 50)
                End If
                If Not IsNull(rsWrk("EE_NEWVALUE")) Then
                    rsWrk("EE_NEWVALUE") = Left(GetTABLCodePub("EDSE", rsWrk("EE_NEWVALUE")), 50)
                End If
            End If
            If Trim(rsWrk("EE_HISTYPE")) = lStr("Status") Then
                If Not IsNull(rsWrk("EE_OLDVALUE")) Then
                    rsWrk("EE_OLDVALUE") = Left(GetTABLCodePub("EDEM", rsWrk("EE_OLDVALUE")), 50)
                End If
                If Not IsNull(rsWrk("EE_NEWVALUE")) Then
                    rsWrk("EE_NEWVALUE") = Left(GetTABLCodePub("EDEM", rsWrk("EE_NEWVALUE")), 50)
                End If
            End If
            If Trim(rsWrk("EE_HISTYPE")) = "Position" Then
                If Not IsNull(rsWrk("EE_OLDVALUE")) Then
                    rsWrk("EE_OLDVALUE") = Left(getPosDesc(rsWrk("EE_OLDVALUE")), 50)
                End If
                If Not IsNull(rsWrk("EE_NEWVALUE")) Then
                    rsWrk("EE_NEWVALUE") = Left(getPosDesc(rsWrk("EE_NEWVALUE")), 50)
                End If
            End If
            'Status rsWRK("EE_OLDVALUE") = Left(GetTABLCodePub("JBGC", rsEmp("JB_GRPCD")), 50)
            rsWrk.Update
            rsWrk.MoveNext
        Loop
    End If
    rsWrk.Close
    
End Sub

Private Sub WFCExcelRpt() 'Ticket #27553 Franks 09/29/2015
Dim exApp As Object, exBook As Object, exSheet As Object
Dim rsWrk As New ADODB.Recordset
Dim xlsFileTmp As String, xlsFileMat As String
Dim SQLQ As String
Dim K As Long
Dim xRow, xRows
Dim xNewEmp As Boolean
Dim xCurName, xNextName
Dim xCurGroup, xNextGroup

    SQLQ = "SELECT HREMPHIS_WRK.*,ED_SURNAME,ED_FNAME FROM HREMPHIS_WRK "
    SQLQ = SQLQ & "LEFT JOIN qry_HREMP ON HREMPHIS_WRK.KEY_EMPNBR = qry_HREMP.KEY_EMPNBR "
    SQLQ = SQLQ & "WHERE EE_WRKEMP='" & glbUserID & "' "
    'Ticket #29660 - filter on Change Type is missing so adding it.
    If Len(glbCode) > 0 Then SQLQ = SQLQ & " AND EE_HISTYPE IN ('" & Replace(glbCode, ",", "','") & "') "
    SQLQ = SQLQ & "ORDER BY EE_GROUP_DESC, ED_SURNAME,ED_FNAME,EE_CHGDATE DESC "

    If rsWrk.State <> 0 Then rsWrk.Close
    rsWrk.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsWrk.EOF Then
        'MsgBox "There is no any record in this Selection Criteria"
        rsWrk.Close
        Exit Sub
    End If
    
    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "WFCEmpHisTmp.xls"
    
    xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "WFCEmpHis(" & glbUserID & ").xls"
    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If
    
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat

    FileCopy xlsFileTmp, xlsFileMat
    
    Set exApp = CreateObject("Excel.Application") 'New Excel.Application
    Set exBook = exApp.Workbooks.Open(xlsFileMat)
    Set exSheet = exBook.Worksheets(1)
    
    exSheet.Cells(2, 1).Font.Bold = True
    exSheet.Cells(2, 1) = "Date: " & Date
    exSheet.Cells(3, 1).Font.Bold = True
    exSheet.Cells(3, 1) = "Time:" & Time$
    
    xRows = rsWrk.RecordCount
    xRow = 0
        
    K = 5
    xNewEmp = True
    xCurGroup = "**"
    xCurName = "***"
    Do While Not rsWrk.EOF
        MDIMain.panHelp(0).FloodPercent = (xRow / xRows) * 100
        DoEvents
        xRow = xRow + 1
        
        If Not comGroup(0).Text = "(none)" Then
            If IsNull(rsWrk("EE_GROUP_DESC")) Then xNextGroup = "" Else xNextGroup = rsWrk("EE_GROUP_DESC")
            If Not (xCurGroup = xNextGroup) Then
                K = K + 1
                xCurGroup = xNextGroup
                exSheet.Cells(K, 1).Font.Bold = True
                exSheet.Cells(K, 1) = xNextGroup
                K = K + 1
            End If
        End If
        
        xNextName = rsWrk("ED_SURNAME") & ", " & rsWrk("ED_FNAME")
        If Not (xCurName = xNextName) Then
            K = K + 1
            exSheet.Cells(K, 1).Font.Bold = True
            exSheet.Cells(K, 1) = "Employee Number and Name: " & rsWrk("EE_EMPNBR") & " " & xNextName
            xCurName = xNextName
            K = K + 1
            exSheet.Cells(K, 1) = "Change Date": exSheet.Cells(K, 1).Font.Bold = True
            exSheet.Cells(K, 2) = "Change Type": exSheet.Cells(K, 2).Font.Bold = True
            exSheet.Cells(K, 3) = "Old Value": exSheet.Cells(K, 3).Font.Bold = True
            exSheet.Cells(K, 4) = "New Value": exSheet.Cells(K, 4).Font.Bold = True
            exSheet.Cells(K, 5) = "Change By User": exSheet.Cells(K, 5).Font.Bold = True
            K = K + 1
        End If
        exSheet.Cells(K, 1) = rsWrk("EE_CHGDATE")
        exSheet.Cells(K, 2) = rsWrk("EE_HISTYPE")
        exSheet.Cells(K, 3) = rsWrk("EE_OLDVALUE")
        exSheet.Cells(K, 4) = rsWrk("EE_NEWVALUE")
        exSheet.Cells(K, 5) = rsWrk("EE_LUSER")
        K = K + 1
        
        rsWrk.MoveNext
    Loop
    rsWrk.Close
    
    If Not exBook Is Nothing Then
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
    End If

    Call Pause(1)
    
    If Not LanchXlsW98(xlsFileMat) Then
        Shell "cmd /c " & GetShortName(xlsFileMat)
    End If
    
End Sub
