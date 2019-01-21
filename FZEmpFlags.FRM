VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmREmpFlags 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Employee Flags Report"
   ClientHeight    =   8055
   ClientLeft      =   375
   ClientTop       =   915
   ClientWidth     =   10020
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
   ScaleHeight     =   8055
   ScaleWidth      =   10020
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtEmpFlag 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Height          =   285
      Index           =   0
      Left            =   4680
      TabIndex        =   39
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox cboEmpFlag 
      Height          =   315
      Index           =   0
      Left            =   1875
      TabIndex        =   19
      Tag             =   "00-Enter Value for this Flag"
      Top             =   6120
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.ListBox lstEmpFlagsList 
      Height          =   1620
      ItemData        =   "FZEmpFlags.frx":0000
      Left            =   1875
      List            =   "FZEmpFlags.frx":0002
      MultiSelect     =   1  'Simple
      TabIndex        =   11
      Top             =   4080
      Width           =   2655
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1875
      MaxLength       =   4
      TabIndex        =   10
      Tag             =   "00-Employee Position Shift"
      Top             =   3740
      Visible         =   0   'False
      Width           =   450
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   1560
      TabIndex        =   8
      Tag             =   "00-Enter Administered By Code"
      Top             =   3070
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
      TabIndex        =   7
      Tag             =   "00-Enter Region Code"
      Top             =   2735
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Tag             =   "EDPT-Category"
      Top             =   2065
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
      Index           =   2
      Left            =   1560
      TabIndex        =   4
      Tag             =   "00-Enter Status Code"
      Top             =   1730
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
      Index           =   1
      Left            =   1560
      TabIndex        =   3
      Tag             =   "00-Enter Union Code"
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   2
      Tag             =   "00-Enter Location Code"
      Top             =   1060
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
      Top             =   725
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
      Left            =   1890
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Tag             =   "Second level of grouping records"
      Top             =   7200
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
      Left            =   1890
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Tag             =   "First Level of grouping records"
      Top             =   6885
      Width           =   2325
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Tag             =   "10-Enter Employee Number"
      Top             =   2400
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   7200
      Top             =   7440
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
      Index           =   5
      Left            =   1560
      TabIndex        =   9
      Tag             =   "00-Enter Section Code"
      Top             =   3405
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin Threed.SSCheck chkShowEmp 
      Height          =   225
      Left            =   4800
      TabIndex        =   12
      Tag             =   "If X-Show All Employee Flags"
      Top             =   4080
      Width           =   2835
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   " Show All Employee Flags"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   27.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCheck chkShowFlagFollowUp 
      Height          =   225
      Left            =   4800
      TabIndex        =   15
      Tag             =   "If X-Show All Employee Flags"
      Top             =   5070
      Visible         =   0   'False
      Width           =   3555
      _Version        =   65536
      _ExtentX        =   6271
      _ExtentY        =   397
      _StockProps     =   78
      Caption         =   " Only Show Employee Flags with Follow Ups"
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
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   3480
      TabIndex        =   18
      Tag             =   "40-Date upto and including this date forward"
      Top             =   5760
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
      TabIndex        =   17
      Tag             =   "40-Date from and including this date forward"
      Top             =   5760
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin Threed.SSCheck chkHideBlankFlags 
      Height          =   225
      Left            =   4800
      TabIndex        =   13
      Tag             =   "If X-Hide Employee Flags with no values"
      Top             =   4410
      Width           =   3555
      _Version        =   65536
      _ExtentX        =   6271
      _ExtentY        =   397
      _StockProps     =   78
      Caption         =   " Hide Employee Flags with no values"
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
   Begin Threed.SSCheck chkNoFlag 
      Height          =   225
      Left            =   4800
      TabIndex        =   16
      Tag             =   "If X-Show Employees who do not have Flag completed"
      ToolTipText     =   "Show Employees who do not have selected Flag completed"
      Top             =   5400
      Width           =   3555
      _Version        =   65536
      _ExtentX        =   6271
      _ExtentY        =   397
      _StockProps     =   78
      Caption         =   " Only Show Employees with No Flag values"
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
   Begin Threed.SSCheck chkHideEmpNoFlags 
      Height          =   225
      Left            =   4800
      TabIndex        =   14
      Tag             =   "If X-Hide Employees with no Employee Flags values"
      Top             =   4740
      Width           =   3675
      _Version        =   65536
      _ExtentX        =   6482
      _ExtentY        =   397
      _StockProps     =   78
      Caption         =   " Hide Employees with no Employee Flags values"
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
   Begin VB.Label lblVal 
      BackStyle       =   0  'Transparent
      Caption         =   "Value"
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
      Index           =   0
      Left            =   120
      TabIndex        =   40
      Top             =   6150
      Visible         =   0   'False
      Width           =   1335
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
      Top             =   5805
      Width           =   1095
   End
   Begin VB.Label lblEmpFlagList 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Flags"
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
      Top             =   4080
      Width           =   1110
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
      TabIndex        =   36
      Top             =   3740
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
      TabIndex        =   35
      Top             =   2065
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
      TabIndex        =   34
      Top             =   3405
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
      TabIndex        =   33
      Top             =   2735
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
      TabIndex        =   32
      Top             =   3070
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
      TabIndex        =   31
      Top             =   1060
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
      Left            =   0
      TabIndex        =   30
      Top             =   6675
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
      Left            =   120
      TabIndex        =   29
      Top             =   7260
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
      TabIndex        =   28
      Top             =   6945
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
      TabIndex        =   27
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
      TabIndex        =   26
      Top             =   2400
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
      Top             =   1730
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
      Top             =   1395
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
      Top             =   725
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
      Top             =   390
      Width           =   555
   End
End
Attribute VB_Name = "frmREmpFlags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkHideBlankFlags_Click(Value As Integer)
    If chkHideBlankFlags Or chkHideEmpNoFlags Then
        chkNoFlag.Value = False
        chkNoFlag.Visible = False
    Else
        chkNoFlag.Visible = True
    End If
End Sub

Private Sub chkHideEmpNoFlags_Click(Value As Integer)
    If chkHideEmpNoFlags Or chkHideBlankFlags Then
        chkNoFlag.Value = False
        chkNoFlag.Visible = False
    Else
        chkNoFlag.Visible = True
    End If
End Sub

Private Sub chkNoFlag_Click(Value As Integer)
    If chkNoFlag Then
        chkHideBlankFlags.Value = False
        chkHideBlankFlags.Visible = False
        chkHideEmpNoFlags.Value = False
        chkHideEmpNoFlags.Visible = False
        lblFromTo.Visible = False
        dlpDateRange(0).Visible = False
        dlpDateRange(1).Visible = False
    Else
        chkHideBlankFlags.Visible = True
        chkHideEmpNoFlags.Visible = True
        lblFromTo.Visible = True
        dlpDateRange(0).Visible = True
        dlpDateRange(1).Visible = True
    End If
End Sub

Private Sub chkShowEmp_Click(Value As Integer)
    Dim Y As Integer
    
    If chkShowEmp Then
        For Y = 0 To lstEmpFlagsList.ListCount - 1
            lstEmpFlagsList.selected(Y) = True
        Next Y
        lstEmpFlagsList.ListIndex = 0
    Else
        For Y = 0 To lstEmpFlagsList.ListCount - 1
            lstEmpFlagsList.selected(Y) = False
        Next Y
        lstEmpFlagsList.Refresh
        lstEmpFlagsList.ListIndex = 1
        lstEmpFlagsList.ListIndex = 0
    End If

End Sub

Private Sub chkShowEmp_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr

If CriCheck() Then

    If Not PrtForm("Employee Flags Report Criteria", Me) Then Exit Sub
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
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

    Screen.MousePointer = HOURGLASS
    x% = Cri_SetAll()
    
    'Ticket #24737 - Samuel, Son & Co., Limited
    If glbCompSerial = "S/N - 2382W" Then
        If chkShowFlagFollowUp.Value = True Then
            GoTo SkipTo
        End If
    End If
    
    Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    
SkipTo:
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
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "ENTITLEMENTS", "VIEW")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

Private Sub comEmpType_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkShowFlagFollowUp_Click(Value As Integer)
    If chkShowFlagFollowUp.Value = True Then
        lblGrp(0).Enabled = False
        comGroup(0).Enabled = False
    Else
        lblGrp(0).Enabled = True
        comGroup(0).Enabled = True
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
    comGroup(0).ListIndex = 0
    comGroup(1).AddItem "Employee Name"
    comGroup(1).AddItem "Employee Number"
    comGroup(1).ListIndex = 0

End Sub

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

Private Sub Cri_EE()
Dim EECri As String

If Len(getEmpnbr(elpEEID.Text)) > 0 Then
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
    Dim Y As Integer
    
    Cri_SetAll = False
    
    On Error GoTo modSetCriteria_Err
    
    Screen.MousePointer = HOURGLASS
    
    glbiOneWhere = False
    glbstrSelCri = ""
    
    Call glbCri_DeptUN(clpDept.Text)
    
    Call Cri_Div
    For x% = 0 To 5
        Call Cri_Code(x%)
    Next x%
    Call Cri_PT
    Call Cri_Shift
    Call Cri_EE
    
    'Populate the temporary table based on the Follow Up Code security
    Call Populate_EmployeeFlag_WorkingTable
    
    'report name
    'Ticket #24737 - Samuel, Son & Co., Limited
    If glbCompSerial = "S/N - 2382W" Then
        If chkShowFlagFollowUp.Value = True Then
            'Excel report with Employee Flags with Follow Up onlys
            Call Generate_ExcelReport
            
            GoTo SkipTo
        End If
    End If
    
    'Ticket #28040 - No Flag - List employees with no selected flag(s) completed
    If chkNoFlag Then
        If comGroup(0) <> "(none)" Then
            strRName$ = glbIHRREPORTS & "rzEmpNoFlags.rpt"
        Else
            strRName$ = glbIHRREPORTS & "rzEmpNoFlags1.rpt"
        End If
    Else
        If comGroup(0) <> "(none)" Then
            strRName$ = glbIHRREPORTS & "rzEmpFlags.rpt"
        Else
            strRName$ = glbIHRREPORTS & "rzEmpFlags1.rpt"
        End If
    End If
    
    Me.vbxCrystal.ReportFileName = strRName$
  
    'set to sorting/grouping criteria
    x% = Cri_Sorts()   ' returns number of sections formated
    
    
    'Set the Show Flags to show Employee Flags selected.
    If chkShowEmp.Value = False Then
        'Set Show Flags based on the list item checked.
        For Y = 0 To lstEmpFlagsList.ListCount - 1
            If lstEmpFlagsList.selected(Y) Then
                Me.vbxCrystal.Formulas(10 & Y + 1) = "showFlag" & Y + 1 & "='1'"
                'Me.vbxCrystal.Formulas(10 & y + 1) = "lblEmpFlag" & y + 1 & "='Employee Flag " & y + 1 & "'"
            Else
                Me.vbxCrystal.Formulas(10 & Y + 1) = "showFlag" & Y + 1 & "='0'"
                'Me.vbxCrystal.Formulas(10 & y + 1) = "lblEmpFlag" & y + 1 & "='Employee Flag " & y + 1 & "'"
            End If
        Next Y
    End If
    
    'In order to hide the Employee Flags that do not meet the From / To Date Range criteria, I am turning the
    'Hide Employee Flags with no values option ON internally. This is because when populating the work table I have
    'put a condition where if an employee's Flag do not meet the date range criteria, then do not populate the work
    'table fields with the values hence it will be considered as "no values" hence this option of
    '"hide flags with no values" will hide those flags.
     
    'Release 8.1
    'Set the Hide Flags to hide blank Employee Flags selected.
    If chkHideBlankFlags Or (IsDate(dlpDateRange(0).Text) Or IsDate(dlpDateRange(1).Text)) Then
        'Set Hide Flags based on the list item checked.
        For Y = 0 To lstEmpFlagsList.ListCount - 1
            If lstEmpFlagsList.selected(Y) Then
                Me.vbxCrystal.Formulas(40 & Y + 1) = "showBlankFlag" & Y + 1 & "='0'"
                'Me.vbxCrystal.Formulas(10 & y + 1) = "lblEmpFlag" & y + 1 & "='Employee Flag " & y + 1 & "'"
            Else
                Me.vbxCrystal.Formulas(40 & Y + 1) = "showBlankFlag" & Y + 1 & "='1'"
                'Me.vbxCrystal.Formulas(10 & y + 1) = "lblEmpFlag" & y + 1 & "='Employee Flag " & y + 1 & "'"
            End If
        Next Y
    End If
    
    'Follow Up security
'    glbstrSelCri = glbstrSelCri & " AND {HR_SECURE_FOLLOW_UP_1.USERID} ='" & glbUserID & "' AND {HR_SECURE_FOLLOW_UP_1.ACCESSABLE} = True"
'    glbstrSelCri = glbstrSelCri & " AND {HR_SECURE_FOLLOW_UP_2.USERID} ='" & glbUserID & "' AND {HR_SECURE_FOLLOW_UP_2.ACCESSABLE} = True"
'    glbstrSelCri = glbstrSelCri & " AND {HR_SECURE_FOLLOW_UP_3.USERID} ='" & glbUserID & "' AND {HR_SECURE_FOLLOW_UP_3.ACCESSABLE} = True"
'    glbstrSelCri = glbstrSelCri & " AND {HR_SECURE_FOLLOW_UP_4.USERID} ='" & glbUserID & "' AND {HR_SECURE_FOLLOW_UP_4.ACCESSABLE} = True"
'    glbstrSelCri = glbstrSelCri & " AND {HR_SECURE_FOLLOW_UP_5.USERID} ='" & glbUserID & "' AND {HR_SECURE_FOLLOW_UP_5.ACCESSABLE} = True"
'    glbstrSelCri = glbstrSelCri & " AND {HR_SECURE_FOLLOW_UP_6.USERID} ='" & glbUserID & "' AND {HR_SECURE_FOLLOW_UP_6.ACCESSABLE} = True"
'    glbstrSelCri = glbstrSelCri & " AND {HR_SECURE_FOLLOW_UP_7.USERID} ='" & glbUserID & "' AND {HR_SECURE_FOLLOW_UP_7.ACCESSABLE} = True"
'    glbstrSelCri = glbstrSelCri & " AND {HR_SECURE_FOLLOW_UP_8.USERID} ='" & glbUserID & "' AND {HR_SECURE_FOLLOW_UP_8.ACCESSABLE} = True"
'    glbstrSelCri = glbstrSelCri & " AND {HR_SECURE_FOLLOW_UP_9.USERID} ='" & glbUserID & "' AND {HR_SECURE_FOLLOW_UP_9.ACCESSABLE} = True"
'    glbstrSelCri = glbstrSelCri & " AND {HR_SECURE_FOLLOW_UP_10.USERID} ='" & glbUserID & "' AND {HR_SECURE_FOLLOW_UP_10.ACCESSABLE} = True"
'    glbstrSelCri = glbstrSelCri & " AND {HR_SECURE_FOLLOW_UP_11.USERID} ='" & glbUserID & "' AND {HR_SECURE_FOLLOW_UP_11.ACCESSABLE} = True"
'    glbstrSelCri = glbstrSelCri & " AND {HR_SECURE_FOLLOW_UP_12.USERID} ='" & glbUserID & "' AND {HR_SECURE_FOLLOW_UP_12.ACCESSABLE} = True"
'    glbstrSelCri = glbstrSelCri & " AND {HR_SECURE_FOLLOW_UP_13.USERID} ='" & glbUserID & "' AND {HR_SECURE_FOLLOW_UP_13.ACCESSABLE} = True"
'    glbstrSelCri = glbstrSelCri & " AND {HR_SECURE_FOLLOW_UP_14.USERID} ='" & glbUserID & "' AND {HR_SECURE_FOLLOW_UP_14.ACCESSABLE} = True"
'    glbstrSelCri = glbstrSelCri & " AND {HR_SECURE_FOLLOW_UP_15.USERID} ='" & glbUserID & "' AND {HR_SECURE_FOLLOW_UP_15.ACCESSABLE} = True"
'    glbstrSelCri = glbstrSelCri & " AND {HR_SECURE_FOLLOW_UP_16.USERID} ='" & glbUserID & "' AND {HR_SECURE_FOLLOW_UP_16.ACCESSABLE} = True"
'    glbstrSelCri = glbstrSelCri & " AND {HR_SECURE_FOLLOW_UP_17.USERID} ='" & glbUserID & "' AND {HR_SECURE_FOLLOW_UP_17.ACCESSABLE} = True"
'    glbstrSelCri = glbstrSelCri & " AND {HR_SECURE_FOLLOW_UP_18.USERID} ='" & glbUserID & "' AND {HR_SECURE_FOLLOW_UP_18.ACCESSABLE} = True"
'    glbstrSelCri = glbstrSelCri & " AND {HR_SECURE_FOLLOW_UP_19.USERID} ='" & glbUserID & "' AND {HR_SECURE_FOLLOW_UP_19.ACCESSABLE} = True"
'    glbstrSelCri = glbstrSelCri & " AND {HR_SECURE_FOLLOW_UP_20.USERID} ='" & glbUserID & "' AND {HR_SECURE_FOLLOW_UP_20.ACCESSABLE} = True"
    
    
    glbstrSelCri = glbstrSelCri & " AND {HREMP_FLAGS.EF_WRKEMP}='" & glbUserID & "'"
    
    'Release 8.0 - Ticket #22682: View Own security
    'If View Own not checked then do not retrieve Employee Flags of the User/Employee No
    If Len(glbUserEmpNo) > 0 And glbUserEmpNo <> 0 And Not gSec_EmpFlags_ViewOwn Then
        'Do not show user's Employee Flags records based on the Employee # associated to the User.
        If Len(glbstrSelCri) > 0 Then
            glbstrSelCri = glbstrSelCri & " AND {HREMP_FLAGS.EF_EMPNBR} <> " & glbUserEmpNo
        Else
            glbstrSelCri = glbstrSelCri & " {HREMP_FLAGS.EF_EMPNBR} <> " & glbUserEmpNo
        End If
    End If
    
    
    If Len(glbstrSelCri) >= 0 Then
        Me.vbxCrystal.SelectionFormula = glbstrSelCri
    End If
  
    Me.vbxCrystal.Connect = RptODBC_SQL

SkipTo:
    Cri_SetAll = True
    Screen.MousePointer = DEFAULT

Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Employee Flags", "HREMP_FLAGS", "Select")
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
'imbeded in report

Cri_Sorts = 0
'first set primary grouping
Y% = 0
grpField$ = getEGroup(comGroup(0).Text)

Call setRptLabel(Me, 0)
Call setRptLabel(Me, 3)

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

GrpIdx% = comGroup(1).ListIndex
Select Case GrpIdx%
    Case 0: grpField$ = "{@EFullName}"
    Case 1: grpField$ = "{HREMP.ED_EMPNBR}" 'GROUP ON EMPLOYEE#
End Select
grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(1) = grpCond$

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

If Not clpPT.ListChecker Then
'If Len(clpPT) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

For x% = 0 To 5
    If Not clpCode(x).ListChecker Then Exit Function
Next x%

If Not elpEEID.ListChecker Then
    Exit Function
End If

If lstEmpFlagsList.SelCount = 0 Then
    MsgBox "At least one Employee Flag should be selected"
    lstEmpFlagsList.SetFocus
    Exit Function
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

CriCheck = True
End Function

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
Dim Y As Integer

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
glbOnTop = "FRMREMPFLAGS"

Screen.MousePointer = HOURGLASS
frmREmpFlags.Caption = "Employee Flags Report"
If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If

Call comGrpLoad
Call setRptCaption(Me)

'Ticket #24737 - Samuel, Son & Co., Limited
If glbCompSerial = "S/N - 2382W" Then
    chkShowFlagFollowUp.Visible = True
    chkShowFlagFollowUp.Caption = "Only Show Flags with " & lStr("Follow-ups")
Else
    chkShowFlagFollowUp.Visible = False
End If

If glbLinamar Then clpCode(3).MaxLength = 8
If glbCompSerial = "S/N - 2227W" Then clpCode(3).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6

'Load Employee Flags list
lstEmpFlagsList.Clear
For Y = 1 To 20
    If lStr("Employee Flag " & Y) <> "" Then
        lstEmpFlagsList.AddItem lStr("Employee Flag " & Y)
    End If
Next Y

Call INI_Controls(Me)

Call PopcboEmpFlag(0) 'Ticket #27811 Franks 12/08/2015

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

Private Sub lstEmpFlagsList_Click()
'    Dim Y As Integer
'    If chkShowEmp.Value = True Then
'        For Y = 0 To lstEmpFlagsList.ListCount - 1
'            If lstEmpFlagsList.selected(Y) Then
'            Else
'                chkShowEmp.Value = False
'                Exit For
'            End If
'        Next Y
'    End If
'    If chkShowEmp.Value = True Then
'    If lstEmpFlagsList.SelCount <> lstEmpFlagsList.ListCount Then
'        chkShowEmp.Value = False
'    End If
'    End If
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

Private Sub Populate_EmployeeFlag_WorkingTable()
    Dim rsHREmp As New ADODB.Recordset
    Dim rsEmpFlgWRK As New ADODB.Recordset
    Dim rsEmpFlg As New ADODB.Recordset
    Dim SQLQ, wrkSelCri As String
    Dim SQLQEmpFlg As String
    Dim xFlagNo As Integer
    Dim counter As Integer
    Dim flgHasValue As Boolean

    On Error GoTo PopulateFlagWrk_Error
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(2).Caption = ""
    MDIMain.panHelp(0).FloodPercent = 1
    MDIMain.panHelp(0).FloodPercent = 3
    
    Screen.MousePointer = HOURGLASS
    
    MDIMain.panHelp(0).FloodPercent = 15
    
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute "DELETE FROM HREMP_FLAGSWRK " & in_SQL(glbIHRDBW) & " WHERE EF_WRKEMP='" & glbUserID & "'"
    gdbAdoIhr001.CommitTrans
    MDIMain.panHelp(0).FloodPercent = 30

    wrkSelCri = Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")")
    
    'Get list of valid employees based on the Dept Security and selection criteria
    SQLQ = "SELECT * FROM HREMP WHERE " & wrkSelCri
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsHREmp.EOF
        'Retrieve Employee Flags of the employee
        SQLQEmpFlg = ""
        SQLQ = ""
        counter = 0
        SQLQEmpFlg = "SELECT * FROM HREMP_FLAGS WHERE EF_EMPNBR = " & rsHREmp("ED_EMPNBR")
        
        'Release 8.1
        'Date Selection Criteria. Compare with each Flag Dates 1 - 20
        For xFlagNo = 1 To lstEmpFlagsList.ListCount  '20
            If lstEmpFlagsList.selected(xFlagNo - 1) Then
                counter = counter + 1
                If IsDate(dlpDateRange(0).Text) And IsDate(dlpDateRange(1).Text) Then
                    If counter = 1 Then
                        SQLQ = SQLQ & " AND "
                    Else
                        SQLQ = SQLQ & " OR "
                    End If
                    SQLQ = SQLQ & "  ((EF_FLAGDTE" & xFlagNo & " >= " & Date_SQL(dlpDateRange(0).Text)
                    'Ticket #28040 - Do not show Employee Flags that do not have Flag Date entered when selection criteria is by Date
                    'SQLQ = SQLQ & " AND EF_FLAGDTE" & xFlagNo & " <= " & Date_SQL(dlpDateRange(1).Text) & ") OR (EF_FLAGDTE" & xFlagNo & " IS NULL))"
                    SQLQ = SQLQ & " AND EF_FLAGDTE" & xFlagNo & " <= " & Date_SQL(dlpDateRange(1).Text) & "))" 'OR (EF_FLAGDTE" & xFlagNo & " IS NULL))"
                Else
                    If IsDate(dlpDateRange(0).Text) Then
                        If counter = 1 Then
                            SQLQ = SQLQ & " AND "
                        Else
                            SQLQ = SQLQ & " OR "
                        End If
                        
                        'Ticket #28040 - Do not show Employee Flags that do not have Flag Date entered when selection criteria is by Date
                        'SQLQ = SQLQ & " AND (EF_FLAGDTE" & xFlagNo & " >= " & Date_SQL(dlpDateRange(0).Text) & " OR EF_FLAGDTE" & xFlagNo & " IS NULL)"
                        SQLQ = SQLQ & "  (EF_FLAGDTE" & xFlagNo & " >= " & Date_SQL(dlpDateRange(0).Text) & ")"  ' OR EF_FLAGDTE" & xFlagNo & " IS NULL)"
                    ElseIf IsDate(dlpDateRange(1).Text) Then
                        If counter = 1 Then
                            SQLQ = SQLQ & " AND "
                        Else
                            SQLQ = SQLQ & " OR "
                        End If
                        
                        'Ticket #28040 - Do not show Employee Flags that do not have Flag Date entered when selection criteria is by Date
                        'SQLQ = SQLQ & " AND (EF_FLAGDTE" & xFlagNo & " <= " & Date_SQL(dlpDateRange(1).Text) & " OR EF_FLAGDTE" & xFlagNo & " IS NULL)"
                        SQLQ = SQLQ & "  (EF_FLAGDTE" & xFlagNo & " <= " & Date_SQL(dlpDateRange(1).Text) & ")"  ' OR EF_FLAGDTE" & xFlagNo & " IS NULL)"
                    End If
                End If
            End If
        Next
        
        If Len(SQLQ) > 0 Or (Not IsDate(dlpDateRange(0).Text) And Not IsDate(dlpDateRange(1).Text)) Then
            SQLQ = SQLQEmpFlg & SQLQ
        Else
            SQLQ = SQLQEmpFlg & " AND 1 = 2"
        End If
        rsEmpFlg.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsEmpFlg.EOF Then
            'Add the employee flag record in the Working table based on the Follow Up Code, is found,
            'with the employee flag that user has access to.
            SQLQ = "SELECT * FROM HREMP_FLAGSWRK WHERE EF_EMPNBR = " & rsEmpFlg("EF_EMPNBR")
            SQLQ = SQLQ & " AND EF_WRKEMP = '" & glbUserID & "'"
            rsEmpFlgWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If rsEmpFlgWRK.EOF Then
                rsEmpFlgWRK.AddNew
                rsEmpFlgWRK("EF_COMPNO") = "001"
                rsEmpFlgWRK("EF_EMPNBR") = rsEmpFlg("EF_EMPNBR")
                'If Not IsMissing(xTerm_Seq) Then
                '    rsEmpFlgWRK("TERM_SEQ") = xTerm_Seq
                'End If
            End If
            
            For xFlagNo = 1 To lstEmpFlagsList.ListCount
                'Only show flags which are selected to show
                If chkShowEmp.Value = False Then
                    'Flag Selected to show?
                    If lstEmpFlagsList.selected(xFlagNo - 1) Then
                        'Check if Follow Up record existing then check if the user has access to it
                        'Ticket #23641 - Field name issue
                        If xFlagNo = 19 Then
                            If Len(rsEmpFlg("EF_FUDTE" & xFlagNo)) > 0 Then
                                If rsEmpFlg("EF_FTREAS" & xFlagNo) > 0 Then
                                    If Not FollowUp_Sec(rsEmpFlg("EF_FTREAS" & xFlagNo)) Then
                                        'User do not have access to this follow up code - do not add this flag record
                                        GoTo Next_EmpFlag
                                    End If
                                End If
                            End If
                        Else
                            If Len(rsEmpFlg("EF_FUDTE" & xFlagNo)) > 0 Then
                                If rsEmpFlg("EF_FUREAS" & xFlagNo) > 0 Then
                                    If Not FollowUp_Sec(rsEmpFlg("EF_FUREAS" & xFlagNo)) Then
                                        'User do not have access to this follow up code - do not add this flag record
                                        GoTo Next_EmpFlag
                                    End If
                                End If
                            End If
                        End If
                        
                        'When Flag Date selection criteria is entered only show Flags that meeting the criteria
                        'The logic below will be used with 'Hide Employee Flags with no values'
                        If (IsDate(dlpDateRange(0).Text) Or IsDate(dlpDateRange(1).Text)) Then
                            If IsDate(rsEmpFlg("EF_FLAGDTE" & xFlagNo)) Then
                                If IsDate(dlpDateRange(0).Text) And IsDate(dlpDateRange(1).Text) Then
                                    If CVDate(rsEmpFlg("EF_FLAGDTE" & xFlagNo)) < CVDate(dlpDateRange(0).Text) Or _
                                       CVDate(rsEmpFlg("EF_FLAGDTE" & xFlagNo)) > CVDate(dlpDateRange(1).Text) Then
                                       
                                        'Flag date outside the date range so do not show this flag in the report
                                        GoTo Next_EmpFlag
                                    End If
                                Else
                                    If IsDate(dlpDateRange(0).Text) Then
                                        If CVDate(rsEmpFlg("EF_FLAGDTE" & xFlagNo)) < CVDate(dlpDateRange(0).Text) Then
                                            'Flag date earlier than the From date so do not show this flag in the report
                                            GoTo Next_EmpFlag
                                        End If
                                    ElseIf IsDate(dlpDateRange(1).Text) Then
                                        If CVDate(rsEmpFlg("EF_FLAGDTE" & xFlagNo)) > CVDate(dlpDateRange(1).Text) Then
                                            'Flag date after To date so do not show this flag in the report
                                            GoTo Next_EmpFlag
                                        End If
                                    End If
                                End If
                            Else
                                'No Flag date entered so do not show this flag in the report
                                GoTo Next_EmpFlag
                            End If
                        End If
                                                
                        'Flag Value
                        If Len(rsEmpFlg("EF_FLAGVAL" & xFlagNo)) > 0 Then
                            rsEmpFlgWRK("EF_FLAGVAL" & xFlagNo) = rsEmpFlg("EF_FLAGVAL" & xFlagNo)
                        End If
                        
                        'Flag Date
                        If Len(rsEmpFlg("EF_FLAGDTE" & xFlagNo)) > 0 Then
                            If IsDate(rsEmpFlg("EF_FLAGDTE" & xFlagNo)) Then
                                rsEmpFlgWRK("EF_FLAGDTE" & xFlagNo) = CVDate(rsEmpFlg("EF_FLAGDTE" & xFlagNo))
                            End If
                        End If
                        
                        'Follow Up Date & Reason
                        If Len(rsEmpFlg("EF_FUDTE" & xFlagNo)) > 0 Then
                            If IsDate(rsEmpFlg("EF_FUDTE" & xFlagNo)) Then
                                rsEmpFlgWRK("EF_FUDTE" & xFlagNo) = CVDate(rsEmpFlg("EF_FUDTE" & xFlagNo))
                                'Ticket #23641 - Field name issue
                                If xFlagNo = 19 Then
                                    rsEmpFlgWRK("EF_FTREAS" & xFlagNo) = rsEmpFlg("EF_FTREAS" & xFlagNo)
                                Else
                                    rsEmpFlgWRK("EF_FUREAS" & xFlagNo) = rsEmpFlg("EF_FUREAS" & xFlagNo)
                                End If
                            End If
                        End If
                    End If
                Else
                    'Show All Flags
                    'Check if Follow Up record existing then check if the user has access to it
                    'Ticket #23641 - Field name issue
                    If xFlagNo = 19 Then
                        If Len(rsEmpFlg("EF_FUDTE" & xFlagNo)) > 0 Then
                            If rsEmpFlg("EF_FTREAS" & xFlagNo) > 0 Then
                                If Not FollowUp_Sec(rsEmpFlg("EF_FTREAS" & xFlagNo)) Then
                                    'User do not have access to this follow up code - do not add this flag record
                                    GoTo Next_EmpFlag
                                End If
                            End If
                        End If
                    Else
                        If Len(rsEmpFlg("EF_FUDTE" & xFlagNo)) > 0 Then
                            If rsEmpFlg("EF_FUREAS" & xFlagNo) > 0 Then
                                If Not FollowUp_Sec(rsEmpFlg("EF_FUREAS" & xFlagNo)) Then
                                    'User do not have access to this follow up code - do not add this flag record
                                    GoTo Next_EmpFlag
                                End If
                            End If
                        End If
                    End If
                    
                    'When Flag Date selection criteria is entered only show Flags that meeting the criteria
                    'The logic below will be used with 'Hide Employee Flags with no values'
                    If (IsDate(dlpDateRange(0).Text) Or IsDate(dlpDateRange(1).Text)) Then
                        If IsDate(rsEmpFlg("EF_FLAGDTE" & xFlagNo)) Then
                            If IsDate(dlpDateRange(0).Text) And IsDate(dlpDateRange(1).Text) Then
                                If CVDate(rsEmpFlg("EF_FLAGDTE" & xFlagNo)) < CVDate(dlpDateRange(0).Text) Or _
                                   CVDate(rsEmpFlg("EF_FLAGDTE" & xFlagNo)) > CVDate(dlpDateRange(1).Text) Then
                                   
                                    'Flag date outside the date range so do not show this flag in the report
                                    GoTo Next_EmpFlag
                                End If
                            Else
                                If IsDate(dlpDateRange(0).Text) Then
                                    If CVDate(rsEmpFlg("EF_FLAGDTE" & xFlagNo)) < CVDate(dlpDateRange(0).Text) Then
                                        'Flag date earlier than the From date so do not show this flag in the report
                                        GoTo Next_EmpFlag
                                    End If
                                ElseIf IsDate(dlpDateRange(1).Text) Then
                                    If CVDate(rsEmpFlg("EF_FLAGDTE" & xFlagNo)) > CVDate(dlpDateRange(1).Text) Then
                                        'Flag date after To date so do not show this flag in the report
                                        GoTo Next_EmpFlag
                                    End If
                                End If
                            End If
                        Else
                            'No Flag date entered so do not show this flag in the report
                            GoTo Next_EmpFlag
                        End If
                    End If
                    
                    'Flag Value
                    If Len(rsEmpFlg("EF_FLAGVAL" & xFlagNo)) > 0 Then
                        rsEmpFlgWRK("EF_FLAGVAL" & xFlagNo) = rsEmpFlg("EF_FLAGVAL" & xFlagNo)
                    End If
                    
                    'Flag Date
                    If Len(rsEmpFlg("EF_FLAGDTE" & xFlagNo)) > 0 Then
                        If IsDate(rsEmpFlg("EF_FLAGDTE" & xFlagNo)) Then
                            rsEmpFlgWRK("EF_FLAGDTE" & xFlagNo) = CVDate(rsEmpFlg("EF_FLAGDTE" & xFlagNo))
                        End If
                    End If
                    
                    'Follow Up Date & Reason
                    If Len(rsEmpFlg("EF_FUDTE" & xFlagNo)) > 0 Then
                        If IsDate(rsEmpFlg("EF_FUDTE" & xFlagNo)) Then
                            rsEmpFlgWRK("EF_FUDTE" & xFlagNo) = CVDate(rsEmpFlg("EF_FUDTE" & xFlagNo))
                            'Ticket #23641 - Field name issue
                            If xFlagNo = 19 Then
                                rsEmpFlgWRK("EF_FTREAS" & xFlagNo) = rsEmpFlg("EF_FTREAS" & xFlagNo)
                            Else
                                rsEmpFlgWRK("EF_FUREAS" & xFlagNo) = rsEmpFlg("EF_FUREAS" & xFlagNo)
                            End If
                        End If
                    End If
                End If
Next_EmpFlag:
            Next
            
            rsEmpFlgWRK("EF_LDATE") = rsEmpFlg("EF_LDATE")
            rsEmpFlgWRK("EF_LTIME") = rsEmpFlg("EF_LTIME")
            rsEmpFlgWRK("EF_LUSER") = rsEmpFlg("EF_LUSER")
            rsEmpFlgWRK("EF_WRKEMP") = glbUserID
            rsEmpFlgWRK.Update
            
            rsEmpFlgWRK.Close
            Set rsEmpFlgWRK = Nothing
        End If
        rsEmpFlg.Close
        Set rsEmpFlg = Nothing
        
        rsHREmp.MoveNext
    Loop
    rsHREmp.Close
    Set rsHREmp = Nothing
    
    'Hide Employees with no Employee Flags values
    'Do this by deleting the Employee from the work table with not Employee Flag values
    If chkHideEmpNoFlags Then
        SQLQ = "SELECT * FROM HREMP_FLAGSWRK "
        SQLQ = SQLQ & " WHERE EF_WRKEMP = '" & glbUserID & "'"
        SQLQ = SQLQ & " ORDER BY EF_WRKEMP"
        rsEmpFlgWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Do While Not rsEmpFlgWRK.EOF
            'Flag to determine Delete or Not for each employee
            flgHasValue = False
            
            'Check each Flag to see if it has value.
            'At least one flag have to have a value to exit the loop and not delete the employee, otherwise the
            'employee gets deleted
            For xFlagNo = 1 To lstEmpFlagsList.ListCount
                'Flag Selected to show?
                If lstEmpFlagsList.selected(xFlagNo - 1) Then
                    'Flag Value
                    If Len(rsEmpFlgWRK("EF_FLAGVAL" & xFlagNo)) > 0 Then
                        flgHasValue = True
                        Exit For
                    End If
                    
                    'Flag Date
                    If Len(rsEmpFlgWRK("EF_FLAGDTE" & xFlagNo)) > 0 Then
                        If IsDate(rsEmpFlgWRK("EF_FLAGDTE" & xFlagNo)) Then
                            flgHasValue = True
                            Exit For
                        End If
                    End If
                    
                    'Follow Up Date & Reason
                    If Len(rsEmpFlgWRK("EF_FUDTE" & xFlagNo)) > 0 Then
                        If IsDate(rsEmpFlgWRK("EF_FUDTE" & xFlagNo)) Then
                            
                            'Ticket #23641 - Field name issue
                            If xFlagNo = 19 Then
                                If Len(rsEmpFlgWRK("EF_FTREAS" & xFlagNo)) > 0 Then
                                    flgHasValue = True
                                    Exit For
                                End If
                            Else
                                If Len(rsEmpFlgWRK("EF_FUREAS" & xFlagNo)) > 0 Then
                                    flgHasValue = True
                                    Exit For
                                End If
                            End If
                            
                            flgHasValue = True
                            Exit For
                        End If
                    End If
                End If
            Next
            
            'Delete?
            If flgHasValue = False Then
                'No Flag values found for this employee, delete the employee from the work table so it does not display
                'with blank values in the report.
                rsEmpFlgWRK.Delete
            End If
            rsEmpFlgWRK.MoveNext
        Loop
        rsEmpFlgWRK.Close
        Set rsEmpFlgWRK = Nothing
    End If
    
MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""

Exit Sub

PopulateFlagWrk_Error:
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    
    Exit Sub
End Sub

Private Function FollowUp_Sec(xCode) As Boolean
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim retval As Boolean
    Dim xTemplate As String
    
    '????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
    xTemplate = ""
    xTemplate = Get_Template(glbUserID)
    
    
    'strSQL = "SELECT MAINTAINABLE FROM HR_SECURE_FOLLOW_UP WHERE "
    strSQL = "SELECT ACCESSABLE FROM HR_SECURE_FOLLOW_UP WHERE "
    If xTemplate = "" Or xTemplate = "TEMPLATE" Then
        strSQL = strSQL & "CODENAME='" & xCode & "' AND USERID='" & Replace(glbUserID, "'", "''") & "'"
    Else
        '????Ticket #24808 -  Retrieve template's security profile
        strSQL = strSQL & "CODENAME='" & xCode & "' AND USERID='" & Replace(xTemplate, "'", "''") & "'"
    End If
    rs.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If Not rs.EOF Then
        'retVal = Abs(rs("MAINTAINABLE"))
        retval = Abs(rs("ACCESSABLE"))
    Else
        retval = False
    End If
    
    FollowUp_Sec = retval
End Function

Private Sub Generate_ExcelReport()
    Dim rsHREmp As New ADODB.Recordset
    Dim rsEmpFlagLbl As New ADODB.Recordset
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim SQLQ, sSQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim I, totNum
    Dim Total As Integer
    Dim xFlagNo As Integer
    Dim xCol As Integer
    Dim xEmpFlag As Boolean
    Dim xShowFlag As Boolean
    Dim xSelFlgNo As Integer
        
    On Error GoTo Err_Generate_ExcelReport
        
    Screen.MousePointer = HOURGLASS

    Total = 0
    
    SQLQ = "SELECT ED_EMPNBR, ED_SURNAME, ED_FNAME, ED_SECTION, HREMP_FLAGSWRK.* FROM HREMP"
    SQLQ = SQLQ & " INNER JOIN HREMP_FLAGSWRK ON ED_EMPNBR = EF_EMPNBR"
    SQLQ = SQLQ & " WHERE EF_WRKEMP='" & glbUserID & "'"
    SQLQ = SQLQ & " AND ((EF_FUDTE1 is not null AND EF_FUDTE1 <> '') OR"
    SQLQ = SQLQ & " (EF_FUDTE2 is not null AND EF_FUDTE2 <> '') OR"
    SQLQ = SQLQ & " (EF_FUDTE3 is not null AND EF_FUDTE3 <> '') OR"
    SQLQ = SQLQ & " (EF_FUDTE4 is not null AND EF_FUDTE4 <> '') OR"
    SQLQ = SQLQ & " (EF_FUDTE5 is not null AND EF_FUDTE5 <> '') OR"
    SQLQ = SQLQ & " (EF_FUDTE6 is not null AND EF_FUDTE6 <> '') OR"
    SQLQ = SQLQ & " (EF_FUDTE7 is not null AND EF_FUDTE7 <> '') OR"
    SQLQ = SQLQ & " (EF_FUDTE8 is not null AND EF_FUDTE8 <> '') OR"
    SQLQ = SQLQ & " (EF_FUDTE9 is not null AND EF_FUDTE9 <> '') OR"
    SQLQ = SQLQ & " (EF_FUDTE10 is not null AND EF_FUDTE10 <> '') OR"
    SQLQ = SQLQ & " (EF_FUDTE11 is not null AND EF_FUDTE11 <> '') OR"
    SQLQ = SQLQ & " (EF_FUDTE12 is not null AND EF_FUDTE12 <> '') OR"
    SQLQ = SQLQ & " (EF_FUDTE13 is not null AND EF_FUDTE13 <> '') OR"
    SQLQ = SQLQ & " (EF_FUDTE14 is not null AND EF_FUDTE14 <> '') OR"
    SQLQ = SQLQ & " (EF_FUDTE15 is not null AND EF_FUDTE15 <> '') OR"
    SQLQ = SQLQ & " (EF_FUDTE16 is not null AND EF_FUDTE16 <> '') OR"
    SQLQ = SQLQ & " (EF_FUDTE17 is not null AND EF_FUDTE17 <> '') OR"
    SQLQ = SQLQ & " (EF_FUDTE18 is not null AND EF_FUDTE18 <> '') OR"
    SQLQ = SQLQ & " (EF_FUDTE19 is not null AND EF_FUDTE19 <> '') OR"
    SQLQ = SQLQ & " (EF_FUDTE20 is not null AND EF_FUDTE20 <> ''))"
    
    'Release 8.0 - Ticket #22682: View Own security
    'If View Own not checked then do not retrieve Employee Flags of the User/Employee No
    If Len(glbUserEmpNo) > 0 And glbUserEmpNo <> 0 And Not gSec_EmpFlags_ViewOwn Then
        'Do not show user's Employee Flags records based on the Employee # associated to the User.
        SQLQ = SQLQ & " AND EF_EMPNBR <> " & glbUserEmpNo
    End If
    
    If comGroup(1).ListIndex = 0 Then
        SQLQ = SQLQ & " ORDER BY ED_SURNAME,ED_FNAME,ED_EMPNBR "
    Else
        SQLQ = SQLQ & " ORDER BY ED_EMPNBR,ED_SURNAME,ED_FNAME "
    End If
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsHREmp.EOF Then
        rsHREmp.MoveFirst
        totNum = rsHREmp.RecordCount: I = 0
                
        'File to export to
        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "EmployeeFlagsTmp.xls"
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "EmployeeFlags_" & Trim(glbUserID) & ".xls"
    
        If Dir(xlsFileTmp) = "" Then
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
        End If
        If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait"
        MDIMain.panHelp(0).FloodPercent = 0
        
        FileCopy xlsFileTmp, xlsFileMat
    
        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)
        
        exSheet.Cells(3, 1) = "Employee #"
        exSheet.Cells(3, 2) = "Employee Name"
        exSheet.Cells(3, 3) = lStr("Section") & " Description"
        
        'Print Employee Flags headings
        xCol = 4
        SQLQ = "SELECT LB_ORG,LB_NEW FROM HRLABEL WHERE LB_ORG LIKE 'Employee Flag %' AND LB_LANG = 'EN'"
        rsEmpFlagLbl.Open SQLQ, gdbAdoIhr001, adOpenStatic
        Do While Not rsEmpFlagLbl.EOF
            'Only show the flags that has been renamed
            If rsEmpFlagLbl("LB_ORG") <> rsEmpFlagLbl("LB_NEW") And rsEmpFlagLbl("LB_NEW") <> "" Then
                'Show Flags which are selected to show.
                If chkShowEmp.Value = False Then
                    'Find out which Flag is checked using the original employee flag label and flaglist
                    If lstEmpFlagsList.selected(Trim(Right(rsEmpFlagLbl("LB_ORG"), 2)) - 1) Then
                        'Flag checked - print the label in excel
                        exSheet.Cells(3, xCol) = rsEmpFlagLbl("LB_NEW")
                        
                        'Move to next column in excel for next employee flag
                        xCol = xCol + 1
                    End If
                Else
                    'Show all Employee flags
                    exSheet.Cells(3, xCol) = rsEmpFlagLbl("LB_NEW")
                    
                    'Move to next column in excel for next employee flag
                    xCol = xCol + 1
                End If
            End If
            
            rsEmpFlagLbl.MoveNext
        Loop
        rsEmpFlagLbl.Close
        Set rsEmpFlagLbl = Nothing
        
        xRow = 4
                
        Do While Not rsHREmp.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
                        
            'Check if user has access to Follow Up Codes then only display the record of the employee
            xEmpFlag = False
            xShowFlag = False
            xSelFlgNo = 0
            
            'Only check flags which are selected
            If chkShowEmp.Value = False Then
                'Check Security on Flags selected only
                For xFlagNo = 1 To lstEmpFlagsList.ListCount
                    If lstEmpFlagsList.selected(xFlagNo - 1) Then
                        If xFlagNo = 19 Then
                            If Len(rsHREmp("EF_FUDTE" & xFlagNo)) > 0 Then
                                If rsHREmp("EF_FTREAS" & xFlagNo) > 0 Then
                                    If FollowUp_Sec(rsHREmp("EF_FTREAS" & xFlagNo)) Then
                                        'User have access to this follow up code
                                        xEmpFlag = True
                                    End If
                                End If
                            End If
                        Else
                            If Len(rsHREmp("EF_FUDTE" & xFlagNo)) > 0 Then
                                If rsHREmp("EF_FUREAS" & xFlagNo) > 0 Then
                                    If FollowUp_Sec(rsHREmp("EF_FUREAS" & xFlagNo)) Then
                                        'User have access to this follow up code
                                        xEmpFlag = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            Else
                For xFlagNo = 1 To 20
                    If xFlagNo = 19 Then
                        If Len(rsHREmp("EF_FUDTE" & xFlagNo)) > 0 Then
                            If rsHREmp("EF_FTREAS" & xFlagNo) > 0 Then
                                If FollowUp_Sec(rsHREmp("EF_FTREAS" & xFlagNo)) Then
                                    'User have access to this follow up code
                                    xEmpFlag = True
                                End If
                            End If
                        End If
                    Else
                        If Len(rsHREmp("EF_FUDTE" & xFlagNo)) > 0 Then
                            If rsHREmp("EF_FUREAS" & xFlagNo) > 0 Then
                                If FollowUp_Sec(rsHREmp("EF_FUREAS" & xFlagNo)) Then
                                    'User have access to this follow up code
                                    xEmpFlag = True
                                End If
                            End If
                        End If
                    End If
                Next
            End If
            
            If xEmpFlag = True Then
                exSheet.Cells(xRow, 1) = rsHREmp("ED_EMPNBR")
                exSheet.Cells(xRow, 2) = rsHREmp("ED_SURNAME") & ", " & rsHREmp("ED_FNAME")
                exSheet.Cells(xRow, 3) = GetTABLDesc("EDSE", rsHREmp("ED_SECTION"))
                
                'Only show flags which are selected to show
                If chkShowEmp.Value = False Then
                    'Print Follow Up in the respective column if the user has rights to the Follow Up Code
                    For xFlagNo = 1 To lstEmpFlagsList.ListCount '20
                        'Only show flags which are selected to show
                        If lstEmpFlagsList.selected(xFlagNo - 1) Then
                            'This flag is selected to show
                            xShowFlag = True
                            xSelFlgNo = xSelFlgNo + 1
                        Else
                            'This flag is not selected to show
                            xShowFlag = False
                        End If
                        
                        'Show only flag selected to show
                        If xShowFlag Then
                            If xFlagNo = 19 Then
                                If rsHREmp("EF_FTREAS" & xFlagNo) > 0 Then
                                    If FollowUp_Sec(rsHREmp("EF_FTREAS" & xFlagNo)) Then
                                        exSheet.Cells(xRow, xSelFlgNo + 3) = rsHREmp("EF_FTREAS" & xFlagNo) & " - " & rsHREmp("EF_FUDTE" & xFlagNo)
                                    End If
                                End If
                            Else
                                If rsHREmp("EF_FUREAS" & xFlagNo) > 0 Then
                                    If FollowUp_Sec(rsHREmp("EF_FUREAS" & xFlagNo)) Then
                                        exSheet.Cells(xRow, xSelFlgNo + 3) = rsHREmp("EF_FUREAS" & xFlagNo) & " - " & rsHREmp("EF_FUDTE" & xFlagNo)
                                    End If
                                End If
                            End If
                        End If
                    Next
                Else
                    'Show all
                    'Print Follow Up in the respective column if the user has rights to the Follow Up Code
                    For xFlagNo = 1 To 20
                        'Show all Flags
                        xShowFlag = True
                        
                        'Show only flag selected to show
                        If xShowFlag Then
                            If xFlagNo = 19 Then
                                If rsHREmp("EF_FTREAS" & xFlagNo) > 0 Then
                                    If FollowUp_Sec(rsHREmp("EF_FTREAS" & xFlagNo)) Then
                                        exSheet.Cells(xRow, xFlagNo + 3) = rsHREmp("EF_FTREAS" & xFlagNo) & " - " & rsHREmp("EF_FUDTE" & xFlagNo)
                                    End If
                                End If
                            Else
                                If rsHREmp("EF_FUREAS" & xFlagNo) > 0 Then
                                    If FollowUp_Sec(rsHREmp("EF_FUREAS" & xFlagNo)) Then
                                        exSheet.Cells(xRow, xFlagNo + 3) = rsHREmp("EF_FUREAS" & xFlagNo) & " - " & rsHREmp("EF_FUDTE" & xFlagNo)
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
                xRow = xRow + 1
            End If
            
            rsHREmp.MoveNext
        Loop
        
        'exSheet.AutoFilterMode = True
        'exSheet.Range("K1:K20").AutoFilter Field:=1
        'exSheet.Columns.AutoFit

        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing
    

        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
    End If
    
    rsHREmp.Close
    Set rsHREmp = Nothing

    Screen.MousePointer = vbDefault
    
Exit Sub

Err_Generate_ExcelReport:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", "HREMP_FLAGSWRK", "SELECT")
Resume Next
    
End Sub

Private Sub PopcboEmpFlag(x As Long)
Dim n
Dim vList
    cboEmpFlag(x).Clear
    vList = FlgList(x)
    n = 1
    Do While n > 0
        n = InStr(vList, "&")
        If n > 0 Then
            cboEmpFlag(x).AddItem Trim(Left(vList, n - 1))
            vList = Mid(vList, n + 1)
        Else
            cboEmpFlag(x).AddItem Trim(vList)
        End If
    Loop
End Sub

Private Function FlgList(idx As Long) As String
Dim retval As String, FlgFile
retval = ""
FlgFile = glbIHRREPORTS & "FlagList" & CStr(idx) & ".MTF"

On Error GoTo ErrorHandler

If File(FlgFile) Then
    Open FlgFile For Input As #1
    Input #1, retval
    Close #1
End If

ResumeHere:
If InStr(retval, Trim(cboEmpFlag(idx))) = 0 And Trim(cboEmpFlag(idx)) <> "" Then
    retval = retval & "&" & Trim(cboEmpFlag(idx))
    cboEmpFlag(idx).AddItem Trim(cboEmpFlag(idx))
End If
Open FlgFile For Output As #1
Print #1, retval
Close #1
FlgList = retval
Exit Function

ErrorHandler:
If Err.Number = 62 Then
    ' Corrupted CountryList.MTF, kill it and regenerate
    Close #1
    MsgBox "Found corrupt FlagList" & CStr(idx) & ".MTF.  info:HR will re-create this file.", vbInformation + vbOKOnly, "Corrupted Country List"
    Kill FlgFile
    Resume ResumeHere
Else
    'MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number & " in CountryList"
    Resume Next
End If
End Function
