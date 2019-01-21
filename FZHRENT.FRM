VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRHrEnt 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Hourly Entitlements Report"
   ClientHeight    =   9000
   ClientLeft      =   255
   ClientTop       =   600
   ClientWidth     =   11085
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
   ScaleHeight     =   9000
   ScaleWidth      =   11085
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2355
      MaxLength       =   4
      TabIndex        =   19
      Tag             =   "00-00-Employee Position Shift"
      Top             =   5190
      Visible         =   0   'False
      Width           =   450
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   7
      Tag             =   "00-Entitlement Codes"
      Top             =   2820
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "ADRE"
   End
   Begin VB.CheckBox chkVDesc 
      Caption         =   "View Entitlement Descriptions"
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
      Left            =   7290
      TabIndex        =   11
      Top             =   2820
      Width           =   2535
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
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Tag             =   "First Level of grouping records"
      Top             =   6990
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
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Tag             =   "First Level of grouping records"
      Top             =   6630
      Width           =   2325
   End
   Begin VB.ComboBox cmbHours 
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
      Left            =   2355
      TabIndex        =   13
      Tag             =   "Choose display in hours or days"
      Text            =   "Combo1"
      Top             =   3510
      Width           =   1215
   End
   Begin Threed.SSFrame fraCodeDesc 
      Height          =   1395
      Left            =   6960
      TabIndex        =   40
      Top             =   3600
      Visible         =   0   'False
      Width           =   2790
      _Version        =   65536
      _ExtentX        =   4921
      _ExtentY        =   2461
      _StockProps     =   14
      Caption         =   "Entitlement Code Descriptions"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   1
      Begin VB.Label lblCodeDesc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Unassigned"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   9
         Left            =   300
         TabIndex        =   44
         Top             =   975
         Width           =   840
      End
      Begin VB.Label lblCodeDesc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Unassigned"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   8
         Left            =   285
         TabIndex        =   43
         Top             =   765
         Width           =   840
      End
      Begin VB.Label lblCodeDesc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Unassigned"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   7
         Left            =   285
         TabIndex        =   42
         Top             =   555
         Width           =   840
      End
      Begin VB.Label lblCodeDesc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Unassigned"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   6
         Left            =   285
         TabIndex        =   41
         Top             =   345
         Width           =   840
      End
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   2
      Tag             =   "00-Enter Location Code"
      Top             =   1140
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
      Top             =   810
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
      Top             =   480
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
      Index           =   2
      Left            =   2040
      TabIndex        =   6
      Tag             =   "00-Enter Status Code"
      Top             =   2490
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
      Left            =   2040
      TabIndex        =   5
      Tag             =   "00-Enter Union Code"
      Top             =   2160
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
      TabIndex        =   12
      Tag             =   "EDPT-Category"
      Top             =   3165
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
      Index           =   4
      Left            =   2040
      TabIndex        =   15
      Tag             =   "00-Enter Administered By Code"
      Top             =   4200
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
      TabIndex        =   16
      Tag             =   "00-Enter Section Code"
      Top             =   4530
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
      TabIndex        =   14
      Tag             =   "00-Enter Region Code"
      Top             =   3870
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   3690
      TabIndex        =   18
      Tag             =   "40-Date upto and including this date forward"
      Top             =   4860
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
      TabIndex        =   17
      Tag             =   "40-Date from and including this date forward"
      Top             =   4860
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   7
      Left            =   3330
      TabIndex        =   8
      Tag             =   "00-Entitlement Codes"
      Top             =   2820
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "ADRE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   8
      Left            =   4650
      TabIndex        =   9
      Tag             =   "00-Entitlement Codes"
      Top             =   2820
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "ADRE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   9
      Left            =   5940
      TabIndex        =   10
      Tag             =   "00-Entitlement Codes"
      Top             =   2820
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "ADRE"
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Tag             =   "10-Enter Employee Number"
      Top             =   1470
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   0
      Top             =   7485
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowWidth     =   480
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   2
      WindowControls  =   -1  'True
      MarginTop       =   720
      MarginBottom    =   720
      PrintFileLinesPerPage=   60
   End
   Begin INFOHR_Controls.EmployeeLookup elpReptau 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Tag             =   "10-Reporting Authority"
      Top             =   1800
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Threed.SSCheck chkAtt 
      Height          =   255
      Left            =   180
      TabIndex        =   20
      Tag             =   "If X-Show Attendance Details"
      Top             =   5640
      Width           =   2835
      _Version        =   65536
      _ExtentX        =   5001
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "   Show Attendance Details"
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
   Begin VB.Label lblRepAuth 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "AttSupervisor"
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
      Left            =   180
      TabIndex        =   46
      Top             =   1845
      Width           =   945
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
      Left            =   180
      TabIndex        =   45
      Top             =   5235
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label lblAttCodes 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Entitlement Codes"
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
      Left            =   180
      TabIndex        =   39
      Top             =   2865
      Width           =   1275
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
      Left            =   180
      TabIndex        =   38
      Top             =   3210
      Width           =   630
   End
   Begin VB.Label Label2 
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
      Left            =   180
      TabIndex        =   37
      Top             =   4905
      Width           =   1095
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
      Left            =   180
      TabIndex        =   36
      Top             =   4575
      Width           =   540
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
      Left            =   180
      TabIndex        =   35
      Top             =   3915
      Width           =   510
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
      Left            =   180
      TabIndex        =   34
      Top             =   4245
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
      Left            =   180
      TabIndex        =   33
      Top             =   1185
      Width           =   615
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
      Top             =   6360
      Width           =   1695
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
      Height          =   375
      Left            =   120
      TabIndex        =   31
      Top             =   120
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
      Index           =   1
      Left            =   60
      TabIndex        =   30
      Top             =   7050
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
      Left            =   60
      TabIndex        =   29
      Top             =   6690
      Width           =   885
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Display?"
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
      Index           =   5
      Left            =   180
      TabIndex        =   28
      Top             =   3570
      Width           =   600
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
      Left            =   180
      TabIndex        =   27
      Top             =   2535
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
      Left            =   180
      TabIndex        =   26
      Top             =   2205
      Width           =   420
   End
   Begin VB.Label lblTitle 
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
      Index           =   2
      Left            =   180
      TabIndex        =   25
      Top             =   1515
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
      Left            =   180
      TabIndex        =   24
      Top             =   855
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
      Left            =   180
      TabIndex        =   23
      Top             =   525
      Width           =   555
   End
End
Attribute VB_Name = "frmRHrEnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim localWdate$, localEntOutstanding$

Private Sub chkVDesc_Click()
fraCodeDesc.Visible = Not chkVDesc = 0
End Sub

Private Sub clpCode_LostFocus(Index As Integer)
On Error Resume Next: lblCodeDesc(Index).Caption = clpCode(Index).Caption
End Sub

Private Sub cmbHours_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim X%

On Error GoTo PrntErr

If CriCheck() Then

    If Not PrtForm("Entitlements Report Criteria", Me) Then Exit Sub
     Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    Screen.MousePointer = HOURGLASS

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

'Private Sub cmdPrint_GotFocus()
' Call SetPanHelp(Me.ActiveControl)
'End Sub

Public Sub cmdView_Click()
Dim X%
Dim strWHand As String
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "ATTEND", "SELECT")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Private Sub cmdView_GotFocus()
' Call SetPanHelp(Me.ActiveControl)
'End Sub

Private Sub comGroup_GotFocus(Index As Integer)
 Call SetPanHelp(Me.ActiveControl)
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

Private Sub Cri_Sup()
Dim EECri As String

If Len(elpReptau.Text) > 0 Then
    EECri = "{HR_JOB_HISTORY.JH_REPTAU} IN [" & getEmpnbr(elpReptau.Text) & "] "
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
Call Cri_PT
Call Cri_Shift
Call Cri_EE
Call Cri_Sup
' dkostka - 07/05/2001 - 'Section' selection criteria was being ignored, fixed.
For X% = 0 To 5
    Call Cri_Code(X%)
Next X%
Call Cri_FTDates

Call Cri_Attend


If glbCompSerial = "S/N - 2235W" Or glbCompSerial = "S/N - 2236W" Then   'laura 03/09/98
    If comGroup(0) = "(none)" Then   'was 4 laura dec 10, 1997
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rshenths.rpt"
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rshenth1.rpt"
    End If
Else
    If comGroup(0) = "(none)" Then   'was 4 laura dec 10, 1997
        If cmbHours.ListIndex = 0 Then
            If chkAtt Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzhenths_D.rpt"
            Else
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzhenths.rpt"
            End If
        Else
            If chkAtt Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzhentds_D.rpt"
            Else
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzhentds.rpt"
            End If
        End If
    Else
        If cmbHours.ListIndex = 0 Then
            If chkAtt Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzhenth1_D.rpt"
            Else
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzhenth1.rpt"
            End If
        Else
            If chkAtt Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzhentd1_D.rpt"
            Else
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzhentd1.rpt"
            End If
        End If
    End If
End If

'set to sorting/grouping criteria
X% = Cri_Sorts()   ' returns number of sections formated

'set location for database tables
If Len(glbstrSelCri) >= 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If

Me.vbxCrystal.Connect = RptODBC_SQL

'window title if appropriate
Me.vbxCrystal.WindowTitle = "Hourly Entitlements Report"

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
Dim X%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

Cri_Sorts = 0
If comGroup(0).Text = lStr("Rept. Authority 1") Then
    grpField$ = "{@fldAdSuper}"
Else
    If comGroup(0).Text = "Employee Name" Then
        grpField$ = "{@EEFullName}"
    Else
        grpField$ = getEGroup(comGroup(0).Text)
        If grpField$ = "(none)" Then grpField$ = "{HRPARCO.PC_NAME}"
    End If
End If
dscGroup$ = comGroup(0).Text
dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
Me.vbxCrystal.Formulas(0) = dscGroup$

grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(0) = grpCond$
grpCond$ = "GROUP" & CStr(2) & ";" & "{@EEFullName}" & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(1) = grpCond$
Me.vbxCrystal.SectionFormat(0) = "GH1;F;T;X;X;X;X;X"
Me.vbxCrystal.SectionFormat(1) = "GF1;F,F;F;X;X;X;X"

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

If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

For X% = 0 To 5
    If Not clpCode(X).ListChecker Then Exit Function
Next X%

For X% = 0 To 1
 If Len(dlpDateRange(X%).Text) > 0 Then
    If Not IsDate(dlpDateRange(X%).Text) Then
        MsgBox "     Not a valid date      "
        dlpDateRange(X%).Text = ""
        dlpDateRange(X%).SetFocus
        Exit Function
    End If
 End If
Next X%

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
glbOnTop = "FRMRHRENT"
                        
Screen.MousePointer = HOURGLASS

If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If

If Not glbSQL Then
    chkAtt.Visible = False
End If

Call setRptCaption(Me)
Call setCaption(lblRepAuth)
Call comGrpLoad

If lblRepAuth.Caption = "AttSupervisor" Then lblRepAuth.Caption = "Supervisor"

If glbCompSerial = "S/N - 2235W" Then   'laura 03/09/98
    lblTitle(5).Visible = False
    cmbHours.Visible = False
ElseIf glbCompSerial = "S/N - 2236W" Then
    lblTitle(5).Visible = False
    cmbHours.Visible = False
Else
    lblTitle(5).Visible = True
    cmbHours.Visible = True
End If

If glbLinamar Then clpCode(3).MaxLength = 8
If glbCompSerial = "S/N - 2227W" Then clpCode(3).MaxLength = 6
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

Private Sub Cri_FTDates()
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%, X%
Dim FromDate, ToDate, SQLQ
Dim RsHRPARCO As New ADODB.Recordset

If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    TempCri = "({HRENTHRS.HE_FDATE} "
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    TempCri = TempCri & " >= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) and "
    
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    TempCri = TempCri & " ({HRENTHRS.HE_TDATE} <= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    GoTo Cri_FTDatst
Else
    SQLQ = "Select * from HRPARCO"
    RsHRPARCO.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not RsHRPARCO.EOF Then
        FromDate = RsHRPARCO("PC_FDATE")
        ToDate = RsHRPARCO("PC_TDATE")
    End If
    RsHRPARCO.Close
End If

If Len(dlpDateRange(0).Text) > 0 Or Len(dlpDateRange(1).Text) > 0 Then
    If Len(dlpDateRange(0).Text) > 0 Then
        TempCri = "({HRENTHRS.HE_FDATE} "
        TempCri = TempCri & " >= "
        dtYYY% = Year(dlpDateRange(0).Text)
        dtMM% = month(dlpDateRange(0).Text)
        dtDD% = Day(dlpDateRange(0).Text)
        TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        If IsDate(ToDate) Then
            dtYYY% = Year(ToDate)
            dtMM% = month(ToDate)
            dtDD% = Day(ToDate)
            TempCri = TempCri & " and ({HRENTHRS.HE_TDATE} <= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
            dlpDateRange(1).Text = ToDate
        End If
        GoTo Cri_FTDatst
    End If
    If Len(dlpDateRange(1).Text) > 0 Then
        If IsDate(FromDate) Then
            dtYYY% = Year(FromDate)
            dtMM% = month(FromDate)
            dtDD% = Day(FromDate)
            TempCri = TempCri & " ({HRENTHRS.HE_FDATE} >= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) and "
            dlpDateRange(0).Text = FromDate
        End If
        TempCri = TempCri & "({HRENTHRS.HE_TDATE}  "
        TempCri = TempCri & " <= "
        dtYYY% = Year(dlpDateRange(1).Text)
        dtMM% = month(dlpDateRange(1).Text)
        dtDD% = Day(dlpDateRange(1).Text)
        TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        GoTo Cri_FTDatst
    End If
Else
'Hemu 06/03/2003 Begin - Commented. If Hourly Entitlements exists outside the
'                        HRPARCO.PC_FDATE and HRPARCO.PC_TDATE range then
'                        it won't be displayed in the report
'                        As Jerry suggested the Date range should not be defaulted.
'    If IsDate(FromDate) Then
'        dtYYY% = Year(FromDate)
'        dtMM% = Month(FromDate)
'        dtDD% = Day(FromDate)
'        TempCri = TempCri & " ({HRENTHRS.HE_FDATE} >= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) and "
'        dlpDateRange(0).Text = FromDate
'    End If
'    If IsDate(ToDate) Then
'        dtYYY% = Year(ToDate)
'        dtMM% = Month(ToDate)
'        dtDD% = Day(ToDate)
'        TempCri = TempCri & " ({HRENTHRS.HE_TDATE} <= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
'        dlpDateRange(1).Text = ToDate
'    End If
'Hemu 06/03/2003 End

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

Private Sub Cri_Attend()
Dim EECri As String, OneSet%, X%
Dim xHrEntCodes As String
Dim xFlex As String
Dim xFlexCode As String
Dim xCodes As String
Dim xfilteredCodes As String

OneSet% = False
For X% = 6 To 9
    If Len(clpCode(X%).Text) > 0 Then
        OneSet% = OneSet% + 1
    End If
Next X%

xfilteredCodes = ""
If OneSet% = 0 Then
    If chkAtt Then
        'Filter the attendance table records with only the codes that are hourly entitlement codes. This will help
        'in the performance when generating the report esp. when the Attendance table is large
        xfilteredCodes = Filter_AttendanceBy_HourlyEntitleCodes
        
        If Len(xfilteredCodes) > 0 Then
            EECri = xfilteredCodes
        End If
        
        GoTo combine_SelCri
    Else
        Exit Sub
    End If
End If

xFlex = ""
xFlexCode = ""
For X% = 6 To 9
    If Len(clpCode(X%).Text) > 0 Then
        EECri = EECri & "'" & clpCode(X%).Text & "' "
        
        If Right(clpCode(X%).Text, 1) = "+" Then
            If Len(xFlex) > 0 Then
                xFlex = xFlex & ","
            End If

            xFlexCode = Left(clpCode(X%).Text, Len(clpCode(X%).Text) - 1)
            xFlex = xFlex & "'" & xFlexCode & "-'"
        End If
        
        OneSet% = OneSet% - 1
        If OneSet% > 0 Then
            EECri = EECri & ","
        End If
    End If
Next X%

xCodes = EECri
EECri = "{HRENTHRS.HE_TYPE} IN [" & EECri & "]"

'If chkAtt Then
'    EECri = EECri & " AND ({HR_ATTENDANCE.AD_REASON} IN [" & xCodes & "]"
'    If Len(xFlex) > 0 Then
'        EECri = EECri & " OR {HR_ATTENDANCE.AD_REASON} startswith [" & xFlex & "])"
'    Else
'        EECri = EECri & ")"
'    End If
'End If

combine_SelCri:

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub '

Private Sub comGrpLoad()
comGroup(0).AddItem lStr("Division")
comGroup(0).AddItem lStr("Department")
comGroup(0).AddItem lStr("Location")
comGroup(0).AddItem lStr("Administered By")
comGroup(0).AddItem "Employee Name"
comGroup(0).AddItem lStr("G/L")
comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
If glbLinamar Then ' Frank May 2,2001
    comGroup(0).AddItem "Employment Type"
    comGroup(0).AddItem ("Home Line")
End If
If Not glbMulti Then comGroup(0).AddItem "Shift"
comGroup(0).AddItem lStr("Region")
comGroup(0).AddItem lStr("Rept. Authority 1")
comGroup(0).AddItem "(none)"
comGroup(0).ListIndex = 0
comGroup(1).AddItem "Employee Name"
comGroup(1).ListIndex = 0
comGroup(1).Enabled = False

cmbHours.AddItem "Hours"
cmbHours.AddItem "Days"
cmbHours.ListIndex = 0
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

Private Function Filter_AttendanceBy_HourlyEntitleCodes()
    Dim rsHrlyEnt As New ADODB.Recordset
    Dim xHrlyCodes As String
    Dim xFlexCode As String
    Dim xFlex As String
    Dim SQLQ As String
    Dim xRtrnStr As String
    
    'Initialize
    xHrlyCodes = ""
    xFlexCode = ""
    xFlex = ""
    xRtrnStr = ""
    
    'Get list of all the hourly entitlement codes used for the employees
    SQLQ = "SELECT DISTINCT(HE_TYPE) FROM HRENTHRS"
    rsHrlyEnt.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If Not rsHrlyEnt.EOF Then
        Do While Not rsHrlyEnt.EOF
            'Any hourly entitlement code
            If Len(xHrlyCodes) > 0 Then
                xHrlyCodes = xHrlyCodes & ","
            End If

            xHrlyCodes = xHrlyCodes & "'" & rsHrlyEnt("HE_TYPE") & "' "

            'Flex hourly entitlement code
            If Right(rsHrlyEnt("HE_TYPE"), 1) = "+" Then
                If Len(xFlex) > 0 Then
                    xFlex = xFlex & ","
                End If

                xFlexCode = Left(rsHrlyEnt("HE_TYPE"), Len(rsHrlyEnt("HE_TYPE")) - 1)
                xFlex = xFlex & "'" & xFlexCode & "-'"
            End If
            
            rsHrlyEnt.MoveNext
        Loop
    End If
    rsHrlyEnt.Close
    Set rsHrlyEnt = Nothing
    
    'Build the return string
    If Len(xHrlyCodes) > 0 Then
        xRtrnStr = "({HR_ATTENDANCE.AD_REASON} IN [" & xHrlyCodes & "]"
    End If
    
    If Len(xFlex) > 0 Then
        xRtrnStr = xRtrnStr & " OR {HR_ATTENDANCE.AD_REASON} startswith [" & xFlex & "])"
    Else
        xRtrnStr = xRtrnStr & ")"
    End If
    
    Filter_AttendanceBy_HourlyEntitleCodes = xRtrnStr
    
End Function

