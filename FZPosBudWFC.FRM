VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmRPosBudWFC 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Budgeted Position Report"
   ClientHeight    =   8205
   ClientLeft      =   375
   ClientTop       =   915
   ClientWidth     =   11475
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
   ScaleHeight     =   8205
   ScaleWidth      =   11475
   WindowState     =   2  'Maximized
   Begin VB.Frame frmTabExp 
      Caption         =   "TableExps"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4440
      TabIndex        =   43
      Top             =   6480
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CheckBox chkIncJobPos 
         Caption         =   "Include Inactive Job/Position Codes"
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
         Left            =   1830
         TabIndex        =   47
         Top             =   600
         Width           =   3255
      End
      Begin VB.OptionButton optExpBud 
         Caption         =   "Budgeted Positions"
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
         Left            =   3720
         TabIndex        =   46
         Tag             =   "Male vs Female"
         Top             =   240
         Width           =   1755
      End
      Begin VB.OptionButton optExpPos 
         Caption         =   "Position Master"
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
         TabIndex        =   45
         Tag             =   "Male vs Female"
         Top             =   240
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.OptionButton optExpJob 
         Caption         =   "Job Master"
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
         Left            =   0
         TabIndex        =   44
         Tag             =   "Male vs Female"
         Top             =   240
         Width           =   1635
      End
   End
   Begin VB.CheckBox chkShowEmp 
      Caption         =   "Show Employees"
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
      TabIndex        =   42
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Frame frmOpt2 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1800
      TabIndex        =   38
      Top             =   5760
      Width           =   7935
      Begin VB.OptionButton optAll 
         Caption         =   "All"
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
         Left            =   0
         TabIndex        =   41
         Tag             =   "Male vs Female"
         Top             =   0
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.OptionButton optOverBud 
         Caption         =   "Over Budget"
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
         Left            =   4920
         TabIndex        =   40
         Tag             =   "Marital Status"
         Top             =   0
         Width           =   2445
      End
      Begin VB.OptionButton optVacancy 
         Caption         =   "Vacancy"
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
         Left            =   2160
         TabIndex        =   39
         Tag             =   "Male vs Female"
         Top             =   0
         Width           =   1755
      End
   End
   Begin VB.Frame frmOpt1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1800
      TabIndex        =   34
      Top             =   5400
      Width           =   7815
      Begin VB.OptionButton optUseBud 
         Caption         =   "Use Budgeted # of Positions"
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
         Left            =   2160
         TabIndex        =   37
         Tag             =   "Marital Status"
         Top             =   0
         Width           =   2445
      End
      Begin VB.OptionButton optUseAll 
         Caption         =   "Use all Positions"
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
         Left            =   0
         TabIndex        =   36
         Tag             =   "Male vs Female"
         Top             =   0
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.OptionButton optUseFTE 
         Caption         =   "Use FTE #"
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
         Left            =   4920
         TabIndex        =   35
         Tag             =   "Average Employee Age"
         Top             =   0
         Width           =   2025
      End
   End
   Begin VB.TextBox txtYear 
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
      DataField       =   "JG_YEAR"
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
      Left            =   1880
      MaxLength       =   4
      TabIndex        =   13
      Tag             =   "01-Number of positions that exist for this job"
      Top             =   4920
      Width           =   855
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   510
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
      TabIndex        =   9
      Tag             =   "00-Enter Section Code"
      Top             =   3480
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDSE"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   1560
      TabIndex        =   8
      Tag             =   "00-Enter Administered By Code"
      Top             =   3150
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
      Top             =   2820
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
      Top             =   2160
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
      Top             =   1830
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
      Top             =   1500
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
      Top             =   1170
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
      Top             =   840
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
      TabIndex        =   16
      Tag             =   "Second level of grouping records"
      Top             =   7200
      Visible         =   0   'False
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
      TabIndex        =   15
      Tag             =   "First Level of grouping records"
      Top             =   6885
      Width           =   2325
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Tag             =   "10-Enter Employee Number"
      Top             =   2490
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   8280
      Top             =   7320
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
   Begin INFOHR_Controls.CodeLookup clpJob 
      Height          =   285
      Left            =   1560
      TabIndex        =   10
      Tag             =   "00-Enter Position Code"
      Top             =   3840
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   6
      Left            =   1560
      TabIndex        =   11
      Tag             =   "00-Position Group  Code"
      Top             =   4200
      Width           =   4080
      _ExtentX        =   7197
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "JBGC"
   End
   Begin INFOHR_Controls.CodeLookup clpJobMaster 
      Height          =   285
      Left            =   1560
      TabIndex        =   12
      Tag             =   "01-Job code"
      Top             =   4560
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   25
      LookupType      =   13
   End
   Begin VB.Label lblJob 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Job Code"
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
      Top             =   4600
      Width           =   675
   End
   Begin VB.Label lblYear 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
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
      Top             =   4920
      Width           =   330
   End
   Begin VB.Label lblPosGroup 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Job/Position Group"
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
      Top             =   4245
      Width           =   1365
   End
   Begin VB.Label lblPosition 
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   30
      Top             =   3885
      Width           =   975
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
      TabIndex        =   29
      Top             =   2205
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
      TabIndex        =   28
      Top             =   3525
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
      TabIndex        =   27
      Top             =   2865
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
      TabIndex        =   26
      Top             =   3195
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
      TabIndex        =   25
      Top             =   1215
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
      TabIndex        =   24
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
      TabIndex        =   23
      Top             =   7260
      Visible         =   0   'False
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
      TabIndex        =   22
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
      TabIndex        =   21
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
      TabIndex        =   20
      Top             =   2535
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
      TabIndex        =   19
      Top             =   1875
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
      TabIndex        =   18
      Top             =   1545
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
      TabIndex        =   17
      Top             =   885
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
      TabIndex        =   14
      Top             =   555
      Width           =   555
   End
End
Attribute VB_Name = "frmRPosBudWFC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xFunName As String

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim X%

On Error GoTo PrntErr

If xFunName = "TableExps" Then
    Call TableExpProgs
    Exit Sub
End If

If CriCheck() Then
    If Not PrtForm("Budgeted Position Report Criteria", Me) Then Exit Sub
    Call set_PrintState(False)
    X% = Cri_SetAll()
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
Dim X%
Dim strWHand As String
On Error GoTo CRW_Err

If xFunName = "TableExps" Then
    Call TableExpProgs
    Exit Sub
End If

If CriCheck() Then
      Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    Screen.MousePointer = HOURGLASS
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    X% = Cri_SetAll()
    'Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    'Me.vbxCrystal.Action = 1
    'vbxCrystal.Reset
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

Private Sub comGroup_GotFocus(Index As Integer)
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")
    comGroup(0).AddItem lStr("Section")
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem "Job"
    comGroup(0).AddItem "Position Group"
    comGroup(0).AddItem "Position"
    comGroup(0).AddItem "(none)"
    comGroup(0).ListIndex = 0
    'comGroup(1).AddItem "Employee Name"
    'comGroup(1).AddItem "Employee Number"
    'comGroup(1).ListIndex = 0

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
Dim X%, strRName$

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
For X% = 0 To 5
    Call Cri_Code(X%)
Next X%
Call Cri_PT
Call Cri_EE

If chkShowEmp.Value Then
    Call WFC_BudPos_ReportShowEmp
Else
    Call WFC_BudPos_ReportShowPos
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
Dim X%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

'for labels - sort by name always
'imbeded in report

Cri_Sorts = 0
'first set primary grouping
Y% = 0
grpField$ = getEGroup(comGroup(0).Text)

Call setRptLabel(Me, 0)

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
    
Y% = X% + 1
dscGroup$ = comGroup(X%).Text
dscGroup$ = "descGroup" & CStr(Y%) & "= '" & dscGroup$ & "'"
Me.vbxCrystal.Formulas(X%) = dscGroup$

grpCond$ = "GROUP" & CStr(Y%) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(X%) = grpCond$

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

Private Function CriChkExpJob()
Dim X%

CriChkExpJob = False

If Not clpDiv.ListChecker Then
    Exit Function
End If

If Not clpDept.ListChecker Then
    Exit Function
End If


For X% = 0 To 6
    If Not clpCode(X).ListChecker Then Exit Function
Next X%

CriChkExpJob = True

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
'If Len(clpPT) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

For X% = 0 To 5
    If Not clpCode(X).ListChecker Then Exit Function
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
glbOnTop = "frmRPosBudWFC"

Screen.MousePointer = HOURGLASS
frmRPosBudWFC.Caption = glbFrmCaption$ ' "Budgeted Position Report"

If glbFrmCaption$ = "Budgeted Position Report" Then xFunName = "BudPosRpt"
If glbFrmCaption$ = "Position Control Table Exports" Then xFunName = "TableExps"

Call comGrpLoad
Call setRptCaption(Me)
'If glbLinamar Then clpCode(3).MaxLength = 8
'If glbCompSerial = "S/N - 2227W" Then clpCode(3).MaxLength = 6
'If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6

If xFunName = "TableExps" Then
    Call ScreenSetTableExps
End If

Call PosExpScreenSetup

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

Private Sub WFC_BudPos_ReportShowPos()
    Dim rsEmp As New ADODB.Recordset
    Dim rsWrk As New ADODB.Recordset
    Dim rsDiv As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim snapBudget As New ADODB.Recordset
    Dim rsJOB As New ADODB.Recordset
    Dim rsJOB2 As New ADODB.Recordset
    Dim exApp As Object, exBook As Object, exSheet As Object
    Dim SQLQ, sSQLQ As String
    Dim locSQL As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim xCol As Long
    Dim xRptSele As String
    Dim I, totNum
    Dim xOrderBy
    Dim xStartLine
    Dim xGroup1Desc
    Dim xTotBUDGNBR, xTotNBRFIL, xTotFTENUM, xTotFTENUMFILL
    Dim xSubBUDGNBR, xSubNBRFIL, xSubFTENUM, xSubFTENUMFILL
    Dim xPosBUDGNBR, xPosNBRFIL, xPosFTENUM, xPosFTENUMFILL
    Dim tmpGroupDesc, tmpPos
    Dim xCurPosCode
    Dim xDisFlag As Boolean
    
    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "WFC Budgeted Position Report Tmp.xls"
    xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "WFC Budgeted Position Report(" & Trim(glbUserID) & ").xls"

    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    FileCopy xlsFileTmp, xlsFileMat

    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(0).FloodPercent = 0
    Screen.MousePointer = HOURGLASS
        
    SQLQ = "DELETE FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE HRJOBBUD SET JG_BUDGNBR = 0 WHERE JG_BUDGNBR IS NULL"
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRJOBBUD SET JG_NBRFIL = 0 WHERE JG_NBRFIL IS NULL"
    gdbAdoIhr001.Execute SQLQ
    
    sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
    'SQLQ = "SELECT HR_JOB_HISTORY.JH_EMPNBR, HRJOB.JB_CODE, HRJOB.JB_DESCR"
    SQLQ = "SELECT ED_SURNAME, ED_FNAME, ED_DIV,ED_DEPTNO,ED_REGION,ED_SECTION,ED_LOC,ED_ORG,JB_ORG,ED_GLNO,ED_COUNTRY, HR_JOB_HISTORY.*, HRJOB.JB_CODE, HRJOB.JB_DESCR,JB_GRPCD,JB_JOBCODE,JB_BAND "
    SQLQ = SQLQ & "FROM HR_JOB_HISTORY INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
    SQLQ = SQLQ & "LEFT JOIN HREMP ON HR_JOB_HISTORY.JH_EMPNBR = HREMP.ED_EMPNBR "
    SQLQ = SQLQ & "WHERE HR_JOB_HISTORY.JH_CURRENT = 1 "
    SQLQ = SQLQ & "AND " & sSQLQ
    If Len(clpJob.Text) > 0 Then
        SQLQ = SQLQ & " AND (HR_JOB_HISTORY.JH_JOB in ('" & Replace(clpJob.Text, ",", "','") & "')) "
    End If
    If Len(clpCode(6).Text) > 0 Then 'Position Group
        SQLQ = SQLQ & " AND (HRJOB.JB_GRPCD = '" & clpCode(6).Text & "') "
    End If
    If Len(clpJobMaster.Text) > 0 Then
        SQLQ = SQLQ & " AND (HRJOB.JB_JOBCODE = '" & clpJobMaster.Text & "') "
    End If
    If optVacancy.Value Or optOverBud.Value Then
        If optVacancy.Value Then
            SQLQ = SQLQ & " AND (HR_JOB_HISTORY.JH_JOB IN (SELECT JG_CODE FROM HRJOBBUD WHERE (JG_BUDGNBR - JG_NBRFIL) > 0 )) "
        End If
        If optOverBud.Value Then
            SQLQ = SQLQ & " AND (HR_JOB_HISTORY.JH_JOB IN (SELECT JG_CODE FROM HRJOBBUD WHERE (JG_BUDGNBR - JG_NBRFIL) < 0 )) "
        End If
    Else
        SQLQ = SQLQ & " AND (HR_JOB_HISTORY.JH_JOB IN (SELECT JG_CODE FROM HRJOBBUD)) "
    End If
    
    If rsEmp.State <> 0 Then rsEmp.Close
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        totNum = rsEmp.RecordCount
    End If
    I = 0
    Do While Not rsEmp.EOF
        If (I / totNum) <= 1 Then
            MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
            I = I + 1
        End If
        DoEvents

        'SQLQ = "SELECT TT_WRKEMP, TT_PT,TT_EMP,TT_LANG1,TT_WHRS,TT_DHRS, TT_SEDATE FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
        SQLQ = "SELECT * FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
        'SQLQ = SQLQ & "AND TT_EMPNBR = '" & rsEmp("JH_EMPNBR") & "' " '
        SQLQ = SQLQ & "AND TT_JREAS = '" & rsEmp("JH_JOB") & "' " 'position based
        If rsWrk.State <> 0 Then rsWrk.Close
        rsWrk.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        
        If Not rsWrk.EOF Then
            GoTo Next_Emp
        End If
        
        If rsWrk.EOF Then
            rsWrk.AddNew
            rsWrk("TT_WRKEMP") = glbUserID
            rsWrk("TT_EMPNBR") = rsEmp("JH_EMPNBR")
        End If
        rsWrk("TT_CHAR10") = rsEmp("ED_COUNTRY") '
        rsWrk("TT_LANG1") = rsEmp("ED_REGION")
        If Not IsNull(rsEmp("ED_SECTION")) Then
            rsWrk("TT_SHIFT") = Left(rsEmp("ED_SECTION"), 4)
        End If
        'If Not IsNull(rsEmp("ED_LOC")) Then
        '    rsWrk("TT_OLDDIV") = Left(rsEmp("ED_LOC"), 4) 'LOCATION
        'End If
        'If Not IsNull(rsEmp("ED_ORG")) Then
        '    rsWrk("TT_ORG") = Left(rsEmp("ED_ORG"), 4)
        'End If '
        If Not IsNull(rsEmp("JB_ORG")) Then
            rsWrk("TT_ORG") = Left(rsEmp("JB_ORG"), 4)
        End If
        rsWrk("TT_JREAS") = rsEmp("JH_JOB") 'employee positin code
        
        If Not IsNull(rsEmp("ED_DIV")) Then
            rsWrk("TT_NEWDIV") = Left(rsEmp("ED_DIV"), 4)
            SQLQ = "SELECT * FROM HR_DIVISION WHERE DIV = '" & rsEmp("ED_DIV") & "' "
            If rsDiv.State <> 0 Then rsDiv.Close
            rsDiv.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsDiv.EOF Then
                If Not IsNull(rsDiv("DV_LOC")) Then
                    rsWrk("TT_OLDDIV") = Left(rsDiv("DV_LOC"), 4) 'LOCATION
                End If
            End If
        End If
        rsWrk("TT_LANG2") = rsEmp("ED_GLNO")
        rsWrk("TT_NAMEFLD") = Left(getDeptDescPub(rsEmp("ED_DEPTNO")), 40) '
        If Not IsNull(rsEmp("JB_BAND")) Then
            rsWrk("TT_BRELATE") = Left(rsEmp("JB_BAND"), 6)
        End If
        If Not IsNull(rsEmp("JB_JOBCODE")) Then
            rsWrk("TT_OLDJOB") = Left(rsEmp("JB_JOBCODE"), 6)
        End If
        rsWrk("TT_JOB") = Left(getNewJobCodeDescPub(rsEmp("JB_JOBCODE")), 50) 'Job Desc

        If Not IsNull(rsEmp("JB_DESCR")) Then '
            'rsWrk("TT_COURSE") = Left(getPosDesc(rsEmp("JH_JOB")), 60) 'Position desc 'JB_DESCR
            rsWrk("TT_COURSE") = Left(rsEmp("JB_DESCR"), 60) 'Position desc 'JB_DESCR
        End If
        rsWrk("TT_BNAME") = Left(GetTABLCodePub("JBGC", rsEmp("JB_GRPCD")), 40) 'Pos Group desc
        rsWrk("TT_SALR1") = Left(rsEmp("ED_SURNAME"), 30)
        rsWrk("TT_SALR2") = Left(rsEmp("ED_FNAME"), 30)
        'find Reports to Position - begin
        SQLQ = "SELECT JB_CODE, JB_REPTAU FROM HRJOB WHERE JB_CODE = '" & rsEmp("JH_JOB") & "' "
        If rsJOB.State <> 0 Then rsJOB.Close
        rsJOB.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsJOB.EOF Then
            If Not IsNull(rsJOB("JB_REPTAU")) Then
                rsWrk("TT_DTYPE") = Left(rsJOB("JB_REPTAU"), 50)
                SQLQ = "SELECT JB_CODE, JB_REPTAU,JB_DESCR FROM HRJOB WHERE JB_CODE = '" & rsJOB("JB_REPTAU") & "' "
                If rsJOB2.State <> 0 Then rsJOB2.Close
                rsJOB2.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsJOB2.EOF Then
                    rsWrk("TT_ETYPE") = Left(rsJOB2("JB_DESCR"), 50)
                End If
            End If
        End If
        rsJOB.Close
        'find Reports to Position - end
        
        'populate Budgeted Position - begin
        rsWrk("TT_CCOST") = 0 'JG_BUDGNBR
        rsWrk("TT_ECOST") = 0 'JG_NBRFIL
        rsWrk("TT_PCE") = 0 'JG_FTENUM
        rsWrk("TT_PCC") = 0 'JG_FTENUMFILL
        If Not IsNull(rsEmp("ED_DIV")) Then
            SQLQ = "SELECT * FROM HRJOBBUD WHERE (1=1) "
            'SQLQ = SQLQ & "AND JG_CURRENT = 1 "
            SQLQ = SQLQ & "AND JG_CODE = '" & rsEmp("JH_JOB") & "' "
            SQLQ = SQLQ & "AND JG_DIV = '" & rsEmp("ED_DIV") & "' "
            SQLQ = SQLQ & "ORDER BY JG_FRDATE DESC " 'Ticket #28341 Franks 03/21/2016
            If snapBudget.State <> 0 Then snapBudget.Close
            snapBudget.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not snapBudget.EOF Then '
                If IsNull(snapBudget("JG_BUDGNBR")) Then
                    rsWrk("TT_CCOST") = 0
                Else
                    rsWrk("TT_CCOST") = snapBudget("JG_BUDGNBR")
                End If
                If IsNull(snapBudget("JG_NBRFIL")) Then
                    rsWrk("TT_ECOST") = 0
                Else
                    rsWrk("TT_ECOST") = snapBudget("JG_NBRFIL")
                End If
                If IsNull(snapBudget("JG_FTENUM")) Then
                    rsWrk("TT_PCE") = 0
                Else
                    rsWrk("TT_PCE") = snapBudget("JG_FTENUM")
                End If
                If IsNull(snapBudget("JG_FTENUMFILL")) Then
                    rsWrk("TT_PCC") = 0
                Else
                    rsWrk("TT_PCC") = snapBudget("JG_FTENUMFILL")
                End If
            End If
            snapBudget.Close
        End If
        'populate Budgeted Position - end
        rsWrk.Update
        
        rsWrk("TT_SKLCOM") = getGroup1Desc(rsEmp, rsWrk) '50 Group 1 Desc
        rsWrk.Update
        
Next_Emp:
        rsEmp.MoveNext
    Loop
    rsEmp.Close
    MDIMain.panHelp(0).FloodPercent = 100
    
    'If optVacancy.Value Then 'Vacancy
    '    SQLQ = "DELETE FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
    '    SQLQ = SQLQ & "AND (TT_CCOST - TT_ECOST) > 0 "
    '    gdbAdoIhr001.Execute SQLQ
    'End If
    'If optOverBud.Value Then 'optOverBud
    '    SQLQ = "DELETE FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
    '    SQLQ = SQLQ & "AND (TT_CCOST - TT_ECOST) < 0 "
    '    gdbAdoIhr001.Execute SQLQ
    'End If
    
    'Populate Excel file - begin
    'Create new WorkBook of Excel
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(xlsFileMat)
    Set exSheet = exBook.Worksheets(1)
    
    ''exSheet.Cells(2, 2) = "Status of WPS Training for " & Year(Date)
    exSheet.Cells(1, 2) = "" & Format(Date, "MMM dd, YYYY") & ""
    exSheet.Cells(2, 2) = "" & Time$ & ""

    'First line of data
    xStartLine = 5
    
    SQLQ = "SELECT * FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
    xOrderBy = getExcelOrderBy
    SQLQ = SQLQ & "ORDER BY " & xOrderBy & " "
    If rsWrk.State <> 0 Then rsWrk.Close
    rsWrk.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsWrk.EOF Then
        totNum = rsWrk.RecordCount: I = 0
    End If
            
    xRow = xStartLine
    'xTotStaff = 0
    'xTotPlant = 0
    xTotBUDGNBR = 0: xTotNBRFIL = 0: xTotFTENUM = 0: xTotFTENUMFILL = 0
    xSubBUDGNBR = 0: xSubNBRFIL = 0: xSubFTENUM = 0: xSubFTENUMFILL = 0
    xPosBUDGNBR = 0: xPosNBRFIL = 0: xPosFTENUM = 0: xPosFTENUMFILL = 0

    If Not rsWrk.EOF Then
        tmpGroupDesc = rsWrk("TT_SKLCOM")
        xCurPosCode = rsWrk("TT_JREAS")
    End If
    xDisFlag = True
    Do While Not rsWrk.EOF
        If (I / totNum) <= 1 Then
            MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
            I = I + 1
        End If
        DoEvents
        
        
        'sub total
        If Not (tmpGroupDesc = rsWrk("TT_SKLCOM")) Then
            exSheet.Cells(xRow, 7 + 1).Font.Bold = True
            exSheet.Cells(xRow, 11 + 2 + 1).Font.Bold = True
            exSheet.Cells(xRow, 12 + 2 + 1).Font.Bold = True
            exSheet.Cells(xRow, 13 + 2 + 1).Font.Bold = True
            exSheet.Cells(xRow, 14 + 2 + 1).Font.Bold = True
            exSheet.Cells(xRow, 15 + 2 + 1).Font.Bold = True
            exSheet.Cells(xRow, 16 + 2 + 1).Font.Bold = True
            
            exSheet.Cells(xRow, 7 + 1) = "Sub Total of " & tmpGroupDesc
            exSheet.Cells(xRow, 11 + 2 + 1) = xSubBUDGNBR
            exSheet.Cells(xRow, 12 + 2 + 1) = xSubNBRFIL
            exSheet.Cells(xRow, 13 + 2 + 1) = xSubBUDGNBR - xSubNBRFIL
            exSheet.Cells(xRow, 14 + 2 + 1) = xSubFTENUM
            exSheet.Cells(xRow, 15 + 2 + 1) = xSubFTENUMFILL
            exSheet.Cells(xRow, 16 + 2 + 1) = xSubFTENUM - xSubFTENUMFILL
            xRow = xRow + 2
            xSubBUDGNBR = 0: xSubNBRFIL = 0: xSubFTENUM = 0: xSubFTENUMFILL = 0
            tmpGroupDesc = rsWrk("TT_SKLCOM")
        End If
        ''xSubBUDGNBR = xSubBUDGNBR + rsWrk("TT_CCOST")
        ''xSubNBRFIL = xSubNBRFIL + rsWrk("TT_ECOST")
        ''xSubFTENUM = xSubFTENUM + rsWrk("TT_PCE")
        ''xSubFTENUMFILL = xSubFTENUMFILL + rsWrk("TT_PCC")
        ''
        ''
        '''groud total
        ''xTotBUDGNBR = xTotBUDGNBR + rsWrk("TT_CCOST")
        ''xTotNBRFIL = xTotNBRFIL + rsWrk("TT_ECOST")
        ''xTotFTENUM = xTotFTENUM + rsWrk("TT_PCE")
        ''xTotFTENUMFILL = xTotFTENUMFILL + rsWrk("TT_PCC")
                
        
        'Plant/Division
        exSheet.Cells(xRow, 1) = rsWrk("TT_CHAR10") 'COUNTRY
        exSheet.Cells(xRow, 2) = rsWrk("TT_LANG1") 'REGION
        exSheet.Cells(xRow, 3) = rsWrk("TT_SHIFT") 'SECTION
        exSheet.Cells(xRow, 4) = rsWrk("TT_OLDDIV") 'LOC
        exSheet.Cells(xRow, 5) = rsWrk("TT_ORG") ' UNION
        'exSheet.Cells(xRow, 6) = rsWrk("TT_LANG2") ' GL - not employee based
        'exSheet.Cells(xRow, 7) = rsWrk("TT_NAMEFLD") ' DEPT - not employee based
        exSheet.Cells(xRow, 6) = rsWrk("TT_BRELATE")
        exSheet.Cells(xRow, 6 + 1) = rsWrk("TT_OLDJOB") ' JOB
        exSheet.Cells(xRow, 7 + 1) = rsWrk("TT_JOB") ' JOB Desc
        exSheet.Cells(xRow, 8 + 1) = rsWrk("TT_JREAS") ' POS
        exSheet.Cells(xRow, 9 + 1) = rsWrk("TT_COURSE") ' POS Desc
        exSheet.Cells(xRow, 10 + 1) = rsWrk("TT_BNAME") ' POS GROUP
        exSheet.Cells(xRow, 11 + 1) = rsWrk("TT_DTYPE") '
        exSheet.Cells(xRow, 12 + 1) = rsWrk("TT_ETYPE") '
        
        If Not xCurPosCode = rsWrk("TT_JREAS") Then
            xCurPosCode = rsWrk("TT_JREAS")
            xDisFlag = True
        End If
        If xDisFlag Then
            exSheet.Cells(xRow, 11 + 2 + 1) = rsWrk("TT_CCOST")
            exSheet.Cells(xRow, 12 + 2 + 1) = rsWrk("TT_ECOST")
            exSheet.Cells(xRow, 13 + 2 + 1) = rsWrk("TT_CCOST") - rsWrk("TT_ECOST")
            exSheet.Cells(xRow, 14 + 2 + 1) = rsWrk("TT_PCE")
            exSheet.Cells(xRow, 15 + 2 + 1) = rsWrk("TT_PCC")
            exSheet.Cells(xRow, 16 + 2 + 1) = rsWrk("TT_PCE") - rsWrk("TT_PCC")
            
            'This position amount
            xPosBUDGNBR = rsWrk("TT_CCOST")
            xPosNBRFIL = rsWrk("TT_ECOST")
            xPosFTENUM = rsWrk("TT_PCE")
            xPosFTENUMFILL = rsWrk("TT_PCC")
        
            xSubBUDGNBR = xSubBUDGNBR + rsWrk("TT_CCOST")
            xSubNBRFIL = xSubNBRFIL + rsWrk("TT_ECOST")
            xSubFTENUM = xSubFTENUM + rsWrk("TT_PCE")
            xSubFTENUMFILL = xSubFTENUMFILL + rsWrk("TT_PCC")
            
            
            'groud total
            xTotBUDGNBR = xTotBUDGNBR + rsWrk("TT_CCOST")
            xTotNBRFIL = xTotNBRFIL + rsWrk("TT_ECOST")
            xTotFTENUM = xTotFTENUM + rsWrk("TT_PCE")
            xTotFTENUMFILL = xTotFTENUMFILL + rsWrk("TT_PCC")
        
            xDisFlag = False
        End If

        xRow = xRow + 1
        rsWrk.MoveNext
    Loop
    
    'Total ----
    If xRow > xStartLine Then
            If comGroup(0).Text = "(none)" Then
            Else
            'xRow = xRow + 1
            exSheet.Cells(xRow, 7 + 1).Font.Bold = True
            exSheet.Cells(xRow, 11 + 2 + 1).Font.Bold = True
            exSheet.Cells(xRow, 12 + 2 + 1).Font.Bold = True
            exSheet.Cells(xRow, 13 + 2 + 1).Font.Bold = True
            exSheet.Cells(xRow, 14 + 2 + 1).Font.Bold = True
            exSheet.Cells(xRow, 15 + 2 + 1).Font.Bold = True
            exSheet.Cells(xRow, 26 + 2 + 1).Font.Bold = True
            exSheet.Cells(xRow, 7 + 1) = "Sub Total of " & tmpGroupDesc
            exSheet.Cells(xRow, 11 + 2 + 1) = xSubBUDGNBR
            exSheet.Cells(xRow, 12 + 2 + 1) = xSubNBRFIL
            exSheet.Cells(xRow, 13 + 2 + 1) = xSubBUDGNBR - xSubNBRFIL
            exSheet.Cells(xRow, 14 + 2 + 1) = xSubFTENUM
            exSheet.Cells(xRow, 15 + 2 + 1) = xSubFTENUMFILL
            exSheet.Cells(xRow, 16 + 2 + 1) = xSubFTENUM - xSubFTENUMFILL
            xRow = xRow + 1
            End If
            
            xRow = xRow + 1
            exSheet.Cells(xRow, 7 + 1).Font.Bold = True
            exSheet.Cells(xRow, 11 + 2 + 1).Font.Bold = True
            exSheet.Cells(xRow, 12 + 2 + 1).Font.Bold = True
            exSheet.Cells(xRow, 13 + 2 + 1).Font.Bold = True
            exSheet.Cells(xRow, 14 + 2 + 1).Font.Bold = True
            exSheet.Cells(xRow, 15 + 2 + 1).Font.Bold = True
            exSheet.Cells(xRow, 16 + 2 + 1).Font.Bold = True
            exSheet.Cells(xRow, 7 + 1) = "Grand Total "
            exSheet.Cells(xRow, 11 + 2 + 1) = xTotBUDGNBR
            exSheet.Cells(xRow, 12 + 2 + 1) = xTotNBRFIL
            exSheet.Cells(xRow, 13 + 2 + 1) = xTotBUDGNBR - xTotNBRFIL
            exSheet.Cells(xRow, 14 + 2 + 1) = xTotFTENUM
            exSheet.Cells(xRow, 15 + 2 + 1) = xTotFTENUMFILL
            exSheet.Cells(xRow, 16 + 2 + 1) = xTotFTENUM - xTotFTENUMFILL
            xRow = xRow + 1
            
            'exSheet.Cells(xRow, 1) = "TOTAL RECORDED"
            'exSheet.Cells(xRow, 2) = xTotStaff
            'exSheet.Cells(xRow, 3) = xTotPlant
            'xRow = xRow + 1
            'exSheet.Cells(xRow, 1) = "% trained"
    End If
    'rsWrk.Close
    
    exBook.Save
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing

    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = " "
    Screen.MousePointer = DEFAULT

    Call Pause(1)
    
    'If Not comGroup(0).Text = "(none)" Then
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
    'End If
    

End Sub

Private Sub WFC_BudPos_ReportShowEmp() 'WFC_WPS_Training_Report
    Dim rsEmp As New ADODB.Recordset
    Dim rsWrk As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim rsPos As New ADODB.Recordset
    Dim rsRptPos As New ADODB.Recordset
    Dim snapBudget As New ADODB.Recordset
    Dim exApp As Object, exBook As Object, exSheet As Object
    Dim SQLQ, sSQLQ As String
    Dim locSQL As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim xCol As Long
    Dim xRptSele As String
    Dim I, totNum
    Dim xOrderBy
    Dim xStartLine
    Dim xGroup1Desc
    Dim xTotBUDGNBR, xTotNBRFIL, xTotFTENUM, xTotFTENUMFILL
    Dim xSubBUDGNBR, xSubNBRFIL, xSubFTENUM, xSubFTENUMFILL
    Dim xPosBUDGNBR, xPosNBRFIL, xPosFTENUM, xPosFTENUMFILL
    Dim tmpGroupDesc, tmpPos
    Dim xCurPosCode
    Dim xDisFlag As Boolean
    Dim xStr
    Dim xFTENo, xFTEFilled, xCnt, xReptNo1, xReptNo2
    Dim xReptPos1, xReptPos2, xReptPos1Desc, xReptPos2Desc
    
    On Error GoTo ReportShowEmp_Err
    
    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "WFC Budgeted Position ReportEmp Tmp.xls"
    xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "WFC Budgeted Position ReportEmp(" & Trim(glbUserID) & ").xls"

    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    FileCopy xlsFileTmp, xlsFileMat

    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(0).FloodPercent = 0
    Screen.MousePointer = HOURGLASS
        
    SQLQ = "DELETE FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE HRJOBBUD SET JG_BUDGNBR = 0 WHERE JG_BUDGNBR IS NULL"
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRJOBBUD SET JG_NBRFIL = 0 WHERE JG_NBRFIL IS NULL"
    gdbAdoIhr001.Execute SQLQ
    
    sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
    'SQLQ = "SELECT HR_JOB_HISTORY.JH_EMPNBR, HRJOB.JB_CODE, HRJOB.JB_DESCR"
    SQLQ = "SELECT ED_SURNAME, ED_FNAME, ED_DIV,ED_DEPTNO,ED_REGION,ED_SECTION,ED_LOC,ED_ORG,ED_GLNO,ED_COUNTRY,ED_EMP, HR_JOB_HISTORY.*, HRJOB.JB_CODE, HRJOB.JB_DESCR,JB_GRPCD,JB_JOBCODE,JB_REPTAU,JB_BAND "
    SQLQ = SQLQ & "FROM HR_JOB_HISTORY INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
    SQLQ = SQLQ & "LEFT JOIN HREMP ON HR_JOB_HISTORY.JH_EMPNBR = HREMP.ED_EMPNBR "
    SQLQ = SQLQ & "WHERE HR_JOB_HISTORY.JH_CURRENT = 1 "
    SQLQ = SQLQ & "AND " & sSQLQ
    If Len(clpJob.Text) > 0 Then
        SQLQ = SQLQ & " AND (HR_JOB_HISTORY.JH_JOB in ('" & Replace(clpJob.Text, ",", "','") & "')) "
    End If
    If Len(clpCode(6).Text) > 0 Then 'Position Group
        SQLQ = SQLQ & " AND (HRJOB.JB_GRPCD = '" & clpCode(6).Text & "') "
    End If
    If Len(clpJobMaster.Text) > 0 Then
        SQLQ = SQLQ & " AND (HRJOB.JB_JOBCODE = '" & clpJobMaster.Text & "') "
    End If
    If optVacancy.Value Or optOverBud.Value Then
        If optVacancy.Value Then
            SQLQ = SQLQ & " AND (HR_JOB_HISTORY.JH_JOB IN (SELECT JG_CODE FROM HRJOBBUD WHERE (JG_BUDGNBR - JG_NBRFIL) > 0 )) "
        End If
        If optOverBud.Value Then
            SQLQ = SQLQ & " AND (HR_JOB_HISTORY.JH_JOB IN (SELECT JG_CODE FROM HRJOBBUD WHERE (JG_BUDGNBR - JG_NBRFIL) < 0 )) "
        End If
    Else
        SQLQ = SQLQ & " AND (HR_JOB_HISTORY.JH_JOB IN (SELECT JG_CODE FROM HRJOBBUD)) "
    End If
    
    If rsEmp.State <> 0 Then rsEmp.Close
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        totNum = rsEmp.RecordCount
    End If
    I = 0
    Do While Not rsEmp.EOF
        If (I / totNum) <= 1 Then
            MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
            I = I + 1
        End If
        DoEvents
        'If rsEmp("JH_EMPNBR") = 10601162 Then
        'Debug.Print ""
        'End If

        'SQLQ = "SELECT TT_WRKEMP, TT_PT,TT_EMP,TT_LANG1,TT_WHRS,TT_DHRS, TT_SEDATE FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
        SQLQ = "SELECT * FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
        SQLQ = SQLQ & "AND TT_EMPNBR = '" & rsEmp("JH_EMPNBR") & "' "
        If rsWrk.State <> 0 Then rsWrk.Close
        rsWrk.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsWrk.EOF Then
            rsWrk.AddNew
            rsWrk("TT_WRKEMP") = glbUserID
            rsWrk("TT_EMPNBR") = rsEmp("JH_EMPNBR")
        End If
        rsWrk("TT_CHAR10") = rsEmp("ED_COUNTRY") '
        rsWrk("TT_LANG1") = rsEmp("ED_REGION")
        If Not IsNull(rsEmp("ED_SECTION")) Then
            rsWrk("TT_SHIFT") = Left(rsEmp("ED_SECTION"), 4)
        End If
        If Not IsNull(rsEmp("ED_LOC")) Then
            rsWrk("TT_OLDDIV") = Left(rsEmp("ED_LOC"), 4) 'LOCATION
        End If
        If Not IsNull(rsEmp("ED_ORG")) Then
            rsWrk("TT_ORG") = Left(rsEmp("ED_ORG"), 4)
        End If
        If Not IsNull(rsEmp("ED_EMP")) Then
            rsWrk("TT_EMP") = Left(rsEmp("ED_EMP"), 4)
        End If
        rsWrk("TT_JREAS") = rsEmp("JH_JOB") 'employee positin code
        If Not IsNull(rsEmp("ED_DIV")) Then
            rsWrk("TT_NEWDIV") = Left(rsEmp("ED_DIV"), 4)
        End If
        rsWrk("TT_LANG2") = rsEmp("ED_GLNO")
        rsWrk("TT_NAMEFLD") = Left(getDeptDescPub(rsEmp("ED_DEPTNO")), 40) '
        If Not IsNull(rsEmp("JB_BAND")) Then
            rsWrk("TT_BRELATE") = Left(rsEmp("JB_BAND"), 6)
        End If
        
        If Not IsNull(rsEmp("JB_JOBCODE")) Then
            rsWrk("TT_OLDJOB") = Left(rsEmp("JB_JOBCODE"), 6)
        End If
        rsWrk("TT_JOB") = Left(getNewJobCodeDescPub(rsEmp("JB_JOBCODE")), 50) 'Job Desc

        If Not IsNull(rsEmp("JB_DESCR")) Then '
            'rsWrk("TT_COURSE") = Left(getPosDesc(rsEmp("JH_JOB")), 60) 'Position desc 'JB_DESCR
            rsWrk("TT_COURSE") = Left(rsEmp("JB_DESCR"), 60) 'Position desc 'JB_DESCR
        End If
        If Not IsNull(rsEmp("JB_REPTAU")) Then 'Reports To Postion
            rsWrk("TT_JOBCODE") = Left(rsEmp("JB_REPTAU"), 25)
            rsWrk("TT_GRID") = Left(getPosMasterValueByField(rsEmp("JB_REPTAU"), "JB_DESCR"), 50)
        End If
        If Not IsNull(rsEmp("JH_REPTAU")) Then 'Reports To Employee #
            rsWrk("TT_REPTAU") = rsEmp("JH_REPTAU")
            rsWrk("TT_RESULTD") = Left((GetEmpData(rsEmp("JH_REPTAU"), "ED_SURNAME") & "," & GetEmpData(rsEmp("JH_REPTAU"), "ED_FNAME")), 50)
        End If
        rsWrk("TT_BNAME") = Left(GetTABLCodePub("JBGC", rsEmp("JB_GRPCD")), 40) 'Pos Group desc
        rsWrk("TT_SALR1") = Left(rsEmp("ED_SURNAME"), 30)
        rsWrk("TT_SALR2") = Left(rsEmp("ED_FNAME"), 30)
        
        'populate Budgeted Position - begin
        rsWrk("TT_CCOST") = 0 'JG_BUDGNBR
        rsWrk("TT_ECOST") = 0 'JG_NBRFIL
        rsWrk("TT_PCE") = 0 'JG_FTENUM
        rsWrk("TT_PCC") = 0 'JG_FTENUMFILL
        If Not IsNull(rsEmp("ED_DIV")) Then
            SQLQ = "SELECT * FROM HRJOBBUD WHERE (1=1) "
            'SQLQ = SQLQ & "AND JG_CURRENT = 1 "
            SQLQ = SQLQ & "AND JG_CODE = '" & rsEmp("JH_JOB") & "' "
            SQLQ = SQLQ & "AND JG_DIV = '" & rsEmp("ED_DIV") & "' "
            SQLQ = SQLQ & "ORDER BY JG_FRDATE DESC " 'Ticket #28341 Franks 03/21/2016
            If snapBudget.State <> 0 Then snapBudget.Close
            snapBudget.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not snapBudget.EOF Then '
                If IsNull(snapBudget("JG_BUDGNBR")) Then
                    rsWrk("TT_CCOST") = 0
                Else
                    rsWrk("TT_CCOST") = snapBudget("JG_BUDGNBR")
                End If
                If IsNull(snapBudget("JG_NBRFIL")) Then
                    rsWrk("TT_ECOST") = 0
                Else
                    rsWrk("TT_ECOST") = snapBudget("JG_NBRFIL")
                End If
                If IsNull(snapBudget("JG_FTENUM")) Then
                    rsWrk("TT_PCE") = 0
                Else
                    rsWrk("TT_PCE") = snapBudget("JG_FTENUM")
                End If

                If IsNull(snapBudget("JG_FTENUMFILL")) Then
                    rsWrk("TT_PCC") = 0
                Else
                    rsWrk("TT_PCC") = snapBudget("JG_FTENUMFILL")
                End If
                
                '''Ticket #27820 Franks 11/25/2015 - employee's FTE # from their current position
                ''SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & rsEmp("JH_EMPNBR")
                ''If rsPos.State <> 0 Then rsPos.Close
                ''rsPos.Open SQLQ, gdbAdoIhr001, adOpenStatic
                ''If Not rsPos.EOF Then
                ''    If Not IsNull(rsPos("JH_FTENUM")) Then
                ''        rsWrk("TT_PCC") = rsPos("JH_FTENUM")
                ''    End If
                ''End If
                '''Ticket #27820 Franks 11/25/2015 - end
            End If
            snapBudget.Close
        End If
        'populate Budgeted Position - end
        rsWrk.Update
        
        rsWrk("TT_SKLCOM") = getGroup1Desc(rsEmp, rsWrk) '50 Group 1 Desc
        rsWrk.Update
        
        rsEmp.MoveNext
    Loop
    rsEmp.Close
    MDIMain.panHelp(0).FloodPercent = 100
    
    'If optVacancy.Value Then 'Vacancy
    '    SQLQ = "DELETE FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
    '    SQLQ = SQLQ & "AND (TT_CCOST - TT_ECOST) > 0 "
    '    gdbAdoIhr001.Execute SQLQ
    'End If
    'If optOverBud.Value Then 'optOverBud
    '    SQLQ = "DELETE FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
    '    SQLQ = SQLQ & "AND (TT_CCOST - TT_ECOST) < 0 "
    '    gdbAdoIhr001.Execute SQLQ
    'End If
    
    'Populate Excel file - begin
    'Create new WorkBook of Excel
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(xlsFileMat)
    Set exSheet = exBook.Worksheets(1)
    
    ''''exSheet.Cells(2, 2) = "Status of WPS Training for " & Year(Date)
    ''exSheet.Cells(1, 2) = "" & Format(Date, "MMM dd, YYYY") & ""
    ''exSheet.Cells(2, 2) = "" & Time$ & ""
    
    'First line of data
    xStartLine = 2 '5
    
    SQLQ = "SELECT * FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
    xOrderBy = getExcelOrderBy
    SQLQ = SQLQ & "ORDER BY " & xOrderBy & " "
    If rsWrk.State <> 0 Then rsWrk.Close
    rsWrk.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsWrk.EOF Then
        totNum = rsWrk.RecordCount: I = 0
    End If
            
    xRow = xStartLine
    'xTotStaff = 0
    'xTotPlant = 0
    xTotBUDGNBR = 0: xTotNBRFIL = 0: xTotFTENUM = 0: xTotFTENUMFILL = 0
    xSubBUDGNBR = 0: xSubNBRFIL = 0: xSubFTENUM = 0: xSubFTENUMFILL = 0
    xPosBUDGNBR = 0: xPosNBRFIL = 0: xPosFTENUM = 0: xPosFTENUMFILL = 0

    If Not rsWrk.EOF Then
        tmpGroupDesc = rsWrk("TT_SKLCOM")
        xCurPosCode = rsWrk("TT_JREAS")
        xFTENo = rsWrk("TT_PCE")
        xFTEFilled = rsWrk("TT_PCC")
    End If
    xDisFlag = True
    Do While Not rsWrk.EOF
        If (I / totNum) <= 1 Then
            MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
            I = I + 1
        End If
        DoEvents
        
'If rsWrk("TT_EMPNBR") = 10601162 Then
'Debug.Print ""
'End If

        'sub total
        If Not (tmpGroupDesc = rsWrk("TT_SKLCOM")) Then
            exSheet.Cells(xRow, 7 + 3 + 2 + 6 + 1 + 2 + 1 + 2).Font.Bold = True
            exSheet.Cells(xRow, 13 + 6 + 2 + 1 + 2 + 1 + 2).Font.Bold = True
            exSheet.Cells(xRow, 14 + 6 + 2 + 1 + 2 + 1 + 2).Font.Bold = True
            exSheet.Cells(xRow, 15 + 6 + 2 + 1 + 2 + 1 + 2).Font.Bold = True
            exSheet.Cells(xRow, 16 + 6 + 2 + 1 + 2 + 1 + 2).Font.Bold = True
            exSheet.Cells(xRow, 17 + 6 + 2 + 1 + 2 + 1 + 2).Font.Bold = True
            exSheet.Cells(xRow, 18 + 6 + 2 + 1 + 2 + 1 + 2).Font.Bold = True
            
            exSheet.Cells(xRow, 7 + 3 + 2 + 6 + 1 + 2 + 1 + 2) = "Sub Total of " & tmpGroupDesc
            exSheet.Cells(xRow, 13 + 6 + 2 + 1 + 2 + 1 + 2) = xSubBUDGNBR
            exSheet.Cells(xRow, 14 + 6 + 2 + 1 + 2 + 1 + 2) = xSubNBRFIL
            exSheet.Cells(xRow, 15 + 6 + 2 + 1 + 2 + 1 + 2) = xSubBUDGNBR - xSubNBRFIL
            exSheet.Cells(xRow, 16 + 6 + 2 + 1 + 2 + 1 + 2) = xSubFTENUM
            exSheet.Cells(xRow, 17 + 6 + 2 + 1 + 2 + 1 + 2) = xSubFTENUMFILL
            exSheet.Cells(xRow, 18 + 6 + 2 + 1 + 2 + 1 + 2) = xSubFTENUM - xSubFTENUMFILL
            xRow = xRow + 2
            xSubBUDGNBR = 0: xSubNBRFIL = 0: xSubFTENUM = 0: xSubFTENUMFILL = 0
            tmpGroupDesc = rsWrk("TT_SKLCOM")
        End If
        
        If Not xCurPosCode = rsWrk("TT_JREAS") Then
            xCurPosCode = rsWrk("TT_JREAS")
            xDisFlag = True
            
            'Ticket #27820 Franks 11/27/2015 - begin
            '"   If Column (Vacancy/Over Budget FTE) is GT 0, add a "Vacant" employee list with column U completed only
            If xFTENo - xFTEFilled > 0 Then
                For xCnt = 1 To xFTENo - xFTEFilled
                    
                    'If Not IsEmpty(exSheet.Cells(xRow - 1, 8 + 1 + 1)) Then
                        'Plant/Division
                        exSheet.Cells(xRow - 1 + xCnt, 1) = exSheet.Cells(xRow - 1, 1)  'POS Code
                        exSheet.Cells(xRow - 1 + xCnt, 1 + 1) = exSheet.Cells(xRow - 1, 1 + 1) 'COUNTRY
                        exSheet.Cells(xRow - 1 + xCnt, 2 + 1) = exSheet.Cells(xRow - 1, 2 + 1) 'REGION
                        exSheet.Cells(xRow - 1 + xCnt, 3 + 1) = exSheet.Cells(xRow - 1, 3 + 1) 'SECTION
                        exSheet.Cells(xRow - 1 + xCnt, 4 + 1) = exSheet.Cells(xRow - 1, 4 + 1) 'LOC
                        exSheet.Cells(xRow - 1 + xCnt, 5 + 1) = exSheet.Cells(xRow - 1, 5 + 1) ' UNION
                        exSheet.Cells(xRow - 1 + xCnt, 6 + 1) = exSheet.Cells(xRow - 1, 6 + 1) 'Status
                        exSheet.Cells(xRow - 1 + xCnt, 6 + 1 + 1) = exSheet.Cells(xRow - 1, 6 + 1 + 1) ' GL
                        exSheet.Cells(xRow - 1 + xCnt, 7 + 1 + 1) = exSheet.Cells(xRow - 1, 7 + 1 + 1) ' DEPT
                        exSheet.Cells(xRow - 1 + xCnt, 8 + 1 + 1) = exSheet.Cells(xRow - 1, 8 + 1 + 1) ' Band
                        exSheet.Cells(xRow - 1 + xCnt, 8 + 1 + 1 + 1) = exSheet.Cells(xRow - 1, 8 + 1 + 1 + 1) ' JOB
                        exSheet.Cells(xRow - 1 + xCnt, 9 + 1 + 1 + 1) = exSheet.Cells(xRow - 1, 9 + 1 + 1 + 1) ' JOB Desc
                        exSheet.Cells(xRow - 1 + xCnt, 10 + 1 + 1 + 1) = exSheet.Cells(xRow - 1, 10 + 1 + 1 + 1) ' POS
                        exSheet.Cells(xRow - 1 + xCnt, 11 + 1 + 1 + 1) = exSheet.Cells(xRow - 1, 11 + 1 + 1 + 1) ' POS Desc
                        exSheet.Cells(xRow - 1 + xCnt, 12 + 1 + 1 + 1) = exSheet.Cells(xRow - 1, 12 + 1 + 1 + 1) ' POS GROUP
                        exSheet.Cells(xRow - 1 + xCnt, 13 + 1 + 1 + 1) = exSheet.Cells(xRow - 1, 13 + 1 + 1 + 1) 'Position Status Description
                        exSheet.Cells(xRow - 1 + xCnt, 15 + 1 + 1) = exSheet.Cells(xRow - 1, 15 + 1 + 1) 'Reports To Position
                        exSheet.Cells(xRow - 1 + xCnt, 16 + 1 + 1) = exSheet.Cells(xRow - 1, 16 + 1 + 1) 'Reports To Position Description
                        
                        exSheet.Cells(xRow - 1 + xCnt, 17 + 1 + 1) = exSheet.Cells(xRow - 1, 17 + 1 + 1)  'Reports To Employee # '?
                        exSheet.Cells(xRow - 1 + xCnt, 18 + 1 + 1) = exSheet.Cells(xRow - 1, 18 + 1 + 1)  'Reports To Employee Name '?
                        
                        'exSheet.Cells(xRow - 1 + xCnt, 15) = "Vacant " & Trim(Str(xCnt))
                        exSheet.Cells(xRow - 1 + xCnt, 15 + 2 + 1 + 2 + 1 + 2 + 1 + 2) = "Vacant " 'Ticket #28970 Franks 07/26/2016 - remove 1,2,3...
                    'End If
                Next
                xRow = xRow + (xCnt - 1)
            End If
            'Ticket #27820 Franks 11/27/2015 - end
        End If
        
        'Plant/Division
        exSheet.Cells(xRow, 1) = rsWrk("TT_JREAS")  '
        exSheet.Cells(xRow, 1 + 1) = rsWrk("TT_CHAR10") 'COUNTRY
        exSheet.Cells(xRow, 2 + 1) = rsWrk("TT_LANG1") 'REGION
        exSheet.Cells(xRow, 3 + 1) = rsWrk("TT_SHIFT") 'SECTION
        exSheet.Cells(xRow, 4 + 1) = rsWrk("TT_OLDDIV") 'LOC
        exSheet.Cells(xRow, 5 + 1) = rsWrk("TT_ORG") ' UNION
        exSheet.Cells(xRow, 6 + 1) = rsWrk("TT_EMP") 'STATUS
        exSheet.Cells(xRow, 6 + 1 + 1) = rsWrk("TT_LANG2") ' GL
        exSheet.Cells(xRow, 7 + 1 + 1) = rsWrk("TT_NAMEFLD") ' DEPT
        exSheet.Cells(xRow, 8 + 1 + 1) = rsWrk("TT_BRELATE") 'Band
        exSheet.Cells(xRow, 8 + 1 + 1 + 1) = rsWrk("TT_OLDJOB") ' JOB
        exSheet.Cells(xRow, 9 + 1 + 1 + 1) = rsWrk("TT_JOB") ' JOB Desc
        exSheet.Cells(xRow, 10 + 1 + 1 + 1) = rsWrk("TT_JREAS") ' POS
        exSheet.Cells(xRow, 11 + 1 + 1 + 1) = rsWrk("TT_COURSE") ' POS Desc
        exSheet.Cells(xRow, 12 + 1 + 1 + 1) = rsWrk("TT_BNAME") ' POS GROUP
        
        'Ticket #27820 Franks 11/25/2015
        '"   Between L & M, add Position Status Description.
        xStr = GetJobData(xCurPosCode, "JB_STATUS", "")
        If Len(xStr) > 0 Then
            xStr = GetTABLCodePub("JBST", xStr)
            If Len(xStr) > 0 Then
                exSheet.Cells(xRow, 13 + 1 + 1 + 1) = xStr
            End If
        End If
        
        'Ticket #29320 Franks 10/11/2016
        'add Reports To Position (from HRJOB) and the Reports To Position Description in the report between columns M & O. Both code and description
        xReptPos1 = "": xReptPos1Desc = ""
        If Not IsNull(rsWrk("TT_JOBCODE")) Then
            exSheet.Cells(xRow, 14 + 1 + 1 + 1) = rsWrk("TT_JOBCODE")
            exSheet.Cells(xRow, 15 + 1 + 1 + 1) = rsWrk("TT_GRID")
            xReptPos1 = rsWrk("TT_JOBCODE")
            xReptPos1Desc = rsWrk("TT_GRID")
        End If
        
        'Ticket #29406 Franks 11/02/2016
        '"   Insert between Cell Q and R - Reports To Employee # (Position Master) and Reports to Employee Name (Position). If there isn't an active employee, write Vacant in each column. "Vacant" should be in RED
        xStr = GetReportingAuth1EmpNoBasePosMaster(rsWrk("TT_JREAS"))
        If Len(xStr) = 0 Then
            'exSheet.Cells(xRow, 19).Interior.Color = RED
            exSheet.Cells(xRow, 19).Font.Color = RED
            exSheet.Cells(xRow, 19) = "Vacant"
            exSheet.Cells(xRow, 20).Font.Color = RED
            exSheet.Cells(xRow, 20) = "Vacant"
        Else
            exSheet.Cells(xRow, 19) = xStr
            exSheet.Cells(xRow, 20) = (GetEmpData(xStr, "ED_SURNAME") & "," & GetEmpData(xStr, "ED_FNAME"))
        End If
        xReptNo1 = xStr
        If IsNull(rsWrk("TT_REPTAU")) Then
            xReptNo2 = ""
        Else
            xReptNo2 = rsWrk("TT_REPTAU")
        End If
        
        xReptPos2 = "": xReptPos2Desc = ""
        If Not IsNull(rsWrk("TT_REPTAU")) Then
            'Reports To Position ( Employee Record)
            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & rsWrk("TT_REPTAU")
            If rsPos.State <> 0 Then rsPos.Close
            rsPos.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsPos.EOF Then
                'If Not IsNull(rsPos("JH_FTENUM")) Then
                '    rsWrk("TT_PCC") = rsPos("JH_FTENUM")
                'End If
                xReptPos2 = rsPos("JH_JOB")
                xReptPos2Desc = Left(getPosDesc(rsPos("JH_JOB")), 60)
                
                If Len(xReptPos1) > 0 And Len(xReptPos2) > 0 Then
                    If Not (xReptPos1 = xReptPos2) Then
                        exSheet.Cells(xRow, 21).Font.Color = RED
                    End If
                End If
                exSheet.Cells(xRow, 18 + 1 + 2) = rsPos("JH_JOB")
                
                If Len(xReptPos1Desc) > 0 And Len(xReptPos2Desc) > 0 Then
                    If Not (xReptPos1Desc = xReptPos2Desc) Then
                        exSheet.Cells(xRow, 22).Font.Color = RED
                    End If
                End If
                exSheet.Cells(xRow, 19 + 1 + 2) = xReptPos2Desc ' Left(getPosDesc(rsPos("JH_JOB")), 60)
                
            End If
            rsPos.Close
            
            'Reports To Employee #
            If Len(xReptNo1) > 0 And Len(xReptNo2) > 0 Then
                If Not (xReptNo1 = xReptNo2) Then
                    exSheet.Cells(xRow, 23).Font.Color = RED
                    exSheet.Cells(xRow, 24).Font.Color = RED
                End If
            End If
            exSheet.Cells(xRow, 17 + 1 + 2 + 1 + 2) = rsWrk("TT_REPTAU")
            exSheet.Cells(xRow, 18 + 1 + 2 + 1 + 2) = rsWrk("TT_RESULTD")
        End If
        
        exSheet.Cells(xRow, 13 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2) = rsWrk("TT_EMPNBR") '
        exSheet.Cells(xRow, 14 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2) = rsWrk("TT_SALR1") & "," & rsWrk("TT_SALR2") '
        
        'If Not xCurPosCode = rsWrk("TT_JREAS") Then
        '    xCurPosCode = rsWrk("TT_JREAS")
        '    xDisFlag = True
        'End If
        
        If xDisFlag Then
            xFTENo = rsWrk("TT_PCE")
            xFTEFilled = rsWrk("TT_PCC")
        
            exSheet.Cells(xRow, 15 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2) = rsWrk("TT_CCOST")
            exSheet.Cells(xRow, 16 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2) = rsWrk("TT_ECOST")
            exSheet.Cells(xRow, 17 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2) = rsWrk("TT_CCOST") - rsWrk("TT_ECOST")
            exSheet.Cells(xRow, 18 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2) = rsWrk("TT_PCE")
            exSheet.Cells(xRow, 19 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2) = rsWrk("TT_PCC")
            exSheet.Cells(xRow, 20 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2) = rsWrk("TT_PCE") - rsWrk("TT_PCC")
            
            'This position amount
            xPosBUDGNBR = rsWrk("TT_CCOST")
            xPosNBRFIL = rsWrk("TT_ECOST")
            xPosFTENUM = rsWrk("TT_PCE")
            xPosFTENUMFILL = rsWrk("TT_PCC")
        
            xSubBUDGNBR = xSubBUDGNBR + rsWrk("TT_CCOST")
            xSubNBRFIL = xSubNBRFIL + rsWrk("TT_ECOST")
            xSubFTENUM = xSubFTENUM + rsWrk("TT_PCE")
            xSubFTENUMFILL = xSubFTENUMFILL + rsWrk("TT_PCC")
            
            
            'groud total
            xTotBUDGNBR = xTotBUDGNBR + rsWrk("TT_CCOST")
            xTotNBRFIL = xTotNBRFIL + rsWrk("TT_ECOST")
            xTotFTENUM = xTotFTENUM + rsWrk("TT_PCE")
            xTotFTENUMFILL = xTotFTENUMFILL + rsWrk("TT_PCC")
        
            xDisFlag = False
        End If

        xRow = xRow + 1
        rsWrk.MoveNext
    Loop
    
    'Total ----
    If xRow > xStartLine Then
    
            'Ticket #27820 Franks 11/27/2015 - begin
            '"   If Column (Vacancy/Over Budget FTE) is GT 0, add a "Vacant" employee list with column U completed only
            If xFTENo - xFTEFilled > 0 Then
                For xCnt = 1 To xFTENo - xFTEFilled

                    'Plant/Division
                    exSheet.Cells(xRow - 1 + xCnt, 1) = exSheet.Cells(xRow - 1, 1) ' POS
                    exSheet.Cells(xRow - 1 + xCnt, 1 + 1) = exSheet.Cells(xRow - 1, 1 + 1) 'COUNTRY
                    exSheet.Cells(xRow - 1 + xCnt, 2 + 1) = exSheet.Cells(xRow - 1, 2 + 1) 'REGION
                    exSheet.Cells(xRow - 1 + xCnt, 3 + 1) = exSheet.Cells(xRow - 1, 3 + 1) 'SECTION
                    exSheet.Cells(xRow - 1 + xCnt, 4 + 1) = exSheet.Cells(xRow - 1, 4 + 1) 'LOC
                    exSheet.Cells(xRow - 1 + xCnt, 5 + 1) = exSheet.Cells(xRow - 1, 5 + 1) ' UNION
                    exSheet.Cells(xRow - 1 + xCnt, 6 + 1) = exSheet.Cells(xRow - 1, 6 + 1) ' GL
                    exSheet.Cells(xRow - 1 + xCnt, 7 + 1) = exSheet.Cells(xRow - 1, 7 + 1) ' DEPT
                    exSheet.Cells(xRow - 1 + xCnt, 8 + 1) = exSheet.Cells(xRow - 1, 8 + 1) ' Band
                    exSheet.Cells(xRow - 1 + xCnt, 8 + 1 + 1) = exSheet.Cells(xRow - 1, 8 + 1 + 1) ' JOB
                    exSheet.Cells(xRow - 1 + xCnt, 9 + 1 + 1) = exSheet.Cells(xRow - 1, 9 + 1 + 1) ' JOB Desc
                    exSheet.Cells(xRow - 1 + xCnt, 10 + 1 + 1) = exSheet.Cells(xRow - 1, 10 + 1 + 1) ' POS
                    exSheet.Cells(xRow - 1 + xCnt, 11 + 1 + 1) = exSheet.Cells(xRow - 1, 11 + 1 + 1) ' POS Desc
                    exSheet.Cells(xRow - 1 + xCnt, 12 + 1 + 1) = exSheet.Cells(xRow - 1, 12 + 1 + 1) ' POS GROUP
                    exSheet.Cells(xRow - 1 + xCnt, 13 + 1 + 1) = exSheet.Cells(xRow - 1, 13 + 1 + 1)
                    exSheet.Cells(xRow - 1 + xCnt, 15 + 1) = exSheet.Cells(xRow - 1, 15 + 1) 'Reports To Position
                    exSheet.Cells(xRow - 1 + xCnt, 16 + 1) = exSheet.Cells(xRow - 1, 16 + 1) 'Reports To Position Description
                    exSheet.Cells(xRow - 1 + xCnt, 18) = exSheet.Cells(xRow - 1, 18)
                    exSheet.Cells(xRow - 1 + xCnt, 19) = exSheet.Cells(xRow - 1, 19)
                    exSheet.Cells(xRow - 1 + xCnt, 20) = exSheet.Cells(xRow - 1, 20)
                    
                    'exSheet.Cells(xRow - 1 + xCnt, 15) = "Vacant " & Trim(Str(xCnt))
                    exSheet.Cells(xRow - 1 + xCnt, 15 + 2 + 1 + 2 + 2 + 1 + 1 + 2) = "Vacant " 'Ticket #28970 Franks 07/26/2016 - remove 1,2,3...
                Next
                xRow = xRow + (xCnt - 1)
            End If
            'Ticket #27820 Franks 11/27/2015 - end
            
    
            If comGroup(0).Text = "(none)" Then
            Else
            'xRow = xRow + 1
            exSheet.Cells(xRow, 7 + 2 + 6 + 1 + 1 + 2 + 1 + 2).Font.Bold = True
            exSheet.Cells(xRow, 15 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2).Font.Bold = True
            exSheet.Cells(xRow, 16 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2).Font.Bold = True
            exSheet.Cells(xRow, 17 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2).Font.Bold = True
            exSheet.Cells(xRow, 18 + 1 + 2 + 1 + 1 + 2 + 1 + 2).Font.Bold = True
            exSheet.Cells(xRow, 19 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2).Font.Bold = True
            exSheet.Cells(xRow, 20 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2).Font.Bold = True
            exSheet.Cells(xRow, 7 + 2 + 6 + 1 + 1 + 2 + 1 + 2) = "Sub Total of " & tmpGroupDesc
            exSheet.Cells(xRow, 15 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2) = xSubBUDGNBR
            exSheet.Cells(xRow, 16 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2) = xSubNBRFIL
            exSheet.Cells(xRow, 17 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2) = xSubBUDGNBR - xSubNBRFIL
            exSheet.Cells(xRow, 18 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2) = xSubFTENUM
            exSheet.Cells(xRow, 19 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2) = xSubFTENUMFILL
            exSheet.Cells(xRow, 20 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2) = xSubFTENUM - xSubFTENUMFILL
            xRow = xRow + 1
            End If
            
            xRow = xRow + 1
            exSheet.Cells(xRow, 7 + 2 + 6 + 1 + 1 + 2 + 1 + 2).Font.Bold = True
            exSheet.Cells(xRow, 15 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2).Font.Bold = True
            exSheet.Cells(xRow, 16 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2).Font.Bold = True
            exSheet.Cells(xRow, 17 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2).Font.Bold = True
            exSheet.Cells(xRow, 18 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2).Font.Bold = True
            exSheet.Cells(xRow, 19 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2).Font.Bold = True
            exSheet.Cells(xRow, 20 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2).Font.Bold = True
            exSheet.Cells(xRow, 7 + 2 + 6 + 1 + 1 + 2 + 1 + 2) = "Grand Total "
            exSheet.Cells(xRow, 15 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2) = xTotBUDGNBR
            exSheet.Cells(xRow, 16 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2) = xTotNBRFIL
            exSheet.Cells(xRow, 17 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2) = xTotBUDGNBR - xTotNBRFIL
            exSheet.Cells(xRow, 18 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2) = xTotFTENUM
            exSheet.Cells(xRow, 19 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2) = xTotFTENUMFILL
            exSheet.Cells(xRow, 20 + 1 + 2 + 1 + 2 + 1 + 2 + 1 + 2) = xTotFTENUM - xTotFTENUMFILL
            xRow = xRow + 1
    End If
    'rsWrk.Close
    
    exBook.Save
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing

    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = " "
    Screen.MousePointer = DEFAULT

    Call Pause(1)
    
    'If Not comGroup(0).Text = "(none)" Then
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
    'End If
    
    Exit Sub
ReportShowEmp_Err:
    MsgBox Err.Description
    'Resume Next
End Sub

Private Function getExcelOrderBy()
Dim xStr As String
Dim retval

'TT_SKLCOM - Group 1 Desc
'TT_SHIFT - Plant
'TT_ORG - Union
'TT_JREAS - employee positin code
'TT_COURSE - Position desc
'TT_SALR1 - ED_SURNAME
'TT_SALR2 - ED_FNAME

    'retVal = "TT_SHIFT,TT_NEWDIV,TT_COURSE,TT_SALR1,TT_SALR2"
    retval = "TT_SKLCOM,TT_SALR1,TT_SALR2" 'sort by name - none ; TT_SKLCOM is blank
    
    Select Case comGroup(0).Text
    Case lStr("Division")
        'retval = "TT_NEWDIV"
        'retVal = "TT_NEWDIV,TT_SHIFT,TT_COURSE,TT_SALR1,TT_SALR2"
        'retval = "TT_SKLCOM,TT_SHIFT,TT_JREAS,TT_COURSE,TT_SALR1,TT_SALR2"
        retval = "TT_SKLCOM,TT_COURSE,TT_JREAS,TT_SALR1,TT_SALR2"
    Case lStr("Department")
        'retval = "TT_NAMEFLD"
        'retVal = "TT_NAMEFLD,TT_SHIFT,TT_NEWDIV,TT_COURSE,TT_SALR1,TT_SALR2"
        'retval = "TT_SKLCOM,TT_SHIFT,TT_NEWDIV,TT_JREAS,TT_COURSE,TT_SALR1,TT_SALR2"
        retval = "TT_SKLCOM,TT_COURSE,TT_JREAS,TT_SALR1,TT_SALR2"
    Case lStr("Location")
        'retval = "TT_OLDDIV"
        'retVal = "TT_OLDDIV,TT_SHIFT,TT_NEWDIV,TT_COURSE,TT_SALR1,TT_SALR2"
        'retval = "TT_SKLCOM,TT_SHIFT,TT_NEWDIV,TT_JREAS,TT_COURSE,TT_SALR1,TT_SALR2"
        retval = "TT_SKLCOM,TT_COURSE,TT_JREAS,TT_SALR1,TT_SALR2"
    Case lStr("Section")
        'retval = "TT_SHIFT"
        'retVal = "TT_SHIFT,TT_NEWDIV,TT_COURSE,TT_SALR1,TT_SALR2"
        'retval = "TT_SKLCOM,TT_NEWDIV,TT_JREAS,TT_COURSE,TT_SALR1,TT_SALR2"
        retval = "TT_SKLCOM,TT_COURSE,TT_JREAS,TT_SALR1,TT_SALR2"
    Case lStr("Region")
        'retval = "TT_LANG1"
        'retVal = "TT_LANG1,TT_SHIFT,TT_NEWDIV,TT_COURSE,TT_SALR1,TT_SALR2"
        'retval = "TT_SKLCOM,TT_SHIFT,TT_NEWDIV,TT_JREAS,TT_COURSE,TT_SALR1,TT_SALR2"
        retval = "TT_SKLCOM,TT_COURSE,TT_JREAS,TT_SALR1,TT_SALR2"
    Case "Job"
        'retval = "TT_JOB"
        'retVal = "TT_JOB,TT_SHIFT,TT_NEWDIV,TT_COURSE,TT_SALR1,TT_SALR2"
        'retval = "TT_SKLCOM,TT_SHIFT,TT_NEWDIV,TT_JREAS,TT_COURSE,TT_SALR1,TT_SALR2"
        retval = "TT_SKLCOM,TT_COURSE,TT_JREAS,TT_SALR1,TT_SALR2"
    Case "Position Group"
        'retval = "TT_BNAME"
        'retVal = "TT_BNAME,TT_SHIFT,TT_NEWDIV,TT_COURSE,TT_SALR1,TT_SALR2"
        'retval = "TT_SKLCOM,TT_SHIFT,TT_NEWDIV,TT_JREAS,TT_COURSE,TT_SALR1,TT_SALR2"
        retval = "TT_SKLCOM,TT_COURSE,TT_JREAS,TT_SALR1,TT_SALR2"
    Case "Position"
        'retval = "TT_COURSE"
        'retVal = "TT_COURSE,TT_SHIFT,TT_NEWDIV,TT_COURSE,TT_SALR1,TT_SALR2"
        'retval = "TT_SKLCOM,TT_SHIFT,TT_NEWDIV,TT_JREAS,TT_COURSE,TT_SALR1,TT_SALR2"
        retval = "TT_SKLCOM,TT_COURSE,TT_JREAS,TT_SALR1,TT_SALR2"
    Case "(none)" 'G1=None. Report order should be Union/Job/Position
        'retval = "TT_SKLCOM,TT_ORG,TT_JOB,TT_JREAS,TT_COURSE"
        'retval = "TT_SKLCOM,TT_ORG,TT_OLDJOB,TT_JREAS,TT_COURSE" 'TT_OLDJOB - job code;TT_JOB - job desc
        retval = "TT_COURSE,TT_JREAS,TT_SALR1,TT_SALR2"
    End Select
    getExcelOrderBy = retval
End Function

Private Function getGroup1Desc(rs As ADODB.Recordset, rsWrk As ADODB.Recordset)
'Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim retval
    retval = "" 'none
    Select Case comGroup(0).Text
    Case lStr("Division")
        retval = getDivDescPub(rs("ED_DIV"))
    Case lStr("Department")
        retval = getDeptDescPub(rs("ED_DEPTNO"))
    Case lStr("Location")
        If Not IsNull(rs("ED_LOC")) Then
            retval = GetTABLDesc("EDLC", rs("ED_LOC"))
        End If
    Case lStr("Section")
        If Not IsNull(rs("ED_SECTION")) Then
            retval = GetTABLDesc("EDSE", rs("ED_SECTION"))
        End If
    Case lStr("Region")
        If Not IsNull(rs("ED_REGION")) Then
            retval = GetTABLDesc("EDRG", rs("ED_REGION"))
        End If
    Case "Job"
        retval = rsWrk("TT_JOB")
    Case "Position Group"
        retval = rsWrk("TT_BNAME")
    Case "Position"
        'retval = rsWrk("TT_COURSE")
        retval = rsWrk("TT_JREAS")
    End Select
    getGroup1Desc = Left(retval, 60)
End Function

Private Sub ScreenSetTableExps()
    lblStatus.Visible = False
    clpCode(2).Visible = False
    lblPT.Visible = False
    clpPT.Visible = False
    lblAdmin.Visible = False
    clpCode(4).Visible = False
    lblEENum(0).Visible = False
    elpEEID.Visible = False
    
    lblRepGrp.Visible = False
    lblGrp(0).Visible = False
    lblGrp(3).Visible = False
    comGroup(0).Visible = False
    comGroup(1).Visible = False
    
    chkShowEmp.Visible = False
    frmOpt1.Visible = False
    frmOpt2.Visible = False

    lblRegion.Top = lblStatus.Top
    clpCode(3).Top = clpCode(2).Top
    lblSection.Top = lblStatus.Top + 330
    clpCode(5).Top = clpCode(2).Top + 330

    lblPosition.Top = lblStatus.Top + 330 * 2
    clpJob.Top = clpCode(2).Top + 330 * 2

    lblPosGroup.Top = lblStatus.Top + 330 * 3
    clpCode(6).Top = clpCode(2).Top + 330 * 3

    lblJob.Top = lblStatus.Top + 330 * 4
    clpJobMaster.Top = clpCode(2).Top + 330 * 4
    lblYear.Top = lblStatus.Top + 330 * 5
    txtYear.Top = clpCode(2).Top + 330 * 5
    
    frmTabExp.Left = 1880
    frmTabExp.Top = 3870
    frmTabExp.BorderStyle = 0
    frmTabExp.Visible = True
    
End Sub

Private Sub TableExpProgs()

If optExpJob.Value Then
    If CriChkExpJob() Then
        Call set_PrintState(False)
        Call Cri_SetAlExpJob
        Call set_PrintState(True)
    End If
End If

If optExpPos.Value Then
    If CriChkExpJob() Then
        Call set_PrintState(False)
        Call Cri_SetAlExpPosition
        Call set_PrintState(True)
    End If
End If

If optExpBud.Value Then
    If CriChkExpJob() Then
        Call set_PrintState(False)
        Call Cri_SetAlExpBud
        Call set_PrintState(True)
    End If
End If

End Sub

Private Sub Cri_SetAlExpJob()
Dim X%, strRName$
Dim rsJOB As New ADODB.Recordset
Dim exApp As Object, exBook As Object, exSheet As Object
Dim SQLQ, sSQLQ As String
Dim locSQL As String
Dim xlsFileTmp As String
Dim xlsFileMat As String
Dim xRow As Long
Dim xCol As Long
Dim xDisFlag As Boolean
Dim xStartLine, totNum, I

On Error GoTo modSetCriteria_Err

    SQLQ = "SELECT * FROM HRJOBMASTER WHERE (1=1) "
    If Len(clpCode(6).Text) > 0 Then
        SQLQ = SQLQ & " AND JB_GRPCD = '" & clpCode(6).Text & "' "
    End If
    If Len(clpJobMaster.Text) > 0 Then
        SQLQ = SQLQ & " AND JB_JOBCODE = '" & clpJobMaster.Text & "' "
    End If
    If chkIncJobPos.Value Then
    Else
        SQLQ = SQLQ & " AND NOT ((LEFT(JB_JOBDESCR,2) = 'Z ' OR JB_STATUS = 'INAC')) "
    End If
    
    SQLQ = SQLQ & "ORDER BY JB_JOBDESCR "
    If rsJOB.State <> 0 Then rsJOB.Close
    rsJOB.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsJOB.EOF Then
        MsgBox "No record found."
        Exit Sub
    End If
    
    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "WFC Position Control Table Tmp.xls"
    xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "WFC Position Control Job Table(" & Trim(glbUserID) & ").xls"
    
    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    FileCopy xlsFileTmp, xlsFileMat
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(0).FloodPercent = 0
    Screen.MousePointer = HOURGLASS

    'Populate Excel file - begin
    'Create new WorkBook of Excel
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(xlsFileMat)
    Set exSheet = exBook.Worksheets(1)
    
    exSheet.Cells(2, 4) = "Job Master Table Export"
    exSheet.Cells(1, 2) = "" & Format(Date, "MMM dd, YYYY") & ""
    exSheet.Cells(2, 2) = "" & Time$ & ""
    
    exSheet.Cells(4, 1) = "Job Code"
    exSheet.Cells(4, 2) = "Job Description"
    exSheet.Cells(4, 3) = lStr("Job Group")
    exSheet.Cells(4, 4) = lStr("Job Status")
    exSheet.Cells(4, 5) = lStr("Job User Defined 1")
    exSheet.Cells(4, 6) = lStr("Job User Defined 2")
    
    'First line of data
    xStartLine = 5
    

    If Not rsJOB.EOF Then
        totNum = rsJOB.RecordCount: I = 0
    End If
            
    xRow = xStartLine
    
    Do While Not rsJOB.EOF
        If (I / totNum) <= 1 Then
            MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
            I = I + 1
        End If
        DoEvents
        
        'details
        exSheet.Cells(xRow, 1) = rsJOB("JB_JOBCODE")
        exSheet.Cells(xRow, 2) = rsJOB("JB_JOBDESCR")
        exSheet.Cells(xRow, 3) = rsJOB("JB_GRPCD")
        exSheet.Cells(xRow, 4) = rsJOB("JB_STATUS")
        exSheet.Cells(xRow, 5) = rsJOB("JB_USERDEF1")
        exSheet.Cells(xRow, 6) = rsJOB("JB_USERDEF2")

        xRow = xRow + 1
        rsJOB.MoveNext
    Loop
    
   
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

Screen.MousePointer = DEFAULT

Exit Sub

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Job Master", "Job Master", "Select")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub Cri_SetAlExpPosition()
Dim X%, strRName$
Dim rsJOB As New ADODB.Recordset
Dim exApp As Object, exBook As Object, exSheet As Object
Dim SQLQ, sSQLQ As String
Dim locSQL As String
Dim xlsFileTmp As String
Dim xlsFileMat As String
Dim xRow As Long
Dim xCol As Long
Dim xDisFlag As Boolean
Dim xStartLine, totNum, I

On Error GoTo modSetCriteria_Err

    SQLQ = "SELECT * FROM HRJOB WHERE (1=1) "
    If Len(clpDiv.Text) > 0 Then
        SQLQ = SQLQ & " AND (JB_DIV in ('" & Replace(clpDiv.Text, ",", "','") & "')) "
    End If
    If Len(clpCode(1).Text) > 0 Then
        SQLQ = SQLQ & " AND (JB_ORG in ('" & Replace(clpCode(1).Text, ",", "','") & "')) "
    End If '
    If Len(clpCode(3).Text) > 0 Then
        SQLQ = SQLQ & " AND (JB_REGION = '" & clpCode(3).Text & "') "
    End If
    If Len(clpCode(5).Text) > 0 Then
        SQLQ = SQLQ & " AND (JB_SECTION in ('" & Replace(clpCode(5).Text, ",", "','") & "')) "
    End If
    If Len(clpJob.Text) > 0 Then
        SQLQ = SQLQ & " AND (JB_CODE in ('" & Replace(clpJob.Text, ",", "','") & "')) "
    End If
    If Len(clpCode(6).Text) > 0 Then 'Position Group
        SQLQ = SQLQ & " AND (JB_GRPCD = '" & clpCode(6).Text & "') "
    End If
    If Len(clpJobMaster.Text) > 0 Then
        SQLQ = SQLQ & " AND (JB_JOBCODE = '" & clpJobMaster.Text & "') "
    End If
    If chkIncJobPos.Value Then
    Else
        SQLQ = SQLQ & " AND NOT ((LEFT(JB_DESCR,2) = 'Z ' OR JB_STATUS = 'INAC')) "
    End If
    
    SQLQ = SQLQ & "ORDER BY JB_DESCR "
    If rsJOB.State <> 0 Then rsJOB.Close
    rsJOB.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsJOB.EOF Then
        MsgBox "No record found."
        Exit Sub
    End If
    
    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "WFC Position Control Table Tmp.xls"
    xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "WFC Position Control Position Table(" & Trim(glbUserID) & ").xls"
    
    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    FileCopy xlsFileTmp, xlsFileMat
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(0).FloodPercent = 0
    Screen.MousePointer = HOURGLASS

    'Populate Excel file - begin
    'Create new WorkBook of Excel
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(xlsFileMat)
    Set exSheet = exBook.Worksheets(1)
    
    exSheet.Cells(2, 4) = "Position Master Table Export"
    exSheet.Cells(1, 2) = "" & Format(Date, "MMM dd, YYYY") & ""
    exSheet.Cells(2, 2) = "" & Time$ & ""
    
    exSheet.Cells(4, 1) = "Position"
    exSheet.Cells(4, 2) = lStr("Position Description")
    exSheet.Cells(4, 3) = lStr("Position Alternate")
    exSheet.Cells(4, 4) = lStr("Division")
    exSheet.Cells(4, 5) = lStr("Region")
    exSheet.Cells(4, 6) = lStr("Section")
    exSheet.Cells(4, 7) = "Mercer Code"
    exSheet.Cells(4, 8) = "Job Code"
    exSheet.Cells(4, 9) = "Total Points"
    exSheet.Cells(4, 10) = lStr("Position Status")
    exSheet.Cells(4, 11) = lStr("Position Group")
    exSheet.Cells(4, 12) = lStr("Position Level")
    exSheet.Cells(4, 13) = "Band"
    exSheet.Cells(4, 14) = lStr("Union")
    exSheet.Cells(4, 15) = "Market Line"
    exSheet.Cells(4, 16) = "N.O.C. Code"
    exSheet.Cells(4, 17) = "Reports To 1"
    exSheet.Cells(4, 18) = "Reports To 2"
    exSheet.Cells(4, 19) = "Reports To 3"
    exSheet.Cells(4, 20) = "Reports To 4"
    exSheet.Cells(4, 21) = "FTE #"
    exSheet.Cells(4, 22) = "Total FTE #"
    exSheet.Cells(4, 23) = "FTE Hours/Year"
    exSheet.Cells(4, 24) = "Total FTE Hours/Year"
    exSheet.Cells(4, 25) = "Hours/Day"
    exSheet.Cells(4, 26) = lStr("Position User Defined 1")
    exSheet.Cells(4, 27) = lStr("Position User Defined 2")

    'First line of data
    xStartLine = 5
    

    If Not rsJOB.EOF Then
        totNum = rsJOB.RecordCount: I = 0
    End If
            
    xRow = xStartLine
    
    Do While Not rsJOB.EOF
        If (I / totNum) <= 1 Then
            MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
            I = I + 1
        End If
        DoEvents
        
        'details
        exSheet.Cells(xRow, 1) = rsJOB("JB_CODE")
        exSheet.Cells(xRow, 2) = rsJOB("JB_DESCR")
        exSheet.Cells(xRow, 3) = rsJOB("JB_DESCR2")
        exSheet.Cells(xRow, 4) = rsJOB("JB_DIV")
        exSheet.Cells(xRow, 5) = rsJOB("JB_REGION")
        exSheet.Cells(xRow, 6) = rsJOB("JB_SECTION")
        exSheet.Cells(xRow, 7) = rsJOB("JB_MERCER_NO")
        exSheet.Cells(xRow, 8) = rsJOB("JB_JOBCODE")
        exSheet.Cells(xRow, 9) = rsJOB("JB_POINTS")
        exSheet.Cells(xRow, 10) = rsJOB("JB_STATUS")
        exSheet.Cells(xRow, 11) = rsJOB("JB_GRPCD")
        exSheet.Cells(xRow, 12) = rsJOB("JB_LEVEL")
        exSheet.Cells(xRow, 13) = rsJOB("JB_BAND")
        exSheet.Cells(xRow, 14) = rsJOB("JB_ORG")
        exSheet.Cells(xRow, 15) = rsJOB("JB_MARKETLINE")
        exSheet.Cells(xRow, 16) = rsJOB("JB_FEDGRP")
        exSheet.Cells(xRow, 17) = rsJOB("JB_REPTAU")
        exSheet.Cells(xRow, 18) = rsJOB("JB_REPTAU2")
        exSheet.Cells(xRow, 19) = rsJOB("JB_REPTAU3")
        exSheet.Cells(xRow, 20) = rsJOB("JB_REPTAU4")
        exSheet.Cells(xRow, 21) = rsJOB("JB_FTENUM")
        exSheet.Cells(xRow, 22) = rsJOB("JB_FTETOTNU")
        exSheet.Cells(xRow, 23) = rsJOB("JB_FTEHRS")
        exSheet.Cells(xRow, 24) = rsJOB("JB_FTETOTHR")
        exSheet.Cells(xRow, 25) = rsJOB("JB_DHRS")
        exSheet.Cells(xRow, 26) = rsJOB("JB_USERDEF1")
        exSheet.Cells(xRow, 27) = rsJOB("JB_USERDEF1")
        xRow = xRow + 1
        rsJOB.MoveNext
    Loop
   
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

Screen.MousePointer = DEFAULT

Exit Sub

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Position Master", "Position Master", "Select")
'Cri_SetAlExpJob = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub Cri_SetAlExpBud()
Dim X%, strRName$
Dim rsJOB As New ADODB.Recordset
Dim exApp As Object, exBook As Object, exSheet As Object
Dim SQLQ, sSQLQ As String
Dim locSQL As String
Dim xlsFileTmp As String
Dim xlsFileMat As String
Dim xRow As Long
Dim xCol As Long
Dim xDisFlag As Boolean
Dim xStartLine, totNum, I
Dim xTmp

On Error GoTo modSetCriteria_Err

    SQLQ = "SELECT HRJOBBUD.*, JB_DESCR FROM HRJOBBUD LEFT JOIN HRJOB ON HRJOBBUD.JG_CODE = HRJOB.JB_CODE WHERE (1=1) "
    If Len(clpDiv.Text) > 0 Then
        SQLQ = SQLQ & " AND (JG_DIV in ('" & Replace(clpDiv.Text, ",", "','") & "')) "
    End If
    If Len(clpCode(1).Text) > 0 Then
        SQLQ = SQLQ & " AND (JB_ORG in ('" & Replace(clpCode(1).Text, ",", "','") & "')) "
    End If '
    If Len(clpCode(3).Text) > 0 Then
        SQLQ = SQLQ & " AND (JB_REGION = '" & clpCode(3).Text & "') "
    End If
    If Len(clpCode(5).Text) > 0 Then
        SQLQ = SQLQ & " AND (JG_SECTION in ('" & Replace(clpCode(5).Text, ",", "','") & "')) "
    End If
    If Len(clpJob.Text) > 0 Then
        SQLQ = SQLQ & " AND (JG_CODE in ('" & Replace(clpJob.Text, ",", "','") & "')) "
    End If
    If Len(clpCode(6).Text) > 0 Then 'Position Group
        SQLQ = SQLQ & " AND (JB_GRPCD = '" & clpCode(6).Text & "') "
    End If
    If Len(clpJobMaster.Text) > 0 Then
        SQLQ = SQLQ & " AND (JB_JOBCODE = '" & clpJobMaster.Text & "') "
    End If
    'year
    If IsNumeric(txtYear.Text) Then
        If Len(txtYear.Text) = 4 Then
            SQLQ = SQLQ & " AND (JG_YEAR = " & txtYear.Text & ") "
        End If
    End If
    If chkIncJobPos.Value Then
    Else
        SQLQ = SQLQ & " AND NOT ((LEFT(JB_DESCR,2) = 'Z ' OR JB_STATUS = 'INAC')) "
    End If
    
    SQLQ = SQLQ & "ORDER BY JB_DESCR "
    If rsJOB.State <> 0 Then rsJOB.Close
    rsJOB.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsJOB.EOF Then
        MsgBox "No record found."
        Exit Sub
    End If
    
    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "WFC Position Control Table Tmp.xls"
    xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "WFC Position Control Budgeted Pos Table(" & Trim(glbUserID) & ").xls"
    
    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    FileCopy xlsFileTmp, xlsFileMat
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(0).FloodPercent = 0
    Screen.MousePointer = HOURGLASS

    'Populate Excel file - begin
    'Create new WorkBook of Excel
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(xlsFileMat)
    Set exSheet = exBook.Worksheets(1)
    
    exSheet.Cells(2, 4) = "Budgeted Position Master Table Export"
    exSheet.Cells(1, 2) = "" & Format(Date, "MMM dd, YYYY") & ""
    exSheet.Cells(2, 2) = "" & Time$ & ""
    
    exSheet.Cells(4, 1) = "Year"
    exSheet.Cells(4, 2) = "From Date"
    exSheet.Cells(4, 3) = "To Date"
    exSheet.Cells(4, 4) = "Effective Date "
    exSheet.Cells(4, 5) = "Reason"
    exSheet.Cells(4, 6) = "Current"
    exSheet.Cells(4, 7) = "Position"
    exSheet.Cells(4, 8) = lStr("Position Description")
    exSheet.Cells(4, 9) = lStr("Section")
    exSheet.Cells(4, 10) = lStr("Division")
    exSheet.Cells(4, 11) = "Budgeted #Pos'ns"
    exSheet.Cells(4, 12) = "Positions Filled"
    exSheet.Cells(4, 13) = "Vacancy Positions"
    exSheet.Cells(4, 14) = "FTE #"
    exSheet.Cells(4, 15) = "FTE # Filled"
    exSheet.Cells(4, 16) = "Vacancy # FTE"
    exSheet.Cells(4, 17) = "FTE Hours/Year"
    exSheet.Cells(4, 18) = "Total FTE Hours/Year"
    
    'First line of data
    xStartLine = 5
    

    If Not rsJOB.EOF Then
        totNum = rsJOB.RecordCount: I = 0
    End If
            
    xRow = xStartLine
    
    Do While Not rsJOB.EOF
        If (I / totNum) <= 1 Then
            MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
            I = I + 1
        End If
        DoEvents
        
        'details
        exSheet.Cells(xRow, 1) = rsJOB("JG_YEAR")
        exSheet.Cells(xRow, 2) = rsJOB("JG_FRDATE")
        exSheet.Cells(xRow, 3) = rsJOB("JG_TODATE")
        exSheet.Cells(xRow, 4) = rsJOB("JG_EFDATE")
        exSheet.Cells(xRow, 5) = rsJOB("JG_JREASON")
        xTmp = "N"
        If Not IsNull(rsJOB("JG_CURRENT")) Then
            If rsJOB("JG_CURRENT") Then
                xTmp = "Y"
            End If
        End If
        exSheet.Cells(xRow, 6) = xTmp ' rsJOB("JG_CURRENT")
        exSheet.Cells(xRow, 7) = rsJOB("JG_CODE")
        exSheet.Cells(xRow, 8) = rsJOB("JB_DESCR")
        exSheet.Cells(xRow, 9) = rsJOB("JG_SECTION")
        exSheet.Cells(xRow, 10) = rsJOB("JG_DIV")
        exSheet.Cells(xRow, 11) = rsJOB("JG_BUDGNBR")
        exSheet.Cells(xRow, 12) = rsJOB("JG_NBRFIL")
        If Not IsNull(rsJOB("JG_BUDGNBR")) And Not IsNull(rsJOB("JG_BUDGNBR")) Then
            exSheet.Cells(xRow, 13) = rsJOB("JG_BUDGNBR") - rsJOB("JG_BUDGNBR")
        End If
        exSheet.Cells(xRow, 14) = rsJOB("JG_FTENUM")
        exSheet.Cells(xRow, 15) = rsJOB("JG_FTENUMFILL")
        If Not IsNull(rsJOB("JG_FTENUM")) And Not IsNull(rsJOB("JG_FTENUMFILL")) Then
            exSheet.Cells(xRow, 16) = rsJOB("JG_FTENUM") - rsJOB("JG_FTENUMFILL")
        End If
        exSheet.Cells(xRow, 17) = rsJOB("JG_FTEHRS")
        exSheet.Cells(xRow, 18) = rsJOB("JG_FTETOTHR")
        xRow = xRow + 1
        rsJOB.MoveNext
    Loop
   
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

Screen.MousePointer = DEFAULT

Exit Sub

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Budgeted Position Master", "Budgeted Position Master", "Select")
'Cri_SetAlExpJob = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub PosExpScreenSetup()
'disable all fields -  begin
lblDiv.Enabled = False
clpDiv.Enabled = False
lblDept.Enabled = False
clpDept.Enabled = False
lblLocation.Enabled = False
clpCode(0).Enabled = False
lblUnion.Enabled = False
clpCode(1).Enabled = False
lblRegion.Enabled = False
clpCode(3).Enabled = False
lblSection.Enabled = False
clpCode(5).Enabled = False
lblPosition.Enabled = False
clpJob.Enabled = False
lblPosGroup.Enabled = False
clpCode(6).Enabled = False
lblJob.Enabled = False
clpJobMaster.Enabled = False
lblYear.Enabled = False
txtYear.Enabled = False
'disable all fields -  end

If optExpJob.Value Then
    lblPosGroup.Enabled = True
    clpCode(6).Enabled = True
    lblJob.Enabled = True
    clpJobMaster.Enabled = True
End If
If optExpPos.Value Then
    lblDiv.Enabled = True
    clpDiv.Enabled = True
    lblUnion.Enabled = True
    clpCode(1).Enabled = True
    lblRegion.Enabled = True
    clpCode(3).Enabled = True
    lblSection.Enabled = True
    clpCode(5).Enabled = True
    lblPosition.Enabled = True
    clpJob.Enabled = True
    lblPosGroup.Enabled = True
    clpCode(6).Enabled = True
    lblJob.Enabled = True
    clpJobMaster.Enabled = True
End If
If optExpBud.Value Then
    lblDiv.Enabled = True
    clpDiv.Enabled = True
    lblUnion.Enabled = True
    clpCode(1).Enabled = True
    lblRegion.Enabled = True
    clpCode(3).Enabled = True
    lblSection.Enabled = True
    clpCode(5).Enabled = True
    lblPosition.Enabled = True
    clpJob.Enabled = True
    lblPosGroup.Enabled = True
    clpCode(6).Enabled = True
    lblJob.Enabled = True
    clpJobMaster.Enabled = True
    lblYear.Enabled = True
    txtYear.Enabled = True
End If

End Sub

Private Sub optExpBud_Click()
Call PosExpScreenSetup
End Sub

Private Sub optExpJob_Click()
Call PosExpScreenSetup
End Sub

Private Sub optExpPos_Click()
Call PosExpScreenSetup
End Sub
