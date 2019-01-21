VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmRHomeA 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Home Address/Telephone Report"
   ClientHeight    =   7785
   ClientLeft      =   375
   ClientTop       =   915
   ClientWidth     =   11010
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
   ScaleHeight     =   7785
   ScaleWidth      =   11010
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
      TabIndex        =   15
      Tag             =   "Check to Exclude Employees with RET Employment Status"
      Top             =   4980
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
      TabIndex        =   14
      Tag             =   "Check to Exclude Employees with CONP Employment Status"
      Top             =   4680
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.Frame frmGroup 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   120
      TabIndex        =   44
      Top             =   6600
      Width           =   5055
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
         Left            =   1820
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Tag             =   "First Level of grouping records"
         Top             =   225
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
         Left            =   1820
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Tag             =   "Second level of grouping records"
         Top             =   565
         Width           =   2325
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
         TabIndex        =   47
         Top             =   285
         Width           =   885
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
         TabIndex        =   46
         Top             =   625
         Width           =   660
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
         TabIndex        =   45
         Top             =   0
         Width           =   1575
      End
   End
   Begin Threed.SSCheck chkOnlyTerm 
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   5760
      Width           =   2895
      _Version        =   65536
      _ExtentX        =   5106
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Only Terminated Employees"
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
   Begin Threed.SSCheck chkIncludeTerm 
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   5400
      Width           =   2775
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Include Terminated Employees"
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
      Left            =   3600
      TabIndex        =   11
      Top             =   3720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2100
      MaxLength       =   4
      TabIndex        =   10
      Tag             =   "00-Employee Position Shift"
      Top             =   3660
      Visible         =   0   'False
      Width           =   435
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
      Index           =   4
      Left            =   1800
      TabIndex        =   8
      Tag             =   "00-Enter Administered By Code"
      Top             =   3000
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
      Left            =   1800
      TabIndex        =   9
      Tag             =   "00-Enter Section Code"
      Top             =   3330
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   7
      Tag             =   "00-Enter Region Code"
      Top             =   2670
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
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
      Left            =   8640
      Top             =   6480
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
   Begin VB.Frame frmDepend 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   4920
      TabIndex        =   38
      Top             =   4660
      Visible         =   0   'False
      Width           =   4455
      Begin VB.ComboBox comCountry 
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
         Left            =   1980
         TabIndex        =   20
         Tag             =   "00-Country of Employment"
         Top             =   1070
         Width           =   2040
      End
      Begin VB.ComboBox ComStatus 
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
         Left            =   1980
         TabIndex        =   16
         Tag             =   "00-Status i.e. Student "
         Top             =   10
         Width           =   2055
      End
      Begin INFOHR_Controls.DateLookup dlpAsDate 
         Height          =   285
         Left            =   1680
         TabIndex        =   19
         Tag             =   "40-As of Date"
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1255
      End
      Begin MSMask.MaskEdBox medDOW 
         Height          =   285
         Index           =   0
         Left            =   1995
         TabIndex        =   17
         Tag             =   "10-Day of week 1=Sunday, 2=Monday..."
         Top             =   375
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   3
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
         Mask            =   "###"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medDOW 
         Height          =   285
         Index           =   1
         Left            =   2955
         TabIndex        =   18
         Tag             =   "10-Day of week 1=Sunday, 2=Monday..."
         Top             =   375
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   3
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
         Mask            =   "###"
         PromptChar      =   "_"
      End
      Begin Threed.SSCheck chkShowEmp 
         Height          =   225
         Left            =   1980
         TabIndex        =   21
         Tag             =   "If X-Show All Employees"
         Top             =   1515
         Width           =   2835
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   " Show All Employees"
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
      Begin VB.Label lblCountry 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Country"
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
         Left            =   10
         TabIndex        =   43
         Top             =   1130
         Width           =   540
      End
      Begin VB.Label lblDOW 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   2580
         TabIndex        =   42
         Top             =   375
         Width           =   195
      End
      Begin VB.Label lblDOW 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dependent Age Range"
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
         Left            =   15
         TabIndex        =   41
         Top             =   425
         Width           =   1650
      End
      Begin VB.Label lblFromTo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "As of Date"
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
         Left            =   10
         TabIndex        =   40
         Top             =   765
         Width           =   750
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dependent Status"
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
         Left            =   15
         TabIndex        =   39
         Top             =   75
         Width           =   1575
      End
   End
   Begin INFOHR_Controls.CodeLookup clpProv 
      Height          =   285
      Left            =   1800
      TabIndex        =   12
      Tag             =   "31-Province of Residence - Code"
      Top             =   3990
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   4
   End
   Begin INFOHR_Controls.CodeLookup clpProvEmp 
      Height          =   285
      Left            =   1800
      TabIndex        =   13
      Tag             =   "31-Province of Employment - Code"
      Top             =   4320
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   4
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Prov. of Employment"
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
      TabIndex        =   49
      Top             =   4365
      Width           =   1455
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Prov. of Residence"
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
      TabIndex        =   48
      Top             =   4035
      Width           =   1365
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
      Top             =   3705
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
      Top             =   2055
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
      Top             =   3375
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
      Top             =   2715
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
      Top             =   3045
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
      Top             =   1065
      Width           =   615
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
      TabIndex        =   31
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
      TabIndex        =   30
      Top             =   2385
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
      TabIndex        =   29
      Top             =   1725
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
      TabIndex        =   28
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
      TabIndex        =   27
      Top             =   735
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
      TabIndex        =   26
      Top             =   405
      Width           =   555
   End
End
Attribute VB_Name = "frmRHomeA"
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
  If FormHomeAddress% = True Then
    If Not PrtForm("Home Address/Telephone Report Criteria", Me) Then Exit Sub
  Else
    If Not PrtForm("Dependents Report Criteria", Me) Then Exit Sub
  End If
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

Private Sub chkExclCONP_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkExclRET_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkIncludeTerm_Click(Value As Integer)
    If chkIncludeTerm Then
        chkOnlyTerm.Value = False
        elpEEID.LookupType = TERM
    Else
        elpEEID.LookupType = 0
    End If
End Sub

Private Sub chkOnlyTerm_Click(Value As Integer)
    If chkOnlyTerm Then
        chkIncludeTerm.Value = False
        elpEEID.LookupType = TERM
    Else
        elpEEID.LookupType = 0
    End If
End Sub

Private Sub comCountry_Click()
    Call DependStatus
End Sub

Private Sub comGroup_GotFocus(Index As Integer)
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()
If FormHomeAddress% = True Then 'laura Oct 28, 1997
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
    comGroup(1).AddItem "Employee Number"
    If Not glbSQL Then
        comGroup(1).AddItem "Date of Hire"
    End If
    comGroup(0).ListIndex = 0
    comGroup(1).ListIndex = 0
ElseIf FormDepend% = True Then
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")
    comGroup(0).AddItem lStr("Section")
    comGroup(0).AddItem "Employee Name"
    If glbLinamar Then ' Frank May 2,2001
        comGroup(0).AddItem "Employment Type"
        comGroup(0).AddItem ("Home Line")
    End If
    If Not glbMulti Then comGroup(0).AddItem "Shift"
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem "(none)"
    comGroup(1).AddItem "Employee Name"
    comGroup(1).AddItem "Employee Number"
    comGroup(0).ListIndex = 0
    comGroup(1).ListIndex = 0
End If

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

Private Sub Cri_DepFromDate(xDate)
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim x%

If Not IsDate(xDate) Then Exit Sub

TempCri = "({HRDEPEND.DP_DOB} "
TempCri = TempCri & " <= "

dtYYY% = Year(xDate)
dtMM% = month(xDate)
dtDD% = Day(xDate)
TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "

GoTo Cri_FTDatst

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

Private Sub Cri_DepToDate(xDate)
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim x%

If Not IsDate(xDate) Then Exit Sub

TempCri = "({HRDEPEND.DP_DOB} "

TempCri = TempCri & " >= "

dtYYY% = Year(xDate)
dtMM% = month(xDate)
dtDD% = Day(xDate)
TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
GoTo Cri_FTDatst


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

Private Sub Cri_DepStatus(xStr)
Dim EECri As String, OneSet%, x%

If UCase(xStr) = "ALL" Then Exit Sub
If Len(xStr) < 1 Then
    EECri = "(isnull({HRDEPEND.DP_STATUS}) or len({HRDEPEND.DP_STATUS})=0)"
Else
    EECri = "{HRDEPEND.DP_STATUS}= '" & xStr & "'"
End If

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

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

Private Sub Cri_ProvResidence()
Dim EECri As String, OneSet%, x%

If Len(clpProv.Text) < 1 Then Exit Sub
EECri = "{HREMP.ED_PROV} in ['" & Replace(clpProv.Text, ",", "','") & "']"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_ProvEmployment()
Dim EECri As String, OneSet%, x%

If Len(clpProvEmp.Text) < 1 Then Exit Sub
EECri = "{HREMP.ED_PROVEMP} in ['" & Replace(clpProvEmp.Text, ",", "','") & "']"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Function Cri_SetAll()
Dim x%, strRName$
Dim SQLQ As String, strCriteria As String
Dim rsIn As New ADODB.Recordset, rsOUT As New ADODB.Recordset

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
For x% = 0 To 5
    Call Cri_Code(x%)
Next x%
Call Cri_PT
Call Cri_Shift
Call Cri_EE
Call Cri_ProvResidence
Call Cri_ProvEmployment

'Ticket #14231
If frmDepend.Visible Then
    Dim xDepStatus As String
    Dim xFromDate As String, xToDate As String
    xDepStatus = GetStrBetweenBrackets(ComStatus.Text)
    Call Cri_DepStatus(xDepStatus)
    If IsNumeric(medDOW(0).Text) Then
        xFromDate = DateAdd("YYYY", -medDOW(0).Text, dlpAsDate.Text)
        Call Cri_DepFromDate(xFromDate)
    End If
    If IsNumeric(medDOW(1).Text) Then
        xToDate = DateAdd("YYYY", -medDOW(1).Text, dlpAsDate.Text)
        Call Cri_DepToDate(xToDate)
    End If
    Call Cri_Country
End If

If FormHomeAddress% = True Then
    If Not (glbSQL Or glbOracle) And chkIncludeTerm Then
        SQLQ = "DELETE FROM HR_HRLABEL_WRK WHERE ED_WRKEMP='" & glbUserID & "'"
        gdbAdoIhr001W.Execute SQLQ
        
        strCriteria = Replace(Replace(Replace(Replace(glbstrSelCri, "{", "("), "}", ")"), "[", "("), "]", ")")
        SQLQ = "SELECT ED_COMPNO, ED_EMPNBR, ED_SURNAME, ED_FNAME, ED_EMPTYPE, ED_ADDR1, ED_ADDR2, ED_CITY, ED_PCODE, ED_PROV, ED_BUSNBR, ED_PHONE, "
        SQLQ = SQLQ & "ED_CELLPHONE, ED_COUNTRY,ED_TITLE , ED_DOH, ED_DIV, ED_DEPTNO, ED_LOC, ED_SECTION, ED_EMP,ED_REGION "
        SQLQ = SQLQ & "FROM HREMP WHERE " & strCriteria
        rsIn.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
        If rsIn.EOF = False And rsIn.BOF = False Then
            rsOUT.Open "HR_HRLABEL_WRK", gdbAdoIhr001W, adOpenDynamic, adLockOptimistic, adCmdTable
            Do
                rsOUT.AddNew
                    rsOUT("ED_COMPNO") = rsIn("ED_COMPNO")
                    rsOUT("ED_EMPNBR") = rsIn("ED_EMPNBR")
                    rsOUT("ED_SURNAME") = rsIn("ED_SURNAME")
                    rsOUT("ED_FNAME") = rsIn("ED_FNAME")
                    rsOUT("ED_EMPTYPE") = rsIn("ED_EMPTYPE")
                    rsOUT("ED_ADDR1") = rsIn("ED_ADDR1")
                    rsOUT("ED_ADDR2") = rsIn("ED_ADDR2")
                    rsOUT("ED_CITY") = rsIn("ED_CITY")
                    rsOUT("ED_PCODE") = rsIn("ED_PCODE")
                    rsOUT("ED_PROV") = rsIn("ED_PROV")
                    rsOUT("ED_BUSNBR") = rsIn("ED_BUSNBR")
                    rsOUT("ED_PHONE") = rsIn("ED_PHONE")
                    rsOUT("ED_CELLPHONE") = rsIn("ED_CELLPHONE")
                    rsOUT("ED_COUNTRY") = rsIn("ED_COUNTRY")
                    rsOUT("ED_TITLE") = rsIn("ED_TITLE")
                    rsOUT("ED_DOH") = rsIn("ED_DOH")
                    rsOUT("ED_DIV") = rsIn("ED_DIV")
                    rsOUT("ED_DEPTNO") = rsIn("ED_DEPTNO")
                    rsOUT("ED_LOC") = rsIn("ED_LOC")
                    rsOUT("ED_SECTION") = rsIn("ED_SECTION")
                    rsOUT("ED_EMP") = rsIn("ED_EMP")
                    rsOUT("ED_REGION") = rsIn("ED_REGION")
                    rsOUT("ED_WRKEMP") = glbUserID
                    
                rsOUT.Update
                rsIn.MoveNext
            Loop Until rsIn.EOF
            rsOUT.Close
        End If
        rsIn.Close
        
        strCriteria = Replace(strCriteria, "HREMP.", "Term_HREMP.")
        SQLQ = "SELECT ED_COMPNO, ED_EMPNBR, ED_SURNAME, ED_FNAME, ED_EMPTYPE, ED_ADDR1, ED_ADDR2, ED_CITY, ED_PCODE, ED_PROV, ED_BUSNBR, ED_PHONE, "
        SQLQ = SQLQ & "ED_CELLPHONE, ED_COUNTRY,ED_TITLE , ED_DOH, ED_DIV, ED_DEPTNO, ED_LOC, ED_SECTION, ED_EMP,ED_REGION "
        SQLQ = SQLQ & "FROM Term_HREMP WHERE " & strCriteria
        rsIn.Open SQLQ, gdbAdoIhr001X, adOpenStatic, adLockOptimistic, adCmdText
        If rsIn.EOF = False And rsIn.BOF = False Then
            rsOUT.Open "HR_HRLABEL_WRK", gdbAdoIhr001W, adOpenDynamic, adLockOptimistic, adCmdTable
            Do
                rsOUT.AddNew
                    rsOUT("ED_COMPNO") = rsIn("ED_COMPNO")
                    rsOUT("ED_EMPNBR") = rsIn("ED_EMPNBR")
                    rsOUT("ED_SURNAME") = rsIn("ED_SURNAME")
                    rsOUT("ED_FNAME") = rsIn("ED_FNAME")
                    rsOUT("ED_EMPTYPE") = rsIn("ED_EMPTYPE")
                    rsOUT("ED_ADDR1") = rsIn("ED_ADDR1")
                    rsOUT("ED_ADDR2") = rsIn("ED_ADDR2")
                    rsOUT("ED_CITY") = rsIn("ED_CITY")
                    rsOUT("ED_PCODE") = rsIn("ED_PCODE")
                    rsOUT("ED_PROV") = rsIn("ED_PROV")
                    rsOUT("ED_BUSNBR") = rsIn("ED_BUSNBR")
                    rsOUT("ED_PHONE") = rsIn("ED_PHONE")
                    rsOUT("ED_CELLPHONE") = rsIn("ED_CELLPHONE")
                    rsOUT("ED_COUNTRY") = rsIn("ED_COUNTRY")
                    rsOUT("ED_TITLE") = rsIn("ED_TITLE")
                    rsOUT("ED_DOH") = rsIn("ED_DOH")
                    rsOUT("ED_DIV") = rsIn("ED_DIV")
                    rsOUT("ED_DEPTNO") = rsIn("ED_DEPTNO")
                    rsOUT("ED_LOC") = rsIn("ED_LOC")
                    rsOUT("ED_SECTION") = rsIn("ED_SECTION")
                    rsOUT("ED_EMP") = rsIn("ED_EMP")
                    rsOUT("ED_REGION") = rsIn("ED_REGION")
                    rsOUT("ED_WRKEMP") = glbUserID
                rsOUT.Update
                rsIn.MoveNext
            Loop Until rsIn.EOF
            rsOUT.Close
        End If
        rsIn.Close
    
    End If
End If

' report name
If FormHomeAddress% = True Then   'laura Oct 28, 1997
  ' danielk - 03/19/2003 - added Date of Hire as a sort
  
    If chkIncludeTerm Then
        If glbLinamar Then
            If comGroup(0) <> "(none)" Then
                strRName$ = glbIHRREPORTS & "rlhomeadti.rpt"
            Else
                strRName$ = glbIHRREPORTS & "rlhomad1ti.rpt"
            End If
        Else
            If comGroup(0) <> "(none)" Then
                'If comGroup(1) = "Employee Name" Or comGroup(1) = "Employee Number" Then
                    '  strRName$ = glbIHRREPORTS & "rzhomeadti.rpt"
                'Else
                    '  strRName$ = glbIHRREPORTS & "rzhomeahti.rpt"
                'End If
                strRName$ = glbIHRREPORTS & "rzhomeadti.rpt"
            Else
                'If comGroup(1) = "Employee Name" Or comGroup(1) = "Employee Number" Then
                    '  strRName$ = glbIHRREPORTS & "rzhomad1ti.rpt"
                'Else
                    '  strRName$ = glbIHRREPORTS & "rzhomah1ti.rpt"
                'End If
                strRName$ = glbIHRREPORTS & "rzhomad1ti.rpt"
            End If
        End If
    ElseIf chkOnlyTerm Then
        If glbLinamar Then
            If comGroup(0) <> "(none)" Then
                strRName$ = glbIHRREPORTS & "rlhomeadt.rpt"
            Else
                strRName$ = glbIHRREPORTS & "rlhomad1t.rpt"
            End If
        Else
            If comGroup(0) <> "(none)" Then
                If comGroup(1) = "Employee Name" Or comGroup(1) = "Employee Number" Then
                    strRName$ = glbIHRREPORTS & "rzhomeadt.rpt"
                Else
                    strRName$ = glbIHRREPORTS & "rzhomeaht.rpt"
                End If
            Else
                If comGroup(1) = "Employee Name" Or comGroup(1) = "Employee Number" Then
                    strRName$ = glbIHRREPORTS & "rzhomad1t.rpt"
                Else
                    strRName$ = glbIHRREPORTS & "rzhomah1t.rpt"
                End If
            End If
        End If
    Else
        If glbLinamar Then
            If comGroup(0) <> "(none)" Then
                strRName$ = glbIHRREPORTS & "rlhomead.rpt"
            Else
                strRName$ = glbIHRREPORTS & "rlhomad1.rpt"
            End If
        Else
            If comGroup(0) <> "(none)" Then
                If comGroup(1) = "Employee Name" Or comGroup(1) = "Employee Number" Then
                    strRName$ = glbIHRREPORTS & "rzhomead.rpt"
                Else
                    strRName$ = glbIHRREPORTS & "rzhomeah.rpt"
                End If
            Else
                If comGroup(1) = "Employee Name" Or comGroup(1) = "Employee Number" Then
                    strRName$ = glbIHRREPORTS & "rzhomad1.rpt"
                Else
                    strRName$ = glbIHRREPORTS & "rzhomah1.rpt"
                End If
            End If
        End If
    End If
  
    Me.vbxCrystal.ReportFileName = strRName$
    
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
        Me.vbxCrystal.SelectionFormula = glbstrSelCri
    End If
  
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        If chkOnlyTerm Then 'Terminated ONLY Employee
            For x% = 0 To 4
                If x% = 2 Then
                    Me.vbxCrystal.DataFiles(x%) = glbIHRAUDIT
                Else
                    Me.vbxCrystal.DataFiles(x%) = glbIHRDB
                End If
            Next x%
        ElseIf chkIncludeTerm Then 'Include Terminated
            For x% = 0 To 4
                If x% = 2 Then
                    Me.vbxCrystal.DataFiles(x%) = glbIHRDBW
                Else
                    Me.vbxCrystal.DataFiles(x%) = glbIHRDB
                End If
            Next x%
        Else
            Me.vbxCrystal.Connect = RptODBC_SQL
        End If
    End If

    ' window title if appropriate
    Me.vbxCrystal.WindowTitle = "Home Address/Telelphone Report"
    'Ticket #25935 08/28/2014 - begin
    If Not glbLinamar Then
        Me.vbxCrystal.Formulas(51) = "lblPhone2='" & lStr("Telephone #2") & "'"
        Me.vbxCrystal.Formulas(52) = "lblCell='" & lStr("Cellular Telephone") & "'"
        Me.vbxCrystal.Formulas(53) = "lblPager='" & lStr("Pager Number") & "'"
    End If
    'Ticket #25935 08/28/2014 - end
ElseIf FormDepend% = True Then  'Laura Oct 28, 1997
'~~~~~~~~~~
    If chkShowEmp.Value Then 'Ticket #14328 Frank 1/31/08
        If comGroup(0) <> "(none)" Then
            strRName$ = glbIHRREPORTS & "rzdependa.rpt"
        Else
            strRName$ = glbIHRREPORTS & "rzdep1a.rpt"
        End If
    Else
        If comGroup(0) <> "(none)" Then
            strRName$ = glbIHRREPORTS & "rzdepend.rpt"
        Else
            strRName$ = glbIHRREPORTS & "rzdep1.rpt"
        End If
    End If
    Me.vbxCrystal.ReportFileName = strRName$
    
    ' set to sorting/grouping criteria
    x% = Cri_SortsDepend()    ' returns number of sections formated
    
    'set location for database tables
    If Len(glbstrSelCri) >= 0 Then
        Me.vbxCrystal.SelectionFormula = glbstrSelCri
    End If
    Me.vbxCrystal.Connect = RptODBC_SQL

    ' window title if appropriate
    Me.vbxCrystal.WindowTitle = "Dependents Report"
'~~~~~~~~~~~~~
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
        Case 2: grpField$ = "{HREMP.ED_DOH}"
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
    Case 2: grpField$ = "{HREMP.ED_DOH}"
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
If glbCompSerial = "S/N - 2375W" Then   'TIMMIS
    If chkComments Then
        Me.vbxCrystal.Formulas(101) = "ShowComments=true"
    Else
        Me.vbxCrystal.Formulas(101) = "ShowComments=false"
    End If
End If
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

GrpIdx% = comGroup(1).ListIndex
Select Case GrpIdx%
    Case 0: grpField$ = "{@EFullName}"
    Case 1: grpField$ = "{HREMP.ED_EMPNBR}" 'GROUP ON EMPLOYEE#
End Select
grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(1) = grpCond$

Cri_SortsDepend = z% ' next section number to format

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
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
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

'Ticket #14231
If frmDepend.Visible Then
    If Len(medDOW(0).Text) > 0 Or Len(medDOW(1).Text) > 0 Then
        If Not IsDate(dlpAsDate.Text) Then
            MsgBox "As of Date is required if Age Range entered."
            dlpAsDate.SetFocus
            Exit Function
        End If
    End If
    If IsDate(dlpAsDate.Text) Then
        If Len(medDOW(0).Text) = 0 And Len(medDOW(1).Text) = 0 Then
            MsgBox "Age Range is required if As of Date entered."
            medDOW(0).SetFocus
            Exit Function
        End If
    End If
End If

If clpProv.Caption = "Unassigned" Then
    MsgBox "Invalid Prov. of Residence"
    clpProv.SetFocus
    Exit Function
End If

If clpProvEmp.Caption = "Unassigned" Then
    MsgBox "Invalid Prov. of Employment"
    clpProvEmp.SetFocus
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

glbOnTop = "FRMRHOMEA"

Screen.MousePointer = HOURGLASS

If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If

Call comGrpLoad
Call setRptCaption(Me)

If FormHomeAddress% = True Then
   frmRHomeA.Caption = "Home Address/Telephone Report"
   chkIncludeTerm.Visible = True
   chkOnlyTerm.Visible = True
   
    'Ticket #29660 - Contract Employees Enhancement
    If glbWFC Then
        chkExclCONP.Visible = True
        chkExclRET.Visible = True
    Else
        chkExclCONP.Visible = False
        chkExclRET.Visible = False
    End If
Else
   frmRHomeA.Caption = "Dependents Report"
   chkIncludeTerm.Visible = False
   chkOnlyTerm.Visible = False
   'Ticket #14231 - Begin
   Call addCountryItems
   If glbWFC Then
        'comCountry.Text = "CANADA"
        comCountry.Text = "ALL" 'Ticket #25352 Franks 04/22/2014
   End If
   Call DependStatus
   frmDepend.Top = 4660
   frmDepend.Left = 120
   frmGroup.Top = frmGroup.Top + 900
   frmDepend.Visible = True
   'Ticket #14231 - End
End If

If glbLinamar Then clpCode(3).MaxLength = 8
If glbCompSerial = "S/N - 2227W" Then clpCode(3).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6

If FormDepend% Then
    If glbCompSerial = "S/N - 2375W" Then
        chkComments.Visible = True
    End If
End If

Call INI_Controls(Me)

Screen.MousePointer = DEFAULT

End Sub

Private Sub addCountryItems()
Dim ctylist, x

ctylist = CountryList
x = 1
Do While x > 0
    x = InStr(ctylist, "&")
    If x > 0 Then
        comCountry.AddItem Left(ctylist, x - 1)
        ctylist = Mid(ctylist, x + 1)
    Else
        comCountry.AddItem ctylist
    End If
Loop

End Sub

Private Sub DependStatus()
ComStatus.Clear
If glbWFC And comCountry.Text = "CANADA" Then 'Ticket #13448
    ComStatus.AddItem "ALL"
    ComStatus.AddItem "Other( )" '"Child( )"
    ComStatus.AddItem "(S)"
    ComStatus.AddItem "Handicap(H)"
    ComStatus.AddItem "Twin/Multiple Birth(M)"
Else
    ComStatus.AddItem "ALL"
    ComStatus.AddItem "Other( )"
    ComStatus.AddItem "Student(1)"
    ComStatus.AddItem "Disability(2)"
    ComStatus.AddItem "Fast Yes(Y)"
    ComStatus.AddItem "Fast No(N)"
    ComStatus.AddItem "Fast(3)"
    ComStatus.AddItem "Fast(S)"
End If
End Sub

Private Function GetStrBetweenBrackets(xStr)
Dim X1Pos As Integer, X2Pos As Integer
Dim xRetVal As String
    xRetVal = xStr
    X1Pos = RetPosition("(", xStr)
    X2Pos = RetPosition(")", xStr)
    If X1Pos > 0 And X2Pos > 0 And X2Pos > X1Pos Then
        xRetVal = Mid(xStr, X1Pos + 1, X2Pos - X1Pos - 1)
    End If
    GetStrBetweenBrackets = Trim(xRetVal)
End Function

Private Function RetPosition(xChar, xStr)
Dim xPos As Integer
Dim I, xlen
    xPos = 0
    xlen = Len(xStr)
    For I = 1 To xlen
        If Mid(xStr, I, 1) = xChar Then
            xPos = I
            GoTo EndRec
        End If
    Next
EndRec:
    RetPosition = xPos
End Function

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

Private Function CountryList() As String
Dim xCountryList As String, ctyFile
xCountryList = ""
ctyFile = glbIHRREPORTS & "CountryList.MTF"

On Error GoTo ErrorHandler

If File(ctyFile) Then
    Open ctyFile For Input As #1
    Input #1, xCountryList
    Close #1
End If

ResumeHere:
'If InStr(xCountryList, BasicCountry) = 0 Then
'    xCountryList = BasicCountry
'End If
If InStr(xCountryList, comCountry) = 0 And comCountry <> "" Then
    xCountryList = xCountryList & "&" & comCountry
    comCountry.AddItem comCountry
End If
Open ctyFile For Output As #1
Print #1, xCountryList
Close #1
CountryList = xCountryList
Exit Function

ErrorHandler:
If Err.Number = 62 Then
    ' Corrupted CountryList.MTF, kill it and regenerate
    Close #1
    MsgBox "Found corrupt CountryList.MTF.  info:HR will re-create this file.", vbInformation + vbOKOnly, "Corrupted Country List"
    Kill ctyFile
    Resume ResumeHere
Else
    'MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number & " in CountryList"
    Resume Next
End If
End Function

Private Sub Cri_Country()
Dim CountryCri As String

If UCase(comCountry.Text) = "ALL" Then Exit Sub

If Len(comCountry.Text) > 0 Then
    CountryCri = "({HREMP.ED_COUNTRY} = '" & comCountry.Text & "')"
End If

If Len(CountryCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = CountryCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & CountryCri
    End If
    glbiOneWhere = True
End If
End Sub
