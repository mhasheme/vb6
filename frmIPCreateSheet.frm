VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmIPCreateSheet 
   Caption         =   "Create Incentive Plan Spreadsheet"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16590
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9315
   ScaleWidth      =   16590
   WindowState     =   2  'Maximized
   Begin VB.Frame fraEmpLetter 
      Height          =   6015
      Left            =   12720
      TabIndex        =   55
      Top             =   0
      Visible         =   0   'False
      Width           =   5295
      Begin VB.Frame fraTypeRpt 
         Height          =   495
         Left            =   2400
         TabIndex        =   78
         Top             =   4800
         Visible         =   0   'False
         Width           =   4575
         Begin VB.OptionButton OptSum 
            Caption         =   "Summary - Dollars Only"
            Height          =   195
            Left            =   0
            TabIndex        =   80
            Top             =   120
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton OptFullPlan 
            Caption         =   "Full Plan"
            Height          =   195
            Left            =   2400
            TabIndex        =   79
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.CommandButton cmdClose2 
         Appearance      =   0  'Flat
         Caption         =   "Close"
         Height          =   375
         Left            =   4680
         TabIndex        =   65
         Tag             =   "Print all Vacation Entitlement Report"
         Top             =   5520
         Width           =   1305
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Create"
         Height          =   375
         Left            =   2280
         TabIndex        =   64
         Top             =   5520
         Width           =   1425
      End
      Begin VB.ComboBox comGroup 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   2355
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Tag             =   "Second level of grouping records"
         Top             =   4365
         Width           =   2325
      End
      Begin VB.ComboBox comGroup 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   2355
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Tag             =   "First Level of grouping records"
         Top             =   4050
         Width           =   2325
      End
      Begin VB.TextBox txtComment 
         Appearance      =   0  'Flat
         Height          =   975
         Left            =   2355
         MaxLength       =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   62
         Tag             =   "00-Local Message"
         Top             =   2400
         Width           =   7245
      End
      Begin MSMask.MaskEdBox MskFiscalYeaLetter 
         Height          =   315
         Left            =   2355
         TabIndex        =   56
         Tag             =   "01-High Dollars"
         Top             =   210
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0"
         PromptChar      =   "_"
      End
      Begin INFOHR_Controls.CodeLookup clpDiv2 
         Height          =   285
         Left            =   2040
         TabIndex        =   57
         Top             =   600
         Width           =   7635
         _ExtentX        =   13467
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
         Left            =   2040
         TabIndex        =   58
         Tag             =   "00-Enter Section Code"
         Top             =   960
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   6
         Left            =   2040
         TabIndex        =   59
         Tag             =   "00-Enter Region Code"
         Top             =   1320
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDRG"
      End
      Begin INFOHR_Controls.EmployeeLookup elpEEID2 
         Height          =   285
         Left            =   2040
         TabIndex        =   60
         Tag             =   "10-Enter Employee Number"
         Top             =   1680
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         TextBoxWidth    =   7195
         RefreshDescriptionWhen=   2
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.DateLookup dlpPayDate 
         Height          =   285
         Left            =   2040
         TabIndex        =   61
         Tag             =   "40-As of Date"
         Top             =   2040
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin VB.Label lblTypeRpt 
         AutoSize        =   -1  'True
         Caption         =   "Type of Report"
         Height          =   195
         Left            =   120
         TabIndex        =   77
         Top             =   4920
         Width           =   1425
      End
      Begin VB.Label lblRepGrp 
         BackStyle       =   0  'Transparent
         Caption         =   "Groupings"
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
         Left            =   120
         TabIndex        =   76
         Top             =   3840
         Width           =   1575
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
         Left            =   180
         TabIndex        =   75
         Top             =   4395
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
         Left            =   180
         TabIndex        =   74
         Top             =   4080
         Width           =   885
      End
      Begin VB.Label lblComment 
         AutoSize        =   -1  'True
         Caption         =   "Local Message"
         Height          =   195
         Left            =   120
         TabIndex        =   73
         Top             =   2400
         Width           =   1080
      End
      Begin VB.Label lblPayDate 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Scheduled Payment Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   72
         Top             =   2085
         Width           =   1815
      End
      Begin VB.Label lblEENum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Number"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   71
         Top             =   1725
         Width           =   1290
      End
      Begin VB.Label lblRegion2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Region"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   70
         Top             =   1365
         Width           =   1095
      End
      Begin VB.Label lblSectio2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   69
         Top             =   1005
         Width           =   1260
      End
      Begin VB.Label lblDiv2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   68
         Top             =   645
         Width           =   915
      End
      Begin VB.Label lblFiscalYea3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Incentive Letter for Year"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   67
         Top             =   270
         Width           =   1710
      End
   End
   Begin VB.Frame fraSunLife 
      Height          =   4335
      Left            =   120
      TabIndex        =   32
      Top             =   4440
      Visible         =   0   'False
      Width           =   6615
      Begin VB.OptionButton cptExc2003 
         Caption         =   "Excel 2003"
         Height          =   195
         Left            =   2280
         TabIndex        =   49
         Top             =   3960
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton cptExc2010 
         Caption         =   "Excel 2010"
         Height          =   195
         Left            =   2280
         TabIndex        =   48
         Top             =   3600
         Width           =   2295
      End
      Begin VB.CommandButton cmdCreate 
         Caption         =   "Create File"
         Height          =   375
         Left            =   1680
         TabIndex        =   37
         Top             =   3000
         Width           =   1425
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "Close"
         Height          =   375
         Left            =   3360
         TabIndex        =   36
         Tag             =   "Print all Vacation Entitlement Report"
         Top             =   3000
         Width           =   1305
      End
      Begin VB.DriveListBox Drive1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1680
         TabIndex        =   34
         Tag             =   "Disk Drive"
         Top             =   240
         Width           =   3372
      End
      Begin VB.DirListBox Dir1 
         BackColor       =   &H00FFFFFF&
         Height          =   2115
         Left            =   1680
         TabIndex        =   33
         Tag             =   "Path"
         Top             =   600
         Width           =   3372
      End
      Begin VB.Label lblPath 
         BackStyle       =   0  'Transparent
         Caption         =   "Copy to Path"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   1260
      End
   End
   Begin VB.Frame fraToPayroll 
      Height          =   1695
      Left            =   7320
      TabIndex        =   50
      Top             =   4680
      Visible         =   0   'False
      Width           =   3135
      Begin VB.ComboBox comEarnCode 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4890
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Tag             =   "00-Country"
         Top             =   960
         Width           =   840
      End
      Begin VB.ComboBox comCountry 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4890
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Tag             =   "00-Country"
         Top             =   480
         Width           =   1320
      End
      Begin MSMask.MaskEdBox MskFiscalYea2 
         Height          =   315
         Left            =   1500
         TabIndex        =   28
         Tag             =   "01-High Dollars"
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0"
         PromptChar      =   "_"
      End
      Begin INFOHR_Controls.CodeLookup clpVadim2 
         Height          =   285
         Left            =   1200
         TabIndex        =   30
         Top             =   960
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDV2"
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Earning Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   3840
         TabIndex        =   54
         Top             =   1020
         Width           =   960
      End
      Begin VB.Label lblVadim2 
         AutoSize        =   -1  'True
         Caption         =   "Vadim Field 2"
         Height          =   195
         Left            =   120
         TabIndex        =   53
         Top             =   1020
         Width           =   945
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Country"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   19
         Left            =   3840
         TabIndex        =   52
         Top             =   510
         Width           =   660
      End
      Begin VB.Label lblFiscalYea2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fiscal Year"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   51
         Top             =   510
         Width           =   780
      End
   End
   Begin VB.Frame fraImportSheet 
      Height          =   3255
      Left            =   7200
      TabIndex        =   39
      Top             =   5520
      Visible         =   0   'False
      Width           =   8175
      Begin VB.TextBox txtFileName 
         Height          =   315
         Left            =   1800
         TabIndex        =   41
         Tag             =   "00-File Name (Include Extension TXT)"
         Top             =   1320
         Width           =   4905
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   290
         Left            =   6840
         TabIndex        =   45
         Tag             =   "Click to select the location"
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton cmdClos2 
         Appearance      =   0  'Flat
         Caption         =   "Close"
         Height          =   375
         Left            =   3360
         TabIndex        =   44
         Tag             =   "Print all Vacation Entitlement Report"
         Top             =   2520
         Width           =   1065
      End
      Begin VB.CommandButton cmdImp 
         Caption         =   "Import File"
         Height          =   375
         Left            =   1800
         TabIndex        =   43
         Top             =   2520
         Width           =   1065
      End
      Begin MSMask.MaskEdBox MskFiscalYear 
         Height          =   315
         Left            =   1800
         TabIndex        =   40
         Tag             =   "01-High Dollars"
         Top             =   600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0"
         PromptChar      =   "_"
      End
      Begin MSComDlg.CommonDialog AttachmentDialog 
         Left            =   7680
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "*.xls;*.xlsx"
      End
      Begin VB.Label lblCriteria 
         BackStyle       =   0  'Transparent
         Caption         =   "Import From"
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   46
         Top             =   1380
         Width           =   1260
      End
      Begin VB.Label lblFiscalYear 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fiscal Year"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   480
         TabIndex        =   42
         Top             =   630
         Width           =   780
      End
   End
   Begin VB.ComboBox ComMTH 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Tag             =   "01-Month"
      Top             =   4050
      Width           =   2655
   End
   Begin VB.ComboBox comQ1 
      Height          =   315
      Left            =   5280
      TabIndex        =   0
      Tag             =   "00-Yes/No"
      Top             =   200
      Width           =   795
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   990
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
      Left            =   1560
      TabIndex        =   9
      Tag             =   "00-Enter Section Code"
      Top             =   3360
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
      TabIndex        =   8
      Tag             =   "00-Enter Administered By Code"
      Top             =   3030
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
      Top             =   2700
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   3
      Tag             =   "00-Enter Location Code"
      Top             =   1650
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Tag             =   "00-Specific Department Desired"
      Top             =   1320
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
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Tag             =   "10-Enter Employee Number"
      Top             =   1990
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpJob 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Tag             =   "00-Enter Position Code "
      Top             =   2350
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   25
      LookupType      =   5
   End
   Begin INFOHR_Controls.CodeLookup clpJobMaster 
      Height          =   285
      Left            =   6600
      TabIndex        =   6
      Tag             =   "01-Job code"
      Top             =   2350
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   25
      LookupType      =   13
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "JB_POSTYPE"
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   10
      Tag             =   "00-Position Type Code"
      Top             =   3690
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "POTY"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.DateLookup dlpAsOf 
      Height          =   285
      Left            =   6600
      TabIndex        =   13
      Tag             =   "40-As of Date"
      Top             =   4080
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin VB.TextBox txtMTH 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   315
      Left            =   4080
      MaxLength       =   3
      TabIndex        =   26
      Text            =   "MTH"
      Top             =   4080
      Visible         =   0   'False
      Width           =   570
   End
   Begin INFOHR_Controls.DateLookup dlpSalAsOf 
      Height          =   285
      Left            =   6600
      TabIndex        =   12
      Tag             =   "40-As of Date"
      Top             =   3720
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin VB.Label lblSalAsOf 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Salary As of Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5160
      TabIndex        =   47
      Top             =   3765
      Width           =   1230
   End
   Begin VB.Label lblAsOf 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Service As of Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5160
      TabIndex        =   38
      Top             =   4125
      Width           =   1335
   End
   Begin VB.Label lblRate 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Exchange Rate to Use"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   27
      Top             =   4080
      Width           =   1620
   End
   Begin VB.Label lblPosType 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Type"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   25
      Top             =   3720
      Width           =   1140
   End
   Begin VB.Label lblPosition 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   2390
      Width           =   975
   End
   Begin VB.Label lblJob 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Job Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   23
      Top             =   2390
      Width           =   1035
   End
   Begin VB.Label lblQ1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Do the Company Incentive Factors Need to be Updated/Reviewed?"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   22
      Top             =   240
      Width           =   4845
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
      TabIndex        =   21
      Top             =   1035
      Width           =   555
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
      TabIndex        =   20
      Top             =   1365
      Width           =   825
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
      TabIndex        =   19
      Top             =   2055
      Width           =   1290
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
      TabIndex        =   18
      Top             =   600
      Width           =   1575
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
      TabIndex        =   17
      Top             =   1695
      Width           =   615
   End
   Begin VB.Label lblAdmin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Administered By"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   3075
      Width           =   1125
   End
   Begin VB.Label lblRegion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   2745
      Width           =   855
   End
   Begin VB.Label lblSection 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   3405
      Width           =   1260
   End
End
Attribute VB_Name = "frmIPCreateSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public wdAppMain As Object
Public wrdDocMain As Object
Public wdApp As Object 'As Word.Application
Public wrdDoc As Object
Public docrange As Object
Dim ImportFile
Dim ImpSucceeded As Boolean
Dim xFound As Boolean
Dim locBandCurr, locBandPrev
Dim xlocFDate, xlocTDate
Dim xPGList(50)
Dim xPGEarnCode(50)
Dim totPG As Integer
    
Private Sub cmdBrowse_Click()
AttachmentDialog.DialogTitle = "Select the file to import..."
AttachmentDialog.Filter = "*.xls;*.xlsx|*.xls;*.xlsx"
AttachmentDialog.FilterIndex = 1
AttachmentDialog.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
AttachmentDialog.ShowOpen
If Len(AttachmentDialog.FileName) <> 0 Then
    txtFileName.Text = AttachmentDialog.FileName
End If
End Sub

Private Sub cmdClos2_Click()
Unload Me
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdClose2_Click()
Unload Me
End Sub

Private Sub cmdCreate_Click()
If glbWFC_IPPopFormName = "WFCIPSpreadSheetCreate" Then
    Call CreateSpreadsheet
End If
If glbWFC_IPPopFormName = "WFCIPPreparePayroll" Then
    Call CreatePreparePayrollFiles
End If
End Sub

Private Sub CreateEmpLetter()
Dim rsEPos As New ADODB.Recordset
Dim rsTermEmp As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset
Dim SQLQ
Dim xlsFileTmp As String
Dim xlsFileMat As String
Dim xlsFileTmpl1 As String
Dim xlsFileTmplM As String
Dim xSaveFileName As String
Dim xlsFileMain As String
Dim I, J, K, M, totNum
Dim xYear, xFName
Dim X%
Dim strWHand As String
Dim Msg As String, a%
Dim locPath, xLocPath
Dim tFlag
Dim xEmpNo, xRep1EmpNo
Dim xROIC, xROICDesc, xTemp, xCURRENCY, xTITLE
Dim xSalary
Dim xEmpFile(5000)
Dim xTotFile As Integer
Dim xIsROIC As Boolean
Dim xTmpMsg
Dim xLocMsg
Dim sSQLQ

On Error GoTo CRW_Err

    If chkEmpLetter() Then
        Msg = "Are You Sure You Want To Create Employee Incentive Letter? "
        
        a% = MsgBox(Msg, 36, "Confirm")
        If a% <> 6 Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If

    locPath = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\")
    'locPath = Dir1.Path & IIf(Right(Dir1.Path, 1) = "\", "", "\")
    'Get the Word Template
    If OptFullPlan.Value Then 'Portuguese
        xlsFileTmpl1 = locPath & "WFC_IncentiveLetterTemplateP.dot"
    Else
        xlsFileTmpl1 = locPath & "WFC_IncentiveLetterTemplate.dot"
    End If
    
    If Dir(xlsFileTmpl1) = "" Then
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(0).FloodPercent = 0
        MDIMain.panHelp(0).Caption = ""
        Screen.MousePointer = DEFAULT
        MsgBox "There is no " & xlsFileTmpl1
        Exit Sub
    End If
    
    'create employee letter ---------------- begin
    
    xYear = MskFiscalYeaLetter.Text
    Call getDateRange(xYear)
    
    SQLQ = "SELECT HREARN.*, ED_PAYROLL_ID,ED_SECTION, ED_SURNAME,ED_FNAME,ED_TITLE,ED_SEX  FROM HREARN LEFT JOIN HREMP ON HREARN.EMPNBR = HREMP.ED_EMPNBR WHERE EARN_TYPE = 'BON4'  "
    SQLQ = SQLQ & "AND FDATE = " & Date_SQL(xlocFDate) & " "
    SQLQ = SQLQ & "AND TDATE = " & Date_SQL(xlocTDate) & " "
    
    If Len(glbPlantCode) > 0 Then
        SQLQ = SQLQ & "AND ED_SECTION = '" & glbPlantCode & "' "
    End If

    '---- check HREMP table --- begin
    If Len(clpDiv2.Text) > 0 Then
        SQLQ = SQLQ & "AND ED_DIV IN ('" & Replace(clpDiv2.Text, ",", "','") & "') "
    End If
    If Len(clpCode(5).Text) > 0 Then
        SQLQ = SQLQ & "AND ED_SECTION IN ('" & Replace(clpCode(5).Text, ",", "','") & "') "
    End If
    If Len(clpCode(6).Text) > 0 Then
        SQLQ = SQLQ & "AND ED_REGION IN ('" & Replace(clpCode(6).Text, ",", "','") & "') "
    End If
    If Len(elpEEID2.Text) > 0 Then
        SQLQ = SQLQ & "AND ED_EMPNBR IN (" & getEmpnbr(elpEEID2.Text) & ") "
    End If
    '---- check HREMP table --- end
    If comGroup(0).Text = lStr("Section") Then 'by plant, employee, date
        SQLQ = SQLQ & "ORDER BY ED_SECTION, ED_SURNAME,ED_FNAME"
    ElseIf comGroup(0).Text = lStr("Region") Then
        SQLQ = SQLQ & "ORDER BY ED_REGION, ED_SURNAME,ED_FNAME"
    ElseIf comGroup(0).Text = lStr("Division") Then
        SQLQ = SQLQ & "ORDER BY ED_DIV, ED_SURNAME,ED_FNAME"
    Else
        SQLQ = SQLQ & "ORDER BY  ED_SURNAME,ED_FNAME"
    End If
    
    If rsEmp.State <> 0 Then rsEmp.Close
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsEmp.EOF Then
        MsgBox "No 'BON4' record found in this Selection Criteria "
        Exit Sub
    End If

    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(0).FloodPercent = 0
    
    'Main document ------------------
    xLocPath = Dir1.Path & IIf(Right(Dir1.Path, 1) = "\", "", "\") ' glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\")
    'Create new hidden instance of Word.
    Set wdAppMain = CreateObject("Word.Application")
    tFlag = True
    
    'Get the Word Template
    xlsFileTmplM = locPath & "WFC_IncentiveLetterTemplate_All.dot" ' "WFC_IncentiveLetterTemplate_tmp.dot"
        
    'Filename for the Word Form to Save As
    'xlsFileMain = locPath & "WFC_IncentiveLetter" & xYear & ".doc"
    If Len(clpCode(5).Text) > 0 Then
        xlsFileMain = xLocPath & "WFC_IncentiveLetter" & xYear & "_" & Trim(clpCode(5).Text) & "(" & glbUserID & ").doc"
    Else
        xlsFileMain = xLocPath & "WFC_IncentiveLetter" & xYear & "(" & glbUserID & ").doc"
    End If
    
    'Delete the word document if already exists
    If (Dir(xlsFileMain)) <> "" Then Kill xlsFileMain
    
    Set wrdDocMain = wdAppMain.Documents.Add(xlsFileTmplM, False)
    wdAppMain.Documents(wrdDocMain).Activate

    'Create new hidden instance of Word.
    Set wdApp = CreateObject("Word.Application")
    
    cmdPrint.Enabled = False
    cmdClose2.Enabled = False
    
    I = 0
    totNum = rsEmp.RecordCount
    'xRow = 2
    Do While Not rsEmp.EOF
        If (I / totNum) <= 1 Then
            MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
            I = I + 1
        End If
        DoEvents
        
        xEmpNo = rsEmp("EMPNBR")
        With wdApp
                'xLocPath = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\")
                xFName = "WFC_IncentiveLetter" & xYear & "-" & Trim(Str(I)) & ".doc" ' "WFC Statement of Retire Benefit & Options(" & empName & ").doc"
                xTemp = Trim(rsEmp("ED_FNAME")) & " " & Trim(rsEmp("ED_SURNAME"))
                'Get the Word Template
                If OptFullPlan.Value Then 'Portuguese
                    xFName = "WFC_IPLetter" & xYear & "_" & xEmpNo & "(" & xTemp & ")P.doc"
                Else
                    xFName = "WFC_IPLetter" & xYear & "_" & xEmpNo & "(" & xTemp & ").doc"
                End If
                xEmpFile(I) = xFName
                
                xlsFileMat = xLocPath & "" & xFName

                'delete this file if it exists
                If (Dir(xlsFileMat)) <> "" Then
                    Kill xlsFileMat
                End If
                xSaveFileName = xlsFileMat

                'Set Word object as the template
                Set wrdDoc = .Documents.Add(xlsFileTmpl1, False)
                'Make the word doc Active
                .Documents(wrdDoc).Activate
                
                xRep1EmpNo = ""
                SQLQ = "SELECT HR_JOB_HISTORY.*,JB_BAND FROM HR_JOB_HISTORY LEFT JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE  WHERE JH_EMPNBR = " & xEmpNo & " AND JH_CURRENT = 1 "
                If rsEPos.State <> 0 Then rsEPos.Close
                rsEPos.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsEPos.EOF Then
                    If Not IsNull(rsEPos("JH_REPTAU")) Then
                        xRep1EmpNo = rsEPos("JH_REPTAU")
                    End If
                End If
                
                'Update the bookmark fields in the Word template with database values
                If OptFullPlan.Value Then 'Portuguese
                    .ActiveDocument.FormFields("txtToday").Result = Date
                Else
                    .ActiveDocument.FormFields("txtToday").Result = Format(Date, "MMM dd, YYYY")
                End If
                xTemp = Trim(rsEmp("ED_FNAME")) & " " & Trim(rsEmp("ED_SURNAME"))
                xTITLE = ""
                If Not IsNull(rsEmp("ED_TITLE")) Then
                    If Len(Trim(rsEmp("ED_TITLE"))) > 0 Then
                        xTITLE = Trim(rsEmp("ED_TITLE"))
                    End If
                End If
                If Len(xTITLE) > 0 Then
                    xTemp = xTITLE & " " & xTemp
                Else
                    If Not IsNull(rsEmp("ED_SEX")) Then
                        If rsEmp("ED_SEX") = "M" Then xTemp = "Mr. " & xTemp
                        If rsEmp("ED_SEX") = "F" Then xTemp = "Ms. " & xTemp
                    End If
                End If
                .ActiveDocument.FormFields("txtEmpName").Result = xTemp ' rsEmp("ED_SURNAME") & "," & rsEmp("ED_FNAME")
                If Not rsEPos.EOF Then
                    .ActiveDocument.FormFields("txtPosDesc").Result = getPosDesc(rsEPos("JH_JOB")) ' getPosDesc(getEmpPostion(xEmpNo)) ' "CEO"
                End If
                .ActiveDocument.FormFields("txtPlant").Result = GetTABLCodePub("EDSE", rsEmp("ED_SECTION"))
                .ActiveDocument.FormFields("txtFName").Result = rsEmp("ED_FNAME")
                
                .ActiveDocument.FormFields("txtYear").Result = xYear
                .ActiveDocument.FormFields("txtYear2").Result = xYear + 1
                .ActiveDocument.FormFields("txtYear3").Result = xYear + 1
                'ROIC from the Company Incentive Factors
                xROIC = getEmpROIC(xYear, rsEmp("ED_SECTION"))
                xIsROIC = False
                If Len(xROIC) > 0 Then
                    If xROIC > 1 Then
                        xIsROIC = True
                    End If
                End If
                xLocMsg = Trim(txtComment.Text)
                '.ActiveDocument.FormFields("txtROIC1").Result = xROIC
                '.ActiveDocument.FormFields("txtROIC2").Result = xROIC
                    
                If xIsROIC Then
                    If OptFullPlan.Value Then 'Portuguese
                        xTmpMsg = ""
                        'xROICDesc = Trim(Str(xROIC * 100)) & "%"
                        'xTmpMsg = Chr(13) & Chr(13) & "Woodbridge has achieved a payout of " & xROICDesc & " of target.  Based on financial and individual performance, "
                        'xTmpMsg = xTmpMsg & "Teammates will be allocated their incentive payout on the basis (i.e. Target incentive amount * " & xROIC & ")."
                    
                    Else
                        xROICDesc = Trim(Str(xROIC * 100)) & "%"
                        xTmpMsg = Chr(13) & Chr(13) & "Woodbridge has achieved a payout of " & xROICDesc & " of target.  Based on financial and individual performance, "
                        xTmpMsg = xTmpMsg & "Teammates will be allocated their incentive payout on the basis (i.e. Target incentive amount * " & xROIC & ")."
                    End If

                    'If Len(xLocMsg) > 0 Then
                    '    xTmpMsg = xTmpMsg & Chr(13) & Chr(13) & xLocMsg
                    'End If
                    '.ActiveDocument.FormFields("txtROICMsg").Result = xTmpMsg
                    .ActiveDocument.FormFields("txtExtraMsg").Result = xTmpMsg
                    If Len(xLocMsg) > 0 Then
                        xLocMsg = Chr(13) & Chr(13) & xLocMsg
                        .ActiveDocument.FormFields("txtLocMsg").Result = xLocMsg
                    End If
                Else
                    If OptFullPlan.Value Then 'Portuguese
                        xTmpMsg = "Embroa sua planta não tenha atingido as metas esperadas, o Comite Executivo do WAI garantiu um reconhecimento especial a ser pago baseando-se no sucesso geral da compahia."
                    Else
                        xTmpMsg = "Although your plant did not meet its expected targets, the WAI Executive team has granted "
                        xTmpMsg = xTmpMsg & "a one time special consideration to be paid based on the company's overall success."
                        '.ActiveDocument.FormFields("txtAlthough").Result = xTmpMsg
                    End If
                    'If Len(xLocMsg) > 0 Then
                    '    xTmpMsg = xTmpMsg & Chr(13) & Chr(13) & xLocMsg
                    '    '.ActiveDocument.FormFields("txtROICMsg").Result = xTmpMsg
                    'End If
                    .ActiveDocument.FormFields("txtExtraMsg").Result = xTmpMsg

                    If Len(xLocMsg) > 0 Then
                        xLocMsg = Chr(13) & Chr(13) & xLocMsg
                        .ActiveDocument.FormFields("txtLocMsg").Result = xLocMsg
                    End If
                    
                End If
                

                
                '.ActiveDocument.FormFields("txtLocMessage").Result = Trim(txtComment.Text)
                
                'Other Earnings – BON4 for the Incentive Letter Year
                xCURRENCY = getEmpcurrencyIndi(xEmpNo)
                xTemp = getEmpOtherEarnByCode(xEmpNo, "BON4", xlocFDate, xlocTDate)
                If IsNumeric(xTemp) Then
                    'xTemp = "$" & Format(xTemp, "#,###.##")
                    xTemp = "$" & Format(xTemp, "#,###") & " " & xCURRENCY
                End If
                .ActiveDocument.FormFields("txtIncentive").Result = xTemp ' getEmpOtherEarnByCode(xEmpNo, "BON4", xlocFDate, xlocTDate)
                If OptFullPlan.Value Then 'Portuguese
                    .ActiveDocument.FormFields("txtPayoutDate").Result = dlpPayDate.Text
                Else
                    .ActiveDocument.FormFields("txtPayoutDate").Result = Format(dlpPayDate.Text, "MMM dd, YYYY")
                End If
                'Current Salary * % of Salary: Find the employee's current salary, Band and Market Line. Lookup the Salary Grid table to get the % of salary
                If Not rsEPos.EOF Then
                    If Not IsNull(rsEPos("JB_BAND")) And Not IsNull(rsEmp("ED_SECTION")) Then
                        'get the % of salary
                        xTemp = getSalPercentageByBand(rsEmp("ED_SECTION"), rsEPos("JB_BAND"), xYear)
                        xSalary = GetEmpSalary(xEmpNo)
                        xTemp = xTemp * xSalary
                        If IsNumeric(xTemp) Then
                            'xTemp = "$" & Format(xTemp, "#,###.##")
                            xTemp = "$" & Format(xTemp, "#,###") & " " & xCURRENCY
                        End If
                        .ActiveDocument.FormFields("txtAnnInc").Result = xTemp ' getSalPercentageByBand(rsEmp("ED_SECTION"), rsEPos("JB_BAND"), xYear)
                    End If
                End If

                If Len(xRep1EmpNo) > 0 Then
                    .ActiveDocument.FormFields("txtRA1Name").Result = GetEmpData(xRep1EmpNo, "ED_FNAME") & " " & GetEmpData(xRep1EmpNo, "ED_SURNAME")
                    .ActiveDocument.FormFields("txtRA1Pos").Result = getPosDesc(getEmpPostion(xRep1EmpNo))
                End If
                
                'Save the template as the Word Document - with the filename generated above
                wrdDoc.SaveAs xlsFileMat
                .ActiveDocument.Close
                Set wrdDoc = Nothing
                
                'add to the main file
                Set docrange = wdAppMain.ActiveDocument.Range    'can use Content instead of Range
                docrange.Collapse wdCollapseEnd
                docrange.InsertFile xlsFileMat ' xFName
                If I > 1 Then
                docrange.InsertBreak wdSectionBreakNextPage
                End If
                
                
        End With
        
        rsEmp.MoveNext
    Loop
    
    'term employees
    
'    wdAppMain.NormalTemplate.Saved = True
'    wdAppMain.Quit
'    Set wdAppMain = Nothing
    
    Set wrdDoc = Nothing
    wdApp.NormalTemplate.Saved = True
    wdApp.Quit
    Set wdApp = Nothing
        
    'Main document
    wrdDocMain.SaveAs xlsFileMain
    wdAppMain.ActiveDocument.Close
    Set wrdDocMain = Nothing
    wdAppMain.NormalTemplate.Saved = True
    wdAppMain.Quit
    Set wdAppMain = Nothing
        
    'create employee letter ---------------- end
    
    'delete the single files - begin
    K = I
    ''For I = 1 To K
    ''    'xlsFileMat = xLocPath & "" & "WFC_IncentiveLetter" & xYear & "-" & Trim(Str(I)) & ".doc"
    ''    xlsFileMat = xEmpFile(I)
    ''    'delete this file if it exists
    ''    If (Dir(xlsFileMat)) <> "" Then
    ''        Kill xlsFileMat
    ''    End If
    ''Next
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = " "
    
    Screen.MousePointer = DEFAULT
    MDIMain.Timer1.Enabled = True
    
    cmdPrint.Enabled = True
    cmdClose2.Enabled = True
    
    'MsgBox "Please open " & xlsFileMain
    MsgBox "Please check the files from " & xLocPath & " folder"
    Exit Sub

CRW_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
MsgBox Err.Description
End Sub

Private Sub CreatePreparePayrollFiles()
Dim X%
Dim strWHand As String
Dim Msg As String, a%

On Error GoTo CRW_Err

If chkPreparePayroll() Then

    Msg = "Are You Sure You Want To Create Payroll Transaction File? "
    
    a% = MsgBox(Msg, 36, "Confirm")
    If a% <> 6 Then
        Exit Sub
    End If

    Screen.MousePointer = HOURGLASS
    If comCountry.Text = "Canada" Then
        Call WFC_CreatePayrollFilesCAD
    End If
    If comCountry.Text = "U.S.A." Then
        Call WFC_CreatePayrollFilesUSA
    End If
    
    Screen.MousePointer = DEFAULT
    MDIMain.Timer1.Enabled = True
End If
Exit Sub

CRW_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
MsgBox Err.Description

End Sub

'Public Sub cmdView_Click()
Private Sub CreateSpreadsheet()
Dim X%
Dim strWHand As String
Dim Msg As String, a%

On Error GoTo CRW_Err

If CriCheck() Then

    Msg = "Are You Sure You Want To Create Incentive Plan Spreadsheet? "
    
    a% = MsgBox(Msg, 36, "Confirm")
    If a% <> 6 Then
        Exit Sub
    End If

    Screen.MousePointer = HOURGLASS
    'x% = Cri_SetAll()
    Call WFC_CreateIncentivePlan
    Screen.MousePointer = DEFAULT
    MDIMain.Timer1.Enabled = True
End If
Exit Sub

CRW_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
MsgBox Err.Description

End Sub

Private Sub cmdImp_Click()
Dim SQLQ As String
Dim Msg As String, a%
Dim I As Integer
    
If glbWFC_IPPopFormName = "WFCIPSpreadSheetImport" Then
    If Not chkIMPSheet() Then Exit Sub
    Msg = ""
    If xFound Then
        Msg = "Found " & MskFiscalYear.Text & " Incentive Plan spreadsheet in the database. The program will delete all these records and then do the Import." & Chr(10) & Chr(10)
    End If
    Msg = Msg & "Are you sure you want to Import Incentive Plan Spreadsheet from "
    Msg = Msg & Chr(10) & txtFileName.Text & "?"
    
    a% = MsgBox(Msg, 36, "Confirm Import")
    If a% <> 6 Then
        Exit Sub
    End If
    
    If xFound Then
        SQLQ = "DELETE FROM HRIP_SPREADSHEET WHERE IP_YEAR = " & MskFiscalYear.Text & " "
        gdbAdoIhr001.Execute SQLQ, I
    End If
    
    ImpSucceeded = False
    Call WFCIPImpSpreadSheet
    
    If ImpSucceeded Then
        'MsgBox "   Finished.   "
        'Unload Me
    End If
End If

If glbWFC_IPPopFormName = "WFCIPUptOtherEarnings" Then 'Ticket #29015 Franks 01/10/2017
    If Not chkUptOtherEarnings() Then Exit Sub
    Msg = ""
    'If xFound Then
    '    Msg = "Found " & MskFiscalYear.Text & " Incentive Plan spreadsheet in the database. The program will delete all these records and then do the Import." & Chr(10) & Chr(10)
    'End If
    Msg = Msg & "Are you sure you want to update Other Earnings from Incentive Plan table?"
    'Msg = Msg & Chr(10) & txtFileName.Text & "?"
    
    a% = MsgBox(Msg, 36, "Confirm Update")
    If a% <> 6 Then
        Exit Sub
    End If


    ImpSucceeded = False
    'Call WFCIPIUptOtherEarningsFromFile 'not use
    Call WFCIPIUptOtherEarningsFromTable
    
    If ImpSucceeded Then
        'MsgBox "   Finished.   "
        'Unload Me
    End If
End If

End Sub
Private Sub WFCIPIUptOtherEarningsFromTable()
Dim exApp As Object, exBook As Object, exSheet As Object
Dim rsAdd As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset
Dim rsOEarn As New ADODB.Recordset
Dim rsIPSheet As New ADODB.Recordset
Dim xDiv, xPlant
Dim SQLQ As String
Dim I As Integer
Dim K As Integer
Dim xCode, xRate
Dim xYear, xEmpNo
Dim xRow, xRows, xTmp, xTm1, xTm2
Dim xNewRec, xUptRec, xTermSEQ
Dim xOECode
Dim xROIC

    Screen.MousePointer = vbHourglass
    
    ImpSucceeded = False
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = "Please Wait..."
    MDIMain.panHelp(2).Caption = ""
    
    xOECode = "BON4"

    
    xYear = MskFiscalYear.Text
    Call getDateRange(xYear)
    
    xROIC = ""
    'If Not IsEmpty(exSheet.Cells(6, 34)) Then
    '    xROIC = exSheet.Cells(6, 34)
    'End If
    
    xNewRec = 0
    xUptRec = 0
    
    SQLQ = "SELECT * FROM HRIP_SPREADSHEET WHERE IP_YEAR = " & xYear & " "
    'SQLQ = SQLQ & "AND IP_EMPNBR = " & xEmpNo & " "
    If rsIPSheet.State <> 0 Then rsIPSheet.Close
    rsIPSheet.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    I = 0
    If Not rsIPSheet.EOF Then
        xRows = rsIPSheet.RecordCount
    End If
    
    Do While Not rsIPSheet.EOF
        MDIMain.panHelp(0).FloodPercent = (I / xRows) * 100
        I = I + 1
        DoEvents
        
        xEmpNo = rsIPSheet("IP_EMPNBR")
        
        If Not IsNumeric(xEmpNo) Then GoTo next_emp

        xTermSEQ = 0
        SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
        If rsEmp.State <> 0 Then rsEmp.Close
        rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsEmp.EOF Then
            'check Term table
            SQLQ = "SELECT * FROM Term_HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
            If rsEmp.State <> 0 Then rsEmp.Close
            rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsEmp.EOF Then
                xTermSEQ = rsEmp("TERM_SEQ")
            End If
        End If
        If rsEmp.EOF Then 'not found in both Active and Term tables
            GoTo next_emp
        End If
        
        'find match Other Earning table
        If xTermSEQ = 0 Then
            'active table
            SQLQ = "SELECT * FROM HREARN WHERE EMPNBR = " & xEmpNo & " "
        Else
            'term table
            SQLQ = "SELECT * FROM Term_EARN WHERE EMPNBR = " & xEmpNo & " "
            SQLQ = SQLQ & "AND TERM_SEQ = " & xTermSEQ & " "
        End If
        SQLQ = SQLQ & "AND EARN_TYPE = '" & xOECode & "' "
        SQLQ = SQLQ & "AND FDATE = " & Date_SQL(xlocFDate) & " "
        SQLQ = SQLQ & "AND TDATE = " & Date_SQL(xlocTDate) & " "
        If rsOEarn.State <> 0 Then rsOEarn.Close
        rsOEarn.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsOEarn.EOF Then
            xNewRec = xNewRec + 1
            rsOEarn.AddNew
            rsOEarn("COMPNO") = "001"
            rsOEarn("EMPNBR") = xEmpNo
            rsOEarn("EARN_TYPE") = xOECode
            rsOEarn("FDATE") = CVDate(xlocFDate)
            rsOEarn("TDATE") = CVDate(xlocTDate)
            If xTermSEQ > 0 Then
                rsOEarn("TERM_SEQ") = xTermSEQ
            End If
        Else
            xUptRec = xUptRec + 1
        End If
        
        'xTmp = rsIPSheet("IP_PAYOUT") 'exSheet.Cells(I, 34) 'AH 'amount
        If Not IsNull(rsIPSheet("IP_PAYOUT")) Then
            rsOEarn("ACT_DOLLAR") = Round(rsIPSheet("IP_PAYOUT"), 2)
        End If
        'xTmp = exSheet.Cells(I, 35) 'AI 'Corporate Equivalent
        If Not IsNull(rsIPSheet("IP_ROIC_ADJ")) Then
            rsOEarn("CORP_EQUI") = Round(rsIPSheet("IP_ROIC_ADJ"), 2)
        End If
        rsOEarn("COST_OF_EMPLOYMENT") = 1 'COE
        
        'Position Code
        rsOEarn("OE_JOB") = rsIPSheet("IP_POSCODE")

        'If IsNumeric(xROIC) Then
            rsOEarn("IP_ROIC") = rsIPSheet("IP_ROIC") 'xROIC
        'End If
        rsOEarn("LDATE") = Date
        rsOEarn("LTIME") = Time$
        rsOEarn("LUSER") = glbUserID
        rsOEarn.Update
        
next_emp:
        rsIPSheet.MoveNext
    Loop
    
    Screen.MousePointer = vbDefault
        
    ImpSucceeded = True
    
    MsgBox xNewRec & " new record(s) and " & xUptRec & " record(s) updated. "
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

End Sub
Private Sub WFCIPIUptOtherEarningsFromFile() 'not use
Dim exApp As Object, exBook As Object, exSheet As Object
Dim rsAdd As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset
Dim rsOEarn As New ADODB.Recordset
Dim xDiv, xPlant
Dim SQLQ As String
Dim I As Integer
Dim K As Integer
Dim xCode, xRate
Dim xYear, xEmpNo
Dim xRow, xRows, xTmp, xTm1, xTm2
Dim xNewRec, xUptRec, xTermSEQ
Dim xOECode
Dim xROIC

    Screen.MousePointer = vbHourglass
    
    ImpSucceeded = False
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = "Please Wait..."
    MDIMain.panHelp(2).Caption = ""
    
    xOECode = "BON4"
    
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(ImportFile)
    Set exSheet = exBook.Worksheets(1)
        
    xCode = exSheet.Cells(7, 8)
    If Not UCase(xCode) = UCase("Manager") Then
        MsgBox "Invalid File Layout."
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing
        Screen.MousePointer = vbDefault
        MDIMain.panHelp(0).FloodType = 0
        Exit Sub
    End If
        
    xYear = MskFiscalYear.Text
    xRow = 8 'First Row
    xRows = getRows(exSheet, xRow)
    
    Call getDateRange(xYear)
    If Len(xlocFDate) = 0 Then
        MsgBox "Invalid From Date."
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing
        Screen.MousePointer = vbDefault
        MDIMain.panHelp(0).FloodType = 0
        Exit Sub
    End If
    
    xROIC = ""
    If Not IsEmpty(exSheet.Cells(6, 34)) Then
        xROIC = exSheet.Cells(6, 34)
    End If
    
    xNewRec = 0
    xUptRec = 0
    For I = xRow To xRows
        MDIMain.panHelp(0).FloodPercent = (I / xRows) * 100
        DoEvents
        
        xEmpNo = exSheet.Cells(I, 1)

        If Not IsNumeric(xEmpNo) Then GoTo next_emp

        xTermSEQ = 0
        SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
        If rsEmp.State <> 0 Then rsEmp.Close
        rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsEmp.EOF Then
            'check Term table
            SQLQ = "SELECT * FROM Term_HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
            If rsEmp.State <> 0 Then rsEmp.Close
            rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsEmp.EOF Then
                xTermSEQ = rsEmp("TERM_SEQ")
            End If
        End If
        If rsEmp.EOF Then 'not found in both Active and Term tables
            GoTo next_emp
        End If
        
        'find match Other Earning table
        If xTermSEQ = 0 Then
            'active table
            SQLQ = "SELECT * FROM HREARN WHERE EMPNBR = " & xEmpNo & " "
        Else
            'term table
            SQLQ = "SELECT * FROM Term_EARN WHERE EMPNBR = " & xEmpNo & " "
            SQLQ = SQLQ & "AND TERM_SEQ = " & xTermSEQ & " "
        End If
        SQLQ = SQLQ & "AND EARN_TYPE = '" & xOECode & "' "
        SQLQ = SQLQ & "AND FDATE = " & Date_SQL(xlocFDate) & " "
        SQLQ = SQLQ & "AND TDATE = " & Date_SQL(xlocTDate) & " "
        If rsOEarn.State <> 0 Then rsOEarn.Close
        rsOEarn.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsOEarn.EOF Then
            xNewRec = xNewRec + 1
            rsOEarn.AddNew
            rsOEarn("COMPNO") = "001"
            rsOEarn("EMPNBR") = xEmpNo
            rsOEarn("EARN_TYPE") = xOECode
            rsOEarn("FDATE") = CVDate(xlocFDate)
            rsOEarn("TDATE") = CVDate(xlocTDate)
            If xTermSEQ > 0 Then
                rsOEarn("TERM_SEQ") = xTermSEQ
            End If
        Else
            xUptRec = xUptRec + 1
        End If
        
        xTmp = exSheet.Cells(I, 34) 'AH 'amount
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then
            rsOEarn("ACT_DOLLAR") = xTmp
        End If
        xTmp = exSheet.Cells(I, 35) 'AI 'Corporate Equivalent
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then
            rsOEarn("CORP_EQUI") = xTmp
        End If
        rsOEarn("COST_OF_EMPLOYMENT") = 1 'COE
        
        'Position Code
        xTmp = exSheet.Cells(I, 14)
        xTm1 = "": xTm1 = ""
        K = InStr(1, xTmp, "/")
        If K > 0 Then
            xTm1 = Trim(Left(xTmp, K - 1))
            xTm2 = Trim(Right(xTmp, Len(xTmp) - K))
        End If
        If Len(xTm1) > 0 Then
            rsOEarn("OE_JOB") = Left(xTm1, 20)
        End If
        If IsNumeric(xROIC) Then
            rsOEarn("IP_ROIC") = xROIC
        End If
        rsOEarn("LDATE") = Date
        rsOEarn("LTIME") = Time$
        rsOEarn("LUSER") = glbUserID
        rsOEarn.Update

next_emp:

    Next

    
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing
    
    Screen.MousePointer = vbDefault
        
    ImpSucceeded = True
    
    MsgBox xNewRec & " new record(s) and " & xUptRec & " record(s) updated. "
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

End Sub

Private Sub WFCIPImpSpreadSheet()
Dim exApp As Object, exBook As Object, exSheet As Object
Dim rsIPSheet As New ADODB.Recordset
Dim rsAdd As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset
Dim xDiv, xPlant
Dim SQLQ As String
Dim I As Integer
Dim K As Integer
Dim xCode, xRate
Dim xYear, xEmpNo
Dim xRow, xRows, xTmp, xTm1, xTm2
Dim xROIC
Dim xSection

    Screen.MousePointer = vbHourglass
    
    ImpSucceeded = False
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = "Please Wait..."
    MDIMain.panHelp(2).Caption = ""
    
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(ImportFile)
    Set exSheet = exBook.Worksheets(1)
        
    xCode = exSheet.Cells(7, 8)
    If Not UCase(xCode) = UCase("Manager") Then
        MsgBox "Invalid File Layout."
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing
        Screen.MousePointer = vbDefault
        MDIMain.panHelp(0).FloodType = 0
        Exit Sub
    End If
        
    xYear = MskFiscalYear.Text
    xRow = 8 'First Row
    xRows = getRows(exSheet, xRow)

    Call getDateRange(xYear)
    
    If IsEmpty(exSheet.Cells(6, 34)) Then
        xROIC = ""
    Else
        xROIC = exSheet.Cells(6, 34)
    End If
    
    For I = xRow To xRows
        MDIMain.panHelp(0).FloodPercent = (I / xRows) * 100
        DoEvents
        
        xEmpNo = exSheet.Cells(I, 1)
        If Not IsNumeric(xEmpNo) Then GoTo next_emp
        SQLQ = "SELECT * FROM HRIP_SPREADSHEET WHERE IP_YEAR = " & xYear & " "
        SQLQ = SQLQ & "AND IP_EMPNBR = " & xEmpNo & " "
        If rsIPSheet.State <> 0 Then rsIPSheet.Close
        rsIPSheet.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsIPSheet.EOF Then
            rsIPSheet.AddNew
        End If
        rsIPSheet("IP_YEAR") = xYear
        If IsDate(xlocFDate) Then rsIPSheet("IP_FDATE") = CVDate(xlocFDate)
        If IsDate(xlocTDate) Then rsIPSheet("IP_TDATE") = CVDate(xlocTDate)
            
        rsIPSheet("IP_EMPNBR") = xEmpNo
        
        'find the Section(Plant) from INFOHR
        SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
        If rsEmp.State <> 0 Then rsEmp.Close
        rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        xSection = ""
        If rsEmp.EOF Then
            SQLQ = "SELECT * FROM Term_HREMP WHERE ED_EMPNBR = " & xEmpNo & " ORDER BY TERM_SEQ DESC "
            If rsEmp.State <> 0 Then rsEmp.Close
            rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsEmp.EOF Then
                If Not IsNull(rsEmp("ED_SECTION")) Then
                    xSection = rsEmp("ED_SECTION")
                End If
            End If
        Else
            If Not IsNull(rsEmp("ED_SECTION")) Then
                xSection = rsEmp("ED_SECTION")
            End If
        End If
        'If rsEmp.EOF Then
        '    Debug.Print xEmpNO
        'End If
        xTmp = exSheet.Cells(I, 2)
        rsIPSheet("IP_COUNTRY") = Left(xTmp, 10)
        xTmp = exSheet.Cells(I, 3)
        rsIPSheet("IP_DIV") = Left(xTmp, 4)
        xTmp = exSheet.Cells(I, 4)
        rsIPSheet("IP_LOC") = Left(xTmp, 10)
        xTmp = exSheet.Cells(I, 5)
        rsIPSheet("IP_REGION") = Left(xTmp, 20)
        xTmp = exSheet.Cells(I, 6)
        rsIPSheet("IP_POSTYPE") = Left(xTmp, 20)
        xTmp = exSheet.Cells(I, 7)
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then
            rsIPSheet("IP_REPTAU") = xTmp
        End If
        xTmp = exSheet.Cells(I, 8)
        If Len(xTmp) > 0 Then rsIPSheet("IP_REPTAUNAME") = Left(xTmp, 60)
        xTmp = exSheet.Cells(I, 10)
        If Len(xTmp) > 0 Then rsIPSheet("IP_EMPSTATUS") = Left(xTmp, 10)
        xTmp = exSheet.Cells(I, 11)
        If Len(xTmp) > 0 Then rsIPSheet("IP_EMPNAME") = Left(xTmp, 60)
        xTmp = exSheet.Cells(I, 12)
        If Len(xTmp) > 0 Then
            If IsDate(xTmp) Then rsIPSheet("IP_DOB") = CVDate(xTmp)
        End If
        xTmp = exSheet.Cells(I, 13)
        If Len(xTmp) > 0 Then
            If IsDate(xTmp) Then rsIPSheet("IP_NEXTDAT") = CVDate(xTmp)
        End If
        
        xTmp = exSheet.Cells(I, 14)
        xTm1 = "": xTm1 = ""
        K = InStr(1, xTmp, "/")
        If K > 0 Then
            xTm1 = Trim(Left(xTmp, K - 1))
            xTm2 = Trim(Right(xTmp, Len(xTmp) - K))
        End If
        If Len(xTm1) > 0 Then rsIPSheet("IP_POSCODE") = Left(xTm1, 20)
        If Len(xTm2) > 0 Then rsIPSheet("IP_JOBCODE") = Left(xTm2, 20)
        If Len(xTmp) > 0 Then rsIPSheet("IP_POSJOBCODE") = Left(xTmp, 40)
        xTmp = exSheet.Cells(I, 15)
        If Len(xTmp) > 0 Then rsIPSheet("IP_POSDESCR") = Left(xTmp, 100)
        xTmp = exSheet.Cells(I, 16)
        If Len(xTmp) > 0 Then rsIPSheet("IP_JBSTATUS") = Left(xTmp, 6)
        xTmp = exSheet.Cells(I, 17)
        If Len(xTmp) > 0 Then rsIPSheet("IP_JBGRPCD") = Left(xTmp, 6)
        xTmp = exSheet.Cells(I, 18)
        If Len(xTmp) > 0 Then
            If IsDate(xTmp) Then rsIPSheet("IP_POSDATE") = CVDate(xTmp)
        End If
        xTmp = exSheet.Cells(I, 19)
        If Len(xTmp) > 0 Then rsIPSheet("IP_MARKETLINE") = Left(xTmp, 4)
        xTmp = exSheet.Cells(I, 20)
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_MDOLLARS") = xTmp
        xTmp = exSheet.Cells(I, 21)
        If Len(xTmp) > 0 Then rsIPSheet("IP_BAND") = Left(xTmp, 3)
        
        xTmp = exSheet.Cells(I, 22)
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_ANNSALARY") = xTmp
        xTmp = exSheet.Cells(I, 23)
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_ANNCURRENCY") = xTmp
        xTmp = exSheet.Cells(I, 24)
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_PAYOUT_TARGET") = xTmp
        xTmp = exSheet.Cells(I, 25)
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_CDN_PAYOUT") = xTmp
        xTmp = exSheet.Cells(I, 26)
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_MONTHS1") = xTmp
        
        xTmp = exSheet.Cells(I, 28)
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_MONTHS2") = xTmp
        xTmp = exSheet.Cells(I, 29)
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_LOCAL_CURRENCY") = xTmp
        xTmp = exSheet.Cells(I, 30)
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_EXCH_RATE") = xTmp
        xTmp = exSheet.Cells(I, 31)
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_PAYROLL_CDN") = xTmp
        xTmp = exSheet.Cells(I, 32)
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_CDN_ADJUSTED") = xTmp
        xTmp = exSheet.Cells(I, 33)
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_CDN_AFTER_ADJ") = Round(xTmp, 0)
        xTmp = exSheet.Cells(I, 34)
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_PAYOUT") = Round(xTmp, 0)
        xTmp = exSheet.Cells(I, 35) 'AI
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_ROIC_ADJ") = Round(xTmp, 0) ' xTmp
        xTmp = exSheet.Cells(I, 38) 'AL
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_PLANT_PLANT_OBJ") = Round(xTmp, 0) ' xTmp
        xTmp = exSheet.Cells(I, 39) 'AM
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_PLANT_BU_FIN") = Round(xTmp, 0) '  xTmp
        xTmp = exSheet.Cells(I, 40) 'AN
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_PLANT_CORP_FIN") = Round(xTmp, 0) '  xTmp
        xTmp = exSheet.Cells(I, 41) 'AO
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_PLANT_SUM") = Round(xTmp, 0) '  xTmp
        xTmp = exSheet.Cells(I, 42) 'AP
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_PLANT_PERCENT") = xTmp
        
        
        xTmp = exSheet.Cells(I, 43) 'AQ
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_BU_BU_FIN") = Round(xTmp, 0) ' xTmp
        xTmp = exSheet.Cells(I, 44) 'AR
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_BU_CORP_FIN") = Round(xTmp, 0) '  xTmp
        xTmp = exSheet.Cells(I, 45) 'AS
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_BU_SUM") = Round(xTmp, 0) '  xTmp
        xTmp = exSheet.Cells(I, 46) 'AT
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_BU_PERCENT") = xTmp
        xTmp = exSheet.Cells(I, 47) 'AU
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_COMM_SALES_IND") = Round(xTmp, 0) '  xTmp
        xTmp = exSheet.Cells(I, 48) 'AV
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_COMM_SALES_COMM") = Round(xTmp, 0) '  xTmp
        xTmp = exSheet.Cells(I, 49) 'AW
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_COMM_CORP_FIN") = Round(xTmp, 0) '  xTmp
        xTmp = exSheet.Cells(I, 50) 'AX
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_COMM_SUM") = Round(xTmp, 0) '  xTmp
        xTmp = exSheet.Cells(I, 51) 'AY
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_COMM_PERCENT") = xTmp
        
        xTmp = exSheet.Cells(I, 52) 'AZ
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_CORP_CORP_OBJ") = Round(xTmp, 0) '  xTmp
        xTmp = exSheet.Cells(I, 53) 'BA
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_CORP_CORP_FIN") = Round(xTmp, 0) '  xTmp
        xTmp = exSheet.Cells(I, 54) 'BB
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_CORP_SUM") = Round(xTmp, 0) '  xTmp
        xTmp = exSheet.Cells(I, 55) 'BC
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_SUM_AZ_BA") = Round(xTmp, 0) '  xTmp
        xTmp = exSheet.Cells(I, 56) 'BD
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_SUM_TOTAL") = Round(xTmp, 0) '  xTmp
        xTmp = exSheet.Cells(I, 57) 'BE
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_PAYOUT_LOCAL") = Round(xTmp, 0) ' xTmp
        xTmp = exSheet.Cells(I, 58) 'BF
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_PAYOUT_ROIC") = Round(xTmp, 0) '  xTmp
        xTmp = exSheet.Cells(I, 59) 'BG
        If Not IsEmpty(xTmp) And IsNumeric(xTmp) Then rsIPSheet("IP_PAYOUT_ROICLOC") = Round(xTmp, 0) ' xTmp
        xTmp = exSheet.Cells(I, 60) 'BH
        If Len(xTmp) > 0 Then rsIPSheet("IP_SYSTEM_COMM") = Left(xTmp, 100)
        xTmp = exSheet.Cells(I, 61) 'BI
        If Len(xTmp) > 0 Then rsIPSheet("IP_PH_COMMENTS") = Left(xTmp, 1000)
        
        If Len(xSection) > 0 Then
            xROIC = getEmpROIC(xYear, xSection)
        End If
        If IsNumeric(xROIC) Then
            rsIPSheet("IP_ROIC") = xROIC
        End If
        
        If Len(xSection) > 0 Then
            rsIPSheet("IP_SECTION") = Left(xSection, 4)
        End If
        
        rsIPSheet("IP_LDATE") = Date
        rsIPSheet("IP_LTIME") = Time$
        rsIPSheet("IP_LUSER") = glbUserID
        rsIPSheet.Update
        

next_emp:

    Next

    
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing
    
    Screen.MousePointer = vbDefault
        
    ImpSucceeded = True
    
    MsgBox (xRows - 7) & " records imported. "
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

End Sub

Private Function chkPrintIP()
Dim rsIPSheet As New ADODB.Recordset
Dim SQLQ As String
Dim X%, Y%
Dim xStr
Dim xlsFileTmp As String

chkPrintIP = False

On Error GoTo chkPrintIP_Err

If Len(MskFiscalYeaLetter.Text) > 0 Then
    If Not IsNumeric(MskFiscalYeaLetter.Text) Then
        MsgBox "Invalid Incentive Letter for Year."
        MskFiscalYeaLetter.SetFocus
        Exit Function
    End If
    If Not Len(MskFiscalYeaLetter.Text) = 4 Then
        MsgBox "Invalid Incentive Letter for Year."
        MskFiscalYeaLetter.SetFocus
        Exit Function
    End If
Else
    MsgBox "Incentive Letter for Year is a required field"
    MskFiscalYeaLetter.SetFocus
    Exit Function
End If

If Not clpDiv2.ListChecker Then
    Exit Function
End If


If Not elpEEID.ListChecker Then
    Exit Function
End If

chkPrintIP = True

Exit Function

chkPrintIP_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, Me.Caption, "chkPrintIP", "Print Incentive Plan", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Function chkEmpLetter()
Dim rsIPSheet As New ADODB.Recordset
Dim SQLQ As String
Dim X%, Y%
Dim xStr
Dim xlsFileTmp As String

chkEmpLetter = False

On Error GoTo chkEmpLetter_Err

If Len(MskFiscalYeaLetter.Text) > 0 Then
    If Not IsNumeric(MskFiscalYeaLetter.Text) Then
        MsgBox "Invalid Incentive Letter for Year."
        MskFiscalYeaLetter.SetFocus
        Exit Function
    End If
    If Not Len(MskFiscalYeaLetter.Text) = 4 Then
        MsgBox "Invalid Incentive Letter for Year."
        MskFiscalYeaLetter.SetFocus
        Exit Function
    End If
Else
    MsgBox "Incentive Letter for Year is a required field"
    MskFiscalYeaLetter.SetFocus
    Exit Function
End If

If Not clpDiv2.ListChecker Then
    Exit Function
End If


For X% = 5 To 6
    If Not clpCode(X).ListChecker Then Exit Function
Next X%

If Len(dlpPayDate.Text) = 0 Then
    MsgBox "Scheduled Payment Date is required!"
    dlpPayDate.SetFocus
    Exit Function
End If
If Not IsDate(dlpPayDate.Text) Then
    MsgBox "Not a valid date"
    dlpPayDate.SetFocus
    Exit Function
End If

If Not elpEEID.ListChecker Then
    Exit Function
End If

chkEmpLetter = True

Exit Function

chkEmpLetter_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, Me.Caption, "chkEmpLetter", "Employee Incentive Letter", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Function chkPreparePayroll()
Dim rsIPSheet As New ADODB.Recordset
Dim SQLQ As String
Dim X%, Y%
Dim xStr
Dim xlsFileTmp As String

chkPreparePayroll = False

On Error GoTo chkPreparePayroll_Err

If Len(MskFiscalYea2.Text) > 0 Then
    If Not IsNumeric(MskFiscalYea2.Text) Then
        MsgBox "Invalid Fiscal Year."
        MskFiscalYea2.SetFocus
        Exit Function
    End If
    If Not Len(MskFiscalYea2.Text) = 4 Then
        MsgBox "Invalid Fiscal Year."
        MskFiscalYea2.SetFocus
        Exit Function
    End If
Else
    MsgBox "Fiscal Year is a required field"
    MskFiscalYea2.SetFocus
    Exit Function
End If

If comCountry.Text = "Canada" Then 'Earning Code cannot be entered if Country equals "Canada".
    If Len(comEarnCode.Text) > 0 Then
        MsgBox "Earning Code cannot be entered if Country is 'Canada'"
        comEarnCode.SetFocus
        Exit Function
    End If
End If
    
'If Country equals "U.S.A." and Pay Group is not blank, Earning Code must be entered.
If comCountry.Text = "U.S.A." Then
    If Len(clpVadim2.Text) > 0 Then
        If Len(comEarnCode.Text) = 0 Then
            MsgBox "If Country equals 'U.S.A.' and Pay Group is not blank, Earning Code must be entered"
            comEarnCode.SetFocus
            Exit Function
        End If
    End If
End If
    
    
xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "WFCPayCodeList.xls"
If Dir(xlsFileTmp) = "" Then
    MsgBox "There is no " & xlsFileTmp
    Exit Function
End If
    
''check if there are Incentive Plan data in the table
'SQLQ = "SELECT TOP 10 * FROM HRIP_SPREADSHEET WHERE IP_YEAR = " & MskFiscalYea2.Text & " "
'If rsIPSheet.State <> 0 Then rsIPSheet.Close
'rsIPSheet.Open SQLQ, gdbAdoIhr001, adOpenStatic
'If rsIPSheet.EOF Then
'    MsgBox "There is no any record of " & MskFiscalYea2.Text & " in the Incentive Plan table(HRIP_SPREADSHEET) "
'    Exit Function
'End If
'rsIPSheet.Close

chkPreparePayroll = True

Exit Function

chkPreparePayroll_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, Me.Caption, "chkPreparePayroll", "Prepare Payroll", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

Private Function chkUptOtherEarnings()
Dim rsIPSheet As New ADODB.Recordset
Dim SQLQ As String
Dim X%, Y%
Dim xStr

chkUptOtherEarnings = False

On Error GoTo chkUptOtherEarnings_Err

If Len(MskFiscalYear.Text) > 0 Then
    If Not IsNumeric(MskFiscalYear.Text) Then
        MsgBox "Invalid Fiscal Year."
        MskFiscalYear.SetFocus
        Exit Function
    End If
    If Not Len(MskFiscalYear.Text) = 4 Then
        MsgBox "Invalid Fiscal Year."
        MskFiscalYear.SetFocus
        Exit Function
    End If
Else
    MsgBox "Fiscal Year is a required field"
    MskFiscalYear.SetFocus
    Exit Function
End If

'Ticket #29679 Franks 01/18/2016 - update from the Incentive Plan table
''If Len(txtFileName.Text) = 0 Then
''    MsgBox "Please enter Import From File"
''    txtFileName.SetFocus
''    Exit Function
''End If
''ImportFile = txtFileName.Text
''If Dir(ImportFile) = "" Then
''  MsgBox ImportFile & " File not Found."
''  txtFileName.SetFocus
''  Exit Function
''End If
    
'check if there are Incentive Plan data in the table
SQLQ = "SELECT TOP 10 * FROM HRIP_SPREADSHEET WHERE IP_YEAR = " & MskFiscalYear.Text & " "
If rsIPSheet.State <> 0 Then rsIPSheet.Close
rsIPSheet.Open SQLQ, gdbAdoIhr001, adOpenStatic
If rsIPSheet.EOF Then
    MsgBox "There is no any record of " & MskFiscalYear.Text & " in the Incentive Plan table(HRIP_SPREADSHEET) "
    Exit Function
End If
rsIPSheet.Close

chkUptOtherEarnings = True

Exit Function

chkUptOtherEarnings_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, Me.Caption, "chkUptOtherEarnings", "Other Earnings", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

Private Function chkIMPSheet()
Dim rsIPSheet As New ADODB.Recordset
Dim SQLQ As String
Dim X%, Y%
Dim xStr

chkIMPSheet = False

On Error GoTo chkIMPSheet_Err

If Len(MskFiscalYear.Text) > 0 Then
    If Not IsNumeric(MskFiscalYear.Text) Then
        MsgBox "Invalid Fiscal Year."
        MskFiscalYear.SetFocus
        Exit Function
    End If
    If Not Len(MskFiscalYear.Text) = 4 Then
        MsgBox "Invalid Fiscal Year."
        MskFiscalYear.SetFocus
        Exit Function
    End If
Else
    MsgBox "Fiscal Year is a required field"
    MskFiscalYear.SetFocus
    Exit Function
End If


If Len(txtFileName.Text) = 0 Then
    MsgBox "Please enter Import From File"
    txtFileName.SetFocus
    Exit Function
End If

ImportFile = txtFileName.Text

If Dir(ImportFile) = "" Then
  MsgBox ImportFile & " File not Found."
  txtFileName.SetFocus
  Exit Function
End If
    
xFound = False
SQLQ = "SELECT TOP 10 * FROM HRIP_SPREADSHEET WHERE IP_YEAR = " & MskFiscalYear.Text
rsIPSheet.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsIPSheet.EOF Then
    xFound = True
End If
rsIPSheet.Close


chkIMPSheet = True

Exit Function

chkIMPSheet_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkIMPSheet", "HRIP_HRIP_SPREADSHEET", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

Private Sub cmdPrint_Click()
    If glbWFC_IPPopFormName = "WFCIPPrintEmpLetter" Then 'Ticket #29016 Franks 02/01/2017
        Call CreateEmpLetter
    End If
    If glbWFC_IPPopFormName = "WFCIPPrintSpreadsheet" Then 'Ticket #29810 Franks 02/07/2017
        MsgBox "This funciton is not done yet."
        Exit Sub
        
        If OptFullPlan.Value Then
            Call PrintIP_Plan_Details
        End If
        If OptSum.Value Then
            'Call PrintIP_Plan_Summary
        End If
    End If
End Sub

Private Sub PrintIP_Plan_Details()
Dim exApp As Object, exBook As Object, exSheet As Object
Dim rsIPSheet As New ADODB.Recordset
Dim rsAdd As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset
Dim Msg As String, a%
Dim xDiv, xPlant
Dim locPath, xLocPath
Dim xlsFileTmp As String
Dim xlsFileMat As String
Dim xlsFileTmpl1 As String
Dim xlsFileTmplM As String
Dim xSaveFileName As String
Dim xlsFileMain As String
Dim SQLQ As String
Dim I As Integer
Dim totNum As Integer
Dim K As Integer
Dim xCode, xRate
Dim xYear, xEmpNo, xStartLine
Dim xRow, xRows, xTmp, xTm1, xTm2
Dim xROIC
    
    Exit Sub 'Jerry said: MZ and Peter not use Import Spreadsheet, so we can't print it here
    
    'If chkEmpLetter() Then
    If chkPrintIP() Then
        Msg = "Are You Sure You Want To Print Incentive Full Plan? "
        
        a% = MsgBox(Msg, 36, "Confirm")
        If a% <> 6 Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    

    
    'create employee letter ---------------- begin
    
    xYear = MskFiscalYeaLetter.Text
    Call getDateRange(xYear)
    
    SQLQ = "SELECT * FROM HRIP_SPREADSHEET WHERE (1=1) "
    SQLQ = SQLQ & "AND IP_YEAR = " & xYear & " "
    'SQLQ = SQLQ & "AND FDATE = " & Date_SQL(xlocFDate) & " "
    'SQLQ = SQLQ & "AND TDATE = " & Date_SQL(xlocTDate) & " "
    '---- check HREMP table --- begin
    If Len(clpDiv2.Text) > 0 Then
        SQLQ = SQLQ & "AND IP_DIV IN ('" & Replace(clpDiv2.Text, ",", "','") & "') "
    End If
    'If Len(clpCode(5).Text) > 0 Then
    '    SQLQ = SQLQ & "AND ED_SECTION IN ('" & Replace(clpCode(5).Text, ",", "','") & "') "
    'End If
    If Len(clpCode(6).Text) > 0 Then
        SQLQ = SQLQ & "AND IP_REGION IN ('" & Replace(clpCode(6).Text, ",", "','") & "') "
    End If
    If Len(elpEEID2.Text) > 0 Then
        SQLQ = SQLQ & "AND IP_EMPNBR IN (" & getEmpnbr(elpEEID2.Text) & ") "
    End If
    '---- check HREMP table --- end
    If comGroup(0).Text = lStr("Section") Then 'by plant, employee, date
        SQLQ = SQLQ & "ORDER BY ED_SECTION, ED_SURNAME,ED_FNAME"
    ElseIf comGroup(0).Text = lStr("Region") Then
        SQLQ = SQLQ & "ORDER BY ED_REGION, ED_SURNAME,ED_FNAME"
    ElseIf comGroup(0).Text = lStr("Division") Then
        SQLQ = SQLQ & "ORDER BY ED_DIV, ED_SURNAME,ED_FNAME"
    Else
        SQLQ = SQLQ & "ORDER BY  ED_SURNAME,ED_FNAME"
    End If
    
    If rsEmp.State <> 0 Then rsEmp.Close
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsEmp.EOF Then
        MsgBox "No 'BON4' record found in this Selection Criteria "
        Exit Sub
    End If

    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(0).FloodPercent = 0
    

    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "WFC_IncentivePlanDetailed_Tmp.xls"
    xlsFileMat = Dir1.Path & IIf(Right(Dir1.Path, 1) = "\", "", "\") & "WFC_IncentivePlanDetailed(" & Trim(glbUserID) & ").xls"

    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    FileCopy xlsFileTmp, xlsFileMat
    
    'Populate Excel file - begin
    'Create new WorkBook of Excel
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(xlsFileMat)
    Set exSheet = exBook.Worksheets(1)
            
    'exSheet.Cells(2, 2) = "Status of WPS Training for " & Year(Date)
    exSheet.Cells(2, 1) = "" & Format(Date, "MMM dd, YYYY") & ""
    
    exSheet.Cells(4, 26) = dlpAsOf.Text 'Col Z
    
    
    Screen.MousePointer = vbHourglass
    
    'First line of data
    xStartLine = 8
    xRow = xStartLine
    'xJun1stDate = CVDate("Jun 1, " & Year(Date) - 1)
    'xSalAsDate = CVDate(dlpSalAsOf.Text)
    Do While Not rsEmp.EOF
        If (I / totNum) <= 1 Then
            MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
            I = I + 1
        End If
        DoEvents
    

        exSheet.Cells(xRow, 1) = rsEmp("ED_EMPNBR")
        exSheet.Cells(xRow, 2) = rsEmp("ED_COUNTRY")
        exSheet.Cells(xRow, 3) = rsEmp("ED_DIV")
        exSheet.Cells(xRow, 4) = rsEmp("ED_LOC")
        exSheet.Cells(xRow, 5) = rsEmp("ED_REGION")
        
        xRow = xRow + 1

Next_Rec:
        rsEmp.MoveNext
    Loop
    rsEmp.Close

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
    
    
    Screen.MousePointer = vbDefault
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    
End Sub

Private Sub comQ1_Click()
If comQ1.Text = "Yes" Then
    Screen.MousePointer = HOURGLASS
    Load frmIPFactors
    'frmIPFactors.ZOrder 0
    Screen.MousePointer = DEFAULT
    
    'Unload Me
    'If glbOnTop = "FRMEESTATS" Then glbOnTop = ""
End If

End Sub

Private Sub Drive1_Change()
Dim xdir, xerror
On Error GoTo CKERROR
xerror = False
Dir1.Path = Drive1.Drive
Exit Sub
CKERROR:
    If Err = 68 Then
         MsgBox "Invalid Drive Selected"
         Drive1.Drive = App.Path
         xerror = True
         Resume Next
    End If
    MsgBox "ERROR " & Str(Err)
    xerror = True
    Resume Next
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
Dim xTmp

glbOnTop = UCase("frmIPCreateSheet")

If glbWFC_IPPopFormName = "WFCIPSpreadSheetCreate" Then
    comQ1.AddItem "Yes"
    comQ1.AddItem "No"
    comQ1.ListIndex = -1
    
    Call MonthDescAdd
    
    Call SetLocLabels
    
    dlpAsOf.Text = CVDate("Oct 31, " & Year(Date))
    dlpSalAsOf.Text = dlpAsOf.Text
    
    clpJob.TextBoxWidth = 1315
    
    fraSunLife.Visible = True
    'Call INI_Controls(Me)
End If

If glbWFC_IPPopFormName = "WFCIPSpreadSheetImport" Then
    Me.Caption = "Import Incentive Plan Spreadsheet"
    fraImportSheet.Left = 0
    fraImportSheet.Top = 0
    fraImportSheet.Height = 9135
    fraImportSheet.Width = 14535
    fraImportSheet.BorderStyle = 0
    fraImportSheet.Visible = True
End If

If glbWFC_IPPopFormName = "WFCIPUptOtherEarnings" Then 'Ticket #29015 Franks 01/10/2017
    Me.Caption = "Update info:HR Other Earnings"
    fraImportSheet.Left = 0
    fraImportSheet.Top = 0
    fraImportSheet.Height = 9135
    fraImportSheet.Width = 14535
    fraImportSheet.BorderStyle = 0
    fraImportSheet.Visible = True
    cmdImp.Caption = "Update"
    'Ticket #29679 Franks 01/18/2016 - update from the Incentive Plan table
    '"Jerry: I think should update from the Incentive Plan table for the year and don't ask for the import from. - Talked this with Jerry, he let me do it.
    lblCriteria(4).Visible = False
    txtFileName.Visible = False
    cmdBrowse.Visible = False
End If

If glbWFC_IPPopFormName = "WFCIPPreparePayroll" Then
    Me.Caption = "Prepare Payroll Transaction File"
    fraToPayroll.Left = 0
    fraToPayroll.Top = 0
    'fraToPayroll.Height = 9135
    fraToPayroll.Width = 14535
    fraToPayroll.BorderStyle = 0
    fraToPayroll.Visible = True
    
    fraSunLife.Visible = True
    fraSunLife.Top = 1500
    fraSunLife.Height = 3615
    fraSunLife.Width = 9000
    cmdCreate.Caption = "Export"
    
    comCountry.Clear
    comCountry.AddItem "Canada"
    comCountry.AddItem "U.S.A."
    comCountry.ListIndex = 0
    
    comEarnCode.Clear
    comEarnCode.AddItem ""
    comEarnCode.AddItem "3"
    comEarnCode.AddItem "4"
    comEarnCode.AddItem "5"
    comEarnCode.ListIndex = 0
    
    'lblVadim2.Caption = lStr("Vadim Field 2")
    
End If

If glbWFC_IPPopFormName = "WFCIPPrintEmpLetter" Then 'Ticket #29016 Franks 02/01/2017
    Me.Caption = "Print Employee Incentive Letter"
    fraEmpLetter.Left = 0
    fraEmpLetter.Top = 200
    fraEmpLetter.Height = 6015 ' 9135
    fraEmpLetter.Width = 14535
    fraEmpLetter.BorderStyle = 0
    fraEmpLetter.Visible = True
    
    comGroup(0).Clear
    comGroup(0).AddItem lStr("Section")
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem "(none)"
    comGroup(0).ListIndex = 0
    comGroup(1).Clear
    comGroup(1).AddItem "Employee Name"
    comGroup(1).ListIndex = 0

    fraSunLife.Visible = True
    fraSunLife.Left = 600
    fraSunLife.Top = 6720 - 400
    fraSunLife.Height = 3615 - 600
    fraSunLife.Width = 9000
    
    OptSum.Caption = "English"
    OptFullPlan.Caption = "Portuguese"
    fraTypeRpt.BorderStyle = 0
    fraTypeRpt.Top = 3840 - 300
    fraTypeRpt.Visible = True
    
End If

If glbWFC_IPPopFormName = "WFCIPPrintSpreadsheet" Then 'Ticket #29810 Franks 02/07/2017
    Me.Caption = "Print Incentive Plan Spreadsheet"
    fraEmpLetter.Left = 0
    fraEmpLetter.Top = 200
    fraEmpLetter.Height = 6015 ' 9135
    fraEmpLetter.Width = 14535
    fraEmpLetter.BorderStyle = 0
    fraEmpLetter.Visible = True
    
    lblPayDate.Visible = False
    dlpPayDate.Visible = False
    lblComment.Visible = False
    txtComment.Visible = False
    
    xTmp = 1200
    lblRepGrp.Top = lblRepGrp.Top - xTmp
    lblGrp(0).Top = lblGrp(0).Top - xTmp
    comGroup(0).Top = comGroup(0).Top - xTmp
    lblGrp(3).Top = lblGrp(3).Top - xTmp
    comGroup(1).Top = comGroup(1).Top - xTmp
    cmdPrint.Top = cmdPrint.Top - xTmp
    cmdClose2.Top = cmdClose2.Top - xTmp
    fraEmpLetter.Height = fraEmpLetter.Height - xTmp
    
    comGroup(0).Clear
    comGroup(0).AddItem lStr("Section")
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem "(none)"
    comGroup(0).ListIndex = 0
    comGroup(1).Clear
    comGroup(1).AddItem "Employee Name"
    comGroup(1).ListIndex = 0
    
    lblTypeRpt.Top = 2400 - 250
    lblTypeRpt.Visible = True
    fraTypeRpt.BorderStyle = 0
    fraTypeRpt.Top = 2400 - 300
    fraTypeRpt.Visible = True
    
    fraSunLife.Visible = True
    fraSunLife.Left = 600
    fraSunLife.Top = 6720 - 400 - xTmp
    fraSunLife.Height = 3615 - 600
    fraSunLife.Width = 9000
    
    cmdPrint.Caption = "Print"
End If

lblSection.Caption = lStr("Section")
lblRegion.Caption = lStr("Region")
lblVadim2.Caption = lStr("Vadim Field 2")

lblSectio2.Caption = lStr("Section")
lblRegion2.Caption = lStr("Region")

Call INI_Controls(Me)

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
Printable = False
End Property

Public Sub SET_UP_MODE()
Call set_Buttons
End Sub

Private Sub SetLocLabels()
lblLocation.Caption = lStr("Location")
lblRegion.Caption = lStr("Region")
lblAdmin.Caption = lStr("Administered By")
lblSection.Caption = lStr("Section")
End Sub

Private Sub MonthDescAdd()
ComMTH.AddItem "00-Annual Average Rate"
ComMTH.AddItem "01-January"
ComMTH.AddItem "02-February"
ComMTH.AddItem "03-March"
ComMTH.AddItem "04-April"
ComMTH.AddItem "05-May"
ComMTH.AddItem "06-June"
ComMTH.AddItem "07-July"
ComMTH.AddItem "08-August"
ComMTH.AddItem "09-September"
ComMTH.AddItem "10-October"
ComMTH.AddItem "11-November"
ComMTH.AddItem "12-December"
ComMTH.ListIndex = 0
End Sub


Private Sub WFC_CreatePayrollFilesUSA()
    Dim rsEPos As New ADODB.Recordset
    Dim rsTermEmp As New ADODB.Recordset
    Dim rsEmp As New ADODB.Recordset
    Dim exApp As Object, exBook As Object, exSheet As Object
    Dim SQLQ
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim xCol As Long
    Dim I, J, K, M, totNum
    Dim xYear
    
    xYear = MskFiscalYea2.Text
    Call getDateRange(xYear)
    
    For I = 1 To 50
        xPGList(I) = ""
        xPGEarnCode(I) = ""
    Next
    
    totPG = 0
    
    If Len(clpVadim2.Text) > 0 Then
        xPGList(1) = clpVadim2.Text
        xPGEarnCode(1) = comEarnCode.Text
        totPG = 1
    Else
        Call getPayGroupList("U.S.A.")
    End If
    
    For I = 1 To totPG
        Call WriteBON2FileUSA(xYear, xPGList(I), xPGEarnCode(I))
    Next
    
    MsgBox "   Finished.   "

End Sub

Private Sub WFC_CreatePayrollFilesCAD()
    Dim rsEPos As New ADODB.Recordset
    Dim rsTermEmp As New ADODB.Recordset
    Dim rsEmp As New ADODB.Recordset
    Dim exApp As Object, exBook As Object, exSheet As Object
    Dim SQLQ
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim xCol As Long
    Dim I, J, K, M, totNum
    Dim xYear
    
    xYear = MskFiscalYea2.Text
    Call getDateRange(xYear)
    
    For I = 1 To 50
        xPGList(I) = ""
        xPGEarnCode(I) = ""
    Next
    
    totPG = 0
    
    If Len(clpVadim2.Text) > 0 Then
        xPGList(1) = clpVadim2.Text
        xPGEarnCode(1) = comEarnCode.Text
        totPG = 1
    Else
        Call getPayGroupList("Canada")
    End If
    
    For I = 1 To totPG
        Call WriteBON2FileCanada(xYear, xPGList(I))
    Next
    
    MsgBox "   Finished.   "
End Sub

Private Sub WriteBON2FileUSA(xYear, xPayGroup, xEarnCode)
    Dim rsEmp As New ADODB.Recordset
    Dim SQLQ
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim xCol As Long
    Dim I, J, K, M, totNum
    Dim buf
    
    'xlsFileMat = Dir1.Path & IIf(Right(Dir1.Path, 1) = "\", "", "\") & "USA Incentive Payout-" & xPayGroup & "-" & xYear & ".txt"
    xlsFileMat = Dir1.Path & IIf(Right(Dir1.Path, 1) = "\", "", "\") & "USA Incentive Payout-" & xPayGroup & "-" & xYear & ".csv"
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    
    SQLQ = "SELECT HREARN.*, ED_PAYROLL_ID  FROM HREARN LEFT JOIN HREMP ON HREARN.EMPNBR = HREMP.ED_EMPNBR WHERE EARN_TYPE = 'BON4'  "
    SQLQ = SQLQ & "AND FDATE = " & Date_SQL(xlocFDate) & " "
    SQLQ = SQLQ & "AND TDATE = " & Date_SQL(xlocTDate) & " "
    SQLQ = SQLQ & "AND ED_VADIM2 = '" & xPayGroup & "' "
    SQLQ = SQLQ & "ORDER BY ED_PAYROLL_ID"

    If rsEmp.State <> 0 Then rsEmp.Close
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsEmp.EOF Then
        'MsgBox "No 'BON4' record found in this Selection Criteria for Pay Group '" & xPayGroup & "' "
        Exit Sub
    End If
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(0).FloodPercent = 0
    
    'xlsFileMat = Dir1.Path & IIf(Right(Dir1.Path, 1) = "\", "", "\") & "USA Incentive Payout-" & xPayGroup & "-" & xYear & ".txt"
    'If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    Open xlsFileMat For Output As #1
    
    'buf = "Co Code,Batch ID,File #,Pay #,Shift,Temp Rate,Temp Dept ,Reg Hours,O/T Hours,Hours 3 Code,Hours 3 Amount,Earnings 3 Code,Earnings 3 Amount,Earnings 4 Code,Earnings 4 Amount,Memo Code,Memo Amount,Tax Frequency"
    buf = "Co Code,Batch ID,File #,Pay #,Shift,Temp Rate,Temp Dept ,Reg Hours,O/T Hours,Hours 3 Code,Hours 3 Amount,Earnings 3 Code,Earnings 3 Amount,Earnings 4 Code,Earnings 4 Amount,Earnings 5 Code,Earnings 5 Amount,Memo Code,Memo Amount,Tax Frequency"
    
    Print #1, buf
    
    
    I = 0
    totNum = rsEmp.RecordCount
    xRow = 2
    Do While Not rsEmp.EOF
        If (I / totNum) <= 1 Then
            MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
            I = I + 1
        End If
        DoEvents

        buf = xPayGroup
        buf = buf & "," & "Bon" & Trim(xYear) & ""
        buf = buf & "," & rsEmp("ED_PAYROLL_ID") & ""
        buf = buf & "," & "2" & ""
        buf = buf & ","
        buf = buf & ","
        buf = buf & ","
        buf = buf & ","
        buf = buf & ","
        buf = buf & ","
        buf = buf & ","
        'buf = buf & ","
        'buf = buf & ","
        
        If xEarnCode = "3" Then
            buf = buf & "," & "B" & ""
            If Not IsNull(rsEmp("ACT_DOLLAR")) Then
                buf = buf & "," & Round(rsEmp("ACT_DOLLAR"), 2) & ""
            Else
                buf = buf & ","
            End If
            buf = buf & ","
            buf = buf & ","
            buf = buf & "," 'Earnings 5 Code
            buf = buf & "," 'Earnings 5 Amount
        ElseIf xEarnCode = "4" Then
            buf = buf & ","
            buf = buf & ","
            buf = buf & "," & "B" & ""
            'buf = buf & "," & rsEmp("ACT_DOLLAR") & ""
            If Not IsNull(rsEmp("ACT_DOLLAR")) Then
                buf = buf & "," & Round(rsEmp("ACT_DOLLAR"), 2) & ""
            Else
                buf = buf & ","
            End If
            buf = buf & "," 'Earnings 5 Code
            buf = buf & "," 'Earnings 5 Amount
        ElseIf xEarnCode = "5" Then
            buf = buf & ","
            buf = buf & ","
            buf = buf & ","
            buf = buf & ","
            buf = buf & "," & "B" & ""
            'buf = buf & "," & rsEmp("ACT_DOLLAR") & ""
            If Not IsNull(rsEmp("ACT_DOLLAR")) Then
                buf = buf & "," & Round(rsEmp("ACT_DOLLAR"), 2) & ""
            Else
                buf = buf & ","
            End If
        Else
            buf = buf & ","
            buf = buf & ","
            buf = buf & ","
            buf = buf & ","
            buf = buf & "," 'Earnings 5 Code
            buf = buf & "," 'Earnings 5 Amount
        End If
        'buf = buf & "," 'Earnings 5 Code
        'buf = buf & "," 'Earnings 5 Amount
        
        buf = buf & ","
        buf = buf & ","
        buf = buf & "," & "B" & ""
        Print #1, buf
        
        xRow = xRow + 1
        
        rsEmp.MoveNext
    Loop
    rsEmp.Close
    
    'For terminated employees
    ''SQLQ = "SELECT Term_EARN.*, ED_PAYROLL_ID  FROM Term_EARN LEFT JOIN Term_HREMP ON Term_EARN.TERM_SEQ = Term_HREMP.TERM_SEQ WHERE EARN_TYPE = 'BON4'  "
    ''SQLQ = SQLQ & "AND FDATE = " & Date_SQL(xlocFDate) & " "
    ''SQLQ = SQLQ & "AND TDATE = " & Date_SQL(xlocTDate) & " "
    ''SQLQ = SQLQ & "AND ED_VADIM2 = '" & xPayGroup & "' "
    ''SQLQ = SQLQ & "ORDER BY ED_PAYROLL_ID"
    ''If rsEmp.State <> 0 Then rsEmp.Close
    ''rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    ''If rsEmp.EOF Then
    ''    GoTo file_end
    ''End If
    ''
    ''I = 0
    ''totNum = rsEmp.RecordCount
    ''Do While Not rsEmp.EOF
    ''    If (I / totNum) <= 1 Then
    ''        MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
    ''        I = I + 1
    ''    End If
    ''    DoEvents
    ''
    ''    buf = xPayGroup
    ''    buf = buf & "," & "Bon" & Trim(xYear) & ""
    ''    buf = buf & "," & rsEmp("ED_PAYROLL_ID") & ""
    ''    buf = buf & "," & "2" & ""
    ''    buf = buf & ","
    ''    buf = buf & ","
    ''    buf = buf & ","
    ''    buf = buf & ","
    ''    buf = buf & ","
    ''    buf = buf & ","
    ''    buf = buf & ","
    ''    'buf = buf & ","
    ''    'buf = buf & ","
    ''    If xEarnCode = "3" Then
    ''        buf = buf & "," & "B" & ""
    ''        'buf = buf & "," & rsEmp("ACT_DOLLAR") & ""
    ''        If Not IsNull(rsEmp("ACT_DOLLAR")) Then
    ''            buf = buf & "," & Round(rsEmp("ACT_DOLLAR"), 2) & ""
    ''        Else
    ''            buf = buf & ","
    ''        End If
    ''        buf = buf & ","
    ''        buf = buf & ","
    ''    Else ' xEarnCode = "4"
    ''        buf = buf & ","
    ''        buf = buf & ","
    ''        buf = buf & "," & "B" & ""
    ''        'buf = buf & "," & rsEmp("ACT_DOLLAR") & ""
    ''        If Not IsNull(rsEmp("ACT_DOLLAR")) Then
    ''            buf = buf & "," & Round(rsEmp("ACT_DOLLAR"), 2) & ""
    ''        Else
    ''            buf = buf & ","
    ''        End If
    ''    End If
    ''    buf = buf & "," 'Earnings 5 Code
    ''    buf = buf & "," 'Earnings 5 Amount
    ''    buf = buf & ","
    ''    buf = buf & ","
    ''    buf = buf & "," & "B" & ""
    ''    Print #1, buf
    ''
    ''    xRow = xRow + 1
    ''
    ''    rsEmp.MoveNext
    ''Loop
    ''rsEmp.Close
    
file_end:

    Close #1
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = " "
End Sub

Private Sub WriteBON2FileCanada(xYear, xPayGroup)
    Dim exApp As Object, exBook As Object, exSheet As Object
    Dim rsEmp As New ADODB.Recordset
    Dim SQLQ
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim xCol As Long
    Dim I, J, K, M, totNum

    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "Cdn Incentive PayoutTmp.xls"
    xlsFileMat = Dir1.Path & IIf(Right(Dir1.Path, 1) = "\", "", "\") & "Cdn Incentive Payout-" & xPayGroup & "-" & xYear & ".xls"
    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    
    SQLQ = "SELECT HREARN.*, ED_PAYROLL_ID,ED_SURNAME,ED_FNAME  FROM HREARN LEFT JOIN HREMP ON HREARN.EMPNBR = HREMP.ED_EMPNBR WHERE EARN_TYPE = 'BON4'  "
    SQLQ = SQLQ & "AND FDATE = " & Date_SQL(xlocFDate) & " "
    SQLQ = SQLQ & "AND TDATE = " & Date_SQL(xlocTDate) & " "
    SQLQ = SQLQ & "AND ED_VADIM2 = '" & xPayGroup & "' "
    SQLQ = SQLQ & "ORDER BY ED_PAYROLL_ID"

    If rsEmp.State <> 0 Then rsEmp.Close
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsEmp.EOF Then
        'MsgBox "No 'BON4' record found in this Selection Criteria for Pay Group '" & xPayGroup & "' "
        Exit Sub
    End If
    
    
    ''If Dir(xlsFileTmp) = "" Then
    ''    MsgBox "There is no " & xlsFileTmp
    ''    Exit Sub
    ''End If
    ''If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    FileCopy xlsFileTmp, xlsFileMat
    
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(0).FloodPercent = 0
        
    'Populate Excel file - begin
    'Create new WorkBook of Excel
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(xlsFileMat)
    Set exSheet = exBook.Worksheets(1)
    
    I = 0
    totNum = rsEmp.RecordCount
    xRow = 2
    Do While Not rsEmp.EOF
        If (I / totNum) <= 1 Then
            MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
            I = I + 1
        End If
        DoEvents
    
        exSheet.Cells(xRow, 1) = xPayGroup
        exSheet.Cells(xRow, 2) = rsEmp("ED_FNAME") & " " & rsEmp("ED_SURNAME")
        exSheet.Cells(xRow, 3) = rsEmp("ED_PAYROLL_ID")
        exSheet.Cells(xRow, 4) = 20
        'exSheet.Cells(xRow, 4) = rsEmp("ACT_DOLLAR")
        If Not IsNull(rsEmp("ACT_DOLLAR")) Then
            exSheet.Cells(xRow, 5) = Round(rsEmp("ACT_DOLLAR"), 2)
        End If
        'exSheet.Cells(xRow, 5) = ""
        xRow = xRow + 1
        
        rsEmp.MoveNext
    Loop
    rsEmp.Close
    
    'For terminated employees
    ''SQLQ = "SELECT Term_EARN.*, ED_PAYROLL_ID  FROM Term_EARN LEFT JOIN Term_HREMP ON Term_EARN.TERM_SEQ = Term_HREMP.TERM_SEQ WHERE EARN_TYPE = 'BON4'  "
    ''SQLQ = SQLQ & "AND FDATE = " & Date_SQL(xlocFDate) & " "
    ''SQLQ = SQLQ & "AND TDATE = " & Date_SQL(xlocTDate) & " "
    ''SQLQ = SQLQ & "AND ED_VADIM2 = '" & xPayGroup & "' "
    ''SQLQ = SQLQ & "ORDER BY ED_PAYROLL_ID"
    ''If rsEmp.State <> 0 Then rsEmp.Close
    ''rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    ''If rsEmp.EOF Then
    ''    GoTo file_end
    ''End If
    ''
    ''I = 0
    ''totNum = rsEmp.RecordCount
    ''Do While Not rsEmp.EOF
    ''    If (I / totNum) <= 1 Then
    ''        MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
    ''        I = I + 1
    ''    End If
    ''    DoEvents
    ''
    ''    exSheet.Cells(xRow, 1) = xPayGroup
    ''    exSheet.Cells(xRow, 2) = Round(rsEmp("ED_PAYROLL_ID"), 2)
    ''    exSheet.Cells(xRow, 3) = 20
    ''    If Not IsNull(rsEmp("ACT_DOLLAR")) Then
    ''        exSheet.Cells(xRow, 4) = Round(rsEmp("ACT_DOLLAR"), 2)
    ''    End If
    ''    'exSheet.Cells(xRow, 5) = ""
    ''    xRow = xRow + 1
    ''
    ''    rsEmp.MoveNext
    ''Loop
    ''rsEmp.Close
    ''
file_end:

    exBook.Save
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = " "
    
End Sub

Private Sub WFC_CreateIncentivePlan()
    On Error GoTo WFC_CreateIncentivePlan_Err
    Dim rsEPos As New ADODB.Recordset
    Dim rsEPo2 As New ADODB.Recordset
    Dim rsESal As New ADODB.Recordset
    Dim rsESa2 As New ADODB.Recordset
    Dim rsTermEmp As New ADODB.Recordset
    Dim rsEmp As New ADODB.Recordset
    Dim rsJOB As New ADODB.Recordset
    Dim rsJOBMASTER As New ADODB.Recordset
    Dim rsBand As New ADODB.Recordset
    Dim rsFactors As New ADODB.Recordset
    Dim rsFac2 As New ADODB.Recordset
    Dim rsPerf As New ADODB.Recordset
    Dim exApp As Object, exBook As Object, exSheet As Object
    Dim SQLQ
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim xCol As Long
    Dim I, J, K, M, totNum
    Dim xStartLine As Integer
    Dim xStartColu As Integer
    Dim xPlant As String
    Dim hasEmpPos As Boolean
    Dim hasEmpSal As Boolean
    Dim hasPosMaster As Boolean
    Dim hasFactors As Boolean
    Dim xCurAnnSal, xJun1stAnnSal
    Dim xSalAsDate ' xJun1stDate
    Dim xRateToCAD, xSerMonths, xCurMonth, xTemp, xStr
    Dim xYear, xType, xBU
    Dim xLocCurrency, xCDNPayout, xROIC
    Dim xIncEmpStatusList
    Dim xFiscalFrom, xFiscalTo
    Dim xBand1Date, xBand2Date, xBand1Mths, xBand2Mths, xBand1SalPerc, xBand2SalPerc, xCol_23Amt
    Dim xBandChgFlag As Boolean
    Dim xPreSalary
    
    xIncEmpStatusList = "'ACP,','ACT,','CB,','FMLA,','FS,','LOA,','MAT,','MIL,','MODW,','PAT,','STD,'"
    
    SQLQ = "SELECT * FROM HREMP WHERE (1=1) "
    'xIncEmpStatusList
    '"'ACP','ACT','CB','FMLA','FS','LOA','MAT','MIL','MODW','PAT','STD'"
    'SQLQ = SQLQ & "AND ED_EMP + ',' IN (" & xIncEmpStatusList & ") " 'LTRIM(RTRIM(ED_EMP))
    SQLQ = SQLQ & "AND LTRIM(RTRIM(ED_EMP)) + ',' IN (" & xIncEmpStatusList & ") " '
    SQLQ = SQLQ & "AND NOT (LEFT(ED_DIV,1) = '9') " 'Exclude all divisions/employees whose division is 9000 or higher
    'o   Exclude employees whose Position Status is one of the highlighted statuses
    SQLQ = SQLQ & "AND (ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE WHERE JH_CURRENT = 1 AND (HRJOB.JB_STATUS = 'CLER' OR HRJOB.JB_STATUS = 'MGMT' OR HRJOB.JB_STATUS = 'SUPR')))"
    If Len(clpDiv.Text) > 0 Then
        SQLQ = SQLQ & "AND ED_DIV IN ('" & Replace(clpDiv.Text, ",", "','") & "') "
    End If
    If Len(clpDept.Text) > 0 Then
        SQLQ = SQLQ & "AND ED_DEPTNO IN ('" & Replace(clpDept.Text, ",", "','") & "') "
    End If
    If Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & "AND ED_LOC IN ('" & Replace(clpCode(0).Text, ",", "','") & "') "
    End If
    If Len(clpCode(3).Text) > 0 Then
        SQLQ = SQLQ & "AND ED_REGION IN ('" & Replace(clpCode(3).Text, ",", "','") & "') "
    End If
    If Len(clpCode(4).Text) > 0 Then
        SQLQ = SQLQ & "AND ED_ADMINBY IN ('" & Replace(clpCode(4).Text, ",", "','") & "') "
    End If
    If Len(clpCode(2).Text) > 0 Then
        SQLQ = SQLQ & "AND ED_SECTION IN ('" & Replace(clpCode(2).Text, ",", "','") & "') "
    End If
    If Len(elpEEID.Text) > 0 Then
        SQLQ = SQLQ & "AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
    End If
    If Len(clpJob.Text) > 0 Then
        SQLQ = SQLQ & "AND ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT = 1 AND JH_JOB = '" & clpJob.Text & "') "
    End If
    If Len(clpJobMaster.Text) > 0 Then
        SQLQ = SQLQ & "AND ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE WHERE JH_CURRENT = 1 AND HRJOB.JB_JOBCODE = '" & clpJobMaster.Text & "') "
    End If
    If Len(clpCode(1).Text) > 0 Then
        SQLQ = SQLQ & "AND ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE WHERE JH_CURRENT = 1 AND HRJOB.JB_POSTYPE = '" & clpCode(1).Text & "') "
    End If
    SQLQ = SQLQ & " ORDER BY ED_SURNAME, ED_FNAME "
    
    If rsEmp.State <> 0 Then rsEmp.Close
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsEmp.EOF Then
        MsgBox "No record found in this Selection Criteria."
        Exit Sub
    End If
    If Not rsEmp.EOF Then
        rsEmp.MoveFirst
        totNum = rsEmp.RecordCount: I = 0
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait"
        MDIMain.panHelp(0).FloodPercent = 0
        Screen.MousePointer = HOURGLASS
    End If
    
    If cptExc2010.Value Then
        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "WFC_IncentivePlan_Tmp.xlsx"
        xlsFileMat = Dir1.Path & IIf(Right(Dir1.Path, 1) = "\", "", "\") & "WFC_IncentivePlan(" & Trim(glbUserID) & ").xlsx"
    Else
        'Excel 2003
        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "WFC_IncentivePlan_Tmp.xls"
        xlsFileMat = Dir1.Path & IIf(Right(Dir1.Path, 1) = "\", "", "\") & "WFC_IncentivePlan(" & Trim(glbUserID) & ").xls"
    End If

    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    FileCopy xlsFileTmp, xlsFileMat
    
    'Populate Excel file - begin
    'Create new WorkBook of Excel
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(xlsFileMat)
    Set exSheet = exBook.Worksheets(1)
            
    'exSheet.Cells(2, 2) = "Status of WPS Training for " & Year(Date)
    exSheet.Cells(2, 1) = "" & Format(Date, "MMM dd, YYYY") & ""
    
    exSheet.Cells(4, 26) = dlpAsOf.Text 'Col Z
    
    xYear = Year(dlpAsOf.Text)
    xFiscalFrom = getWFCFiscalYearStartDate(CVDate(dlpAsOf.Text))
    xFiscalTo = getWFCFiscalYearToDate(CVDate(dlpAsOf.Text))
    
    SQLQ = "SELECT * FROM HRIP_FACTORS WHERE IP_YEAR = " & xYear & " " 'IP_ROIC is same for all plants within one year
    If Len(clpCode(2).Text) > 0 Then
        SQLQ = SQLQ & "AND IP_SECTION = '" & clpCode(2).Text & "' "
    End If
    SQLQ = SQLQ & "ORDER BY IP_ROIC DESC "
    
    If rsFac2.State <> 0 Then rsFac2.Close
    rsFac2.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xROIC = ""
    If Not rsFac2.EOF Then
        xROIC = rsFac2("IP_ROIC")
        exSheet.Cells(6, 34) = xROIC 'AH6
        exSheet.Cells(6, 57) = xROIC 'AH6
    End If
    rsFac2.Close
            
    'PLANT Factors ------------------------------
    xType = "PLANT"
    SQLQ = "SELECT * FROM HRIP_FACTORS WHERE IP_YEAR = " & xYear & " "
    SQLQ = SQLQ & "AND IP_POSTYPE = '" & xType & "' "
    If rsFac2.State <> 0 Then rsFac2.Close
    rsFac2.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsFac2.EOF Then
        'Adjusted By
        exSheet.Cells(3, 38) = rsFac2("IP_A_PLANT_OBJ")   'AL3
        exSheet.Cells(3, 39) = rsFac2("IP_A_BU_FIN")   'AM3
        exSheet.Cells(3, 40) = rsFac2("IP_A_CORP_FIN")   'AN3
        'lblTarget
        exSheet.Cells(5, 38) = rsFac2("IP_T_PLANT_OBJ")   'AL5
        exSheet.Cells(5, 39) = rsFac2("IP_T_BU_FIN")   'AM5
        exSheet.Cells(5, 40) = rsFac2("IP_T_CORP_FIN")   'AN5
    End If
    rsFac2.Close
    
    'BU Factors ------------------------------
    xType = "BU"
    SQLQ = "SELECT * FROM HRIP_FACTORS WHERE IP_YEAR = " & xYear & " "
    SQLQ = SQLQ & "AND IP_POSTYPE = '" & xType & "' "
    If rsFac2.State <> 0 Then rsFac2.Close
    rsFac2.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsFac2.EOF Then
        'Adjusted By
        exSheet.Cells(3, 43) = rsFac2("IP_A_BU_FIN")   'AQ3
        exSheet.Cells(3, 44) = rsFac2("IP_A_CORP_FIN")  'AR3
        'lblTarget
        exSheet.Cells(5, 43) = rsFac2("IP_T_BU_FIN")   'AQ5
        exSheet.Cells(5, 44) = rsFac2("IP_T_CORP_FIN")   'AR5
    End If
    rsFac2.Close
    
    'COMM Factors ------------------------------
    xType = "COMM"
    SQLQ = "SELECT * FROM HRIP_FACTORS WHERE IP_YEAR = " & xYear & " "
    SQLQ = SQLQ & "AND IP_POSTYPE = '" & xType & "' "
    If rsFac2.State <> 0 Then rsFac2.Close
    rsFac2.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsFac2.EOF Then
        'Adjusted By
        exSheet.Cells(3, 47) = rsFac2("IP_A_SALES_IND")   'AU3
        exSheet.Cells(3, 48) = rsFac2("IP_A_SALES_COMM")   'AV3
        exSheet.Cells(3, 49) = rsFac2("IP_A_CORP_FIN")   'AW3
        'lblTarget
        exSheet.Cells(5, 47) = rsFac2("IP_T_SALES_IND")   'AU5
        exSheet.Cells(5, 48) = rsFac2("IP_T_SALES_COMM")   'AV5
        exSheet.Cells(5, 49) = rsFac2("IP_T_CORP_FIN")   'AW5
    End If
    rsFac2.Close
        
    'CORP Factors ------------------------------
    xType = "CORP"
    SQLQ = "SELECT * FROM HRIP_FACTORS WHERE IP_YEAR = " & xYear & " "
    SQLQ = SQLQ & "AND IP_POSTYPE = '" & xType & "' "
    If rsFac2.State <> 0 Then rsFac2.Close
    rsFac2.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsFac2.EOF Then
        'Adjusted By
        exSheet.Cells(3, 52) = rsFac2("IP_A_CORP_OBJ")   'AZ3
        exSheet.Cells(3, 53) = rsFac2("IP_A_CORP_FIN")  'BA3
        'lblTarget
        exSheet.Cells(5, 52) = rsFac2("IP_T_CORP_OBJ")   'AZ5
        exSheet.Cells(5, 53) = rsFac2("IP_T_CORP_FIN")   'BA5
    End If
    rsFac2.Close
    
    
        
    'First line of data
    xStartLine = 8
    xRow = xStartLine
    'xJun1stDate = CVDate("Jun 1, " & Year(Date) - 1)
    xSalAsDate = CVDate(dlpSalAsOf.Text)
    Do While Not rsEmp.EOF
        If (I / totNum) <= 1 Then
            MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
            I = I + 1
        End If
        DoEvents
                    
        'Emp Salary:Salaries will be as of June 1st of the previous year or the starting salary for new employees.
        SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE SH_EMPNBR = " & rsEmp("ED_EMPNBR") & " "
        'SQLQ = SQLQ & "AND SH_EDATE <= " & Date_SQL(xJun1stDate) & " ORDER BY SH_EDATE DESC"
        SQLQ = SQLQ & "AND SH_EDATE <= " & Date_SQL(xSalAsDate) & " ORDER BY SH_EDATE DESC"
        If rsESal.State <> 0 Then rsESal.Close
        rsESal.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If rsESal.EOF Then 'can not find the previous record, treate as new hire, use current Salary
            SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE SH_CURRENT = 1 AND SH_EMPNBR = " & rsEmp("ED_EMPNBR") & " "
            If rsESal.State <> 0 Then rsESal.Close
            rsESal.Open SQLQ, gdbAdoIhr001, adOpenStatic
        End If
        If rsESal.EOF Then hasEmpSal = False Else hasEmpSal = True
        If rsESal.EOF Then
            GoTo Next_Rec
        End If
                            
        'Emp Pos
        ''SQLQ = "SELECT JB_JOBCODE, JB_POSTYPE, HR_JOB_HISTORY.* FROM HR_JOB_HISTORY LEFT JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE  WHERE JH_CURRENT = 1 AND JH_EMPNBR = " & rsEmp("ED_EMPNBR")
        'SQLQ = "SELECT HR_JOB_HISTORY.* FROM HR_JOB_HISTORY WHERE JH_EMPNBR = " & rsEmp("ED_EMPNBR") 'JH_CURRENT = 1 AND
        SQLQ = "SELECT HR_JOB_HISTORY.*,JB_BAND FROM HR_JOB_HISTORY LEFT JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE WHERE JH_EMPNBR = " & rsEmp("ED_EMPNBR") 'JH_CURRENT = 1 AND
        SQLQ = SQLQ & "AND JH_JOB = '" & rsESal("SH_JOB") & "' "
        SQLQ = SQLQ & "AND JH_SDATE = " & Date_SQL(rsESal("SH_SDATE")) & " "
        If rsEPos.State <> 0 Then rsEPos.Close
        rsEPos.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If rsEPos.EOF Then hasEmpPos = False Else hasEmpPos = True
        hasPosMaster = False
        If hasEmpPos Then
            If rsJOB.State <> 0 Then rsJOB.Close
            SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & rsEPos("JH_JOB") & "' "
            rsJOB.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If rsJOB.EOF Then hasPosMaster = False Else hasPosMaster = True
            
            SQLQ = "SELECT * FROM HRJOBMASTER WHERE JB_JOBCODE = '" & rsJOB("JB_JOBCODE") & "' "
            If rsJOBMASTER.State <> 0 Then rsJOBMASTER.Close
            rsJOBMASTER.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsJOBMASTER.EOF Then
                If Not IsNull(rsJOBMASTER("JB_STATUS")) Then
                    If rsJOBMASTER("JB_STATUS") = "CLER" Or rsJOBMASTER("JB_STATUS") = "MGMT" Or rsJOBMASTER("JB_STATUS") = "SUPR" Then
                        'continue
                    Else
                        'Ticket #29562 Franks 12/13/2016
                        'This position is based on "Salary As of Date", it maybe is not as same as current position
                        'if the JB_STATUS(Job Class) is not CLER,MGMT,SUPR, then skip it
                        GoTo Next_Rec
                    End If
                End If
            End If
        End If
        
        
        'open performance recrod which is matching the Salary
        SQLQ = "SELECT * FROM HR_PERFORM_HISTORY WHERE PH_EMPNBR = " & rsEmp("ED_EMPNBR") & " "
        SQLQ = SQLQ & "AND YEAR(PH_PREVIEW) = " & xYear & " "
        SQLQ = SQLQ & " ORDER BY PH_PREVIEW DESC, PH_PNEXT DESC"
        If rsPerf.State <> 0 Then rsPerf.Close
        rsPerf.Open SQLQ, gdbAdoIhr001, adOpenStatic
        
        xType = ""
        If hasPosMaster Then
            If Not IsNull(rsJOB("JB_POSTYPE")) Then
                xType = rsJOB("JB_POSTYPE")
            End If
        End If
        xBU = ""
        If Not IsNull(rsEmp("ED_REGION")) Then
            xBU = rsEmp("ED_REGION")
        End If
        xPlant = ""
        If Not IsNull(rsEmp("ED_SECTION")) Then
            xPlant = rsEmp("ED_SECTION")
        End If
        'Open Factors records - begin ******************************************************************
        hasFactors = False
        'on Company Incentive Factors screen, BU and Plant are optional fields, so the program need to check all the conditions
        'option 1 - check all fields
        SQLQ = "SELECT * FROM HRIP_FACTORS WHERE IP_YEAR = " & xYear & " "
        SQLQ = SQLQ & "AND IP_POSTYPE = '" & xType & "' "
        SQLQ = SQLQ & "AND IP_REGION = '" & xBU & "' "
        SQLQ = SQLQ & "AND IP_SECTION = '" & xPlant & "' "
        If rsFactors.State <> 0 Then rsFactors.Close
        rsFactors.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsFactors.EOF Then
            hasFactors = True
        Else 'not found then check it without Plant
            SQLQ = "SELECT * FROM HRIP_FACTORS WHERE IP_YEAR = " & xYear & " "
            SQLQ = SQLQ & "AND IP_POSTYPE = '" & xType & "' "
            SQLQ = SQLQ & "AND IP_REGION = '" & xBU & "' "
            'SQLQ = SQLQ & "AND IP_SECTION = '" & xPlant & "' "
            If rsFactors.State <> 0 Then rsFactors.Close
            rsFactors.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsFactors.EOF Then
                hasFactors = True
            Else 'not found then check it without BU
                SQLQ = "SELECT * FROM HRIP_FACTORS WHERE IP_YEAR = " & xYear & " "
                SQLQ = SQLQ & "AND IP_POSTYPE = '" & xType & "' "
                'SQLQ = SQLQ & "AND IP_REGION = '" & xBU & "' "
                SQLQ = SQLQ & "AND IP_SECTION = '" & xPlant & "' "
                If rsFactors.State <> 0 Then rsFactors.Close
                rsFactors.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsFactors.EOF Then
                    hasFactors = True
                Else 'not found then check it without BU and Plant
                    SQLQ = "SELECT * FROM HRIP_FACTORS WHERE IP_YEAR = " & xYear & " "
                    SQLQ = SQLQ & "AND IP_POSTYPE = '" & xType & "' "
                    'SQLQ = SQLQ & "AND IP_REGION = '" & xBU & "' "
                    'SQLQ = SQLQ & "AND IP_SECTION = '" & xPlant & "' "
                    If rsFactors.State <> 0 Then rsFactors.Close
                    rsFactors.Open SQLQ, gdbAdoIhr001, adOpenStatic
                    If Not rsFactors.EOF Then
                        hasFactors = True
                    End If
                End If
            End If
        End If
        
        'Open Factors records - end ******************************************************************
        
        exSheet.Cells(xRow, 1) = rsEmp("ED_EMPNBR")
        exSheet.Cells(xRow, 2) = rsEmp("ED_COUNTRY")
        exSheet.Cells(xRow, 3) = rsEmp("ED_DIV")
        exSheet.Cells(xRow, 4) = rsEmp("ED_LOC")
        exSheet.Cells(xRow, 5) = rsEmp("ED_REGION")
        If hasPosMaster Then
            If Not IsNull(rsJOB("JB_POSTYPE")) Then exSheet.Cells(xRow, 6) = rsJOB("JB_POSTYPE")
        End If
        If hasEmpPos Then
            If Not IsNull(rsEPos("JH_REPTAU")) Then
                exSheet.Cells(xRow, 7) = rsEPos("JH_REPTAU")
                exSheet.Cells(xRow, 8) = GetEmpData(rsEPos("JH_REPTAU"), "ED_SURNAME", "") & ", " & GetEmpData(rsEPos("JH_REPTAU"), "ED_FNAME", "")
            End If
        End If
        exSheet.Cells(xRow, 9) = rsEmp("ED_EMPNBR")
        If Not IsNull(rsEmp("ED_EMP")) Then
            exSheet.Cells(xRow, 10) = rsEmp("ED_EMP") 'J - Status
        End If
        exSheet.Cells(xRow, 11) = GetEmpData(rsEmp("ED_EMPNBR"), "ED_SURNAME", "") & ", " & GetEmpData(rsEmp("ED_EMPNBR"), "ED_FNAME", "")
        If Not IsNull(rsEmp("ED_DOH")) Then exSheet.Cells(xRow, 12) = rsEmp("ED_DOH")
        If hasEmpSal Then
            If Not IsNull(rsESal("SH_NEXTDAT")) Then
                exSheet.Cells(xRow, 13) = rsESal("SH_NEXTDAT")
            End If
        End If
        If hasPosMaster Then
            exSheet.Cells(xRow, 14) = rsEPos("JH_JOB") & "/" & rsJOB("JB_JOBCODE")
            exSheet.Cells(xRow, 15) = rsJOB("JB_DESCR")
            SQLQ = "SELECT * FROM HRJOBMASTER WHERE JB_JOBCODE = '" & rsJOB("JB_JOBCODE") & "' "
            
            If rsJOBMASTER.State <> 0 Then rsJOBMASTER.Close
            rsJOBMASTER.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsJOBMASTER.EOF Then
                If Not IsNull(rsJOBMASTER("JB_STATUS")) Then exSheet.Cells(xRow, 16) = rsJOBMASTER("JB_STATUS")
                If Not IsNull(rsJOBMASTER("JB_GRPCD")) Then exSheet.Cells(xRow, 17) = rsJOBMASTER("JB_GRPCD")
            End If
        End If
        If hasEmpPos Then
            exSheet.Cells(xRow, 18) = rsEPos("JH_SDATE")
            If Not IsNull(rsEPos("JB_BAND")) Then exSheet.Cells(xRow, 21) = rsEPos("JB_BAND")
        End If
        If hasEmpSal Then
            If Not IsNull(rsESal("SH_MARKETLINE")) Then exSheet.Cells(xRow, 19) = rsESal("SH_MARKETLINE")
            'If Not IsNull(rsESal("SH_MDOLLARS")) Then exSheet.Cells(xRow, 20) = rsESal("SH_MDOLLARS")
            'If Not IsNull(rsESal("SH_BAND")) Then exSheet.Cells(xRow, 21) = rsESal("SH_BAND")
            
            'Annual Currency - Current Annual Salary for all countries except ARGENTINA, BRAZIL or MEXICO. If one of those countries, multiple Cell V by 12.
            'xCurAnnSal = getWFCAnnualSalary(rsESal("SH_SALARY"), rsESal("SH_SALCD"), rsESal("SH_WHRS"))
            xCurAnnSal = rsESal("SH_SALARY")
            xRateToCAD = getRateToCAD(Year(dlpAsOf.Text), Left(ComMTH.Text, 2), rsESal("SH_CURRENCYINDI"), 1)
            exSheet.Cells(xRow, 30) = xRateToCAD 'AD EXCH
            '''xJun1stAnnSal
            ''SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE SH_EMPNBR = " & rsEmp("ED_EMPNBR") & " "
            ''SQLQ = SQLQ & "AND SH_EDATE <= " & Date_SQL(xJun1stDate) & " ORDER BY SH_EDATE DESC"
            ''If rsESa2.State <> 0 Then rsESa2.Close
            ''rsESa2.Open SQLQ, gdbAdoIhr001, adOpenStatic
            ''xJun1stAnnSal = 0
            ''If Not rsESa2.EOF Then
            ''    xJun1stAnnSal = getWFCAnnualSalary(rsESa2("SH_SALARY"), rsESa2("SH_SALCD"), rsESa2("SH_WHRS"))
            ''End If
            '''exSheet.Cells(xRow, 22) = xJun1stAnnSal
            '''If Not xJun1stAnnSal = xCurAnnSal Then '??? testing
            '''    Debug.Print rsEmp("ED_EMPNBR")
            '''End If
            exSheet.Cells(xRow, 22) = xCurAnnSal
            'Current Annual Salary for all countries except ARGENTINA, BRAZIL or MEXICO. If one of those countries, multiple Cell V by 12.
            If rsEmp("ED_COUNTRY") = "ARGENTINA" Or rsEmp("ED_COUNTRY") = "BRAZIL" Or rsEmp("ED_COUNTRY") = "MEXICO" Then
                exSheet.Cells(xRow, 23) = xCurAnnSal * 12 ' xCurAnnSal * 12
            Else
                exSheet.Cells(xRow, 23) = xCurAnnSal ' xCurAnnSal
            End If
            xCol_23Amt = exSheet.Cells(xRow, 23)
            
            If Not IsNull(rsESal("SH_MARKETLINE")) And Not IsNull(rsESal("SH_BAND")) And Not IsNull(rsESal("SH_SECTION")) And Not IsNull(rsESal("SH_FISCALYEAR")) Then
                SQLQ = "SELECT * FROM WFC_Salary_Administration " 'LDOLLARS,MDOLLARS,HDOLLARS
                SQLQ = SQLQ & " WHERE [BAND]='" & rsESal("SH_BAND") & "'"
                SQLQ = SQLQ & " AND [MARKETLINE]='" & rsESal("SH_MARKETLINE") & "'"
                SQLQ = SQLQ & " AND SectionCode='" & rsESal("SH_SECTION") & "' "
                If IsNumeric(rsESal("SH_FISCALYEAR")) Then
                    SQLQ = SQLQ & " AND FiscalYear='" & rsESal("SH_FISCALYEAR") & "' "
                End If
                If rsBand.State <> 0 Then rsBand.Close
                rsBand.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsBand.EOF Then
                    If Not IsNull(rsBand("MIDPOINT_PER")) Then
                        exSheet.Cells(xRow, 24) = exSheet.Cells(xRow, 23) * rsBand("MIDPOINT_PER")
                    End If
                    If Not IsNull(rsBand("MDollars")) Then exSheet.Cells(xRow, 20) = rsBand("MDollars") 'Mid Point
                End If
            End If
            'get Exch Rate
            'exSheet.Cells(xRow, 25) = exSheet.Cells(xRow, 24) * xRateToCAD 'Y - use formula
            'exSheet.Cells(xRow, 30) = xRateToCAD 'column AD EXCH
            xSerMonths = (DateDiff("D", rsEmp("ED_DOH"), CVDate(dlpAsOf.Text))) / 30 'keep the same logic as what on their file '/ (365 / 12)

            'If rsEmp("ED_EMPNBR") = 10980076 Then
            'Debug.Print ""
            'End If
            xSerMonths = Round(xSerMonths, 2)
            exSheet.Cells(xRow, 26) = xSerMonths 'column Z
            If xSerMonths < 12 Then
                'exSheet.Cells(xRow, 28) = xSerMonths
                If xSerMonths < 6 Then
                    xCurMonth = xSerMonths '0
                Else
                    xCurMonth = Round(xSerMonths, 0)
                End If
            Else
                'exSheet.Cells(xRow, 28) = 12
                xCurMonth = 12
            End If
            'xCurMonth = exSheet.Cells(xRow, 28)
            exSheet.Cells(xRow, 28) = xCurMonth
            
            xLocCurrency = ""
            xCDNPayout = ""
            'Local Currency
            'If employee hasn't changed positions within the fiscal year, cell AC = (cell X*cell AB)/12. (talked this with Jerry, leave Otherwise for now)Otherwise, the calculation takes the target & # of months and adds the
            '---------------------- Band change -------------- begin
            xBandChgFlag = False
            locBandCurr = ""
            locBandPrev = ""
            xBand1Date = ""
            xBand2Date = ""
            If hasEmpPos Then
                If Not IsNull(rsEPos("JB_BAND")) Then 'Band not blank
                    locBandCurr = rsEPos("JB_BAND")
                    xBand1Date = rsEPos("JH_SDATE")
    
                    If CVDate(xBand1Date) > CVDate(xFiscalFrom) And CVDate(xBand1Date) < CVDate(xFiscalTo) Then 'Have New Position in the fiscal year
                    
                        'If CVDate(rsEmp("ED_DOH")) < CVDate(xFiscalFrom) Then 'DOH not in this year
                            'check the previous position band, if it was changed then using then calculate use both Band
                            SQLQ = "SELECT HR_JOB_HISTORY.*,JB_BAND FROM HR_JOB_HISTORY LEFT JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE WHERE JH_EMPNBR = " & rsEmp("ED_EMPNBR") 'JH_CURRENT = 1 AND
                            'SQLQ = SQLQ & "AND JH_JOB = '" & rsESal("SH_JOB") & "' "
                            SQLQ = SQLQ & " AND JH_SDATE < " & Date_SQL(rsEPos("JH_SDATE")) & " "
                            SQLQ = SQLQ & "ORDER BY JH_SDATE DESC"
                            If rsEPo2.State <> 0 Then rsEPo2.Close
                            rsEPo2.Open SQLQ, gdbAdoIhr001, adOpenStatic
                            xPreSalary = ""
                            If Not rsEPo2.EOF Then
                                'found previous position
                                If Not IsNull(rsEPo2("JB_BAND")) Then 'Previous Band not blank
                                    locBandPrev = rsEPo2("JB_BAND")
                                    If Len(locBandCurr) > 0 And Len(locBandPrev) > 0 Then
                                        If Not (locBandCurr = locBandPrev) Then 'Band was changed in this fiscal year
                                            'get the months and % of Salary
                                            xBand2Date = rsEPo2("JH_SDATE")
                                            
                                            xPreSalary = xCol_23Amt 'default to the latest Salary
                                            'find the recent Salary with the previous Band and position
                                            SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE SH_EMPNBR = " & rsEmp("ED_EMPNBR") & " "
                                            SQLQ = SQLQ & "AND SH_SDATE = " & Date_SQL(rsEPo2("JH_SDATE")) & " "
                                            SQLQ = SQLQ & "AND SH_JOB = '" & rsEPo2("JH_JOB") & "' "
                                            SQLQ = SQLQ & "AND SH_BAND = '" & locBandPrev & "' "
                                            SQLQ = SQLQ & "ORDER BY SH_EDATE DESC"
                                            If rsESa2.State <> 0 Then rsESa2.Close
                                            rsESa2.Open SQLQ, gdbAdoIhr001, adOpenStatic
                                            If Not rsESa2.EOF Then
                                                xPreSalary = rsESa2("SH_SALARY")
                                            End If
                                            rsESa2.Close
        
                                            'xBand1Mths, xBand2Mths
                                            xBand1Mths = getLocIncMonths(xBand1Date, xFiscalTo)
                                            If CVDate(rsEmp("ED_DOH")) > CVDate(xFiscalFrom) Then
                                                xBand2Mths = getLocIncMonths(rsEmp("ED_DOH"), xBand1Date)
                                            Else
                                                xBand2Mths = getLocIncMonths(xFiscalFrom, xBand1Date)
                                            End If
                                            If xBand2Mths < 0 Then
                                                xBand2Mths = 0
                                            End If
                                            
                                            xBand1SalPerc = getSalPercentageByBand(xPlant, locBandCurr, xYear)
                                            xBand2SalPerc = getSalPercentageByBand(xPlant, locBandPrev, xYear)
                                            
                                            'xTemp = (exSheet.Cells(xRow, 24) * xCurMonth) / 12  'xSerMonths
                                            'xTemp = (xCol_23Amt * xBand1Mths / xCurMonth) * xBand1SalPerc + (xCol_23Amt * xBand2Mths / xCurMonth) * xBand2SalPerc
                                            'Ticket #29586 Franks 12/20/2016 - use the previous Salary with band change(xPreSalary)
                                            'Ticket #29586 Franks 12/20/2016 - use the previous Salary with band change(xPreSalary)
                                            If xCurMonth = 0 Then 'Ticket #30528 Franks 08/22/2017
                                                xTemp = 0
                                            Else
                                                xTemp = (xCol_23Amt * xBand1Mths / xCurMonth) * xBand1SalPerc + (xPreSalary * xBand2Mths / xCurMonth) * xBand2SalPerc
                                            End If
                                            xLocCurrency = Round(xTemp, 2)
                                            If xRateToCAD = 0 Then
                                                xCDNPayout = 0
                                                'exSheet.Cells(xRow, 58) = "Divide by zero"
                                            Else
                                                xCDNPayout = Round(xLocCurrency / xRateToCAD, 2) 'AG
                                            End If
    
                                            'xStr = xBand2Mths & " months @ band " & locBandPrev & "(" & xBand2SalPerc * 100 & "%) and " & xBand1Mths & " months @ band " & locBandCurr & "(" & xBand1SalPerc * 100 & "%)" '
                                            'Ticket #29586 Franks 12/20/2016 - use the previous Salary with band change(xPreSalary)
                                            xStr = xBand2Mths & " months @ band " & locBandPrev & "($" & xPreSalary & ")" & "(" & xBand2SalPerc * 100 & "%) and " & xBand1Mths & " months @ band " & locBandCurr & "($" & xCol_23Amt & ")" & "(" & xBand1SalPerc * 100 & "%)"  '
                                                
                                            If xBand2Mths = 0 Then
                                                'Debug.Print ""
                                            Else
                                                'Debug.Print rsEmp("ED_EMPNBR") & xStr ' & "   " & xBand2Mths & " months @ band " & locBandPrev & " %" & xBand1SalPerc * 100 & " and " & xBand1Mths & " months @ band " & " %" & xBand2SalPerc * 100 & locBandCurr '" " & locBandPrev & "-" & xBand2Mths & locBandCurr & "-" & xBand1Mths
                                                '4 months @ band e and 8 months @ band f
                                                exSheet.Cells(xRow, 58 + 3) = xStr
                                                exSheet.Cells(xRow, 29) = Round(xLocCurrency, 2) 'AC
                                                xBandChgFlag = True
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        'End If
                    End If
                End If
            End If
            '---------------------- Band change -------------- end
            
            If xBandChgFlag Then
                'Band was changed then program gets xLocCurrency and xCDNPayout in section above
            Else
                If Not IsEmpty(exSheet.Cells(xRow, 24)) Then
                    xTemp = (exSheet.Cells(xRow, 24) * xCurMonth) / 12  'xSerMonths
                    xLocCurrency = Round(xTemp, 2)
                    'exSheet.Cells(xRow, 29) = Round(xLocCurrency, 2) 'AC
                    'exSheet.Cells(xRow, 32) = Round(xLocCurrency * xRateToCAD, 2) 'AF - AC * Exchange Rate (Cell AD)
                    'exSheet.Cells(xRow, 33) = Round(xLocCurrency * xRateToCAD, 0) 'AG - AC * Exchange Rate (Cell AD) rounded to the nearest whole dollar
                    'xCDNPayout = Round(xLocCurrency * xRateToCAD, 0) 'AG
                    If xRateToCAD = 0 Then
                        xCDNPayout = 0
                        exSheet.Cells(xRow, 58 + 3) = "Divide by zero"
                    Else
                        xCDNPayout = Round(xLocCurrency / xRateToCAD, 2) 'AG
                    End If
                End If
            End If
            'If xLocCurrency > 0 Then
            'Debug.Print ""
            'End If
            
            xTemp = xCurAnnSal * xRateToCAD
            'exSheet.Cells(xRow, 31) = xTemp 'AE
            ''SQLQ = "SELECT * FROM HRIP_FACTORS WHERE IP_YEAR = " & xYear & " " 'IP_ROIC is same for all plants within one year
            ''If rsFac2.State <> 0 Then rsFac2.Close
            ''rsFac2.Open SQLQ, gdbAdoIhr001, adOpenStatic
            ''xROIC = ""
            ''If Not rsFac2.EOF Then
            ''    xROIC = rsFac2("IP_ROIC")
            ''End If
            ''rsFac2.Close
            If IsNumeric(xROIC) Then
                If IsNumeric(xLocCurrency) Then
                    'AH 34 AH = AC * factor stored in Cell AH7  - ROIC
                    'exSheet.Cells(xRow, 34) = Round(xLocCurrency * xROIC, 2)
                    'AI 35 AI = AG * factor
                    'exSheet.Cells(xRow, 35) = Round(xLocCurrency * xRateToCAD, 0) * xROIC
                End If
            End If
        End If
        
        
        'AJ 36  - blank for now
        If hasPosMaster Then
            'AK 37  - If Position Group Code = "EXE" or "LEG", cell AK = cell AI
            If Not IsNull(rsJOB("JB_GRPCD")) Then
                If rsJOB("JB_GRPCD") = "EXE" Or rsJOB("JB_GRPCD") = "LEG" Then
                    exSheet.Cells(xRow, 37) = exSheet.Cells(xRow, 35)
                End If
            End If
        End If
        If hasFactors Then
            If UCase(xType) = "PLANT" Then
                'AL 38  - If Type = plant, then cell AG(col 33) * 50% (Plant percentage). Need to store the percentage somewhere.
                If Not IsNull(rsFactors("IP_T_PLANT_OBJ")) Then
                    'If Not IsEmpty(exSheet.Cells(xRow, 33)) Then
                        If IsNumeric(xCDNPayout) Then
                            exSheet.Cells(xRow, 38) = xCDNPayout * rsFactors("IP_T_PLANT_OBJ") * rsFactors("IP_A_PLANT_OBJ")  '
                        End If
                    'End If
                End If
                'AM 39
                If Not IsNull(rsFactors("IP_T_BU_FIN")) Then
                    'If Not IsEmpty(exSheet.Cells(xRow, 33)) Then
                        If IsNumeric(xCDNPayout) Then
                            exSheet.Cells(xRow, 39) = xCDNPayout * rsFactors("IP_T_BU_FIN") * rsFactors("IP_A_BU_FIN")
                        End If
                    'End If
                End If
                'AN 40
                If Not IsNull(rsFactors("IP_T_CORP_FIN")) Then
                    'If Not IsEmpty(exSheet.Cells(xRow, 33)) Then
                        If IsNumeric(xCDNPayout) Then
                            exSheet.Cells(xRow, 40) = xCDNPayout * rsFactors("IP_T_CORP_FIN") * rsFactors("IP_A_CORP_FIN")
                        End If
                    'End If
                End If
                'AO 41 - use formula
                'exSheet.Cells(xRow, 41) = exSheet.Cells(xRow, 38) + exSheet.Cells(xRow, 39) + exSheet.Cells(xRow, 40)
                'AP 42
                'If IsNumeric(xCDNPayout) Then
                '    If xCDNPayout > 0 Then
                '        If IsNumeric(exSheet.Cells(xRow, 41)) Then
                '            exSheet.Cells(xRow, 42) = (exSheet.Cells(xRow, 41) / xCDNPayout) * 100
                '        End If
                '    End If
                'End If
            End If 'UCase(xType) = "PLANT" - end
            
            If UCase(xType) = "BU" Then
                'AQ 43  If Type = bu, then Cell AG * .80 (BU percentage)
                If Not IsNull(rsFactors("IP_T_BU_FIN")) Then
                    'If Not IsEmpty(exSheet.Cells(xRow, 33)) Then
                        If IsNumeric(xCDNPayout) Then
                            exSheet.Cells(xRow, 43) = xCDNPayout * rsFactors("IP_T_BU_FIN") * rsFactors("IP_A_BU_FIN")
                        End If
                    'End If
                End If
                '44 AR
                If Not IsNull(rsFactors("IP_T_CORP_FIN")) Then
                    'If Not IsEmpty(exSheet.Cells(xRow, 33)) Then
                        If IsNumeric(xCDNPayout) Then
                            exSheet.Cells(xRow, 44) = xCDNPayout * rsFactors("IP_T_CORP_FIN") * rsFactors("IP_A_CORP_FIN")
                        End If
                    'End If
                End If
                '45 AS
                'exSheet.Cells(xRow, 45) = exSheet.Cells(xRow, 43) + exSheet.Cells(xRow, 44)
                '46 AT
                'If IsNumeric(xCDNPayout) Then
                '    If xCDNPayout > 0 Then
                '        If IsNumeric(exSheet.Cells(xRow, 45)) Then
                '            exSheet.Cells(xRow, 46) = (exSheet.Cells(xRow, 45) / xCDNPayout) * 100
                '        End If
                '    End If
                'End If
            End If 'UCase(xType) = "BU" - end
            
            
            If UCase(xType) = "COMM" Then
                '47 AU If Type = comm, then cell AG * 50% (Commercial percentage)
                If Not IsNull(rsFactors("IP_T_SALES_IND")) Then
                    'If Not IsEmpty(exSheet.Cells(xRow, 33)) Then
                        If IsNumeric(xCDNPayout) Then
                            exSheet.Cells(xRow, 47) = xCDNPayout * rsFactors("IP_T_SALES_IND") * rsFactors("IP_A_SALES_IND")
                        End If
                    'End If
                End If
                '48 AV
                If Not IsNull(rsFactors("IP_T_SALES_COMM")) Then
                    'If Not IsEmpty(exSheet.Cells(xRow, 33)) Then
                        If IsNumeric(xCDNPayout) Then
                            exSheet.Cells(xRow, 48) = xCDNPayout * rsFactors("IP_T_SALES_COMM") * rsFactors("IP_A_SALES_COMM")
                        End If
                    'End If
                End If
                '49 AW
                If Not IsNull(rsFactors("IP_T_CORP_FIN")) Then
                    'If Not IsEmpty(exSheet.Cells(xRow, 33)) Then
                        If IsNumeric(xCDNPayout) Then
                            exSheet.Cells(xRow, 49) = xCDNPayout * rsFactors("IP_T_CORP_FIN") * rsFactors("IP_A_CORP_FIN")
                        End If
                    'End If
                End If
                '50 AX
                'exSheet.Cells(xRow, 50) = exSheet.Cells(xRow, 47) + exSheet.Cells(xRow, 48) + exSheet.Cells(xRow, 49)
                '51 AY
                'If IsNumeric(xCDNPayout) Then
                '    If xCDNPayout > 0 Then
                '        If IsNumeric(exSheet.Cells(xRow, 450)) Then
                '            exSheet.Cells(xRow, 51) = (exSheet.Cells(xRow, 50) / xCDNPayout) * 100
                '        End If
                '    End If
                'End If
            End If 'UCase(xType) = "COMM" - end
            
            If UCase(xType) = "CORP" Then
                '52 AZ If Type = corp, then cell AG * 20% (Corporate teammate)
                If Not IsNull(rsFactors("IP_T_CORP_OBJ")) Then
                    'If Not IsEmpty(exSheet.Cells(xRow, 33)) Then
                        If IsNumeric(xCDNPayout) Then
                            exSheet.Cells(xRow, 52) = xCDNPayout * rsFactors("IP_T_CORP_OBJ") * rsFactors("IP_A_CORP_OBJ")
                        End If
                    'End If
                End If
                '53 BA
                If Not IsNull(rsFactors("IP_T_CORP_FIN")) Then
                    'If Not IsEmpty(exSheet.Cells(xRow, 33)) Then
                        If IsNumeric(xCDNPayout) Then
                            exSheet.Cells(xRow, 53) = xCDNPayout * rsFactors("IP_T_CORP_FIN") * rsFactors("IP_A_CORP_FIN")
                        End If
                    'End If
                End If
                '54 BB
                'exSheet.Cells(xRow, 54) = exSheet.Cells(xRow, 52) + exSheet.Cells(xRow, 53)
                '55 BC = Sum of cells AZ + BA.
                'exSheet.Cells(xRow, 55) = exSheet.Cells(xRow, 52) + exSheet.Cells(xRow, 53)
            End If 'UCase(xType) = "CORP"
            '65 BD = Sum of cells BC + AX + AS + AO (Summing the total for each group)
            'exSheet.Cells(xRow, 56) = exSheet.Cells(xRow, 55) + exSheet.Cells(xRow, 50) + exSheet.Cells(xRow, 45) + exSheet.Cells(xRow, 41)
        End If
        
        'Performance Review
        If Not rsPerf.EOF Then
            If Not IsNull(rsPerf("PH_PCODE2")) Then
                'xTemp = GetTABLDesc("SDPC", rsPerf("PH_PCODE2"))
                'exSheet.Cells(xRow, 57 + 3) = xTemp
                'exSheet.Cells(xRow, 57 + 3) = rsPerf("PH_PCODE2") ' xTemp '
                xTemp = ""
                If rsPerf("PH_PCODE2") = "1" Or rsPerf("PH_PCODE2") = "01" Then xTemp = "01 NOT ACHEIVED WAI"
                If rsPerf("PH_PCODE2") = "2" Or rsPerf("PH_PCODE2") = "02" Then xTemp = "02 PARTIALLY ACHEIVED WAI"
                If rsPerf("PH_PCODE2") = "3" Or rsPerf("PH_PCODE2") = "03" Then xTemp = "03 FULLY ACHEIVED WAI"
                If Len(xTemp) > 0 Then
                    exSheet.Cells(xRow, 57 + 3) = xTemp
                End If
                
                'xTemp = GetTABLDesc("SDPC", rsPerf("PH_PCODE2"))
                'xTemp = UCase(xTemp)
                'If xTemp = UCase("Achieved") Then
                '    exSheet.Cells(xRow, 67) = "X"
                'End If
                'If xTemp = UCase("Patially Achieved") Then
                '    exSheet.Cells(xRow, 68) = "X"
                'End If
                'If xTemp = UCase("Not achieved") Then
                '    exSheet.Cells(xRow, 69) = "X"
                'End If
            End If
            'If Not IsNull(rsPerf("PH_COMMENTS")) Then
            '    exSheet.Cells(xRow, 71) = rsPerf("PH_COMMENTS")
            'End If
        End If

        xRow = xRow + 1

Next_Rec:
        rsEmp.MoveNext
    Loop
    rsEmp.Close

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
    'Populate Excel file - end
    
    Exit Sub

'    exSheet.Cells(2, 2) = "Status of WPS Training for " & Year(Date)
''    exSheet.Cells(3, 2) = "(" & Format(Date, "MMM dd, YYYY") & ")"
'    'Less than 75% have been trained
'    'exSheet.Cells(4, 4).Interior.Color = RED
'    '75% and more have been trained
'    'exSheet.Cells(5, 4).Interior.Color = RGB(34, 139, 34) 'Dark Green
'    xRow = xStartLine
'    'xTotStaff = 0
'    'xTotPlant = 0
'    Do While Not rsDiv.EOF
'        If (I / totNum) <= 1 Then
'            MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
'            I = I + 1
'        End If
'        DoEvents
'        'Plant/Division
'        exSheet.Cells(xRow, 1) = rsDiv("JH_LUSER") & "/" & rsDiv("JH_COMMENT2") '"Woodbridge Corporate" ' rsDiv("JH_COMMENT2")
'        exSheet.Cells(xRow, 2) = rsDiv("JH_REPTAU2") 'Staff
'        exSheet.Cells(xRow, 3) = rsDiv("JH_REPTAU3") 'Plant
'        xRow = xRow + 1
'        exSheet.Cells(xRow, 1) = "% trained"
'        xRow = xRow + 1
'        xRow = xRow + 1
'        rsDiv.MoveNext
'    Loop
'    'Total ----
'    If xRow > xStartLine Then
'        exSheet.Cells(xRow, 1) = "TOTAL RECORDED"
'        'exSheet.Cells(xRow, 2) = xTotStaff
'        'exSheet.Cells(xRow, 3) = xTotPlant
'        xRow = xRow + 1
'        exSheet.Cells(xRow, 1) = "% trained"
'    End If

        
    Exit Sub
'-------------- End

WFC_CreateIncentivePlan_Err:


glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", "", "SELECT")
Exit Sub
Resume Next

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


For X% = 0 To 4
    If Not clpCode(X).ListChecker Then Exit Function
Next X%


If Len(clpJob.Text) > 0 And clpJob.Caption = "Unassigned" Then
    MsgBox "Position Code must be valid"
    clpJob.SetFocus
    Exit Function
End If
If Len(clpJobMaster.Text) > 0 And clpJobMaster.Caption = "Unassigned" Then
    MsgBox "Job Code must be valid"
    clpJobMaster.SetFocus
    Exit Function
End If

'If Len(clpCode(2)) = 0 Then
'    MsgBox lStr("Section is required.")
'    clpCode(2).SetFocus
'    Exit Function
'End If

If Len(dlpSalAsOf.Text) = 0 Then
    MsgBox "Salary As of Date is required!"
    dlpSalAsOf.SetFocus
    Exit Function
End If
If Not IsDate(dlpSalAsOf.Text) Then
    MsgBox "Not a valid date"
    dlpSalAsOf.SetFocus
    Exit Function
End If

If Len(dlpAsOf.Text) = 0 Then
    MsgBox "Service As of Date is required!"
    dlpAsOf.SetFocus
    Exit Function
End If
If Not IsDate(dlpAsOf.Text) Then
    MsgBox "Not a valid date"
    dlpAsOf.SetFocus
    Exit Function
End If

If Not elpEEID.ListChecker Then
    Exit Function
End If



CriCheck = True

End Function

Private Function getRows(exSheet As Object, xFirstLine)
Dim X
X = xFirstLine
Do While True
    If exSheet.Cells(X, 1) = "" Then
        Exit Do
    Else
        X = X + 1
    End If
Loop
getRows = X - 1
End Function

Private Function getLocIncMonths(xFr, xTo)
Dim xFrD, xToD
Dim xDay, xDNum, xMths
    xFrD = xFr
    xToD = xTo
    
    xDay = Day(xFrD)
    If xDay < 15 Then
        xFrD = CVDate((MonthName(month(xFrD)) & " 1," & Year(xFrD)))
    Else
        xFrD = CVDate((MonthName(month(xFrD)) & " 1," & Year(xFrD)))
        xFrD = DateAdd("M", 1, xFrD)
        xFrD = DateAdd("d", -1, xFrD)
    End If
    
    xDay = Day(xToD)
    If xDay < 15 Then
        xToD = CVDate((MonthName(month(xToD)) & " 1," & Year(xToD)))
    Else
        xToD = CVDate((MonthName(month(xToD)) & " 1," & Year(xToD)))
        xToD = DateAdd("M", 1, xToD)
        xToD = DateAdd("d", -1, xToD)
    End If
    
    xDNum = DateDiff("d", CVDate(xFrD), CVDate(xToD))
    xMths = xDNum / 30
    xMths = Round(xMths, 0)
    
    getLocIncMonths = xMths
End Function

Private Function getSalPercentageByBand(xPlant, xBand, xYear)
Dim rsBand As New ADODB.Recordset
Dim SQLQ
Dim retval
    retval = 0
    
    SQLQ = "SELECT * FROM WFC_Salary_Administration " 'LDOLLARS,MDOLLARS,HDOLLARS
    SQLQ = SQLQ & " WHERE [BAND]='" & xBand & "'"
    'SQLQ = SQLQ & " AND [MARKETLINE]='" & rsESal("SH_MARKETLINE") & "'"
    SQLQ = SQLQ & " AND SectionCode='" & xPlant & "' "
    If IsNumeric(xYear) Then
        SQLQ = SQLQ & " AND FiscalYear='" & xYear & "' "
    End If
    If rsBand.State <> 0 Then rsBand.Close
    rsBand.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsBand.EOF Then
        If Not IsNull(rsBand("MIDPOINT_PER")) Then
            retval = rsBand("MIDPOINT_PER")
        End If
    End If
    
    getSalPercentageByBand = retval
End Function


Private Sub getDateRange(xYear)
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim I As Integer
    xlocFDate = ""
    xlocTDate = ""
    If IsNumeric(xYear) Then
        If Len(xYear) = 4 Then
            SQLQ = "SELECT * FROM HRPARCO"
            If rs.State <> 0 Then rs.Close
            rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rs.EOF Then
                If Not IsNull(rs("PC_TDATE")) Then
                    I = xYear - Year(rs("PC_TDATE"))
                    xlocFDate = DateAdd("YYYY", I, rs("PC_FDATE"))
                    xlocTDate = DateAdd("YYYY", I, rs("PC_TDATE"))
                End If
            End If
            rs.Close
        End If
    End If
End Sub

Private Sub getPayGroupList(xCountry)
Dim exApp As Object, exBook As Object, exSheet As Object
Dim xlsFileTmp As String
Dim xRow, xRows
Dim I As Integer

    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "WFCPayCodeList.xls"
    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If

    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(xlsFileTmp)
    If xCountry = "U.S.A." Then
        Set exSheet = exBook.Worksheets(1)
    End If
    If xCountry = "Canada" Then
        Set exSheet = exBook.Worksheets(2)
    End If
    
    xRow = 2
    xRows = getRows(exSheet, xRow)
    
    totPG = 0
    For I = xRow To xRows
        xPGList(I - 1) = exSheet.Cells(I, 1)
        If Not IsEmpty(exSheet.Cells(I, 3)) Then
            xPGEarnCode(I - 1) = Left(Trim(exSheet.Cells(I, 3)), 1)
        Else
            xPGEarnCode(I - 1) = ""
        End If
        
        totPG = totPG + 1
    
    Next

    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing
    
'Dim xPGList(50)
'Dim totPG As Integer
End Sub

Private Function getEmpROIC(xYear, xSection)
Dim rsFac2 As New ADODB.Recordset
Dim SQLQ
Dim retval
    retval = ""
    
    SQLQ = "SELECT * FROM HRIP_FACTORS WHERE IP_YEAR = " & xYear & " " 'IP_ROIC is same for all plants within one year
    If Not IsNull(xSection) Then
        SQLQ = SQLQ & "AND IP_SECTION = '" & xSection & "' "
    End If
    SQLQ = SQLQ & "ORDER BY IP_ROIC DESC "
    
    If rsFac2.State <> 0 Then rsFac2.Close
    rsFac2.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsFac2.EOF Then
        retval = rsFac2("IP_ROIC")
    End If
    rsFac2.Close
    
    getEmpROIC = retval
End Function

Private Function getEmpOtherEarnByCode(xEmpNo, xCode, xFDate, xTDate)
Dim rsE As New ADODB.Recordset
Dim SQLQ
Dim retval
    retval = ""
    SQLQ = "SELECT *  FROM HREARN WHERE EARN_TYPE = '" & xCode & "' "
    SQLQ = SQLQ & "AND FDATE = " & Date_SQL(xFDate) & " "
    SQLQ = SQLQ & "AND TDATE = " & Date_SQL(xTDate) & " "
    SQLQ = SQLQ & "AND EMPNBR = " & xEmpNo & " "
    If rsE.State <> 0 Then rsE.Close
    rsE.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsE.EOF Then
        If Not IsNull(rsE("ACT_DOLLAR")) Then
            retval = rsE("ACT_DOLLAR")
        End If
    End If
    rsE.Close
    getEmpOtherEarnByCode = retval
End Function
